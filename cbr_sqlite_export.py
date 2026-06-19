#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Собирает данные ЦБ из Power Query книги в SQLite по списку банков.

Скрипт переиспользует M-исполнитель из OFUKB_CBR_PQ_alt_parser.py и не
создает/изменяет Excel-файлы.
"""

from __future__ import annotations

import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
import io
import json
import logging
import re
import sqlite3
import sys
import time
import traceback
import warnings
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup

from OFUKB_CBR_PQ_alt_parser import (
    MiniMEngine,
    PowerQueryWorkbookInspector,
    df_debug_summary,
    replace_regnum_in_m_code,
)


@dataclass(frozen=True)
class BankRef:
    regnum: str
    name: str = ""


@dataclass
class QueryFrame:
    query_name: str
    table: str
    df: pd.DataFrame


@dataclass
class BankExportData:
    bank: BankRef
    frames: List[QueryFrame]
    query_count: int
    rows_written: int


def setup_logging(verbose: bool) -> logging.Logger:
    logger = logging.getLogger("cbr_sqlite_export")
    logger.handlers.clear()
    logger.propagate = False
    logger.setLevel(logging.DEBUG if verbose else logging.INFO)

    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(logging.Formatter("%(asctime)s | %(levelname)-7s | %(message)s", "%Y-%m-%d %H:%M:%S"))
    handler.setLevel(logging.DEBUG if verbose else logging.INFO)
    logger.addHandler(handler)
    return logger


def parse_bank_refs_from_file(path: Path) -> List[BankRef]:
    banks: List[BankRef] = []
    seen = set()
    for raw_line in path.read_text(encoding="utf-8-sig").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        parts = [p.strip() for p in re.split(r"[,;\t]", line, maxsplit=1)]
        m = re.search(r"\d+", parts[0])
        if not m:
            continue
        regnum = m.group(0)
        if regnum in seen:
            continue
        seen.add(regnum)
        banks.append(BankRef(regnum=regnum, name=parts[1] if len(parts) > 1 else ""))
    return banks


def merge_bank_refs(*groups: Iterable[BankRef]) -> List[BankRef]:
    merged: List[BankRef] = []
    seen = set()
    for group in groups:
        for bank in group:
            if not bank.regnum or bank.regnum in seen:
                continue
            seen.add(bank.regnum)
            merged.append(bank)
    return merged


def is_empty_cell(value: object) -> bool:
    return value is None or pd.isna(value) or not str(value).strip()


def clean_cell_text(value: object) -> str:
    if is_empty_cell(value):
        return ""
    return " ".join(str(value).split())


def bank_refs_from_full_list_frame(df: pd.DataFrame) -> List[BankRef]:
    reg_col = next((col for col in df.columns if str(col) in {"cregnum", "Регистрационный номер"}), None)
    name_col = next((col for col in df.columns if str(col) in {"bnk_name", "Наименование"}), None)
    type_col = next((col for col in df.columns if str(col) in {"bnk_type", "Вид"}), None)
    status_col = next((col for col in df.columns if str(col) in {"lic_status", "Статус лицензии"}), None)
    if reg_col is None or name_col is None or status_col is None:
        return []

    banks: List[BankRef] = []
    for _, row in df.iterrows():
        status = clean_cell_text(row[status_col])
        bank_type = clean_cell_text(row[type_col]) if type_col is not None else ""
        if status != "Действующая" or bank_type:
            continue

        raw_regnum = clean_cell_text(row[reg_col])
        m = re.search(r"\d+", raw_regnum)
        if not m:
            continue
        banks.append(BankRef(regnum=m.group(0), name=clean_cell_text(row[name_col])))
    return banks


def find_full_list_xlsx_url(page_html: str, page_url: str) -> Optional[str]:
    soup = BeautifulSoup(page_html, "html.parser")
    for link in soup.find_all("a", href=True):
        text = clean_cell_text(link.get_text(" ", strip=True)).upper()
        href = str(link.get("href") or "")
        if text == "XLSX" or "DownloadExcel" in href:
            return urljoin(page_url, href)
    return None


def discover_bank_refs_from_cbr(timeout: int, logger: logging.Logger) -> List[BankRef]:
    """
    Пытается получить список regnum действующих банков с публичной страницы ЦБ.

    Это вспомогательный режим: структура сайта ЦБ может измениться, поэтому для
    воспроизводимых прогонов надежнее передавать --regnums-file.
    """
    page_url = "https://www.cbr.ru/banking_sector/credit/FullCoList/"
    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": (
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15"
            ),
            "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
        }
    )

    logger.info("Запрашиваю страницу списка кредитных организаций ЦБ: %s", page_url)
    response = session.get(page_url, timeout=timeout)
    response.raise_for_status()
    response.encoding = response.encoding or "utf-8"
    html = response.text

    banks: List[BankRef] = []
    xlsx_url = find_full_list_xlsx_url(html, page_url)
    if xlsx_url:
        logger.info("Скачиваю XLSX списка кредитных организаций ЦБ: %s", xlsx_url)
        try:
            xlsx_response = session.get(xlsx_url, timeout=timeout)
            xlsx_response.raise_for_status()
            with warnings.catch_warnings():
                warnings.filterwarnings("ignore", message="Workbook contains no default style.*")
                sheets = pd.read_excel(io.BytesIO(xlsx_response.content), sheet_name=None)
            for df in sheets.values():
                banks = merge_bank_refs(banks, bank_refs_from_full_list_frame(df))
            logger.info("XLSX списка ЦБ обработан: листов=%d, банков=%d", len(sheets), len(banks))
        except Exception as exc:
            logger.warning("Не удалось обработать XLSX списка ЦБ, пробую HTML-таблицу: %s", exc)

    if not banks:
        logger.info("Разбираю HTML-таблицу списка кредитных организаций ЦБ")
        try:
            tables = pd.read_html(io.StringIO(html), flavor="lxml")
        except Exception:
            tables = []
        for table in tables:
            banks = merge_bank_refs(banks, bank_refs_from_full_list_frame(table))

    result = sorted(banks, key=lambda item: int(item.regnum))
    if not result:
        raise RuntimeError("Не удалось найти действующие банки на странице ЦБ. Передайте список через --regnums-file.")
    logger.info("Найдено действующих банков на странице ЦБ: %d", len(result))
    return result


def sqlite_table_name(query_name: str, used: Dict[str, str]) -> str:
    base = "data_" + re.sub(r"\W+", "_", query_name, flags=re.UNICODE).strip("_")
    if not base or base == "data_":
        base = "data_query"
    if re.match(r"^\d", base):
        base = "data_" + base
    table = base
    suffix = 2
    while table in used and used[table] != query_name:
        table = f"{base}_{suffix}"
        suffix += 1
    used[table] = query_name
    return table


def ensure_meta_tables(conn: sqlite3.Connection) -> None:
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS export_banks (
            regnum TEXT PRIMARY KEY,
            bank_name TEXT,
            status TEXT NOT NULL,
            started_at TEXT NOT NULL,
            finished_at TEXT,
            error TEXT
        );

        CREATE TABLE IF NOT EXISTS export_queries (
            query_name TEXT PRIMARY KEY,
            sqlite_table TEXT NOT NULL,
            sheet_name TEXT,
            excel_table_name TEXT,
            excel_table_ref TEXT,
            columns_json TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS export_errors (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            regnum TEXT NOT NULL,
            bank_name TEXT,
            query_name TEXT,
            error TEXT NOT NULL,
            traceback TEXT NOT NULL,
            created_at TEXT NOT NULL
        );
        """
    )


def sqlite_now(conn: sqlite3.Connection) -> str:
    return conn.execute("SELECT strftime('%Y-%m-%dT%H:%M:%SZ', 'now')").fetchone()[0]


def init_database(
    db_path: Path,
    if_exists: str,
    loaded_queries: Sequence,
    query_tables: Dict[str, str],
) -> sqlite3.Connection:
    if db_path.exists() and if_exists == "fail":
        raise FileExistsError(f"SQLite-файл уже существует: {db_path}. Используйте --replace или --append.")
    if db_path.exists() and if_exists == "replace":
        db_path.unlink()

    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    ensure_meta_tables(conn)

    if if_exists == "replace":
        for table in query_tables.values():
            conn.execute(f'DROP TABLE IF EXISTS "{table}"')
    for item in loaded_queries:
        conn.execute(
            """
            INSERT OR REPLACE INTO export_queries
                (query_name, sqlite_table, sheet_name, excel_table_name, excel_table_ref, columns_json)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                item.query_name,
                query_tables[item.query_name],
                item.sheet_name,
                item.table_name,
                item.table_ref,
                json.dumps(item.columns, ensure_ascii=False),
            ),
        )
    conn.commit()
    return conn


def write_query_frame(
    conn: sqlite3.Connection,
    *,
    table: str,
    regnum: str,
    bank_name: str,
    query_name: str,
    df: pd.DataFrame,
) -> None:
    out = df.copy()
    out.insert(0, "row_number", range(1, len(out) + 1))
    out.insert(0, "collected_at", sqlite_now(conn))
    out.insert(0, "query_name", query_name)
    out.insert(0, "bank_name", bank_name)
    out.insert(0, "regnum", regnum)
    out.to_sql(table, conn, if_exists="append", index=False)


def mark_bank_started(conn: sqlite3.Connection, bank: BankRef) -> None:
    started_at = sqlite_now(conn)
    conn.execute(
        """
        INSERT OR REPLACE INTO export_banks
            (regnum, bank_name, status, started_at, finished_at, error)
        VALUES (?, ?, 'running', ?, NULL, NULL)
        """,
        (bank.regnum, bank.name, started_at),
    )
    conn.commit()


def collect_bank_data(
    *,
    bank: BankRef,
    base_m_code: str,
    loaded_queries: Sequence,
    query_tables: Dict[str, str],
    selected_queries: Optional[set],
    cache_dir: Path,
    use_cache: bool,
    timeout: int,
    verbose: bool,
    logger: logging.Logger,
) -> BankExportData:
    m_code = replace_regnum_in_m_code(base_m_code, bank.regnum)
    engine = MiniMEngine(
        m_code,
        cache_dir=cache_dir,
        use_cache=use_cache,
        timeout=timeout,
        logger=logger,
        verbose=verbose,
    )

    frames: List[QueryFrame] = []
    rows_written = 0
    query_count = 0
    for item in loaded_queries:
        if selected_queries is not None and item.query_name not in selected_queries:
            continue
        df = engine.evaluate_query(item.query_name)
        frames.append(QueryFrame(query_name=item.query_name, table=query_tables[item.query_name], df=df))
        rows_written += len(df)
        query_count += 1
        logger.info(
            "regnum=%s query=%s получено строк=%d %s",
            bank.regnum,
            item.query_name,
            len(df),
            df_debug_summary(df),
        )
    return BankExportData(bank=bank, frames=frames, query_count=query_count, rows_written=rows_written)


def export_bank(
    *,
    conn: sqlite3.Connection,
    bank: BankRef,
    base_m_code: str,
    loaded_queries: Sequence,
    query_tables: Dict[str, str],
    selected_queries: Optional[set],
    cache_dir: Path,
    use_cache: bool,
    timeout: int,
    verbose: bool,
    logger: logging.Logger,
) -> Tuple[int, int]:
    mark_bank_started(conn, bank)
    data = collect_bank_data(
        bank=bank,
        base_m_code=base_m_code,
        loaded_queries=loaded_queries,
        query_tables=query_tables,
        selected_queries=selected_queries,
        cache_dir=cache_dir,
        use_cache=use_cache,
        timeout=timeout,
        verbose=verbose,
        logger=logger,
    )
    for frame in data.frames:
        try:
            write_query_frame(
                conn,
                table=frame.table,
                regnum=bank.regnum,
                bank_name=bank.name,
                query_name=frame.query_name,
                df=frame.df,
            )
            logger.info(
                "regnum=%s query=%s записано строк=%d %s",
                bank.regnum,
                frame.query_name,
                len(frame.df),
                df_debug_summary(frame.df),
            )
        except Exception as exc:
            conn.execute(
                """
                INSERT INTO export_errors (regnum, bank_name, query_name, error, traceback, created_at)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (bank.regnum, bank.name, frame.query_name, str(exc), traceback.format_exc(), sqlite_now(conn)),
            )
            raise
    conn.execute(
        "UPDATE export_banks SET status='ok', finished_at=?, error=NULL WHERE regnum=?",
        (sqlite_now(conn), bank.regnum),
    )
    conn.commit()
    return data.query_count, data.rows_written


def write_bank_data(conn: sqlite3.Connection, data: BankExportData, logger: logging.Logger) -> None:
    for frame in data.frames:
        write_query_frame(
            conn,
            table=frame.table,
            regnum=data.bank.regnum,
            bank_name=data.bank.name,
            query_name=frame.query_name,
            df=frame.df,
        )
        logger.info(
            "regnum=%s query=%s записано строк=%d %s",
            data.bank.regnum,
            frame.query_name,
            len(frame.df),
            df_debug_summary(frame.df),
        )
    conn.execute(
        "UPDATE export_banks SET status='ok', finished_at=?, error=NULL WHERE regnum=?",
        (sqlite_now(conn), data.bank.regnum),
    )
    conn.commit()


def parse_args(argv: Optional[List[str]]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Собрать данные ЦБ из Power Query книги в SQLite по списку regnum.")
    parser.add_argument("xlsx", type=Path, help="Исходный .xlsx файл с Power Query")
    parser.add_argument("-o", "--output", type=Path, default=Path("cbr_banks.sqlite"), help="SQLite-файл результата")
    parser.add_argument("--regnum", action="append", default=[], help="Один regnum. Можно передать несколько раз")
    parser.add_argument("--regnums-file", type=Path, default=None, help="TXT/CSV со списком regnum; второй столбец может быть названием банка")
    parser.add_argument("--all-banks", action="store_true", help="Получить с сайта ЦБ список банков с действующими лицензиями; НКО/НБКО исключаются")
    parser.add_argument("--limit", type=int, default=None, help="Ограничить количество банков для тестового прогона")
    parser.add_argument("--query", action="append", default=[], help="Выполнять только указанный Power Query. Можно передать несколько раз")
    parser.add_argument("--replace", action="store_true", help="Пересоздать SQLite-файл, если он уже существует")
    parser.add_argument("--append", action="store_true", help="Дописать данные в существующий SQLite-файл")
    parser.add_argument("--cache-dir", type=Path, default=None, help="Папка HTML-кэша")
    parser.add_argument("--no-cache", action="store_true", help="Не использовать HTML-кэш")
    parser.add_argument("--timeout", type=int, default=60, help="HTTP timeout в секундах")
    parser.add_argument("--workers", type=int, default=1, help="Количество параллельных загрузчиков банков. 1 = последовательный режим")
    parser.add_argument("--fail-fast", action="store_true", help="Остановиться на первой ошибке банка/запроса")
    parser.add_argument("--verbose", action="store_true", help="Подробный вывод")
    return parser.parse_args(argv)


def main(argv: Optional[List[str]] = None) -> int:
    args = parse_args(argv)
    if args.replace and args.append:
        raise SystemExit("Нельзя одновременно использовать --replace и --append.")
    if args.workers < 1:
        raise SystemExit("--workers должен быть не меньше 1.")

    logger = setup_logging(args.verbose)
    xlsx_path = args.xlsx.resolve()
    db_path = args.output.resolve()
    cache_dir = (args.cache_dir or xlsx_path.with_name("pq_html_cache")).resolve()

    banks = merge_bank_refs(BankRef(regnum=r.strip()) for r in args.regnum if r.strip())
    if args.regnums_file is not None:
        banks = merge_bank_refs(banks, parse_bank_refs_from_file(args.regnums_file))
    if args.all_banks:
        banks = merge_bank_refs(banks, discover_bank_refs_from_cbr(args.timeout, logger))
    if not banks:
        raise SystemExit("Передайте хотя бы один --regnum, --regnums-file или --all-banks.")
    if args.limit is not None:
        banks = banks[: max(args.limit, 0)]
    bad_regnums = [bank.regnum for bank in banks if not re.fullmatch(r"\d+", bank.regnum)]
    if bad_regnums:
        raise SystemExit(f"Некорректные regnum: {', '.join(bad_regnums[:10])}")

    inspector = PowerQueryWorkbookInspector(xlsx_path)
    base_m_code = inspector.extract_m_code()
    loaded_queries = inspector.list_loaded_queries()
    if not loaded_queries:
        raise SystemExit("В книге не найдены загружаемые Power Query таблицы.")

    selected_queries = set(args.query) if args.query else None
    if selected_queries is not None:
        known_queries = {item.query_name for item in loaded_queries}
        unknown = sorted(selected_queries - known_queries)
        if unknown:
            raise SystemExit(f"В книге нет таких загружаемых запросов: {', '.join(unknown)}")

    used_tables: Dict[str, str] = {}
    query_tables = {item.query_name: sqlite_table_name(item.query_name, used_tables) for item in loaded_queries}
    if_exists = "append" if args.append else ("replace" if args.replace else "fail")
    conn = init_database(db_path, if_exists, loaded_queries, query_tables)

    logger.info("SQLite: %s", db_path)
    logger.info("Банков к обработке: %d", len(banks))
    logger.info("Загружаемых запросов в книге: %d", len(loaded_queries))
    logger.info("Параллельных загрузчиков: %d", min(args.workers, len(banks)))

    ok_count = 0
    failed_count = 0
    total_rows = 0
    try:
        if args.workers == 1 or len(banks) <= 1:
            for idx, bank in enumerate(banks, start=1):
                logger.info("[%d/%d] regnum=%s %s", idx, len(banks), bank.regnum, bank.name)
                try:
                    query_count, rows_written = export_bank(
                        conn=conn,
                        bank=bank,
                        base_m_code=base_m_code,
                        loaded_queries=loaded_queries,
                        query_tables=query_tables,
                        selected_queries=selected_queries,
                        cache_dir=cache_dir,
                        use_cache=not args.no_cache,
                        timeout=args.timeout,
                        verbose=args.verbose,
                        logger=logger,
                    )
                    ok_count += 1
                    total_rows += rows_written
                    logger.info("regnum=%s готов: запросов=%d строк=%d", bank.regnum, query_count, rows_written)
                except Exception as exc:
                    failed_count += 1
                    conn.execute(
                        "UPDATE export_banks SET status='error', finished_at=?, error=? WHERE regnum=?",
                        (sqlite_now(conn), str(exc), bank.regnum),
                    )
                    conn.commit()
                    logger.error("regnum=%s ошибка: %s", bank.regnum, exc)
                    logger.debug(traceback.format_exc())
                    if args.fail_fast:
                        raise
                time.sleep(0.05)
        else:
            worker_count = min(args.workers, len(banks))
            with ThreadPoolExecutor(max_workers=worker_count) as executor:
                futures = {}
                for idx, bank in enumerate(banks, start=1):
                    logger.info("[%d/%d] regnum=%s %s", idx, len(banks), bank.regnum, bank.name)
                    mark_bank_started(conn, bank)
                    future = executor.submit(
                        collect_bank_data,
                        bank=bank,
                        base_m_code=base_m_code,
                        loaded_queries=loaded_queries,
                        query_tables=query_tables,
                        selected_queries=selected_queries,
                        cache_dir=cache_dir,
                        use_cache=not args.no_cache,
                        timeout=args.timeout,
                        verbose=args.verbose,
                        logger=logger,
                    )
                    futures[future] = bank

                for done_no, future in enumerate(as_completed(futures), start=1):
                    bank = futures[future]
                    try:
                        data = future.result()
                        write_bank_data(conn, data, logger)
                        ok_count += 1
                        total_rows += data.rows_written
                        logger.info(
                            "[%d/%d] regnum=%s готов: запросов=%d строк=%d",
                            done_no,
                            len(banks),
                            bank.regnum,
                            data.query_count,
                            data.rows_written,
                        )
                    except Exception as exc:
                        failed_count += 1
                        conn.execute(
                            "UPDATE export_banks SET status='error', finished_at=?, error=? WHERE regnum=?",
                            (sqlite_now(conn), str(exc), bank.regnum),
                        )
                        conn.commit()
                        logger.error("[%d/%d] regnum=%s ошибка: %s", done_no, len(banks), bank.regnum, exc)
                        logger.debug(traceback.format_exc())
                        if args.fail_fast:
                            for pending in futures:
                                pending.cancel()
                            raise
                    finally:
                        time.sleep(0.05)
    finally:
        conn.close()

    print(f"Готово: {db_path}")
    print(f"Банков успешно: {ok_count}; с ошибками: {failed_count}; строк записано: {total_rows}")
    return 0 if failed_count == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
