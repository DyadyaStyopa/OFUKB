#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Python-копия загрузки Power Query из книги Excel.

Версия: добавлена поддержка выбора банка через --regnum; добавлены --verbose/--debug; запись Excel выполнена напрямую через ZIP/XML без сохранения книги через openpyxl.

Что делает скрипт:
1) Достаёт M-код Power Query из customXml/item1.xml внутри .xlsx.
2) Находит запросы, которые реально загружаются на листы Excel как таблицы.
3) Выполняет используемое в книге подмножество M-операций на Python.
4) Записывает результаты в те же листы и диапазоны таблиц Excel.

Зависимости:
    pip install pandas requests beautifulsoup4 lxml

Запуск:
    python wtf2_xmlsafe_cellonly.py "ОФУКБ_АО_ТБанк_2026_вер 2.xlsx"

Выбор банка по регистрационному номеру ЦБ:
    python wtf2_xmlsafe_cellonly.py "...xlsx" --regnum 1481

Полезные режимы:
    python wtf2_xmlsafe_cellonly.py "...xlsx" --list       # только показать найденные загрузки
    python wtf2_xmlsafe_cellonly.py "...xlsx" --dump-m     # сохранить извлечённый M-код рядом с файлом
    python wtf2_xmlsafe_cellonly.py "...xlsx" --no-cache   # не использовать HTML-кэш

Ограничение: это не полноценный интерпретатор M, а реализация именно того набора
операций, который обнаружен в этой книге: Web.Page/Web.Contents, Web.BrowserContents,
Html.Table, PromoteHeaders, TransformColumnTypes, ReplaceErrorValues, RemoveColumns,
RenameColumns, Skip, NestedJoin, ExpandTableColumn, ReorderColumns, TransformColumns.
"""

from __future__ import annotations

import argparse
import copy
import base64
import hashlib
import io
import json
import logging
import os
import re
import sys
import time
import traceback
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple
from xml.etree import ElementTree as ET

import pandas as pd
import requests
from bs4 import BeautifulSoup

# -----------------------------------------------------------------------------
# Логирование и диагностика
# -----------------------------------------------------------------------------

class TeeStdout:
    """Пишет stdout/stderr одновременно в консоль и в файл лога."""

    def __init__(self, stream: Any, log_path: Path):
        self.stream = stream
        self.log_path = Path(log_path)
        self.file = self.log_path.open("a", encoding="utf-8", buffering=1)

    def write(self, data: str) -> int:
        self.stream.write(data)
        self.file.write(data)
        return len(data)

    def flush(self) -> None:
        self.stream.flush()
        self.file.flush()

    def isatty(self) -> bool:
        return getattr(self.stream, "isatty", lambda: False)()


def safe_filename(name: str, max_len: int = 120) -> str:
    """Безопасное имя файла для debug-артефактов."""
    name = re.sub(r"[^0-9A-Za-zА-Яа-яЁё._ -]+", "_", str(name)).strip(" ._")
    return (name[:max_len] or "unnamed")


def setup_run_logging(
    *,
    xlsx_path: Path,
    log_file: Optional[Path],
    verbose: bool,
    debug: bool,
) -> Tuple[logging.Logger, Optional[Path]]:
    """Настраивает логгер. При --verbose/--debug обязательно пишет текстовый лог."""
    logger = logging.getLogger("pq2py")
    logger.handlers.clear()
    logger.propagate = False
    logger.setLevel(logging.DEBUG)

    fmt = logging.Formatter("%(asctime)s | %(levelname)-7s | %(message)s", "%Y-%m-%d %H:%M:%S")

    console = logging.StreamHandler(sys.__stdout__)
    console.setFormatter(fmt)
    console.setLevel(logging.DEBUG if debug else (logging.INFO if verbose else logging.WARNING))
    logger.addHandler(console)

    effective_log_file: Optional[Path] = log_file
    if effective_log_file is None and (verbose or debug):
        effective_log_file = xlsx_path.with_name(xlsx_path.stem + "_run_log.txt")
    if effective_log_file is not None:
        effective_log_file = effective_log_file.resolve()
        effective_log_file.parent.mkdir(parents=True, exist_ok=True)
        file_handler = logging.FileHandler(effective_log_file, mode="w", encoding="utf-8")
        file_handler.setFormatter(fmt)
        file_handler.setLevel(logging.DEBUG if debug else logging.INFO)
        logger.addHandler(file_handler)
        logger.info("Лог выполнения: %s", effective_log_file)

    return logger, effective_log_file


def df_debug_summary(df: pd.DataFrame, max_cols: int = 25) -> str:
    """Краткое описание DataFrame для текстового лога."""
    cols = [str(c) for c in df.columns]
    if len(cols) > max_cols:
        cols = cols[:max_cols] + [f"... +{len(df.columns)-max_cols} колонок"]
    return f"shape={df.shape}; columns={cols}"


def write_df_debug_artifacts(debug_dir: Optional[Path], query_name: str, step_name: str, df: pd.DataFrame) -> None:
    """Сохраняет маленькие CSV/JSON-снимки DataFrame при --debug."""
    if debug_dir is None:
        return
    q_dir = debug_dir / safe_filename(query_name)
    q_dir.mkdir(parents=True, exist_ok=True)
    base = safe_filename(step_name)
    schema = {
        "query": query_name,
        "step": step_name,
        "rows": int(len(df)),
        "columns_count": int(len(df.columns)),
        "columns": [str(c) for c in df.columns],
        "dtypes": {str(k): str(v) for k, v in df.dtypes.items()},
    }
    (q_dir / f"{base}.schema.json").write_text(json.dumps(schema, ensure_ascii=False, indent=2), encoding="utf-8")
    df.head(30).to_csv(q_dir / f"{base}.head.csv", index=False, encoding="utf-8-sig")


def xlsx_zip_diagnostics(xlsx_path: Path, out_txt: Path) -> None:
    """Пишет диагностический отчёт по внутренним XML .xlsx."""
    lines: List[str] = []
    lines.append(f"Файл: {xlsx_path}")
    try:
        with zipfile.ZipFile(xlsx_path) as zf:
            bad = zf.testzip()
            lines.append(f"zip testzip: {bad or 'OK'}")
            names = zf.namelist()
            lines.append(f"entries: {len(names)}")
            for name in names:
                if name.startswith("xl/tables/table") and name.endswith(".xml"):
                    xml = zf.read(name).decode("utf-8", errors="replace")
                    ref_match = re.search(r'ref="([^"]+)"', xml)
                    count_match = re.search(r'<tableColumns count="(\d+)"', xml)
                    table_name = re.search(r'name="([^"]+)"', xml)
                    ref = ref_match.group(1) if ref_match else "?"
                    cnt = count_match.group(1) if count_match else "?"
                    actual = len(re.findall(r'<tableColumn\b', xml))
                    lines.append(f"{name}: name={table_name.group(1) if table_name else '?'} ref={ref} tableColumns_count={cnt} actual_tableColumn={actual}")
                    for col_match in re.finditer(r'<tableColumn\b[^>]*>', xml):
                        tag = col_match.group(0)
                        cm_id = re.search(r'id="([^"]+)"', tag)
                        cm_name = re.search(r'name="([^"]*)"', tag)
                        cm_uid = re.search(r'uniqueName="([^"]*)"', tag)
                        cm_qid = re.search(r'queryTableFieldId="([^"]*)"', tag)
                        lines.append(
                            "    column "
                            f"id={cm_id.group(1) if cm_id else '?'} "
                            f"name={cm_name.group(1) if cm_name else '?'} "
                            f"uniqueName={cm_uid.group(1) if cm_uid else '-'} "
                            f"queryTableFieldId={cm_qid.group(1) if cm_qid else '-'}"
                        )
    except Exception:
        lines.append("ОШИБКА диагностики ZIP/XML:")
        lines.append(traceback.format_exc())
    out_txt.write_text("\n".join(lines), encoding="utf-8")

# -----------------------------------------------------------------------------
# XML namespaces
# -----------------------------------------------------------------------------

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"m": NS_MAIN}


@dataclass
class LoadedQuery:
    """Куда Power Query загружает результат в книге."""

    query_name: str
    sheet_name: str
    table_name: str
    table_ref: str
    connection_id: str
    columns: List[str]


class PowerQueryWorkbookInspector:
    """Извлекает из .xlsx M-код и карту загрузок Query -> Excel Table."""

    def __init__(self, xlsx_path: Path):
        self.xlsx_path = Path(xlsx_path)
        if not self.xlsx_path.exists():
            raise FileNotFoundError(self.xlsx_path)

    def extract_m_code(self) -> str:
        """Достаёт Formulas/Section1.m из DataMashup."""
        with zipfile.ZipFile(self.xlsx_path) as zf:
            # В этой книге Power Query хранится в customXml/item1.xml.
            candidates = [n for n in zf.namelist() if n.startswith("customXml/item") and n.endswith(".xml")]
            last_error: Optional[Exception] = None
            for name in candidates:
                raw = zf.read(name)
                for encoding in ("utf-16", "utf-8-sig", "utf-8"):
                    try:
                        text = raw.decode(encoding)
                        root = ET.fromstring(text)
                        if "DataMashup" not in root.tag:
                            continue
                        payload = "".join(root.itertext()).strip()
                        mashup = base64.b64decode(payload)
                        package_len = int.from_bytes(mashup[4:8], "little")
                        package = mashup[8 : 8 + package_len]
                        with zipfile.ZipFile(io.BytesIO(package)) as mz:
                            return mz.read("Formulas/Section1.m").decode("utf-8-sig")
                    except Exception as exc:  # пробуем следующий item/encoding
                        last_error = exc
                        continue
            raise RuntimeError(f"Не удалось извлечь M-код из DataMashup: {last_error}")

    def list_loaded_queries(self) -> List[LoadedQuery]:
        """Находит таблицы Excel, связанные с запросами Power Query."""
        with zipfile.ZipFile(self.xlsx_path) as zf:
            workbook = ET.fromstring(zf.read("xl/workbook.xml"))
            wb_rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
            rid_to_target = {rel.attrib["Id"]: rel.attrib["Target"] for rel in wb_rels}

            sheet_paths: List[Tuple[str, str]] = []
            for sh in workbook.find("m:sheets", NS):
                sheet_name = sh.attrib["name"]
                rid = sh.attrib[f"{{{NS_REL}}}id"]
                target = rid_to_target[rid]
                sheet_path = target.lstrip("/") if target.startswith("/") else "xl/" + target
                sheet_paths.append((sheet_name, sheet_path))

            connections = self._read_connections(zf)
            table_info = self._read_tables(zf)

            loaded: List[LoadedQuery] = []
            for sheet_name, sheet_path in sheet_paths:
                if sheet_path not in zf.namelist():
                    continue
                sheet_root = ET.fromstring(zf.read(sheet_path))
                rel_path = sheet_path.replace("worksheets/", "worksheets/_rels/") + ".rels"
                sheet_rels: Dict[str, str] = {}
                if rel_path in zf.namelist():
                    root_rels = ET.fromstring(zf.read(rel_path))
                    for rel in root_rels:
                        target = rel.attrib["Target"]
                        if target.startswith("../"):
                            normalized = "xl/" + target[3:]
                        elif target.startswith("/"):
                            normalized = target.lstrip("/")
                        else:
                            normalized = str(Path(sheet_path).parent / target).replace("\\", "/")
                        sheet_rels[rel.attrib["Id"]] = normalized

                table_parts = sheet_root.find("m:tableParts", NS)
                if table_parts is None:
                    continue
                for part in table_parts:
                    rid = part.attrib[f"{{{NS_REL}}}id"]
                    table_path = sheet_rels.get(rid)
                    info = table_info.get(table_path or "")
                    if not info:
                        continue
                    conn_id = info.get("connection_id") or ""
                    conn = connections.get(conn_id, {})
                    query_name = self._query_name_from_connection(conn)
                    if not query_name:
                        continue
                    loaded.append(
                        LoadedQuery(
                            query_name=query_name,
                            sheet_name=sheet_name,
                            table_name=info["table_name"],
                            table_ref=info["ref"],
                            connection_id=conn_id,
                            columns=info["columns"],
                        )
                    )
            return loaded

    def _read_connections(self, zf: zipfile.ZipFile) -> Dict[str, Dict[str, Any]]:
        if "xl/connections.xml" not in zf.namelist():
            return {}
        root = ET.fromstring(zf.read("xl/connections.xml"))
        result: Dict[str, Dict[str, Any]] = {}
        for conn in root:
            cid = conn.attrib.get("id", "")
            item: Dict[str, Any] = dict(conn.attrib)
            for child in conn:
                tag = child.tag.split("}")[-1]
                item[tag] = dict(child.attrib)
            result[cid] = item
        return result

    def _read_tables(self, zf: zipfile.ZipFile) -> Dict[str, Dict[str, Any]]:
        tables: Dict[str, Dict[str, Any]] = {}
        for name in zf.namelist():
            if not re.fullmatch(r"xl/tables/table\d+\.xml", name):
                continue
            root = ET.fromstring(zf.read(name))
            columns = []
            tcols = root.find("m:tableColumns", NS)
            if tcols is not None:
                columns = [c.attrib.get("name", "") for c in tcols]

            rel_path = "xl/tables/_rels/" + Path(name).name + ".rels"
            query_table_path: Optional[str] = None
            connection_id: Optional[str] = None
            if rel_path in zf.namelist():
                rels = ET.fromstring(zf.read(rel_path))
                for rel in rels:
                    target = rel.attrib.get("Target", "")
                    if "queryTable" not in target:
                        continue
                    if target.startswith("../"):
                        query_table_path = "xl/" + target[3:]
                    elif target.startswith("/"):
                        query_table_path = target.lstrip("/")
                    else:
                        query_table_path = str(Path(name).parent / target).replace("\\", "/")
                    break
            if query_table_path and query_table_path in zf.namelist():
                qt = ET.fromstring(zf.read(query_table_path))
                connection_id = qt.attrib.get("connectionId")

            tables[name] = {
                "table_name": root.attrib.get("displayName") or root.attrib.get("name") or Path(name).stem,
                "ref": root.attrib.get("ref", "A1"),
                "columns": columns,
                "connection_id": connection_id,
            }
        return tables

    @staticmethod
    def _query_name_from_connection(conn: Dict[str, Any]) -> str:
        dbpr = conn.get("dbPr") or {}
        connection = dbpr.get("connection", "")
        # Provider=...;Location=Ф_102_2021;Extended Properties=""
        m = re.search(r"(?:^|;)Location=([^;]+)", connection)
        if m:
            return m.group(1)
        command = dbpr.get("command", "")
        m = re.search(r"SELECT \* FROM \[([^\]]+)\]", command)
        if m:
            return m.group(1)
        name = conn.get("name", "")
        return name.replace("Запрос — ", "").strip()


# -----------------------------------------------------------------------------
# Мини-исполнитель M-кода для нужного подмножества Power Query
# -----------------------------------------------------------------------------


# -----------------------------------------------------------------------------
# Подстановка регистрационного номера банка в URL ЦБ
# -----------------------------------------------------------------------------

def detect_regnums_in_m_code(m_code: str) -> List[str]:
    """Возвращает все regnum, найденные в URL внутри M-кода."""
    return sorted(set(re.findall(r"[?&]regnum=(\d+)", m_code)), key=lambda x: int(x))


def replace_regnum_in_m_code(m_code: str, regnum: str) -> str:
    """
    Заменяет параметр regnum во всех URL ЦБ, сохраняя остальные параметры и даты.

    Пример:
        ...?regnum=2673&dt=2024-01-01 -> ...?regnum=1481&dt=2024-01-01
        ...?when=201701&regnum=2673&view=0409806 -> ...?when=201701&regnum=1481&view=0409806
    """
    regnum = str(regnum).strip()
    if not re.fullmatch(r"\d+", regnum):
        raise ValueError("regnum должен состоять только из цифр, например: 2673 или 1481")
    return re.sub(r"([?&]regnum=)\d+", lambda m: m.group(1) + regnum, m_code)


class MiniMEngine:
    def __init__(
        self,
        m_code: str,
        cache_dir: Path,
        use_cache: bool = True,
        timeout: int = 60,
        logger: Optional[logging.Logger] = None,
        verbose: bool = False,
        debug: bool = False,
        debug_dir: Optional[Path] = None,
    ):
        self.queries = self._split_queries(m_code)
        self.cache: Dict[str, pd.DataFrame] = {}
        self.cache_dir = Path(cache_dir)
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.use_cache = use_cache
        self.timeout = timeout
        self.logger = logger or logging.getLogger("pq2py")
        self.verbose = verbose
        self.debug = debug
        self.debug_dir = Path(debug_dir) if debug_dir is not None else None
        self._eval_depth = 0
        self._current_query = ""
        self.session = requests.Session()
        self.session.headers.update(
            {
                "User-Agent": (
                    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                    "AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15"
                ),
                "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
            }
        )
        self.logger.debug(
            "MiniMEngine initialized: queries=%d, cache_dir=%s, use_cache=%s, timeout=%s, debug_dir=%s",
            len(self.queries), self.cache_dir, self.use_cache, self.timeout, self.debug_dir,
        )

    @staticmethod
    def _normalize_identifier(raw: str) -> str:
        raw = raw.strip().rstrip(",;")
        if raw.startswith('#"') and raw.endswith('"'):
            return raw[2:-1].replace('""', '"')
        return raw

    def _split_queries(self, code: str) -> Dict[str, str]:
        positions = [(m.start(), m.group(1)) for m in re.finditer(r"(?m)^shared\s+(.+?)\s*=\s*let", code)]
        result: Dict[str, str] = {}
        for i, (start, raw_name) in enumerate(positions):
            end = positions[i + 1][0] if i + 1 < len(positions) else len(code)
            name = self._normalize_identifier(raw_name)
            result[name] = code[start:end]
        return result

    def evaluate_query(self, query_name: str) -> pd.DataFrame:
        query_name = self._normalize_identifier(query_name)
        if query_name in self.cache:
            self.logger.debug("CACHE query result: %s -> %s", query_name, df_debug_summary(self.cache[query_name]))
            return self.cache[query_name].copy()
        if query_name not in self.queries:
            raise KeyError(f"Запрос M не найден: {query_name}")

        prev_query = self._current_query
        self._current_query = query_name
        self._eval_depth += 1
        indent = "  " * (self._eval_depth - 1)
        self.logger.info("%sM query START: %s", indent, query_name)
        try:
            code = self.queries[query_name]
            steps, final_ref = self._parse_query_steps(code)
            self.logger.debug("%sM query parsed: %s steps=%d final_ref=%s", indent, query_name, len(steps), final_ref)
            if self.debug_dir is not None:
                q_dir = self.debug_dir / safe_filename(query_name)
                q_dir.mkdir(parents=True, exist_ok=True)
                (q_dir / "query.m").write_text(code, encoding="utf-8")
                (q_dir / "steps.json").write_text(
                    json.dumps([{"step": n, "expression": e} for n, e in steps], ensure_ascii=False, indent=2),
                    encoding="utf-8",
                )

            context: Dict[str, Any] = {}
            for step_no, (step_name, expression) in enumerate(steps, start=1):
                self.logger.debug("%sSTEP %03d %s.%s expr=%s", indent, step_no, query_name, step_name, expression[:1000])
                t0 = time.perf_counter()
                try:
                    value = self._eval_expression(expression, context)
                except Exception:
                    self.logger.error("%sERROR in step %s.%s", indent, query_name, step_name)
                    self.logger.error("%sExpression: %s", indent, expression)
                    self.logger.error(traceback.format_exc())
                    raise
                context[step_name] = value
                elapsed = time.perf_counter() - t0
                if isinstance(value, pd.DataFrame):
                    self.logger.debug("%sSTEP OK %s.%s: %s; %.3fs", indent, query_name, step_name, df_debug_summary(value), elapsed)
                    write_df_debug_artifacts(self.debug_dir, query_name, f"{step_no:03d}_{step_name}", value)
                else:
                    self.logger.debug("%sSTEP OK %s.%s: type=%s; %.3fs", indent, query_name, step_name, type(value).__name__, elapsed)
            final_name = self._normalize_identifier(final_ref)
            if final_name in context:
                df = self._ensure_df(context[final_name])
            else:
                df = self.evaluate_query(final_name)
            self.cache[query_name] = df.copy()
            self.logger.info("%sM query END: %s -> %s", indent, query_name, df_debug_summary(df))
            write_df_debug_artifacts(self.debug_dir, query_name, "999_final", df)
            return df.copy()
        finally:
            self._eval_depth -= 1
            self._current_query = prev_query

    def _parse_query_steps(self, code: str) -> Tuple[List[Tuple[str, str]], str]:
        let_pos = code.find("let")
        in_match = re.search(r"(?m)^in\s*$", code)
        if let_pos < 0 or not in_match:
            raise ValueError("Не удалось разобрать let/in в запросе")
        body = code[let_pos + 3 : in_match.start()].strip()
        after_in = code[in_match.end() :].strip().rstrip(";").strip()
        final_ref = after_in.splitlines()[0].strip() if after_in else ""

        chunks = self._split_top_level_commas(body)
        steps: List[Tuple[str, str]] = []
        for chunk in chunks:
            if not chunk.strip():
                continue
            name, expr = self._split_assignment(chunk)
            steps.append((self._normalize_identifier(name), expr.strip()))
        return steps, final_ref

    @staticmethod
    def _split_top_level_commas(text: str) -> List[str]:
        chunks: List[str] = []
        start = 0
        depth_par = depth_brace = depth_bracket = 0
        in_string = False
        i = 0
        while i < len(text):
            ch = text[i]
            if in_string:
                if ch == '"':
                    if i + 1 < len(text) and text[i + 1] == '"':
                        i += 2
                        continue
                    in_string = False
                i += 1
                continue
            if ch == '"':
                in_string = True
            elif ch == '(':
                depth_par += 1
            elif ch == ')':
                depth_par -= 1
            elif ch == '{':
                depth_brace += 1
            elif ch == '}':
                depth_brace -= 1
            elif ch == '[':
                depth_bracket += 1
            elif ch == ']':
                depth_bracket -= 1
            elif ch == ',' and depth_par == depth_brace == depth_bracket == 0:
                chunks.append(text[start:i].strip())
                start = i + 1
            i += 1
        tail = text[start:].strip()
        if tail:
            chunks.append(tail)
        return chunks

    @staticmethod
    def _split_assignment(chunk: str) -> Tuple[str, str]:
        in_string = False
        depth_par = depth_brace = depth_bracket = 0
        for i, ch in enumerate(chunk):
            if in_string:
                if ch == '"':
                    if i + 1 < len(chunk) and chunk[i + 1] == '"':
                        continue
                    in_string = False
                continue
            if ch == '"':
                in_string = True
            elif ch == '(':
                depth_par += 1
            elif ch == ')':
                depth_par -= 1
            elif ch == '{':
                depth_brace += 1
            elif ch == '}':
                depth_brace -= 1
            elif ch == '[':
                depth_bracket += 1
            elif ch == ']':
                depth_bracket -= 1
            elif ch == '=' and depth_par == depth_brace == depth_bracket == 0:
                return chunk[:i].strip(), chunk[i + 1 :].strip()
        raise ValueError(f"Не найдено присваивание в шаге M: {chunk[:100]}")

    def _eval_expression(self, expr: str, ctx: Dict[str, Any]) -> Any:
        expr = expr.strip().rstrip(",")

        # Web.BrowserContents("url") -> HTML-строка
        m = re.search(r"Web\.BrowserContents\(\"([^\"]+)\"\)", expr)
        if m:
            return self._fetch_html(m.group(1))

        # Web.Page(Web.Contents("url")) -> список таблиц; для этой книги нужен Источник{0}[Data]
        m = re.search(r"Web\.Page\(\s*Web\.Contents\(\"([^\"]+)\"\)\s*\)", expr)
        if m:
            return [self._web_page_first_table(m.group(1))]

        # Источник{0}[Data]
        m = re.fullmatch(r"(.+?)\{0\}\[Data\]", expr)
        if m:
            source = self._resolve_ref(m.group(1), ctx)
            if isinstance(source, list) and source:
                return self._ensure_df(source[0])
            return self._ensure_df(source)

        # Html.Table(Источник, ..., [RowSelector="..."])
        if expr.startswith("Html.Table"):
            source_ref = self._first_arg(expr)
            html = self._resolve_ref(source_ref, ctx)
            row_selector = self._extract_row_selector(expr)
            col_count = self._extract_html_table_column_count(expr)
            return self._html_table(str(html), row_selector, col_count, expr)

        if expr.startswith("Table.PromoteHeaders"):
            df = self._first_df_arg(expr, ctx)
            return self._promote_headers(df)

        if expr.startswith("Table.TransformColumnTypes"):
            df = self._first_df_arg(expr, ctx)
            return self._transform_column_types(df, expr)

        if expr.startswith("Table.ReplaceErrorValues"):
            # Ошибки типов в Python уже переводятся в None.
            return self._first_df_arg(expr, ctx)

        if expr.startswith("Table.Skip"):
            df = self._first_df_arg(expr, ctx)
            m = re.search(r",\s*(\d+)\s*\)$", expr)
            n = int(m.group(1)) if m else 0
            return df.iloc[n:].reset_index(drop=True)

        if expr.startswith("Table.RenameColumns"):
            df = self._first_df_arg(expr, ctx)
            pairs = self._parse_string_pairs(expr)
            rename_map = {old: new for old, new in pairs if old in df.columns}
            return df.rename(columns=rename_map)

        if expr.startswith("Table.RemoveColumns"):
            df = self._first_df_arg(expr, ctx)
            cols = self._parse_string_list_after_first_arg(expr)
            return df.drop(columns=[c for c in cols if c in df.columns], errors="ignore")

        if expr.startswith("Table.ReorderColumns"):
            df = self._first_df_arg(expr, ctx)
            cols = self._parse_string_list_after_first_arg(expr)
            for col in cols:
                if col not in df.columns:
                    df[col] = None
            return df.loc[:, cols]

        if expr.startswith("Table.NestedJoin"):
            return self._nested_join(expr, ctx)

        if expr.startswith("Table.ExpandTableColumn"):
            return self._expand_table_column(expr, ctx)

        if expr.startswith("Table.TransformColumns"):
            df = self._first_df_arg(expr, ctx)
            return self._transform_columns(df, expr)

        # Прямая ссылка на шаг или на другой запрос.
        return self._resolve_ref(expr, ctx)

    @staticmethod
    def _first_arg(expr: str) -> str:
        start = expr.find("(") + 1
        depth_par = depth_brace = depth_bracket = 0
        in_string = False
        for i in range(start, len(expr)):
            ch = expr[i]
            if in_string:
                if ch == '"':
                    if i + 1 < len(expr) and expr[i + 1] == '"':
                        continue
                    in_string = False
                continue
            if ch == '"':
                in_string = True
            elif ch == '(':
                depth_par += 1
            elif ch == ')':
                if depth_par == 0:
                    break
                depth_par -= 1
            elif ch == '{':
                depth_brace += 1
            elif ch == '}':
                depth_brace -= 1
            elif ch == '[':
                depth_bracket += 1
            elif ch == ']':
                depth_bracket -= 1
            elif ch == ',' and depth_par == depth_brace == depth_bracket == 0:
                return expr[start:i].strip()
        return expr[start:].strip()

    def _first_df_arg(self, expr: str, ctx: Dict[str, Any]) -> pd.DataFrame:
        return self._ensure_df(self._resolve_ref(self._first_arg(expr), ctx))

    def _resolve_ref(self, ref: str, ctx: Dict[str, Any]) -> Any:
        name = self._normalize_identifier(ref.strip())
        if name in ctx:
            return ctx[name]
        if name in self.queries:
            return self.evaluate_query(name)
        raise KeyError(f"Не удалось найти ссылку M: {ref}")

    @staticmethod
    def _ensure_df(value: Any) -> pd.DataFrame:
        if isinstance(value, pd.DataFrame):
            return value.copy()
        raise TypeError(f"Ожидалась таблица pandas.DataFrame, получено: {type(value)!r}")

    def _fetch_html(self, url: str) -> str:
        key = hashlib.sha1(url.encode("utf-8")).hexdigest() + ".html"
        path = self.cache_dir / key
        self.logger.debug("HTML request: url=%s cache_key=%s", url, key)
        if self.use_cache and path.exists():
            html = path.read_text(encoding="utf-8", errors="replace")
            self.logger.debug("HTML cache HIT: %s bytes=%d", path, len(html.encode("utf-8", errors="replace")))
            return html
        self.logger.info("HTML download: %s", url)
        response = self.session.get(url, timeout=self.timeout)
        self.logger.debug("HTTP response: status=%s encoding=%s url=%s", response.status_code, response.encoding, url)
        response.raise_for_status()
        response.encoding = response.encoding or "utf-8"
        html = response.text
        if self.use_cache:
            path.write_text(html, encoding="utf-8")
            self.logger.debug("HTML cache SAVED: %s bytes=%d", path, len(html.encode("utf-8", errors="replace")))
        # Небольшая пауза, чтобы не долбить сайт ЦБ подряд десятками запросов.
        time.sleep(0.25)
        return html

    def _web_page_first_table(self, url: str) -> pd.DataFrame:
        html = self._fetch_html(url)
        # Для Web.Page ближе всего pandas.read_html: он учитывает rowspan/colspan и заголовки.
        # Сначала используем lxml, чтобы не требовать html5lib; если lxml не справился,
        # пробуем стандартный autodetect pandas.
        try:
            tables = pd.read_html(io.StringIO(html), flavor="lxml")
        except Exception:
            tables = pd.read_html(io.StringIO(html))
        if not tables:
            raise ValueError(f"На странице нет HTML-таблиц: {url}")
        self.logger.debug("Web.Page tables found: url=%s count=%d", url, len(tables))
        df = tables[0]
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = [self._clean_text(" ".join(str(x) for x in tup if str(x) != "nan")) for tup in df.columns]
        else:
            df.columns = [self._clean_text(c) for c in df.columns]
        df = self._normalise_dataframe(df)
        self.logger.debug("Web.Page selected first table: %s", df_debug_summary(df))
        return df

    def _html_table(self, html: str, row_selector: str, col_count: int, expr: str = "") -> pd.DataFrame:
        """
        Аналог Html.Table из Power Query.

        В оригинальной книге для части страниц ЦБ используются очень хрупкие CSS-селекторы
        вида DIV.table-wrapper:nth-child(...). При смене банка количество блоков на странице
        меняется, и такой селектор может попасть в соседнюю таблицу. Поэтому сначала пробуем
        исходный селектор, а затем для типовых страниц ЦБ делаем fallback: выбираем таблицу
        по содержанию заголовков.
        """
        soup = BeautifulSoup(html, "html.parser")
        selector = self._normalise_css_selector(row_selector)
        rows = soup.select(selector) if selector else []
        self.logger.debug("Html.Table selector=%s normalized=%s rows=%d col_count=%s", row_selector, selector, len(rows), col_count)
        if not rows:
            # Fallback: берём первую таблицу класса data / data spaced.
            table = soup.select_one("table.data") or soup.find("table")
            rows = table.find_all("tr") if table else []
            self.logger.debug("Html.Table selector fallback to first table: rows=%d", len(rows))

        df = self._rows_to_dataframe(rows, col_count)
        self.logger.debug("Html.Table initial result: %s", df_debug_summary(df))

        # На страницах Ф.135 ЦБ селекторы nth-child часто начинают указывать на таблицу
        # "Код обозначения / Сумма, тыс. руб." вместо таблицы нормативов. Исправляем
        # выбор по заголовкам: нужна таблица с "Краткое наименование норматива".
        fallback = self._cbr_f135_table_fallback(soup, df, row_selector, col_count)
        if fallback is not None:
            self.logger.debug("Html.Table F135 fallback selected: %s", df_debug_summary(fallback))
            return fallback
        return df

    def _rows_to_dataframe(self, rows: List[Any], col_count: int) -> pd.DataFrame:
        data: List[List[Any]] = []
        for tr in rows:
            cells = tr.find_all(["th", "td"], recursive=False)
            if not cells:
                continue
            values: List[Any] = []
            for cell in cells:
                text = self._clean_text(cell.get_text(" ", strip=True))
                colspan = int(cell.get("colspan") or 1)
                for _ in range(max(colspan, 1)):
                    values.append(text)
            if col_count:
                if len(values) < col_count:
                    values.extend([None] * (col_count - len(values)))
                elif len(values) > col_count:
                    values = values[:col_count]
            data.append(values)
        if not data:
            raise ValueError("Html.Table: не удалось извлечь строки из HTML")
        if not col_count:
            col_count = max(len(r) for r in data)
            data = [r + [None] * (col_count - len(r)) for r in data]
        columns = [f"Column{i}" for i in range(1, col_count + 1)]
        return pd.DataFrame(data, columns=columns)

    def _cbr_f135_table_fallback(
        self,
        soup: BeautifulSoup,
        current_df: pd.DataFrame,
        row_selector: str,
        col_count: int,
    ) -> Optional[pd.DataFrame]:
        """Ищет правильную таблицу нормативов Ф.135, если nth-child попал не туда."""
        selector_l = (row_selector or "").lower()
        if col_count != 4 or "table.data" not in selector_l:
            return None

        first_row = " ".join(str(x) for x in (current_df.iloc[0].tolist() if not current_df.empty else []))
        if "Краткое наименование норматива" in first_row:
            return None

        for table in soup.select("table.data"):
            rows = table.find_all("tr")
            if not rows:
                continue
            header_text = " ".join(
                self._clean_text(cell.get_text(" ", strip=True)) or ""
                for cell in rows[0].find_all(["th", "td"], recursive=False)
            )
            # В разные годы заголовок мог быть "... (требования)", но начало стабильно.
            if "Краткое наименование норматива" in header_text and "Фактическое значение" in header_text:
                return self._rows_to_dataframe(rows, col_count)
        return None

    @staticmethod
    def _normalise_css_selector(selector: str) -> str:
        selector = selector.replace('""', '"')
        # SoupSieve обычно понимает верхний регистр, но нижний стабильнее для html.parser.
        selector = re.sub(r"\b(TABLE|TR|TH|TD|DIV)\b", lambda m: m.group(1).lower(), selector)
        return selector

    @staticmethod
    def _extract_row_selector(expr: str) -> str:
        m = re.search(r"RowSelector\s*=\s*\"((?:\"\"|[^\"])*)\"", expr)
        return m.group(1).replace('""', '"') if m else ""

    @staticmethod
    def _extract_html_table_column_count(expr: str) -> int:
        nums = [int(x) for x in re.findall(r'\{\"Column(\d+)\"\s*,', expr)]
        return max(nums) if nums else 0

    @staticmethod
    def _clean_text(value: Any) -> Any:
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return None
        text = str(value).replace("\xa0", " ")
        text = re.sub(r"\s+", " ", text).strip()
        return text if text != "nan" else None

    def _normalise_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        out = df.copy()
        out.columns = [self._clean_text(c) for c in out.columns]
        for col in out.columns:
            out[col] = out[col].map(self._clean_text)
        return out

    def _promote_headers(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy().reset_index(drop=True)
        if df.empty:
            return df
        headers = [self._clean_text(v) or f"Column{i+1}" for i, v in enumerate(df.iloc[0].tolist())]
        headers = self._make_unique(headers)
        out = df.iloc[1:].reset_index(drop=True)
        out.columns = headers[: len(out.columns)]
        return out

    @staticmethod
    def _make_unique(names: List[str]) -> List[str]:
        seen: Dict[str, int] = {}
        result: List[str] = []
        for name in names:
            base = name or "Column"
            if base not in seen:
                seen[base] = 0
                result.append(base)
            else:
                seen[base] += 1
                result.append(f"{base}.{seen[base]}")
        return result

    @staticmethod
    def _parse_string_pairs(expr: str) -> List[Tuple[str, str]]:
        return [(a.replace('""', '"'), b.replace('""', '"')) for a, b in re.findall(r'\{\s*\"((?:\"\"|[^\"])*)\"\s*,\s*\"((?:\"\"|[^\"])*)\"\s*\}', expr)]

    @staticmethod
    def _parse_string_list_after_first_arg(expr: str) -> List[str]:
        # Берём первую {...} после первой запятой верхнего уровня.
        pos = expr.find(",")
        if pos < 0:
            return []
        text = expr[pos + 1 :]
        return [s.replace('""', '"') for s in re.findall(r'\"((?:\"\"|[^\"])*)\"', text)]

    def _transform_column_types(self, df: pd.DataFrame, expr: str) -> pd.DataFrame:
        out = df.copy()
        pairs = re.findall(r'\{\s*\"((?:\"\"|[^\"])*)\"\s*,\s*([^\}]+?)\s*\}', expr)
        for col, mtype in pairs:
            col = col.replace('""', '"')
            if col not in out.columns:
                continue
            if "Int64.Type" in mtype:
                out[col] = out[col].map(self._parse_int)
            elif "Percentage.Type" in mtype:
                out[col] = out[col].map(self._parse_number)
            elif "type number" in mtype:
                out[col] = out[col].map(self._parse_number)
            elif "type text" in mtype:
                out[col] = out[col].map(lambda x: None if pd.isna(x) else str(x))
        return out

    @staticmethod
    def _parse_int(value: Any) -> Optional[int]:
        num = MiniMEngine._parse_number(value)
        if num is None:
            return None
        try:
            return int(round(num))
        except Exception:
            return None

    @staticmethod
    def _parse_number(value: Any) -> Optional[float]:
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return None
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            return float(value)
        text = str(value).replace("\xa0", " ").strip()
        if text in {"", "-", "—", "x", "X"}:
            return None
        text = text.replace("%", "")
        # Скобки как отрицательное число.
        negative = text.startswith("(") and text.endswith(")")
        text = text.strip("()")
        text = text.replace(" ", "")
        # Русские десятичные запятые и англ. точки.
        if "," in text and "." in text:
            # 1,234.56 или 1.234,56 — выбираем по последнему разделителю.
            if text.rfind(",") > text.rfind("."):
                text = text.replace(".", "").replace(",", ".")
            else:
                text = text.replace(",", "")
        else:
            text = text.replace(",", ".")
        try:
            val = float(text)
            return -val if negative else val
        except Exception:
            return None

    def _transform_columns(self, df: pd.DataFrame, expr: str) -> pd.DataFrame:
        out = df.copy()
        # Поддерживает вид: { {"2024", each _ / 100, type number} }
        matches = re.findall(r'\{\s*\"((?:\"\"|[^\"])*)\"\s*,\s*each\s+_\s*/\s*([0-9]+)', expr)
        for col, divisor in matches:
            col = col.replace('""', '"')
            if col in out.columns:
                div = float(divisor)
                out[col] = out[col].map(lambda x: None if self._parse_number(x) is None else self._parse_number(x) / div)
        return out

    def _nested_join(self, expr: str, ctx: Dict[str, Any]) -> pd.DataFrame:
        args = self._function_args(expr)
        if len(args) < 6:
            raise ValueError(f"Не удалось разобрать Table.NestedJoin: {expr[:200]}")
        left_ref, left_keys_expr, right_ref, right_keys_expr, nested_expr = args[:5]
        nested_name = self._unquote(nested_expr)
        left_keys_requested = self._parse_string_list(left_keys_expr)
        right_keys_requested = self._parse_string_list(right_keys_expr)
        left = self._ensure_df(self._resolve_ref(left_ref, ctx))
        right = self._ensure_df(self._resolve_ref(right_ref, ctx))

        left_keys = [self._resolve_column_name(left, key, "левой", expr) for key in left_keys_requested]
        right_keys = [self._resolve_column_name(right, key, "правой", expr) for key in right_keys_requested]

        # Оставляем ключи + неключевые столбцы справа; неключевые переименовываем в nested.col.
        right_work = right.copy()
        rename_map: Dict[str, str] = {}
        for col in right_work.columns:
            if col not in right_keys:
                rename_map[col] = f"{nested_name}.{col}"
        right_work = right_work.rename(columns=rename_map)

        merged = left.merge(right_work, how="left", left_on=left_keys, right_on=right_keys, sort=False)
        # Удаляем дубли ключей справа, если названия отличались.
        for rk in right_keys:
            if rk not in left_keys and rk in merged.columns:
                merged = merged.drop(columns=[rk])
        return merged

    def _resolve_column_name(self, df: pd.DataFrame, requested: str, side_name: str, expr: str) -> str:
        """Находит колонку с учётом мелких отличий заголовков на страницах ЦБ."""
        if requested in df.columns:
            return requested
        requested_norm = self._normalise_column_for_match(requested)
        for col in df.columns:
            if self._normalise_column_for_match(col) == requested_norm:
                return col
        columns_preview = ", ".join(str(c) for c in list(df.columns)[:12])
        raise KeyError(
            f"Не найдена колонка соединения {requested!r} в {side_name} таблице. "
            f"Доступные колонки: {columns_preview}. Выражение M: {expr[:250]}"
        )

    @staticmethod
    def _normalise_column_for_match(value: Any) -> str:
        text = MiniMEngine._clean_text(value) or ""
        text = str(text).lower().replace("ё", "е")
        # Старые страницы ЦБ иногда добавляют уточнение: "(требования)".
        text = re.sub(r"\s*\([^)]*\)", "", text)
        text = re.sub(r"[^0-9a-zа-я]+", "", text)
        return text

    def _expand_table_column(self, expr: str, ctx: Dict[str, Any]) -> pd.DataFrame:
        args = self._function_args(expr)
        if len(args) < 4:
            raise ValueError(f"Не удалось разобрать Table.ExpandTableColumn: {expr[:200]}")
        source_ref, nested_expr, cols_expr, new_cols_expr = args[:4]
        nested_name = self._unquote(nested_expr)
        cols = self._parse_string_list(cols_expr)
        new_cols = self._parse_string_list(new_cols_expr)
        df = self._ensure_df(self._resolve_ref(source_ref, ctx))
        out = df.copy()
        for old, new in zip(cols, new_cols):
            prefixed = f"{nested_name}.{old}"
            if prefixed in out.columns:
                out = out.rename(columns={prefixed: new})
            elif old in out.columns and new not in out.columns:
                out = out.rename(columns={old: new})
            elif new not in out.columns:
                out[new] = None
        # Удаляем неразвёрнутые nested.* остатки.
        leftovers = [c for c in out.columns if isinstance(c, str) and c.startswith(nested_name + ".")]
        if leftovers:
            out = out.drop(columns=leftovers)
        return out

    @staticmethod
    def _function_args(expr: str) -> List[str]:
        start = expr.find("(") + 1
        end = expr.rfind(")")
        inner = expr[start:end]
        args: List[str] = []
        depth_par = depth_brace = depth_bracket = 0
        in_string = False
        arg_start = 0
        i = 0
        while i < len(inner):
            ch = inner[i]
            if in_string:
                if ch == '"':
                    if i + 1 < len(inner) and inner[i + 1] == '"':
                        i += 2
                        continue
                    in_string = False
                i += 1
                continue
            if ch == '"':
                in_string = True
            elif ch == '(':
                depth_par += 1
            elif ch == ')':
                depth_par -= 1
            elif ch == '{':
                depth_brace += 1
            elif ch == '}':
                depth_brace -= 1
            elif ch == '[':
                depth_bracket += 1
            elif ch == ']':
                depth_bracket -= 1
            elif ch == ',' and depth_par == depth_brace == depth_bracket == 0:
                args.append(inner[arg_start:i].strip())
                arg_start = i + 1
            i += 1
        tail = inner[arg_start:].strip()
        if tail:
            args.append(tail)
        return args

    @staticmethod
    def _parse_string_list(expr: str) -> List[str]:
        return [s.replace('""', '"') for s in re.findall(r'\"((?:\"\"|[^\"])*)\"', expr)]

    @staticmethod
    def _unquote(expr: str) -> str:
        expr = expr.strip()
        if expr.startswith('"') and expr.endswith('"'):
            return expr[1:-1].replace('""', '"')
        return MiniMEngine._normalize_identifier(expr)




def make_excel_table_headers_unique(headers: List[str]) -> List[str]:
    """Excel-таблица не допускает пустые и повторяющиеся имена колонок."""
    result: List[str] = []
    seen: Dict[str, int] = {}
    for idx, header in enumerate(headers, start=1):
        base = str(header).strip() if header is not None else ""
        if not base:
            base = f"Column{idx}"
        key = base.lower()
        if key not in seen:
            seen[key] = 1
            result.append(base)
            continue
        seen[key] += 1
        while True:
            candidate = f"{base}_{seen[key]}"
            candidate_key = candidate.lower()
            if candidate_key not in seen:
                seen[candidate_key] = 1
                result.append(candidate)
                break
            seen[key] += 1
    return result


# -----------------------------------------------------------------------------
# Низкоуровневая запись XLSX через ZIP/XML
# -----------------------------------------------------------------------------

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", REL_NS)
ET.register_namespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
ET.register_namespace("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")
ET.register_namespace("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision")
ET.register_namespace("xr6", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6")
ET.register_namespace("xr10", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10")
ET.register_namespace("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2")
ET.register_namespace("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3")
ET.register_namespace("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main")
ET.register_namespace("xcalcf", "http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures")

KNOWN_XMLNS_PREFIXES = {
    "r": REL_NS,
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
    "x15": "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
    "xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
    "xr2": "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
    "xr3": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3",
    "xr6": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6",
    "xr10": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10",
    "xcalcf": "http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures",
}


def xq(tag: str) -> str:
    return f"{{{MAIN_NS}}}{tag}"


def rq(tag: str) -> str:
    return f"{{{REL_NS}}}{tag}"


def ensure_mc_ignorable_namespaces(xml_bytes: bytes) -> bytes:
    """
    ElementTree при сериализации удаляет неиспользуемые xmlns:* и может переименовывать
    префиксы расширений Excel в ns0/ns1. Для Excel это критично: если mc:Ignorable
    содержит, например, xr3, то prefix xr3 обязан быть объявлен в корневом теге.
    """
    text = xml_bytes.decode("utf-8")
    root_match = re.search(r"<([A-Za-z_][\w:.-]*)(\s[^<>]*?)?>", text)
    if not root_match:
        return xml_bytes
    root_tag_text = root_match.group(0)
    ign = re.search(r'\bmc:Ignorable="([^"]+)"', root_tag_text)
    if not ign:
        return xml_bytes
    declared = set(re.findall(r'\bxmlns:([A-Za-z_][\w.-]*)="', root_tag_text))
    missing_decls: List[str] = []
    for prefix in ign.group(1).split():
        if prefix in declared:
            continue
        uri = KNOWN_XMLNS_PREFIXES.get(prefix)
        if uri:
            missing_decls.append(f' xmlns:{prefix}="{uri}"')
    if not missing_decls:
        return xml_bytes
    insert_pos = root_match.end() - 1  # перед закрывающей > корневого тега
    text = text[:insert_pos] + "".join(missing_decls) + text[insert_pos:]
    return text.encode("utf-8")


def xml_tostring_excel_safe(root: ET.Element) -> bytes:
    """Сериализует XML так, чтобы Excel не ругался на mc:Ignorable/namespace."""
    return ensure_mc_ignorable_namespaces(ET.tostring(root, encoding="utf-8", xml_declaration=True))


def col_to_num(col: str) -> int:
    num = 0
    for ch in col.upper():
        if "A" <= ch <= "Z":
            num = num * 26 + ord(ch) - ord("A") + 1
    return num


def num_to_col(num: int) -> str:
    result = []
    while num:
        num, rem = divmod(num - 1, 26)
        result.append(chr(ord("A") + rem))
    return "".join(reversed(result))


def split_cell_ref(cell_ref: str) -> Tuple[int, int]:
    m = re.fullmatch(r"([A-Z]+)(\d+)", cell_ref.upper())
    if not m:
        raise ValueError(f"Некорректная ссылка на ячейку: {cell_ref!r}")
    return col_to_num(m.group(1)), int(m.group(2))


def parse_range_ref(ref: str) -> Tuple[int, int, int, int]:
    if ":" not in ref:
        c, r = split_cell_ref(ref)
        return c, r, c, r
    start, end = ref.split(":", 1)
    min_col, min_row = split_cell_ref(start)
    max_col, max_row = split_cell_ref(end)
    return min_col, min_row, max_col, max_row


def make_range_ref(min_col: int, min_row: int, max_col: int, max_row: int) -> str:
    return f"{num_to_col(min_col)}{min_row}:{num_to_col(max_col)}{max_row}"


def normalize_zip_target(base_dir: str, target: str) -> str:
    """Нормализует Target из .rels к пути внутри ZIP."""
    import posixpath

    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(base_dir, target))


def worksheet_and_table_maps(zip_bytes: Dict[str, bytes]) -> Tuple[Dict[str, str], Dict[Tuple[str, str], str]]:
    """Возвращает: имя листа -> worksheet xml; (лист, имя таблицы) -> table xml."""
    workbook = ET.fromstring(zip_bytes["xl/workbook.xml"])
    wb_rels = ET.fromstring(zip_bytes["xl/_rels/workbook.xml.rels"])

    rel_targets: Dict[str, str] = {}
    for rel in wb_rels:
        rid = rel.attrib.get("Id")
        target = rel.attrib.get("Target", "")
        if rid:
            rel_targets[rid] = normalize_zip_target("xl", target)

    sheet_to_path: Dict[str, str] = {}
    for sheet in workbook.findall(f".//{xq('sheet')}"):
        name = sheet.attrib.get("name")
        rid = sheet.attrib.get(rq("id"))
        if name and rid in rel_targets:
            sheet_to_path[name] = rel_targets[rid]

    table_map: Dict[Tuple[str, str], str] = {}
    for sheet_name, sheet_path in sheet_to_path.items():
        rel_path = f"{str(Path(sheet_path).parent).replace(os.sep, '/')}/_rels/{Path(sheet_path).name}.rels"
        # Path() на Windows может менять разделители; подстрахуемся posix-вариантом.
        rel_path = "/".join(sheet_path.split("/")[:-1]) + "/_rels/" + sheet_path.split("/")[-1] + ".rels"
        if rel_path not in zip_bytes:
            continue
        sheet_rels = ET.fromstring(zip_bytes[rel_path])
        for rel in sheet_rels:
            target = rel.attrib.get("Target", "")
            rel_type = rel.attrib.get("Type", "")
            if "table" not in rel_type and "tables/" not in target:
                continue
            table_path = normalize_zip_target("/".join(sheet_path.split("/")[:-1]), target)
            if table_path not in zip_bytes:
                continue
            table_root = ET.fromstring(zip_bytes[table_path])
            for attr_name in ("displayName", "name"):
                tname = table_root.attrib.get(attr_name)
                if tname:
                    table_map[(sheet_name, tname)] = table_path
    return sheet_to_path, table_map


def read_shared_strings(zip_bytes: Dict[str, bytes]) -> List[str]:
    """Читает sharedStrings для диагностики/сравнения; для записи используем inlineStr."""
    if "xl/sharedStrings.xml" not in zip_bytes:
        return []
    root = ET.fromstring(zip_bytes["xl/sharedStrings.xml"])
    result: List[str] = []
    for si in root.findall(xq("si")):
        parts: List[str] = []
        for t in si.iter(xq("t")):
            parts.append(t.text or "")
        result.append("".join(parts))
    return result


def get_row_map(sheet_root: ET.Element) -> Dict[int, ET.Element]:
    sheet_data = sheet_root.find(xq("sheetData"))
    if sheet_data is None:
        sheet_data = ET.SubElement(sheet_root, xq("sheetData"))
    result: Dict[int, ET.Element] = {}
    for row in sheet_data.findall(xq("row")):
        r = row.attrib.get("r")
        if r and str(r).isdigit():
            result[int(r)] = row
    return result


def get_or_create_row(sheet_root: ET.Element, row_num: int) -> ET.Element:
    sheet_data = sheet_root.find(xq("sheetData"))
    if sheet_data is None:
        sheet_data = ET.SubElement(sheet_root, xq("sheetData"))
    row_map = get_row_map(sheet_root)
    if row_num in row_map:
        return row_map[row_num]
    new_row = ET.Element(xq("row"), {"r": str(row_num)})
    rows = sheet_data.findall(xq("row"))
    insert_at = len(sheet_data)
    for i, row in enumerate(rows):
        try:
            if int(row.attrib.get("r", "0")) > row_num:
                insert_at = list(sheet_data).index(row)
                break
        except Exception:
            pass
    sheet_data.insert(insert_at, new_row)
    return new_row


def get_cell_map(row: ET.Element) -> Dict[int, ET.Element]:
    result: Dict[int, ET.Element] = {}
    for cell in row.findall(xq("c")):
        ref = cell.attrib.get("r")
        if not ref:
            continue
        try:
            col_num, _ = split_cell_ref(ref)
            result[col_num] = cell
        except Exception:
            continue
    return result


def get_or_create_cell(row: ET.Element, col_num: int, row_num: int, style_template: Optional[ET.Element] = None) -> ET.Element:
    cell_map = get_cell_map(row)
    if col_num in cell_map:
        return cell_map[col_num]
    ref = f"{num_to_col(col_num)}{row_num}"
    attrs = {"r": ref}
    if style_template is not None and style_template.attrib.get("s") is not None:
        attrs["s"] = style_template.attrib["s"]
    cell = ET.Element(xq("c"), attrs)
    children = list(row)
    insert_at = len(children)
    for i, existing in enumerate(children):
        if existing.tag != xq("c"):
            continue
        eref = existing.attrib.get("r", "")
        try:
            ecol, _ = split_cell_ref(eref)
            if ecol > col_num:
                insert_at = i
                break
        except Exception:
            pass
    row.insert(insert_at, cell)
    return cell


def clear_cell_value(cell: ET.Element) -> None:
    """Очищает значение/формулу ячейки, сохраняя стиль и адрес."""
    for child in list(cell):
        if child.tag in {xq("v"), xq("is"), xq("f")}:  # значения, inline-string, формулы
            cell.remove(child)
    for attr in ("t", "cm", "vm", "ph"):
        cell.attrib.pop(attr, None)


def set_cell_value(cell: ET.Element, value: Any) -> None:
    """Ставит значение в ячейку Open XML. Строки пишутся как inlineStr."""
    clear_cell_value(cell)
    if value is None:
        return
    try:
        if pd.isna(value):
            return
    except Exception:
        pass
    if hasattr(value, "item"):
        try:
            value = value.item()
        except Exception:
            pass

    if isinstance(value, bool):
        cell.attrib["t"] = "b"
        ET.SubElement(cell, xq("v")).text = "1" if value else "0"
        return
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        # Excel Open XML всегда использует точку как десятичный разделитель.
        cell.attrib.pop("t", None)
        ET.SubElement(cell, xq("v")).text = repr(float(value)) if isinstance(value, float) else str(value)
        return

    text = str(value)
    cell.attrib["t"] = "inlineStr"
    is_el = ET.SubElement(cell, xq("is"))
    t_el = ET.SubElement(is_el, xq("t"))
    if text.startswith(" ") or text.endswith(" ") or "\n" in text or "\t" in text:
        t_el.attrib["{http://www.w3.org/XML/1998/namespace}space"] = "preserve"
    t_el.text = text


def find_style_template(sheet_root: ET.Element, col_num: int, preferred_rows: Iterable[int]) -> Optional[ET.Element]:
    """Ищет ячейку-шаблон стиля для новых строк в той же колонке."""
    row_map = get_row_map(sheet_root)
    for row_num in preferred_rows:
        row = row_map.get(row_num)
        if row is None:
            continue
        cell = get_cell_map(row).get(col_num)
        if cell is not None and cell.attrib.get("s") is not None:
            return cell
    return None


def update_sheet_dimension(sheet_root: ET.Element, min_col: int, min_row: int, max_col: int, max_row: int) -> None:
    dim = sheet_root.find(xq("dimension"))
    if dim is None:
        return
    old_ref = dim.attrib.get("ref", "A1")
    try:
        o_min_col, o_min_row, o_max_col, o_max_row = parse_range_ref(old_ref)
        n_min_col = min(o_min_col, min_col)
        n_min_row = min(o_min_row, min_row)
        n_max_col = max(o_max_col, max_col)
        n_max_row = max(o_max_row, max_row)
    except Exception:
        n_min_col, n_min_row, n_max_col, n_max_row = min_col, min_row, max_col, max_row
    dim.attrib["ref"] = make_range_ref(n_min_col, n_min_row, n_max_col, n_max_row)


def write_values_to_sheet_xml(sheet_xml: bytes, old_ref: str, new_ref: str, headers: List[str], rows: List[List[Any]]) -> bytes:
    """Очищает старую Excel-таблицу и записывает новые header+data в XML листа."""
    root = ET.fromstring(sheet_xml)
    old_min_col, old_min_row, old_max_col, old_max_row = parse_range_ref(old_ref)
    new_min_col, new_min_row, new_max_col, new_max_row = parse_range_ref(new_ref)

    # Очищаем старый диапазон таблицы, но стили оставляем.
    row_map = get_row_map(root)
    for r in range(old_min_row, old_max_row + 1):
        row = row_map.get(r)
        if row is None:
            continue
        cells = get_cell_map(row)
        for c in range(old_min_col, old_max_col + 1):
            cell = cells.get(c)
            if cell is not None:
                clear_cell_value(cell)

    values = [headers] + rows
    for r_offset, row_values in enumerate(values):
        row_num = new_min_row + r_offset
        row = get_or_create_row(root, row_num)
        for c_offset, value in enumerate(row_values):
            col_num = new_min_col + c_offset
            template = find_style_template(root, col_num, [row_num, old_min_row + min(r_offset, max(0, old_max_row - old_min_row)), old_max_row, old_min_row + 1, old_min_row])
            cell = get_or_create_cell(row, col_num, row_num, template)
            set_cell_value(cell, value)

    update_sheet_dimension(root, new_min_col, new_min_row, new_max_col, new_max_row)
    return xml_tostring_excel_safe(root)


def write_values_to_sheet_xml_preserve_table_range(
    sheet_xml: bytes,
    table_ref: str,
    headers: List[str],
    rows: List[List[Any]],
    logger: Optional[logging.Logger] = None,
    context: str = "",
) -> bytes:
    """
    Безопасная запись значений без изменения размеров Excel-таблицы.

    Для книг с Power Query/QueryTable нельзя надёжно менять xl/tables/table*.xml
    и xl/queryTables/queryTable*.xml вручную: Excel строго проверяет эти части
    и при малейшем расхождении удаляет диапазоны внешних данных. Поэтому этот
    режим оставляет существующий table_ref неизменным и пишет данные только в
    уже существующие клетки таблицы. Лишние старые значения внутри table_ref
    очищаются, стили и сама таблица сохраняются.
    """
    logger = logger or logging.getLogger("pq2py")
    min_col, min_row, max_col, max_row = parse_range_ref(table_ref)
    capacity_cols = max_col - min_col + 1
    capacity_data_rows = max(0, max_row - min_row)

    safe_headers = [str(h) if h is not None else "" for h in headers[:capacity_cols]]
    if len(safe_headers) < capacity_cols:
        safe_headers += [""] * (capacity_cols - len(safe_headers))

    safe_rows: List[List[Any]] = []
    for src_row in rows[:capacity_data_rows]:
        row = list(src_row[:capacity_cols])
        if len(row) < capacity_cols:
            row += [None] * (capacity_cols - len(row))
        safe_rows.append(row)

    if len(headers) != capacity_cols:
        logger.warning(
            "%s: число колонок результата (%s) отличается от ширины существующей Excel-таблицы (%s); "
            "записываю по позиции, без изменения table/queryTable XML",
            context, len(headers), capacity_cols,
        )
    if len(rows) > capacity_data_rows:
        logger.warning(
            "%s: строк результата больше, чем помещается в исходную Excel-таблицу: %s > %s. "
            "В безопасном режиме лишние строки не записаны, чтобы не ломать QueryTable.",
            context, len(rows), capacity_data_rows,
        )
    if len(rows) < capacity_data_rows:
        logger.info(
            "%s: строк результата меньше исходного диапазона таблицы: %s < %s; хвост диапазона очищен и оставлен пустым.",
            context, len(rows), capacity_data_rows,
        )

    return write_values_to_sheet_xml(sheet_xml, table_ref, table_ref, safe_headers, safe_rows)



def table_column_headers_from_xml(table_xml: bytes) -> List[str]:
    """Читает имена столбцов Excel-таблицы из tableColumns."""
    root = ET.fromstring(table_xml)
    table_columns = root.find(xq("tableColumns"))
    if table_columns is None:
        return []
    return [c.attrib.get("name", "") for c in table_columns.findall(xq("tableColumn"))]


def querytable_fields_for_table(zip_bytes: Dict[str, bytes], table_path: str) -> List[Dict[str, str]]:
    """
    Возвращает поля связанного queryTable*.xml для Excel-таблицы.

    Это критично для файлов с Power Query: tableColumn.id / queryTableFieldId
    нельзя перенумеровывать по порядку, потому что Excel сверяет их со связанным
    queryTable*.xml. При несовпадении Excel открывает книгу как повреждённую.
    """
    import posixpath

    rel_path = posixpath.join(posixpath.dirname(table_path), "_rels", posixpath.basename(table_path) + ".rels")
    if rel_path not in zip_bytes:
        return []
    try:
        rel_root = ET.fromstring(zip_bytes[rel_path])
    except Exception:
        return []
    qt_path = ""
    for rel in rel_root:
        rel_type = rel.attrib.get("Type", "")
        target = rel.attrib.get("Target", "")
        if "queryTable" in rel_type or "queryTables/" in target:
            qt_path = normalize_zip_target(posixpath.dirname(table_path), target)
            break
    if not qt_path or qt_path not in zip_bytes:
        return []
    try:
        qt_root = ET.fromstring(zip_bytes[qt_path])
    except Exception:
        return []
    fields: List[Dict[str, str]] = []
    for field in qt_root.findall(f".//{xq('queryTableField')}"):
        fid = field.attrib.get("id", "")
        fields.append({
            "id": fid,
            "name": field.attrib.get("name", ""),
            "tableColumnId": field.attrib.get("tableColumnId", fid),
            "queryTablePath": qt_path,
        })
    return fields

def update_table_xml(
    table_xml: bytes,
    new_ref: str,
    headers: List[str],
    query_fields: Optional[List[Dict[str, str]]] = None,
) -> bytes:
    """
    Обновляет XML Excel-таблицы: ref и autoFilter.ref.

    Важно: для Power Query / queryTable нельзя просто перенумеровать
    tableColumn.id, uniqueName и queryTableFieldId как 1..N. Эти id связаны
    с xl/queryTables/queryTable*.xml. Excel на macOS/Windows проверяет эту
    связку и помечает файл повреждённым, если она нарушена.
    """
    root = ET.fromstring(table_xml)
    root.attrib["ref"] = new_ref
    auto_filter = root.find(xq("autoFilter"))
    if auto_filter is not None:
        auto_filter.attrib["ref"] = new_ref

    table_columns = root.find(xq("tableColumns"))
    if table_columns is None:
        table_columns = ET.Element(xq("tableColumns"))
        children = list(root)
        insert_at = len(children)
        for idx, child in enumerate(children):
            if child.tag == xq("tableStyleInfo"):
                insert_at = idx
                break
        root.insert(insert_at, table_columns)

    old_columns = list(table_columns.findall(xq("tableColumn")))

    # Безопасный путь: количество колонок не изменилось. Сохраняем исходные
    # элементы tableColumn, включая xr3:uid и dataDxfId. Если есть связанный
    # queryTable*.xml, дополнительно синхронизируем id/name/queryTableFieldId
    # с queryTableField, потому что Excel проверяет эту связь.
    if len(old_columns) == len(headers):
        fields = query_fields or []
        for idx, col in enumerate(old_columns):
            if fields and idx < len(fields):
                field = fields[idx]
                tcid = str(field.get("tableColumnId") or field.get("id") or idx + 1)
                qfid = str(field.get("id") or tcid)
                qname = str(field.get("name") or headers[idx])
                col.attrib["id"] = tcid
                col.attrib["name"] = qname
                if "uniqueName" in col.attrib:
                    col.attrib["uniqueName"] = tcid
                col.attrib["queryTableFieldId"] = qfid
            else:
                col.attrib["name"] = str(headers[idx])
        table_columns.attrib["count"] = str(len(headers))
        return xml_tostring_excel_safe(root)

    # Если количество колонок всё-таки изменилось, пересобираем, но при наличии
    # queryTable берём id/name из queryTableField, а не генерируем их заново.
    fields = query_fields or []
    for child in list(table_columns):
        table_columns.remove(child)

    for idx, header in enumerate(headers):
        if idx < len(old_columns):
            col = copy.deepcopy(old_columns[idx])
        elif old_columns:
            col = copy.deepcopy(old_columns[-1])
        else:
            col = ET.Element(xq("tableColumn"))

        if fields and idx < len(fields):
            field = fields[idx]
            tcid = str(field.get("tableColumnId") or field.get("id") or idx + 1)
            qfid = str(field.get("id") or tcid)
            qname = str(field.get("name") or header)
            col.attrib["id"] = tcid
            col.attrib["name"] = qname
            if "uniqueName" in col.attrib:
                col.attrib["uniqueName"] = tcid
            col.attrib["queryTableFieldId"] = qfid
        else:
            # Последний резервный вариант для не-queryTable.
            col.attrib["id"] = str(idx + 1)
            col.attrib["name"] = str(header)
            if "uniqueName" in col.attrib:
                col.attrib["uniqueName"] = str(idx + 1)
            if "queryTableFieldId" in col.attrib:
                col.attrib["queryTableFieldId"] = str(idx + 1)
        table_columns.append(col)

    table_columns.attrib["count"] = str(len(headers))
    return xml_tostring_excel_safe(root)


def force_workbook_recalculation(workbook_xml: bytes) -> bytes:
    """Просит Excel пересчитать формулы при открытии."""
    root = ET.fromstring(workbook_xml)
    calc_pr = root.find(xq("calcPr"))
    if calc_pr is None:
        calc_pr = ET.SubElement(root, xq("calcPr"))
    calc_pr.attrib["calcMode"] = "auto"
    calc_pr.attrib["fullCalcOnLoad"] = "1"
    calc_pr.attrib["forceFullCalc"] = "1"
    return xml_tostring_excel_safe(root)


def validate_xlsx_table_xml(xlsx_path: Path) -> None:
    """Проверяет XML-структуру workbook/sheets/tables после сохранения."""
    errors: List[str] = []
    with zipfile.ZipFile(xlsx_path) as zf:
        bad = zf.testzip()
        if bad:
            errors.append(f"ZIP CRC error: {bad}")
        for name in zf.namelist():
            if name.endswith(".xml") or name.endswith(".rels"):
                try:
                    ET.fromstring(zf.read(name))
                except Exception as exc:
                    errors.append(f"Некорректный XML {name}: {exc}")
                    continue
                # Строгая проверка для Excel: все префиксы из mc:Ignorable должны быть объявлены.
                try:
                    xml_text = zf.read(name).decode("utf-8", errors="replace")
                    root_tag = re.search(r"<([A-Za-z_][\w:.-]*)(\s[^<>]*?)?>", xml_text)
                    if root_tag:
                        root_tag_text = root_tag.group(0)
                        ign = re.search(r'\bmc:Ignorable="([^"]+)"', root_tag_text)
                        if ign:
                            declared = set(re.findall(r'\bxmlns:([A-Za-z_][\w.-]*)="', root_tag_text))
                            missing = [p for p in ign.group(1).split() if p not in declared]
                            if missing:
                                errors.append(f"{name}: mc:Ignorable ссылается на необъявленные префиксы: {missing}")
                except Exception as exc:
                    errors.append(f"Не удалось проверить mc:Ignorable в {name}: {exc}")
            if not (name.startswith("xl/tables/table") and name.endswith(".xml")):
                continue
            xml = zf.read(name).decode("utf-8", errors="replace")
            ref_match = re.search(r'ref="([^"]+)"', xml)
            count_match = re.search(r'<(?:\w+:)?tableColumns\b[^>]*count="(\d+)"', xml)
            if not ref_match or not count_match:
                continue
            ref = ref_match.group(1)
            min_col, _, max_col, _ = parse_range_ref(ref)
            ref_col_count = max_col - min_col + 1
            xml_count = int(count_match.group(1))
            # Ищем только сами tableColumn, не tableColumns.
            actual_count = len(re.findall(r'<(?:\w+:)?tableColumn\b', xml))
            if xml_count != ref_col_count or actual_count != ref_col_count:
                errors.append(
                    f"{name}: ref={ref}, columns_in_ref={ref_col_count}, "
                    f"tableColumns_count={xml_count}, actual_tableColumn={actual_count}"
                )

        # Проверяем связку tableColumn <-> queryTableField. Обычные XML-парсеры
        # это не считают ошибкой, но Excel считает такие книги повреждёнными.
        names_set = set(zf.namelist())
        import posixpath
        for table_name in [n for n in names_set if n.startswith("xl/tables/table") and n.endswith(".xml")]:
            rel_path = posixpath.join(posixpath.dirname(table_name), "_rels", posixpath.basename(table_name) + ".rels")
            if rel_path not in names_set:
                continue
            try:
                table_root = ET.fromstring(zf.read(table_name))
                rel_root = ET.fromstring(zf.read(rel_path))
            except Exception:
                continue
            qt_path = ""
            for rel in rel_root:
                target = rel.attrib.get("Target", "")
                rel_type = rel.attrib.get("Type", "")
                if "queryTable" in rel_type or "queryTables/" in target:
                    qt_path = normalize_zip_target(posixpath.dirname(table_name), target)
                    break
            if not qt_path or qt_path not in names_set:
                continue
            try:
                qt_root = ET.fromstring(zf.read(qt_path))
            except Exception:
                continue
            table_cols = {
                c.attrib.get("id", ""): c.attrib.get("name", "")
                for c in table_root.findall(f".//{xq('tableColumn')}")
            }
            for field in qt_root.findall(f".//{xq('queryTableField')}"):
                tcid = field.attrib.get("tableColumnId", "")
                qname = field.attrib.get("name", "")
                if tcid not in table_cols:
                    errors.append(f"{table_name}: queryTable {qt_path} ссылается на отсутствующий tableColumnId={tcid} ({qname})")
                elif table_cols[tcid] != qname:
                    errors.append(f"{table_name}: tableColumnId={tcid} name mismatch: table={table_cols[tcid]!r}, queryTable={qname!r}")
    if errors:
        details = "\n".join(errors[:20])
        raise ValueError("После сохранения обнаружена неконсистентная XML-разметка XLSX:\n" + details)


def validate_only_worksheet_xml_changed(source_path: Path, output_path: Path) -> None:
    """
    Проверяет инвариант безопасной записи: меняются только XML листов.

    Power Query книги чувствительны к xl/tables, xl/queryTables, connections,
    customXml/DataMashup и workbook metadata. Если какая-то из этих частей
    отличается от исходной книги, Excel может открыть файл как повреждённый
    и удалить внешние диапазоны данных.
    """
    allowed = re.compile(r"^xl/worksheets/sheet\d+\.xml$")
    errors: List[str] = []
    with zipfile.ZipFile(source_path) as src, zipfile.ZipFile(output_path) as out:
        src_names = set(src.namelist())
        out_names = set(out.namelist())
        if src_names != out_names:
            missing = sorted(src_names - out_names)
            extra = sorted(out_names - src_names)
            if missing:
                errors.append(f"В выходном XLSX отсутствуют части: {missing[:10]}")
            if extra:
                errors.append(f"В выходном XLSX появились лишние части: {extra[:10]}")

        for name in sorted(src_names & out_names):
            if allowed.fullmatch(name):
                continue
            if src.read(name) != out.read(name):
                errors.append(f"Недопустимо изменена часть XLSX: {name}")
                if len(errors) >= 20:
                    break
    if errors:
        raise ValueError(
            "После сохранения изменились служебные части XLSX, которые должны оставаться byte-for-byte неизменными:\n"
            + "\n".join(errors)
        )


# -----------------------------------------------------------------------------
# Запись результатов в те же таблицы Excel
# -----------------------------------------------------------------------------

def write_dataframe_to_excel_table(
    xlsx_path: Path,
    output_path: Path,
    loaded: List[LoadedQuery],
    results: Dict[str, pd.DataFrame],
    logger: Optional[logging.Logger] = None,
    debug_dir: Optional[Path] = None,
) -> None:
    """
    Записывает результаты в .xlsx напрямую через ZIP/XML.

    В этой версии меняются только XML нужных листов. XML Excel-таблиц
    (xl/tables/table*.xml), queryTables, connections и Power Query metadata
    остаются byte-for-byte как в исходном файле. Это самый безопасный режим для
    Excel на macOS/Windows: он не удаляет диапазоны внешних данных при открытии.
    """
    logger = logger or logging.getLogger("pq2py")
    logger.info("Открываю XLSX как ZIP/XML для безопасной записи: %s", xlsx_path)

    with zipfile.ZipFile(xlsx_path, "r") as zin:
        zip_bytes: Dict[str, bytes] = {name: zin.read(name) for name in zin.namelist()}
        zip_infos: Dict[str, zipfile.ZipInfo] = {info.filename: info for info in zin.infolist()}
        zip_order: List[str] = zin.namelist()

    sheet_to_path, table_map = worksheet_and_table_maps(zip_bytes)
    logger.debug("Worksheet map: %s", sheet_to_path)
    logger.debug("Table map keys: %s", sorted([f"{k[0]}::{k[1]}" for k in table_map]))

    for item in loaded:
        if item.query_name not in results:
            continue
        df = results[item.query_name].copy()
        logger.info("Запись таблицы: query=%s -> %s!%s", item.query_name, item.sheet_name, item.table_ref)
        logger.debug("DataFrame before Excel write: %s", df_debug_summary(df))
        df = df.where(pd.notna(df), None)

        sheet_path = sheet_to_path.get(item.sheet_name)
        if not sheet_path or sheet_path not in zip_bytes:
            raise KeyError(f"Не найден XML листа {item.sheet_name!r}")
        table_path = table_map.get((item.sheet_name, item.table_name))
        if not table_path or table_path not in zip_bytes:
            # Иногда имя таблицы из connections совпадает с name/displayName не полностью.
            candidates = [k for k in table_map if k[0] == item.sheet_name]
            raise KeyError(f"На листе {item.sheet_name!r} не найдена таблица {item.table_name!r}. Доступны: {candidates}")

        table_root_before = ET.fromstring(zip_bytes[table_path])
        old_ref = table_root_before.attrib.get("ref", item.table_ref)
        old_columns = [c.attrib.get("name", "") for c in table_root_before.findall(f".//{xq('tableColumn')}")]
        logger.debug("Excel table before XML patch: sheet=%s table=%s path=%s ref=%s columns=%s", item.sheet_name, item.table_name, table_path, old_ref, old_columns)

        query_fields = querytable_fields_for_table(zip_bytes, table_path)
        query_headers = [f.get("name", "") for f in query_fields]
        original_headers = old_columns
        df_headers = make_excel_table_headers_unique([str(c) for c in df.columns])

        # Для Power Query-таблиц заголовки должны совпадать с queryTable*.xml.
        # Поэтому при совпадении количества колонок пишем в лист именно исходные
        # Excel/queryTable-заголовки, а данные берём из DataFrame по позиции.
        if query_headers and len(query_headers) == df.shape[1]:
            headers = query_headers
        elif original_headers and len(original_headers) == df.shape[1]:
            headers = original_headers
        else:
            headers = df_headers
            logger.warning(
                "Количество колонок DataFrame (%s) не совпало с Excel/queryTable (%s/%s), "
                "использую заголовки DataFrame: %s",
                df.shape[1], len(original_headers), len(query_headers), headers,
            )

        data_rows = df.values.tolist()

        # Ключевое отличие от предыдущих версий: НЕ меняем table XML и queryTable XML.
        # Excel удалял queryTables при открытии именно из-за ручного изменения этих
        # служебных частей. Поэтому оставляем старый ref таблицы и только очищаем/
        # перезаписываем значения внутри уже существующего диапазона.
        zip_bytes[sheet_path] = write_values_to_sheet_xml_preserve_table_range(
            zip_bytes[sheet_path],
            old_ref,
            headers,
            data_rows,
            logger=logger,
            context=f"{item.query_name} -> {item.sheet_name}!{old_ref}",
        )
        logger.debug(
            "Excel table XML preserved byte-for-byte: sheet=%s table=%s path=%s ref=%s headers=%s query_fields=%s",
            item.sheet_name, item.table_name, table_path, old_ref, headers, query_fields,
        )

    # Не трогаем xl/workbook.xml и calcChain.xml: изменение workbook.xml через XML-парсер
    # может повлиять на Excel-specific metadata. При необходимости пересчёт можно выполнить
    # в Excel вручную: Cmd+Option+Shift+F9.

    output_path.parent.mkdir(parents=True, exist_ok=True)
    logger.info("Сохраняю книгу ZIP/XML без переписывания неподронутых частей: %s", output_path)
    with zipfile.ZipFile(output_path, "w") as zout:
        written = set()
        for name in zip_order:
            info = zip_infos[name]
            data = zip_bytes[name]
            new_info = zipfile.ZipInfo(filename=name, date_time=info.date_time)
            new_info.compress_type = info.compress_type
            new_info.comment = info.comment
            new_info.extra = info.extra
            new_info.internal_attr = info.internal_attr
            new_info.external_attr = info.external_attr
            new_info.create_system = info.create_system
            zout.writestr(new_info, data)
            written.add(name)
        for name, data in zip_bytes.items():
            if name not in written:
                zout.writestr(name, data)

    logger.info("Книга сохранена, размер: %.2f MB", output_path.stat().st_size / 1024 / 1024)
    if debug_dir is not None:
        Path(debug_dir).mkdir(parents=True, exist_ok=True)
        diag_path = Path(debug_dir) / "xlsx_zip_diagnostics.txt"
        xlsx_zip_diagnostics(output_path, diag_path)
        logger.info("Диагностика XML/ZIP сохранена: %s", diag_path)
    validate_xlsx_table_xml(output_path)
    validate_only_worksheet_xml_changed(xlsx_path, output_path)
    logger.info("Проверка XML/ZIP XLSX пройдена")
    print("Проверка XML/ZIP XLSX пройдена")


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Выполнить Power Query этой книги на Python и заполнить те же Excel-таблицы.")
    parser.add_argument("xlsx", type=Path, help="Исходный .xlsx файл")
    parser.add_argument("-o", "--output", type=Path, default=None, help="Куда сохранить заполненную копию")
    parser.add_argument("--regnum", type=str, default=None, help="Регистрационный номер банка ЦБ РФ, которым заменить regnum из M-кода")
    parser.add_argument("--list", action="store_true", help="Только показать, какие запросы загружаются в какие таблицы")
    parser.add_argument("--dump-m", action="store_true", help="Сохранить извлечённый Section1.m рядом с файлом")
    parser.add_argument("--no-cache", action="store_true", help="Не использовать HTML-кэш")
    parser.add_argument("--cache-dir", type=Path, default=None, help="Папка HTML-кэша")
    parser.add_argument("--verbose", action="store_true", help="Подробный вывод в консоль и текстовый лог")
    parser.add_argument("--debug", action="store_true", help="Максимальная диагностика: подробный лог, снимки DataFrame и XML-отчёт")
    parser.add_argument("--log-file", type=Path, default=None, help="Путь к текстовому логу. По умолчанию рядом с xlsx при --verbose/--debug")
    parser.add_argument("--debug-dir", type=Path, default=None, help="Папка для debug-артефактов. По умолчанию рядом с xlsx при --debug")
    args = parser.parse_args(argv)

    if args.regnum is not None:
        args.regnum = args.regnum.strip()
        if not re.fullmatch(r"\d+", args.regnum):
            parser.error("--regnum должен состоять только из цифр, например: --regnum 2673")

    xlsx_path = args.xlsx.resolve()
    if args.output is None:
        suffix = f"_regnum_{args.regnum}_python_filled.xlsx" if args.regnum else "_python_filled.xlsx"
        output_path = xlsx_path.with_name(xlsx_path.stem + suffix)
    else:
        output_path = args.output.resolve()
    cache_dir = args.cache_dir or xlsx_path.with_name("pq_html_cache")
    debug_dir = args.debug_dir.resolve() if args.debug_dir else (xlsx_path.with_name(xlsx_path.stem + "_debug") if args.debug else None)
    if debug_dir is not None:
        debug_dir.mkdir(parents=True, exist_ok=True)

    logger, log_path = setup_run_logging(
        xlsx_path=xlsx_path,
        log_file=args.log_file,
        verbose=args.verbose or args.debug,
        debug=args.debug,
    )
    logger.info("Старт скрипта")
    logger.info("Аргументы: xlsx=%s output=%s regnum=%s list=%s dump_m=%s no_cache=%s cache_dir=%s verbose=%s debug=%s debug_dir=%s",
                xlsx_path, output_path, args.regnum, args.list, args.dump_m, args.no_cache, cache_dir, args.verbose, args.debug, debug_dir)
    logger.info("Python: %s", sys.version.replace("\n", " "))
    logger.info("pandas=%s requests=%s openpyxl loaded", pd.__version__, requests.__version__)

    inspector = PowerQueryWorkbookInspector(xlsx_path)
    m_code = inspector.extract_m_code()
    logger.info("M-код извлечён: символов=%d, shared-запросов=%d", len(m_code), len(re.findall(r"(?m)^shared\s+", m_code)))
    if debug_dir is not None:
        (debug_dir / "Section1_effective.m").write_text(m_code, encoding="utf-8")
    original_regnums = detect_regnums_in_m_code(m_code)
    if args.regnum:
        m_code = replace_regnum_in_m_code(m_code, args.regnum)
        print(f"regnum в URL заменён: {', '.join(original_regnums) if original_regnums else 'не найден'} -> {args.regnum}")
    else:
        print(f"regnum из M-кода: {', '.join(original_regnums) if original_regnums else 'не найден'}")
    loaded = inspector.list_loaded_queries()
    logger.info("Загружаемых Excel-таблиц найдено: %d", len(loaded))
    if debug_dir is not None:
        (debug_dir / "loaded_queries.json").write_text(
            json.dumps([item.__dict__ for item in loaded], ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    if args.dump_m:
        m_path = xlsx_path.with_suffix(".Section1.m")
        m_path.write_text(m_code, encoding="utf-8")
        print(f"M-код сохранён: {m_path}")

    if args.list:
        print(f"Найдено загружаемых таблиц: {len(loaded)}")
        for item in loaded:
            print(f"- {item.query_name} -> {item.sheet_name}!{item.table_ref} [{item.table_name}]")
        return 0

    engine = MiniMEngine(
        m_code,
        cache_dir=cache_dir,
        use_cache=not args.no_cache,
        logger=logger,
        verbose=args.verbose or args.debug,
        debug=args.debug,
        debug_dir=debug_dir,
    )
    results: Dict[str, pd.DataFrame] = {}
    print(f"Найдено загружаемых таблиц: {len(loaded)}")
    for item in loaded:
        print(f"Выполняю {item.query_name} -> {item.sheet_name}!{item.table_ref} ...", flush=True)
        try:
            df = engine.evaluate_query(item.query_name)
        except Exception:
            logger.error("Ошибка при выполнении загружаемого запроса %s", item.query_name)
            logger.error(traceback.format_exc())
            raise
        results[item.query_name] = df
        logger.info("Загружаемый запрос готов: %s -> %s", item.query_name, df_debug_summary(df))
        print(f"  строк: {len(df)}, столбцов: {len(df.columns)}")

    try:
        write_dataframe_to_excel_table(xlsx_path, output_path, loaded, results, logger=logger, debug_dir=debug_dir)
    except Exception:
        logger.error("Ошибка при записи/валидации Excel-файла")
        logger.error(traceback.format_exc())
        raise
    print(f"Готово: {output_path}")
    if log_path is not None:
        print(f"Лог выполнения: {log_path}")
    if debug_dir is not None:
        print(f"Debug-папка: {debug_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
