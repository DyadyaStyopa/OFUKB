#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Bundled CLI entrypoint for the macOS app."""

from __future__ import annotations

import sys
from typing import List, Optional


def main(argv: Optional[List[str]] = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    mode = args.pop(0) if args and args[0] in {"excel", "sqlite"} else "excel"

    if mode == "sqlite":
        from cbr_sqlite_export import main as sqlite_main

        return int(sqlite_main(args) or 0)

    from OFUKB_CBR_PQ_alt_parser import main as excel_main

    return int(excel_main(args) or 0)


if __name__ == "__main__":
    raise SystemExit(main())
