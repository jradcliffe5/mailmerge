"""Enable `python -m mailmerge_cli` execution."""

from __future__ import annotations

from .cli import main


if __name__ == "__main__":
    raise SystemExit(main())
