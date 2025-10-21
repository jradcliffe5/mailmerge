#!/usr/bin/env python3
"""Backward-compatible CLI wrapper for the mailmerge package."""

from __future__ import annotations

from mailmerge_cli.cli import main


if __name__ == "__main__":
    raise SystemExit(main())
