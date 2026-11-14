#!/bin/bash
# Chạy thử trên macOS/Linux: cần Python 3.12.
cd "$(dirname "$0")"
if [ ! -d .venv ]; then python3.12 -m venv .venv && .venv/bin/python -m pip install -r requirements.txt; fi
.venv/bin/python -m ui.giao_dien
