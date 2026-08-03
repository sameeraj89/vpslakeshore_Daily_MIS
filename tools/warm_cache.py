#!/usr/bin/env python3
"""Warm the per-flash-file parse cache under a time budget.
Usage: python3 warm_cache.py <folder> [budget_seconds]
Run repeatedly until it prints 'remaining 0'."""
import sys, os, glob, json, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_dashboard import file_date, parse_flash_details

folder = sys.argv[1]
budget = float(sys.argv[2]) if len(sys.argv) > 2 else 30
tools = os.path.join(folder, "_dashboard_tools")
path = os.path.join(tools, "flash_cache.json")
cache = {}
if os.path.exists(path):
    try: cache = json.load(open(path))
    except Exception: cache = {}

by_date = {}
for f in glob.glob(os.path.join(folder, "**", "Daily Revenue Flash_New_*.xlsx"), recursive=True):
    d = file_date(f)
    if d and (d not in by_date or os.path.getmtime(f) > os.path.getmtime(by_date[d])):
        by_date[d] = f

t0, done, remaining = time.time(), 0, 0
for d in sorted(by_date):
    f = by_date[d]
    ck = f"{os.path.basename(f)}|{int(os.path.getmtime(f))}"
    if ck in cache: continue
    if time.time() - t0 > budget:
        remaining += 1; continue
    cache[ck] = parse_flash_details(f)
    done += 1

json.dump(cache, open(path, "w"))
print(f"cached {done} this run; remaining {remaining}")
