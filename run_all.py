#!/usr/bin/env python3
"""run_all.py — run build_data.py then build_excel.py; --open launches index.html"""

import subprocess, sys, os, argparse, json, re

HERE = os.path.dirname(os.path.abspath(__file__))

def run(script, label):
    print(f"\n{'='*60}")
    print(f"  {label}")
    print(f"{'='*60}")
    result = subprocess.run(
        [sys.executable, os.path.join(HERE, script)],
        cwd=HERE
    )
    if result.returncode != 0:
        print(f"\nERROR: {script} failed (exit {result.returncode}). Stopping.")
        sys.exit(result.returncode)

def latest_week_summary():
    data_js = os.path.join(HERE, "data.js")
    if not os.path.exists(data_js):
        return
    with open(data_js, "r", encoding="utf-8") as f:
        js = f.read()
    json_str = re.sub(r"^window\.DASHBOARD\s*=\s*", "", js.strip()).rstrip(";").strip()
    try:
        D = json.loads(json_str)
        print("\n--- Latest-Week Summary ---")
        for s in D.get("sheets", []):
            print(f"  {s.get('name',''):<42} {s.get('latest_week','?')}")
    except Exception:
        pass

def main():
    ap = argparse.ArgumentParser(description="Build UDS Dashboard")
    ap.add_argument("--open", action="store_true", help="Open index.html in browser after build")
    args = ap.parse_args()

    run("build_data.py",  "Step 1 - Parsing workbook -> data.js")
    run("build_excel.py", "Step 2 - Building UDS_Dashboard.xlsx")

    latest_week_summary()

    html_path = os.path.join(HERE, "index.html")
    xlsx_path = os.path.join(HERE, "UDS_Dashboard.xlsx")
    js_path   = os.path.join(HERE, "data.js")

    print("\n--- Output Paths ---")
    print(f"  data.js          -> {js_path}")
    print(f"  index.html       -> {html_path}")
    print(f"  UDS_Dashboard.xlsx -> {xlsx_path}")

    if args.open:
        import webbrowser
        webbrowser.open(f"file:///{html_path.replace(os.sep, '/')}")

if __name__ == "__main__":
    main()
