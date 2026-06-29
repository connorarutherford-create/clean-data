#!/usr/bin/env python3
"""drop.py — Process a file from the drop folder.

Usage:
    python3 drop.py filename.csv
    python3 drop.py filename.xlsx

Drop a file in ~/Desktop/Hermes/Inbox/, run this, get the cleaned
version in ~/Desktop/Hermes/Outbox/.
"""

import os, sys, subprocess

INBOX = os.path.expanduser("~/Desktop/Hermes/Inbox")
OUTBOX = os.path.expanduser("~/Desktop/Hermes/Outbox")
SCRIPT = os.path.expanduser("~/Desktop/Hermes/Projects/clean-data/clean_data.py")

def main():
    os.makedirs(INBOX, exist_ok=True)
    os.makedirs(OUTBOX, exist_ok=True)
    
    if len(sys.argv) > 1:
        files = [os.path.join(INBOX, sys.argv[1])]
    else:
        files = [os.path.join(INBOX, f) for f in os.listdir(INBOX) 
                 if f.endswith(('.csv','.xlsx','.xls'))]
    
    if not files:
        print(f"No files in {INBOX}/")
        print("Drop a .csv or .xlsx file there and run: python3 drop.py filename.csv")
        return
    
    for f in files:
        if not os.path.exists(f):
            print(f"Not found: {f}")
            continue
        
        name = os.path.basename(f)
        base, ext = os.path.splitext(name)
        out = os.path.join(OUTBOX, f"{base}_clean{ext}")
        
        print(f"Processing: {name}")
        r = subprocess.run(
            ["python3", SCRIPT, f, "-o", out, "--auto", "--summary"],
            capture_output=True, text=True, timeout=30
        )
        print(r.stdout)
        if r.stderr:
            print(f"Error: {r.stderr}")

if __name__ == "__main__":
    main()
