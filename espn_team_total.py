# Launcher: run 爬取ESPN各隊Total.py (ASCII filename for .bat)
import os
import sys

_dir = os.path.dirname(os.path.abspath(__file__))
_main = os.path.join(_dir, "爬取ESPN各隊Total.py")
if not os.path.exists(_main):
    print("Error: 爬取ESPN各隊Total.py not found")
    sys.exit(1)

with open(_main, encoding="utf-8") as f:
    exec(compile(f.read(), _main, "exec"))

