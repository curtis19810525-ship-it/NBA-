# Launcher: run main script (ASCII filename for .bat encoding)
import os
import sys
_dir = os.path.dirname(os.path.abspath(__file__))
_main = os.path.join(_dir, '爬取ESPN球員狀態.py')
if not os.path.exists(_main):
    print('Error: 爬取ESPN球員狀態.py not found')
    sys.exit(1)
with open(_main, encoding='utf-8') as f:
    exec(compile(f.read(), _main, 'exec'))
