# Launcher: run 爬取今日先發名單 (ASCII filename for .bat)
import os
import sys
_dir = os.path.dirname(os.path.abspath(__file__))
_main = os.path.join(_dir, '爬取今日先發名單.py')
if not os.path.exists(_main):
    print('Error: 爬取今日先發名單.py not found')
    sys.exit(1)
with open(_main, encoding='utf-8') as f:
    exec(compile(f.read(), _main, 'exec'))
