# -*- coding: utf-8 -*-
"""S559 — run the Phase-1 pagination gate for 3a4f under a given
OXI_S559_CELLMAR mode: regenerate the Oxi pagination json with the env set,
then invoke pagination_diff for 3a4f. Usage: python _s559_paginate.py <mode>
"""
import os
import subprocess
import sys

REPO = r'c:\Users\ryuji\oxi-main'
mode = sys.argv[1] if len(sys.argv) > 1 else 'OFF'
env = dict(os.environ)
if mode == 'OFF':
    env.pop('OXI_S559_CELLMAR', None)
else:
    env['OXI_S559_CELLMAR'] = mode

print('=== mode=%s : regen oxi pagination ===' % mode)
subprocess.run([sys.executable, os.path.join(REPO, 'tools', 'metrics', 'measure_pagination_oxi.py'), '3a4f'],
               env=env, cwd=REPO)
print('=== diff ===')
subprocess.run([sys.executable, os.path.join(REPO, 'tools', 'metrics', 'pagination_diff.py'), '3a4f'],
               cwd=REPO)
