# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.
"""Recompile and verify Automation name dispatch in owned native Excel.

Requires caller-enabled VBA project access; never changes trust settings.
Each source fixture is exported again and each entry called three times.
"""
import argparse
from pathlib import Path
from measure_vba_unicode_modules_com import measure

if __name__ == '__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output',type=Path,required=True)
    parser.add_argument('--names',nargs='+')
    args=parser.parse_args()
    measure(args.output,args.names,fixture_name='automation_names.json')
