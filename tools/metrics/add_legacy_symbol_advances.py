# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.
"""Refresh legacy symbol advances from the Windows symbol cmap, without font files."""
import json
import os
from pathlib import Path
from fontTools.ttLib import TTFont


def main():
    target = Path(__file__).resolve().parents[2] / 'crates/oxidocs-core/src/font/data/font_metrics_compact.json'
    rows = json.loads(target.read_text(encoding='utf-8'))
    fonts = Path(os.environ.get('WINDIR', 'C:/Windows')) / 'Fonts'
    for family, filename in [('Symbol', 'symbol.ttf'), ('Wingdings', 'wingding.ttf')]:
        row = next(r for r in rows if r['family'] == family)
        with TTFont(fonts / filename) as font:
            assert row['units_per_em'] == font['head'].unitsPerEm
            cmap = font['cmap'].getcmap(3, 0)
            if cmap is None:
                raise ValueError(f'{family}: Windows symbol cmap missing')
            widths = {str(cp): font['hmtx'].metrics[glyph][0]
                      for cp, glyph in cmap.cmap.items() if 0xF000 <= cp <= 0xF0FF}
            row['widths'].update(widths)
            print(f'{family}: {len(widths)} symbol advances')
    target.write_text(json.dumps(rows, ensure_ascii=False, separators=(',', ':')) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
