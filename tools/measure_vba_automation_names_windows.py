# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.
"""Export portable Automation name weights and ANSI encoding from Windows NLS.

Writes only the explicitly supplied data directory and vector JSON path.
Does not start Office or change machine settings.
"""
import argparse
import ctypes
import ctypes.wintypes as w
import json
from pathlib import Path
import random
import struct


def measure(directory, vectors):
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    encode = kernel.WideCharToMultiByte
    encode.argtypes = [w.UINT, w.DWORD, w.LPCWSTR, ctypes.c_int,
                       ctypes.c_void_p, ctypes.c_int, ctypes.c_void_p,
                       ctypes.POINTER(w.BOOL)]
    encode.restype = ctypes.c_int
    sort = kernel.LCMapStringA
    sort.argtypes = [w.DWORD, w.DWORD, ctypes.c_char_p, ctypes.c_int,
                     ctypes.c_void_p, ctypes.c_int]
    sort.restype = ctypes.c_int

    def sort_key(source):
        output = ctypes.create_string_buffer(len(source) * 16 + 64)
        count = sort(0, 0x401, source, len(source), output, len(output))
        assert count, ctypes.get_last_error()
        return output.raw[:count]

    directory.mkdir(parents=True, exist_ok=True)
    weights = bytearray(10)  # NUL terminates Automation names.
    for byte in range(1, 256):
        levels = sort_key(bytes([byte]))[:-1].split(b'\x01')
        assert len(levels) == 5
        primary, secondary, tertiary, unused, special = levels
        assert len(primary) in (0, 2, 4)
        assert len(secondary) <= 1 and len(tertiary) <= 1 and not unused
        assert not special or len(special) == 4 and special[:2] == b'\xff\xff'
        weights.extend(bytes([len(primary)]) + primary.ljust(4, b'\0')
                       + secondary.ljust(1, b'\0') + tertiary.ljust(1, b'\0')
                       + bytes([bool(special)]) + special[2:].ljust(2, b'\0'))
    (directory / 'vba_automation_weights.bin').write_bytes(weights)

    counts = {}
    for page in [932, 1252]:
        mappings = bytearray()
        for code in range(128, 65536):
            if 0xd800 <= code <= 0xdfff:
                continue
            output = ctypes.create_string_buffer(8)
            used = w.BOOL()
            count = encode(page, 0, chr(code), 1, output, len(output), None,
                           ctypes.byref(used))
            assert count in (1, 2), (page, code, count)
            if not used.value:
                encoded = int.from_bytes(output.raw[:count], 'big')
                mappings.extend(struct.pack('<HH', code, encoded))
        (directory / f'vba_automation_cp{page}.bin').write_bytes(mappings)
        counts[page] = len(mappings) // 4

    rng = random.Random(20261018)
    sources = [bytes([i]) for i in range(1, 256)]
    sources.extend(bytes(rng.randrange(1, 256) for _ in range(rng.randrange(2, 30)))
                   for _ in range(128))
    vectors.parent.mkdir(parents=True, exist_ok=True)
    vectors.write_text(json.dumps([dict(input=list(s), key=list(sort_key(s)))
                                   for s in sources], indent=2) + '\n', encoding='utf-8')
    print(json.dumps(dict(encoding_mappings=counts, sort_key_vectors=len(sources))))


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output-directory', type=Path, required=True)
    parser.add_argument('--vectors', type=Path, required=True)
    args = parser.parse_args()
    measure(args.output_directory, args.vectors)
