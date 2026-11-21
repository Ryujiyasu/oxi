# -*- coding: utf-8 -*-
"""oxi-xll-host as a browser's native-messaging host: started the way Chrome
starts one (the extension's origin as its argument), fed length-prefixed
JSON on stdin, answering on stdout. Lists the test XLL and calls a scalar, a
range and an FP12 function through it, then an unknown one.

    python tools/metrics/_xll_native_messaging_check.py
"""
import json
import struct
import subprocess
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
HOST = REPO / "tools/oxi-xll-host/target/release/oxi-xll-host.exe"
XLL = str(REPO / "tools/oxi-xll-host/testxll/target/release/oxi_test.xll")


def main():
    host = subprocess.Popen([str(HOST), "chrome-extension://oxi-test/"], stdin=subprocess.PIPE, stdout=subprocess.PIPE)

    def ask(message):
        body = json.dumps(message).encode("utf-8")
        host.stdin.write(struct.pack("<I", len(body)) + body)
        host.stdin.flush()
        size = struct.unpack("<I", host.stdout.read(4))[0]
        return json.loads(host.stdout.read(size))

    cases = [
        ({"id": 1, "xll": XLL, "op": "list"}, lambda a: a["id"] == 1 and any(f["function"] == "OxiAdd" for f in a["functions"])),
        ({"id": 2, "xll": XLL, "op": "call", "function": "OxiAdd", "args": [2, 3]}, lambda a: a == {"id": 2, "value": 5.0}),
        ({"id": 3, "xll": XLL, "op": "call", "function": "OxiRangeSum", "args": [[[1, 2], [3, "x"]]]}, lambda a: a["value"] == 6.0),
        ({"id": 4, "xll": XLL, "op": "call", "function": "OxiFpTwice", "args": [[[1, 2]]]}, lambda a: a["value"] == [[2.0, 4.0]]),
        ({"id": 5, "xll": XLL, "op": "call", "function": "Nope", "args": []}, lambda a: a == {"id": 5, "error": "#NAME?"}),
        ({"id": 6, "xll": "C:/no/such.xll", "op": "list"}, lambda a: a["id"] == 6 and "error" in a),
    ]
    ok = 0
    try:
        for message, good in cases:
            answer = ask(message)
            fine = good(answer)
            ok += fine
            print(f"  {message.get('op')} {message.get('function', '')!s:12s} -> {json.dumps(answer)[:90]} {'OK' if fine else 'DIFF'}")
    finally:
        host.stdin.close()
        host.wait(timeout=10)
    print(f"{ok}/{len(cases)} native-messaging exchanges answered")


if __name__ == "__main__":
    main()
