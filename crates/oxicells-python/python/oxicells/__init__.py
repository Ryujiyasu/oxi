# SPDX-License-Identifier: MIT OR Apache-2.0
"""Run supported workbook VBA through the persistent OxiCells native engine."""
from __future__ import annotations

import datetime as _datetime
import decimal
import json
import math
import os
from pathlib import Path
import queue
import secrets
import shutil
import struct
import subprocess
import sysconfig
import threading
from dataclasses import dataclass
from typing import Any, Optional


class OxiCellsError(RuntimeError):
    """The engine could not complete a request."""


class MacroError(OxiCellsError):
    def __init__(self, details: dict[str, Any]) -> None:
        self.kind = details.get("kind")
        self.line = details.get("line")
        self.details = details
        super().__init__(details.get("message", "VBA execution failed"))


class Byte(int): pass
class Integer(int): pass
class Long(int): pass
class LongLong(int): pass
class Single(float): pass
class Date(float):
    """A VBA Date represented by its exact OLE date serial."""
class Currency(decimal.Decimal): pass


@dataclass(frozen=True)
class _Unit:
    name: str


EMPTY = _Unit("empty")
MISSING = _Unit("missing")
NOTHING = _Unit("nothing")
NULL_STRING = _Unit("null_string")


@dataclass(frozen=True)
class BStr:
    """A BSTR's raw bytes, including a possible final unpaired byte."""
    data: bytes


@dataclass(frozen=True)
class ErrorValue:
    code: int


@dataclass(frozen=True)
class Array:
    """Flat row-major values with their original VBA bounds and default type."""
    values: tuple[Any, ...]
    dimensions: tuple[tuple[int, int], ...]
    element_default: Any = EMPTY
    resizable: bool = True

    def at(self, *indices: int) -> Any:
        if len(indices) != len(self.dimensions):
            raise IndexError("array rank does not match supplied indices")
        offset = 0
        for index, (lower, length) in zip(indices, self.dimensions):
            if not lower <= index < lower + length:
                raise IndexError("VBA array subscript is out of range")
            offset = offset * length + index - lower
        return self.values[offset]


@dataclass(frozen=True)
class Record:
    type_name: str
    fields: tuple[tuple[str, Any], ...]


class Object:
    def __init__(self, workbook: Workbook, descriptor: dict[str, Any]) -> None:
        self._workbook = workbook
        self._descriptor = descriptor
        self._released = False

    def invoke(self, member: str, *args: Any, kind: int = 1) -> Any:
        if self._released:
            raise OxiCellsError("VBA object reference has been released")
        return self._workbook._call("invoke", reference=self._descriptor["reference"],
                                    member=member, kind=kind, args=[_encode(v) for v in args])

    def get(self, member: str, *indices: Any) -> Any:
        return self.invoke(member, *indices, kind=2)

    def set(self, member: str, value: Any, *indices: Any) -> Any:
        return self.invoke(member, *indices, value, kind=4)

    def release(self) -> None:
        if not self._released:
            try:
                self._workbook._call("release", references=[self._descriptor["reference"]])
            finally:
                self._released = True


def _float(tag: str, value: float) -> dict[str, Any]:
    payload: Any = float(value)
    if not math.isfinite(value):
        payload = {"bits": struct.pack(">f" if tag == "single" else ">d", value).hex()}
    return {"type": tag, "value": payload}


def _rectangular(value: Any) -> tuple[tuple[int, ...], list[Any]]:
    if not isinstance(value, (list, tuple)):
        return (), [value]
    if not value:
        return (0,), []
    children = [_rectangular(item) for item in value]
    if all(shape == children[0][0] for shape, _ in children):
        return (len(value),) + children[0][0], [v for _, items in children for v in items]
    return (len(value),), list(value)


def _encode(value: Any) -> dict[str, Any]:
    if isinstance(value, _Unit): return {"type": value.name}
    if value is None: return {"type": "null"}
    if isinstance(value, bool): return {"type": "boolean", "value": value}
    for cls, tag in [(Byte, "byte"), (Integer, "integer"), (Long, "long"), (LongLong, "long_long")]:
        if isinstance(value, cls):
            return {"type": tag, "value": str(value) if tag == "long_long" else int(value)}
    if isinstance(value, int):
        if -(1 << 31) <= value < 1 << 31: return {"type": "long", "value": value}
        if -(1 << 63) <= value < 1 << 63: return {"type": "long_long", "value": str(value)}
        raise OverflowError("VBA integers must fit a signed 64-bit value")
    if isinstance(value, Single): return _float("single", value)
    if isinstance(value, Date): return _float("date", value)
    if isinstance(value, float): return _float("double", value)
    if isinstance(value, Currency): return {"type": "currency", "value": format(value, "f")}
    if isinstance(value, decimal.Decimal):
        if not value.is_finite(): raise ValueError("VBA Decimal must be finite")
        parts = value.as_tuple()
        magnitude = int("".join(map(str, parts.digits)))
        exponent = int(parts.exponent)
        if exponent > 0: magnitude *= 10 ** exponent
        return {"type": "decimal", "value": {"negative": bool(parts.sign),
                "magnitude": str(magnitude), "scale": max(0, -exponent)}}
    if isinstance(value, str):
        if any(0xd800 <= ord(c) <= 0xdfff for c in value):
            raw = value.encode("utf-16-le", errors="surrogatepass")
            return {"type": "utf16", "value": list(struct.unpack("<" + "H" * (len(raw) // 2), raw))}
        return {"type": "string", "value": value}
    if isinstance(value, BStr): return {"type": "bstr", "value": list(value.data)}
    if isinstance(value, ErrorValue): return {"type": "error", "value": value.code}
    if isinstance(value, Object):
        if value._released: raise OxiCellsError("VBA object reference has been released")
        return {"type": "object", "value": value._descriptor}
    if isinstance(value, (list, tuple)):
        shape, values = _rectangular(value)
        value = Array(tuple(values), tuple((0, length) for length in shape))
    if isinstance(value, Array):
        return {"type": "array", "value": {"dimensions": [
            {"lower_bound": lower, "length": length} for lower, length in value.dimensions],
            "element_default": _encode(value.element_default), "resizable": value.resizable,
            "values": [_encode(v) for v in value.values]}}
    if isinstance(value, Record):
        return {"type": "record", "value": {"type_name": value.type_name,
                "fields": [[name, _encode(v)] for name, v in value.fields]}}
    raise TypeError(f"unsupported VBA argument type: {type(value).__name__}")


def _decode(descriptor: dict[str, Any], workbook: Workbook) -> Any:
    tag, value = descriptor["type"], descriptor.get("value")
    if tag == "null": return None
    if tag in ("empty", "missing", "nothing", "null_string"):
        return {"empty": EMPTY, "missing": MISSING, "nothing": NOTHING, "null_string": NULL_STRING}[tag]
    if tag in ("byte", "integer", "long", "long_long"):
        return {"byte": Byte, "integer": Integer, "long": Long, "long_long": LongLong}[tag](value)
    if tag in ("single", "double", "date"):
        if isinstance(value, dict):
            value = struct.unpack(">f" if tag == "single" else ">d", bytes.fromhex(value["bits"]))[0]
        return {"single": Single, "double": float, "date": Date}[tag](value)
    if tag == "currency": return Currency(value)
    if tag == "decimal":
        return decimal.Decimal((int(value["negative"]), tuple(map(int, value["magnitude"])), -value["scale"]))
    if tag == "utf16": return struct.pack("<" + "H" * len(value), *value).decode("utf-16-le", errors="surrogatepass")
    if tag == "bstr": return BStr(bytes(value))
    if tag == "error": return ErrorValue(value)
    if tag in ("boolean", "string"): return value
    if tag == "object": return Object(workbook, value)
    if tag == "array":
        return Array(tuple(_decode(v, workbook) for v in value["values"]),
                     tuple((d["lower_bound"], d["length"]) for d in value["dimensions"]),
                     _decode(value["element_default"], workbook), value["resizable"])
    if tag == "record": return Record(value["type_name"], tuple((name, _decode(v, workbook)) for name, v in value["fields"]))
    raise OxiCellsError(f"unknown native VBA result type: {tag}")


def binary_path() -> str:
    override = os.environ.get("OXICELLS_BINARY")
    if override: return override
    suffix = ".exe" if os.name == "nt" else ""
    installed = Path(sysconfig.get_path("scripts")) / ("oxicells" + suffix)
    if installed.is_file(): return str(installed)
    found = shutil.which("oxicells")
    if found: return found
    raise OxiCellsError("OxiCells engine not found; install the platform wheel or set OXICELLS_BINARY")


class Workbook:
    def __init__(self, path: os.PathLike[str] | str, *, binary: Optional[str] = None,
                 source_code_page: Optional[int] = None, timeout: Optional[float] = None) -> None:
        self.path = Path(path).resolve()
        self.timeout = timeout
        self.last_outcome: dict[str, Any] = {}
        self._lock = threading.RLock()
        self._answers: queue.Queue[Optional[str]] = queue.Queue()
        self._sequence = 0
        self._closed = False
        self._process = subprocess.Popen([binary or binary_path(), "serve"], stdin=subprocess.PIPE,
            stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, text=True, encoding="utf-8", bufsize=1,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        threading.Thread(target=self._read, daemon=True).start()
        try:
            info = self._rpc("open", path=str(self.path), sourceCodePage=source_code_page,
                             seed=secrets.randbits(64))
            if info.get("protocol") != 1: raise OxiCellsError("unsupported OxiCells engine protocol")
            self.sheets = tuple(info["sheets"])
            self.has_macros = bool(info["hasMacros"])
        except BaseException:
            self._stop()
            raise

    def _read(self) -> None:
        assert self._process.stdout is not None
        try:
            for line in self._process.stdout: self._answers.put(line)
        finally:
            self._answers.put(None)

    def _rpc(self, operation: str, **fields: Any) -> dict[str, Any]:
        with self._lock:
            if self._closed: raise OxiCellsError("workbook session is closed")
            self._sequence += 1
            local = (_datetime.datetime.now() - _datetime.datetime(1899, 12, 30)).total_seconds() / 86400
            request = {"id": self._sequence, "op": operation, "currentTime": local, **fields}
            data = json.dumps(request, ensure_ascii=True, allow_nan=False)
            try:
                assert self._process.stdin is not None
                self._process.stdin.write(data + "\n")
                self._process.stdin.flush()
                line = self._answers.get(timeout=self.timeout)
            except queue.Empty as error:
                self._stop()
                raise TimeoutError("VBA request timed out; workbook session was terminated") from error
            except (OSError, ValueError) as error:
                self._stop()
                raise OxiCellsError("OxiCells engine connection closed") from error
            if line is None:
                self._stop()
                raise OxiCellsError("OxiCells engine exited")
            try:
                answer = json.loads(line)
                if answer.get("id") != self._sequence: raise ValueError("response id does not match")
            except ValueError as error:
                self._stop()
                raise OxiCellsError("invalid OxiCells engine response") from error
            if not answer.get("ok"): raise OxiCellsError(answer.get("error", "engine request failed"))
            return answer["outcome"]

    def _call(self, operation: str, **fields: Any) -> Any:
        outcome = self._rpc(operation, **fields)
        self.last_outcome = outcome
        if "sheets" in outcome: self.sheets = tuple(outcome["sheets"])
        if outcome.get("error"): raise MacroError(outcome["error"])
        return _decode(outcome["result"], self) if "result" in outcome else None

    def run(self, procedure: str, *args: Any) -> Any:
        return self._call("run", procedure=procedure, args=[_encode(v) for v in args])

    def compile(self) -> dict[str, Any]:
        self._call("compile")
        return self.last_outcome["compilation"]

    def save(self, path: os.PathLike[str] | str) -> None:
        self._call("save", path=str(Path(path).resolve()))

    def _stop(self) -> None:
        self._closed = True
        try:
            if self._process.stdin: self._process.stdin.close()
            self._process.wait(timeout=5)
        except (OSError, subprocess.TimeoutExpired):
            self._process.kill()
            self._process.wait()
        if self._process.stdout: self._process.stdout.close()

    def close(self) -> None:
        with self._lock:
            if not self._closed:
                try: self._call("close")
                finally: self._stop()

    def __enter__(self) -> Workbook: return self

    def __exit__(self, kind: Any, error: Any, traceback: Any) -> None:
        try: self.close()
        except Exception:
            if kind is None: raise


def open(path: os.PathLike[str] | str, **options: Any) -> Workbook:
    """Open a workbook without executing macros; use ``run`` explicitly."""
    return Workbook(path, **options)
