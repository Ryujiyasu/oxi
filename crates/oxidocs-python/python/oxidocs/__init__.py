# SPDX-License-Identifier: MIT OR Apache-2.0
"""Convert .docx to PDF with the Oxi engine, without LibreOffice or Word.

    import oxidocs

    result = oxidocs.docx_to_pdf("report.docx", "report.pdf")
    print(result.pages)

    # Many documents: keep one engine process running.
    with oxidocs.Converter(threads=4) as converter:
        for name in names:
            converter.convert(f"{name}.docx", f"{name}.pdf")

With ``strict=True`` a document that holds content Oxi does not draw (a
chart type it does not support, SmartArt with no saved drawing, less than
90% of the document's own text drawn) raises :class:`UnsupportedContent`
instead of producing an incomplete PDF, so the caller can fall back to
another converter.

The package ships the ``oxidocs`` command-line binary (also on your PATH
after installing) and drives it through its ``serve`` mode: one JSON request
per line in, one JSON answer per line out.
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import sysconfig
import threading
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional, Union

__all__ = [
    "Converter",
    "ConversionError",
    "Result",
    "UnsupportedContent",
    "binary_path",
    "docx_to_pdf",
]

PathLike = Union[str, "os.PathLike[str]"]

# Exit statuses of `oxidocs docx-to-pdf`, which serve answers carry as "exit".
EXIT_UNSUPPORTED = 3


@dataclass(frozen=True)
class Result:
    """A finished conversion."""

    output: Path
    pages: int
    bytes: int
    ms: float


class ConversionError(Exception):
    """Oxi could not convert the document (unreadable file, engine failure)."""

    def __init__(self, message: str, exit_status: Optional[int] = None) -> None:
        super().__init__(message)
        self.exit_status = exit_status


class UnsupportedContent(ConversionError):
    """``strict=True`` and the document holds content Oxi does not draw.

    ``kinds`` lists what was found, e.g. ``["chart"]`` or
    ``["undrawn-text (84% drawn)"]``; ``details`` keeps the engine's full
    report (kind, package part, count).
    """

    def __init__(self, details: list[dict[str, Any]]) -> None:
        self.details = details
        self.kinds = list(dict.fromkeys(str(d.get("kind", "unknown")) for d in details)) or ["unknown"]
        super().__init__("content Oxi does not draw: " + ", ".join(self.kinds), EXIT_UNSUPPORTED)


def binary_path() -> str:
    """The oxidocs binary this package runs.

    ``OXIDOCS_BINARY`` overrides it; otherwise the one installed with the
    package (the wheel puts it in the environment's scripts directory), then
    one on PATH.
    """
    override = os.environ.get("OXIDOCS_BINARY")
    if override:
        return override
    exe = "oxidocs.exe" if sys.platform == "win32" else "oxidocs"
    for scheme_path in (sysconfig.get_path("scripts"), sysconfig.get_path("scripts", f"{os.name}_user")):
        if scheme_path:
            candidate = Path(scheme_path) / exe
            if candidate.is_file():
                return str(candidate)
    found = shutil.which("oxidocs")
    if found:
        return found
    raise ConversionError(
        "oxidocs binary not found; reinstall the package, or set OXIDOCS_BINARY to its path"
    )


def _answer_to_result(answer: dict[str, Any], output: Path) -> Result:
    if answer.get("ok"):
        return Result(output=output, pages=int(answer["pages"]), bytes=int(answer["bytes"]), ms=float(answer.get("ms", 0.0)))
    if answer.get("exit") == EXIT_UNSUPPORTED:
        raise UnsupportedContent(list(answer.get("unsupported") or []))
    raise ConversionError(str(answer.get("error", "conversion failed")), answer.get("exit"))


class Converter:
    """One running ``oxidocs serve`` process, reused for many documents.

    Thread-safe: concurrent :meth:`convert` calls are sent to the engine as
    they come and each waits for its own answer; the engine converts up to
    ``threads`` of them in parallel. If the process dies it is restarted for
    the next call, and the calls it was serving fail with
    :class:`ConversionError`.
    """

    def __init__(self, threads: int = 2, binary: Optional[str] = None) -> None:
        if threads < 1:
            raise ValueError("threads must be at least 1")
        self._threads = threads
        self._binary = binary
        self._lock = threading.Lock()
        self._proc: Optional[subprocess.Popen[str]] = None
        self._pending: dict[str, _Waiter] = {}
        self._next_id = 0

    # -- lifecycle ---------------------------------------------------------

    def __enter__(self) -> "Converter":
        return self

    def __exit__(self, *exc: object) -> None:
        self.close()

    def close(self) -> None:
        """Stop the engine process (it finishes the requests it holds)."""
        with self._lock:
            proc, self._proc = self._proc, None
        if proc is not None:
            try:
                if proc.stdin:
                    proc.stdin.close()
                proc.wait(timeout=60)
            except (OSError, subprocess.TimeoutExpired):
                proc.kill()

    def _start(self) -> subprocess.Popen[str]:
        proc = subprocess.Popen(
            [self._binary or binary_path(), "serve", "--threads", str(self._threads)],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL,
            text=True,
            encoding="utf-8",
            bufsize=1,
        )
        threading.Thread(target=self._read, args=(proc,), daemon=True).start()
        return proc

    def _read(self, proc: subprocess.Popen[str]) -> None:
        assert proc.stdout is not None
        for line in proc.stdout:
            try:
                answer = json.loads(line)
            except ValueError:
                continue
            waiter = self._pending.pop(str(answer.get("id")), None)
            if waiter is not None:
                waiter.answer = answer
                waiter.done.set()
        # The process is gone: fail everything it still held.
        with self._lock:
            if self._proc is proc:
                self._proc = None
            gone = [w for w in self._pending.values() if w.proc is proc]
            for w in gone:
                self._pending.pop(w.id, None)
        for w in gone:
            w.answer = {"ok": False, "exit": None, "error": "oxidocs process exited"}
            w.done.set()

    # -- conversion --------------------------------------------------------

    def convert(self, docx: PathLike, pdf: PathLike, *, strict: bool = False, timeout: Optional[float] = None) -> Result:
        """Convert ``docx`` to ``pdf``; see :func:`docx_to_pdf`."""
        source, target = Path(docx).resolve(), Path(pdf).resolve()
        if not source.is_file():
            raise ConversionError(f"no such file: {source}")
        with self._lock:
            if self._proc is None or self._proc.poll() is not None:
                self._proc = self._start()
            self._next_id += 1
            waiter = _Waiter(id=str(self._next_id), proc=self._proc)
            self._pending[waiter.id] = waiter
            request = {"id": waiter.id, "input": str(source), "output": str(target), "strict": strict}
            try:
                assert self._proc.stdin is not None
                self._proc.stdin.write(json.dumps(request) + "\n")
                self._proc.stdin.flush()
            except OSError as e:
                self._pending.pop(waiter.id, None)
                raise ConversionError(f"oxidocs process unavailable: {e}") from e
        if not waiter.done.wait(timeout):
            self._pending.pop(waiter.id, None)
            # A stuck document would hold an engine thread: restart the process.
            with self._lock:
                proc, self._proc = self._proc, None
            if proc is not None:
                proc.kill()
            raise ConversionError(f"timed out after {timeout} s: {source}")
        return _answer_to_result(waiter.answer, target)


@dataclass
class _Waiter:
    id: str
    proc: Optional[subprocess.Popen[str]]
    done: threading.Event = field(default_factory=threading.Event)
    answer: dict[str, Any] = field(default_factory=dict)


def docx_to_pdf(
    docx: PathLike,
    pdf: PathLike,
    *,
    strict: bool = False,
    timeout: Optional[float] = None,
) -> Result:
    """Convert one .docx file to PDF.

    Returns the output path, page count, size and time taken. Raises
    :class:`UnsupportedContent` (``strict=True`` only) when the document
    holds content Oxi does not draw, and :class:`ConversionError` when it
    cannot be converted at all. ``timeout`` is in seconds.

    For many documents, a :class:`Converter` saves starting the engine for
    each one.
    """
    with Converter(threads=1) as converter:
        return converter.convert(docx, pdf, strict=strict, timeout=timeout)
