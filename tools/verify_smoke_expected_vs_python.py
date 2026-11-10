#!/usr/bin/env python3
"""
Cross-check SmokeTest_MathParser.bas numeric *expected* strings against a Python
reference (math + statistics + explicit Smart-Math-like rules).

Expressions use the Smart Math surface syntax; we translate a large subset to
Python. Cases we cannot translate safely are reported as SKIP (not mismatches).

When an expression *looks like* a complex or time-duration case (heuristic), we
switch to a dedicated reference path:
  * **Complex:** map Smart-Math `i` suffix / bare `i` to Python ``j``, evaluate with
    ``complex`` arithmetic, then compare after normalizing ``j`` <-> ``i`` spelling.
    Tuple + scalar / tuple * scalar / element-wise tuple/tuple uses the same
    broadcast lowering as the ``RunComplexNumberSupportOptionTests`` block in
    ``SmokeTest_MathParser.bas``. Table rows are discovered from:
    ``StemCases``/``StemOk`` + ``StemExpect``; ``StemExpr`` + ``StemExpect`` (with
    ``&`` string concat via ``dim … as String`` constants); ``dim var`` +
    ``Parser_TryEvaluateEx(var,…) orelse rt <> varExpect``; inline literal probes;
    and ``*ErrExpr`` error tables. ``hex``/``bin``/``oct``/``uhex``/``uoct``/``ubin``
    are evaluated for reference strings (uppercase ``0x``). User functions,
    ``ratio()``, unsigned display, array indexing, and wide integer ``mod``
    are evaluated. Rows that only record an error string stay skipped.
    A bare trailing formatter (``hex``/``oct``/``bin``/
    ``deg``/``rad`` and the unsigned names, optional empty ``()``), including a
    final ``;``, is rewritten to ``name(ans)``.
  * **Real (recent exact-int math):** ``sqrt``/``sqr`` use Smart-Math-style exact
    integer rules; ``tests(1150)``..``1156`` and ``raw-api/sqrt-non-perfect-square``
    are included. ``hex(sqr(9))`` etc. use the same display builtins as smoke.
  * **Time:** colon literals (MM:SS / HH:MM:SS / DD:HH:MM:SS), named duration
    constants (``second`` … ``day``), compact suffix forms (``1d2h3m4s5ms``),
    converters (``milliseconds``, ``seconds``, …), and ``sum`` of durations —
    enough to mirror the smoke *expected* strings for the time block.

Float agreement uses the same tolerance as ``ScalarCloseEnough`` in
``SmokeTest_MathParser.bas``:
  tol = max(256 * eps * max(1, |a|, |b|), 1e-14)

Smart-Math rules plain Python does not model:
  * ``min`` / ``max`` skip ``nan`` once any non-NaN value is present (NaN is absence).
    ``clamp`` uses the same absence rules for a NaN value or a NaN bound.
  * Arrays are flattened as they are built, so a result never contains a nested array.
  * A bare trailing formatter is ``name(ans)``; a final ``;`` still yields that result.

Sources parsed (in addition to ``tests(N)`` rows):
  * ``tools/neg_band_cases_generated.bas`` (neg/pos real + complex magnitude bands)
  * ``tools/complex_coverage_generated.bas`` (``cxCov*`` tables)
  * ``tools/naninf_mirror_generated.bas`` (``naninf*`` tables)
  * ``tools/float_magnitude_cases_generated.bas`` (``floatMag*`` tables)
  * ``tools/smoke_parity_append_generated.bas`` (``smokeAppend*``; merged into main table)
  * ``RunOperatorPrecedenceDocTests`` (``precOk`` / ``precExpect``)
  * ``RunTrigAngleReductionTests`` (magnitude ``|result|`` thresholds + ``exactOne*``)
  * Paired ``*Ok``/``*Expect`` and ``*Expr``/``*Expect`` rows anywhere in smoke source
  * ``*Err`` / ``errExprs`` error-probe tables in smoke subs (time, complex, arity, …)
  * Inline ``Parser_TryEvaluateEx(...) orelse rt <> "..."`` numeric probes in smoke subs
  * All ``Run*Tests`` subs in ``SmokeTest_MathParser.bas`` (paired ``*Ok/*Expect``,
    ``*Cases/*Expect``, ``*Exprs/*Need``, ``*Expr/*Expect`` with ``&`` concat,
    error tables, inline ``rt <>`` probes)
  * ``RunComplexNumberSupportOptionTests`` (large complex-opt block)
  * ``RunRawResultApiTests`` (display-string probes where comparable)

Use ``--audit-coverage`` to print discovered vs parsed row counts per source and per
``Run*Tests`` sub (before the Python reference pass). Use ``--audit-only`` to exit
after the audit (fast). C++ ``Tests_MathParser.cpp`` is not parsed (Basic smoke corpus).
"""

from __future__ import annotations

import ast
import cmath
import math
import os
import re
import statistics
import sys
import threading
import time
import warnings
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal
from fractions import Fraction
from typing import Any, Optional

EPS = 2.2204460492503131e-16

# --- Duration (time) reference model (milliseconds, signed) -----------------

_MS = 1
_SECOND_MS = 1000
_MINUTE_MS = 60 * _SECOND_MS
_HOUR_MS = 60 * _MINUTE_MS
_DAY_MS = 24 * _HOUR_MS

_UNIT_ORDER = {"d": 0, "h": 1, "m": 2, "s": 3, "ms": 4}


@dataclass(frozen=True)
class Duration:
    """Signed duration stored in milliseconds (may be fractional)."""

    ms: float

    def __neg__(self) -> "Duration":
        return Duration(-self.ms)

    def _other_ms(self, other: object) -> Optional[float]:
        if isinstance(other, Duration):
            return other.ms
        if isinstance(other, bool) or not isinstance(other, (int, float)):
            return None
        return float(other) * _SECOND_MS

    def __add__(self, other: object) -> Any:
        extra = self._other_ms(other)
        if extra is None:
            return NotImplemented
        return Duration(self.ms + extra)

    def __radd__(self, other: object) -> Any:
        return self.__add__(other)

    def __sub__(self, other: object) -> Any:
        if isinstance(other, Duration):
            return Duration(self.ms - other.ms)
        return NotImplemented

    def __rsub__(self, other: object) -> Any:
        if isinstance(other, Duration):
            return Duration(other.ms - self.ms)
        return NotImplemented

    def __mul__(self, other: object) -> Any:
        if isinstance(other, bool) or not isinstance(other, (int, float)):
            return NotImplemented
        return Duration(self.ms * float(other))

    def __rmul__(self, other: object) -> Any:
        return self.__mul__(other)

    def __lt__(self, other: object) -> Any:
        extra = self._other_ms(other)
        if extra is None:
            return NotImplemented
        return self.ms < extra

    def __le__(self, other: object) -> Any:
        extra = self._other_ms(other)
        if extra is None:
            return NotImplemented
        return self.ms <= extra

    def __gt__(self, other: object) -> Any:
        extra = self._other_ms(other)
        if extra is None:
            return NotImplemented
        return self.ms > extra

    def __ge__(self, other: object) -> Any:
        extra = self._other_ms(other)
        if extra is None:
            return NotImplemented
        return self.ms >= extra

    def __truediv__(self, other: object) -> Any:
        if isinstance(other, bool) or not isinstance(other, (int, float)):
            return NotImplemented
        if other == 0:
            return Duration(math.copysign(math.inf, self.ms) if self.ms != 0 else math.nan)
        return Duration(self.ms / float(other))


@dataclass(frozen=True)
class ExactComplex:
    """Cartesian complex with exact integer components (Smart-Math int64 path)."""

    real: int
    imag: int

    def conjugate(self) -> "ExactComplex":
        return ExactComplex(self.real, -self.imag)

    def __add__(self, other: object) -> Any:
        pair = _exact_complex_pair(other)
        if pair is None:
            return NotImplemented
        br, bi, exact = pair
        return _exact_complex_make(self.real + br, self.imag + bi, exact)

    def __radd__(self, other: object) -> Any:
        return self.__add__(other)

    def __sub__(self, other: object) -> Any:
        pair = _exact_complex_pair(other)
        if pair is None:
            return NotImplemented
        br, bi, exact = pair
        return _exact_complex_make(self.real - br, self.imag - bi, exact)

    def __rsub__(self, other: object) -> Any:
        pair = _exact_complex_pair(other)
        if pair is None:
            return NotImplemented
        br, bi, exact = pair
        return _exact_complex_make(br - self.real, bi - self.imag, exact)

    def __mul__(self, other: object) -> Any:
        pair = _exact_complex_pair(other)
        if pair is None:
            return NotImplemented
        br, bi, exact = pair
        rr = self.real * br - self.imag * bi
        ii = self.real * bi + self.imag * br
        return _exact_complex_make(rr, ii, exact)

    def __rmul__(self, other: object) -> Any:
        return self.__mul__(other)

    def __truediv__(self, other: object) -> Any:
        pair = _exact_complex_pair(other)
        if pair is None:
            return NotImplemented
        br, bi, exact = pair
        den = br * br + bi * bi
        if den == 0:
            raise ZeroDivisionError
        nr = self.real * br + self.imag * bi
        ni = self.imag * br - self.real * bi
        if exact and isinstance(den, int) and nr % den == 0 and ni % den == 0:
            return ExactComplex(nr // den, ni // den)
        return complex(nr / den, ni / den)

    def __rtruediv__(self, other: object) -> Any:
        pair = _exact_complex_pair(other)
        if pair is None:
            return NotImplemented
        br, bi, exact = pair
        den = self.real * self.real + self.imag * self.imag
        if den == 0:
            raise ZeroDivisionError
        nr = br * self.real + bi * self.imag
        ni = bi * self.real - br * self.imag
        if exact and nr % den == 0 and ni % den == 0:
            return ExactComplex(nr // den, ni // den)
        return complex(nr / den, ni / den)

    def __neg__(self) -> "ExactComplex":
        return ExactComplex(-self.real, -self.imag)

    def __abs__(self) -> float:
        return math.hypot(self.real, self.imag)


def _exact_complex_pair(other: object) -> Optional[tuple[Any, Any, bool]]:
    """Return (real, imag, both_components_are_ints) or None."""
    if isinstance(other, ExactComplex):
        return other.real, other.imag, True
    if isinstance(other, complex):
        return other.real, other.imag, False
    if isinstance(other, bool):
        return int(other), 0, True
    if isinstance(other, int):
        return other, 0, True
    if isinstance(other, float):
        return other, 0.0, False
    return None


def _exact_complex_make(real: Any, imag: Any, exact: bool) -> Any:
    if (
        exact
        and isinstance(real, int)
        and isinstance(imag, int)
        and not isinstance(real, bool)
        and not isinstance(imag, bool)
    ):
        return ExactComplex(real, imag)
    return complex(float(real), float(imag))


def _parse_colon_groups(groups: list[str]) -> Duration:
    """Parse 2/3/4 colon-separated groups into total ms (Smart-Math segment rules)."""
    parts = [g.strip() for g in groups if g.strip() != ""]
    n = len(parts)
    if n < 2 or n > 4:
        raise ValueError("colon time segment count")
    last = parts[-1]
    frac = 0.0
    if "." in last:
        base, frac_s = last.split(".", 1)
        parts[-1] = base
        frac = float("0." + frac_s) * _SECOND_MS if frac_s else 0.0
    ints = [int(p) for p in parts]
    if n == 2:
        mm, ss = ints
        return Duration(mm * _MINUTE_MS + ss * _SECOND_MS + frac)
    if n == 3:
        hh, mm, ss = ints
        return Duration(hh * _HOUR_MS + mm * _MINUTE_MS + ss * _SECOND_MS + frac)
    dd, hh, mm, ss = ints
    return Duration(dd * _DAY_MS + hh * _HOUR_MS + mm * _MINUTE_MS + ss * _SECOND_MS + frac)


def dur_parse_colon_literal(text: str) -> Duration:
    s = text.strip()
    return _parse_colon_groups(s.split(":"))


def dur_parse_compact(s: str) -> Duration:
    """
    Parse compact suffix literal: ``1d2h3m4s5ms``, ``5000ms``, ``-1s-1m`` (spaces ignored).
    Units must appear in non-decreasing order d,h,m,s,ms.
    Per-token leading ``-`` is allowed on each field (``-1s-1m``).
    For unary minus on a whole literal written like ``-1m 1s``, use ``(-_dur_lit_compact('1m1s'))``.
    """
    s0 = re.sub(r"\s+", "", s.strip())
    tokens = re.findall(r"-?\d+(?:ms|[dhms])", s0)
    if not tokens:
        raise ValueError("empty compact time")
    total = 0.0
    for tok in tokens:
        m = re.fullmatch(r"(-?)(\d+)(ms|[dhms])", tok)
        if not m:
            raise ValueError("bad compact token")
        sign = -1.0 if m.group(1) == "-" else 1.0
        val = int(m.group(2)) * sign
        unit = m.group(3)
        # Each token keeps its own sign, so ``-1s-1m`` is -1 second plus -1 minute.
        if unit == "d":
            total += val * _DAY_MS
        elif unit == "h":
            total += val * _HOUR_MS
        elif unit == "m":
            total += val * _MINUTE_MS
        elif unit == "s":
            total += val * _SECOND_MS
        elif unit == "ms":
            total += val * _MS
    return Duration(total)


def _replace_compact_time_literals(expr: str) -> str:
    """Replace compact duration runs with ``_dur_lit_compact('...')`` or negated form."""

    def scan_replace(s: str) -> str:
        out: list[str] = []
        i = 0
        n = len(s)
        while i < n:
            m = re.match(r"-?\d+(?:ms|[dhms])", s[i:])
            if not m:
                out.append(s[i])
                i += 1
                continue
            j = i + m.end()
            parts = [m.group(0)]
            sp_mid = j
            while j < n and s[j].isspace():
                j += 1
            had_space_between = j > sp_mid
            while j < n:
                m2 = re.match(r"-?\d+(?:ms|[dhms])", s[j:])
                if not m2:
                    break
                parts.append(m2.group(0))
                j += m2.end()
                while j < n and s[j].isspace():
                    j += 1
            blob = "".join(parts)
            if re.fullmatch(r"-?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?", blob):
                out.append(blob)
                i = j
                continue
            if ":" in blob:
                out.append(blob)
                i = j
                continue
            if (
                had_space_between
                and len(parts) == 2
                and parts[1]
                and parts[0].startswith("-")
                and not parts[1].startswith("-")
                and parts[0] == "-" + parts[1][0] + "m"
            ):
                inner = parts[1][0] + "m" + parts[1]
                try:
                    dur_parse_compact(inner)
                except Exception:
                    out.append(blob)
                    i = j
                    continue
                out.append(f"(-_dur_lit_compact({inner!r}))")
                i = j
                continue
            try:
                dur_parse_compact(blob)
            except Exception:
                out.append(blob)
                i = j
                continue
            out.append(f"_dur_lit_compact({blob!r})")
            i = j
        return "".join(out)

    return scan_replace(expr)


def _replace_colon_time_literals(expr: str) -> str:
    """Replace ``H:MM(:SS)?`` colon durations with ``dur_parse_colon_literal('...')``."""

    def find_match(s: str, pos: int) -> Optional[re.Match]:
        if pos >= len(s) or not s[pos].isdigit():
            return None
        for nseg in (4, 3, 2):
            if nseg == 4:
                rgx = re.compile(
                    r"^(\d+):(\d{2}):(\d{2}):(\d{2})(\.\d+)?(?![:\d])"
                )
            elif nseg == 3:
                rgx = re.compile(r"^(\d+):(\d{2}):(\d{2})(\.\d+)?(?!\s*:)")
            else:
                rgx = re.compile(r"^(\d+):(\d{2})(\.\d+)?(?!\s*:)")
            m = rgx.match(s[pos:])
            if m:
                return m
        return None

    out: list[str] = []
    i = 0
    n = len(expr)
    depth = 0
    while i < n:
        ch = expr[i]
        if ch == "[":
            depth += 1
        elif ch == "]" and depth:
            depth -= 1
        if depth == 0:
            m = find_match(expr, i)
            if m:
                lit = m.group(0)
                out.append(f"dur_parse_colon_literal({lit!r})")
                i += len(lit)
                continue
        out.append(ch)
        i += 1
    return "".join(out)


def _dur_lit_compact(s: str) -> Duration:
    return dur_parse_compact(s)


def duration_format_ms(ms: float) -> str:
    """
    Format ``ms`` into Smart-Math-like canonical strings used in smoke tests.
    Heuristic but tuned to the SmokeTest_MathParser.bas time block.
    """
    if ms != ms:
        return "nan"
    sign = "-" if ms < 0 else ""
    ms_abs = abs(ms)
    total_sec = ms_abs / _SECOND_MS

    if total_sec + 1e-9 < _DAY_MS / _SECOND_MS:
        if total_sec < 3600 - 1e-9:
            mm = int(math.floor(total_sec / 60.0 + 1e-12))
            ss = total_sec - mm * 60.0
            if abs(ss - round(ss)) < 1e-6:
                return f"{sign}{mm:02d}:{int(round(ss)):02d}"
            return f"{sign}{mm:02d}:{ss:06.3f}".rstrip("0").rstrip(".")
        hh = int(math.floor(total_sec / 3600.0 + 1e-12))
        rem = total_sec - hh * 3600.0
        mm = int(math.floor(rem / 60.0 + 1e-12))
        ss = rem - mm * 60.0
        if abs(ss) < 1e-6:
            return f"{sign}{hh:02d}:{mm:02d}"
        if abs(ss - round(ss)) < 1e-6:
            return f"{sign}{hh:02d}:{mm:02d}:{int(round(ss)):02d}"
        return f"{sign}{hh:02d}:{mm:02d}:{ss:06.3f}".rstrip("0").rstrip(".")
    dd = int(math.floor(ms_abs / _DAY_MS + 1e-12))
    rem_ms = ms_abs - dd * _DAY_MS
    hh = int(rem_ms // _HOUR_MS)
    rem_ms -= hh * _HOUR_MS
    mm = int(rem_ms // _MINUTE_MS)
    rem_ms -= mm * _MINUTE_MS
    sec = rem_ms / _SECOND_MS
    if abs(sec - round(sec)) < 1e-6:
        return f"{sign}{dd}:{hh:02d}:{mm:02d}:{int(round(sec)):02d}"
    return f"{sign}{dd}:{hh:02d}:{mm:02d}:{sec:06.3f}".rstrip("0").rstrip(".")


def milliseconds_sm(x: object) -> Any:
    if isinstance(x, Duration):
        return int(round(x.ms))
    if isinstance(x, tuple):
        return tuple(milliseconds_sm(t) for t in x)
    raise TypeError("milliseconds expects duration or tuple of durations")


def seconds_sm(x: object) -> Any:
    if isinstance(x, Duration):
        return x.ms / _SECOND_MS
    if isinstance(x, tuple):
        return tuple(seconds_sm(t) for t in x)
    raise TypeError("seconds expects duration or tuple of durations")


def minutes_sm(x: object) -> Any:
    if isinstance(x, Duration):
        return x.ms / _MINUTE_MS
    if isinstance(x, tuple):
        return tuple(minutes_sm(t) for t in x)
    raise TypeError("minutes expects duration or tuple of durations")


def hours_sm(x: object) -> Any:
    if isinstance(x, Duration):
        return x.ms / _HOUR_MS
    if isinstance(x, tuple):
        return tuple(hours_sm(t) for t in x)
    raise TypeError("hours expects duration or tuple of durations")


def days_sm(x: object) -> Any:
    if isinstance(x, Duration):
        return x.ms / _DAY_MS
    if isinstance(x, tuple):
        return tuple(days_sm(t) for t in x)
    raise TypeError("days expects duration or tuple of durations")


def sum_dur_sm(*args: object) -> Duration:
    if len(args) == 1 and isinstance(args[0], tuple):
        args = args[0]  # type: ignore[assignment]
    acc = Duration(0.0)
    for a in args:
        if not isinstance(a, Duration):
            raise TypeError("sum_dur expects durations")
        acc = Duration(acc.ms + a.ms)
    return acc


# --- Complex helpers ---------------------------------------------------------


_RE_TIGHT_NUM_IMAG = re.compile(
    r"(?<![A-Za-z0-9_.])"
    r"("
    r"0[xX][0-9a-fA-F]+"
    r"|0[bB][01]+"
    r"|0[oO][0-7]+"
    r"|\d+\.?\d*|\.\d+"
    r"(?:[eE][-+]?\d+)?"
    r")"
    r"[iI](?![A-Za-z0-9_])",
    re.I,
)

# Back-compat alias for heuristics that only need decimal/scientific suffix probes.
_RE_SMART_IMAG_SUFFIX = re.compile(
    r"(?<!\*)(?<![A-Za-z0-9_.])(\d+\.?\d*(?:[eE][-+]?\d+)?)i(?![A-Za-z0-9_])",
    re.I,
)

_IMAG_UNIT_PY = "1j"


def _is_tight_imag_at(s: str, pos: int) -> bool:
    if pos < 0 or pos >= len(s) or s[pos] not in "iI":
        return False
    nxt = s[pos + 1] if pos + 1 < len(s) else ""
    return not (nxt.isalnum() or nxt == "_")


def _find_matching_open_paren(s: str, close_index: int) -> Optional[int]:
    depth = 1
    j = close_index - 1
    while j >= 0:
        if s[j] == ")":
            depth += 1
        elif s[j] == "(":
            depth -= 1
            if depth == 0:
                return j
        j -= 1
    return None


def _ident_start_before(s: str, pos: int) -> Optional[int]:
    """Return identifier start when an identifier ends immediately before ``pos``."""
    q = pos - 1
    while q >= 0 and s[q].isspace():
        q -= 1
    if q < 0 or not (s[q].isalnum() or s[q] == "_"):
        return None
    end = q + 1
    while q >= 0 and (s[q].isalnum() or s[q] == "_"):
        q -= 1
    start = q + 1
    if start > 0 and (s[start - 1].isalnum() or s[start - 1] in "._"):
        return None
    return start


def _tight_imag_replacement(s: str, i_pos: int) -> Optional[tuple[int, int, str]]:
    """
    Smart-Math tight-``i`` lowering (rightmost ``i_pos`` only):

    * ``Xi`` -> ``(X*j)`` for numeric literals ``X`` (decimal/scientific/``0x``/``0b``/``0o``)
    * ``(expr)i`` -> ``((expr)*j)``
    * ``func(expr)i`` -> ``(func(expr)*j)`` (``func`` is a normal identifier call)
    """
    if not _is_tight_imag_at(s, i_pos):
        return None

    k = i_pos - 1
    while k >= 0 and s[k].isspace():
        k -= 1
    if k < 0:
        return None

    if s[k] == ")":
        open_paren = _find_matching_open_paren(s, k)
        if open_paren is None:
            return None
        call_start = _ident_start_before(s, open_paren)
        if call_start is not None:
            call_text = s[call_start:i_pos]
            return (call_start, i_pos + 1, f"({call_text}*{_IMAG_UNIT_PY})")
        expr = s[open_paren + 1 : k]
        return (open_paren, i_pos + 1, f"(({expr})*{_IMAG_UNIT_PY})")

    tail = s[: i_pos + 1]
    m = None
    for m in _RE_TIGHT_NUM_IMAG.finditer(tail):
        pass
    if m is None or m.end() != i_pos + 1:
        return None
    x = m.group(1)
    return (m.start(), i_pos + 1, f"({x}*{_IMAG_UNIT_PY})")


def _apply_tight_imag_suffix_rules(expr: str) -> str:
    """Lower every Smart-Math tight ``i`` suffix using the three primary rules."""
    s = expr
    while True:
        i_pos = -1
        for p in range(len(s) - 1, -1, -1):
            if _is_tight_imag_at(s, p):
                i_pos = p
                break
        if i_pos < 0:
            break
        repl = _tight_imag_replacement(s, i_pos)
        if repl is None:
            break
        start, end, text = repl
        s = s[:start] + text + s[end:]
    return s


_RE_EXACT_INT_COMPLEX_ADD = re.compile(
    r"(?<![\d./])(-?\d+)\s*\+\s*\(\s*(-?\d+)\s*\*\s*1j\s*\)"
)
_RE_EXACT_INT_COMPLEX_SUB = re.compile(
    r"(?<![\d./])(-?\d+)\s*-\s*\(\s*(-?\d+)\s*\*\s*1j\s*\)"
)


def exact_int_complex(re: Any, im: Any = 0) -> ExactComplex:
    """Build exact integer complex (avoids IEEE float loss on int64-sized literals)."""
    return ExactComplex(int(re), int(im))


def _lower_exact_int_complex_literals(expr: str) -> str:
    """
    ``922...+922...*1j`` cannot be evaluated as native Python ``complex`` without
    rounding int64-sized parts; lower to ``exact_int_complex(re, im)`` instead.
    """

    def repl_sub(m: re.Match) -> str:
        return f"exact_int_complex({m.group(1)}, -{m.group(2)})"

    def repl_add(m: re.Match) -> str:
        return f"exact_int_complex({m.group(1)}, {m.group(2)})"

    s = _RE_EXACT_INT_COMPLEX_SUB.sub(repl_sub, expr)
    return _RE_EXACT_INT_COMPLEX_ADD.sub(repl_add, s)


def _preprocess_complex_surface(expr: str) -> str:
    """
    Turn Smart-Math complex surface syntax into Python ``complex`` literals.

    Tight ``i`` suffixes (no space, not part of a longer identifier):

    * ``Xi`` -> ``(X*1j)`` for numeric ``X``
    * ``(expr)i`` -> ``((expr)*1j)``
    * ``func(expr)i`` -> ``(func(expr)*1j)``

    Remaining bare ``i`` tokens (e.g. ``2*i``, ``1/2*i``) become ``(1j)``.
    An assignment ``i=...`` keeps that name as a variable.
    """
    s = _apply_tight_imag_suffix_rules(expr)
    if not _assigns_name_i(s):
        s = re.sub(r"(?<![A-Za-z0-9_])i(?![A-Za-z0-9_])", "(1j)", s)
    return s


def _assigns_name_i(expr: str) -> bool:
    """``i=2`` is a variable. ``i==2`` and ``2i`` are not assignments to ``i``."""
    return re.search(r"(?<![A-Za-z0-9_])i\s*=(?!=)", expr) is not None


def _mask_bracket_interiors(expr: str) -> str:
    """Blank the inside of ``[]`` so a slice colon is not read as a time."""
    out: list[str] = []
    depth = 0
    for ch in expr:
        if ch == "[":
            depth += 1
            out.append(ch)
        elif ch == "]" and depth:
            depth -= 1
            out.append(ch)
        elif depth:
            out.append(" ")
        else:
            out.append(ch)
    return "".join(out)


def _looks_like_complex_expr(expr: str) -> bool:
    e = _strip_line_comment(expr)
    if _RE_TIGHT_NUM_IMAG.search(e) or _RE_SMART_IMAG_SUFFIX.search(e):
        return True
    if re.search(r"\)[iI](?![A-Za-z0-9_])", e):
        return True
    if not _assigns_name_i(e) and re.search(r"(?<![A-Za-z0-9_])i(?![A-Za-z0-9_])", e):
        return True
    return False


def _looks_like_complex_expected(expected: str) -> bool:
    s = expected.strip()
    if "i" not in s and "I" not in s:
        return False
    if re.search(r"(?<![A-Za-z0-9_])(\d+\.?\d*|\.\d+)i\b", s, flags=re.I):
        return True
    if re.search(r"[+-]\s*i\b", s, flags=re.I):
        return True
    if re.search(r"\(\s*i\b", s, flags=re.I):
        return True
    return False


def _looks_like_time(expr: str, expected: str) -> bool:
    ex = _strip_line_comment(expr)
    visible = _mask_bracket_interiors(ex)
    if re.search(
        r"\b(second|minute|hour|day|millisecond|milliseconds|seconds|minutes|hours|days)\b",
        ex,
        flags=re.I,
    ):
        return True
    if re.search(r"\b\d+\s*:\s*\d{2}\b", visible):
        return True
    if re.search(r"(?:^|[^\w.])(-?\s*(?:\d+)(?:ms|[dhms])(?:\s*(?:\d+)(?:ms|[dhms]))+)(?:[^\w.]|$)", visible):
        if not re.search(r"\d+e[-+]?\d", ex, flags=re.I):
            return True
    if re.search(r"\b\d+\s*:\s*\d{2}\b", expected):
        return True
    return False


def infer_ref_mode(expr: str, expected: str) -> str:
    if _looks_like_time(expr, expected):
        return "time"
    if _looks_like_complex_expr(expr) or _looks_like_complex_expected(expected):
        return "complex"
    return "real"


def _flat_ast_tuple(elts: list[ast.AST]) -> ast.Tuple:
    """Splice nested tuple nodes. Smart-Math array values are always flat."""
    flat: list[ast.AST] = []
    for el in elts:
        if isinstance(el, ast.Tuple):
            flat.extend(_flat_ast_tuple(el.elts).elts)
        else:
            flat.append(el)
    return ast.Tuple(elts=flat, ctx=ast.Load())


def _fold_tuple_broadcast_ast(node: ast.AST) -> ast.AST:
    """
    Lower Smart-Math element-wise tuple algebra to a flat tuple of scalars
    (``(a,b)+c -> (a+c, b+c)``, same for ``- * /`` when one side is a tuple).
    Nested array literals are spliced into that flat tuple before the broadcast.
    """

    def fold(n: ast.AST) -> ast.AST:
        if isinstance(n, ast.BinOp):
            left, right = fold(n.left), fold(n.right)
            op = n.op
            if isinstance(left, ast.Tuple) and isinstance(right, ast.Tuple):
                if len(left.elts) != len(right.elts):
                    return ast.BinOp(left=left, op=op, right=right)
                return _flat_ast_tuple(
                    [
                        ast.BinOp(left=le, op=op, right=re)
                        for le, re in zip(left.elts, right.elts)
                    ]
                )
            if isinstance(left, ast.Tuple) and not isinstance(right, ast.Tuple):
                return _flat_ast_tuple(
                    [ast.BinOp(left=le, op=op, right=right) for le in left.elts]
                )
            if isinstance(right, ast.Tuple) and not isinstance(left, ast.Tuple):
                return _flat_ast_tuple(
                    [ast.BinOp(left=left, op=op, right=re) for re in right.elts]
                )
            return ast.BinOp(left=left, op=op, right=right)
        if isinstance(n, ast.UnaryOp):
            operand = fold(n.operand)
            if isinstance(operand, ast.Tuple) and isinstance(
                n.op, (ast.UAdd, ast.USub, ast.Invert)
            ):
                return _flat_ast_tuple(
                    [ast.UnaryOp(op=n.op, operand=el) for el in operand.elts]
                )
            return ast.UnaryOp(op=n.op, operand=operand)
        if isinstance(n, ast.Tuple):
            return _flat_ast_tuple([fold(e) for e in n.elts])
        if isinstance(n, ast.Call):
            return ast.Call(
                func=fold(n.func),
                args=[fold(a) for a in n.args],
                keywords=n.keywords,
            )
        if isinstance(n, ast.IfExp):
            return ast.IfExp(test=fold(n.test), body=fold(n.body), orelse=fold(n.orelse))
        if isinstance(n, ast.Attribute):
            return ast.Attribute(value=fold(n.value), attr=n.attr, ctx=n.ctx)
        if isinstance(n, ast.Subscript):
            sl = n.slice
            if isinstance(sl, ast.Slice):
                sl = ast.Slice(
                    lower=fold(sl.lower) if sl.lower else None,
                    upper=fold(sl.upper) if sl.upper else None,
                    step=fold(sl.step) if sl.step else None,
                )
            else:
                sl = fold(sl)
            return ast.Subscript(value=fold(n.value), slice=sl, ctx=n.ctx)
        return n

    return fold(node)


def _broadcast_tuple_expr_string(e: str) -> str:
    """Apply tuple broadcast then re-emit (requires Python 3.9+ ``ast.unparse``)."""
    tree = ast.parse(e, mode="eval")
    new_body = _fold_tuple_broadcast_ast(tree.body)
    wrapped = ast.Expression(body=new_body)
    ast.fix_missing_locations(wrapped)
    return ast.unparse(new_body)


def _preprocess_time_expression(expr: str) -> str:
    s = _replace_colon_time_literals(expr)
    s = _replace_compact_time_literals(s)
    s = re.sub(
        r"\bsum\(\s*(dur_parse_colon_literal\([^)]+\))\s*,\s*(dur_parse_colon_literal\([^)]+\))\s*\)",
        r"sum_dur_sm(\1, \2)",
        s,
    )
    return s


def _smoke_str_to_complex(s: str) -> Optional[complex]:
    t = s.strip().replace(" ", "")
    t = re.sub(
        r"(?<![A-Za-z0-9_])(\d+\.?\d*|\.\d+)i(?![A-Za-z0-9_])",
        r"\1j",
        t,
        flags=re.I,
    )
    t = re.sub(r"(?<![A-Za-z0-9_])i(?![A-Za-z0-9_])", "(1j)", t, flags=re.I)
    try:
        v = eval(compile(ast.parse(t, mode="eval"), "<cpx>", "eval"), {"__builtins__": {}}, {})
    except Exception:
        return None
    if isinstance(v, complex):
        return v
    if isinstance(v, (int, float)):
        return complex(float(v), 0.0)
    return None


def _complex_close(a: complex, b: complex) -> bool:
    return abs(a.real - b.real) <= smoke_tol(a.real, b.real) and abs(a.imag - b.imag) <= smoke_tol(
        a.imag, b.imag
    )


def _format_float_scalar_smoke(x: float) -> str:
    """One float as smoke expects (scientific mantissa; never ``rstrip('0')`` on ``e+20``)."""
    if math.isnan(x):
        return "nan"
    if math.isinf(x):
        return "-inf" if x < 0 else "inf"
    if math.isfinite(x) and x != 0 and (abs(x) >= 1e16 or abs(x) < 1e-4):
        s = f"{x:.15e}"
        s = re.sub(r"e\+0+(\d)", r"e+\1", s, flags=re.I)
        s = re.sub(r"e-0+(\d)", r"e-\1", s, flags=re.I)
        return s
    if abs(x - round(x)) < 1e-12 and abs(x) < 1e15:
        return str(int(round(x)))
    s = str(x)
    if "e" in s.lower():
        return s
    return s.rstrip("0").rstrip(".")


def _format_exact_imag_coeff(ai: int) -> str:
    if ai == 1:
        return "i"
    if ai == -1:
        return "-i"
    return f"{ai}i"


def _format_exact_complex_smoke(z: ExactComplex) -> str:
    if z.imag == 0:
        return str(z.real)
    if z.real == 0:
        return _format_exact_imag_coeff(z.imag)
    ti = _format_exact_imag_coeff(z.imag)
    tr = str(z.real)
    if z.imag > 0:
        return f"{tr}+{ti}"
    if z.imag == -1:
        return f"{tr}-i"
    return f"{tr}{ti}"


def _format_complex_smoke(z: complex) -> str:
    """Emit strings closer to Smart-Math smoke (``-5+10i``, ``5-i``, ``i``)."""
    r, i = z.real, z.imag
    # Smart-Math collapses any complex value with NaN in either component to "nan".
    # (It does not preserve "+2i"/"-2i" around a NaN real or imag part.)
    if math.isnan(r) or math.isnan(i):
        return "nan"
    if abs(i) < 1e-15:
        if abs(r) < 1e-15:
            return "0"
        if math.isinf(r):
            return "-inf" if r < 0 else "inf"
        return _format_float_scalar_smoke(r)
    # non-zero imaginary
    eps = 1e-12

    def fmt_re(x: float) -> str:
        if abs(x) < eps:
            return ""
        return _format_float_scalar_smoke(x)

    def imag_tail(ai: float) -> str:
        if math.isnan(ai):
            return "nan*i"
        if math.isinf(ai):
            return ("-" if ai < 0 else "") + "inf*i"
        a = abs(ai)
        neg = ai < 0
        if abs(a - 1.0) < eps:
            return "-i" if neg else "i"
        if abs(a - round(a)) < 1e-12 and a < 1e15:
            n = int(round(a))
            return f"-{n}i" if neg else f"{n}i"
        body = _format_float_scalar_smoke(a) + "i"
        return "-" + body if neg else body

    tr = fmt_re(r)
    ti = imag_tail(i)
    if not tr:
        return ti
    if i > 0:
        return f"{tr}+{ti}".replace("+-", "-")
    return f"{tr}{ti}".replace("+-", "-")


def smoke_tol(a: float, b: float) -> float:
    """Mirror ``ScalarCloseEnough`` in ``SmokeTest_MathParser.bas``."""
    scale = max(1.0, abs(a), abs(b))
    tol = 256.0 * EPS * scale
    if tol < 1e-14:
        tol = 1e-14
    return tol


_RE_DECIMAL_SCALAR = re.compile(r"^[-+]?(\d+(\.\d*)?|\.\d+)([eE][-+]?\d+)?$")
_RE_RATIONAL_SCALAR = re.compile(r"^([-+]?\d+)\s*/\s*(\d+)\s*$")


def try_parse_decimal_scalar(s: str) -> Optional[float]:
    s = s.strip()
    if _RE_DECIMAL_SCALAR.fullmatch(s):
        return float(s)
    return None


def try_parse_rational_scalar(s: str) -> Optional[float]:
    """Smoke ``ratio()`` / ``n/d`` display token (optional ``*`` suffix already stripped)."""
    s = s.strip()
    if s.endswith("*"):
        s = s[:-1].strip()
    m = _RE_RATIONAL_SCALAR.fullmatch(s)
    if not m:
        return None
    num = int(m.group(1))
    den = int(m.group(2))
    if den == 0:
        return None
    return num / den


def try_parse_smoke_scalar(s: str) -> Optional[float]:
    """Decimal/scientific literal or reduced rational ``n/d``."""
    s = s.strip()
    if not s:
        return 0.0
    rat = try_parse_rational_scalar(s)
    if rat is not None:
        return rat
    return try_parse_decimal_scalar(s)


def _parse_smoke_imag_coeff(im_s: str) -> Optional[float]:
    """Imaginary coefficient after ``SplitComplexText`` (``+i``/``-i`` -> ``±1``)."""
    t = im_s.strip()
    if not t:
        return 1.0
    if t == "-":
        return -1.0
    return try_parse_smoke_scalar(t)


def split_complex_text_smoke(s: str) -> Optional[tuple[str, str]]:
    """Mirror ``SplitComplexText`` in ``SmokeTest_MathParser.bas``."""
    t = s.strip()
    if not t:
        return ("0", "0")
    if t[-1].lower() != "i":
        return None
    split_at = 0
    for i in range(len(t) - 1, 0, -1):
        ch = t[i]
        if ch in "+-":
            if i > 0 and t[i - 1].lower() not in "e":
                split_at = i
                break
            if ch == "-" and i == 0:
                split_at = i
                break
    if split_at == 0:
        return ("0", t[:-1])
    re_part = t[:split_at].strip() or "0"
    im_part = t[split_at:].strip()
    if im_part.startswith("+"):
        im_part = im_part[1:].strip()
    if im_part.endswith(("i", "I")):
        im_part = im_part[:-1].strip()
    if im_part.endswith("*"):
        im_part = im_part[:-1].strip()
    return (re_part, im_part)


def try_parse_smoke_complex(s: str) -> Optional[complex]:
    t = s.strip().replace(" ", "")
    if not t:
        return complex(0.0, 0.0)
    parts = split_complex_text_smoke(t)
    if parts is None:
        re_v = try_parse_smoke_scalar(t)
        return complex(re_v, 0.0) if re_v is not None else None
    re_s, im_s = parts
    re_v = try_parse_smoke_scalar(re_s)
    im_v = _parse_smoke_imag_coeff(im_s)
    if re_v is None or im_v is None:
        return None
    return complex(re_v, im_v)


def scalar_display_close_enough(actual: str, expected: str) -> bool:
    """Float/scientific or rational ``n/d`` tokens (smoke ``ratio()`` display family)."""
    ea, ee = actual.strip(), expected.strip()
    if ea == ee:
        return True
    da = try_parse_smoke_scalar(ea)
    de = try_parse_smoke_scalar(ee)
    if da is None or de is None:
        return False
    if da == de:
        return True
    return abs(da - de) <= smoke_tol(da, de)


def complex_display_close_enough(actual: str, expected: str) -> bool:
    """Cartesian complex strings; rationals in either part compare numerically."""
    ca = try_parse_smoke_complex(actual)
    ce = try_parse_smoke_complex(expected)
    if ca is not None and ce is not None:
        return _complex_close(ca, ce)
    a_parts = split_complex_text_smoke(actual.strip())
    e_parts = split_complex_text_smoke(expected.strip())
    if a_parts is None or e_parts is None:
        return False
    re_ok = scalar_display_close_enough(a_parts[0], e_parts[0])
    im_a = _parse_smoke_imag_coeff(a_parts[1])
    im_e = _parse_smoke_imag_coeff(e_parts[1])
    if im_a is None or im_e is None:
        return False
    im_ok = im_a == im_e or abs(im_a - im_e) <= smoke_tol(im_a, im_e)
    return re_ok and im_ok


def close_enough_str(actual: str, expected: str) -> bool:
    return scalar_display_close_enough(actual, expected)


def split_top_level_tuple(s: str) -> Optional[list[str]]:
    s = s.strip()
    if len(s) < 2 or not (s.startswith("(") and s.endswith(")")):
        return None
    inner = s[1:-1].strip()
    if not inner:
        return []
    parts: list[str] = []
    depth = 0
    start = 0
    for i, ch in enumerate(inner):
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        elif ch == "," and depth == 0:
            parts.append(inner[start:i].strip())
            start = i + 1
    parts.append(inner[start:].strip())
    return parts


def tuple_display_close_enough(actual: str, expected: str) -> bool:
    pa, pe = split_top_level_tuple(actual), split_top_level_tuple(expected)
    if pa is None or pe is None or len(pa) != len(pe):
        return False
    for a, e in zip(pa, pe):
        if scalar_display_close_enough(a, e):
            continue
        if complex_display_close_enough(a, e):
            continue
        if a.strip() != e.strip():
            return False
    return True


def tuple_close_enough(actual: str, expected: str) -> bool:
    return tuple_display_close_enough(actual, expected)


def tuple_close_enough_complex(actual: str, expected: str) -> bool:
    return tuple_display_close_enough(actual, expected)


def _looks_like_ratio_display(expected: str, got: str, *, expr: str = "") -> bool:
    if "ratio(" in expr:
        return True
    exp = expected.strip()
    if _RE_RATIONAL_SCALAR.search(exp.replace(" ", "")):
        return True
    if "i" in exp.lower() and "/" in exp:
        return True
    if got.strip().startswith("(") and "/" in exp:
        return True
    return False


def ratio_display_close_enough(expected: str, got: str, *, expr: str = "") -> bool:
    """
    ``ratio()`` smoke output is often ``n/d`` or ``a/b+c/d*i``; Python reference uses floats.
    Compare numerically with smoke tolerance (not string identity).
    """
    exp, got_s = expected.strip(), got.strip()
    if exp == got_s:
        return True
    if exp.startswith("(") or got_s.startswith("("):
        return tuple_display_close_enough(got_s, exp)
    if "i" in exp.lower() or "i" in got_s.lower():
        return complex_display_close_enough(got_s, exp)
    return scalar_display_close_enough(got_s, exp)


def round_half_away_from_zero(x: float) -> float:
    """Smart Math-style round (e.g. round(2.5) -> 3). Non-finites pass through."""
    xf = float(x)
    if not math.isfinite(xf):
        return xf
    return float(Decimal(str(xf)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def is_multiple_of(x: float, x_mult: float) -> bool:
    """Mirror MathParser ``IsMultipleOf`` (round quotient + residual)."""
    if x_mult == 0.0:
        return False
    q = x / x_mult
    if not math.isfinite(q):
        return False
    n = round(q)
    if abs(q - n) > 1e-9:
        return False
    residual = abs(x - n * x_mult)
    eps = max(1e-14 * abs(x_mult), 8.0 * EPS * abs(x_mult))
    return residual <= eps


def sin_sm(x: float) -> float:
    if x == 0.0:
        return 0.0
    if math.isfinite(x) and is_multiple_of(x, math.pi):
        return 0.0
    return math.sin(x)


def cos_sm(x: float) -> float:
    if math.isfinite(x):
        if not is_multiple_of(x, math.pi) and is_multiple_of(x, math.pi / 2.0):
            return 0.0
    return math.cos(x)


def tan_sm(x: float) -> float:
    if x == 0.0:
        return 0.0
    if math.isfinite(x):
        if is_multiple_of(x, math.pi):
            return 0.0
        if is_multiple_of(x, math.pi / 2.0):
            t = math.tan(x)
            return math.inf if t > 0.0 else -math.inf
    return math.tan(x)


def dsl_truth(x: Any) -> bool:
    if isinstance(x, Duration):
        return x.ms != 0
    if isinstance(x, complex):
        # Smart-Math treats complex values with NaN in either component as falsey.
        if math.isnan(x.real) or math.isnan(x.imag):
            return False
        return x.real != 0.0 or x.imag != 0.0
    if isinstance(x, float):
        if math.isnan(x):
            return False
        return x != 0.0
    if isinstance(x, int):
        return x != 0
    return bool(x)


def dsl_lnot(x: Any) -> int:
    return 0 if dsl_truth(x) else 1


def dsl_and(a: Any, b: Any) -> int:
    return 1 if dsl_truth(a) and dsl_truth(b) else 0


def dsl_or(a: Any, b: Any) -> int:
    return 1 if dsl_truth(a) or dsl_truth(b) else 0


def dsl_cmp(op: str, a: Any, b: Any) -> int:
    if op == "eq":
        if isinstance(a, float) and isinstance(b, float) and math.isnan(a) and math.isnan(b):
            return 0
        if isinstance(a, (int, float)) and isinstance(b, (int, float)):
            return 1 if close_enough_str(str(a), str(b)) or a == b else 0
        return 1 if a == b else 0
    if op == "ne":
        if isinstance(a, float) and isinstance(b, float) and math.isnan(a) and math.isnan(b):
            return 1
        return 1 - dsl_cmp("eq", a, b)
    if op == "lt":
        return 1 if a < b else 0
    if op == "le":
        return 1 if a <= b else 0
    if op == "gt":
        return 1 if a > b else 0
    if op == "ge":
        return 1 if a >= b else 0
    raise ValueError(op)


def dsl_bitinv(x: Any) -> Any:
    if isinstance(x, tuple):
        return tuple(dsl_bitinv(item) for item in _flatten_value(x))
    xi = _as_int_exact(x)
    if xi is None:
        xi = int(math.trunc(float(x)))
    return ~xi


def fractional_power_is_odd_unit_root(p: float) -> Optional[int]:
    if p <= 0.0 or p >= 1.0:
        return None
    inv = 1.0 / p
    if inv < 2.0 or inv > 63.0:
        return None
    n = int(round(inv))
    if n < 2 or n > 63:
        return None
    if abs(inv - float(n)) > 1e-6:
        return None
    if (n % 2) == 0:
        return None
    return n


def _real_scalar_mag(v: Any) -> float:
    if isinstance(v, complex):
        return abs(v)
    if isinstance(v, (int, float)):
        return abs(float(v))
    raise TypeError("not numeric")


class _SmokeAstTransform(ast.NodeTransformer):
    """Comparisons and boolean ops return Smart-Math 0/1 ints."""

    def __init__(self, mode: str = "real") -> None:
        self._mode = mode

    def visit_BinOp(self, node: ast.BinOp) -> ast.AST:
        if isinstance(node.op, ast.Pow):
            fn = "pow_sm_cx" if self._mode == "complex" else "pow_sm"
        elif isinstance(node.op, ast.Mod):
            fn = "mod_sm"
        elif isinstance(node.op, ast.Add):
            fn = "add_sm"
        elif isinstance(node.op, ast.Sub):
            fn = "sub_sm"
        elif isinstance(node.op, ast.Div):
            fn = "div_sm_cx" if self._mode == "complex" else "div_sm_real"
        elif isinstance(node.op, ast.Mult):
            fn = "mul_sm_cx" if self._mode == "complex" else "mul_sm_real"
        elif isinstance(node.op, ast.LShift):
            fn = "lshift_sm"
        elif isinstance(node.op, ast.RShift):
            fn = "rshift_sm"
        elif isinstance(node.op, ast.BitAnd):
            fn = "bitand_sm"
        elif isinstance(node.op, ast.BitOr):
            fn = "bitor_sm"
        elif isinstance(node.op, ast.BitXor):
            fn = "bitxor_sm"
        else:
            return self.generic_visit(node)
        return ast.Call(
            func=ast.Name(id=fn, ctx=ast.Load()),
            args=[self.visit(node.left), self.visit(node.right)],
            keywords=[],
        )

    def visit_Compare(self, node: ast.Compare) -> ast.AST:
        left = self.visit(node.left)
        for op, comp in zip(node.ops, node.comparators):
            right = self.visit(comp)
            op_name = {
                ast.Eq: "eq",
                ast.NotEq: "ne",
                ast.Lt: "lt",
                ast.LtE: "le",
                ast.Gt: "gt",
                ast.GtE: "ge",
            }.get(type(op))
            if op_name is None:
                raise ValueError("compare op")
            left = ast.Call(
                func=ast.Name(id="dsl_cmp", ctx=ast.Load()),
                args=[ast.Constant(op_name), left, right],
                keywords=[],
            )
        return left

    def visit_BoolOp(self, node: ast.BoolOp) -> ast.AST:
        fn = "dsl_and" if isinstance(node.op, ast.And) else "dsl_or"
        acc = self.visit(node.values[0])
        for v in node.values[1:]:
            acc = ast.Call(
                func=ast.Name(id=fn, ctx=ast.Load()),
                args=[acc, self.visit(v)],
                keywords=[],
            )
        return acc

    def visit_UnaryOp(self, node: ast.UnaryOp) -> ast.AST:
        if isinstance(node.op, ast.Not):
            return ast.Call(
                func=ast.Name(id="dsl_lnot", ctx=ast.Load()),
                args=[self.visit(node.operand)],
                keywords=[],
            )
        if isinstance(node.op, ast.Invert):
            return ast.Call(
                func=ast.Name(id="dsl_bitinv", ctx=ast.Load()),
                args=[self.visit(node.operand)],
                keywords=[],
            )
        return self.generic_visit(node)

    def visit_Call(self, node: ast.Call) -> ast.AST:
        """Spread ``unpack(...)`` results into the surrounding call."""
        node = self.generic_visit(node)
        return ast.Call(
            func=ast.Name(id="expand_call_sm", ctx=ast.Load()),
            args=[node.func, *node.args],
            keywords=list(node.keywords),
        )


def _preprocess_unary_bang_real(expr: str) -> str:
    """``!`` / ``!!`` on a number, name, or parenthesized value. Rightmost first."""
    s = expr
    while True:
        matches = list(re.finditer(r"!(?!=)", s))
        if not matches:
            break
        pos = matches[-1].start()
        rest = s[pos + 1 :].lstrip()
        mnum = re.match(r"(\d+\.?\d*(?:[eE][-+]?\d+)?)", rest)
        if mnum:
            num = mnum.group(1)
            s = s[:pos] + f"dsl_lnot({num})" + rest[len(num) :]
            continue
        if rest.startswith("("):
            s = s[:pos] + "dsl_lnot(" + rest[1:]
            continue
        mname = re.match(r"([A-Za-z_][A-Za-z0-9_]*)", rest)
        if mname:
            end = mname.end()
            if end < len(rest) and rest[end] == "(":
                depth = 0
                j = end
                while j < len(rest):
                    if rest[j] == "(":
                        depth += 1
                    elif rest[j] == ")":
                        depth -= 1
                        if depth == 0:
                            j += 1
                            break
                    j += 1
                end = j
            operand = rest[:end]
            s = s[:pos] + f"dsl_lnot({operand})" + rest[end:]
            continue
        break
    return s


def _preprocess_real_dsl(expr: str) -> str:
    s = expr.replace("<>", "!=")
    s = _preprocess_unary_bang_real(s)
    out: list[str] = []
    i = 0
    while i < len(s):
        if s[i : i + 2] in ("==", "!=", "<=", ">="):
            out.append(s[i : i + 2])
            i += 2
            continue
        if s[i] == "=":
            out.append("==")
            i += 1
            continue
        out.append(s[i])
        i += 1
    s = "".join(out)
    s = re.sub(r"\s*&&\s*", " and ", s)
    s = re.sub(r"\s*\|\|\s*", " or ", s)
    return s


def _preprocess_word_not(expr: str) -> str:
    """``not`` binds looser than comparison, so ``3 + not 4 + 5`` is ``3 + not(4+5)``."""
    s = expr
    while True:
        matches = list(re.finditer(r"\bnot\b(?!\s*\()", s, flags=re.I))
        if not matches:
            break
        matched = matches[-1]
        rest = s[matched.end() :]
        depth = 0
        end = 0
        while end < len(rest):
            ch = rest[end]
            if ch == "(":
                depth += 1
            elif ch == ")":
                if depth == 0:
                    break
                depth -= 1
            elif depth == 0 and (
                rest.startswith("&&", end)
                or rest.startswith("||", end)
                or re.match(r"and\b|or\b", rest[end:], flags=re.I)
            ):
                break
            end += 1
        operand = rest[:end].strip()
        if not operand:
            break
        s = s[: matched.start()] + f"dsl_lnot({operand})" + rest[end:]
    return s


def _matching_paren(s: str, open_idx: int) -> int:
    depth = 0
    for i in range(open_idx, len(s)):
        if s[i] == "(":
            depth += 1
        elif s[i] == ")":
            depth -= 1
            if depth == 0:
                return i
    return -1


def _rewrite_all_base_vals(s: str) -> str:
    return re.sub(
        r"\b(uhex|uoct|ubin|hex|oct|bin)\s*\(",
        "base_val_sm(",
        s,
        flags=re.I,
    )


def _rewrite_formatters_for_value(expr: str) -> str:
    """A formatter that is the whole result stays a display. Nested ones return the integer."""
    s = expr.strip()
    matched = re.match(r"(hex|oct|bin|uhex|uoct|ubin)\s*\(", s, flags=re.I)
    if matched:
        open_idx = matched.end() - 1
        close_idx = _matching_paren(s, open_idx)
        if close_idx == len(s) - 1:
            inner = _rewrite_all_base_vals(s[open_idx + 1 : close_idx])
            return s[: open_idx + 1] + inner + s[close_idx:]
    return _rewrite_all_base_vals(s)


def trig_near_zero_ok(expected: str, got: str) -> bool:
    """sin/tan near mathematical zero often prints as '0' in smoke tests."""
    if expected.strip() != "0":
        return False
    g = got.strip()
    if re.fullmatch(r"[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?", g):
        return abs(float(g)) <= 5e-11
    return False


def pole_inf_ok(expected: str, got: str) -> bool:
    if expected.strip() in ("inf", "+inf"):
        if got.strip().lower() == "inf":
            return True
        if re.fullmatch(r"[-+]?(?:\d+\.?\d*)(?:[eE][-+]?\d+)?", got.strip()):
            return float(got) > 1e10
    if expected.strip() == "-inf":
        if got.strip().lower() == "-inf":
            return True
        if re.fullmatch(r"[-+]?(?:\d+\.?\d*)(?:[eE][-+]?\d+)?", got.strip()):
            return float(got) < -1e10
    return False


def results_match(
    expected: str, got_str: str, *, mode: str = "real", expr: str = ""
) -> bool:
    exp, got = expected.strip(), got_str.strip()
    if got == exp:
        return True
    if _looks_like_ratio_display(exp, got, expr=expr):
        if ratio_display_close_enough(exp, got, expr=expr):
            return True
    if mode == "complex":
        if complex_display_close_enough(got, exp):
            return True
        if tuple_close_enough_complex(got, exp):
            return True
        ca, ce = try_parse_smoke_complex(got), try_parse_smoke_complex(exp)
        if ca is None:
            ca = _smoke_str_to_complex(got)
        if ce is None:
            ce = _smoke_str_to_complex(exp)
        if ca is not None and ce is not None and _complex_close(ca, ce):
            return True
        if close_enough_str(got, exp) or tuple_close_enough(got, exp):
            return True
    elif mode == "time":
        def _strip_hhmmss_trailing_zero_seconds(x: str) -> str:
            x = x.strip()
            ps = x.split(":")
            if len(ps) == 3 and ps[2] == "00":
                return ":".join(ps[:2])
            return x

        ge, gg = _strip_hhmmss_trailing_zero_seconds(exp), _strip_hhmmss_trailing_zero_seconds(got)
        if ge == gg or tuple_close_enough(ge, gg) or close_enough_str(ge, gg):
            return True
        return False
    if complex_display_close_enough(got, exp):
        return True
    if close_enough_str(got, exp) or tuple_close_enough(got, exp):
        return True
    if trig_near_zero_ok(exp, got) or pole_inf_ok(exp, got):
        return True
    return False


def parse_smoke_cases(path: str) -> list[dict]:
    """Parse tests(N) fields from SmokeTest_MathParser.bas (same-line assignments)."""
    text = open(path, encoding="utf-8", errors="replace").read()
    lines = text.splitlines()
    cases: dict[int, dict] = {}

    re_expr = re.compile(
        r'tests\((\d+)\)\.expr\s*=\s*"((?:[^"\\]|\\.)*)"\s*:'
    )
    re_exp = re.compile(r'tests\(\d+\)\.expected\s*=\s*"((?:[^"\\]|\\.)*)"')
    re_err = re.compile(
        r'tests\(\d+\)\.expectedErrContains\s*=\s*"((?:[^"\\]|\\.)*)"'
    )
    re_nores = re.compile(r"tests\(\d+\)\.expectNoResult\s*=\s*TRUE")

    for line in lines:
        m = re_expr.search(line)
        if not m:
            continue
        idx = int(m.group(1))
        expr = m.group(2).replace('""', '"')  # FB string escape if any
        entry = cases.setdefault(
            idx,
            {
                "idx": idx,
                "expr": expr,
                "source": "smoke_main",
                "ref_mode": infer_ref_mode(expr, ""),
            },
        )
        if re_nores.search(line):
            entry["expect_no_result"] = True
        me = re_exp.search(line)
        if me:
            entry["expected"] = me.group(1).replace('""', '"')
            entry["ref_mode"] = infer_ref_mode(expr, entry["expected"])
        mer = re_err.search(line)
        if mer:
            entry["expected_err"] = mer.group(1).replace('""', '"')

    return [cases[k] for k in sorted(cases)]


_RE_RUN_COMPLEX_SUB = re.compile(
    r"(?is)private\s+sub\s+RunComplexNumberSupportOptionTests\s*\([^)]*\)\s*(.*?)(?:^\s*end\s+sub\s*$)",
    re.MULTILINE,
)

# Quoted FB string or unquoted rhs (rhs may contain ``:`` only inside quotes).
_RE_FB_RHS_TOKEN = r'(?:"(?:[^"\\]|\\.)*"|[^:\r\n]+)'

# Same-line paired table row: fooCases|fooOk|fooExprs(n) = <expr> : fooExpect|fooNeed(n) = <expect>
_RE_PAIRED_TABLE_ROW = re.compile(
    r"([A-Za-z_][A-Za-z0-9_]*)(?:Cases|Ok|Exprs)\s*\(\s*(\d+)\s*\)\s*=\s*"
    rf"({_RE_FB_RHS_TOKEN})\s*:\s*"
    r"\1(?:Expect|Need)\s*\(\s*\2\s*\)\s*=\s*"
    rf"({_RE_FB_RHS_TOKEN})",
    re.I,
)

# expr(n) = "..." : want(n) = "..."  (RunInexactFloatDisplayTests, RunFilterExactIntegerTests)
_RE_EXPR_WANT_ROW = re.compile(
    r'(?im)^\s*expr\s*\(\s*(\d+)\s*\)\s*=\s*"((?:[^"\\]|\\.)*)"\s*:\s*'
    r'want\s*\(\s*\1\s*\)\s*=\s*"((?:[^"\\]|\\.)*)"'
)

# fooExpr(n) = <expr> : fooExpect(n) = <expect>  (expr may use & string concat)
_RE_PAIRED_EXPR_EXPECT_ROW = re.compile(
    r"([A-Za-z_][A-Za-z0-9_]*)Expr\s*\(\s*(\d+)\s*\)\s*=\s*"
    rf"({_RE_FB_RHS_TOKEN}(?:\s*&\s*{_RE_FB_RHS_TOKEN})*)\s*:\s*"
    r"\1Expect\s*\(\s*\2\s*\)\s*=\s*"
    rf"({_RE_FB_RHS_TOKEN}(?:\s*&\s*{_RE_FB_RHS_TOKEN})*)",
    re.I,
)

# Legacy alias (quote-only rows are a subset of _RE_PAIRED_TABLE_ROW).
_RE_COMPLEX_OPT_PAIRED_ROW = _RE_PAIRED_TABLE_ROW

_RE_COMPLEX_OPT_PAIRED_EXPR_ROW = _RE_PAIRED_EXPR_EXPECT_ROW

# Split-line table rows (expr and expect on separate lines).
_RE_SPLIT_TABLE_LHS = re.compile(
    r"([A-Za-z_][A-Za-z0-9_]*)(?:Cases|Ok|Exprs)\s*\(\s*(\d+)\s*\)\s*=\s*(.+?)\s*$",
    re.I,
)
_RE_SPLIT_TABLE_RHS = re.compile(
    r"([A-Za-z_][A-Za-z0-9_]*)(?:Expect|Need)\s*\(\s*(\d+)\s*\)\s*=\s*(.+?)\s*$",
    re.I,
)

_RE_RUN_SUB_BODY = re.compile(
    r"(?is)(?:private\s+)?sub\s+(Run\w+Tests)\s*\([^)]*\)\s*(.*?)(?:^\s*end\s+sub\s*$)",
    re.MULTILINE,
)

# SmokeTest_MathParser.bas Run* subs (Main() driver order).
SMOKE_RUN_SUB_SOURCES: dict[str, str] = {
    "RunRawResultApiTests": "raw_api",
    "RunOperatorPrecedenceDocTests": "precedence_doc",
    "RunTrigAngleReductionTests": "trig_reduction",
    "RunComplexNumberSupportOptionTests": "complex_opt",
    "RunNegativeArgumentMagnitudeBandTests": "magnitude_band",
    "RunPositiveArgumentMagnitudeBandTests": "magnitude_band",
    "RunFloatMagnitudeLiteralTests": "float_magnitude",
    "RunNanInfTests": "naninf_mirror",
    "RunBuiltinArityTableTests": "builtin_arity",
    "RunIncompleteFunctionCallHintTests": "open_paren_hint",
    "RunGcdLcmNcrNprArrayBroadcastTests": "gcd_ncr_array",
    "RunRatioInExpressionTests": "ratio_in_expression",
    "RunMinMaxPreserveWinnerTests": "minmax_winner",
    "RunExactIntegerDivisionTests": "exact_int_div",
    "RunExactIntegerMultiplicationTests": "exact_int_mul",
    "RunBinaryBuiltinArrayLengthMismatchTests": "bin_array_len",
    "RunTimeValuesSupportOptionTests": "time_opt",
    "RunLambdaFunctionsSupportOptionTests": "lambda_opt",
}

_RE_DIM_STRING_CONST = re.compile(
    r'dim\s+([A-Za-z_][A-Za-z0-9_]*)\s+as\s+String\s*=\s*"((?:[^"\\]|\\.)*)"',
    re.I,
)

# Parser_TryEvaluateEx(var, ...) = FALSE orelse rt <> varExpect | "literal"
_RE_COMPLEX_OPT_VAR_PROBE = re.compile(
    r"Parser_TryEvaluateEx\s*\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*,\s*r\s*,\s*rt\s*,\s*ia\s*\)\s*=\s*FALSE\s+"
    r"orelse\s+rt\s+<>\s+(?:\"((?:[^\"\\]|\\.)*)\"|([A-Za-z_][A-Za-z0-9_]*))",
    re.I,
)

_RE_RUN_RAW_SUB = re.compile(
    r"(?is)private\s+sub\s+RunRawResultApiTests\s*\([^)]*\)\s*(.*?)(?:^\s*end\s+sub\s*$)",
    re.MULTILINE,
)

# Error-only table: cmpErrExpr(1) = "1+2i > 0"
_RE_COMPLEX_OPT_ERR_ROW = re.compile(
    r'([A-Za-z_][A-Za-z0-9_]*Err(?:Expr)?)\s*\(\s*(\d+)\s*\)\s*=\s*"((?:[^"\\]|\\.)*)"',
    re.I,
)

# Inline literal probe: if Parser_TryEvaluateEx("expr", r, rt, ia) then ... expected error ...
_RE_COMPLEX_OPT_INLINE_PROBE = re.compile(
    r'if\s+Parser_TryEvaluateEx\("((?:[^"\\]|\\.)*)"\s*,\s*r\s*,\s*rt\s*,\s*ia\s*\)\s+then',
    re.I,
)

# Inline success probe: ... = FALSE orelse rt <> "expect"
_RE_COMPLEX_OPT_INLINE_SUCCESS = re.compile(
    r'if\s+Parser_TryEvaluateEx\("((?:[^"\\]|\\.)*)"\s*,\s*r\s*,\s*rt\s*,\s*ia\s*\)\s*=\s*FALSE\s+'
    r'orelse\s+rt\s+<>\s+"((?:[^"\\]|\\.)*)"',
    re.I,
)


def _extract_run_complex_number_support_sub_body(text: str) -> Optional[str]:
    """Return the body of ``RunComplexNumberSupportOptionTests`` only (no outer subs)."""
    m = _RE_RUN_COMPLEX_SUB.search(text)
    return m.group(1) if m else None


def _extract_run_raw_result_sub_body(text: str) -> Optional[str]:
    m = _RE_RUN_RAW_SUB.search(text)
    return m.group(1) if m else None


def _unescape_fb_string(s: str) -> str:
    return s.replace('""', '"')


def _collect_dim_string_constants(block: str) -> dict[str, str]:
    consts: dict[str, str] = {}
    for m in _RE_DIM_STRING_CONST.finditer(block):
        consts[m.group(1)] = _unescape_fb_string(m.group(2))
    return consts


def _resolve_fb_string_piece(piece: str, consts: dict[str, str]) -> str:
    piece = piece.strip()
    m = re.fullmatch(r'"((?:[^"\\]|\\.)*)"', piece)
    if m:
        return _unescape_fb_string(m.group(1))
    if piece in consts:
        return consts[piece]
    raise ValueError(f"unresolved string piece: {piece!r}")


def _resolve_fb_string_expr(rhs: str, consts: dict[str, str]) -> str:
    rhs = rhs.strip()
    if "&" not in rhs:
        return _resolve_fb_string_piece(rhs, consts)
    return "".join(_resolve_fb_string_piece(p, consts) for p in re.split(r"\s*&\s*", rhs))


def _resolve_fb_expect_rhs(rhs: str, consts: dict[str, str]) -> str:
    rhs = rhs.strip()
    if rhs.startswith('"'):
        return _resolve_fb_string_expr(rhs, consts)
    if rhs in consts:
        return consts[rhs]
    return rhs


def _table_stem_prefix(name: str) -> str:
    for suffix in ("Exprs", "Cases", "Ok", "Expect", "Need", "Expr"):
        if name.endswith(suffix):
            return name[: -len(suffix)]
    return name


def _paired_table_idx_from_match(m: re.Match) -> str:
    lhs_m = re.match(
        r"([A-Za-z_][A-Za-z0-9_]*(?:Cases|Ok|Exprs))\s*\(\s*(\d+)\s*\)",
        m.group(0),
        re.I,
    )
    if lhs_m:
        return f"{lhs_m.group(1)}({lhs_m.group(2)})"
    stem, n_s = m.group(1), m.group(2)
    return f"{stem}({n_s})"


def _infer_case_ref_mode(expr: str, expected: str, *, default: str = "real") -> str:
    return infer_ref_mode(expr, expected) if expected else infer_ref_mode(expr, "")


def _add_paired_table_case(
    add_case,
    *,
    idx: str,
    expr: str,
    expected: str,
    source: str,
    ref_mode: Optional[str] = None,
) -> None:
    mode = ref_mode or _infer_case_ref_mode(expr, expected)
    add_case(
        {
            "idx": idx,
            "expr": expr,
            "expected": expected,
            "ref_mode": mode,
            "source": source,
        }
    )


def _parse_paired_table_rows_in_block(
    block: str,
    *,
    source: str,
    consts: dict[str, str],
    add_case,
) -> None:
    """Parse same-line *Cases/*Ok/*Exprs paired with *Expect/*Need."""
    for m in _RE_PAIRED_TABLE_ROW.finditer(block):
        stem, n_s, expr_rhs, exp_rhs = m.group(1), m.group(2), m.group(3), m.group(4)
        n = int(n_s)
        try:
            expr = _resolve_fb_string_expr(expr_rhs, consts)
            expected = _resolve_fb_expect_rhs(exp_rhs, consts)
        except ValueError:
            continue
        _add_paired_table_case(
            add_case,
            idx=_paired_table_idx_from_match(m),
            expr=expr,
            expected=expected,
            source=source,
        )

    for m in _RE_PAIRED_EXPR_EXPECT_ROW.finditer(block):
        stem, n_s, expr_rhs, exp_rhs = m.group(1), m.group(2), m.group(3), m.group(4)
        n = int(n_s)
        try:
            expr = _resolve_fb_string_expr(expr_rhs, consts)
            expected = _resolve_fb_expect_rhs(exp_rhs, consts)
        except ValueError:
            continue
        _add_paired_table_case(
            add_case,
            idx=f"{stem}Expr({n})",
            expr=expr,
            expected=expected,
            source=source,
        )


def _parse_split_line_table_rows_in_block(
    block: str,
    *,
    source: str,
    consts: dict[str, str],
    add_case,
) -> None:
    """Pair ``fooOk(3) = ...`` on one line with ``fooExpect(3) = ...`` on another."""
    lhs: dict[tuple[str, int], tuple[str, str]] = {}
    for line in block.splitlines():
        m = _RE_SPLIT_TABLE_LHS.match(line.strip())
        if not m:
            continue
        stem, n_s, rhs = m.group(1), m.group(2), m.group(3)
        if ":" in rhs:
            continue
        try:
            expr = _resolve_fb_string_expr(rhs, consts)
        except ValueError:
            continue
        lhs[(_table_stem_prefix(stem), int(n_s))] = (stem, expr)
    for line in block.splitlines():
        m = _RE_SPLIT_TABLE_RHS.match(line.strip())
        if not m:
            continue
        stem, n_s, rhs = m.group(1), m.group(2), m.group(3)
        key = (_table_stem_prefix(stem), int(n_s))
        if key not in lhs:
            continue
        lhs_stem, expr = lhs[key]
        try:
            expected = _resolve_fb_expect_rhs(rhs, consts)
        except ValueError:
            continue
        _add_paired_table_case(
            add_case,
            idx=f"{lhs_stem}({n_s})",
            expr=expr,
            expected=expected,
            source=source,
        )


_RE_ERR_ONLY_EXPRS_ROW = re.compile(
    r"(?:badExprs|sortbyExprs|rejectExprs|noExprs|timeAggErr|badCloserExprs)\s*\(\s*(\d+)\s*\)\s*=\s*"
    r'"((?:[^"\\]|\\.)*)"',
    re.I,
)


def _parse_err_tables_in_block(block: str, *, source: str, add_case) -> None:
    for m in _RE_ERR_ONLY_EXPRS_ROW.finditer(block):
        n_s, expr_s = m.group(1), m.group(2)
        expr = _unescape_fb(expr_s)
        add_case(
            {
                "idx": f"err-only/{expr[:40]}({n_s})",
                "expr": expr,
                "expect_error": True,
                "ref_mode": infer_ref_mode(expr, ""),
                "source": source,
            }
        )
    for m in _RE_COMPLEX_OPT_ERR_ROW.finditer(block):
        stem, n_s, expr_s = m.group(1), m.group(2), m.group(3)
        expr = _unescape_fb(expr_s)
        add_case(
            {
                "idx": f"errtbl/{stem}({n_s})",
                "expr": expr,
                "expect_error": True,
                "ref_mode": infer_ref_mode(expr, ""),
                "source": source,
            }
        )
    for m in _RE_ERR_EXPR_NEED.finditer(block):
        n_s, expr_s, need_s = m.group(1), m.group(2), m.group(3)
        expr = _unescape_fb(expr_s)
        add_case(
            {
                "idx": f"arity-err/{n_s}",
                "expr": expr,
                "expect_error": True,
                "expected_err_substr": _unescape_fb(need_s),
                "ref_mode": infer_ref_mode(expr, ""),
                "source": source,
            }
        )


def _parse_inline_probes_in_block(block: str, *, source: str, add_case) -> None:
    for m in _RE_INLINE_NUMERIC.finditer(block):
        expr = _unescape_fb(m.group(1))
        expected = _unescape_fb(m.group(2))
        add_case(
            {
                "idx": f"inline:{expr[:80]}",
                "expr": expr,
                "expected": expected,
                "ref_mode": _infer_case_ref_mode(expr, expected),
                "source": source,
            }
        )
    for m in _RE_COMPLEX_OPT_INLINE_SUCCESS.finditer(block):
        expr = _unescape_fb_string(m.group(1))
        expected = _unescape_fb_string(m.group(2))
        add_case(
            {
                "idx": f"inline:{expr[:80]}",
                "expr": expr,
                "expected": expected,
                "ref_mode": _infer_case_ref_mode(expr, expected),
                "source": source,
            }
        )
    for m in _RE_COMPLEX_OPT_INLINE_PROBE.finditer(block):
        expr = _unescape_fb_string(m.group(1))
        tail = block[m.end() : m.end() + 400]
        if re.search(r"expected\s+error|expected\s+failure", tail, re.I):
            add_case(
                {
                    "idx": f"inline-err:{expr[:80]}",
                    "expr": expr,
                    "expect_error": True,
                    "ref_mode": infer_ref_mode(expr, ""),
                    "source": source,
                }
            )


def _parse_paired_expr_expect_tables(block: str, consts: dict[str, str], add_case, *, source: str = "complex_opt") -> None:
    _parse_paired_table_rows_in_block(block, source=source, consts=consts, add_case=add_case)


def parse_smoke_synthetic_recent_cases() -> list[dict]:
    """Probes that live only in raw/API subs (no tests(N).expected)."""
    return [
        {
            "idx": "raw-api/sqrt-non-perfect-square-display",
            "expr": "sqrt(4611686014132420611)",
            "expected": "2147483647",
            "ref_mode": "real",
            "source": "raw_api",
        },
    ]


_RE_PREC_PAIR = re.compile(
    r'precOk\s*\(\s*(\d+)\s*\)\s*=\s*"((?:[^"\\]|\\.)*)"\s*:\s*precExpect\s*\(\s*\1\s*\)\s*=\s*"((?:[^"\\]|\\.)*)"',
    re.I,
)

_RE_TRIG_NONZERO = re.compile(
    r'nonzeroExpr\s*\(\s*(\d+)\s*\)\s*=\s*"((?:[^"\\]|\\.)*)"',
    re.I,
)

_RE_TRIG_NEARZERO = re.compile(
    r'nearZeroExpr\s*\(\s*(\d+)\s*\)\s*=\s*"((?:[^"\\]|\\.)*)"',
    re.I,
)

_RE_PAIRED_EXPR_EXPECT = re.compile(
    r"([A-Za-z_][A-Za-z0-9_]*)Expr\s*\(\s*(\d+)\s*\)\s*=\s*\"((?:[^\"\\]|\\.)*)\"\s*:\s*"
    r"\1Expect\s*\(\s*\2\s*\)\s*=\s*\"((?:[^\"\\]|\\.)*)\"",
    re.I,
)

_RE_ERR_EXPR_NEED = re.compile(
    r'errExprs\s*\(\s*(\d+)\s*\)\s*=\s*"((?:[^"\\]|\\.)*)"\s*:\s*'
    r'errNeed\s*\(\s*\1\s*\)\s*=\s*"((?:[^"\\]|\\.)*)"',
    re.I,
)

_RE_INLINE_NUMERIC = re.compile(
    r'Parser_TryEvaluateEx\s*\(\s*"((?:[^"\\]|\\.)*)"\s*,\s*r\s*,\s*rt\s*,\s*ia\s*\)\s*=\s*FALSE\s+'
    r'orelse\s+rt\s+<>\s+"((?:[^"\\]|\\.)*)"',
    re.I,
)


def _unescape_fb(s: str) -> str:
    return s.replace('""', '"')


def parse_smoke_sub_array_cases(path: str) -> list[dict]:
    """``precOk``/``precExpect`` and trig magnitude tables from SmokeTest_MathParser.bas."""
    text = open(path, encoding="utf-8", errors="replace").read()
    out: list[dict] = []
    seen: set[str] = set()
    for m in _RE_PREC_PAIR.finditer(text):
        n, expr, expected = m.group(1), _unescape_fb(m.group(2)), _unescape_fb(m.group(3))
        key = f"precedence:{expr}"
        if key in seen:
            continue
        seen.add(key)
        out.append(
            {
                "idx": f"precedence-doc/{n}",
                "expr": expr,
                "expected": expected,
                "ref_mode": infer_ref_mode(expr, expected),
                "source": "precedence_doc",
            }
        )
    for m in _RE_TRIG_NONZERO.finditer(text):
        n, expr = m.group(1), _unescape_fb(m.group(2))
        key = f"trig-nz:{expr}"
        if key in seen:
            continue
        seen.add(key)
        out.append(
            {
                "idx": f"trig-reduction/nonzero/{n}",
                "expr": expr,
                "assert_kind": "abs_gt",
                "assert_threshold": 1e-6,
                "ref_mode": "real",
                "source": "trig_reduction",
            }
        )
    for m in _RE_TRIG_NEARZERO.finditer(text):
        n, expr = m.group(1), _unescape_fb(m.group(2))
        key = f"trig-z:{expr}"
        if key in seen:
            continue
        seen.add(key)
        out.append(
            {
                "idx": f"trig-reduction/nearzero/{n}",
                "expr": expr,
                "assert_kind": "abs_lt",
                "assert_threshold": 1e-9,
                "ref_mode": "real",
                "source": "trig_reduction",
            }
        )
    for m in _RE_PAIRED_EXPR_EXPECT.finditer(text):
        stem, n_s, expr_s, exp_s = m.group(1), m.group(2), m.group(3), m.group(4)
        expr = _unescape_fb(expr_s)
        expected = _unescape_fb(exp_s)
        key = f"{stem}-expr:{expr}"
        if key in seen:
            continue
        seen.add(key)
        out.append(
            {
                "idx": f"{stem}-expr/{n_s}",
                "expr": expr,
                "expected": expected,
                "ref_mode": infer_ref_mode(expr, expected),
                "source": "paired_expr_expect",
            }
        )
    return out


def parse_smoke_inline_numeric_probes(path: str) -> list[dict]:
    """``Parser_TryEvaluateEx(...) orelse rt <> "expect"`` rows across smoke subs."""
    text = open(path, encoding="utf-8", errors="replace").read()
    out: list[dict] = []
    seen: set[str] = set()
    for m in _RE_INLINE_NUMERIC.finditer(text):
        expr = _unescape_fb(m.group(1))
        expected = _unescape_fb(m.group(2))
        if "/" in expected and not re.search(r"\d+/\d+", expected):
            continue
        if expected.strip() in ("0", "1") and re.search(
            r"(?<![<>!=])=(?!=)|<>|!=|==", expr
        ):
            pass
        key = f"{expr}|{expected}"
        if key in seen:
            continue
        seen.add(key)
        out.append(
            {
                "idx": f"inline:{expr[:60]}",
                "expr": expr,
                "expected": expected,
                "ref_mode": infer_ref_mode(expr, expected),
                "source": "inline_probe",
            }
        )
    return out


def parse_mirror_table_stems(
    text: str,
    tables: list[tuple[str, Optional[str], str]],
) -> list[dict]:
    """
  Parse generated mirror tables: ``{stem}Expr``, ``{stem}Expect``, ``{stem}IsErr``, …

  *tables*: list of (stem, ref_mode or None to infer per row, source tag).
  """
    out: list[dict] = []
    for table, ref_mode, source in tables:
        exprs: dict[int, str] = {}
        expects: dict[int, str] = {}
        is_err: dict[int, bool] = {}
        err_txt: dict[int, str] = {}
        for m in re.finditer(
            rf"{table}Expr\s*\(\s*(\d+)\s*\)\s*=\s*\"((?:[^\"\\]|\\.)*)\"",
            text,
        ):
            exprs[int(m.group(1))] = _unescape_fb(m.group(2))
        for m in re.finditer(
            rf"{table}Expect\s*\(\s*(\d+)\s*\)\s*=\s*\"((?:[^\"\\]|\\.)*)\"",
            text,
        ):
            expects[int(m.group(1))] = _unescape_fb(m.group(2))
        for m in re.finditer(
            rf"{table}IsErr\s*\(\s*(\d+)\s*\)\s*=\s*(TRUE|FALSE)",
            text,
            re.I,
        ):
            is_err[int(m.group(1))] = m.group(2).upper() == "TRUE"
        for m in re.finditer(
            rf"{table}Err\s*\(\s*(\d+)\s*\)\s*=\s*\"((?:[^\"\\]|\\.)*)\"",
            text,
        ):
            err_txt[int(m.group(1))] = _unescape_fb(m.group(2))
        for i in sorted(exprs):
            label_m = re.search(
                rf"{table}Label\s*\(\s*{i}\s*\)\s*=\s*\"((?:[^\"\\]|\\.)*)\"",
                text,
            )
            label = _unescape_fb(label_m.group(1)) if label_m else f"{table}/{i}"
            expr = exprs[i]
            mode = ref_mode
            if mode is None:
                mode = infer_ref_mode(expr, expects.get(i, ""))
            entry: dict = {
                "idx": f"{source}/{table}/{i}",
                "expr": expr,
                "source": source,
                "ref_mode": mode,
                "band_label": label,
            }
            if is_err.get(i):
                entry["expect_error"] = True
                if i in err_txt:
                    entry["expected_err_substr"] = err_txt[i]
            elif i in expects:
                entry["expected"] = expects[i]
            out.append(entry)
    return out


def parse_band_cases_file(path: str) -> list[dict]:
    """Parse ``neg_band_cases_generated.bas`` (and positive/complex tables)."""
    text = open(path, encoding="utf-8", errors="replace").read()
    return parse_mirror_table_stems(
        text,
        [
            ("negBandReal", "real", "magnitude_band"),
            ("posBandReal", "real", "magnitude_band"),
            ("negBandCx", "complex", "magnitude_band"),
            ("posBandCx", "complex", "magnitude_band"),
        ],
    )


def parse_tools_generated_cases(tools_dir: str) -> list[dict]:
    """All ``#include`` generated tables under ``tools/`` (except neg band file)."""
    out: list[dict] = []
    paths = [
        (
            f"{tools_dir}/complex_coverage_generated.bas",
            [("cxCov", "complex", "complex_coverage")],
        ),
        (
            f"{tools_dir}/naninf_mirror_generated.bas",
            [("naninf", "real", "naninf_mirror")],
        ),
        (
            f"{tools_dir}/float_magnitude_cases_generated.bas",
            [("floatMag", "real", "float_magnitude")],
        ),
        (
            f"{tools_dir}/smoke_parity_append_generated.bas",
            [("smokeAppend", None, "smoke_append")],
        ),
    ]
    for path, stems in paths:
        try:
            text = open(path, encoding="utf-8", errors="replace").read()
        except OSError:
            continue
        out.extend(parse_mirror_table_stems(text, stems))
    return out


def _parse_expr_want_rows_in_block(block: str, *, source: str, add_case) -> None:
    """Parse ``expr(n) = "...": want(n) = "..."`` tables."""
    for matched in _RE_EXPR_WANT_ROW.finditer(block):
        expr = _unescape_fb_string(matched.group(2))
        expected = _unescape_fb_string(matched.group(3))
        _add_paired_table_case(
            add_case,
            idx=f"expr({matched.group(1)})",
            expr=expr,
            expected=expected,
            source=source,
        )


def parse_smoke_run_sub_cases(path: str) -> list[dict]:
    """
    Parse every ``Run*Tests`` sub in ``SmokeTest_MathParser.bas``: paired tables,
    error tables, and inline numeric/error probes.
    """
    text = open(path, encoding="utf-8", errors="replace").read()
    file_consts = _collect_dim_string_constants(text)
    out: list[dict] = []
    seen: set[str] = set()

    def add(entry: dict) -> None:
        key = "|".join(
            (
                entry.get("source", ""),
                entry.get("idx", ""),
                entry.get("expr", ""),
                entry.get("expected", ""),
                str(entry.get("expect_error", "")),
            )
        )
        if key in seen:
            return
        seen.add(key)
        out.append(entry)

    for m in _RE_RUN_SUB_BODY.finditer(text):
        sub_name = m.group(1)
        body = m.group(2)
        source = SMOKE_RUN_SUB_SOURCES.get(sub_name, sub_name)
        consts = {**file_consts, **_collect_dim_string_constants(body)}
        _parse_paired_table_rows_in_block(body, source=source, consts=consts, add_case=add)
        _parse_expr_want_rows_in_block(body, source=source, add_case=add)
        _parse_split_line_table_rows_in_block(body, source=source, consts=consts, add_case=add)
        _parse_err_tables_in_block(body, source=source, add_case=add)
        _parse_inline_probes_in_block(body, source=source, add_case=add)

        if sub_name == "RunComplexNumberSupportOptionTests":
            for vm in _RE_COMPLEX_OPT_VAR_PROBE.finditer(body):
                var, lit_exp, id_exp = vm.group(1), vm.group(2), vm.group(3)
                if var not in consts:
                    continue
                expr = consts[var]
                if lit_exp is not None:
                    expected = _unescape_fb_string(lit_exp)
                elif id_exp in consts:
                    expected = consts[id_exp]
                else:
                    continue
                add(
                    {
                        "idx": f"var:{var}",
                        "expr": expr,
                        "expected": expected,
                        "ref_mode": "complex",
                        "source": source,
                    }
                )

    return out


def parse_smoke_global_tables(path: str) -> list[dict]:
    """
    Legacy/global scan for paired and error tables outside ``Run*Tests`` subs
    (e.g. ``Main()``). Most rows are covered by ``parse_smoke_run_sub_cases``.
    """
    text = open(path, encoding="utf-8", errors="replace").read()
    run_spans = [(m.start(), m.end()) for m in _RE_RUN_SUB_BODY.finditer(text)]

    def outside_run_subs(pos: int) -> bool:
        return not any(s <= pos < e for s, e in run_spans)

    out: list[dict] = []
    seen: set[str] = set()

    def add(entry: dict) -> None:
        key = f"{entry.get('expr','')}|{entry.get('expected','')}|{entry.get('expect_error')}"
        if key in seen:
            return
        seen.add(key)
        out.append(entry)

    consts = _collect_dim_string_constants(text)
    for m in _RE_PAIRED_TABLE_ROW.finditer(text):
        if not outside_run_subs(m.start()):
            continue
        _parse_paired_table_rows_in_block(
            m.group(0),
            source="paired_ok_expect",
            consts=consts,
            add_case=add,
        )
    for m in _RE_PAIRED_EXPR_EXPECT_ROW.finditer(text):
        if not outside_run_subs(m.start()):
            continue
        try:
            expr = _resolve_fb_string_expr(m.group(3), consts)
            expected = _resolve_fb_expect_rhs(m.group(4), consts)
        except ValueError:
            continue
        add(
            {
                "idx": f"{m.group(1)}Expr({m.group(2)})",
                "expr": expr,
                "expected": expected,
                "ref_mode": _infer_case_ref_mode(expr, expected),
                "source": "paired_expr_expect",
            }
        )

    for m in _RE_COMPLEX_OPT_ERR_ROW.finditer(text):
        if not outside_run_subs(m.start()):
            continue
        stem, n_s, expr_s = m.group(1), m.group(2), m.group(3)
        expr = _unescape_fb(expr_s)
        add(
            {
                "idx": f"errtbl/{stem}({n_s})",
                "expr": expr,
                "expect_error": True,
                "ref_mode": infer_ref_mode(expr, ""),
                "source": "error_expr_table",
            }
        )

    return out


def dedupe_cases(cases: list[dict]) -> list[dict]:
    """Drop duplicate probes (complex-opt block overlaps global paired-row scan)."""
    out: list[dict] = []
    seen: set[str] = set()
    for c in cases:
        key = "|".join(
            (
                c.get("expr", ""),
                c.get("expected", ""),
                str(c.get("expect_error", "")),
                c.get("assert_kind", ""),
                str(c.get("assert_threshold", "")),
            )
        )
        if key in seen:
            continue
        seen.add(key)
        out.append(c)
    return out


def collect_all_numeric_cases(smoke_bas: str, band_bas: str, tools_dir: str) -> list[dict]:
    batches = (
        parse_smoke_cases(smoke_bas)
        + parse_smoke_run_sub_cases(smoke_bas)
        + parse_smoke_sub_array_cases(smoke_bas)
        + parse_smoke_global_tables(smoke_bas)
        + parse_band_cases_file(band_bas)
        + parse_tools_generated_cases(tools_dir)
        + parse_smoke_synthetic_recent_cases()
    )
    return dedupe_cases(batches)


def _count_table_expr_rows(text: str, stem: str) -> int:
    return len(re.findall(rf"{re.escape(stem)}Expr\s*\(\s*\d+\s*\)\s*=", text, re.I))


def _count_discovered_rows_in_block(block: str) -> dict[str, int]:
    return {
        "paired_table": len(_RE_PAIRED_TABLE_ROW.findall(block)),
        "paired_expr": len(_RE_PAIRED_EXPR_EXPECT_ROW.findall(block)),
        "err_table": len(_RE_COMPLEX_OPT_ERR_ROW.findall(block)),
        "err_only_exprs": len(_RE_ERR_ONLY_EXPRS_ROW.findall(block)),
        "arity_err": len(_RE_ERR_EXPR_NEED.findall(block)),
        "inline_success": len(_RE_INLINE_NUMERIC.findall(block))
        + len(_RE_COMPLEX_OPT_INLINE_SUCCESS.findall(block)),
        "inline_err_probe": len(_RE_COMPLEX_OPT_INLINE_PROBE.findall(block)),
        "trig_nonzero": len(_RE_TRIG_NONZERO.findall(block)),
        "trig_nearzero": len(_RE_TRIG_NEARZERO.findall(block)),
    }


def audit_coverage(smoke_bas: str, band_bas: str, tools_dir: str, cases: list[dict]) -> None:
    """Print discovered row counts vs cases loaded (sanity check for parsers)."""
    smoke_text = open(smoke_bas, encoding="utf-8", errors="replace").read()
    parsed_by_source: dict[str, int] = {}
    for c in cases:
        src = c.get("source", "?")
        parsed_by_source[src] = parsed_by_source.get(src, 0) + 1

    re_tests_expr = re.compile(r"tests\(\d+\)\.expr\s*=", re.I)
    re_tests_expected = re.compile(r"tests\(\d+\)\.expected\s*=", re.I)
    re_tests_err = re.compile(r"tests\(\d+\)\.expectedErrContains\s*=", re.I)
    print("=== Parser coverage audit ===")
    print(f"tests(N).expr rows (Main table):      {len(re_tests_expr.findall(smoke_text))}")
    print(f"tests(N).expected rows:             {len(re_tests_expected.findall(smoke_text))}")
    print(f"tests(N).expectedErrContains rows:  {len(re_tests_err.findall(smoke_text))}")
    print(f"Paired *Cases|*Ok|*Exprs rows:      {len(_RE_PAIRED_TABLE_ROW.findall(smoke_text))}")
    print(f"Paired *Expr/*Expect rows:           {len(_RE_PAIRED_EXPR_EXPECT_ROW.findall(smoke_text))}")
    print(f"*Err* expr table rows:              {len(_RE_COMPLEX_OPT_ERR_ROW.findall(smoke_text))}")
    print(f"Inline rt<> success probes:         {len(_RE_INLINE_NUMERIC.findall(smoke_text))}")
    print(f"Run*Tests subs found:               {len(_RE_RUN_SUB_BODY.findall(smoke_text))}")
    print()
    print("Run*Tests sub inventory (discovered rows in source vs parsed by source tag):")
    for m in _RE_RUN_SUB_BODY.finditer(smoke_text):
        sub_name = m.group(1)
        body = m.group(2)
        source = SMOKE_RUN_SUB_SOURCES.get(sub_name, sub_name)
        disc = _count_discovered_rows_in_block(body)
        disc_total = sum(disc.values())
        parsed = parsed_by_source.get(source, 0)
        if parsed >= disc_total or disc_total == 0:
            flag = " OK"
        elif parsed >= disc_total - 6:
            flag = " OK (dedupe with tests(N) likely)"
        else:
            flag = " GAP"
        print(
            f"  {sub_name}: discovered~{disc_total} "
            f"(paired={disc['paired_table']+disc['paired_expr']} "
            f"err={disc['err_table']+disc['err_only_exprs']+disc['arity_err']} "
            f"inline={disc['inline_success']}) parsed={parsed}{flag}"
        )
    print()
    for label, path, stem in (
        ("neg/pos band", band_bas, "negBandReal"),
        ("cxCov", f"{tools_dir}/complex_coverage_generated.bas", "cxCov"),
        ("naninf", f"{tools_dir}/naninf_mirror_generated.bas", "naninf"),
        ("floatMag", f"{tools_dir}/float_magnitude_cases_generated.bas", "floatMag"),
        ("smokeAppend", f"{tools_dir}/smoke_parity_append_generated.bas", "smokeAppend"),
    ):
        try:
            t = open(path, encoding="utf-8", errors="replace").read()
            if stem == "negBandReal":
                n = sum(
                    _count_table_expr_rows(t, s)
                    for s in ("negBandReal", "posBandReal", "negBandCx", "posBandCx")
                )
            else:
                n = _count_table_expr_rows(t, stem)
            src_tag = {
                "negBandReal": "magnitude_band",
                "cxCov": "complex_coverage",
                "naninf": "naninf_mirror",
                "floatMag": "float_magnitude",
                "smokeAppend": "smoke_append",
            }[stem]
            parsed = parsed_by_source.get(src_tag, 0)
            print(f"{label} Expr rows in file: {n}  (parsed tag {src_tag}: {parsed})")
        except OSError:
            print(f"{label}: (file missing)")
    print()
    print("Parsed cases by source tag:")
    for src in sorted(parsed_by_source):
        print(f"  {src}: {parsed_by_source[src]}")
    comparable = sum(
        1
        for c in cases
        if "expected" in c or c.get("assert_kind") or c.get("expect_error")
    )
    print(f"Total deduped cases with expected/assert/error: {comparable}")
    missing_run_subs = [
        name for name in SMOKE_RUN_SUB_SOURCES if name not in {m.group(1) for m in _RE_RUN_SUB_BODY.finditer(smoke_text)}
    ]
    if missing_run_subs:
        print(f"WARNING: expected Run* subs not found: {', '.join(missing_run_subs)}")
    print()


def parse_smoke_raw_api_cases(path: str) -> list[dict]:
    """Inline numeric probes from ``RunRawResultApiTests`` (display string, not raw kind)."""
    text = open(path, encoding="utf-8", errors="replace").read()
    block = _extract_run_raw_result_sub_body(text)
    if not block:
        return []
    out: list[dict] = []
    seen: set[str] = set()

    def add(entry: dict) -> None:
        key = entry["idx"]
        if key in seen:
            return
        seen.add(key)
        out.append(entry)

    for m in _RE_COMPLEX_OPT_INLINE_SUCCESS.finditer(block):
        expr = _unescape_fb_string(m.group(1))
        expected = _unescape_fb_string(m.group(2))
        add(
            {
                "idx": f"raw-inline:{expr}",
                "expr": expr,
                "expected": expected,
                "ref_mode": infer_ref_mode(expr, expected),
                "source": "raw_api",
            }
        )
    return out


def parse_smoke_complex_opt_cases(path: str) -> list[dict]:
    """
    Parse numeric and error probes from ``RunComplexNumberSupportOptionTests``.

    * Paired rows: ``fooCases(n) = "expr": fooExpect(n) = "exp"`` or
      ``fooOk(n): fooExpect(n)`` on one line.
    * ``fooExpr(n) = <string expr> : fooExpect(n) = <expect>`` (supports ``&`` concat).
    * ``dim var as String = "lit"`` with ``Parser_TryEvaluateEx(var, ...) orelse rt <> varExpect``.
    * Error tables: ``cmpErrExpr(n) = "expr"`` (no expected string).
    * Inline ``Parser_TryEvaluateEx("literal", ...)`` success/error probes.
    """
    text = open(path, encoding="utf-8", errors="replace").read()
    block = _extract_run_complex_number_support_sub_body(text)
    if not block:
        return []

    consts = _collect_dim_string_constants(block)

    out: list[dict] = []
    stem_first_pos: dict[str, int] = {}
    seen_keys: set[str] = set()

    def add_case(entry: dict) -> None:
        key = entry["idx"]
        if key in seen_keys:
            return
        seen_keys.add(key)
        stem = re.match(r"([A-Za-z_][A-Za-z0-9_]*)", key)
        if stem:
            stem_first_pos.setdefault(stem.group(1).lower(), len(out))
        out.append(entry)

    _parse_paired_expr_expect_tables(block, consts, add_case)

    for m in _RE_COMPLEX_OPT_VAR_PROBE.finditer(block):
        var, lit_exp, id_exp = m.group(1), m.group(2), m.group(3)
        if var not in consts:
            continue
        expr = consts[var]
        if lit_exp is not None:
            expected = _unescape_fb_string(lit_exp)
        elif id_exp in consts:
            expected = consts[id_exp]
        else:
            continue
        add_case(
            {
                "idx": f"var:{var}",
                "expr": expr,
                "expected": expected,
                "ref_mode": "complex",
                "source": "complex_opt",
            }
        )

    for m in _RE_COMPLEX_OPT_PAIRED_ROW.finditer(block):
        stem, n_s, expr_s, exp_s = m.group(1), m.group(2), m.group(3), m.group(4)
        n = int(n_s)
        expr = expr_s.replace('""', '"')
        expected = exp_s.replace('""', '"')
        add_case(
            {
                "idx": f"{stem}({n})",
                "expr": expr,
                "expected": expected,
                "ref_mode": "complex",
                "source": "complex_opt",
            }
        )

    for m in _RE_COMPLEX_OPT_ERR_ROW.finditer(block):
        stem, n_s, expr_s = m.group(1), m.group(2), m.group(3)
        n = int(n_s)
        expr = expr_s.replace('""', '"')
        add_case(
            {
                "idx": f"{stem}({n})",
                "expr": expr,
                "expect_error": True,
                "ref_mode": "complex",
                "source": "complex_opt",
            }
        )

    for m in _RE_COMPLEX_OPT_INLINE_SUCCESS.finditer(block):
        expr_s, exp_s = m.group(1), m.group(2)
        expr = expr_s.replace('""', '"')
        expected = exp_s.replace('""', '"')
        add_case(
            {
                "idx": f"inline:{expr}",
                "expr": expr,
                "expected": expected,
                "ref_mode": "complex",
                "source": "complex_opt",
            }
        )

    for m in _RE_COMPLEX_OPT_INLINE_PROBE.finditer(block):
        expr_s = m.group(1)
        expr = expr_s.replace('""', '"')
        tail = block[m.end() : m.end() + 400]
        if re.search(r"expected\s+error|expected\s+failure", tail, re.I):
            add_case(
                {
                    "idx": f"inline-err:{expr}",
                    "expr": expr,
                    "expect_error": True,
                    "ref_mode": "complex",
                    "source": "complex_opt",
                }
            )

    def _key(c: dict) -> tuple[int, int, str]:
        lab = c["idx"]
        stem_m = re.match(r"([A-Za-z_][A-Za-z0-9_]*)", lab)
        stem = stem_m.group(1).lower() if stem_m else lab.lower()
        n_m = re.search(r"\((\d+)\)", lab)
        n = int(n_m.group(1)) if n_m else 0
        return (stem_first_pos.get(stem, 10**9), n, lab)

    out.sort(key=_key)
    return out


# --- Smart-Math-ish reference evaluation ---

def _strip_line_comment(expr: str) -> str:
    if "//" in expr:
        expr = expr.split("//", 1)[0]
    return expr.strip()


def _replace_int_literals(expr: str) -> str:
    def hex_repl(m: re.Match) -> str:
        return str(int(m.group(0), 16))

    def bin_repl(m: re.Match) -> str:
        return str(int(m.group(0), 2))

    def oct_repl(m: re.Match) -> str:
        return str(int(m.group(0), 8))

    s = expr
    s = re.sub(r"\b0[xX][0-9a-fA-F]+\b", hex_repl, s)
    s = re.sub(r"\b0[bB][01]+\b", bin_repl, s)
    s = re.sub(r"\b0[oO][0-7]+\b", oct_repl, s)
    return s


def _implicit_mul(expr: str) -> str:
    s = expr
    # Numeric literal before ``(`` only (not ``log10_sm(`` / ``foo2(``).
    s = re.sub(r"(?<![A-Za-z0-9_])(\d+\.?\d*|\.\d+)\s*\(", r"\1*(", s)
    s = re.sub(r"\)\s*\(", r")*(", s)
    return s


_RE_POSTFIX_NUMBER = re.compile(r"(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?$")


def _primary_before(s: str, end: int) -> Optional[tuple[int, str]]:
    """Primary (number or parenthesized group) that ends just before ``end``."""
    k = end - 1
    while k >= 0 and s[k].isspace():
        k -= 1
    if k < 0:
        return None
    if s[k] == ")":
        depth = 0
        j = k
        while j >= 0:
            if s[j] == ")":
                depth += 1
            elif s[j] == "(":
                depth -= 1
                if depth == 0:
                    return j, s[j : k + 1]
            j -= 1
        return None
    matched = _RE_POSTFIX_NUMBER.search(s[: k + 1])
    if matched is None or matched.end() != k + 1:
        return None
    start = matched.start()
    if start > 0 and s[start - 1] in "+-":
        prev = s[start - 2] if start >= 2 else ""
        if start == 1 or prev in "(*+-/^%,":
            start -= 1
    return start, s[start : k + 1]


def _next_postfix_percent(s: str) -> int:
    """Index of a ``%`` that is postfix, not the modulo operator."""
    i = 0
    while i < len(s):
        if s[i] == "%":
            j = i + 1
            while j < len(s) and s[j].isspace():
                j += 1
            if j >= len(s) or not (s[j].isalnum() or s[j] in "._("):
                return i
        i += 1
    return -1


def _postfix_percent(expr: str) -> Optional[str]:
    """
    Postfix ``%`` is a percentage, not modulo.

    ``200+15%`` -> ``200 + 200*(15/100)``. ``2*3%`` -> ``2*(3/100)``.
    ``5+(2*3)%`` -> ``5 + 5*((2*3)/100)``. ``5%3`` stays modulo.
    """
    if "%" not in expr:
        return expr
    s = expr
    for _ in range(32):
        i = _next_postfix_percent(s)
        if i < 0:
            return s
        found = _primary_before(s, i)
        if found is None:
            return None
        primary_start, primary = found
        pct = f"(({primary})/100)"
        stripped = s[:primary_start].rstrip()
        if stripped and stripped[-1] in "+-" and len(stripped) >= 2:
            op_index = len(stripped) - 1
            before_op = stripped[:op_index]
            base = before_op.strip()
            if base and base[-1] not in "+-*/%(,":
                op = stripped[op_index]
                base_start = len(before_op) - len(before_op.lstrip())
                if op == "+":
                    repl = f"({base})+({base})*{pct}"
                else:
                    repl = f"({base})-({base})*{pct}"
                s = s[:base_start] + repl + s[i + 1 :]
                continue
        s = s[:primary_start] + pct + s[i + 1 :]
    return None


def _normalize_nonfinite_names(expr: str) -> str:
    """Accept ``Inf`` / ``INF`` / ``NaN`` the way smoke writes them."""
    return re.sub(r"\b(?:inf|nan)\b", lambda m: m.group(0).lower(), expr, flags=re.I)


_RE_LAMBDA_HEAD = re.compile(
    r"(?<![A-Za-z0-9_])(?:\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*\)|([A-Za-z_][A-Za-z0-9_]*))\s*:"
)
_RE_LAMBDA_MULTI = re.compile(
    r"(?<![A-Za-z0-9_])(\(\s*[A-Za-z_][A-Za-z0-9_]*(?:\s*,\s*[A-Za-z_][A-Za-z0-9_]*)+\s*\)|"
    r"[A-Za-z_][A-Za-z0-9_]*(?:\s*,\s*[A-Za-z_][A-Za-z0-9_]*)+)\s*:"
)


def _strip_lambda_grouping(body: str) -> str:
    """``(x):(1/x)`` parentheses group the body. A comma keeps an array body."""
    b = body.strip()
    if len(b) < 2 or b[0] != "(" or b[-1] != ")":
        return b
    depth = 0
    for i, ch in enumerate(b):
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
            if depth == 0 and i != len(b) - 1:
                return b
    inner = b[1:-1]
    depth = 0
    for ch in inner:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        elif ch == "," and depth == 0:
            return b
    return inner.strip() or b


def _bracket_depth_at(expr: str, pos: int) -> int:
    depth = 0
    for ch in expr[:pos]:
        if ch == "[":
            depth += 1
        elif ch == "]" and depth:
            depth -= 1
    return depth


def _rewrite_lambdas(expr: str) -> str:
    """Turn ``x: body`` and ``x,y: body`` into ``lambda_sm("names", "body")``."""
    out: list[str] = []
    i = 0
    while i < len(expr):
        multi = _RE_LAMBDA_MULTI.search(expr, i)
        single = _RE_LAMBDA_HEAD.search(expr, i)
        empty = re.search(r"(?<![A-Za-z0-9_])\(\s*\)\s*:", expr[i:])
        empty_at = (i + empty.start()) if empty else None
        matched = None
        names = ""
        if empty_at is not None and (
            (multi is None or empty_at <= multi.start())
            and (single is None or empty_at <= single.start())
        ):
            matched = empty
            # Rebase so matched.start/end are indexes into expr.
            class _At:
                def __init__(self, start: int, end: int) -> None:
                    self._start = start
                    self._end = end

                def start(self) -> int:
                    return self._start

                def end(self) -> int:
                    return self._end

            matched = _At(empty_at, i + empty.end())
            names = ""
        elif multi is not None and (single is None or multi.start() <= single.start()):
            matched = multi
            names = ",".join(re.findall(r"[A-Za-z_][A-Za-z0-9_]*", multi.group(1)))
        elif single is not None:
            matched = single
            names = single.group(1) or single.group(2)
        if matched is None:
            out.append(expr[i:])
            break
        if _bracket_depth_at(expr, matched.start()) > 0:
            out.append(expr[i : matched.start() + 1])
            i = matched.start() + 1
            continue
        depth = 0
        j = matched.end()
        while j < len(expr):
            ch = expr[j]
            if ch == "(":
                depth += 1
            elif ch == ")":
                if depth == 0:
                    break
                depth -= 1
            elif ch == "," and depth == 0:
                break
            j += 1
        body = _strip_lambda_grouping(expr[matched.end() : j].strip())
        if not body:
            out.append(expr[i : matched.end()])
            i = matched.end()
            continue
        body_esc = body.replace("\\", "\\\\").replace('"', '\\"')
        out.append(expr[i : matched.start()])
        out.append(f'lambda_sm("{names}","{body_esc}")')
        i = j
    return "".join(out)


def _rewrite_empty_tuples(expr: str) -> str:
    """``()`` is an empty array. ``name()`` stays a call."""
    out: list[str] = []
    i = 0
    while i < len(expr):
        if expr[i] == "(":
            j = i + 1
            while j < len(expr) and expr[j].isspace():
                j += 1
            if j < len(expr) and expr[j] == ")":
                k = i - 1
                while k >= 0 and expr[k].isspace():
                    k -= 1
                if k < 0 or not (expr[k].isalnum() or expr[k] in "_)"):
                    out.append("_empty_tuple")
                    i = j + 1
                    continue
        out.append(expr[i])
        i += 1
    return "".join(out)


def _preprocess_complex_dsl(expr: str) -> str:
    """Smart-Math DSL in complex-option block: ``<>``, ``=``, ``!`` / ``not``."""
    s = expr.replace("<>", "!=")
    s = re.sub(r"!\s*\(", "complex_not_sm(", s)
    s = re.sub(r"\bnot\s+\(", "complex_not_sm(", s, flags=re.I)
    out: list[str] = []
    i = 0
    while i < len(s):
        if s[i : i + 2] in ("==", "!=", "<=", ">="):
            out.append(s[i : i + 2])
            i += 2
            continue
        if s[i] == "=":
            out.append("==")
            i += 1
            continue
        out.append(s[i])
        i += 1
    s = "".join(out)
    # Reuse real-mode boolean operator replacement so complex-mode expressions
    # remain parseable (Smart-Math `&&` / `||` are not Python operators).
    s = re.sub(r"\s*&&\s*", " and ", s)
    s = re.sub(r"\s*\|\|\s*", " or ", s)
    return s


def _translate_functions(expr: str, *, mode: str = "real") -> Optional[str]:
    """Rename / wrap calls so ast.parse + eval with NS works."""
    s = expr
    repl = {
        "ln": "log_ln",
        "log10": "log10_sm",
        "product": "prod_sm",
        "log": "log_sm",
        "round": "round_sm",
        "pow": "pow_sm",
        "sqr": "sqr_sm",
        "prod": "prod_sm",
        "mean": "mean_sm",
        "avg": "avg_sm",
        "fact": "fact_sm",
        "factorial": "fact_sm",
        "mod": "mod_sm",
        "arcsin": "asin",
        "arccos": "acos",
        "arctan": "atan",
        "deg": "deg_sm",
        "rad": "rad_sm",
        "sign": "sign_sm",
        "clamp": "clamp_sm",
        "gcd": "gcd_sm",
        "lcm": "lcm_sm",
        "median": "median_sm",
        "variance": "variance_sm",
        "stddev": "stddev_sm",
        "min": "min_sm",
        "max": "max_sm",
        "sum": "sum_sm",
        "random": "random_sm",
        "rand": "rand_sm",
        "int": "int_sm",
        "trunc": "int_sm",
        "frac": "frac_sm",
        "fract": "frac_sm",
        "ncr": "ncr_sm",
        "npr": "npr_sm",
        "factorint": "factorint_sm",
        "cart": "cart_sm",
        "flat": "flat_sm",
        "len": "len_sm",
        "length": "len_sm",
        "range": "range_sm",
        "repeat": "repeat_sm",
        "reverse": "reverse_sm",
        "reversed": "reverse_sm",
        "sort": "sort_sm",
        "sorted": "sort_sm",
        "unique": "unique_sm",
        "unpack": "unpack_sm",
        "filter": "filter_sm",
        "sortby": "sortby_sm",
    }
    repl.update(
        {
            "hex": "hex_sm",
            "bin": "bin_sm",
            "oct": "oct_sm",
            "uhex": "uhex_sm",
            "uoct": "uoct_sm",
            "ubin": "ubin_sm",
            "sqrt": "sqrt_sm_real",
            "ratio": "ratio_sm",
            "real": "real_sm",
            "imag": "imag_sm",
            "conj": "conj_sm",
            "phase": "phase_sm",
            "polar": "polar_sm",
        }
    )
    if mode == "complex":
        repl.update(
            {
                "int": "int_sm_cx",
                "trunc": "int_sm_cx",
                "frac": "frac_sm_cx",
                "fract": "frac_sm_cx",
                "floor": "floor_sm_cx",
                "ceil": "ceil_sm_cx",
                "round": "round_sm_cx",
                "real": "real_sm",
                "imag": "imag_sm",
                "conj": "conj_sm",
                "phase": "phase_sm",
                "polar": "polar_sm",
                "cart": "cart_sm",
                "sqrt": "sqrt_sm",
                "hypot": "hypot_sm",
            }
        )
    # Longer names first so ``log`` does not alter ``log10`` / ``log10_sm``.
    for old, new in sorted(repl.items(), key=lambda kv: -len(kv[0])):
        s = re.sub(rf"\b{old}\b", new, s)
    return s


def _split_statements(expr: str) -> list[str]:
    parts = []
    depth = 0
    start = 0
    for i, ch in enumerate(expr):
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        elif ch == ";" and depth == 0:
            parts.append(expr[start:i].strip())
            start = i + 1
    parts.append(expr[start:].strip())
    return [p for p in parts if p]


def _flatten_value(v: Any) -> Any:
    """Smart-Math arrays have no nested arrays; splice tuples into one flat tuple."""
    if not isinstance(v, tuple):
        return v
    flat: list[Any] = []
    for item in v:
        if isinstance(item, tuple):
            flat.extend(_flatten_value(item))
        else:
            flat.append(item)
    return tuple(flat)


def _flatten_args(*args: Any) -> list[Any]:
    flat: list[Any] = []
    for arg in args:
        if isinstance(arg, tuple):
            flat.extend(_flatten_value(arg))
        else:
            flat.append(arg)
    return flat


def _eval_tuple_literal(expr: str, ns: dict, *, mode: str = "real") -> Any:
    expr = expr.strip()
    if not (expr.startswith("(") and expr.endswith(")")):
        raise ValueError("not tuple")
    inner = expr[1:-1].strip()
    if not inner:
        return tuple()
    depth = 0
    start = 0
    elems: list[str] = []
    for i, ch in enumerate(inner):
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        elif ch == "," and depth == 0:
            elems.append(inner[start:i].strip())
            start = i + 1
    elems.append(inner[start:].strip())
    return _flatten_value(tuple(ref_eval(e, ns, mode=mode) for e in elems))


def _log_scalar(a, b):
    if a <= 0 or b <= 0 or b == 1:
        raise ValueError("domain")
    return math.log(float(a)) / math.log(float(b))


def log_sm(a, b=None):
    if b is None:
        raise TypeError("log needs 2 args")
    return _broadcast2(_log_scalar, a, b)


def log_ln(x):
    xf = float(x)
    if math.isnan(xf) or xf < 0.0:
        return math.nan
    if xf == 0.0:
        return -math.inf
    return math.log(xf)


def log10_sm(x):
    xf = float(x)
    if math.isnan(xf) or xf < 0.0:
        return math.nan
    if xf == 0.0:
        return -math.inf
    return math.log10(xf)


def atanh_sm(x: Any) -> Any:
    if isinstance(x, tuple):
        return tuple(atanh_sm(item) for item in _flatten_value(x))
    xf = float(x)
    if xf == 1.0:
        return math.inf
    if xf == -1.0:
        return -math.inf
    return _domain_nan(math.atanh)(xf)


def _domain_nan(fn):
    """Real inverse trig/hyperbolic: a domain error is ``nan``, not an exception."""

    def wrapped(x: Any) -> Any:
        try:
            xf = float(x)
        except (TypeError, ValueError):
            return fn(x)
        if math.isnan(xf):
            return math.nan
        try:
            return fn(xf)
        except (ValueError, OverflowError):
            return math.nan

    return wrapped


_UINT64_MAX = (1 << 64) - 1
_INT64_MIN = -(1 << 63)


def _is_plain_int(x: Any) -> bool:
    return isinstance(x, int) and not isinstance(x, bool)


def _exact_int_pow(b: int, e: int) -> Optional[int]:
    """Integer power when the result fits Smart-Math int64/uint64; else None (float path)."""
    if e < 0:
        return None
    result = b**e
    if result > _UINT64_MAX or result < _INT64_MIN:
        return None
    return result


def _pow_overflow(bf: float, ef: float) -> float:
    if ef < 0.0:
        return 0.0
    if bf < 0.0 and math.isfinite(ef) and ef.is_integer() and int(ef) % 2:
        return -math.inf
    return math.inf


def pow_sm(b, e):
    if isinstance(b, tuple) or isinstance(e, tuple):
        return _broadcast2(pow_sm, b, e)
    if _is_plain_int(b) and _is_plain_int(e):
        exact = _exact_int_pow(int(b), int(e))
        if exact is not None:
            return exact
    bf = float(b) if not isinstance(b, complex) else b
    ef = float(e) if not isinstance(e, complex) else e
    if isinstance(bf, complex) or isinstance(ef, complex):
        try:
            return bf**ef
        except (OverflowError, ValueError, ZeroDivisionError):
            return complex(math.nan, math.nan)
    try:
        if bf < 0.0 and math.isfinite(bf) and math.isfinite(ef):
            n_root = fractional_power_is_odd_unit_root(ef)
            if n_root is not None:
                return -(abs(bf) ** ef)
        return bf**ef
    except OverflowError:
        return _pow_overflow(bf, ef)
    except (ValueError, ZeroDivisionError):
        return math.nan


def _as_int_exact(x: Any) -> Optional[int]:
    if isinstance(x, bool):
        return int(x)
    if isinstance(x, int):
        return x
    if isinstance(x, float) and math.isfinite(x) and x.is_integer():
        r = int(x)
        if abs(r) < (1 << 63):
            return r
    return None


def _map_unary_real(fn):
    def wrapped(x: Any) -> Any:
        if isinstance(x, tuple):
            return tuple(fn(e) for e in _flatten_value(x))
        return fn(x)

    return wrapped


def _broadcast2(fn, a, b):
    """Element-wise binary call. One array maps against a scalar; two arrays must match."""
    a_is = isinstance(a, tuple)
    b_is = isinstance(b, tuple)
    if a_is:
        a = _flatten_value(a)
    if b_is:
        b = _flatten_value(b)
    if a_is and b_is:
        if len(a) != len(b):
            raise ValueError("incompatible operands")
        return tuple(fn(x, y) for x, y in zip(a, b))
    if a_is:
        return tuple(fn(x, b) for x in a)
    if b_is:
        return tuple(fn(a, y) for y in b)
    return fn(a, b)


def base_val_sm(x: Any) -> Any:
    """Integer contributed by ``bin``/``oct``/``hex`` when the call is not the displayed result."""
    if isinstance(x, tuple):
        return tuple(base_val_sm(item) for item in _flatten_value(x))
    xi = _as_int_exact(x)
    if xi is None:
        raise ValueError("base value expects integer")
    return xi


def hex_val_sm(x: Any) -> Any:
    return base_val_sm(x)


def _trunc_nonfinite(v: Any) -> Any:
    if isinstance(v, bool):
        return float(int(v))
    if isinstance(v, int):
        return v
    if isinstance(v, float) and not math.isfinite(v):
        return v
    return float(math.trunc(float(v)))


def int_sm(x: Any) -> Any:
    if isinstance(x, tuple):
        return tuple(int_sm(e) for e in _flatten_value(x))
    if isinstance(x, ExactComplex):
        return x
    if isinstance(x, complex):
        return complex(_trunc_nonfinite(x.real), _trunc_nonfinite(x.imag))
    return _trunc_nonfinite(x)


def sqrt_sm_real(x: Any) -> Any:
    """Smart-Math sqrt: exact int only when round(sqrt(n))^2 == n (checked int multiply)."""
    if isinstance(x, tuple):
        return tuple(sqrt_sm_real(e) for e in x)
    if isinstance(x, complex):
        return sqrt_sm(x)
    xf = float(x)
    if math.isnan(xf) or xf < 0.0:
        return math.nan
    xi = _as_int_exact(x)
    if xi is None:
        return math.sqrt(xf)
    r = math.sqrt(float(xi))
    if not math.isfinite(r):
        return r
    n = int(round(r))
    if n * n == xi:
        return n
    return r


def sqr_sm(x):
    xi = _as_int_exact(x)
    if xi is not None:
        return xi * xi
    return x * x


_RATIO_EPS = 1e-14
_RATIO_MAX_DEN = 10_000_000
_RATIO_MAX_POW10 = 18
_EXACT_FLOAT_INT_LIMIT = 1 << 53


class RatioValue(float):
    """Float that prints as Smart-Math's reduced ``n/d`` when it is the result."""

    text: str

    def __new__(cls, value: float, text: str) -> "RatioValue":
        obj = float.__new__(cls, value)
        obj.text = text
        return obj


def _ratio_abs_num(v: float, p: int, q: int) -> float:
    return abs(v * float(q) - float(p))


def _ratio_err_less(n1: float, q1: int, n2: float, q2: int) -> bool:
    return (n1 * float(q2)) < (n2 * float(q1))


def _ratio_consider(v: float, p: int, q: int, best: dict) -> bool:
    if q <= 0 or q > _RATIO_MAX_DEN:
        return False
    abs_num = _ratio_abs_num(v, p, q)
    if _ratio_err_less(abs_num, q, best["abs"], best["den"]):
        best["abs"] = abs_num
        best["num"] = p
        best["den"] = q
        return abs_num == 0.0
    return False


def _ratio_exact_pow10(v: float, best: dict) -> None:
    for k in range(1, _RATIO_MAX_POW10 + 1):
        den_pow = 10**k
        scaled = v * float(den_pow)
        n = int(round(scaled))
        if n == 0:
            continue
        scale_err = max(_RATIO_EPS, _RATIO_EPS * abs(scaled))
        if abs(scaled - float(n)) > scale_err:
            continue
        g = math.gcd(abs(n), den_pow)
        cand_num = n // g
        cand_den = den_pow // g
        cand_abs = _ratio_abs_num(v, cand_num, cand_den)
        if _ratio_err_less(cand_abs, cand_den, best["abs"], best["den"]):
            best["abs"] = cand_abs
            best["num"] = cand_num
            best["den"] = cand_den


def _ratio_scan_semi(v: float, p1: int, q1: int, p0: int, q0: int, h_max: int, best: dict) -> bool:
    if h_max < 0:
        return False
    if h_max > 64:
        if _ratio_consider(v, p1, q1, best):
            return True
        lo, hi = 0, h_max
        while hi - lo > 2:
            span = hi - lo
            m1 = lo + span // 3
            m2 = hi - span // 3
            qh1, qh2 = q1 + m1 * q0, q1 + m2 * q0
            if qh1 <= 0 or qh2 <= 0:
                break
            n1 = _ratio_abs_num(v, p1 + m1 * p0, qh1)
            n2 = _ratio_abs_num(v, p1 + m2 * p0, qh2)
            if _ratio_err_less(n1, qh1, n2, qh2):
                hi = m2
            else:
                lo = m1
        for h_fin in range(lo, hi + 1):
            if _ratio_consider(v, p1 + h_fin * p0, q1 + h_fin * q0, best):
                return True
        return False
    for h_lin in range(h_max + 1):
        if _ratio_consider(v, p1 + h_lin * p0, q1 + h_lin * q0, best):
            return True
    return False


def _approximate_rational(x: float) -> tuple[int, int]:
    """Continued-fraction ratio with the same denominator cap and power-of-ten path."""
    neg = x < 0.0
    v = abs(x)
    scale = v if v >= 1.0 else 1.0
    if v < _RATIO_EPS * scale:
        return 0, 1
    best = {"num": 0, "den": 1, "abs": 1e300}
    p0, p1, q0, q1 = 0, 1, 1, 0
    cf_val = v
    for _ in range(64):
        a = math.trunc(cf_val)
        p = a * p1 + p0
        q = a * q1 + q0
        if q > _RATIO_MAX_DEN:
            if q1 > 0 and q0 > 0:
                h_max = (_RATIO_MAX_DEN - q1) // q0
                _ratio_scan_semi(v, p1, q1, p0, q0, h_max, best)
            break
        if _ratio_consider(v, p, q, best):
            break
        t = cf_val - float(a)
        if t <= _RATIO_EPS:
            break
        cf_val = 1.0 / t
        p0, p1, q0, q1 = p1, p, q1, q
    if v <= 1.0 / float(_RATIO_MAX_DEN):
        _ratio_exact_pow10(v, best)
    num, den = best["num"], best["den"]
    if num == 0:
        return 0, 1
    g = math.gcd(abs(num), den)
    num //= g
    den //= g
    if neg:
        num = -num
    return num, den


def _ratio_text(x: float) -> Optional[str]:
    if not math.isfinite(x) or abs(x) > _EXACT_FLOAT_INT_LIMIT:
        return None
    num, den = _approximate_rational(x)
    if den == 1:
        return str(num)
    return f"{num}/{den}"


def _ratio_number(x: float) -> Any:
    if not math.isfinite(x) or abs(x) > _EXACT_FLOAT_INT_LIMIT:
        return x
    text = _ratio_text(x)
    if text is None:
        return x
    return RatioValue(x, text)


def ratio_sm(x: Any) -> Any:
    """``ratio()`` value. The final result prints as a reduced fraction."""
    if isinstance(x, tuple):
        return tuple(ratio_sm(e) for e in x)
    if isinstance(x, ExactComplex):
        if x.imag == 0:
            return _ratio_number(float(x.real))
        return complex(float(x.real), float(x.imag))
    if isinstance(x, complex):
        if math.isnan(x.real) or math.isnan(x.imag):
            return complex(math.nan, math.nan)
        if abs(x.imag) < 1e-15:
            return _ratio_number(float(x.real))
        return complex(float(x.real), float(x.imag))
    if isinstance(x, Duration):
        return _ratio_number(x.ms / _SECOND_MS)
    return _ratio_number(float(x))


def _zmul(x: float, y: float) -> float:
    """0 * inf is 0. Python complex multiply turns that pair into nan."""
    if x == 0.0 or y == 0.0:
        if math.isinf(x) or math.isinf(y):
            return 0.0
        return x * y
    return x * y


def _cmul(a: Any, b: Any) -> complex:
    aa, bb = _as_complex(a), _as_complex(b)
    return complex(
        _zmul(aa.real, bb.real) - _zmul(aa.imag, bb.imag),
        _zmul(aa.real, bb.imag) + _zmul(aa.imag, bb.real),
    )


def _cdiv(a: Any, b: Any) -> complex:
    bb = _as_complex(b)
    den = _zmul(bb.real, bb.real) + _zmul(bb.imag, bb.imag)
    if den == 0.0 or not math.isfinite(den):
        return complex(math.nan, math.nan)
    numer = _cmul(a, complex(bb.real, -bb.imag))
    return complex(numer.real / den, numer.imag / den)


def mul_sm_cx(a, b):
    if isinstance(a, tuple) or isinstance(b, tuple):
        return _broadcast2(mul_sm_cx, a, b)
    if isinstance(a, (complex, ExactComplex)) or isinstance(b, (complex, ExactComplex)):
        return _cmul(a, b)
    return a * b


def div_sm_cx(a, b):
    if isinstance(a, tuple) or isinstance(b, tuple):
        return _broadcast2(div_sm_cx, a, b)
    if b == 0 or isinstance(a, (complex, ExactComplex)) or isinstance(b, (complex, ExactComplex)):
        return _cdiv(a, b)
    return a / b


def _div_scalar_real(a, b):
    if isinstance(a, Duration) or isinstance(b, Duration):
        return a / b
    if isinstance(a, (complex, ExactComplex)) or isinstance(b, (complex, ExactComplex)):
        return _cdiv(a, b)
    af = float(a)
    bf = float(b)
    if math.isnan(af) or math.isnan(bf) or (af == 0.0 and bf == 0.0):
        return math.nan
    if bf == 0.0:
        return math.copysign(math.inf, af if af != 0.0 else 1.0)
    try:
        return af / bf
    except OverflowError:
        return math.copysign(math.inf, af)


def div_sm_real(a, b):
    if isinstance(a, tuple) or isinstance(b, tuple):
        return _broadcast2(div_sm_real, a, b)
    return _div_scalar_real(a, b)


def add_sm(a, b):
    if isinstance(a, tuple) or isinstance(b, tuple):
        return _broadcast2(add_sm, a, b)
    return a + b


def sub_sm(a, b):
    if isinstance(a, tuple) or isinstance(b, tuple):
        return _broadcast2(sub_sm, a, b)
    return a - b


def mul_sm_real(a, b):
    if isinstance(a, tuple) or isinstance(b, tuple):
        return _broadcast2(mul_sm_real, a, b)
    return a * b


def _as_bit_int(x: Any) -> int:
    xi = _as_int_exact(x)
    if xi is None:
        raise ValueError("bitwise operands must be integer values")
    return xi


def lshift_sm(a, b):
    return _broadcast2(lambda x, y: _as_bit_int(x) << _as_bit_int(y), a, b)


def rshift_sm(a, b):
    return _broadcast2(lambda x, y: _as_bit_int(x) >> _as_bit_int(y), a, b)


def bitand_sm(a, b):
    return _broadcast2(lambda x, y: _as_bit_int(x) & _as_bit_int(y), a, b)


def bitor_sm(a, b):
    return _broadcast2(lambda x, y: _as_bit_int(x) | _as_bit_int(y), a, b)


def bitxor_sm(a, b):
    return _broadcast2(lambda x, y: _as_bit_int(x) ^ _as_bit_int(y), a, b)


def prod_sm(*args):
    flat = []
    for item in _flatten_args(*args):
        flat.append(item.numeric() if isinstance(item, _FactorPower) else item)
    if not flat:
        return 1
    acc = flat[0]
    for item in flat[1:]:
        if isinstance(acc, (complex, ExactComplex)) or isinstance(item, (complex, ExactComplex)):
            acc = _cmul(acc, item)
        else:
            acc = acc * item
    return acc


def mean_sm(*args):
    flat = _flatten_args(*args)
    if flat and all(isinstance(item, Duration) for item in flat):
        return Duration(sum(item.ms for item in flat) / len(flat))
    return statistics.mean(flat)


def avg_sm(*args):
    return mean_sm(*args)


def _flatten_args_cx(*args: Any) -> list[Any]:
    return _flatten_args(*args)


def mean_sm_cx(*args: Any) -> complex:
    flat = _flatten_args_cx(*args)
    if not flat:
        raise ValueError("empty")
    total = complex(0.0, 0.0)
    for item in flat:
        total += _as_complex(item)
    n = float(len(flat))
    return complex(total.real / n, total.imag / n)


def avg_sm_cx(*args: Any) -> complex:
    return mean_sm_cx(*args)


def fact_sm(n):
    """Mirror ``TryApplyFactorialScalarInt``: table through 20, then multiply until non-finite."""
    if not float(n).is_integer() or n < 0:
        raise ValueError("fact")
    n = int(n)
    if n <= 20:
        return float(math.factorial(n))
    d = float(math.factorial(20))
    for fi in range(21, n + 1):
        d *= float(fi)
        if not math.isfinite(d):
            break
    return d


def _trunc_div_int(a_i: int, b_i: int) -> int:
    """Quotient toward zero, so the remainder keeps the dividend's sign."""
    q = abs(a_i) // abs(b_i)
    if (a_i < 0) != (b_i < 0):
        q = -q
    return q


def _mod_scalar(a, b):
    """FreeBASIC ``mod`` / ``%``: remainder with sign of dividend."""
    a_i, b_i = int(a), int(b)
    if b_i == 0:
        raise ValueError("mod")
    return a_i - b_i * _trunc_div_int(a_i, b_i)


def mod_sm(a, b):
    return _broadcast2(_mod_scalar, a, b)


def _deg_one(x):
    if isinstance(x, complex) and abs(x.imag) > 1e-15:
        raise ValueError("incompatible operands")
    xv = x.real if isinstance(x, complex) else float(x)
    return xv * 180.0 / math.pi


def deg_sm(*args):
    if len(args) == 1 and not isinstance(args[0], tuple):
        return _deg_one(args[0])
    return tuple(_deg_one(item) for item in _flatten_args(*args))


def _rad_one(x):
    if isinstance(x, complex) and abs(x.imag) > 1e-15:
        raise ValueError("incompatible operands")
    xv = x.real if isinstance(x, complex) else float(x)
    return xv * math.pi / 180.0


def rad_sm(*args):
    if len(args) == 1 and not isinstance(args[0], tuple):
        return _rad_one(args[0])
    return tuple(_rad_one(item) for item in _flatten_args(*args))


def _as_complex(z: Any) -> complex:
    if isinstance(z, ExactComplex):
        return complex(float(z.real), float(z.imag))
    if isinstance(z, complex):
        return z
    if isinstance(z, (int, float)):
        return complex(float(z), 0.0)
    raise TypeError("not scalar")


def _map_unary_cx(fn):
    """Element-wise on tuples (Smart-Math array builtins)."""

    def wrapped(x: Any) -> Any:
        if isinstance(x, tuple):
            return tuple(wrapped(e) for e in x)
        return fn(x)

    return wrapped


def complex_not_sm(z: Any) -> int:
    """Smart-Math ``!`` / ``not``: true only when both Cartesian parts are ~0."""
    c = _as_complex(z)
    eps = 1e-15
    return 1 if abs(c.real) < eps and abs(c.imag) < eps else 0


def sign_sm(x):
    if isinstance(x, tuple):
        return tuple(sign_sm(e) for e in _flatten_value(x))
    if isinstance(x, (complex, ExactComplex)):
        c = _as_complex(x)
        mag = abs(c)
        if mag < 1e-15:
            return 0
        return c / mag
    return -1 if x < 0 else (1 if x > 0 else 0)


def sign_sm_real(x):
    return -1 if x < 0 else (1 if x > 0 else 0)


def _sqrt_negative_real(real: float) -> complex:
    """Principal square root: non-negative imaginary part (ignore a -0 imag sign)."""
    return cmath.sqrt(complex(real, 0.0))


def sqrt_sm(x):
    if isinstance(x, tuple):
        return tuple(sqrt_sm(e) for e in _flatten_value(x))
    if isinstance(x, ExactComplex):
        if x.imag == 0 and x.real < 0:
            return _sqrt_negative_real(float(x.real))
        return cmath.sqrt(_as_complex(x))
    if isinstance(x, complex):
        if abs(x.imag) <= 1e-15 and x.real < 0:
            return _sqrt_negative_real(x.real)
        return cmath.sqrt(x)
    if isinstance(x, (int, float)) and not isinstance(x, bool) and x < 0:
        return _sqrt_negative_real(float(x))
    return sqrt_sm_real(x)


def _format_hex_component(v: int, *, unsigned: bool = False) -> str:
    if unsigned and v < 0:
        uv = v % (1 << 64)
        return "0x" + format(uv, "X")
    if v < 0:
        return "-0x" + format(-v, "X")
    return "0x" + format(v, "X")


def _format_hex_imag_coeff(ai: float) -> str:
    if abs(ai) < 1e-15:
        return "0"
    neg = ai < 0
    a = abs(ai)
    if abs(a - 1.0) < 1e-12:
        body = "i" if not neg else "-i"
    else:
        iv = int(round(a)) if abs(a - round(a)) < 1e-12 else None
        if iv is not None:
            body = (f"0x{format(iv, 'X')}i" if iv != 1 else "i")
        else:
            body = f"{a}i".rstrip("0").rstrip(".")
        if neg:
            body = "-" + body.lstrip("-")
    return body


def _format_hex_imag_coeff_int(ai: int) -> str:
    if ai == 0:
        return "0"
    neg = ai < 0
    a = abs(ai)
    if a == 1:
        return "-i" if neg else "i"
    body = f"0x{format(a, 'X')}i"
    return "-" + body if neg else body


def _nonfinite_base_text(x: Any) -> Optional[str]:
    if isinstance(x, float) and math.isnan(x):
        return "nan"
    if isinstance(x, float) and math.isinf(x):
        return "-inf" if x < 0 else "inf"
    if isinstance(x, complex) and (math.isnan(x.real) or math.isnan(x.imag)):
        return "nan"
    return None


def hex_sm(*args: Any, unsigned: bool = False) -> Any:
    if len(args) != 1:
        return tuple(hex_sm(e, unsigned=unsigned) for e in _flatten_args(*args))
    x = args[0]
    if isinstance(x, tuple):
        return tuple(hex_sm(e, unsigned=unsigned) for e in _flatten_value(x))
    nonfinite = _nonfinite_base_text(x)
    if nonfinite is not None:
        return nonfinite
    if isinstance(x, ExactComplex):
        re_s = _format_hex_component(x.real, unsigned=unsigned)
        im_s = _format_hex_imag_coeff_int(x.imag)
        if im_s in ("0", "i", "-i"):
            if im_s == "0":
                return re_s
            if not re_s or re_s == "0x0":
                return im_s
        if x.imag > 0:
            return f"{re_s}+{im_s}"
        return f"{re_s}{im_s}"
    if isinstance(x, complex):
        re_s = _format_hex_component(int(round(x.real)), unsigned=unsigned)
        im_s = _format_hex_imag_coeff(x.imag)
        if im_s in ("0", "i", "-i"):
            if im_s == "0":
                return re_s
            if not re_s or re_s == "0x0":
                return im_s
        if x.imag > 0:
            return f"{re_s}+{im_s}"
        return f"{re_s}{im_s}"
    xi = _as_int_exact(x)
    if xi is None:
        raise ValueError("hex expects integer values")
    return _format_hex_component(xi, unsigned=unsigned)


def bin_sm(*args: Any) -> Any:
    if len(args) != 1:
        return tuple(bin_sm(e) for e in _flatten_args(*args))
    x = args[0]
    if isinstance(x, tuple):
        return tuple(bin_sm(e) for e in _flatten_value(x))
    nonfinite = _nonfinite_base_text(x)
    if nonfinite is not None:
        return nonfinite
    xi = _as_int_exact(x)
    if xi is None:
        raise ValueError("bin expects integer values")
    if xi < 0:
        return "-0b" + format(-xi, "b")
    return "0b" + format(xi, "b")


def oct_sm(*args: Any) -> Any:
    if len(args) != 1:
        return tuple(oct_sm(e) for e in _flatten_args(*args))
    x = args[0]
    if isinstance(x, tuple):
        return tuple(oct_sm(e) for e in _flatten_value(x))
    nonfinite = _nonfinite_base_text(x)
    if nonfinite is not None:
        return nonfinite
    xi = _as_int_exact(x)
    if xi is None:
        raise ValueError("oct expects integer values")
    if xi < 0:
        return "-0o" + format(-xi, "o")
    return "0o" + format(xi, "o")


def uhex_sm(x: Any) -> Any:
    return hex_sm(x, unsigned=True)


def uoct_sm(x: Any) -> Any:
    if isinstance(x, tuple):
        return tuple(uoct_sm(e) for e in x)
    nonfinite = _nonfinite_base_text(x)
    if nonfinite is not None:
        return nonfinite
    xi = _as_int_exact(x)
    if xi is None:
        raise ValueError("uoct expects integer values")
    uv = xi % (1 << 64) if xi < 0 else xi
    return "0o" + format(uv, "o")


def ubin_sm(x: Any) -> Any:
    if isinstance(x, tuple):
        return tuple(ubin_sm(e) for e in x)
    nonfinite = _nonfinite_base_text(x)
    if nonfinite is not None:
        return nonfinite
    xi = _as_int_exact(x)
    if xi is None:
        raise ValueError("ubin expects integer values")
    uv = xi % (1 << 64) if xi < 0 else xi
    return "0b" + format(uv, "b")


def _hypot_scalar(a, b):
    aa, bb = _as_complex(a), _as_complex(b)
    return (abs(aa) ** 2 + abs(bb) ** 2) ** 0.5


def hypot_sm(a, b):
    return _broadcast2(_hypot_scalar, a, b)


def _hypot_real_scalar(a, b):
    return math.hypot(float(a), float(b))


def hypot_sm_real(a, b):
    return _broadcast2(_hypot_real_scalar, a, b)


def _log_ln_cx_scalar(x: Any) -> complex:
    c = _as_complex(x)
    if c.imag == 0.0:
        c = complex(c.real, 0.0)
    if c.real == 0.0 and c.imag == 0.0:
        return complex(-math.inf, 0.0)
    return cmath.log(c)


log_ln_cx = _map_unary_cx(_log_ln_cx_scalar)


def _log10_cx_scalar(x: Any) -> complex:
    return cmath.log10(_as_complex(x))


log10_sm_cx = _map_unary_cx(_log10_cx_scalar)


def _log_cx_scalar(a, b):
    return cmath.log(_as_complex(a)) / cmath.log(_as_complex(b))


def log_sm_cx(a, b=None):
    if b is None:
        raise TypeError("log needs 2 args")
    return _broadcast2(_log_cx_scalar, a, b)


def pow_sm_cx(b, e):
    if isinstance(b, tuple) or isinstance(e, tuple):
        return _broadcast2(pow_sm_cx, b, e)
    br = _as_complex(b)
    er = _as_complex(e)
    if abs(br.imag) < 1e-15 and abs(er.imag) < 1e-15:
        r = pow_sm(br.real, er.real)
        if isinstance(r, (int, float)) and math.isfinite(float(r)):
            return complex(float(r), 0.0)
    return br**er


def _int_sm_cx_scalar(z: Any) -> complex:
    c = _as_complex(z)
    return complex(math.trunc(c.real), math.trunc(c.imag))


int_sm_cx = _map_unary_cx(_int_sm_cx_scalar)


def frac_sm_cx(z):
    c = _as_complex(z)
    return complex(c.real - math.trunc(c.real), c.imag - math.trunc(c.imag))


def round_sm_cx(z):
    c = _as_complex(z)
    return complex(round_half_away_from_zero(c.real), round_half_away_from_zero(c.imag))


def _floor_real_scalar(x: Any) -> Any:
    if isinstance(x, bool):
        return float(int(x))
    if isinstance(x, int):
        return x
    xf = float(x)
    if not math.isfinite(xf):
        return xf
    return float(math.floor(xf))


def _ceil_real_scalar(x: Any) -> Any:
    if isinstance(x, bool):
        return float(int(x))
    if isinstance(x, int):
        return x
    xf = float(x)
    if not math.isfinite(xf):
        return xf
    return float(math.ceil(xf))


def floor_sm_cx(z):
    c = _as_complex(z)
    return complex(_floor_real_scalar(c.real), _floor_real_scalar(c.imag))


def ceil_sm_cx(z):
    c = _as_complex(z)
    return complex(_ceil_real_scalar(c.real), _ceil_real_scalar(c.imag))


def real_sm(z):
    if isinstance(z, ExactComplex):
        return z.real
    c = _as_complex(z)
    return c.real


def imag_sm(z):
    if isinstance(z, ExactComplex):
        return z.imag
    return _as_complex(z).imag


def conj_sm(z):
    if isinstance(z, ExactComplex):
        return z.conjugate()
    return _as_complex(z).conjugate()


def phase_sm(z):
    return cmath.phase(_as_complex(z))


def polar_sm(z):
    c = _as_complex(z)
    return (abs(c), cmath.phase(c))


def cart_sm(z, ang=None):
    if ang is not None:
        if isinstance(z, tuple) or isinstance(ang, tuple):
            return _broadcast2(lambda a, b: cmath.rect(float(a), float(b)), z, ang)
        return cmath.rect(float(z), float(ang))
    if isinstance(z, tuple):
        flat = _flatten_value(z)
        if len(flat) == 2:
            return cmath.rect(float(flat[0]), float(flat[1]))
        if len(flat) == 1:
            return cart_sm(flat[0])
        raise ValueError("incompatible operands")
    if isinstance(z, ExactComplex):
        return z
    if isinstance(z, complex):
        return z
    return float(z)


def _atan2_scalar(y: Any, x: Any) -> float:
    if isinstance(y, (complex, ExactComplex)) or isinstance(x, (complex, ExactComplex)):
        raise ValueError("incompatible operands")
    return math.atan2(float(y), float(x))


def atan2_sm(y: Any, x: Any) -> Any:
    return _broadcast2(_atan2_scalar, y, x)


def atan2_sm_cx(y: Any, x: Any) -> Any:
    return atan2_sm(y, x)


def deg_sm_cx(*args: Any) -> Any:
    return deg_sm(*args)


def rad_sm_cx(*args: Any) -> Any:
    return rad_sm(*args)


def _is_nan_number(v: Any) -> bool:
    return isinstance(v, float) and math.isnan(v)


def _clamp_scalar(v: Any, lo: Any, hi: Any) -> Any:
    """Mirror ClampScalarValue: NaN is an absent bound, and a NaN value sits below a real lower bound."""
    if _is_nan_number(v):
        if not _is_nan_number(lo):
            return lo
        return v
    vf = float(v)
    if not _is_nan_number(lo) and vf <= float(lo):
        return lo
    if not _is_nan_number(hi) and vf > float(hi):
        return hi
    if _is_nan_number(hi) and not _is_nan_number(lo) and vf > float(lo):
        return math.nan
    return v


def clamp_sm(x, lo, hi):
    if isinstance(lo, tuple) or isinstance(hi, tuple):
        raise ValueError("clamp bounds must be scalar")
    if isinstance(x, tuple):
        return tuple(_clamp_scalar(item, lo, hi) for item in _flatten_value(x))
    return _clamp_scalar(x, lo, hi)


def _gcd_scalar(a, b):
    return math.gcd(int(a), int(b))


def gcd_sm(a, b):
    return _broadcast2(_gcd_scalar, a, b)


def _lcm_scalar(a, b):
    return math.lcm(int(a), int(b))


def lcm_sm(a, b):
    return _broadcast2(_lcm_scalar, a, b)


def median_sm(*args):
    return statistics.median(_flatten_args(*args))


def _finite_stats(*args):
    flat = _flatten_args(*args)
    for item in flat:
        if isinstance(item, float) and not math.isfinite(item):
            return None
    return flat


def variance_sm(*args):
    flat = _finite_stats(*args)
    if flat is None:
        return math.nan
    return statistics.pvariance(flat)


def stddev_sm(*args):
    flat = _finite_stats(*args)
    if flat is None:
        return math.nan
    return statistics.pstdev(flat)


def _minmax_sm(want_min: bool, *args: Any) -> Any:
    """NaN is absent: kept only while every value seen so far is NaN."""
    flat = _flatten_args(*args)
    if not flat:
        raise ValueError("empty")
    winner: Any = None
    winner_v = 0.0
    has_winner = False
    has_non_nan = False
    for item in flat:
        if _is_nan_number(item):
            if not has_non_nan and not has_winner:
                winner = item
                has_winner = True
            continue
        v = item.ms if isinstance(item, Duration) else float(item)
        replace = (not has_winner) or _is_nan_number(winner)
        if not replace:
            replace = (v < winner_v) if want_min else (v > winner_v)
        if replace:
            winner = item
            winner_v = v
            has_winner = True
            has_non_nan = True
    return winner


def min_sm(*args):
    return _minmax_sm(True, *args)


def max_sm(*args):
    return _minmax_sm(False, *args)


def sum_sm(*args):
    return sum(_flatten_args(*args))


def random_sm(lo, hi):
    return hi


def rand_sm():
    return 0.0


def round_sm(x):
    if isinstance(x, tuple):
        return tuple(round_sm(e) for e in _flatten_value(x))
    if isinstance(x, int) and not isinstance(x, bool):
        return x
    xf = float(x)
    if not math.isfinite(xf) or abs(xf) >= 1e16 or xf.is_integer():
        return xf
    return round_half_away_from_zero(xf)


class _FactorPower:
    """One ``p`` or ``p**e`` term. ``numeric()`` is what ``prod(factorint(...))`` multiplies."""

    __slots__ = ("base", "exp", "sign")

    def __init__(self, base: int, exp: int, sign: int = 1) -> None:
        self.base = int(base)
        self.exp = int(exp)
        self.sign = 1 if sign >= 0 else -1

    def numeric(self) -> int:
        if self.base == 0:
            return 0
        return self.sign * (self.base**self.exp)

    def __str__(self) -> str:
        if self.base == 0:
            return "0"
        if self.base == 1 and self.exp == 1:
            return "-1" if self.sign < 0 else "1"
        body = str(self.base) if self.exp == 1 else f"{self.base}**{self.exp}"
        return f"-{body}" if self.sign < 0 else body

    def __mul__(self, other: object) -> int:
        other_n = other.numeric() if isinstance(other, _FactorPower) else int(other)  # type: ignore[arg-type]
        return self.numeric() * other_n

    def __rmul__(self, other: object) -> int:
        return self.__mul__(other)


def _miller_rabin(n: int) -> bool:
    """Deterministic for every integer below 2**64."""
    if n < 2:
        return False
    small = (2, 3, 5, 7, 11, 13, 17, 19, 23, 29, 31, 37)
    if n in small:
        return True
    for prime in small:
        if n % prime == 0:
            return False
    d = n - 1
    s = 0
    while d % 2 == 0:
        d //= 2
        s += 1
    for base in (2, 325, 9375, 28178, 450775, 9780504, 1795265022):
        if base % n == 0:
            continue
        x = pow(base, d, n)
        if x == 1 or x == n - 1:
            continue
        witness = True
        for _ in range(s - 1):
            x = (x * x) % n
            if x == n - 1:
                witness = False
                break
        if witness:
            return False
    return True


def _pollard_rho(n: int) -> int:
    if n % 2 == 0:
        return 2
    if n % 3 == 0:
        return 3
    for const in (1, 2, 3, 5, 7, 11):
        x = 2
        y = 2
        d = 1
        while d == 1:
            x = (x * x + const) % n
            y = (y * y + const) % n
            y = (y * y + const) % n
            d = math.gcd(abs(x - y), n)
        if d != n:
            return d
    raise ValueError("factor")


def _prime_exponents(n: int) -> dict[int, int]:
    acc: dict[int, int] = {}

    def absorb(m: int) -> None:
        if m <= 1:
            return
        if _miller_rabin(m):
            acc[m] = acc.get(m, 0) + 1
            return
        divisor = _pollard_rho(m)
        absorb(divisor)
        absorb(m // divisor)

    m = n
    p = 2
    while p * p <= m and p <= 10007:
        while m % p == 0:
            acc[p] = acc.get(p, 0) + 1
            m //= p
        p = 3 if p == 2 else p + 2
    absorb(m)
    return acc


def _as_factor_int(n: Any) -> int:
    if isinstance(n, bool) or isinstance(n, tuple) or isinstance(n, (complex, ExactComplex, Duration)):
        raise ValueError("factorint() expects integer values")
    if isinstance(n, int):
        if n > _UINT64_MAX or n < _INT64_MIN:
            raise ValueError("factorint() expects integer values")
        return n
    if isinstance(n, float) and math.isfinite(n) and n.is_integer():
        iv = int(n)
        if _INT64_MIN <= iv <= _UINT64_MAX and float(iv) == n:
            return iv
    raise ValueError("factorint() expects integer values")


def factorint_sm(n: Any) -> tuple:
    if isinstance(n, tuple):
        raise ValueError("factorint() expects scalar values")
    iv = _as_factor_int(n)
    if iv == 0:
        return (_FactorPower(0, 1),)
    if iv == 1:
        return (_FactorPower(1, 1, 1),)
    if iv == -1:
        return (_FactorPower(1, 1, -1),)
    negative = iv < 0
    exponents = _prime_exponents(abs(iv))
    terms = []
    for index, prime in enumerate(sorted(exponents)):
        sign = -1 if negative and index == 0 else 1
        terms.append(_FactorPower(prime, exponents[prime], sign))
    return tuple(terms)


def frac_sm(x: Any) -> Any:
    if isinstance(x, tuple):
        return tuple(frac_sm(e) for e in _flatten_value(x))
    if isinstance(x, float) and not math.isfinite(x):
        return math.nan
    xf = float(x)
    return xf - math.trunc(xf)


def ncr_sm(n, k):
    return _broadcast2(lambda a, b: math.comb(int(a), int(b)), n, k)


def npr_sm(n, k):
    return _broadcast2(lambda a, b: math.perm(int(a), int(b)), n, k)


def _require_int(v: Any, fn: str) -> int:
    if isinstance(v, bool) or isinstance(v, (tuple, complex, ExactComplex, Duration)):
        raise ValueError(f"{fn} expects integer values")
    if isinstance(v, int):
        return v
    if isinstance(v, float) and math.isfinite(v) and v.is_integer():
        return int(v)
    raise ValueError(f"{fn} expects integer values")


_MAX_PRODUCED = 1000


def flat_sm(*args: Any) -> tuple:
    if not args:
        raise ValueError("flat() expects at least 1 argument")
    return tuple(_flatten_args(*args))


def len_sm(*args: Any) -> int:
    if not args:
        return 0
    return len(_flatten_args(*args))


def range_sm(start: Any, stop: Any, step: Any = 1) -> tuple:
    s = _require_int(start, "range()")
    t = _require_int(stop, "range()")
    p = _require_int(step, "range()")
    if p == 0:
        raise ValueError("range() step must not be zero")
    if (p > 0 and s >= t) or (p < 0 and s <= t):
        raise ValueError("range() produces no values")
    count = len(range(s, t, p))
    if count > _MAX_PRODUCED:
        raise ValueError("range() produces too many values")
    return tuple(range(s, t, p))


def repeat_sm(x: Any, n: Any) -> tuple:
    count = _require_int(n, "repeat()")
    if count <= 0:
        raise ValueError("repeat() expects a positive integer")
    items = list(_flatten_value(x)) if isinstance(x, tuple) else [x]
    if len(items) * count > _MAX_PRODUCED:
        raise ValueError("repeat() produces too many values")
    return tuple(items * count)


def reverse_sm(*args: Any) -> tuple:
    if not args:
        raise ValueError("reverse() expects at least 1 argument")
    return tuple(reversed(_flatten_args(*args)))


def _sort_rank(v: Any) -> tuple:
    if isinstance(v, (complex, ExactComplex, Duration)):
        raise TypeError("sort")
    if _is_nan_number(v):
        return (0, 0.0)
    if isinstance(v, float) and math.isinf(v):
        return (1, v)
    if isinstance(v, bool):
        return (1, float(int(v)))
    if isinstance(v, (int, float)):
        return (1, float(v))
    raise TypeError("sort")


def sort_sm(*args: Any) -> tuple:
    if not args:
        raise ValueError("sort() expects at least 1 argument")
    return tuple(sorted(_flatten_args(*args), key=_sort_rank))


def _has_nan_component(v: Any) -> bool:
    if _is_nan_number(v):
        return True
    if isinstance(v, complex):
        return math.isnan(v.real) or math.isnan(v.imag)
    return False


def _same_unique(a: Any, b: Any) -> bool:
    if _has_nan_component(a) and _has_nan_component(b):
        return True
    return a == b


def unique_sm(*args: Any) -> tuple:
    if not args:
        raise ValueError("unique() expects at least 1 argument")
    out: list[Any] = []
    for item in _flatten_args(*args):
        if any(_same_unique(item, prev) for prev in out):
            continue
        out.append(item)
    return tuple(out)


class _UserFunction:
    """A ``name(params)=body`` binding. Printing it yields ``defined name(params)``."""

    def __init__(self, name: str, params: tuple[str, ...], impl) -> None:
        self.name = name
        self.params = params
        self._impl = impl

    def __call__(self, *args):
        return self._impl(*args)


class _Unpack:
    __slots__ = ("values",)

    def __init__(self, values: tuple) -> None:
        self.values = tuple(values)


def unpack_sm(*args: Any) -> _Unpack:
    if not args:
        raise ValueError("unpack() expects at least 1 argument")
    return _Unpack(tuple(_flatten_args(*args)))


def expand_call_sm(fn, *args, **kwargs):
    flat = []
    for arg in args:
        if isinstance(arg, _Unpack):
            flat.extend(arg.values)
        else:
            flat.append(arg)
    return fn(*flat, **kwargs)


def _sortby_key(k: Any) -> tuple:
    if isinstance(k, tuple):
        return tuple(_sortby_key(item) for item in k)
    if _is_nan_number(k):
        return (0, 0.0)
    if isinstance(k, float) and math.isinf(k):
        return (1, k)
    if isinstance(k, bool):
        return (1, float(int(k)))
    if isinstance(k, (int, float)):
        return (1, float(k))
    if isinstance(k, (complex, ExactComplex)):
        c = _as_complex(k)
        return (2, c.real, c.imag)
    if isinstance(k, Duration):
        return (1, k.ms / _SECOND_MS)
    return (4, str(k))


def sortby_sm(arr: Any, keyfn: Any) -> tuple:
    if not callable(keyfn):
        raise TypeError("sortby expects a function")
    items = list(_flatten_value(arr) if isinstance(arr, tuple) else (arr,))
    items.sort(key=lambda el: _sortby_key(keyfn(el)))
    return tuple(items)


def filter_sm(arr: Any, pred: Any) -> tuple:
    if not callable(pred):
        raise TypeError("filter expects a function")
    items = _flatten_value(arr) if isinstance(arr, tuple) else (arr,)
    return tuple(el for el in items if dsl_truth(pred(el)))


def build_ns(ans: Any) -> dict:
    return {
        "pi": math.pi,
        "e": math.e,
        "inf": math.inf,
        "nan": math.nan,
        "ans": ans,
        "dsl_cmp": dsl_cmp,
        "dsl_lnot": dsl_lnot,
        "dsl_and": dsl_and,
        "dsl_or": dsl_or,
        "dsl_bitinv": dsl_bitinv,
        "hex_val_sm": hex_val_sm,
        "base_val_sm": base_val_sm,
        "div_sm_real": div_sm_real,
        "add_sm": add_sm,
        "sub_sm": sub_sm,
        "mul_sm_real": mul_sm_real,
        "lshift_sm": lshift_sm,
        "rshift_sm": rshift_sm,
        "bitand_sm": bitand_sm,
        "bitor_sm": bitor_sm,
        "bitxor_sm": bitxor_sm,
        "real_sm": real_sm,
        "imag_sm": imag_sm,
        "conj_sm": conj_sm,
        "phase_sm": phase_sm,
        "polar_sm": polar_sm,
        "int_sm": int_sm,
        "log_sm": log_sm,
        "log_ln": _map_unary_real(log_ln),
        "log10_sm": _map_unary_real(log10_sm),
        "pow_sm": pow_sm,
        "pow_sm_cx": pow_sm_cx,
        "sqr_sm": sqr_sm,
        "prod_sm": prod_sm,
        "mean_sm": mean_sm,
        "avg_sm": avg_sm,
        "fact_sm": fact_sm,
        "mod_sm": mod_sm,
        "deg_sm": deg_sm,
        "rad_sm": rad_sm,
        "sign_sm": sign_sm,
        "clamp_sm": clamp_sm,
        "gcd_sm": gcd_sm,
        "lcm_sm": lcm_sm,
        "median_sm": median_sm,
        "variance_sm": variance_sm,
        "stddev_sm": stddev_sm,
        "min_sm": min_sm,
        "max_sm": max_sm,
        "sum_sm": sum_sm,
        "random_sm": random_sm,
        "rand_sm": rand_sm,
        "sin": _map_unary_real(sin_sm),
        "cos": _map_unary_real(cos_sm),
        "tan": _map_unary_real(tan_sm),
        "asin": _map_unary_real(_domain_nan(math.asin)),
        "acos": _map_unary_real(_domain_nan(math.acos)),
        "atan": _map_unary_real(math.atan),
        "sinh": _map_unary_real(math.sinh),
        "cosh": _map_unary_real(math.cosh),
        "tanh": _map_unary_real(math.tanh),
        "asinh": _map_unary_real(math.asinh),
        "acosh": _map_unary_real(_domain_nan(math.acosh)),
        "atanh": atanh_sm,
        "exp": _map_unary_real(math.exp),
        "sqrt": sqrt_sm_real,
        "sqrt_sm_real": sqrt_sm_real,
        "hex_sm": hex_sm,
        "bin_sm": bin_sm,
        "oct_sm": oct_sm,
        "uhex_sm": uhex_sm,
        "uoct_sm": uoct_sm,
        "ubin_sm": ubin_sm,
        "abs": _map_unary_real(abs),
        "floor": _map_unary_real(_floor_real_scalar),
        "ceil": _map_unary_real(_ceil_real_scalar),
        "trunc": math.trunc,
        "round_sm": round_sm,
        "atan2": atan2_sm,
        "hypot": hypot_sm_real,
        "ratio_sm": ratio_sm,
        "frac_sm": frac_sm,
        "ncr_sm": ncr_sm,
        "npr_sm": npr_sm,
        "factorint_sm": factorint_sm,
        "cart_sm": cart_sm,
        "flat_sm": flat_sm,
        "len_sm": len_sm,
        "range_sm": range_sm,
        "repeat_sm": repeat_sm,
        "reverse_sm": reverse_sm,
        "sort_sm": sort_sm,
        "unique_sm": unique_sm,
        "unpack_sm": unpack_sm,
        "filter_sm": filter_sm,
        "sortby_sm": sortby_sm,
        "expand_call_sm": expand_call_sm,
        "_dur_lit_compact": _dur_lit_compact,
        "dur_parse_colon_literal": dur_parse_colon_literal,
        "mul_sm_cx": mul_sm_cx,
        "div_sm_cx": div_sm_cx,
        "_empty_tuple": (),
    }


def build_ns_complex(ans: Any) -> dict:
    """Namespace for complex-option reference (``cmath`` + Smart-Math component rules)."""
    ns = build_ns(ans)
    ns.update(
        {
            "log_sm": log_sm_cx,
            "log_ln": log_ln_cx,
            "log10_sm": log10_sm_cx,
            "mean_sm": mean_sm_cx,
            "avg_sm": avg_sm_cx,
            "pow_sm": pow_sm_cx,
            "sqrt_sm": sqrt_sm,
            "hypot_sm": hypot_sm,
            "sign_sm": sign_sm,
            "complex_not_sm": complex_not_sm,
            "int_sm_cx": int_sm_cx,
            "frac_sm_cx": frac_sm_cx,
            "round_sm_cx": round_sm_cx,
            "floor_sm_cx": floor_sm_cx,
            "ceil_sm_cx": ceil_sm_cx,
            "real_sm": real_sm,
            "imag_sm": imag_sm,
            "conj_sm": conj_sm,
            "phase_sm": phase_sm,
            "polar_sm": polar_sm,
            "cart_sm": cart_sm,
            "exact_int_complex": exact_int_complex,
            "sin": _map_unary_cx(cmath.sin),
            "cos": _map_unary_cx(cmath.cos),
            "tan": _map_unary_cx(cmath.tan),
            "asin": _map_unary_cx(cmath.asin),
            "acos": _map_unary_cx(cmath.acos),
            "atan": _map_unary_cx(cmath.atan),
            "sinh": _map_unary_cx(cmath.sinh),
            "cosh": _map_unary_cx(cmath.cosh),
            "tanh": _map_unary_cx(cmath.tanh),
            "asinh": _map_unary_cx(cmath.asinh),
            "acosh": _map_unary_cx(cmath.acosh),
            "atanh": _map_unary_cx(cmath.atanh),
            "exp": _map_unary_cx(cmath.exp),
            "abs": _map_unary_cx(abs),
            "hex_sm": hex_sm,
            "bin_sm": bin_sm,
            "oct_sm": oct_sm,
            "uhex_sm": uhex_sm,
            "uoct_sm": uoct_sm,
            "ubin_sm": ubin_sm,
            "sqrt": _map_unary_cx(sqrt_sm),
            "hypot": hypot_sm,
            "atan2": atan2_sm_cx,
            "deg_sm": deg_sm_cx,
            "rad_sm": rad_sm_cx,
        }
    )
    return ns


def build_ns_time(ans: Any) -> dict:
    ns = build_ns(ans)
    ns.update(
        {
            "millisecond": Duration(1),
            "second": Duration(_SECOND_MS),
            "minute": Duration(_MINUTE_MS),
            "hour": Duration(_HOUR_MS),
            "day": Duration(_DAY_MS),
            "dur_parse_colon_literal": dur_parse_colon_literal,
            "_dur_lit_compact": _dur_lit_compact,
            "sum_dur_sm": sum_dur_sm,
            "milliseconds": milliseconds_sm,
            "seconds": seconds_sm,
            "minutes": minutes_sm,
            "hours": hours_sm,
            "days": days_sm,
        }
    )
    return ns


def ref_eval(
    expr: str, ns: dict, *, mode: str = "real", expected_hint: str = ""
) -> Any:
    e = _strip_line_comment(expr).strip()
    if not e:
        raise ValueError("empty")
    e = _replace_int_literals(e)
    e = _normalize_nonfinite_names(e)
    p = _postfix_percent(e)
    if p is None:
        raise ValueError("postfix percent")
    e = p
    if mode == "complex":
        e = _preprocess_complex_surface(e)
        # Tight-``i`` suffixes leave ``0x…`` inside ``(0x…*1j)``; normalize again.
        e = _replace_int_literals(e)
        e = _lower_exact_int_complex_literals(e)
        e = _preprocess_complex_dsl(e)
        e = _preprocess_word_not(e)
    elif mode == "time":
        e = _preprocess_time_expression(e)
        e = _preprocess_unary_bang_real(e)
        e = _preprocess_word_not(e)
    else:
        e = _preprocess_real_dsl(e)
        e = _preprocess_word_not(e)
        e = _replace_compact_time_literals(e)
    e = _rewrite_formatters_for_value(e)
    e = _implicit_mul(e)
    e = _rewrite_lambdas(e)
    e = _rewrite_empty_tuples(e)
    t = _translate_functions(e, mode=mode)
    if t is None:
        raise ValueError("translate")
    e = t
    if hasattr(ast, "unparse"):
        try:
            e = _broadcast_tuple_expr_string(e)
        except Exception:
            pass
    def _lambda_sm(names, body):
        name_list = [part.strip() for part in str(names).split(",") if part.strip()]

        def fn(*args, _names=tuple(name_list), _body=body):
            if len(args) != len(_names):
                raise ValueError(f"expects {len(_names)} argument(s)")
            local = dict(ns)
            for name, value in zip(_names, args):
                local[name] = value
            return ref_eval(_body, local, mode=mode)

        return fn

    ns["lambda_sm"] = _lambda_sm
    if e.startswith("(") and e.endswith(")") and e.count("(") == 1:
        try:
            return _flatten_value(_eval_tuple_literal(e, ns, mode=mode))
        except Exception:
            pass
    node = ast.parse(e, mode="eval")
    if mode in ("real", "time", "complex"):
        body = _SmokeAstTransform(mode).visit(node.body)
        ast.fix_missing_locations(body)
        node = ast.Expression(body=body)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", SyntaxWarning)
        result = eval(compile(node, "<smoke>", "eval"), {"__builtins__": {}}, ns)
    if isinstance(result, _Unpack):
        result = result.values[0] if len(result.values) == 1 else tuple(result.values)
    return _flatten_value(result)


def ref_format(v: Any, *, mode: str = "real") -> str:
    if isinstance(v, _UserFunction):
        return f"defined {v.name}({','.join(v.params)})"
    if isinstance(v, RatioValue):
        return v.text
    if isinstance(v, Duration):
        return duration_format_ms(v.ms)
    if isinstance(v, str):
        return v
    if isinstance(v, _Unpack):
        return ref_format(tuple(v.values), mode=mode)
    if isinstance(v, _FactorPower):
        return str(v)
    if isinstance(v, ExactComplex):
        return _format_exact_complex_smoke(v)
    if isinstance(v, complex):
        return _format_complex_smoke(v)
    if isinstance(v, bool):
        return "1" if v else "0"
    if isinstance(v, float):
        return _format_float_scalar_smoke(v)
    if isinstance(v, tuple):
        inner = ", ".join(ref_format(x, mode=mode) for x in v)
        return f"({inner})".replace(", ", ",").replace("( ", "(")
    return str(v)


_TRAILING_FORMATTER_STMT_RE = re.compile(
    r"(hex|oct|bin|deg|rad|uhex|uoct|ubin)\s*(\(\s*\))?",
    flags=re.I,
)


def _rewrite_trailing_formatter_stmt(part: str) -> str:
    """Bare ``hex`` / ``hex()`` after a previous statement means ``hex(ans)``."""
    matched = _TRAILING_FORMATTER_STMT_RE.fullmatch(part.strip())
    if matched is None:
        return part
    return f"{matched.group(1).lower()}(ans)"


_RE_USER_FN = re.compile(
    r"^([A-Za-z_][A-Za-z0-9_]*)\s*\(\s*"
    r"([A-Za-z_][A-Za-z0-9_]*(?:\s*,\s*[A-Za-z_][A-Za-z0-9_]*)*)?"
    r"\s*\)\s*=(?!=)\s*(.*)$",
    re.S,
)


_CALL_NS: list = [None]


def _bind_user_function(
    ns: dict, name: str, params: list[str], body: str, mode: str, expected_hint: str
) -> None:
    """Call-time binding: a later ``p(x)=...`` is visible inside an earlier ``f``."""

    prior = ns.get(name)
    if prior is not None and not isinstance(prior, _UserFunction):
        raise ValueError("reserved name")

    def fn(*args, _params=tuple(params), _body=body):
        if len(args) != len(_params):
            raise ValueError(f"expects {len(_params)} argument(s)")
        live = _CALL_NS[0] if _CALL_NS[0] is not None else ns
        local = dict(live)
        for param, value in zip(_params, args):
            local[param] = value
        return ref_eval(_body, local, mode=mode, expected_hint=expected_hint)

    ns[name] = _UserFunction(name, tuple(params), fn)


def _carried_functions(ns: dict) -> dict[str, _UserFunction]:
    return {key: value for key, value in ns.items() if isinstance(value, _UserFunction)}


def try_reference(
    expr_chain: str,
    *,
    mode: str = "real",
    expected_hint: str = "",
    ans_seed: Any = None,
    fn_box: Optional[dict] = None,
) -> tuple[Optional[Any], Optional[str], Any]:
    """Evaluate semicolon chain; last statement value is result.

    The third value is ``ans`` after the last statement that succeeded. A later
    failing statement does not clear it.
    """
    parts = _split_statements(expr_chain)
    if mode == "time":
        ns: dict = build_ns_time(ans_seed)
    elif mode == "complex":
        ns = build_ns_complex(ans_seed)
    else:
        ns = build_ns(ans_seed)
    if fn_box:
        ns.update(fn_box)
    ns["ans"] = ans_seed
    if ans_seed is not None:
        ns["_"] = ans_seed
    _CALL_NS[0] = ns
    last = ans_seed
    try:
        for part in parts:
            part = _rewrite_trailing_formatter_stmt(part.strip())
            user_fn = _RE_USER_FN.match(part)
            if user_fn:
                params = [
                    name.strip()
                    for name in (user_fn.group(2) or "").split(",")
                    if name.strip()
                ]
                _bind_user_function(
                    ns,
                    user_fn.group(1),
                    params,
                    user_fn.group(3).strip(),
                    mode,
                    expected_hint,
                )
                last = ns[user_fn.group(1)]
                ns["ans"] = last
                continue
            if "=" in part and not any(
                part.startswith(pref)
                for pref in ("atan2", "pow", "log", "mod", "gcd", "lcm", "hypot", "random")
            ):
                if re.match(r"^[A-Za-z_][A-Za-z0-9_]*\s*=(?!=)", part):
                    name, rhs = part.split("=", 1)
                    name = name.strip()
                    rhs = rhs.strip()
                    last = ref_eval(rhs, ns, mode=mode, expected_hint=expected_hint)
                    ns[name] = last
                    ns["ans"] = last
                    continue
            last = ref_eval(part, ns, mode=mode, expected_hint=expected_hint)
            ns["ans"] = last
            if not isinstance(last, _UserFunction):
                ns["_"] = last
        if fn_box is not None:
            fn_box.clear()
            fn_box.update(_carried_functions(ns))
        return last, None, last
    except Exception as ex:
        if fn_box is not None:
            fn_box.clear()
            fn_box.update(_carried_functions(ns))
        return None, str(ex), ns.get("ans", ans_seed)


def _last_statement(expr: str) -> str:
    depth = 0
    start = 0
    last = expr.strip()
    for i, ch in enumerate(expr):
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        elif ch == ";" and depth == 0:
            last = expr[i + 1 :].strip()
    return _strip_line_comment(last).strip()


def should_skip_case(case: dict) -> Optional[str]:
    if case.get("assert_kind"):
        return None
    if case.get("expect_no_result"):
        return "expectNoResult"
    if case.get("expected_err"):
        return "expected error path"
    if case.get("expect_error"):
        return "parser error probe (error text not comparable with Python)"
    return None


_PROGRESS_IDLE_ABORT_S = 20.0


class _ProgressWatchdog:
    """Abort the process if nothing is printed for ``idle_s`` seconds."""

    def __init__(self, idle_s: float = _PROGRESS_IDLE_ABORT_S) -> None:
        self._idle_s = idle_s
        self._last = time.monotonic()
        self._lock = threading.Lock()
        self._stop = threading.Event()

    def touch(self) -> None:
        with self._lock:
            self._last = time.monotonic()

    def say(self, msg: str = "") -> None:
        self.touch()
        print(msg, flush=True)

    def start(self) -> None:
        def run() -> None:
            while not self._stop.wait(1.0):
                with self._lock:
                    idle = time.monotonic() - self._last
                if idle >= self._idle_s:
                    print(
                        f"ERROR: no progress for {idle:.0f}s "
                        f"(limit {self._idle_s:.0f}s); aborting.",
                        flush=True,
                    )
                    os._exit(4)

        threading.Thread(target=run, daemon=True).start()

    def stop(self) -> None:
        self._stop.set()


def main() -> int:
    root = __file__.rsplit("\\", 1)[0]
    if "/" in root and "\\" not in root:
        root = __file__.rsplit("/", 1)[0]
    bas = root + "/../SmokeTest_MathParser.bas"
    band = root + "/neg_band_cases_generated.bas"
    tools_dir = root
    argv = [a for a in sys.argv[1:] if not a.startswith("--")]
    if len(argv) > 0:
        bas = argv[0]
    if len(argv) > 1:
        band = argv[1]
    if len(argv) > 2:
        tools_dir = argv[2]

    wd = _ProgressWatchdog()
    if "--no-watchdog" not in sys.argv:
        wd.start()

    wd.say("Loading smoke/band/tools cases...")
    cases = collect_all_numeric_cases(bas, band, tools_dir)
    work_total = sum(
        1
        for c in cases
        if "expected" in c or c.get("assert_kind") or c.get("expect_error")
    )
    wd.say(f"Loaded {len(cases)} rows; {work_total} comparable (expected/assert/error).")
    if "--audit-coverage" in sys.argv or "--audit-only" in sys.argv:
        audit_coverage(bas, band, tools_dir, cases)
    if "--audit-only" in sys.argv:
        wd.stop()
        return 0
    mismatches: list[str] = []
    skipped: list[str] = []
    carried_ans: Any = None
    carried_fns: dict = {}
    ok = 0
    n_complex = 0
    n_complex_opt_compared_ok = 0
    n_complex_opt_parsed = sum(
        1 for c in cases if c.get("source") in ("complex_opt", "raw_api")
    )
    n_complex_opt_skip_should = 0
    n_complex_opt_skip_eval = 0
    n_complex_opt_error_ok = 0
    n_complex_opt_error_fail = 0
    n_time = 0
    n_assert_ok = 0
    n_band = sum(1 for c in cases if c.get("source") == "magnitude_band")
    n_compared = 0

    for case in cases:
        idx = case["idx"]
        if (
            "expected" not in case
            and not case.get("expect_error")
            and not case.get("assert_kind")
        ):
            if case.get("source") == "smoke_main" and case.get("expected_err"):
                _val, err, new_ans = try_reference(
                    case["expr"],
                    mode=case.get("ref_mode")
                    or infer_ref_mode(case["expr"], ""),
                    ans_seed=carried_ans,
                    fn_box=carried_fns,
                )
                if err:
                    carried_ans = new_ans
            continue
        n_compared += 1
        if n_compared % 50 == 0 or n_compared == 1:
            wd.say(
                f"  {n_compared}/{work_total} ok={ok} "
                f"skip={len(skipped)} mismatch={len(mismatches)}"
            )
        reason = should_skip_case(case)
        if reason:
            if case.get("source") == "smoke_main" and (
                case.get("expected_err") or case.get("expect_error")
            ):
                _val, err, new_ans = try_reference(
                    case["expr"],
                    mode=case.get("ref_mode")
                    or infer_ref_mode(case["expr"], case.get("expected", "")),
                    expected_hint=case.get("expected", ""),
                    ans_seed=carried_ans,
                    fn_box=carried_fns,
                )
                if err:
                    carried_ans = new_ans
            skipped.append(f"#{idx}: {reason}")
            if case.get("source") == "complex_opt":
                n_complex_opt_skip_should += 1
            continue
        is_co = case.get("source") == "complex_opt"
        mode = case.get("ref_mode") or infer_ref_mode(
            case["expr"], case.get("expected", "")
        )
        if is_co:
            mode = "complex"

        ans_seed = carried_ans if case.get("source") == "smoke_main" else None
        fn_box = carried_fns if case.get("source") == "smoke_main" else None
        if case.get("expect_error"):
            if mode == "complex":
                n_complex += 1
            val, err, new_ans = try_reference(
                case["expr"],
                mode=mode,
                expected_hint=case.get("expected", ""),
                ans_seed=ans_seed,
                fn_box=fn_box,
            )
            if case.get("source") == "smoke_main":
                carried_ans = new_ans
            if err:
                substr = case.get("expected_err_substr")
                if substr and substr.lower() not in err.lower():
                    mismatches.append(
                        f"#{idx} expr={case['expr']!r}\n  expected err contains {substr!r}, got {err!r}\n"
                    )
                    if is_co:
                        n_complex_opt_error_fail += 1
                else:
                    ok += 1
                    if is_co:
                        n_complex_opt_error_ok += 1
            else:
                got = ref_format(val, mode=mode).strip()
                mismatches.append(
                    f"#{idx} expr={case['expr']!r}\n  expected error, ref_python={got!r}\n"
                )
                if is_co:
                    n_complex_opt_error_fail += 1
            continue

        assert_kind = case.get("assert_kind")
        exp_hint = case.get("expected", "")
        val, err, new_ans = try_reference(
            case["expr"],
            mode=mode,
            expected_hint=exp_hint,
            ans_seed=ans_seed,
            fn_box=fn_box,
        )
        if case.get("source") == "smoke_main":
            carried_ans = new_ans
        if err:
            skipped.append(f"#{idx}: eval failed ({err[:80]})")
            if is_co:
                n_complex_opt_skip_eval += 1
            continue

        if assert_kind == "abs_gt":
            thr = float(case.get("assert_threshold", 1e-6))
            try:
                mag = _real_scalar_mag(val)
            except TypeError:
                skipped.append(f"#{idx}: non-scalar magnitude assert")
                continue
            if mag >= thr:
                ok += 1
                n_assert_ok += 1
            else:
                mismatches.append(
                    f"#{idx} expr={case['expr']!r}\n  expected |result| >= {thr}, got |{mag}|\n"
                )
            continue
        if assert_kind == "abs_lt":
            thr = float(case.get("assert_threshold", 1e-9))
            try:
                mag = _real_scalar_mag(val)
            except TypeError:
                skipped.append(f"#{idx}: non-scalar magnitude assert")
                continue
            if mag < thr:
                ok += 1
                n_assert_ok += 1
            else:
                mismatches.append(
                    f"#{idx} expr={case['expr']!r}\n  expected |result| < {thr}, got |{mag}|\n"
                )
            continue

        exp = case["expected"].strip()
        if mode == "complex":
            n_complex += 1
        elif mode == "time":
            n_time += 1
        got = ref_format(val, mode=mode).strip()
        if results_match(exp, got, mode=mode, expr=case["expr"]):
            ok += 1
            if is_co:
                n_complex_opt_compared_ok += 1
            continue
        mismatches.append(
            f"#{idx} expr={case['expr']!r}\n  expected(smoke)={exp!r}\n  ref_python  ={got!r}\n"
        )

    print("=== Smoke expected vs Python reference ===")
    n_exp = sum(
        1
        for c in cases
        if "expected" in c or c.get("assert_kind") or c.get("expect_error")
    )
    print(f"Cases parsed (numeric / assert / error): {n_exp}  (magnitude-band rows: {n_band})")
    n_raw_api = sum(1 for c in cases if c.get("source") == "raw_api")
    print(
        f"Option/raw subs: parsed {n_complex_opt_parsed} rows from complex-opt + raw API "
        f"({n_complex_opt_compared_ok} numeric OK, {n_complex_opt_error_ok} error probes OK); "
        f"{n_complex_opt_skip_should} skipped (DSL/display), "
        f"{n_complex_opt_skip_eval} eval-failed in Python stub, "
        f"{n_complex_opt_error_fail} error-probe mismatches "
        f"(raw_api rows: {n_raw_api})."
    )
    print(
        f"Compared OK (real/complex/time reference, tol rules): {ok}  "
        f"(complex-mode eval rows: {n_complex}, time-mode: {n_time}, "
        f"magnitude asserts: {n_assert_ok})"
    )
    print(f"Skipped (DSL-only / non-comparable / eval failed): {len(skipped)}")
    print(f"MISMATCH vs plain Python math: {len(mismatches)}")
    print()
    if mismatches:
        print("--- MISMATCHES ---")
        for m in mismatches:
            print(m)
    if skipped and "--verbose-skip" in sys.argv:
        print("--- SKIPPED ---")
        for s in skipped:
            print(s)
    wd.stop()
    return 1 if mismatches else 0


if __name__ == "__main__":
    raise SystemExit(main())
