"""Telling the user what is happening while it happens.

A large workbook takes minutes, and the two phases that take them — reading
the values the file carries, and sweeping it for formulas — are exactly the
ones a pathological file makes pathological. Before this, ``verbose`` printed
each phase's duration *after* it ended: a 50 MB file was several minutes of
silence, then a wall of timings. Silence is what makes someone kill the
process and conclude the tool hangs.

Progress goes to **stderr**, so `--json -` and `-o -` stay pipeable, and it is
drawn only when stderr is a terminal: a CI log or a redirect gets the plain
one-line-per-phase form instead of thousands of redraw escapes.

``rich`` draws it when it is installed (``pip install 'linexcel[progress]'``).
Without it nothing is lost but the bars — the same phases and the same timings
are printed as plain lines, so no behaviour depends on whether it is there.
"""

from __future__ import annotations

import sys
import time
from collections.abc import Iterator
from contextlib import contextmanager


def _rich_console():
    """A rich console on stderr, or ``None`` if rich is not installed."""
    try:
        from rich.console import Console
    except ImportError:
        return None
    return Console(stderr=True)


class Reporter:
    """Phase and per-item progress, or silence.

    The silent one is the default and costs nothing: ``analyze()`` is a
    library call, and a library that writes to a terminal it was not asked to
    write to is a library people wrap in ``contextlib.redirect_stderr``.
    """

    def __init__(self, enabled: bool = False, *, force_plain: bool = False) -> None:
        self.enabled = enabled
        self._console = None if force_plain or not enabled else _rich_console()
        # A progress display redrawing into a pipe is thousands of escape
        # sequences no one will read.
        self._live = bool(self._console) and sys.stderr.isatty()

    @contextmanager
    def phase(self, label: str, total: int | None = None) -> Iterator[Phase]:
        """One named phase, timed, optionally with a known number of steps."""
        if not self.enabled:
            yield _SILENT_PHASE
            return
        started = time.perf_counter()
        if self._live:
            from rich.progress import (
                BarColumn,
                Progress,
                ProgressColumn,
                SpinnerColumn,
                TextColumn,
                TimeElapsedColumn,
            )

            columns: list[ProgressColumn] = [
                SpinnerColumn(),
                TextColumn("[progress.description]{task.description}"),
            ]
            if total:
                columns += [BarColumn(), TextColumn("{task.completed}/{task.total}")]
            columns.append(TimeElapsedColumn())
            with Progress(*columns, console=self._console, transient=True) as bars:
                task = bars.add_task(label, total=total)
                yield _RichPhase(bars, task)
        else:
            yield _SILENT_PHASE
        self._say(f"{label}: {time.perf_counter() - started:.1f}s")

    def note(self, message: str) -> None:
        """Something worth saying that is not a phase."""
        if self.enabled:
            self._say(message)

    def _say(self, message: str) -> None:
        if self._console is not None:
            self._console.print(f"[dim]\\[linexcel][/dim] {message}", highlight=False)
        else:
            print(f"[linexcel] {message}", file=sys.stderr)


class Phase:
    """What a phase hands back: somewhere to report each item done."""

    def step(self, label: str = "", advance: int = 1) -> None: ...


class _SilentPhase(Phase):
    def step(self, label: str = "", advance: int = 1) -> None:
        return


_SILENT_PHASE = _SilentPhase()


class _RichPhase(Phase):
    def __init__(self, bars, task) -> None:
        self._bars = bars
        self._task = task

    def step(self, label: str = "", advance: int = 1) -> None:
        if label:
            self._bars.update(self._task, advance=advance, description=label)
        else:
            self._bars.update(self._task, advance=advance)
