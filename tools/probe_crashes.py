"""Probe de crashs : exécute linexcel.analyze() sur chaque fixture adversariale.

Écrit ``crash_report.md`` (par défaut à côté des fixtures) — la source de
vérité des crashs : table fixture / phase / résultat / symptôme.

Usage :
    uv run python tools/probe_crashes.py [--fixtures DIR] [--report FICHIER]
                                         [--timeout SECONDES]
"""

from __future__ import annotations

import argparse
import signal
import sys
import traceback
from dataclasses import dataclass, field
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))
DEFAULT_FIXTURES = ROOT / "tests" / "fixtures" / "adversarial"

import linexcel  # noqa: E402


class _TimeoutError(Exception):
    pass


def _alarm(signum: int, frame: object) -> None:
    raise _TimeoutError()


@dataclass
class ProbeResult:
    fixture: str
    phase: str
    outcome: str  # OK / EXCEPTION / TIMEOUT
    symptom: str
    none_values: int | None = None
    nodes: int | None = None
    extras: list[str] = field(default_factory=list)


def _last_linexcel_frame(tb: str) -> str:
    """Dernière frame du package linexcel dans la traceback = phase fautive."""
    phase = "?"
    for line in traceback.format_tb(sys.exc_info()[2]):
        if "linexcel" in line:
            # extrait le nom de fonction
            part = line.strip().rsplit(",", 1)[-1].strip()
            phase = part.removeprefix("in ")
    return phase or "?"


def _count_none_values(result: linexcel.LineageResult) -> tuple[int, int]:
    """Compte les valeurs None au niveau cellule, toutes formes de nœuds.

    Le graphe est hétérogène : ``value`` (scalaire), ``values`` (liste de
    cellules pour les plages input), ``samples`` (échantillons de groupes).
    """
    none_values = 0
    total = 0
    for node in result.graph.get("nodes", []):
        if "value" in node:
            total += 1
            none_values += node["value"] is None
        for key in ("values", "samples"):
            for cell in node.get(key) or []:
                if "value" in cell:
                    total += 1
                    none_values += cell["value"] is None
    return none_values, total


def probe_one(path: Path, timeout: int) -> ProbeResult:
    res = ProbeResult(fixture=path.name, phase="analyze", outcome="OK", symptom="")
    old_handler = signal.signal(signal.SIGALRM, _alarm)
    signal.alarm(timeout)
    try:
        result = linexcel.analyze(path)
    except _TimeoutError:
        res.outcome = "TIMEOUT"
        res.symptom = f"analyse > {timeout}s (possible boucle infinie)"
    except Exception as exc:
        res.outcome = "EXCEPTION"
        tb_text = traceback.format_exc()
        res.phase = _last_linexcel_frame(tb_text)
        message = str(exc).replace("\n", " ")[:160]
        res.symptom = f"{type(exc).__name__}: {message}"
    else:
        res.none_values, res.nodes = _count_none_values(result)
        stats = getattr(result, "stats", None)
        if stats:
            res.extras.append(f"stats={stats}")
        if res.nodes and res.none_values == res.nodes:
            res.symptom = "toutes les valeurs None (graphe empoisonné)"
        elif res.none_values:
            res.symptom = f"{res.none_values}/{res.nodes} valeurs cellule None"
        else:
            res.symptom = "analyse complète"
    finally:
        signal.alarm(0)
        signal.signal(signal.SIGALRM, old_handler)
    return res


def write_report(results: list[ProbeResult], report: Path, fixtures_dir: Path) -> None:
    lines = [
        "# crash_report — probe adversarial linexcel",
        "",
        f"Fixtures : `{fixtures_dir}`",
        f"Généré par `tools/probe_crashes.py` — {len(results)} fixtures sondées.",
        "",
        "| fixture | phase | résultat | valeurs None | symptôme |",
        "|---|---|---|---|---|",
    ]
    for r in results:
        none_cell = f"{r.none_values}/{r.nodes}" if r.none_values is not None else "—"
        lines.append(
            f"| `{r.fixture}` | {r.phase} | {r.outcome} | {none_cell} | {r.symptom} |"
        )
    lines += [
        "",
        "## Synthèse",
        "",
    ]
    crashes = [r for r in results if r.outcome != "OK"]
    poisoned = [r for r in results if r.outcome == "OK" and r.none_values]
    lines.append(f"- exceptions/timeouts : **{len(crashes)}**")
    lines.append(f"- analyses OK avec valeurs None : **{len(poisoned)}**")
    lines.append(
        f"- analyses propres : **{len(results) - len(crashes) - len(poisoned)}**"
    )
    report.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--fixtures", type=Path, default=DEFAULT_FIXTURES)
    parser.add_argument("--report", type=Path, default=None)
    parser.add_argument("--timeout", type=int, default=60)
    args = parser.parse_args()
    report = args.report or args.fixtures / "crash_report.md"

    fixtures = sorted(
        p for p in args.fixtures.iterdir() if p.suffix in {".xlsx", ".xlsm"}
    )
    if not fixtures:
        print(f"Aucune fixture dans {args.fixtures} — lancez tools/gen_fixtures.py")
        return 1

    results = []
    for path in fixtures:
        res = probe_one(path, args.timeout)
        print(f"{res.outcome:9s} {res.fixture}: {res.symptom}")
        results.append(res)
    write_report(results, report, args.fixtures)
    print(f"\nRapport écrit : {report}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
