"""Inventory formualizer as installed, declared, locked and published, then compare
its public Python surface with what linexcel mentions.

Run from the repository root:

    uv run python .agents/skills/formualizer/scripts/inventory.py            # offline
    uv run python .agents/skills/formualizer/scripts/inventory.py --online   # + PyPI
    uv run python .agents/skills/formualizer/scripts/inventory.py --json
    uv run python .agents/skills/formualizer/scripts/inventory.py --compare 0.14.0

``--compare`` installs the named release in a throwaway environment
(``uv run --no-project --with formualizer==X``) and lists public names and
class members added or removed relative to the installed version. It never
changes the project lock.

The usage comparison is textual. "Unmentioned" means the member name never
appears in src/linexcel, which makes it a lead to investigate, not proof that a
capability is unused or useful. Generic names (name, trace, cancel...) can be
mentioned for unrelated reasons, so "mentioned" is not proof of use either.
"""

from __future__ import annotations

import argparse
import inspect
import json
import re
import subprocess
import sys
import tempfile
import tomllib
import urllib.request
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path

ROOT = Path.cwd()
SOURCE = ROOT / "src" / "linexcel"
ENGINE = re.compile(r"\b(formualizer|upstream|native|engine)\b", re.IGNORECASE)
WORKAROUND = re.compile(
    r"\b(bug|workaround|bridge|correct\w*|unsupported|not support\w*|limitation|"
    r"fallback|regression|until|instead|rejects?|refuses?|does not|doesn't|cannot)\b",
    re.IGNORECASE,
)


def declared_floor():
    project = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    for item in project["project"]["dependencies"]:
        if item.lower().startswith("formualizer"):
            return item
    return None


def locked_version():
    lock = ROOT / "uv.lock"
    if not lock.exists():
        return None
    data = tomllib.loads(lock.read_text(encoding="utf-8"))
    for package in data.get("package", []):
        if package.get("name") == "formualizer":
            return package.get("version")
    return None


def published(online):
    if not online:
        return None
    url = "https://pypi.org/pypi/formualizer/json"
    request = urllib.request.Request(url, headers={"User-Agent": "linexcel-skill"})
    with urllib.request.urlopen(request, timeout=20) as response:
        info = json.load(response)
    releases = [
        name
        for name, files in info["releases"].items()
        if files and not any(file.get("yanked") for file in files)
    ]

    def key(text):
        return tuple(int(part) if part.isdigit() else 0 for part in text.split("."))

    return {"latest": info["info"]["version"], "recent": sorted(releases, key=key)[-6:]}


def surface(module):
    names = sorted(name for name in vars(module) if not name.startswith("_"))
    classes = {}
    for name in names:
        value = getattr(module, name)
        if inspect.isclass(value) and value.__module__.startswith("formualizer"):
            members = sorted(m for m in vars(value) if not m.startswith("_"))
            if members:
                classes[name] = members
    return names, classes


SURFACE_CODE = """
import inspect, json, formualizer
from importlib.metadata import version
names = sorted(n for n in vars(formualizer) if not n.startswith("_"))
classes = {}
for n in names:
    v = getattr(formualizer, n)
    if inspect.isclass(v) and v.__module__.startswith("formualizer"):
        classes[n] = sorted(m for m in vars(v) if not m.startswith("_"))
report = {"version": version("formualizer"), "names": names, "classes": classes}
print(json.dumps(report))
"""


def other_surface(release):
    if not re.fullmatch(r"[0-9]+(?:\.[0-9]+)*(?:[a-z0-9.]*)", release):
        sys.exit(f"Not a release number: {release!r}")
    command = [
        "uv", "run", "-q", "--no-project", "--with", f"formualizer=={release}",
        "python", "-c", SURFACE_CODE,
    ]  # fmt: skip
    output = subprocess.run(
        command, capture_output=True, text=True, check=True, cwd=tempfile.gettempdir()
    ).stdout
    return json.loads(output.strip().splitlines()[-1])


def compare(names, classes, other):
    added = sorted(set(other["names"]) - set(names))
    removed = sorted(set(names) - set(other["names"]))
    members = {}
    for cls in sorted(set(classes) | set(other["classes"])):
        new = set(other["classes"].get(cls, [])) - set(classes.get(cls, []))
        gone = set(classes.get(cls, [])) - set(other["classes"].get(cls, []))
        if new or gone:
            members[cls] = {"added": sorted(new), "removed": sorted(gone)}
    return {
        "version": other["version"],
        "added": added,
        "removed": removed,
        "members": members,
    }


def linexcel_text():
    return {
        path.relative_to(ROOT).as_posix(): path.read_text(encoding="utf-8")
        for path in sorted(SOURCE.rglob("*.py"))
    }


def mentions(name, corpus):
    pattern = re.compile(r"(?<![\w])" + re.escape(name) + r"(?![\w])")
    return sorted(path for path, text in corpus.items() if pattern.search(text))


def workaround_markers(corpus):
    found = []
    for path, text in corpus.items():
        for number, line in enumerate(text.splitlines(), 1):
            stripped = line.strip()
            if (
                stripped.startswith("#")
                and ENGINE.search(stripped)
                and WORKAROUND.search(stripped)
            ):
                found.append(f"{path}:{number}: {stripped[:160]}")
    return found


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--online", action="store_true", help="query PyPI")
    parser.add_argument("--json", action="store_true", help="emit JSON")
    parser.add_argument(
        "--compare", metavar="VERSION", help="diff the API with another release"
    )
    args = parser.parse_args()
    try:
        import formualizer
    except ImportError:
        sys.exit("formualizer is not importable: run inside the project environment")
    try:
        installed = version("formualizer")
    except PackageNotFoundError:
        installed = None
    package_dir = Path(formualizer.__file__).parent
    names, classes = surface(formualizer)
    corpus = linexcel_text()
    report = {
        "versions": {
            "installed": installed,
            "declared": declared_floor(),
            "locked": locked_version(),
            "pypi": published(args.online),
        },
        "stubs": [p.as_posix() for p in sorted(package_dir.glob("*.pyi"))],
        "moduleNames": {
            name: mentions(name, corpus) for name in names if name != "formualizer_py"
        },
        "classMembers": {
            cls: {member: mentions(member, corpus) for member in members}
            for cls, members in classes.items()
        },
        "workaroundComments": workaround_markers(corpus),
        "comparison": compare(names, classes, other_surface(args.compare))
        if args.compare
        else None,
    }
    if args.json:
        print(json.dumps(report, indent=2))
        return
    v = report["versions"]
    print("# formualizer inventory\n")
    print(f"- installed: {v['installed']}")
    print(f"- declared:  {v['declared']}")
    print(f"- locked:    {v['locked']}")
    if v["pypi"]:
        pypi = v["pypi"]
        print(f"- PyPI:      latest {pypi['latest']} (recent {pypi['recent']})")
    else:
        print("- PyPI:      not queried (add --online)")
    print(f"- stubs:     {', '.join(report['stubs']) or 'none found'}\n")
    unmentioned = [n for n, where in report["moduleNames"].items() if not where]
    print(f"## Module names never mentioned in src/linexcel ({len(unmentioned)})\n")
    print(", ".join(unmentioned) or "none")
    print("\n## Class members never mentioned in src/linexcel\n")
    for cls, members in report["classMembers"].items():
        missing = [m for m, where in members.items() if not where]
        if missing:
            print(f"- {cls} ({len(missing)}/{len(members)}): {', '.join(missing)}")
    print(f"\n## Comments to re-verify ({len(report['workaroundComments'])})\n")
    for line in report["workaroundComments"]:
        print(f"- {line}")
    diff = report["comparison"]
    if diff:
        print(f"\n## API changes from {v['installed']} to {diff['version']}\n")
        print(f"- names added: {', '.join(diff['added']) or 'none'}")
        print(f"- names removed: {', '.join(diff['removed']) or 'none'}")
        for cls, change in diff["members"].items():
            print(
                f"- {cls}: +{', '.join(change['added']) or '-'}"
                f" / -{', '.join(change['removed']) or '-'}"
            )
    print(
        "\nLeads only: confirm each with the stubs, upstream changelog and a probe "
        "before changing linexcel."
    )


if __name__ == "__main__":
    main()
