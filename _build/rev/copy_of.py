"""Print the v2 copy (case / steps / table / so) for given topic tags.  python3 _build/rev/copy_of.py 1.3.11 2.6.4"""
import importlib.util
import sys
from pathlib import Path

REV = Path(__file__).resolve().parent
sys.path.insert(0, str(REV))
sys.path.insert(0, str(REV.parent))


def load_all():
    out = {}
    for f in sorted(REV.glob("t[1-5]_*.py")):
        spec = importlib.util.spec_from_file_location(f.stem, f)
        m = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(m)
        for t in getattr(m, "TOPICS", []):
            out[t["tag"]] = (f.name, t)
    return out


if __name__ == "__main__":
    allt = load_all()
    full = "--full" in sys.argv
    for tag in [a for a in sys.argv[1:] if not a.startswith("--")]:
        if tag not in allt:
            print("??", tag)
            continue
        fn, t = allt[tag]
        print(f"## {tag} ({fn}) {t.get('title') or ''}")
        keys = ["what", "case", "steps", "table", "so", "tryit"] if not full else list(t.keys())
        for k in keys:
            if t.get(k):
                print(f"   {k}: {t[k]}")
