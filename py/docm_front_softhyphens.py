# docm_front_softhyphens.py <docm> [front-matter text file] [--apply]: carries the owner's hand-set optional (soft) hyphens from a SAVED docm into the
# front-matter text file as U+00AD (the docm is the SSOT, RT-23 / 2026-11-12). Without --apply it only reports what would change. A paragraph whose
# wording differs from the docm (soft hyphens ignored) is an error: nothing is written. Paragraphs of the text file that already hold U+00AD are
# replaced by the docm's positions. Reads the file's XML only (no Word). Run: wsl.exe --exec python3 -I py/docm_front_softhyphens.py "<docm>" [text] [--apply]
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import docm_front_verify as v

SOFT = "­"


def raw_with_hyphens(raw, plain_with_hyphens):
    """Insert U+00AD into the raw text-file text (which carries {justify}/{i}/{/i} markers) at the positions the docm text has them."""
    out, k = [], 0
    i = 0
    while i < len(raw):
        for marker in ("{justify}", "{i}", "{/i}"):
            if raw.startswith(marker, i):
                out.append(marker)
                i += len(marker)
                break
        else:
            while k < len(plain_with_hyphens) and plain_with_hyphens[k] == SOFT:
                out.append(SOFT)
                k += 1
            if k >= len(plain_with_hyphens) or plain_with_hyphens[k] != raw[i]:
                raise ValueError("wording differs near: " + raw[max(0, i - 20):i + 20])
            out.append(raw[i])
            k += 1
            i += 1
    while k < len(plain_with_hyphens) and plain_with_hyphens[k] == SOFT:
        out.append(SOFT)
        k += 1
    if k != len(plain_with_hyphens):
        raise ValueError("the docm paragraph is longer than the text-file paragraph")
    return "".join(out)


def main():
    args = [a for a in sys.argv[1:] if a != "--apply"]
    apply = "--apply" in sys.argv[1:]
    docm = args[0]
    spec = Path(args[1]) if len(args) > 1 else v.DEFAULT_TEXT
    lines = spec.read_text(encoding="utf8").split("\n")
    n_para = sum(1 for ln in lines if ln and not ln.startswith("#") and not ln.startswith("@"))
    got, _ = v.read_front(docm, n_para)
    out, p, total, changed = [], 0, 0, 0
    for ln in lines:
        if not ln or ln.startswith("#") or ln.startswith("@"):
            out.append(ln)
            continue
        style, text = ln.split("\t", 1)
        docm_text = got[p]["text"]
        p += 1
        new = raw_with_hyphens(text.replace(SOFT, ""), docm_text)
        total += new.count(SOFT)
        if new != text:
            changed += 1
        out.append(style + "\t" + new)
    print("paragraphs", p, "| soft hyphens in the docm front matter", total, "| paragraphs changed", changed)
    if apply:
        spec.write_text("\n".join(out), encoding="utf8", newline="")
        print("written", spec)
    return 0


if __name__ == "__main__":
    sys.exit(main())
