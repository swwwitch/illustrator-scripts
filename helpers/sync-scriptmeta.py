#!/usr/bin/env python3
"""
SCRIPTMETA（https://github.com/Yamonov/SCRIPTMETA）の同期ツール。

各 .jsx の基本情報ブロック（SCRIPT_NAME / SCRIPT_VERSION / SCRIPT_AUTHOR / SCRIPT_RELEASED）と
ヘッダーの「### 概要」「### Overview」から SCRIPTMETA ブロックを作り直し、
リポジトリ直下の SCRIPTMETA.txt（配布ページ側の SCRIPTMETA-DIST ブロック）を書き出す。

SCRIPT_VERSION を上げたり概要を書き換えたりしたら、コミット前に実行する。

    python3 helpers/sync-scriptmeta.py          # 書き換える
    python3 helpers/sync-scriptmeta.py --check  # 差分があれば一覧を出して終了コード1

Syncs SCRIPTMETA blocks in every .jsx from its basic info block and header overview,
and regenerates SCRIPTMETA.txt at the repository root. Run it before committing.
"""
import glob
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ID_PREFIX = "com.swwwitch.illustrator."
META_URL = "https://github.com/swwwitch/illustrator-scripts"
TARGET_APP = "Illustrator"
EXCLUDE_DIRS = ("/_templates/",)

# Script-ID に使えない文字（A-Z a-z 0-9 . _ - 以外）を含む SCRIPT_NAME の代わりの ID。
# 一度公開した ID は変えない / IDs for names with characters Script-ID does not allow. Never change once published
ID_OVERRIDES = {
    "ガイドをテンプレートに": "GuidesToTemplate",
    "レイヤーをテンプレートに": "LayersToTemplate",
    "丸型先端にする": "SetRoundCap",
    "シンボルに置き換え-sw": "ReplaceWithSymbol-sw",
    "AutoLeadingStep+1": "AutoLeadingStepPlus1",
    "AutoLeadingStep+10": "AutoLeadingStepPlus10",
    "段落スタイル": "ParagraphStyles",
}

BLOCK_RE = re.compile(r"/\*\nSCRIPTMETA-BEGIN\n.*?\nSCRIPTMETA-END\n\*/\n\n", re.S)
SCRIPT_ID_RE = re.compile(r"^[A-Za-z0-9._-]+$")


def read_var(text, key):
    m = re.search(r'^var\s+' + key + r'\s*=\s*"([^"]*)"', text, re.M)
    return m.group(1) if m else None


def read_section(text, heading, stop_prefixes):
    """「### 見出し」から README 案内の手前までの段落を返す"""
    start = text.find(heading)
    if start < 0:
        return ""
    lines = text[start + len(heading):].split("\n")
    out = []
    for line in lines:
        s = line.strip()
        if s.startswith("###") or s.startswith("*/") or any(s.startswith(p) for p in stop_prefixes):
            break
        out.append(line.rstrip())
    return "\n".join(out).strip("\n")


def build_block(path, text):
    name = read_var(text, "SCRIPT_NAME")
    version = read_var(text, "SCRIPT_VERSION")
    if not name or not version:
        raise ValueError("SCRIPT_NAME / SCRIPT_VERSION がありません")
    script_id = ID_PREFIX + ID_OVERRIDES.get(name, name)
    if not SCRIPT_ID_RE.match(script_id):
        raise ValueError("Script-ID に使えない文字があります（ID_OVERRIDES に追加）: " + script_id)
    version = version.lstrip("v")
    if not re.match(r"^\d+(\.\d+)*$", version):
        raise ValueError("Version の形式が不正です: " + version)

    ja = read_section(text, "### 概要", ("詳細は", "note記事"))
    en = read_section(text, "### Overview", ("See the README",))
    desc = "\n\n".join(p for p in (ja, en) if p)
    for line in desc.split("\n"):
        if line.strip() in ("Description-END", "SCRIPTMETA-END") or "*/" in line:
            raise ValueError("説明文に使えない行があります: " + line)

    lines = ["/*", "SCRIPTMETA-BEGIN", "Script-ID=" + script_id, "Name=" + name]
    author = read_var(text, "SCRIPT_AUTHOR")
    if author:
        lines.append("Author=" + author)
    lines += ["Target-App=" + TARGET_APP, "Version=" + version]
    released = read_var(text, "SCRIPT_RELEASED")
    if released:
        lines.append("Release-Date=" + released)
    lines.append("Meta-URL=" + META_URL)
    if desc:
        lines += ["Description-BEGIN", desc, "Description-END"]
    lines += ["SCRIPTMETA-END", "*/", "", ""]
    return script_id, version, "\n".join(lines)


def insert_block(text, block):
    if BLOCK_RE.search(text):
        return BLOCK_RE.sub(lambda m: block, text, count=1)
    # ヘッダーコメント（### 概要）の直前に置く。無ければ #target などの指令行の後
    head = text.find("### 概要")
    pos = text.rfind("/*", 0, head) if head >= 0 else -1
    if pos < 0:
        pos = 0
        for m in re.finditer(r"^(#.*|app\.preferences\..*)\n", text, re.M):
            if text[pos:m.start()].strip():
                break
            pos = m.end()
        block = ("\n" if pos else "") + block
    return text[:pos] + block + text[pos:]


def main():
    check = "--check" in sys.argv
    files = sorted(
        f for f in glob.glob(os.path.join(ROOT, "jsx", "**", "*.jsx"), recursive=True)
        if not any(d in f.replace(os.sep, "/") for d in EXCLUDE_DIRS)
    )
    records, seen, changed, errors = [], {}, [], []
    for path in files:
        rel = os.path.relpath(path, ROOT)
        with open(path, encoding="utf-8", newline="") as fp:
            text = fp.read()
        try:
            script_id, version, block = build_block(path, text)
        except ValueError as e:
            errors.append(rel + ": " + str(e))
            continue
        if script_id in seen:
            errors.append(rel + ": Script-ID が " + seen[script_id] + " と重複しています")
            continue
        seen[script_id] = rel
        records.append((script_id, version))
        new_text = insert_block(text, block)
        if new_text != text:
            changed.append(rel)
            if not check:
                with open(path, "w", encoding="utf-8", newline="") as fp:
                    fp.write(new_text)

    dist = ["SCRIPTMETA-DIST-BEGIN", ""]
    for script_id, version in sorted(records):
        dist += ["Script-ID=" + script_id, "Version=" + version, ""]
    dist += ["SCRIPTMETA-DIST-END", ""]
    dist_text = "\n".join(dist)
    dist_path = os.path.join(ROOT, "SCRIPTMETA.txt")
    old_dist = open(dist_path, encoding="utf-8").read() if os.path.exists(dist_path) else ""
    if dist_text != old_dist:
        changed.append("SCRIPTMETA.txt")
        if not check:
            with open(dist_path, "w", encoding="utf-8", newline="") as fp:
                fp.write(dist_text)

    for e in errors:
        print("ERROR " + e)
    for c in changed:
        print(("NEEDS SYNC " if check else "updated ") + c)
    print("%d scripts, %d changed, %d errors" % (len(records), len(changed), len(errors)))
    sys.exit(1 if errors or (check and changed) else 0)


if __name__ == "__main__":
    main()
