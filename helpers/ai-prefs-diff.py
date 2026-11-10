#!/usr/bin/env python3
"""Illustrator の設定ファイルを新旧バージョンで突き合わせ、差分を Markdown で出す。

比べるもの
  - Adobe Illustrator 環境設定 / Adobe Illustrator Cloud Prefs（app.preferences のキー）
  - .kys（/Menus・/Tools のメニューコマンド・ツール ID）… --kys-old / --kys-new を渡したときだけ
  - 原本 AT-doc/ExtendScript.txt（AiCommandPrefLookup.jsx に埋め込む一覧）に載っているか

使い方
  python3 helpers/ai-prefs-diff.py 30 31
  python3 helpers/ai-prefs-diff.py 30 31 --out /tmp/ai31-diff.md
  python3 helpers/ai-prefs-diff.py 30 "31.1.0 Beta" --kys-old 30.kys --kys-new 31.kys

バージョンは「Adobe Illustrator <版> Settings」の <版> の部分。フォルダーのパスも渡せる。

注意
  - Illustrator の起動中は「環境設定」ファイルが消えている（終了時に書き直される）。比べる版はどちらも終了しておく。
  - 設定ファイルには、一度変更・使用したキーしか書かれない。「ファイルに無い」は廃止の根拠にならない。
  - get〜Preference は存在しないキーも既定値で書き込む。環境設定ファイルにあるキーが実在するとは限らない。
  - メニュー ID の末尾の空白は ID の一部。削らない。
"""

import argparse
import difflib
import json
import os
import re
import sys
from datetime import date

PREF_FILES = (
    ("main", "Adobe Illustrator 環境設定"),
    ("cloud", "Adobe Illustrator Cloud Prefs"),
)
FILE_LABEL = {"main": "環境設定", "cloud": "Cloud Prefs"}
DEFAULT_REF = os.path.expanduser(
    "~/sw Dropbox/takano masahiro/Dropbox-shared/AT-doc/ExtendScript.txt"
)
PREF_ROOT = os.path.expanduser("~/Library/Preferences")
DEFAULT_JSX = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "jsx", "misc", "AiCommandPrefLookup.jsx")

# 原本の行の形式（AiCommandPrefLookup.jsx の PREF_LINE_PATTERN / MENU_LINE_PATTERN と同じ）
REF_PREF_LINE = re.compile(r"^([^\t]*)\t([^\t]+)\t(Boolean|Integer|Real|String)\t([^\t]*)(?:\t([^\t]*))?$")
REF_MENU_LINE = re.compile(r"^([^\t]*)\t([^\t]*)(?:\t([^\t]*))?$")

INT_RE = re.compile(r"^-?\d+$")
REAL_RE = re.compile(r"^-?\d*\.\d+(?:e-?\d+)?$")
LONG_RE = re.compile(r"^-?\d+L$")
STR_RE = re.compile(r"^\((.*)\)$")
ARR_RE = re.compile(r"^\[ \d+$")
BLOB_RE = re.compile(r"^< \d+$")

# 他社製プラグインの ID（Astute Graphics・XtreamPath など）。印を付けるだけ
THIRD_PARTY_RE = re.compile(
    r"(astute|ASTG|XtreamPath|Xtream Path|Panel Menu Item$|^DS.* Menu Item$|NkSD|^AG |^AGCreate|^SF(Make|Release|Expand)$|"
    r"VectorScribe|^Dynamic .* Tool$|^(InkFlow|PathScribe|Path Zipper|Reposition Point|Smart Remove Brush|Extend Path|"
    r"Inflate Deflate|Super Marquee) Tool$|Collision(s)? Tool$)",
    re.IGNORECASE,
)

# 設定ではなくデータを抱えている枝（アクションの中身・ライブラリーのパス・キャッシュ・最近使ったフォント）。
# 既定では一覧から外し、件数だけ出す（--include-data で含める）
DATA_KEY_RE = re.compile(
    r"^plugin/(Action/SavedSets|[^/]* Library: |Swatch Library_|AIProxy_CacheFile_MetaData|RecentFont)(/|$)"
    r"|/(Applications|Users|Volumes)/",
)

# 中身ではなく作業の痕跡に当たりそうなキー（一覧から外すのではなく印を付けるだけ）
NOISE_RE = re.compile(
    r"(recent|mru|lastused|last ?used|timestamp|begintime|time$|count$|shown|firsttime|onboarding|"
    r"coach|seq|location|position|^zz|window|palette.*(pos|size)|\.acb|\.cfg)",
    re.IGNORECASE,
)


# ---------------------------------------------------------------- 共通


def read_cr_text(path):
    """CR 改行の設定ファイルを行に分けて返す。"""
    with open(path, "rb") as f:
        data = f.read()
    text = data.decode("utf-8", errors="surrogateescape")
    return text.replace("\r\n", "\n").replace("\r", "\n").split("\n")


def decode_key(raw):
    """設定ファイル・.kys のキー表記を文字列に戻す。

    `\\ ` は空白、`\\ooo` は8進（上位ビットが立つことがあるので下位8ビットを取る）。
    バイト列は UTF-8、だめなら Shift_JIS（cp932）で読む。
    """
    out = bytearray()
    i = 0
    raw_bytes = raw.encode("utf-8", errors="surrogateescape")
    while i < len(raw_bytes):
        c = raw_bytes[i]
        if c == 0x5C and i + 1 < len(raw_bytes):  # backslash
            m = re.match(rb"[0-7]{3}", raw_bytes[i + 1 : i + 4])
            if m:
                out.append(int(m.group(0), 8) & 0xFF)
                i += 4
                continue
            out.append(raw_bytes[i + 1])
            i += 2
            continue
        out.append(c)
        i += 1
    for enc in ("utf-8", "cp932"):
        try:
            return bytes(out).decode(enc)
        except UnicodeDecodeError:
            pass
    return bytes(out).decode("latin-1")


KEY_LINE_RE = re.compile(r"^/((?:\\.|[^ \\])*)(?: (.*))?$")


def split_key_line(line):
    m = KEY_LINE_RE.match(line)
    if not m:
        return None, None
    return decode_key(m.group(1)), (m.group(2) or "")


# ---------------------------------------------------------------- 設定ファイル


def classify_value(value):
    """値の表記から (型, 表示用の値) を返す。型は ExtendScript.txt の型名に寄せる。"""
    if INT_RE.match(value):
        return "Integer", value
    if REAL_RE.match(value):
        return "Real", value
    m = STR_RE.match(value)
    if m:
        return "String", m.group(1)
    if LONG_RE.match(value):
        return "Long", value
    return "?", value


def parse_pref_file(path):
    """設定ファイルを読み、{キーのパス: (型, 値)} を返す。入れ子は「/」でつなぐ。"""
    result = {}
    stack = []
    lines = read_cr_text(path)
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        i += 1
        if not line:
            continue
        if line == "}":
            if stack:
                stack.pop()
            continue
        if not line.startswith("/"):
            continue
        name, value = split_key_line(line)
        if name is None:
            continue
        full = "/".join(stack + [name])
        if value == "{":
            stack.append(name)
            continue
        if ARR_RE.match(value) or BLOB_RE.match(value):
            closer = "]" if value.startswith("[") else ">"
            size = value.split()[1]
            while i < len(lines) and lines[i].strip() != closer:
                i += 1
            i += 1
            result[full] = ("Array" if closer == "]" else "Blob", "%s bytes" % size)
            continue
        result[full] = classify_value(value)
    return result


def load_prefs(settings_dir):
    """{"main": {...}, "cloud": {...}} を返す。無いファイルは空。"""
    data = {}
    for file_id, filename in PREF_FILES:
        path = os.path.join(settings_dir, filename)
        data[file_id] = parse_pref_file(path) if os.path.exists(path) else {}
    return data


# ---------------------------------------------------------------- .kys


def parse_kys(path):
    """.kys の /Menus・/Tools の ID を {"menu": [...], "tool": [...]} で返す（ファイル内の順）。"""
    result = {"menu": [], "tool": []}
    section = None
    depth = 0
    for raw in read_cr_text(path):
        line = raw.strip()
        if not line:
            continue
        if depth == 0:
            if line.startswith("/Menus"):
                section, depth = "menu", 1
            elif line.startswith("/Tools"):
                section, depth = "tool", 1
            continue
        if line == "}":
            depth -= 1
            continue
        name, value = split_key_line(line)
        if name is None:
            continue
        if value == "{":
            if depth == 1 and name.strip():
                result[section].append(name)
            depth += 1
    return result


# ---------------------------------------------------------------- 原本


def load_reference(path):
    """ExtendScript.txt を読み、(環境設定キー {キー: [型...]}, メニュー・ツール ID の集合) を返す。"""
    prefs = {}
    ids = set()
    if not path or not os.path.exists(path):
        return prefs, ids
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.rstrip("\r\n")
            m = REF_PREF_LINE.match(line)
            if m:
                prefs.setdefault(m.group(2), set()).add(m.group(3))
                continue
            m = REF_MENU_LINE.match(line)
            if m and m.group(2).strip():
                ids.add(m.group(2))  # 末尾の空白は残す
    return prefs, ids


def id_note(command_id, ref_ids):
    notes = []
    if command_id not in ref_ids:
        notes.append("原本に無い")
    if THIRD_PARTY_RE.search(command_id):
        notes.append("他社プラグイン？")
    return "・".join(notes)


def embedded_lines(jsx_path):
    """AiCommandPrefLookup.jsx の getCommandListLines() に並ぶ文字列を返す。"""
    src = open(jsx_path, encoding="utf-8").read()
    start = src.index("function getCommandListLines()")
    end = src.index("];", start)
    return [json.loads(m) for m in re.findall(r'^\s*("(?:[^"\\]|\\.)*"),?\s*$', src[start:end], re.M)]


def reference_lines(ref_path):
    """原本のうち埋め込み対象の行（タブを含み、// で始まらない行）を返す。"""
    with open(ref_path, encoding="utf-8") as f:
        return [l.rstrip("\r\n") for l in f if "\t" in l and not l.startswith("//")]


def type_matches_ref(file_type, file_value, ref_types):
    if file_type in ref_types:
        return True
    if file_type == "Integer" and "Boolean" in ref_types and file_value in ("0", "1"):
        return True
    if file_type == "Integer" and "Real" in ref_types:
        return True
    return False


# ---------------------------------------------------------------- 出力


def md_escape(text):
    return text.replace("|", "\\|").replace("`", "'")


def code(text):
    # 末尾の空白が見えるように、空白で終わるものは引用符で囲む
    if text != text.rstrip() or text != text.lstrip():
        return "`'%s'`" % md_escape(text)
    return "`%s`" % md_escape(text)


def bullet_keys(out, keys, info=None, limit=None):
    shown = keys if limit is None else keys[:limit]
    for k in shown:
        extra = info(k) if info else ""
        out.append("- %s%s" % (code(k), ("　" + extra) if extra else ""))
    if limit is not None and len(keys) > limit:
        out.append("- …ほか %d 件（--all で全件）" % (len(keys) - limit))
    if not keys:
        out.append("- なし")


def parent_of(key):
    return key.rsplit("/", 1)[0] if "/" in key else ""


def find_renames(removed, added):
    """消えたキーと現れたキーの組で、改名らしいもの（同じ親の中で大小違い・似た名前）を返す。"""
    pairs = []
    added_lower = {a.lower(): a for a in added}
    used = set()
    for r in removed:
        if r.lower() in added_lower and added_lower[r.lower()] not in used:
            a = added_lower[r.lower()]
            pairs.append((r, a, "大小のみ"))
            used.add(a)
            continue
        cands = [a for a in added if a not in used and parent_of(a) == parent_of(r)]
        best = difflib.get_close_matches(r, cands, n=1, cutoff=0.8)
        if best:
            pairs.append((r, best[0], "似た名前"))
            used.add(best[0])
    return pairs


def resolve_settings_dir(arg, lang):
    if os.path.isdir(arg):
        return arg
    return os.path.join(PREF_ROOT, "Adobe Illustrator %s Settings" % arg, lang)


def build_report(args):
    old_dir = resolve_settings_dir(args.old, args.lang)
    new_dir = resolve_settings_dir(args.new, args.lang)
    for d in (old_dir, new_dir):
        if not os.path.isdir(d):
            sys.exit("設定フォルダーが見つかりません: %s" % d)
    # 起動中の Illustrator は環境設定ファイルを消し、終了時に書き直す。無いまま比べると全キーが「新だけ」「旧だけ」になる
    missing = [
        os.path.join(d, filename)
        for d in (old_dir, new_dir)
        for _, filename in PREF_FILES
        if not os.path.exists(os.path.join(d, filename))
    ]
    if missing and not args.allow_missing:
        sys.exit(
            "設定ファイルがありません（Illustrator の起動中は消えていて、終了時に書き直されます）。\n"
            "該当するバージョンの Illustrator を終了してから実行してください。\n  "
            + "\n  ".join(missing)
        )

    old = load_prefs(old_dir)
    new = load_prefs(new_dir)
    data_counts = {}
    if not args.include_data:
        for label, prefs in (("旧", old), ("新", new)):
            n = 0
            for file_id in prefs:
                dropped = [k for k in prefs[file_id] if DATA_KEY_RE.search(k)]
                n += len(dropped)
                for k in dropped:
                    del prefs[file_id][k]
            data_counts[label] = n
    ref_prefs, ref_ids = load_reference(args.ref)
    limit = None if args.all else args.limit

    out = []
    out.append("# Illustrator 設定キーの差分 %s → %s" % (args.old, args.new))
    out.append("")
    out.append("作成日 %s" % date.today().isoformat())
    out.append("")
    out.append("- 旧: `%s`" % old_dir)
    out.append("- 新: `%s`" % new_dir)
    out.append("- 原本: `%s`（環境設定キー %d 件・ID %d 件）" % (args.ref, len(ref_prefs), len(ref_ids)))
    out.append("")
    out.append("> 設定ファイルには一度変更・使用したキーしか書かれない。「消えた」は廃止の根拠にならない。")
    out.append("> get〜Preference は存在しないキーも既定値で書き込むので、ファイルにある＝実在とも限らない。新しいキーは GUI で切り替えて値が変わるか確かめる。")
    out.append("> ⚠ は作業の痕跡（最近使った・表示回数・位置など）らしい名前の印。")
    out.append("")

    # 概要
    out.append("## 概要")
    out.append("")
    for file_id, filename in PREF_FILES:
        o, n = old[file_id], new[file_id]
        out.append(
            "- %s：旧 %d キー → 新 %d キー（新だけ %d・旧だけ %d・共通 %d）"
            % (FILE_LABEL[file_id], len(o), len(n), len(set(n) - set(o)), len(set(o) - set(n)), len(set(o) & set(n)))
        )
    if data_counts:
        out.append(
            "- データの枝（アクションの中身・ライブラリーのパス・キャッシュ・最近使ったフォント）は除外：旧 %d・新 %d（--include-data で含める）"
            % (data_counts["旧"], data_counts["新"])
        )
    out.append("")

    def all_keys(data):
        merged = {}
        for file_id, _ in PREF_FILES:
            for k, v in data[file_id].items():
                merged.setdefault(k, []).append((file_id, v))
        return merged

    old_all, new_all = all_keys(old), all_keys(new)

    def describe(k, data_all):
        parts = []
        for file_id, (t, v) in data_all.get(k, []):
            parts.append("%s %s %s" % (FILE_LABEL[file_id], t, md_escape(v)[:40]))
        info = " / ".join(parts)
        if k in ref_prefs:
            info += "　原本:" + ",".join(sorted(ref_prefs[k]))
        if NOISE_RE.search(k):
            info += "　⚠"
        return info

    added = sorted(set(new_all) - set(old_all), key=str.lower)
    removed = sorted(set(old_all) - set(new_all), key=str.lower)
    renames = find_renames(removed, added)

    out.append("## 1. 新バージョンで現れたキー（%d）" % len(added))
    out.append("")
    out.append("新しい機能の設定キーの候補。原本に無いものは追加を検討する。")
    out.append("")
    bullet_keys(out, added, lambda k: describe(k, new_all), limit)
    out.append("")

    out.append("## 2. 改名らしいキー（%d）" % len(renames))
    out.append("")
    if renames:
        for r, a, why in renames:
            out.append("- %s → %s（%s）" % (code(r), code(a), why))
    else:
        out.append("- なし")
    out.append("")

    out.append("## 3. 旧バージョンにだけあるキー（%d）" % len(removed))
    out.append("")
    out.append("新で未使用なだけのことが多い。廃止かどうかは新バージョンで get して確かめる。")
    out.append("")
    bullet_keys(out, removed, lambda k: describe(k, old_all), limit)
    out.append("")

    # 所属ファイルの移動・重複
    moved = []
    for k in sorted(set(old_all) & set(new_all), key=str.lower):
        of = sorted(f for f, _ in old_all[k])
        nf = sorted(f for f, _ in new_all[k])
        if of != nf:
            moved.append((k, of, nf))
    dup = sorted((k for k, v in new_all.items() if len(v) > 1), key=str.lower)
    out.append("## 4. 所属ファイルが変わったキー（%d）・両方にあるキー（%d）" % (len(moved), len(dup)))
    out.append("")
    out.append("両方にあるキーは、Cloud Prefs 側が実体で環境設定側は get で書かれた残骸のことがある（enableBackgroundSave など）。")
    out.append("")
    for k, of, nf in moved:
        out.append("- %s：%s → %s" % (code(k), "+".join(FILE_LABEL[f] for f in of), "+".join(FILE_LABEL[f] for f in nf)))
    for k in dup:
        out.append("- %s：両方（%s）" % (code(k), describe(k, new_all)))
    if not moved and not dup:
        out.append("- なし")
    out.append("")

    # 型の変化
    type_changed = []
    for k in sorted(set(old_all) & set(new_all), key=str.lower):
        ot = {t for _, (t, _) in old_all[k]}
        nt = {t for _, (t, _) in new_all[k]}
        if ot != nt:
            type_changed.append((k, ot, nt))
    out.append("## 5. 型が変わったキー（%d）" % len(type_changed))
    out.append("")
    for k, ot, nt in type_changed:
        out.append("- %s：%s → %s" % (code(k), ",".join(sorted(ot)), ",".join(sorted(nt))))
    if not type_changed:
        out.append("- なし")
    out.append("")

    # 原本との照合
    readable = ("Integer", "Real", "String")
    not_in_ref = sorted(
        (k for k, v in new_all.items() if k not in ref_prefs and any(t in readable for _, (t, _) in v)),
        key=str.lower,
    )
    mismatch = []
    for k, v in new_all.items():
        if k in ref_prefs:
            for _, (t, val) in v:
                if t in readable and not type_matches_ref(t, val, ref_prefs[k]):
                    mismatch.append((k, t, ref_prefs[k]))
                    break
    mismatch.sort(key=lambda x: x[0].lower())
    out.append("## 6. 原本に無いキー（新バージョン・%d）" % len(not_in_ref))
    out.append("")
    out.append("Integer / Real / String のものだけ。Array・Blob・Long（末尾 L）は app.preferences で扱いにくいので除く。")
    out.append("")
    bullet_keys(out, not_in_ref, lambda k: describe(k, new_all), limit)
    out.append("")
    out.append("## 7. 原本と型が合わないキー（%d）" % len(mismatch))
    out.append("")
    for k, t, rt in mismatch:
        out.append("- %s：ファイル %s・原本 %s" % (code(k), t, ",".join(sorted(rt))))
    if not mismatch:
        out.append("- なし")
    out.append("")

    # 値の変化（ユーザーの設定変更も混ざる）
    changed = []
    for k in sorted(set(old_all) & set(new_all), key=str.lower):
        ov = {(f, v) for f, v in old_all[k]}
        nv = {(f, v) for f, v in new_all[k]}
        if ov != nv and all(t in readable for _, (t, _) in new_all[k]):
            o_txt = " / ".join(md_escape(v[1])[:30] for _, v in old_all[k])
            n_txt = " / ".join(md_escape(v[1])[:30] for _, v in new_all[k])
            changed.append((k, o_txt, n_txt))
    out.append("## 8. 値が変わったキー（%d）" % len(changed))
    out.append("")
    out.append("自分で変えた設定も混ざる。既定値が変わった手掛かりとして見る。")
    out.append("")
    shown = changed if limit is None else changed[:limit]
    for k, o_txt, n_txt in shown:
        out.append("- %s：%s → %s" % (code(k), o_txt, n_txt))
    if limit is not None and len(changed) > limit:
        out.append("- …ほか %d 件（--all で全件）" % (len(changed) - limit))
    if not changed:
        out.append("- なし")
    out.append("")

    # .kys
    if args.kys_new:
        kys_new = parse_kys(args.kys_new)
        kys_old = parse_kys(args.kys_old) if args.kys_old else None
        out.append("## 9. メニューコマンド・ツール ID（.kys）")
        out.append("")
        out.append("- 新: `%s`（メニュー %d・ツール %d）" % (args.kys_new, len(kys_new["menu"]), len(kys_new["tool"])))
        if kys_old:
            out.append("- 旧: `%s`（メニュー %d・ツール %d）" % (args.kys_old, len(kys_old["menu"]), len(kys_old["tool"])))
        out.append("")
        out.append("> .kys に無い ID でも動くもの（ショートカット対象外）がある。無い＝廃止とは限らない。")
        out.append("")
        for kind, label in (("menu", "メニュー"), ("tool", "ツール")):
            n_ids = kys_new[kind]
            if kys_old:
                o_ids = kys_old[kind]
                k_added = [x for x in n_ids if x not in set(o_ids)]
                k_removed = [x for x in o_ids if x not in set(n_ids)]
                k_renames = find_renames(k_removed, k_added)
                out.append("### %s：新だけ（%d）" % (label, len(k_added)))
                out.append("")
                bullet_keys(out, k_added, lambda x: id_note(x, ref_ids))
                out.append("")
                k_removed_own = [x for x in k_removed if not THIRD_PARTY_RE.search(x)]
                out.append(
                    "### %s：旧だけ（%d・他社プラグイン %d は除外）"
                    % (label, len(k_removed_own), len(k_removed) - len(k_removed_own))
                )
                out.append("")
                bullet_keys(out, k_removed_own, lambda x: "" if x in ref_ids else "原本に無い")
                out.append("")
                out.append("### %s：改名らしい ID（%d）" % (label, len(k_renames)))
                out.append("")
                for r, a, why in k_renames:
                    out.append("- %s → %s（%s）" % (code(r), code(a), why))
                if not k_renames:
                    out.append("- なし")
                out.append("")
            missing = [x for x in n_ids if x not in ref_ids and not THIRD_PARTY_RE.search(x)]
            n_third = len([x for x in n_ids if x not in ref_ids and THIRD_PARTY_RE.search(x)])
            out.append("### %s：新の .kys にあって原本に無い ID（%d・他社プラグイン %d は除外）" % (label, len(missing), n_third))
            out.append("")
            tool_note = None
            if kind == "tool":
                # /Tools には描画モード・不透明度など selectTool で選べないものも入る
                tool_note = lambda x: "" if x.endswith(" Tool") else "selectTool の対象外？"
            bullet_keys(out, missing, tool_note, limit)
            out.append("")

    # 原本と埋め込み先のずれ
    if args.jsx and os.path.exists(args.jsx) and os.path.exists(args.ref):
        emb = embedded_lines(args.jsx)
        ref = reference_lines(args.ref)
        only_ref = [l for l in ref if l not in set(emb)]
        only_jsx = [l for l in emb if l not in set(ref)]
        out.append("## 10. 原本と AiCommandPrefLookup.jsx のずれ（原本だけ %d・jsx だけ %d）" % (len(only_ref), len(only_jsx)))
        out.append("")
        out.append("- jsx: `%s`（%d 行）" % (os.path.normpath(args.jsx), len(emb)))
        out.append("")
        for l in only_ref:
            out.append("- 原本だけ：%s" % code(l.replace("\t", " ⇥ ")))
        for l in only_jsx:
            out.append("- jsx だけ：%s" % code(l.replace("\t", " ⇥ ")))
        if not only_ref and not only_jsx:
            out.append("- なし")
        out.append("")

    return "\n".join(out) + "\n"


def main():
    p = argparse.ArgumentParser(description="Illustrator の設定キーとメニュー ID を新旧バージョンで突き合わせる")
    p.add_argument("old", help="旧バージョン（例: 30）または設定フォルダーのパス")
    p.add_argument("new", help="新バージョン（例: 31）または設定フォルダーのパス")
    p.add_argument("--lang", default="ja_JP", help="言語フォルダー（既定: ja_JP）")
    p.add_argument("--ref", default=DEFAULT_REF, help="原本 ExtendScript.txt のパス")
    p.add_argument("--jsx", default=DEFAULT_JSX, help="原本を埋め込む AiCommandPrefLookup.jsx（ずれの確認用）")
    p.add_argument("--kys-old", help="旧バージョンの初期設定を別名保存した .kys")
    p.add_argument("--kys-new", help="新バージョンの初期設定を別名保存した .kys")
    p.add_argument("--limit", type=int, default=200, help="長い一覧の表示件数（既定: 200）")
    p.add_argument("--all", action="store_true", help="一覧を省略せずに出す")
    p.add_argument("--allow-missing", action="store_true", help="設定ファイルが欠けていても比べる（通常は使わない）")
    p.add_argument("--include-data", action="store_true", help="アクション・ライブラリー・キャッシュなどデータの枝も比べる")
    p.add_argument("--out", help="Markdown の書き出し先（省略時は標準出力）")
    args = p.parse_args()

    report = build_report(args)
    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(report)
        print("書き出しました: %s" % args.out)
    else:
        sys.stdout.write(report)


if __name__ == "__main__":
    main()
