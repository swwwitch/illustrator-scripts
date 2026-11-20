#!/usr/bin/env python3
"""
SCRIPTMETA（SCRIPTMETAKit 2.1、https://github.com/Yamonov/SCRIPTMETAKit/blob/main/docs/MetadataAuthoring.md）の同期ツール。

スクリプトの基本情報は先頭の SCRIPTMETA ブロックが元で、スクリプトは実行時に「スクリプト情報」の部品
（jsx/_templates/ScriptMeta.jsx）で読んで SCRIPT_NAME などに入れる。
Version・Author・Release-Date・Description・X-com.swwwitch-Script-Name・Updated・README-JA/EN・Article-URL などは
SCRIPTMETA を直接書き換える。このツールは残りを決まった形にそろえる：
- Script-ID が無ければ ID_PREFIX＋Script-Name（ID_OVERRIDES）で作る。あれば変えない
- Meta-Default-language=ja、Target-App
- Name は X-com.swwwitch-README-JA が指す README の見出し、Name[en] は README-EN の見出し（README が無ければ今の値を残す）
- Tags は置き場所のフォルダー（jsx/<フォルダー>/…）から英語・日本語の語を入れる（FOLDER_TAGS）
- X-com.swwwitch-Source-URL は GitHub 上の最新版ソースの URL。URL の値は非 ASCII と空白をパーセントエンコードする
- 項目の並びをそろえ、スクリプト情報の部品をテンプレートと同じ中身に書き直す

SCRIPTMETA を書き換えたり、README の見出しを変えたり、スクリプトを追加・移動・改名したりしたら、コミット前に実行する。

    python3 helpers/sync-scriptmeta.py          # 書き換える
    python3 helpers/sync-scriptmeta.py --check  # 差分があれば一覧を出して終了コード1

The SCRIPTMETA block is the source of each script's basic info; scripts read it at run time with the
Script info part. This tool fills in Script-ID, Name / Name[en] (README titles), Tags (folder), Source-URL,
normalizes the field order and refreshes the Script info part from jsx/_templates/ScriptMeta.jsx.
"""
import glob
import os
import re
import sys
import unicodedata
import urllib.parse

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ID_PREFIX = "com.swwwitch.illustrator."
TARGET_APP = "Illustrator"
SOURCE_URL_BASE = "https://github.com/swwwitch/illustrator-scripts/blob/master/"
EXCLUDE_DIRS = ("/_templates/",)
DEFAULT_LANGUAGE = "ja"
X_PREFIX = "X-com.swwwitch-"  # 独自項目の作成者 ID / author ID of the custom fields
PART_TEMPLATE = os.path.join(ROOT, "jsx", "_templates", "ScriptMeta.jsx")
PART_START = "// スクリプト情報（再利用パーツ） / Script info (reusable)"
PART_END = "// スクリプト情報（再利用パーツ）ここまで / End of the reusable script info"

# 独自項目の並び（ここに無い X- 項目は後ろに元の順で残す）/ order of the custom fields
X_ORDER = (
    "Script-Name",    # スクリプト名（SCRIPT_NAME。設定の保存名にも使うので変えない）
    "Modified-By",    # 改変した人（他の人のスクリプトを元にしたとき。Author は元の作者）
    "Updated",        # 更新日
    "README-JA",      # README（日本語）
    "README-EN",      # README（英語）
    "Article-URL",    # 紹介記事
    "Reference-URL",  # 参照記事
    "Pro-URL",        # 上位版
    "Source-URL",     # 最新版のソース（ツールが書く）
)
# 標準項目の並び（Name[en] は Name の直後、Tags と X- 項目は後ろ）/ order of the standard fields
STANDARD_ORDER = ("Script-ID", "Meta-Default-language", "Name", "Name[en]", "Author", "Target-App", "Version", "Release-Date")

# Script-ID に使えない文字（A-Z a-z 0-9 . _ - 以外）を含むスクリプト名と、改名したスクリプトの旧名の ID。
# Script-ID が無い（新しい）スクリプトにだけ使う。一度公開した ID は変えない
# IDs for names Script-ID does not allow, and the old IDs of renamed scripts. Used only when Script-ID is missing
ID_OVERRIDES = {
    "TextHighlighter": "DrawMarkerBehindText",  # 2026-11-01 改名。公開済みの Script-ID を保つ
    "ガイドをテンプレートに": "GuidesToTemplate",
    "レイヤーをテンプレートに": "LayersToTemplate",
    "丸型先端にする": "SetRoundCap",
    "シンボルに置き換え-sw": "ReplaceWithSymbol-sw",
    "AutoLeadingStep+1": "AutoLeadingStepPlus1",
    "AutoLeadingStep+10": "AutoLeadingStepPlus10",
    "段落スタイル": "ParagraphStyles",
    "AutoApplySwatches": "ApplySwatchesToSelection",
    "TileInsideShape": "PatternFill",
    "AiDistributeColors": "AiApplySwatchesToSelection",
}

# jsx/ 直下のフォルダー → Tags（英語・日本語）/ top folder under jsx/ -> Tags
FOLDER_TAGS = {
    "alignment": ("alignment", "整列"),
    "artboard": ("artboard", "アートボード"),
    "blend": ("blend", "ブレンド"),
    "color": ("color", "カラー"),
    "data": ("data", "データ"),
    "demo": ("demo", "デモ"),
    "document": ("document", "ドキュメント"),
    "export": ("export", "書き出し"),
    "files": ("file", "ファイル"),
    "font": ("font", "フォント"),
    "fonts": ("font", "フォント"),
    "frame": ("frame", "フレーム"),
    "fx": ("effect", "効果"),
    "group": ("group", "グループ"),
    "guide": ("guide", "ガイド"),
    "increment": ("numbering", "連番"),
    "layers": ("layer", "レイヤー"),
    "link": ("link", "リンク"),
    "mask": ("mask", "マスク"),
    "misc": ("misc", "その他"),
    "outline": ("outline", "アウトライン"),
    "page": ("page", "ページ"),
    "path": ("path", "パス"),
    "preference": ("preferences", "環境設定"),
    "runner": ("launcher", "ランチャー"),
    "select": ("selection", "選択"),
    "shape": ("shape", "シェイプ"),
    "sort": ("sort", "並べ替え"),
    "stroke-table": ("stroke", "線", "table", "表"),
    "style": ("style", "スタイル"),
    "symbol": ("symbol", "シンボル"),
    "table": ("table", "表"),
    "text": ("text", "テキスト"),
    "transform": ("transform", "変形"),
}

BLOCK_RE = re.compile(r"/\*\nSCRIPTMETA-BEGIN\n(?P<fields>.*?)(?:\nDescription-BEGIN\n(?P<desc>.*?)\nDescription-END)?\nSCRIPTMETA-END\n\*/\n", re.S)
PART_RE = re.compile(re.escape(PART_START) + r"\n.*?" + re.escape(PART_END) + r"\n", re.S)
SCRIPT_ID_RE = re.compile(r"^[A-Za-z0-9._-]+$")
RESERVED_RE = re.compile(r"(SCRIPTMETA(-DIST)?|Description)-(BEGIN|END)", re.I)


def read_part():
    with open(PART_TEMPLATE, encoding="utf-8") as fp:
        match = PART_RE.search(fp.read())
    if not match:
        sys.exit("ERROR " + PART_TEMPLATE + " にスクリプト情報の部品がありません")
    return match.group(0)


def readme_title(url, is_japanese):
    """README の URL（.../blob/<branch>/<path>）が指すファイルの見出し（# 行）を返す。ファイルが無ければ None"""
    m = re.search(r"/blob/[^/]+/(.+)$", url or "")
    if not m:
        return None
    path = os.path.join(ROOT, urllib.parse.unquote(m.group(1)))
    if not os.path.isfile(path):
        return None
    with open(path, encoding="utf-8") as fp:
        heading = re.search(r"^# (.+)$", fp.read(), re.M)
    if not heading:
        return None
    title = unicodedata.normalize("NFC", heading.group(1).strip())
    if "*/" in title or RESERVED_RE.search(title) or (is_japanese and ('"' in title or "\\" in title)):
        raise ValueError("README の見出しに使えない文字があります: " + title)
    return title


def source_url(rel):
    """GitHub 上の最新版ソースの URL / URL of the latest source on GitHub"""
    return SOURCE_URL_BASE + unicodedata.normalize("NFC", rel.replace(os.sep, "/"))


def encode_url(url):
    """X 値用に非 ASCII と空白をパーセントエンコードする（エンコード済みの % はそのまま）"""
    return urllib.parse.quote(unicodedata.normalize("NFC", url), safe=":/?#[]@!$&'()*+,;=%~")


def folder_tags(rel):
    """jsx/<フォルダー>/… のフォルダーから Tags を返す。表に無いフォルダーはエラー"""
    parts = rel.replace(os.sep, "/").split("/")
    if len(parts) < 3:
        return ()
    if parts[1] not in FOLDER_TAGS:
        raise ValueError("FOLDER_TAGS にフォルダーがありません: " + parts[1])
    return FOLDER_TAGS[parts[1]]


def parse_fields(block_text):
    """SCRIPTMETA の項目（Description より前）を [(キー, 値)] で返す"""
    fields = []
    for line in block_text.split("\n"):
        if not line.strip():
            continue
        key, sep, value = line.partition("=")
        if not sep:
            raise ValueError("「キー=値」でない行があります: " + line)
        fields.append((key.strip(), value.strip()))
    return fields


def build_block(rel, fields, desc):
    values = {}
    for key, value in fields:
        if key == "Tags":
            continue
        if key in values:
            raise ValueError("同じ項目が2つあります: " + key)
        values[key] = value

    script_name = values.get(X_PREFIX + "Script-Name")
    if not script_name:
        raise ValueError(X_PREFIX + "Script-Name がありません")
    version = values.get("Version", "")
    if not re.match(r"^\d+(\.\d+)*$", version):
        raise ValueError("Version の形式が不正です（v は付けない）: " + version)
    script_id = values.get("Script-ID") or ID_PREFIX + ID_OVERRIDES.get(script_name, script_name)
    if not SCRIPT_ID_RE.match(script_id):
        raise ValueError("Script-ID に使えない文字があります（ID_OVERRIDES に追加）: " + script_id)

    for key in list(values):
        if key.startswith(X_PREFIX) and values[key].startswith("http"):
            values[key] = encode_url(values[key])
    values["Script-ID"] = script_id
    values["Meta-Default-language"] = DEFAULT_LANGUAGE
    values["Target-App"] = TARGET_APP
    values[X_PREFIX + "Source-URL"] = encode_url(source_url(rel))
    title = readme_title(values.get(X_PREFIX + "README-JA"), True)
    if title:
        values["Name"] = title
    title_en = readme_title(values.get(X_PREFIX + "README-EN"), False)
    if title_en:
        values["Name[en]"] = title_en
    if not values.get("Name"):
        raise ValueError("Name がありません（README が無いときは手で書く）")
    if '"' in values["Name"] or "\\" in values["Name"]:
        raise ValueError("Name に \" \\ は使えません: " + values["Name"])

    lines = ["/*", "SCRIPTMETA-BEGIN"]
    for key in STANDARD_ORDER:
        if values.get(key):
            lines.append(key + "=" + values.pop(key))
        values.pop(key, None)
    tags = folder_tags(rel)
    if tags:
        lines.append("Tags=" + ",".join(tags))
    custom_keys = [X_PREFIX + name for name in X_ORDER] + [k for k in values if k not in [X_PREFIX + n for n in X_ORDER]]
    for key in custom_keys:
        if key not in values:
            continue
        value = values.pop(key)
        if not key.startswith("X-"):
            # 知らない標準項目もそのまま残す / keep unknown standard fields
            lines.append(key + "=" + value)
            continue
        if "*/" in value or re.search(r"[\r\n]", value) or RESERVED_RE.search(value):
            raise ValueError("独自項目に使えない値です: " + key + "=" + value)
        if re.search(r"(^|\s)(Script-ID|Version|Description|Target-App|Min-Target-Version|Meta-URL|Latest-URL|Latest-Version|"
                     r"Latest-Page-URL|Name|Author|Release-Date|Edit-Password-SHA256)\s*[=:]", value, re.I):
            raise ValueError("独自項目の値が旧1.xで項目と誤読されます: " + key + "=" + value)
        lines.append(key + "=" + value)

    if desc is not None:
        for line in desc.split("\n"):
            if "*/" in line or RESERVED_RE.search(line):
                raise ValueError("説明文に使えない行があります: " + line)
        lines += ["Description-BEGIN", desc, "Description-END"]
    lines += ["SCRIPTMETA-END", "*/", ""]
    return script_id, "\n".join(lines)


def main():
    check = "--check" in sys.argv
    part = read_part()
    files = sorted(
        f for f in glob.glob(os.path.join(ROOT, "jsx", "**", "*.jsx"), recursive=True)
        if not any(d in f.replace(os.sep, "/") for d in EXCLUDE_DIRS)
    )
    seen, changed, errors = {}, [], []
    for path in files:
        rel = os.path.relpath(path, ROOT)
        with open(path, encoding="utf-8", newline="") as fp:
            text = fp.read()
        try:
            match = BLOCK_RE.search(text)
            if not match:
                raise ValueError("SCRIPTMETA がありません（既存のスクリプトからコピーして書く）")
            script_id, block = build_block(rel, parse_fields(match.group("fields")), match.group("desc"))
            if script_id in seen:
                raise ValueError("Script-ID が " + seen[script_id] + " と重複しています")
            seen[script_id] = rel
            new_text = text[:match.start()] + block + text[match.end():]
            if len(PART_RE.findall(new_text)) != 1:
                raise ValueError("スクリプト情報の部品がありません（" + os.path.relpath(PART_TEMPLATE, ROOT) + " から貼る）")
            new_text = PART_RE.sub(lambda m: part, new_text, count=1)
        except ValueError as e:
            errors.append(rel + ": " + str(e))
            continue
        if new_text != text:
            changed.append(rel)
            if not check:
                with open(path, "w", encoding="utf-8", newline="") as fp:
                    fp.write(new_text)

    for e in errors:
        print("ERROR " + e)
    for c in changed:
        print(("NEEDS SYNC " if check else "updated ") + c)
    print("%d scripts, %d changed, %d errors" % (len(seen), len(changed), len(errors)))
    sys.exit(1 if errors or (check and changed) else 0)


if __name__ == "__main__":
    main()
