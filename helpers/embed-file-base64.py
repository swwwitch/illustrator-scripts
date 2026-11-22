#!/usr/bin/env python3
"""
ファイルを Base64 にして .jsx に埋め込む（埋め込み済みなら差し替える）。

.jsx の中の「/* EMBEDDED-FILE-BEGIN …」の行から「/* EMBEDDED-FILE-END */」の行までを、
次の関数で書き直す。スクリプトはこれを読んでファイルを復元する。
- getEmbeddedFileName()        元のファイル名
- getEmbeddedFileSize()        バイト数（復元したファイルの確かめに使う）
- getEmbeddedFileId()          中身の SHA-1 の先頭8桁（中身が変わると変わる。復元先のフォルダー名に使う）
- getEmbeddedFileBase64Lines() Base64 の行の配列（1行 4096 文字 = 3072 バイト。行ごとに単独でデコードできる）

    python3 helpers/embed-file-base64.py <埋め込み先の.jsx> <埋め込むファイル>

Embeds a file into a .jsx as Base64, rewriting the block between the EMBEDDED-FILE-BEGIN / -END marker lines.
"""
import base64
import hashlib
import os
import sys

BEGIN_MARKER = "/* EMBEDDED-FILE-BEGIN"
END_MARKER = "/* EMBEDDED-FILE-END */"
LINE_BYTES = 3072  # 1行に入れるバイト数（3 の倍数なので行ごとにデコードできる）/ bytes per line (a multiple of 3)
INDENT = "    "


def build_block(file_path, indent):
    """埋め込みブロックの行を作る / build the lines of the embedded block"""
    with open(file_path, "rb") as source:
        data = source.read()
    file_name = os.path.basename(file_path)
    file_id = hashlib.sha1(data).hexdigest()[:8]
    base64_lines = []
    for offset in range(0, len(data), LINE_BYTES):
        base64_lines.append(base64.b64encode(data[offset:offset + LINE_BYTES]).decode("ascii"))
    quoted_name = file_name.replace("\\", "\\\\").replace('"', '\\"')

    i1 = indent
    i2 = indent + INDENT
    i3 = indent + INDENT * 2
    block = [
        i1 + BEGIN_MARKER + "（helpers/embed-file-base64.py が書き直す。手で編集しない / rewritten by helpers/embed-file-base64.py; do not edit） */",
        "",
        i1 + "/**",
        i1 + " * 埋め込んだファイルの元のファイル名を返す",
        i1 + " * @returns {string} ファイル名",
        i1 + " */",
        i1 + "function getEmbeddedFileName() {",
        i2 + 'return "' + quoted_name + '";',
        i1 + "}",
        "",
        i1 + "/**",
        i1 + " * 埋め込んだファイルのバイト数を返す",
        i1 + " * @returns {number} バイト数",
        i1 + " */",
        i1 + "function getEmbeddedFileSize() {",
        i2 + "return " + str(len(data)) + ";",
        i1 + "}",
        "",
        i1 + "/**",
        i1 + " * 埋め込んだファイルの中身の ID（SHA-1 の先頭8桁）を返す。中身が変わると変わる",
        i1 + " * @returns {string} ID",
        i1 + " */",
        i1 + "function getEmbeddedFileId() {",
        i2 + 'return "' + file_id + '";',
        i1 + "}",
        "",
        i1 + "/**",
        i1 + " * 埋め込んだファイルの Base64 を行の配列で返す（1行 " + str(LINE_BYTES) + " バイト分。行ごとにデコードできる）",
        i1 + " * @returns {string[]} Base64 の行",
        i1 + " */",
        i1 + "function getEmbeddedFileBase64Lines() {",
        i2 + "return [",
    ]
    for index, line in enumerate(base64_lines):
        block.append(i3 + '"' + line + '"' + ("," if index < len(base64_lines) - 1 else ""))
    block += [
        i2 + "];",
        i1 + "}",
        "",
        i1 + END_MARKER,
    ]
    return block, len(data), file_id


def main():
    if len(sys.argv) != 3:
        sys.stderr.write("usage: embed-file-base64.py <target.jsx> <file>\n")
        return 2
    jsx_path, file_path = sys.argv[1], sys.argv[2]
    with open(jsx_path, "r", encoding="utf-8", newline="") as jsx_file:
        lines = jsx_file.read().split("\n")

    begin_index = end_index = None
    for index, line in enumerate(lines):
        if line.strip().startswith(BEGIN_MARKER) and begin_index is None:
            begin_index = index
        elif line.strip() == END_MARKER and begin_index is not None:
            end_index = index
            break
    if begin_index is None or end_index is None:
        sys.stderr.write("markers not found: " + BEGIN_MARKER + " … " + END_MARKER + "\n")
        return 1

    indent = lines[begin_index][:len(lines[begin_index]) - len(lines[begin_index].lstrip())]
    block, size, file_id = build_block(file_path, indent)
    lines[begin_index:end_index + 1] = block
    with open(jsx_path, "w", encoding="utf-8", newline="") as jsx_file:
        jsx_file.write("\n".join(lines))
    print("embedded %s (%d bytes, id %s) into %s" % (os.path.basename(file_path), size, file_id, jsx_path))
    return 0


if __name__ == "__main__":
    sys.exit(main())
