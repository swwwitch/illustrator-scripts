# AddBulletsAndNumbers

[![Direct](https://img.shields.io/badge/Direct%20Link-AddBulletsAndNumbers.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/AddBulletsAndNumbers.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddBulletsAndNumbers.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択したテキストフレームの各行の先頭に、箇条書き記号または連番を付与します。
ダイアログのラジオボタンで「箇条書き」「番号リスト」「なし」を切り替えられ、
プレビューを見ながら設定を調整できます。行の並べ替えにも対応しています。

ダイアログを開くと現在の行頭マーカーから種類・記号・番号スタイル・区切り文字を
推定し、「現状の続き」として編集できます。

- 箇条書き: 記号（• - ● ○ ◎ ■ □ ◆ ◇）を選択
- 番号リスト: 数字／白丸数字／黒丸数字／ABC／abc、開始番号・ゼロ埋め・区切り文字（. ： |）に対応
  （白丸・黒丸数字は箇条書きと同じくタブストップ1つで配置）
- 位置調整: マーカー位置・本文位置・揃え（左／中央／右）を指定
  （既定のタブストップは種類・記号ごとに自動設定。数字=1.5/2.0、ABC/abc=1.0/2.0、箇条書き=1.2、•/-=本文0.8 など、文字サイズ基準）
- マーカーの書式: フォント・スタイル・比率・ベースラインシフトに加え、記号／番号と区切り文字のカラーを個別に指定
- 段落設定: 行送り・段落後のアキ・インデント（折り返し位置を揃えるに対応）
- 行の並べ替え: 行を複数選択して上下移動／先頭・末尾へ移動、名前順・数値順・数値順（逆）で並べ替え可能
- フレームごとにリセット: 複数テキストフレーム選択時に各フレームで連番を振り直し
- リセット: 比率・ベースライン・カラー・インデント・タブストップをまとめて初期化

複数のテキストフレームを選択した場合は、上→下（同じ高さなら左→右）の順に
通し番号が付きます。実行時、行頭の中黒（・ ･ ·）と、後ろにタブ・空白を伴う
ハイフン（-）は自動的に除去します。

Adds a bullet symbol or sequential numbers to the head of each line in the
selected text frames. Switch between "Bullets", "Numbered List", and "None"
with the radio buttons, reorder lines, and tune the settings while watching a live preview.
On open, the current leading markers are detected so you can keep editing
from the existing state.

- Bullets: choose a glyph (• - ● ○ ◎ ■ □ ◆ ◇)
- Numbered: numbers / circled (white/black) / ABC / abc, with start number, zero padding, and delimiter (. ： |)
  (circled numbers use a single tab stop, like bullets)
- Position: marker position, body position, and alignment (left/center/right)
  (default tab stops are set automatically per type/glyph, relative to font size: numbers = 1.5/2.0, ABC/abc = 1.0/2.0, bullets = 1.2, •/- = 0.8 body, etc.)
- Marker format: font family, style, scale, baseline shift, plus separate colors for the marker/number and the delimiter
- Paragraph settings: leading, space-after, and indent (supports aligning wrapped lines to the body)
- Reorder lines: multi-select rows to move up/down or to top/bottom, or sort by name, number, or reverse number
- Restart each frame: restart numbering from the start number in each selected text frame
- Reset: clears scale, baseline, color, indent, and tab stops at once

With multiple frames selected, numbering runs continuously top→bottom
(then left→right). Leading middle dots (・ ･ ·) and a hyphen (-) followed by a
tab/space are stripped on run.

### スクリプト情報

- バージョン: v1.1.1
