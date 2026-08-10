# SymbolizeAndReplace

[![Direct](https://img.shields.io/badge/Direct%20Link-SymbolizeAndReplace.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/symbol/SymbolizeAndReplace.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SymbolizeAndReplace.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 紹介記事（note)

https://note.com/dtp_tranist/n/n650a4b91329d

Overview

Illustrator JSX script that converts the selected object into a symbol and
replaces matching items in the document with instances of that symbol.

- Registers the symbol via a dialog for its name and a 3×3 registration point (existing names are rejected)
- For a TextFrame, seeds the symbol name from its text and targets frames with the same font, style, and contents
  (enable "Include different font sizes" to ignore size)
- Otherwise targets similar objects via SmartEdit bulk selection
- With multiple groups selected, replaces every selected group with one symbol (a multi-selection containing non-groups is rejected)
- Replaces each target with a symbol instance, aligned by the chosen registration point
- Leaves the new symbol instances selected after replacement
- Reports how many items could not be replaced (locked/hidden, etc.)

### スクリプト情報

- バージョン: v1.0.1
