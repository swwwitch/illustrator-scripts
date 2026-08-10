# AutoLeadingStep+1

[![Direct](https://img.shields.io/badge/Direct%20Link-AutoLeadingStep+1.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/AutoLeadingStep+1.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AutoLeadingStep+1.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択したテキストの行送り（表示値）が整数で 1 ステップ大きくなるように、自動行送り量（％）を
逆算して設定するスクリプトです。AutoLeadingCalc.jsx を元にした姉妹スクリプトで、ダイアログや
パレットは表示せず、実行するとその場で選択中のテキストへ適用します。
グループ内のテキストやテキスト編集モードの範囲選択にも対応します。

- 段落ごとに、現在の行送りを表示単位（pt / mm / Q(H) など）に換算し、次の整数を目標にする
  （例：26.124 → 27）
- その整数の行送りになるよう「目標行送り ÷ フォントサイズ × 100」で自動行送り量（％）を逆算し、
  autoLeadingAmount に設定して常に自動行送り（autoLeading=true）にする
- 手動行送りの段落も、現在の見た目の行送りを基準に整数へ丸めたうえで自動行送り化される
- 行送りの基準（leadingType）は仮想ボディの上（TOPTOTOP）に固定する
- テキスト編集中に段落内の一部の文字だけを選択している場合は、その段落全体へ適用する
  （行送り量は段落単位の属性のため）

### スクリプト情報

- バージョン: v1.0.0
