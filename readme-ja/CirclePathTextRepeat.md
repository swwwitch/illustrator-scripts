# 円周に沿ってテキストを繰り返す

[![Direct](https://img.shields.io/badge/Direct%20Link-CirclePathTextRepeat.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/CirclePathTextRepeat.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CirclePathTextRepeat.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 円（パス）とテキストを 1 つずつ選択し、テキストを指定回数繰り返して、円を複製したパス上の文字に変換する
- 区切り文字は「スペース」または任意の入力文字（初期値は欧文 bullet「•」）を選択でき、末尾にも区切り文字を付与
- 区切り文字の前後に入れる半角スペース数を指定可能（スペースのみの場合は区切りの間隔になる）
- スペース以外の区切り文字は、スケール（水平・垂直比率）とベースライン（環境設定のテキスト単位）を調整可能
- 円周に合わせた文字サイズの自動調整は ON/OFF 可能（補正率で開始・終了の隙間を微調整）
- 生成結果を円の中心基準で回転
- 数値フィールドは ↑↓ キーで増減（Shift で ±10・10 の倍数にスナップ、Option で ±0.1）
- プレビュー対応（確定までは元のテキスト・円を保持し、OK で元を削除して生成結果を選択）

### 紹介記事

[DTP Transit 別館](https://note.com/dtp_tranist/n/na9334a217ec3)

### 更新履歴

- v1.0.0 (2026-06-12) : 初版
- v1.0.1 (2026-09-07) : 区切り文字のスケール・ベースラインを文字の内容ではなく位置で適用するよう修正（元テキストに同じ文字が含まれていても影響しない）。Esc / Enter でキャンセル・OK が効くように修正。文字サイズを Illustrator の上限・下限に収めるよう修正。元テキストの改行をスペースに置き換え。UI文言を実際の動作に合わせて調整（「連結文字」→「区切り文字」、「円周の補正」→「補正率」など）し、数値欄のツールチップに ↑↓ 操作を明記。命名・JSDoc・レイアウトをハウスルールに合わせて整理
