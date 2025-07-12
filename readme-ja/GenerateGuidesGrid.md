# グリッド状にガイドを生成

[![Direct](https://img.shields.io/badge/Direct%20Link-GenerateGuidesGrid.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/guide/GenerateGuidesGrid.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### ムービー

[ムービー](https://youtu.be/7sNMYhyTwfY)

<img alt="" src="https://www.dtp-transit.jp/images/ss-738-1182-72-20250713-082813.png" width="80%" />

### 仕様

- 行数・列数の指定による分割数
- 上・下・左・右マージンの個別指定
- 共通マージン入力欄とチェックボックス（有効時は上下左右マージンを一括適用）
- アートボード外への延長距離
- 対象を「アクティブなアートボードのみに適用」または「すべてのアートボードに適用」の選択
- プリセットからの設定読み込み（行列数・マージン・延長距離）

また、次のようになっています。

- ［適用］ボタンをクリックすると、現在の設定に応じたガイドプレビューが表示される
- ［すべて同じにする］チェックボックスは、上下左右のマージンが一致しない場合に自動でOFFになる
- OKを押すと「ガイドをロック」状態に設定される

### 標準ツールでなく、このスクリプトを使うメリット

- アートボードを対象にする場合、アートボードサイズの長方形を描く手間が省ける
- すべてのアートボードを対象にできる
- アートボードに対して上下左右、個別にマージンを設定できる
- アートボードからどれだけ伸ばすか、指定できる

### 段組設定との違い

スクリプトの方は次のような特徴があります。

- アートボードが対象（段組設定はパスが対象）
- ガイドの長さを調整できる（段組設定はガイドが長い）
- 上下左右のマージンを個別に設定できる
- プリセットを登録できる
- 生成される長方形が別レイヤーに、不透明度を15%に設定
