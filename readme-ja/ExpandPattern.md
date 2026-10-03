# パターンの塗りを通常のパスに展開

[![Direct](https://img.shields.io/badge/Direct%20Link-ExpandPattern.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fx/ExpandPattern.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ExpandPattern.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択オブジェクトをグループ化し、3D 回転（クラシック）と［合流］の効果を掛けてからアピアランスを分割します。パターンなどの塗りを、見た目を変えずに通常のパスへ展開します。

### 使い方

1. 展開したいオブジェクトを選択します。
2. スクリプトを実行します。

ダイアログボックスは表示されず、そのまま実行されます。

### 処理の流れ

1. 選択オブジェクトをグループ化し、グループ名を「パターンの分割」にする
2. グループに次の効果を掛ける（アピアランスパネルでは「内容」の下に、この順で並ぶ）
   - 3D 回転（クラシック）: 位置「前面」（X/Y/Z すべて 0°）、表面「陰影なし」
   - パスファインダー［合流］
3. ［オブジェクト］＞［アピアランスを分割］を実行する

### 注意点

- 効果は LiveEffect の XML で適用しています。パラメーターは [live-effect-functions-for-illustrator](https://github.com/mark1bean/live-effect-functions-for-illustrator) の既定値をもとにしています。
- ［合流］はメニューコマンドで掛けると「内容」の上に入るため、XML で「内容」の下に積んでいます。

### 更新履歴

- v1.0.0（2026-10-04）初版
