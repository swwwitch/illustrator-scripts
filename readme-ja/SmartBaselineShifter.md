# 特定の文字のベースラインシフトを調整

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartBaselineShifter.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/SmartBaselineShifter.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

電話番号のハイフンのように特定の文字のみ、ベースラインシフトを調整したい場面があります。

これを解決するアプローチについて、過去にも取り上げましたが、決定版的なスクリプトを作成しました。

<img alt="" src="https://www.dtp-transit.jp/images/ss-678-596-72-20250713-081901.png" width="70%" />

### 対象文字列

- 選択しているテキストから、数字やスペースなどを除いたものが自動的に入ります。
- 編集可能です。
- 「:-」と入力すれば「:」と「-」の両方が対象になります。

### シフト量

- 「-3.2」のように入力するのが面倒でならないので、整数、小数点以下のテキストフィールドを分け、マイナスにしたい場合にはラジオボタンを選択するだけにしました。
- 結果は「シフト量」にプレビューされます。

### リセット
- 対象文字列だけでなく、選択しているテキストのすべての文字を対象にリセットします。

### 懸念事項
文字スタイルを呼び出して適用したり、文字スタイルに登録できるとよいかも


解説記事：

- [note](https://note.com/dtp_tranist/n/n2e19ad0bdb83)