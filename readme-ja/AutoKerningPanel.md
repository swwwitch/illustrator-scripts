# カーニング設定パレット

[![Direct](https://img.shields.io/badge/Direct%20Link-AutoKerningPanel.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/AutoKerningPanel.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AutoKerningPanel.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

自動カーニングとプロポーショナルメトリクスは、文字パネルと OpenType パネルに分かれています。どちらも和文組版では毎回触る設定なのに、パネルを2枚開いて往復することになります。

この2つだけを1枚の常駐パレットにまとめました。[統合文字組みパネル](UnifiedTypePanel.md)からこの部分だけを抜き出したもので、カーニングまわりだけを触りたいときはこちらのほうが軽快です。

## 使い方

1. 対象のテキストを選択（テキストフレーム、グループ内のテキスト、テキスト編集モードでの範囲選択のいずれでも可）
2. スクリプトを実行

**適用ボタンはありません。** ラジオボタンやチェックボックスを操作した時点で、その場で選択テキストに適用されます。パレットは `Esc` キーで閉じられます。すでに開いているときに再実行すると、前面化するだけで二重には開きません。

パレットにフォーカスが戻るたびに選択中のテキストの現在値を読み直して UI に反映するので、別のテキストを選び直せば表示も追従します。

## 自動カーニング

| 選択肢 | 内部の値 |
| --- | --- |
| 和文等幅 | METRICSROMANONLY（欧文のみメトリクス） |
| 0 | NOAUTOKERN |
| メトリクス | AUTO |
| オプティカル | OPTICAL |

## プロポーショナルメトリクス

「メトリクス」を選んだときだけ、プロポーショナルメトリクスが自動的に ON になります。それ以外を選ぶと OFF に戻ります。

チェックボックスは単独でも操作できます。この場合はカーニング方式を変えずに、プロポーショナルメトリクスだけを切り替えます。

## 適用範囲

**選択が触れた段落全体に適用されます。** テキスト編集モードで数文字だけ選んで操作しても、その文字だけでなく、その段落全体に当たります。

カーニングやプロポーショナルメトリクスは段落内でバラつくと事故になりやすい設定なので、意図的に段落単位に寄せてあります。

## UI への反映について

パレットが読み取る値は、選択のうち**最初に読めた段落の先頭文字**です。複数のテキストフレームを選択していたり、1つのフレーム内で設定が混在している場合、表示は先頭の値になり実態とはズレます。

## 対象

TextFrame、グループ内のテキスト（再帰的に処理）、テキスト編集モードでの TextRange 選択

## 補足

Illustrator の常駐パレットは、表示している間にドキュメント DOM への接続を失うという事情があるため、実際のオブジェクト操作はすべて BridgeTalk 経由でメインエンジンに委譲しています。委譲するコードは `encodeURIComponent` で包んで送信しています。

パレットの位置は移動のたびに `Folder.userData` に保存され、次回起動時に同じ位置で開きます。モニタ構成が変わって画面外になる位置は無視され、既定位置で開きます。

### note

- [自動カーニングとプロポーショナルメトリクスのパレット｜DTP Transit 別館](https://note.com/dtp_tranist/n/ne7a198a4f527)
