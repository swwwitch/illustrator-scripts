# オブジェクトを指定レイヤーへ移動

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartLayerManage.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/SmartLayerManage.jsx)


[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

Illustratorで複雑なアートワークを扱うとき、オブジェクトやテキストがバラバラのレイヤーに散らばってしまいがち。「後からレイヤーを整理したいけど、一つずつ手動で動かすのは大変…」です。

単純に選択しているオブジェクトをほかのレイヤーに移動するのも、ちょっと面倒です。

これらを解決するスクリプトを書きました。

- 選択したオブジェクトだけを移動
- ドキュメント内のすべてのオブジェクトを一括移動
- テキストオブジェクトのみを選んで移動
- 不要になった空のレイヤーを自動削除 （「bg」や「//」で始まるレイヤーは除外）
- 標準機能の「すべてのレイヤーを結合」をダイナミックアクションを使って実行しています。否応なく、一番上のレイヤーに結合されます。
- サブレイヤーはそのままです。
- レイヤーカラーは、ライトブルーに近いカラーになります。
- ロック／非表示のオブジェクト／レイヤーがある場合、アラートが表示されます。［いいえ］をクリックすると、そのままの状態で結合します

![](https://assets.st-note.com/img/1751590763-zd5ArG9bytCocY2wOEX8qmKT.png?width=1200)
    

### 更新履歴：

- v1.0.0（2025-07-03）: 初版リリース
- v1.0.1（2025-07-03）: レイヤーカラー変更機能追加
- v1.0.2（2025-07-03）: 自動選択判定、空レイヤー削除ロジック改善
- v1.0.3（2025-07-04）: 「すべて（強制）」モード追加（すべてのレイヤーを結合）

## note

- [【Illustrator】オブジェクトを指定したレイヤーに移動するスクリプト｜DTP Transit 別館](https://note.com/dtp_tranist/n/n95ec4929ae9d)