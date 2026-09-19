# アクティブなアートボードだけを残す

[![Direct](https://img.shields.io/badge/Direct%20Link-KeepActiveArtboardOnly.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/KeepActiveArtboardOnly.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/KeepActiveArtboardOnly.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

アクティブなアートボードだけを残すか、空のアートボードをまとめて削除します。

### 主な機能

- アクティブなアートボード以外を削除し、その外側にあるオブジェクトとオブジェクトガイドも削除（オフにするとアートボードだけを削除）
- ロック・非表示・テンプレートのレイヤーやグループも一時的に解除して削除し、処理後に元の状態へ戻す
- ガイドは境界ボックスではなくパスの実体で重なりを判定
- 空のアートボードをまとめて削除（非表示のレイヤー／オブジェクトを無視するかを切り替え可能）
- 削除対象アートボード数の表示は、トグルに応じて即時更新
- 日本語／英語UI

### 使い方

1. 残したいアートボードをアクティブにします。
2. スクリプトを実行します。
3. 削除するアートボードとオプションを選び、［OK］をクリックします。

### 注意点

- 元に戻せない変更を加えるため、実行前にファイルを複製しておくことを推奨します。
- ルーラーガイドは削除対象外です。
- 非表示判定は祖先を遡って行うため、グループ内の非表示アイテムや、非表示グループ／レイヤー配下のアイテムも同じ扱いになります。
- アートボードが1枚しかないときは［空のアートボード］を選べません。

### 更新履歴

- v1.2.1: RemoveOtherArtboards.jsx、RemoveEmptyArtboards.jsx を統合し、削除するアートボードをダイアログで選択する形に変更
- v1.2: 初出
