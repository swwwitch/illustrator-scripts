# Adobe Illustrator Scripts

Adobe Illustratorでのデザイン制作に役立つスクリプト集です。主に ChatGPT を活用して開発しています。

## 公開しているスクリプトについて

ここで紹介しているスクリプトは「とりあえず動けば OK」というスタンスで、自分用に作成・利用しているものを掲載しています。そのため、専門的な視点から見るとツッコミどころが多いかもしれません。

以下の方針で公開していますので、ご理解のうえご利用ください。

- 無保証で提供しています。
- 使用によるいかなるトラブルや損害についても責任を負いません。
- 改変・再配布は自由です。
- むしろ「ここはこう直したほうがいい！」という提案や修正は大歓迎です。

### バグ報告

次の情報をお知らせください！

- OS（例：macOS Sequoia 15.5）
- Illustratorのバージョン（例：29.6.1）

エラーメッセージが出る場合、正確にお伝えください。スクショされるのがベストです。

可能でしたら、該当ファイルを共有くださいますと話が早いです（当然ですが、外部には公開しません）。

### アップデート情報

アップデート情報（新規／アップデート）は、noteが一番早いです。

- [DTP Transit 別館｜note](https://note.com/dtp_tranist)

公開後、日々の作業で使う中でバグフィックや調整を行っています。

---

## フォント関連

- [選択したテキストのフォント名、フォントサイズなどをテキストとして生成する](readme-ja/AddTextInfoLabel.md)
- [ドキュメントフォントを適用](readme-ja/ApplyDocumentFonts.md)
- [カテゴリ別ウエイト順にフォントを一覧表示し、フォント見本を一瞬で作成する](readme-ja/TypefaceSampler.md)
- [ドキュメントで使用されているフォント情報を書き出す](readme-ja/ExportFontInfoFromXMP.md)

## テキスト関連

- [文字ばらし](readme-ja/TextSplitterPro.md)
- [特定の文字のベースラインシフトを調整](readme-ja/SmartBaselineShifter.md)
- [PDFをイラレで開いたときのバラバラ文字を、ひとつのエリア内文字に再構成するスクリプト](readme-ja/TextMergeToAreaBox.md)
- [特定の文字のフォントサイズやベースラインの調整](readme-ja/AdjustTextScaleBaseline.md)

## オブジェクトの配置や整列

- [配置したオブジェクトをグリッド状に配置する](readme-ja/SmartObjectDistributor.md)
- [選択したオブジェクトを幅や高さ、不透明度、カラーでソート](readme-ja/SmartObjectSorter.md)
- [オブジェクトを入れ替え](readme-ja/SwapNearestItemWithDialogbox.md)

## 基本図形と変形

- [正方形や正円、正三角形を作成するスクリプト](readme-ja/SmartShapeMaker.md)
- [アスペクト比で変形](readme-ja/AspectRatioScaler.md)

## アートワークやファイル

- [選択したオブジェクトを書き出す](readme-ja/SmartObjectExporter.md)
- [開いているファイルを1つに整列統合](readme-ja/SmartBatchImporter.md)
- [Figmaの⌘ + shift + C（ビットマップとしてコピー）](readme-ja/CopyAsPngLikeFigma.md)

## カラー

- [スウォッチの連続適用](readme-ja/ApplySwatchesToSelection.md)
- [カラーをランダム適用](readme-ja/ShuffleObjectColors.md)

## アートボード

- [アートボード名を最前面のテキストや特定のレイヤーのテキストに設定](readme-ja/SmartArtboardRenamer.md)
- [ページ番号を挿入](readme-ja/AddPageNumberFromTextSelection.md)
- [アートボード名の一括変更](readme-ja/RenameArtboardsPlus.md)
- [カンバス上の並びで［アートボード］パネルの並び順を変更](readme-ja/ReorderArtboardsByPosition.md)
- [アートボード外のオブジェクトを削除](readme-ja/DeleteOutsideArtboard.md)
- [選択オブジェクトに合わせてアートボードサイズを設定するときにマージンを付ける](readme-ja/FitArtboardWithMargin.md)
- [画像を分割してアートボード化](readme-ja/Slice2Artboards.md)

## レイヤー、重ね順

- [オブジェクトを指定レイヤーへ移動](readme-ja/SmartLayerManage.md)
- [選択しているオブジェクトを新規レイヤーに移動し、そのレイヤーを最背面に移動してロック](readme-ja/SendToBgLayer.md)
- [座標を基準に重ね順を変更](readme-ja/SortItemsByPosition.md)
- [選択したグループ内のサブグループを解除して、グループ構造を簡素化](readme-ja/SimplifyGroups.md)

## マスク

- [オブジェクトのまとまりごとにグループ化したり、マスクする](readme-ja/SmartClipAndGroup.md)
- [〈クリッピングマスクを解除〉を拡張](readme-ja/ReleaseClipMask.md)
- [パズル](readme-ja/SmartSliceWithPuzzlify.md)
- [マスクパスのサイズ変更](readme-ja/ResizeClipMask.md)

## ガイド

- [グリッド状にガイドを生成](readme-ja/GenerateGuidesGrid.md)
- [選択したオブジェクトに対してガイドを自動作成](readme-ja/CreateGuidesFromSelection.md)
- [すべてのガイドを削除](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/guide/DeleteAllGuides.jsx)
- [Photoshopの「新規ガイド」をIllustratorでも可能にする](readme-ja/NewGuideMaker.md)

## ドキュメント

- [ドキュメントの切替](readme-ja/SmartSwitchDocs.md)

## その他

- [トンボ作成](readme-ja/AddTrimMark.md)


