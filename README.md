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
- [行のテキストをフォント名とみなして行単位でフォントを適用](readme-ja/ApplyFontByLine.md)

## テキスト関連

- [文字ばらし](readme-ja/TextSplitterPro.md)
- [特定の文字のベースラインシフトを調整](readme-ja/SmartBaselineShifter.md)
- [PDFをイラレで開いたときのバラバラ文字を、ひとつのエリア内文字に再構成するスクリプト](readme-ja/TextMergeToAreaBox.md)
- [特定の文字のフォントサイズやベースラインの調整](readme-ja/AdjustTextScaleBaseline.md)
- [段落を選択](readme-ja/AiSelectParagraph.md)
- [2つのテキストの内容を入れ替える](readme-ja/SwapText.md)
- [テキストと図形をエリア内文字に変換](readme-ja/TextWithShapeToAreaType.md)
- [テキストの分割・結合パレット](readme-ja/TextBreakSplitMergePallete.md)
- [統合文字組みパネル](readme-ja/UnifiedTypePanel.md)
- [テキストのアウトライン化と復元](readme-ja/AiTextOutlineRestorePalette.md)
- [選択したテキストをアーチ状のパス上文字に変換](readme-ja/ArcTextGenerator.md)
- [あふれたテキストの文字サイズ・エリア高さを自動調整](readme-ja/AutoFitTextFrame.md)
- [ポイント文字・パス上文字とエリア内文字の相互変換](readme-ja/ConvertAreaAndPointType.md)
- [テキスト内の日付・曜日・数値を一括で増減](readme-ja/IncrementDatesAndNumbers.md)

## オブジェクトの配置や整列

- [配置したオブジェクトをグリッド状に配置する](readme-ja/SmartObjectDistributor.md)
- [選択したオブジェクトを幅や高さ、不透明度、カラーでソート](readme-ja/SmartObjectSorter.md)
- [オブジェクトを入れ替え](readme-ja/SwapNearestItemWithDialogbox.md)
- [2つのオブジェクトの位置を入れ替え](readme-ja/SwapObjects.md)
- [オブジェクトのリサイズ](readme-ja/SmartObjectResizer.md)

## 基本図形と変形

- [正方形や正円、正三角形を作成するスクリプト](readme-ja/SmartShapeMaker.md)
- [アスペクト比で変形](readme-ja/AspectRatioScaler.md)
- [自由変形（フリーディストート）](readme-ja/SmartFreeDistort.md)
- [パスファインダー](readme-ja/AiSmartPathfinder.md)
- [パスの最適化](readme-ja/PathCleanupTool.md)
- [選択オブジェクトの回転を水平（0°）に補正](readme-ja/ResetRotation.md)

## アートワークやファイル

- [選択したオブジェクトを書き出す](readme-ja/SmartObjectExporter.md)
- [開いているファイルを1つに整列統合](readme-ja/SmartBatchImporter.md)
- [Figmaの⌘ + shift + C（ビットマップとしてコピー）](readme-ja/CopyAsPngLikeFigma.md)
- [リンク画像の管理](readme-ja/LinkedImageManager.md)
- [登録済みAIファイルからスタイル・ブラシ・フォント見本を読み込み](readme-ja/ImportStyles.md)
- [フォルダー内のAI／SVGをバージョン指定で一括保存](readme-ja/AIBatchVersionSave.md)

## カラー

- [スウォッチの連続適用](readme-ja/ApplySwatchesToSelection.md)
- [カラーをランダム適用](readme-ja/ShuffleObjectColors.md)
- [カラーを配色](readme-ja/AiApplySwatchesToSelection.md)
- [選択オブジェクトの色からスウォッチとグラデーションを自動生成](readme-ja/CreateGradientFromSelection.md)
- [塗りと線の入れ替え・変換・消去](readme-ja/FillStrokeSwitcher.md)

## アートボード

- [アートボード名を最前面のテキストや特定のレイヤーのテキストに設定](readme-ja/SmartArtboardRenamer.md)
- [ページ番号を挿入](readme-ja/AddPageNumberFromTextSelection.md)
- [アートボード名の一括変更](readme-ja/RenameArtboardsPlus.md)
- [カンバス上の並びで［アートボード］パネルの並び順を変更](readme-ja/ReorderArtboardsByPosition.md)
- [アートボード外のオブジェクトを削除](readme-ja/DeleteOutsideArtboard.md)
- [選択オブジェクトに合わせてアートボードサイズを設定するときにマージンを付ける](readme-ja/FitArtboardWithMargin.md)
- [画像を分割してアートボード化](readme-ja/Slice2Artboards.md)
- [アートボードサイズ変更](readme-ja/AiArtboardScaler.md)
- [アートボードナビゲーター](readme-ja/ArtboardNavigator.md)
- [選択オブジェクトを各アートボードの9点へ整列](readme-ja/AlignToArtboards.md)
- [アートボード名を解析して行列グリッドに再配置](readme-ja/GridArrangeArtboards.md)

## レイヤー、重ね順

- [オブジェクトを指定レイヤーへ移動](readme-ja/SmartLayerManage.md)
- [選択しているオブジェクトを新規レイヤーに移動し、そのレイヤーを最背面に移動してロック](readme-ja/SendToBgLayer.md)
- [座標を基準に重ね順を変更](readme-ja/SortItemsByPosition.md)
- [選択したグループ内のサブグループを解除して、グループ構造を簡素化](readme-ja/SimplifyGroups.md)
- [水平方向にグループ化](readme-ja/SmartAutoGroup-yoko.md)
- [アートボード単位でオブジェクトをレイヤーに整理](readme-ja/ArtboardLayerOrganizer.md)

## マスク

- [オブジェクトのまとまりごとにグループ化したり、マスクする](readme-ja/SmartClipAndGroup.md)
- [〈クリッピングマスクを解除〉を拡張](readme-ja/ReleaseClipMask.md)
- [パズル](readme-ja/SmartSliceWithPuzzlify.md)
- [マスクパスのサイズ変更](readme-ja/ResizeClipMask.md)
- [クリップグループのマスクと内容を調整](readme-ja/ClipMaskAdjust.md)

## ガイド

- [グリッド状にガイドを生成](readme-ja/GenerateGuidesGrid.md)
- [選択したオブジェクトに対してガイドを自動作成](readme-ja/CreateGuidesFromSelection.md)
- [すべてのガイドを削除](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/guide/DeleteAllGuides.jsx)
- [Photoshopの「新規ガイド」をIllustratorでも可能にする](readme-ja/NewGuideMaker.md)
- [アートボード基準のガイド作成とルーラーガイドの変換](readme-ja/AiCreateArtboardGuides.md)

## ドキュメント

- [ドキュメントの切替](readme-ja/SmartSwitchDocs.md)
- [ファイル名を変更して保存](readme-ja/Ai-FileNameManager.md)

## 環境設定

- [クイック環境設定](readme-ja/AiQuickPrefsPalette.md)
- [環境設定：変形と整列](readme-ja/PreferenceManagerForTransformAndAlign.md)

## その他

- [トンボ作成](readme-ja/AddTrimMark.md)
- [テキストの文字列でグラフィックスタイルを登録](readme-ja/RegisterGraphicStyleWithText.md)
- [選択オブジェクトを条件で絞り込み再選択](readme-ja/SmartSelectionFilter.md)
