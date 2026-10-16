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
- [使用中のフォントをまとめて別のフォントに置換](readme-ja/ReplaceDocumentFonts.md)
- [カテゴリ別ウェイト順にフォントを一覧表示し、フォント見本を一瞬で作成する](readme-ja/TypefaceSampler.md)
- [ドキュメントで使用されているフォント情報を書き出す](readme-ja/ExportFontInfoFromXMP.md)
- [行のテキストをフォント名（＋サイズ・行送り）とみなして行単位で適用](readme-ja/ApplyFontWithFontsize.md)
- [フォントの種別（文字セット・P・UD・N・NT・ウェイト）をまとめて変更するスクリプト](readme-ja/AiFontConverter.md)
- [選択したテキストをフォント情報に変換するIllustrator用スクリプトです](readme-ja/ConvertFontInfo.md)
- [ドキュメント内で使用しているテキストの組み合わせ（フォント・サイズ・行送り・](readme-ja/DocumentFontListSelectorPalette.md)
- [入力したテキストを、インストールされているフォントで表示](readme-ja/FontSampler.md)
- [和文・かな・欧文のフォントを指定して合成フォントを作成](readme-ja/CompositeFontMaker.md)
- [合成フォントを構成フォントに置き換え](readme-ja/DeCompositeFontMaker.md)
- [選択テキストのフォントを一つ上／下のウェイトに切り替え](readme-ja/FontWeightUp.md)
- [似ているフォントのウェイトを見本で見比べて適用](readme-ja/FontWeightPicker.md)
- [よく使うフォントだけを一覧から選んで適用](readme-ja/FavoriteFontPickerPalette.md)
- [インストール済みフォントの見本を自動生成](readme-ja/FontCatalogGenerator.md)
- [ドキュメント内の段落書式を一覧にして、まとめて変更・置換](readme-ja/FontPresetCollector.md)
- [文字設定を「定番」として登録し、選択テキストへ一括適用](readme-ja/FontPresetPickerPalette.md)


## テキスト関連

- [文字ばらし](readme-ja/SmartTextSplitter.md)
- [特定の文字のベースラインシフトを調整](readme-ja/SmartBaselineShifter.md)
- [PDFをイラレで開いたときのバラバラ文字を、ひとつのエリア内文字に再構成するスクリプト](readme-ja/TextMergeToAreaBox.md)
- [特定の文字のフォントサイズやベースラインの調整](readme-ja/AdjustTextScaleBaseline.md)
- [段落を選択](readme-ja/AiSelectParagraph.md)
- [テキストと図形をエリア内文字に変換](readme-ja/TextWithShapeToAreaType.md)
- [テキスト処理パレット](readme-ja/TextProcessingPalette.md)
- [統合文字組みパネル](readme-ja/UnifiedTypePalette.md)
- [テキストのアウトライン化と復元](readme-ja/AiTextOutlineRestorePalette.md)
- [あふれたテキストの文字サイズ・エリア高さを自動調整](readme-ja/FitAreaText.md)
- [エリア内文字に自動サイズ調整をワンクリックで適用](readme-ja/AutoFitTextFrame.md)
- [ポイント文字・パス上文字とエリア内文字の相互変換](readme-ja/ConvertAreaAndPointType.md)
- [あふれを自動サイズ調整で解消してからポイント文字に変換](readme-ja/FitAndConvertToPointType.md)
- [テキスト内の日付・曜日・数値を一括で増減](readme-ja/IncrementDatesAndNumbers.md)
- [テキスト内の数字・英字を増分しながら複製](readme-ja/SmartIncrementText.md)
- [数字・英字・漢数字をソートして連番を振り直す](readme-ja/SmartRenumber.md)
- [エリア内文字ツールキット（作成と調整）](readme-ja/AreaTypeToolkit.md)
- [クリップボードの内容で選択を置換（テキスト以外にも対応）](readme-ja/ReplaceWithPaste.md)
- [クリップボードの複数行テキストを1行ずつ順に流し込む](readme-ja/ReplaceTextWithPasteSequential.md)
- [テキストの検索・置換・整形](readme-ja/SmartTextFindReplace.md)
- [フォントサイズと水平比率／垂直比率を調整](readme-ja/AdjustFontSize.md)
- [選択したテキストフレームの各行の先頭に、箇条書き記号または連番を付与します](readme-ja/AddBulletsAndNumbers.md)
- [Illustrator 用のメモ入力フローティングパレット](readme-ja/AiMemoPalette.md)
- [選択したテキストの各文字にベースライン・比率・回転・カーニングをランダムに付与](readme-ja/AutoTouchType.md)
- [円（パス）とテキストを 1 つずつ選択し、テキストを指定回数繰り返して、円を複製したパス上の文字に変換する](readme-ja/CirclePathTextRepeat.md)
- [ポイント文字・パス上文字・図形・エリア内文字を対象に、エリア内文字の作成と調整を行うツール](readme-ja/ConvertToAreaTypeLikeButton.md)
- [選択した「パス上文字（Path Text）」を、同じ文字内容・段落属性・文字属性をできるだけ維持したまま「ポ…](readme-ja/DetachPathText.md)
- [選択オブジェクトの外接バウンディングボックスを基準に、オフセットを加えた長方形を生成](readme-ja/DrawRectangleBehindSelectedObject.md)
- [選択テキスト内で文字サイズが混在しているとき、各テキストの先頭文字のサイズへ統一する](readme-ja/FontSizeToScaleConverter.md)
- [テキスト内の4桁以上の数値に、3桁ごとの桁区切りのカンマを付けます](readme-ja/FormatNumberWithCommas.md)
- [ポイント文字・パス上文字・図形＋テキストを、見た目を保ったままエリア内文字へ変換する](readme-ja/ImportAndApplyGraphicStylePalette.md)
- [エリア内文字をポイント文字に変換し、改行ごとに分割して個別に再配置するIllustrator用スクリプトです](readme-ja/MimicDynamicText.md)
- [ポイント文字／パス上文字を、用途に応じて「作成」「分離」「調整」できるツールです](readme-ja/PathTextToolkit.md)
- [選択したテキストを段落ごとに分け、指定したアートボードから順に配置](readme-ja/SplitTextToArtboards.md)
- [選択中の2つのテキストオブジェクトの内容を入れ替えます](readme-ja/SwapTextSpecial.md)
- [複数のテキストオブジェクトを1つのエリア内文字に連結します](readme-ja/TextMergeToAreaBox-tab.md)
- [選択したテキストの基本的な文字組み設定（フォントサイズと行送り・自動カーニング・文字ツメ・](readme-ja/TypeBasicsPalette.md)
- [基準フォントサイズと倍率からタイプスケールを自動生成](readme-ja/TypeScaler.md)
- [選択中のテキストフレームを、指定した文字の直後で改行する](readme-ja/titlemaker.md)
- [指定文字のベースラインを自動的に天地中央にする](readme-ja/AdjustBaselineVerticalCenter.md)
- [行頭の文字サイズを基準に行送りを再計算](readme-ja/ApplyLeadingPerTextFramePalette.md)
- [エリア内文字を天地・左右とも中央にそろえる](readme-ja/AreaTypeCenterMiddle.md)
- [ドキュメント内の日付を検索して置換](readme-ja/DateFindReplace.md)
- [選択したテキストをアーチ・円・弓のパス上文字に変換、または行の幅をそろえる](readme-ja/DynamicTextGenerator.md)
- [複数のテキストをエリア内文字にまとめる／分割する](readme-ja/MultiAreaText.md)
- [テキストを一覧で編集してドキュメントに書き戻す](readme-ja/TextScopeEdit.md)
- [複数の条件でテキストフレームを一括選択](readme-ja/TextSelector.md)
- [縦組みの欧文回転・縦中横・数字の書き換えをまとめて適用](readme-ja/VerticalTextFormatter.md)
- [縦組み用のテキストを横組み用に戻す](readme-ja/HorizontalTextFormatter.md)
- [全角・半角の英数字と等幅字形を色で塗り分け](readme-ja/ColorNumberByWidth.md)
- [選択したテキストの体裁を段落スタイルとして登録](readme-ja/%E6%AE%B5%E8%90%BD%E3%82%B9%E3%82%BF%E3%82%A4%E3%83%AB.md)


## オブジェクトの配置や整列

- [配置したオブジェクトをグリッド状に配置する](readme-ja/SmartObjectDistributor.md)
- [選択したオブジェクトを幅や高さ、不透明度、カラーでソート](readme-ja/SmartObjectSorter.md)
- [オブジェクトを入れ替え](readme-ja/SwapNearestItemWithDialogbox.md)
- [2つのオブジェクトの位置を入れ替え](readme-ja/SwapObjects.md)
- [オブジェクトのリサイズ](readme-ja/SmartObjectResizer.md)
- [隣り合うオブジェクトのアキの中央に記号を配置](readme-ja/RightMarkPlacer.md)
- [選択した2つのオブジェクトの上下の間隔を、指定した値にそろえる常駐パレットです](readme-ja/AiAdjustVerticalGapPalette.md)
- [選択したオブジェクトを、アートボードを対象に整列する常駐パレットです](readme-ja/AiAlignToArtboardPalette.md)
- [選択内容に応じて、行送りと配置を次の順で判定して調整します](readme-ja/DistributeDownFromTop.md)
- [選択内容に応じて、行送りと配置を次の順で判定して調整します](readme-ja/DistributeUpFromTop.md)
- [選択したオブジェクトを、グリッド／行／列／ランダムのいずれかの方式で複製・配置します](readme-ja/DuplicateInGridPlus.md)
- [選択したオブジェクトのグループを入れ子ごとすべて解除し、あらためて1つのグループにまとめます](readme-ja/FlattenGroup.md)
- [複数選択したテキストフレームの背面に K15% の長方形を作成し、「_bg-rectangle」レイヤーに配…](readme-ja/GridTextLayout.md)
- [既存のグループに入れる・グループから出す・1つずつグループにするを、ダイアログで選んで実行](readme-ja/GroupMembership.md)
- [選択している配置画像（PlacedItem / RasterItem）の拡大・縮小率（%）を表示し、入力値で…](readme-ja/ImageScaler.md)
- [選択したオブジェクトの移動・複製と反転・回転を、アイコンのクリックで即時実行する常駐パレットです（プレビュー…](readme-ja/QuickTransformPalette.md)
- [選択したオブジェクトをランダムに移動・変形・回転・不透明度を変更するスクリプト](readme-ja/RandomizeObjects.md)
- [選択中のオブジェクトが「だいたいグリッド状」に並んでいることを前提に、左右・上下の間隔で再配置します](readme-ja/RegridObjects.md)
- [グループ・複合パス・ブレンド・リピートなど、解除できるものを入れ子ごとすべて解除](readme-ja/ReleaseEverything.md)
- [配置画像・テキスト・長方形（パス）・クリップグループ・直線パスに対して、回転／シアー（せん断）／スケール／縦…](readme-ja/ResetTransform.md)
- [重なって配置されたオブジェクトを、横方向または縦方向へ指定した間隔で並べ直します。行数・列数を指定すれば…](readme-ja/SmartAlignAndTile.md)
- [選択オブジェクトを「縦/横」に並べて、指定した間隔で分布します。方向は自動判定も可能で、揃え（左右/上下）、…](readme-ja/SmartAlignDistribute.md)
- [選択オブジェクトを拡大・縮小する常駐パレット](readme-ja/SmartScalePalette.md)
- [ポイント文字およびエリア内文字の［字形の境界に整列］を切り替えながら、垂直方向（上・中央・下）に整列します。…](readme-ja/SmartVerticalAlign.md)
- [グループ内のテキストから数値を抽出し、数値に基づいてグループを並び替え縦方向に整列するIllustrator…](readme-ja/SortByNumbers.md)
- [テキストフレーム内のタブ区切りテキストを指定列の値で並び替えるIllustrator用スクリプトです](readme-ja/SortTextByColumn.md)
- [選択したオブジェクトを基準に、指定方向（右／左／上／下）にある最も近いオブジェクトと自然な見た目で位置を入れ…](readme-ja/SwapNearestItem.md)
- [Illustratorでテキストフレームを行・列単位で整列またはグループ化するスクリプト](readme-ja/TextGridAligner.md)
- [選択した複数オブジェクトを行・列として自動判定し、歯抜け（欠け）を許容しつつ、](readme-ja/TransposeGrid.md)
- [オブジェクトの重ね順を位置（X/Y）やZインデックスで並べ替える](readme-ja/ZIndexSorter.md)
- [選択した 2 つのオブジェクトの中心位置を入れ替えます](readme-ja/swap-2-objects.md)
- [基準にする側を決めて、間隔と位置をそろえる](readme-ja/AdjustPairGap.md)
- [パスに沿ってオブジェクトを等間隔に配置](readme-ja/ArrangeObjectsAlongPath.md)
- [複数のオブジェクトをアートボードの端・中央・ガイドへ整列](readme-ja/GroupEdgeAlign.md)
- [行送りと配置を上下方向に調整するパレット](readme-ja/SmartDistributorPalette.md)


## 基本図形と変形

- [正方形や正円、正三角形を作成するスクリプト](readme-ja/SmartShapeMaker.md)
- [縦横比を指定してサイズ変更](readme-ja/AspectRatioScaler.md)
- [自由変形（フリーディストート）](readme-ja/SmartFreeDistort.md)
- [パスファインダー](readme-ja/AiSmartPathfinderPalette.md)
- [パスの最適化](readme-ja/PathCleanupTool.md)
- [重なりをならす（オフセットパス＋合流）](readme-ja/OverlapRemover.md)
- [チケット・クーポン形状の作成（ミシン目、ギザギザ、スリット／ホール）](readme-ja/CouponTicketMaker.md)
- [選択オブジェクトの回転を水平（0°）に補正](readme-ja/ResetRotation.md)
- [選択オブジェクトを複製 → 背面配置 → オフセットパス（Live Effect）→ アウトライン → 合体…](readme-ja/AddOutlineOffsetPath.md)
- [選択オブジェクトの全アンカーポイントに、マーカー（正方形／最前面オブジェクト／シンボル）を配置するユーティリ…](readme-ja/AiAnchorPointMarker.md)
- [Excel などからコピー＆ペーストした長方形から罫線（グリッド・枠線）を自動生成](readme-ja/CenterLineConnectorFromRect.md)
- [RGB または CMYK で構成された黒を安定した K100 の黒に変換する Illustrator 用スク…](readme-ja/ColorToK100Converter.md)
- [選択オブジェクトの境界に合わせて長方形を作成する Illustrator スクリプト](readme-ja/ConvertToRectangle.md)
- [選択したオブジェクトにランダムな変形を加え、手書き・スケッチ風の見た目に整えます](readme-ja/KPTSketchyPalette.md)
- [距離・角度・スケールを指定してロングシャドウを生成](readme-ja/LongShadowMaker.md)
- [選択中のオブジェクトに対して複合パス解除 → パスの合体 → アピアランス拡張 → グループ解除 → オフセ…](readme-ja/PathUniteOffsetTool.md)
- [選択した長方形を、左下隅・上辺中央・右下隅の 3 点を通る円弧に変換](readme-ja/RectangleToArc.md)
- [角丸長方形の角の半径を変更](readme-ja/EditCornerRadius.md)
- [テキストやオブジェクトの背面に図形を敷く](readme-ja/AddBackdrop.md)
- [新聞の見出しのような画像を作る](readme-ja/ShimbunTitleMaker.md)
- [波括弧（カーリーブラケット）をプレビューしながら作成](readme-ja/AiCurlyBracketMaker.md)
- [ブレンドの作成と調整](readme-ja/BlendSp.md)
- [選択全体の外接矩形をひとつの長方形に](readme-ja/BoundsToRectangle.md)
- [アピアランスを消去して塗り・線だけ残す](readme-ja/ClearAppearance.md)
- [破線を1本ずつの線分に分割](readme-ja/DashNipper.md)
- [グラデーションを単色オブジェクトに分割](readme-ja/ExpandGradient.md)
- [パターンの塗りを通常のパスに展開](readme-ja/ExpandPattern.md)
- [テキストの背面に座布団を作る](readme-ja/FitShapeToContent.md)
- [テキストを原稿用紙に組む](readme-ja/GenkoYoshiMaker.md)
- [線を塗りに変換してから合流して分割](readme-ja/MergeExpand.md)
- [選択したオブジェクトから枝分かれするサンキー図を作る](readme-ja/SankeyFlowMaker.md)
- [雨のような罫線を生成](readme-ja/VTRainMaker.md)


## アートワークやファイル

- [選択したオブジェクトを書き出す](readme-ja/SmartObjectExporter.md)
- [開いているファイルを1つに整列統合](readme-ja/SmartBatchImporter.md)
- [Figmaの⌘ + shift + C（ビットマップとしてコピー）](readme-ja/CopyAsPngLikeFigma.md)
- [リンク画像の管理](readme-ja/LinkedImageManagerPalette.md)
- [埋め込み画像をリンク画像に変換（アクション方式）](readme-ja/UnembedToLinks.md)
- [登録済みAIファイルからスタイル・ブラシ・フォント見本を読み込み](readme-ja/ImportStyles.md)
- [フォルダー内のAI／SVGをバージョン指定で一括保存](readme-ja/AIBatchVersionSave.md)
- [CSV／タブ区切りデータをテンプレートに流し込むデータ結合](readme-ja/VariableDataImport.md)
- [クリップボードの申し込み内容から領収書・請求書・納品書のPDFを作成](readme-ja/InvoiceFromClipboard.md)
- [選択オブジェクトを高解像度でラスタライズし、PNG相当のビットマップとしてクリップボードにコピーするIllu…](readme-ja/CopyAsPngLikeFigmaWithDialog.md)
- [登録されているアクションセットを、デスクトップの `Illustrator_Actions` フォルダーへ…](readme-ja/ExportActions.md)
- [PDF/AI ファイルを指定したページ範囲で読み込み、現在のドキュメント上にページを配置します](readme-ja/PDFAIImporter.md)
- [同じリンクファイルを参照する配置画像を一括で差し替え](readme-ja/RelinkSameImages.md)
- [現在選択中のリンク画像（PlacedItem）と同じリンクファイルを参照する](readme-ja/SelectSameLinks.md)
- [Illustrator ドキュメントに登録されたシンボルを一覧表示する専用アートボード「シンボル一覧」を自動…](readme-ja/SymbolListBuilder.md)
- [選択オブジェクトをシンボルとして登録し、ドキュメント内の一致するオブジェクトをそのインスタンスに置き換え](readme-ja/SymbolizeAndReplace.md)
- [選択したオブジェクトをシンボルとして登録](readme-ja/SymbolizeEach.md)
- [アクティブドキュメントの全アートボードを、名前ごとのルールで PNG 書き出しします](readme-ja/export-Event.md)
- [ダイアログで選んだアートボードを PNG・JPEG・PDF で書き出します](readme-ja/ArtboardExporter.md)
- [アクティブなアートボードを PNG24 形式で書き出します](readme-ja/export200.md)
- [指定したファイルを開いて PDF として保存し、閉じる](readme-ja/OpenSaveAsPDFAndClose.md)
- [複数のフォルダーを横断して .ai/.svg ファイルを探す](readme-ja/AiFileFinder.md)
- [同じシンボルのインスタンスをすべて選択](readme-ja/FindAllSymbolInstances.md)
- [配置画像の埋め込み・解除・リンクをまとめて扱う](readme-ja/ImageLinkManager.md)
- [リンク切れの画像を同名ファイルで自動再リンク](readme-ja/RelinkMissingLinks.md)
- [見開きを左右2つの片ページに分割](readme-ja/SplitSpreadToSingle.md)
- [選択オブジェクトを一覧から選んだシンボルに置き換え](readme-ja/%E3%82%B7%E3%83%B3%E3%83%9C%E3%83%AB%E3%81%AB%E7%BD%AE%E3%81%8D%E6%8F%9B%E3%81%88-sw.md)


## カラー

- [画像やオブジェクトからカラーパレットを作成](readme-ja/ColorPaletteFromImage.md)
- [スウォッチの連続適用](readme-ja/ApplySwatchesToSelection.md)
- [カラーをランダム適用](readme-ja/ShuffleObjectColors.md)
- [カラーを配色](readme-ja/AiApplySwatchesToSelection.md)
- [選択オブジェクトの色からスウォッチとグラデーションを自動生成](readme-ja/CreateGradientFromSelection.md)
- [塗りと線の入れ替え・変換・消去](readme-ja/FillStrokeSwitcher.md)
- [不透明度を塗りのカラーに焼き込む](readme-ja/FlattenOpacityPro.md)
- [選択オブジェクト（閉じたパス、テキスト）の塗りおよび線のカラーをスウォッチ（スポットカラー）として登録し、そ…](readme-ja/RegisterAndApplySwatches.md)
- [選択オブジェクトのグラデーションを、指定した数のカラーストップでセパレートグラデーション（縞模様）に変換](readme-ja/Convert2SeparateGradient.md)
- [カラーパレットを生成してスウォッチに登録](readme-ja/ColorGenerator.md)
- [明るさをスライダーで調整](readme-ja/ColorToneSlider.md)
- [元の塗り色を始点にした線形グラデーションを作成](readme-ja/GradientFromFill.md)
- [配置順で塗り色を抽出してスウォッチグループに登録](readme-ja/SwatchGroupFromSelection.md)


## アートボード

- [アートボード名を最前面のテキストや特定のレイヤーのテキストに設定](readme-ja/SmartArtboardRenamer.md)
- [ページ番号を挿入](readme-ja/AddPageNumberFromTextSelection.md)
- [アートボード名の一括変更](readme-ja/RenameArtboardsPlus.md)
- [カンバス上の並びで［アートボード］パネルの並び順を変更](readme-ja/ReorderArtboardsByPosition.md)
- [アートボード外のオブジェクトを削除](readme-ja/DeleteOutsideArtboard.md)
- [オブジェクトに合わせてアートボードのサイズを調整](readme-ja/FitArtboardWithMargin.md)
- [画像を分割してアートボード化](readme-ja/Slice2Artboards.md)
- [アートボードサイズ変更](readme-ja/AiArtboardScaler.md)
- [アートボードナビゲーター](readme-ja/ArtboardNavigatorPalette.md)
- [選択オブジェクトを各アートボードの9点へ整列](readme-ja/AlignToArtboards.md)
- [アートボード名を解析して行列グリッドに再配置](readme-ja/GridArrangeArtboards.md)
- [すべてのアートボードに同じサイズの矩形を描画し、アートボード内のオブジェクトをマスクします](readme-ja/ArtboardMaskAndRelease.md)
- [全アートボード上にあるテキストフレームを収集し、最後のアートボードの右側に縦に並べて配置する](readme-ja/CollectArtboardTexts.md)
- [選択したグループオブジェクトの境界に指定したマージンを加え、その範囲をアートボードとして自動追加するIllu…](readme-ja/Group2Artboards.md)
- [アクティブなアートボードだけを残す、または空のアートボードを削除](readme-ja/KeepActiveArtboardOnly.md)
- [ダイアログで指定した「幅」「高さ」に、アートボードをライブプレビューしながら変形します](readme-ja/ResizeArtboardsAll.md)
- [アクティブまたは全アートボードと同サイズの長方形を、オフセットを考慮して描画します](readme-ja/SmartDrawArtboardRectangle.md)
- [開いているすべてのドキュメントへ同じ位置にペースト](readme-ja/PasteInAllDocs.md)
- [並びを引き継いでアートボードを追加](readme-ja/AddArtboardPlus.md)
- [アートボード全体が正方形に近づくよう再配置](readme-ja/ArrangeArtboardsToSquareGrid.md)
- [アートボードをランダムな順序でグリッドに並べ替え](readme-ja/RandomizeArtboards.md)


## レイヤー、重ね順

- [オブジェクトを指定レイヤーへ移動](readme-ja/SmartLayerManage.md)
- [選択しているオブジェクトを新規レイヤーに移動し、そのレイヤーを最背面に移動してロック](readme-ja/SendToBgLayer.md)
- [座標を基準に重ね順を変更](readme-ja/SortItemsByPosition.md)
- [選択したグループ内のサブグループを解除して、グループ構造を簡素化](readme-ja/SimplifyGroups.md)
- [アートボード単位でオブジェクトをレイヤーに整理](readme-ja/ArtboardLayerOrganizer.md)
- [実行前にダイアログを表示し、処理条件を選択可能](readme-ja/FlattenLayers.md)
- [アクティブレイヤーの「テンプレート」属性（ロック・印刷不可・画像を薄く表示）を ON / OFF する](readme-ja/ToggleTemplateLayer.md)
- [レイヤーの階層・ロック・表示状態を一覧表示](readme-ja/InspectLayerTree.md)
- [「Guides Preview for Trim View」レイヤーを削除](readme-ja/RemoveTrimViewGuideLayer.md)
- [指定したレイヤーをテンプレート化して接頭辞を付ける](readme-ja/SetTemplateLayer.md)
- [アートボードと同じ大きさの背景をテンプレート化して最背面へ](readme-ja/bg-template.md)
- [「Guides Preview for Trim View」レイヤーをテンプレート化](readme-ja/%E3%82%AC%E3%82%A4%E3%83%89%E3%82%92%E3%83%86%E3%83%B3%E3%83%97%E3%83%AC%E3%83%BC%E3%83%88%E3%81%AB.md)
- [「下絵」レイヤーをテンプレート化](readme-ja/%E3%83%AC%E3%82%A4%E3%83%A4%E3%83%BC%E3%82%92%E3%83%86%E3%83%B3%E3%83%97%E3%83%AC%E3%83%BC%E3%83%88%E3%81%AB.md)


## マスク

- [オブジェクトのまとまりごとにグループ化したり、マスクする](readme-ja/SmartClipAndGroup.md)
- [〈クリッピングマスクを解除〉を拡張](readme-ja/ReleaseClipMask.md)
- [パズル](readme-ja/SmartSliceWithPuzzlify.md)
- [マスクパスのサイズ変更](readme-ja/ResizeClipMask.md)
- [クリップグループのマスクと内容を調整](readme-ja/ClipMaskAdjust.md)
- [マスクした複製で、ぼかし・ルーペ・部分拡大をつくる](readme-ja/MaskSpotlight.md)
- [画像を省略線でトリミング](readme-ja/TrimWithBreakLine.md)
- [マスク編集モードへの切り替えとクリッピングマスク作成](readme-ja/ClipEditMask.md)
- [クリップ形状を正方形・正円・六角形に置き換え](readme-ja/ClipMaskShapeChanger.md)
- [複数の図形と画像をまとめてクリッピングマスク化](readme-ja/ImgFitMaskMultiple.md)
- [クリッピングマスクの作成と解除を切り替える](readme-ja/MakeClippingMask.md)


## ガイド

- [グリッド状にガイドを生成](readme-ja/GenerateGuidesGrid.md)
- [囲み罫とグリッドを一括生成](readme-ja/SmartGridMaker.md)
- [選択したオブジェクトに対してガイドを自動作成](readme-ja/CreateGuidesFromSelection.md)
- [Photoshopの「新規ガイド」をIllustratorでも可能にする](readme-ja/NewGuideMaker.md)
- [アートボード基準のガイド作成とルーラーガイドの変換](readme-ja/AiCreateArtboardGuides.md)
- [複数のレイヤー／サブレイヤーに散在するガイドを、1 つのレイヤー（既定は `// guide`）へ集約します](readme-ja/CollectGuides.md)
- [ドキュメント内のガイド（ルーラーガイド含む）の交点を基準に長方形を自動生成します](readme-ja/MakeRectangleFromGuides.md)
- [「_guide」レイヤーのガイドをパスに戻して移動](readme-ja/ReleaseGuidesAsPaths.md)
- [ガイドを線付きのパスに変換してコピー](readme-ja/CopyGuidesAsPaths.md)


## ドキュメント

- [ドキュメントの切替](readme-ja/SmartSwitchDocs.md)
- [ファイル名を変更して保存](readme-ja/AiFileNameManager.md)
- [選択オブジェクトのみを残した複製ドキュメントを作成するInDesign用スクリプトです](readme-ja/CloneDocSelectedOnly.md)
- [複数ドキュメントの選択オブジェクトを同じ座標に揃える](readme-ja/SyncSelectionPosition.md)


## 環境設定

- [よく使う環境設定を常駐パレットで切り替える](readme-ja/AiQuickPrefsPalette.md)
- [ビデオ定規・境界線などの表示を常駐パレットで切り替える](readme-ja/ViewTogglePalette.md)
- [アートボード関連の環境設定](readme-ja/ArtboardDisplayPresetManagerPalette.md)
- [環境設定：変形と整列](readme-ja/PreferenceManagerForTransformAndAlignPalette.md)
- [環境設定をまとめて変更](readme-ja/PresetManager.md)
- [Illustrator の各種環境設定をダイアログボックスから変更可能にします](readme-ja/PreferenceManager-unit.md)
- [環境設定をまとめて変更（ダイアログなし）](readme-ja/PresetManagerNoDialog.md)
- [表示とスナップのオン・オフをまとめて切り替え](readme-ja/ViewToggles.md)


## その他

- [ドキュメントの不要な要素をまとめて削除](readme-ja/AiDocumentCleaner.md)
- [トンボ作成](readme-ja/AddTrimMark.md)
- [テキストの文字列でグラフィックスタイルを登録](readme-ja/RegisterGraphicStyleWithText.md)
- [選択オブジェクトを条件で絞り込み再選択](readme-ja/SmartSelectionFilter.md)
- [選択中とドキュメント全体のオブジェクト数を集計](readme-ja/SelectionInspectorPalette.md)
- [選択オブジェクトから新規レイヤー・アートボード・ドキュメントを作成](readme-ja/SelectionToNew.md)
- [シンボルへのリンクを解除して整理](readme-ja/AiBreakLink.md)
- [開いているドキュメントのアクティブビューの回転角度（表示の回転）を取得](readme-ja/AiSmartRotateViewPalette.md)
- [常駐エンジンで動いている各種フローティングパレットをまとめて閉じるユーティリティ](readme-ja/CloseAllPalettes.md)
- [環境設定の「角度の制限」（Shiftキーを押したときの角度）と「キー増加」（矢印キーの移動量）の変更、](readme-ja/DirectPrefsPalette.md)
- [選択したオブジェクトを 1 つずつ選び直しながら、それぞれにアピアランスの分割を適用します](readme-ja/ExpandAppearanceEachObject.md)
- [ダイアログの「スタイルを読み込み」ボタンで AI ファイルを指定し、そのファイル内のグラフィックスタイルを取…](readme-ja/ImportGraphicStyles.md)
- [表示領域の中心にエリア内文字を作成し、サンプルテキストを流し込む](readme-ja/InsertNewAreaText.md)
- [表示領域の中心に黒い正方形を作成し、シェイプに変換とピクセルグリッドに最適化を適用](readme-ja/InsertNewRectangle.md)
- [表示領域の中心付近に黒い正方形を5つ、重ならないように配置](readme-ja/InsertNewRectangle5Times.md)
- [表示領域の中心に欧文のポイントテキストを作成](readme-ja/InsertNewTextE.md)
- [表示領域の中心に長めの欧文ポイントテキストを作成](readme-ja/InsertNewTextELong.md)
- [表示領域の中心に和文のポイントテキストを作成](readme-ja/InsertNewTextJ.md)
- [ロックしたオブジェクトを履歴に残して戻す](readme-ja/LockHistoryPalette.md)
- [グラフィックスタイル／ブラシ／スウォッチ／シンボルの「名前」を、ダイアログで指定した「検索→置換」で一括変更…](readme-ja/RenameAssets.md)
- [アクティブなドキュメントに、選んだ .ai / .pdf のアートボード（PDF はページ）をグリッド状に並…](readme-ja/SlideCollage.md)
- [基準日をもとに、指定した月数ぶんのカレンダー（月曜はじまり）をアートボード中心に作成します](readme-ja/SmartCalendarMaker.md)
- [Illustrator のアートボード／シンボル／レイヤー名を、接頭辞・接尾辞・名前ソース・検索置換を組み合…](readme-ja/SmartRenamer.md)
- [2つのオブジェクト（テキスト、パス、グループなど）を選択して実行すると、各オブジェクトの背面に2分割の背景を…](readme-ja/SplitBackgroundForTwo.md)
- [1つのオブジェクト（テキスト、パス、グループなど）を選択して実行すると、そのオブジェクトの外接矩形を左右また…](readme-ja/SplitForTwo.md)
- [選択した外枠の長方形を基準に、内部の縦罫／横罫を等間隔に再配置します](readme-ja/TableRuleAverager.md)
- [Illustrator の選択テキストや全体の文字情報を統計的に可視化](readme-ja/TextCountStatsPalette.md)
- [選択オブジェクトに合わせて、アクティブビューをズーム＆センタリングします](readme-ja/ZoomToSelection.md)
- [選択中のオブジェクトの見た目を「名称未設定」で登録する Illustrator 標準の挙動を回避し、固定名の…](readme-ja/register-temp-style.md)
- [アートボード／レイヤー／シンボル／グラフィックスタイルの名前を、検索置換とナンバリングでまとめて変更します](readme-ja/renamer.md)
- [メニューコマンド・ツール・環境設定のコードを引く](readme-ja/AiCommandPrefLookup.md)
- [キーオブジェクトから各図形へコネクターを引く](readme-ja/AiConnectorBuilder.md)
- [スクリプトをキーワードで絞り込んで実行](readme-ja/AiScriptLauncher.md)
- [保存した文字属性をテキストに適用](readme-ja/ApplyTextAttributesFromClipboard.md)
- [外枠の中で罫線を均等に再配置](readme-ja/AverageRulesInRectangle.md)
- [選択オブジェクトの領域に紙吹雪を散らす](readme-ja/ConfettiMaker.md)
- [文字属性・段落属性を取得して保存](readme-ja/CopyTextAttributesToClipboard.md)
- [破線計算機（DashGapCalculator）](readme-ja/DashGapCalculator.md)
- [オブジェクトの間に水平の罫線を描く](readme-ja/DrawLinesBetween.md)
- [Excel由来のデータを表組みとして整形](readme-ja/ExcelTableNormalizer.md)
- [隣接アンカーを結ぶ直線を描画範囲いっぱいに延長](readme-ja/ExtendLines.md)
- [対象のバウンディングボックスを最寄りの基準線に合わせる](readme-ja/FillSnapper.md)
- [ロゴ設計や解析に役立つ補助線を生成](readme-ja/GuideLineBuilder.md)
- [外接矩形から指定角度の引き出し線を作成](readme-ja/LeaderLineBuilder.md)
- [ロゴ用の補助線とクリアスペースを生成](readme-ja/LogoGridMaker.md)
- [容器の中にタイルを等間隔で敷き詰める](readme-ja/PatternFill.md)
- [長方形を矢印に置き換え](readme-ja/RectangleToArrow.md)
- [水平線・垂直線を長方形グリッドとして再構成](readme-ja/RectangularGridReverseTool.md)
- [奇数番目／偶数番目だけを互い違いに選択](readme-ja/SelectAlternateItems.md)
- [複数のオブジェクトを基準に「共通 > アピアランス」を実行](readme-ja/SelectSameAppearanceMulti.md)
- [よく使う矢印と線の設定をまとめて適用](readme-ja/SmartStrokeSettings.md)
- [2つのオブジェクトの見た目または文字列を交換](readme-ja/SwapStyle.md)
- [テキストフレームの見た目に合わせて罫線と背景を生成](readme-ja/TableMaker.md)
- [テキストの周りで罫線が欠けたタイトル帯を作成](readme-ja/TitleBarLineCut.md)
- [選択の周りで罫線を欠けさせる](readme-ja/TitleBarLineCutSp.md)
- [トリミング表示を切り替えてガイドを表示](readme-ja/ToggleTrimView.md)
- [テキストのツリー記号をパスの罫線に変換](readme-ja/TreeSymbolToStroke.md)
- [グラフィックスタイル「temp_style」を適用](readme-ja/apply-temp-style-or-last.md)
- [選択を表として解釈して罫線と塗りを生成](readme-ja/tabularize.md)

