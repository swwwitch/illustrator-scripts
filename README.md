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
- [フォントの種別（文字セット・P・UD・N・NT・ウエイト）をまとめて変更するスクリプト](readme-ja/AiFontConverter.md)
- [選択したテキストをフォント情報に変換するIllustrator用スクリプトです](readme-ja/ConvertFontInfo.md)
- [ドキュメント内で使用しているテキストの組み合わせ（フォント・サイズ・行送り・](readme-ja/DocumentFontListSelector.md)
- [入力したテキストを、インストールされているフォントで表示](readme-ja/FontSampler.md)
- [Illustrator で利用可能なフォントを一覧表示し、ウエイトやスタイル順にアートボード上へ整列描画する…](readme-ja/TypefaceSampler-text.md)


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
- [カーニング設定パレット](readme-ja/AutoKerningPanel.md)
- [テキストのアウトライン化と復元](readme-ja/AiTextOutlineRestorePalette.md)
- [選択したテキストをアーチ状のパス上文字に変換](readme-ja/ArcTextGenerator.md)
- [あふれたテキストの文字サイズ・エリア高さを自動調整](readme-ja/AutoFitTextFrame.md)
- [ポイント文字・パス上文字とエリア内文字の相互変換](readme-ja/ConvertAreaAndPointType.md)
- [テキスト内の日付・曜日・数値を一括で増減](readme-ja/IncrementDatesAndNumbers.md)
- [エリア内文字ツールキット（作成と調整）](readme-ja/AreaTypeToolkit.md)
- [クリップボードのテキストで選択テキストを置換](readme-ja/ReplaceTextWithPaste.md)
- [クリップボードの複数行テキストを1行ずつ順に流し込む](readme-ja/ReplaceTextWithPasteSequential.md)
- [フォントサイズと水平比率／垂直比率を調整](readme-ja/AdjustFontSize.md)
- [選択したテキストフレームの各行の先頭に、箇条書き記号または連番を付与します](readme-ja/AddBulletsAndNumbers.md)
- [選択している文字を対象に、フォントサイズと水平比率／垂直比率を調整する](readme-ja/AdjustFontSizePallete.md)
- [Illustrator 用のメモ入力フローティングパレット](readme-ja/AiMemoPallete.md)
- [選択中の「ポイント文字」を、同時に選択しているパス（PathItem）上の文字に変換します](readme-ja/AttachTextToPath.md)
- [選択したテキストの自動カーニング方式（和文等幅／0／メトリクス／オプティカル）を](readme-ja/AutoKerning.md)
- [選択したテキストの自動カーニング方式を「メトリクス」に設定する](readme-ja/AutoKerning-Metrics.md)
- [選択したテキストの自動カーニング方式を「和文等幅（欧文のみメトリクス）」に設定する](readme-ja/AutoKerning-MetricsRomanOnly.md)
- [選択したテキストの自動カーニング方式を「オプティカル」に設定する](readme-ja/AutoKerning-Optical.md)
- [選択したテキストの自動カーニング方式を「オプティカル」に設定する](readme-ja/AutoKerning-Optical30.md)
- [選択したテキストの自動カーニング方式を「0（カーニングなし）」に設定する](readme-ja/AutoKerning-Zero.md)
- [選択したテキストの自動カーニング方式を「オプティカル」に設定します](readme-ja/AutoKerningWabunSimple.md)
- [選択したテキストの「現在の行送り（絶対値）」とフォントサイズから行送り％を逆算し、](readme-ja/AutoLeadingCalc.md)
- [選択したテキストの行送り（表示値）が整数で 1 ステップ大きくなるように、自動行送り量（％）を](readme-ja/AutoLeadingPlus1.md)
- [行送りを整数 1 ステップ大きくする（自動行送り量で調整）](readme-ja/AutoLeadingStep+1.md)
- [行送りを整数 10 ステップ大きくする（自動行送り量で調整）](readme-ja/AutoLeadingStep+10.md)
- [行送りを整数 1 ステップ小さくする（自動行送り量で調整）](readme-ja/AutoLeadingStep-1.md)
- [行送りを整数 10 ステップ小さくする（自動行送り量で調整）](readme-ja/AutoLeadingStep-10.md)
- [円（パス）とテキストを 1 つずつ選択し、テキストを指定回数繰り返して、円を複製したパス上の文字に変換する](readme-ja/CirclePathTextRepeat.md)
- [ポイント文字・パス上文字・図形・エリア内文字を対象に、エリア内文字の作成と調整を行うツール](readme-ja/ConvertToAreaTypeLikeButton.md)
- [ポイント文字・パス上文字・図形＋テキストを、見た目を保ったままエリア内文字へ変換する](readme-ja/ConvertToAreaTypeLikeButtonAccurateSize.md)
- [選択したテキストを、選択内容に応じて双方向に変換する](readme-ja/ConvertToAreaTypeLikeButtonAccurateSize-v2.md)
- [選択した「パス上文字（Path Text）」を、同じ文字内容・段落属性・文字属性をできるだけ維持したまま「ポ…](readme-ja/DetachPathText.md)
- [選択オブジェクトの外接バウンディングボックスを基準に、オフセットを加えた長方形を生成](readme-ja/DrawRectangleBehindSelectedObject.md)
- [選択テキスト内で文字サイズが混在しているとき、各テキストの先頭文字のサイズへ統一する](readme-ja/FontSizeToScaleConverter.md)
- [ポイント文字・パス上文字・図形＋テキストを、見た目を保ったままエリア内文字へ変換する](readme-ja/ImportAndApplyGraphicStyle.md)
- [選択したテキストフレームの各段落で使われている禁則処理セットの値を集め、一覧をアラートで表示します](readme-ja/InspectKinsokuSimple.md)
- [エリア内文字をポイント文字に変換し、改行ごとに分割して個別に再配置するIllustrator用スクリプトです](readme-ja/MimicDynamicText.md)
- [アウトライン化されたテキストを、メモ情報（note）をもとに元のテキストとして復元するスクリプトです](readme-ja/OutlineTextRestore.md)
- [ポイント文字／パス上文字を、用途に応じて「作成」「分離」「調整」できるツールです](readme-ja/PathTextToolkit.md)
- [ポイント文字／パス上文字を、用途に応じて「作成」「分離」「調整」できるツールです](readme-ja/PathTextToolkit-v2.md)
- [選択中の2つのテキストオブジェクトの内容を入れ替えます](readme-ja/SwapTextSpecial.md)
- [選択したテキストの文字組み設定（フォント・フォントサイズ・自動カーニング・字間・文字揃え・行揃え・行送り・](readme-ja/TextFontPanelReinvented.md)
- [複数のテキストオブジェクトを1つのエリア内文字に連結します](readme-ja/TextMergeToAreaBox-tab.md)
- [選択したテキストフレームの情報を取得してアウトライン化](readme-ja/TextOutlineWithMemo.md)
- [ポイント文字・パス上文字を、見た目（アピアランス）を保ったままエリア内文字へ変換する](readme-ja/TextWithShapeToAreaTypeSimple.md)
- [選択したテキストの文字ツメを 30% に設定します](readme-ja/Tsume30simple.md)
- [選択したテキストの基本的な文字組み設定（フォントサイズと行送り・自動カーニング・文字ツメ・](readme-ja/TypeBasicsPanel.md)
- [基準フォントサイズと倍率からタイプスケールを自動生成](readme-ja/TypeScaler.md)
- [基準フォントサイズと倍率からタイプスケールを自動生成](readme-ja/TypeScaler-v2.md)
- [選択したテキストの文字組み設定（フォント・フォントサイズ・自動カーニング・字間・文字揃え・行揃え・行送り・](readme-ja/UnifiedTypePanel-v2.md)
- [選択したテキストの文字組み設定（フォント・フォントサイズ・自動カーニング・字間・文字揃え・行揃え・行送り・](readme-ja/UnifiedTypePanel-v3.md)
- [Illustrator ドキュメント内の数字に桁区切りのカンマを自動で付与](readme-ja/formatNumberWithCommas.md)
- [選択中のテキストフレームを、指定した文字の直後で改行する](readme-ja/titlemaker.md)


## オブジェクトの配置や整列

- [配置したオブジェクトをグリッド状に配置する](readme-ja/SmartObjectDistributor.md)
- [選択したオブジェクトを幅や高さ、不透明度、カラーでソート](readme-ja/SmartObjectSorter.md)
- [オブジェクトを入れ替え](readme-ja/SwapNearestItemWithDialogbox.md)
- [2つのオブジェクトの位置を入れ替え](readme-ja/SwapObjects.md)
- [オブジェクトのリサイズ](readme-ja/SmartObjectResizer.md)
- [隣り合うオブジェクトのアキの中央に記号を配置](readme-ja/RightMarkPlacer.md)
- [選択した2つ以上のオブジェクトをいったんグループ解除してから、あらためて1つのグループにまとめます](readme-ja/AddToGroup.md)
- [選択した2つのオブジェクトの上下の間隔を、指定した値にそろえる常駐パレットです](readme-ja/AiAdjustVerticalGap.md)
- [選択内容に応じて、行送りと配置を次の順で判定して調整します](readme-ja/DistributeDownFromTop.md)
- [横並びに選択した複数オブジェクトのうち、最も左のものを固定し、以降を左方向へ「キー入力」の値ぶんずつ動かして…](readme-ja/DistributeLL.md)
- [横並びに選択した複数オブジェクトのうち、最も左のものを固定し、それ以降を右方向へ等間隔に再配置します](readme-ja/DistributeLR.md)
- [横並びに選択した複数オブジェクトのうち、最も右のものを固定し、以降を左方向へ「キー入力」の値ぶんずつ等間隔に…](readme-ja/DistributeRL.md)
- [選択内容に応じて、行送りと配置を次の順で判定して調整します](readme-ja/DistributeUpFromTop.md)
- [選択したオブジェクトのグループを入れ子ごとすべて解除し、あらためて1つのグループにまとめます](readme-ja/FlattenGroup.md)
- [複数選択したテキストフレームの背面に K15% の長方形を作成し、「_bg-rectangle」レイヤーに配…](readme-ja/GridTextLayout.md)
- [選択したオブジェクトを、1 つずつ個別のグループにまとめます](readme-ja/GroupEachSelection.md)
- [選択している配置画像（PlacedItem / RasterItem）の拡大・縮小率（%）を表示し、入力値で…](readme-ja/ImageScaler.md)
- [Illustrator の各種環境設定をダイアログボックスから変更可能にします](readme-ja/PreferenceManager.md)
- [選択したオブジェクトの移動・複製と反転・回転を、アイコンのクリックで即時実行する常駐パレットです（プレビュー…](readme-ja/QuickTransformPalette.md)
- [選択したオブジェクトをランダムに移動・変形・回転・不透明度を変更するスクリプト](readme-ja/RandomizeObjects.md)
- [選択中のオブジェクトが「だいたいグリッド状」に並んでいることを前提に、左右・上下の間隔で再配置します](readme-ja/RegridObjects.md)
- [選択中のオブジェクトがグループ内にある場合、親グループをたどって所属レイヤーの直下へ移動します](readme-ja/ReleaseFromGroup.md)
- [テキストに対して、回転／シアー（せん断）／比率を安全にリセット](readme-ja/ResetText.md)
- [配置画像・テキスト・長方形（パス）・クリップグループ・直線パスに対して、回転／シアー（せん断）／スケール／縦…](readme-ja/ResetTransform.md)
- [更新日：2026-02-26](readme-ja/SmartAlignAndTile-tate.md)
- [更新日：2026-01-19](readme-ja/SmartAlignAndTile-yoko.md)
- [選択オブジェクトを「縦/横」に並べて、指定した間隔で分布します。方向は自動判定も可能で、揃え（左右/上下）、…](readme-ja/SmartAlignDistribute.md)
- [選択オブジェクトを自動的に「重なり」「垂直方向」「水平方向」「近接度」などの条件に応じてグループ化するIll…](readme-ja/SmartAutoGroup.md)
- [ポイント文字およびエリア内文字に対して［字形の境界に整列］を制御するスクリプト](readme-ja/SmartVerticalAlign.md)
- [グループ内のテキストから数値を抽出し、数値に基づいてグループを並び替え縦方向に整列するIllustrator…](readme-ja/SortByNumbers.md)
- [テキストフレーム内のタブ区切りテキストを指定列の値で並び替えるIllustrator用スクリプトです](readme-ja/SortTextByColumn.md)
- [選択したオブジェクトを基準に、指定方向（右／左／上／下）にある最も近いオブジェクトと自然な見た目で位置を入れ…](readme-ja/SwapNearestItem.md)
- [Illustratorでテキストフレームを行・列単位で整列またはグループ化するスクリプト](readme-ja/TextGridAligner.md)
- [選択した複数オブジェクトを行・列として自動判定し、歯抜け（欠け）を許容しつつ、](readme-ja/TransposeGrid.md)
- [オブジェクトの重ね順を位置（X/Y）やZインデックスで並べ替える](readme-ja/ZIndexSorter.md)
- [選択した 2 つのオブジェクトの中心位置を入れ替えます](readme-ja/swap-2-objects.md)


## 基本図形と変形

- [正方形や正円、正三角形を作成するスクリプト](readme-ja/SmartShapeMaker.md)
- [アスペクト比で変形](readme-ja/AspectRatioScaler.md)
- [自由変形（フリーディストート）](readme-ja/SmartFreeDistort.md)
- [パスファインダー](readme-ja/AiSmartPathfinder.md)
- [パスの最適化](readme-ja/PathCleanupTool.md)
- [重なりをならす（オフセットパス＋合流）](readme-ja/OverlapRemover.md)
- [チケット・クーポン形状の作成（ミシン目、ギザギザ、スリット／ホール）](readme-ja/CouponTicketMaker.md)
- [選択オブジェクトの回転を水平（0°）に補正](readme-ja/ResetRotation.md)
- [更新日：2025-08-13](readme-ja/AddOutlineOffsetPath.md)
- [選択オブジェクトの全アンカーポイントに、マーカー（正方形／最前面オブジェクト／シンボル）を配置するユーティリ…](readme-ja/AiAnchorPointMarker.md)
- [更新日：2026-04-27](readme-ja/CenterLineConnectorFromRect.md)
- [更新日：2025-06-15](readme-ja/ColorToK100Converter.md)
- [更新日：2026-05-24](readme-ja/ConvertToRectangle.md)
- [選択したオブジェクトにランダムな変形を加え、手書き・スケッチ風の見た目に整えます](readme-ja/KPTSketchy.md)
- [更新日：2026-03-20](readme-ja/PathCleanupTool-v2.md)
- [更新日：2026-05-10](readme-ja/PathUnite.md)
- [更新日：2026-05-10](readme-ja/PathUniteOffsetTool.md)
- [更新日：2026-05-19](readme-ja/RectangleToArc.md)
- [SmartShapeMaker-v2](readme-ja/SmartShapeMaker-v2.md)
- [選択オブジェクトに［形状に変換］のライブエフェクトを適用する常駐パレット。パレットで「長方形／楕円」と「値を…](readme-ja/fx-all.md)
- [更新日：2026-05-01](readme-ja/test-20260501-040758.md)
- [更新日：2026-05-20](readme-ja/長方形に変換.md)


## アートワークやファイル

- [選択したオブジェクトを書き出す](readme-ja/SmartObjectExporter.md)
- [開いているファイルを1つに整列統合](readme-ja/SmartBatchImporter.md)
- [Figmaの⌘ + shift + C（ビットマップとしてコピー）](readme-ja/CopyAsPngLikeFigma.md)
- [リンク画像の管理](readme-ja/LinkedImageManager.md)
- [埋め込み画像をリンク画像に変換](readme-ja/UnembedRasterItems.md)
- [埋め込み画像をリンク画像に変換（アクション方式）](readme-ja/UnembedToLinks.md)
- [登録済みAIファイルからスタイル・ブラシ・フォント見本を読み込み](readme-ja/ImportStyles.md)
- [フォルダー内のAI／SVGをバージョン指定で一括保存](readme-ja/AIBatchVersionSave.md)
- [CSV／タブ区切りデータをテンプレートに流し込むデータ結合](readme-ja/VariableDataImport.md)
- [選択オブジェクトを高解像度でラスタライズし、PNG相当のビットマップとしてクリップボードにコピーするIllu…](readme-ja/CopyAsPngLikeFigmaWithDialog.md)
- [登録されているアクションセットを、デスクトップの `Illustrator_Actions` フォルダーへ…](readme-ja/ExportActions.md)
- [PDF/AI ファイルを指定したページ範囲で読み込み、現在のドキュメント上にページを配置します](readme-ja/PDFAIImporter.md)
- [選択した配置画像を基準に、同じリンクファイルを参照している配置画像をドキュメント内から検索し、指定ファイルへ…](readme-ja/RelinkSameImages.md)
- [現在選択中のリンク画像（PlacedItem）と同じリンクファイルを参照する](readme-ja/SelectSameLinks.md)
- [Illustrator ドキュメントに登録されたシンボルを一覧表示する専用アートボード「シンボル一覧」を自動…](readme-ja/SymbolListBuilder.md)
- [SymbolizeAndReplace](readme-ja/SymbolizeAndReplace.md)
- [アクティブドキュメントの全アートボードを、名前ごとのルールで PNG 書き出しします](readme-ja/export-Event.md)
- [アクティブなアートボードを PNG24 形式で書き出します](readme-ja/export200.md)


## カラー

- [画像やオブジェクトからカラーパレットを作成](readme-ja/ColorPaletteFromImage.md)
- [スウォッチの連続適用](readme-ja/ApplySwatchesToSelection.md)
- [カラーをランダム適用](readme-ja/ShuffleObjectColors.md)
- [カラーを配色](readme-ja/AiApplySwatchesToSelection.md)
- [選択オブジェクトの色からスウォッチとグラデーションを自動生成](readme-ja/CreateGradientFromSelection.md)
- [塗りと線の入れ替え・変換・消去](readme-ja/FillStrokeSwitcher.md)
- [選択したオブジェクトやテキストに、スウォッチや定義済みカラーを適用するモーダルダイアログです](readme-ja/AiApplySwatchesToSelection-dialog.md)
- [一時的なアクションを読み込んで実行し、選択オブジェクトの「透明部分を分割・統合」を適用します](readme-ja/FlattenTransparency.md)
- [選択オブジェクト（閉じたパス、テキスト）の塗りおよび線のカラーをスウォッチ（スポットカラー）として登録し、そ…](readme-ja/RegisterAndApplySwatches.md)
- [convert2separategradient](readme-ja/convert2separategradient.md)


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
- [すべてのアートボードに同じサイズの矩形を描画し、アートボード内のオブジェクトをマスクします](readme-ja/ArtboardMaskAndRelease.md)
- [全アートボード上にあるテキストフレームを収集し、最後のアートボードの右側に縦に並べて配置する](readme-ja/CollectArtboardTexts.md)
- [現在のアートボードとまったく同じ大きさの長方形を作成します](readme-ja/DrawArtboardRectangle.md)
- [選択したオブジェクトのバウンディングボックスに **上下マージン** を加味して、作業中のアートボードの *…](readme-ja/FitArtboardHeight.md)
- [選択したグループオブジェクトの境界に指定したマージンを加え、その範囲をアートボードとして自動追加するIllu…](readme-ja/Group2Artboards.md)
- [アクティブなアートボード以外を削除](readme-ja/RemoveOtherArtboards-v2.md)
- [ダイアログで指定した「幅」「高さ」に、アートボードをライブプレビューしながら変形します](readme-ja/ResizeArtboardsAll.md)
- [アクティブまたは全アートボードと同サイズの長方形を、オフセットを考慮して描画します](readme-ja/SmartDrawArtboardRectangle.md)


## レイヤー、重ね順

- [オブジェクトを指定レイヤーへ移動](readme-ja/SmartLayerManage.md)
- [選択しているオブジェクトを新規レイヤーに移動し、そのレイヤーを最背面に移動してロック](readme-ja/SendToBgLayer.md)
- [座標を基準に重ね順を変更](readme-ja/SortItemsByPosition.md)
- [選択したグループ内のサブグループを解除して、グループ構造を簡素化](readme-ja/SimplifyGroups.md)
- [水平方向にグループ化](readme-ja/SmartAutoGroup-yoko.md)
- [アートボード単位でオブジェクトをレイヤーに整理](readme-ja/ArtboardLayerOrganizer.md)
- [実行前にダイアログを表示し、処理条件を選択可能](readme-ja/FlattenLayers.md)
- [アクティブレイヤーを「テンプレート」属性（ロック・印刷不可・画像を薄く表示）にする（ON 専用）](readme-ja/MakeTemplateLayer.md)
- [アクティブレイヤーの「テンプレート」属性（ロック・印刷不可・画像を薄く表示）を ON / OFF する](readme-ja/ToggleTemplateLayer.md)
- [一時的なアクションを読み込んで実行し、対象レイヤーに「テンプレート」属性（ロック・印刷不可・画像を薄く表示）…](readme-ja/bg-template-only.md)


## マスク

- [オブジェクトのまとまりごとにグループ化したり、マスクする](readme-ja/SmartClipAndGroup.md)
- [〈クリッピングマスクを解除〉を拡張](readme-ja/ReleaseClipMask.md)
- [パズル](readme-ja/SmartSliceWithPuzzlify.md)
- [マスクパスのサイズ変更](readme-ja/ResizeClipMask.md)
- [クリップグループのマスクと内容を調整](readme-ja/ClipMaskAdjust.md)
- [選択された画像（配置画像/埋め込み画像）やクリッピングマスクグループ内の画像に対して、中心を基準とした最小の…](readme-ja/ClipWithSquare.md)


## ガイド

- [グリッド状にガイドを生成](readme-ja/GenerateGuidesGrid.md)
- [囲み罫とグリッド](readme-ja/SmartGridMaker.md)
- [選択したオブジェクトに対してガイドを自動作成](readme-ja/CreateGuidesFromSelection.md)
- [すべてのガイドを削除](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/guide/DeleteAllGuides.jsx)
- [Photoshopの「新規ガイド」をIllustratorでも可能にする](readme-ja/NewGuideMaker.md)
- [アートボード基準のガイド作成とルーラーガイドの変換](readme-ja/AiCreateArtboardGuides.md)
- [アートボードを基準にガイドを整理・作成するツール。次の3系統をダイアログでまとめて設定できる](readme-ja/AiCreateArtboardGuides-v2.md)
- [複数のレイヤー／サブレイヤーに散在するガイドを、1 つのレイヤー（既定は `// guide`）へ集約します](readme-ja/CollectGuides.md)
- [ドキュメント内のすべてのガイドを削除します](readme-ja/DeleteAllGuides.md)
- [ドキュメント内のガイド（ルーラーガイド含む）の交点を基準に長方形を自動生成します](readme-ja/MakeRectangleFromGuides.md)
- [unlockGuideLayerAndClearGuides](readme-ja/unlockGuideLayerAndClearGuides.md)


## ドキュメント

- [ドキュメントの切替](readme-ja/SmartSwitchDocs.md)
- [ファイル名を変更して保存](readme-ja/Ai-FileNameManager.md)
- [選択オブジェクトのみを残した複製ドキュメントを作成するInDesign用スクリプトです](readme-ja/CloneDocSelectedOnly.md)


## 環境設定

- [クイック環境設定](readme-ja/AiQuickPrefsPalette.md)
- [クイック環境設定（SuperSimple）](readme-ja/AiQuickPrefsPalette-SuperSimple.md)
- [アートボード関連の環境設定](readme-ja/ArtboardDisplayPresetManager.md)
- [環境設定：変形と整列](readme-ja/PreferenceManagerForTransformAndAlign.md)
- [環境設定をまとめて変更](readme-ja/PresetManager.md)
- [Illustrator の各種環境設定の切り替えと、選択オブジェクトの反転・回転を、常駐パレットでまとめて操…](readme-ja/AiQuickPrefsPalette-simple.md)
- [Illustrator の各種環境設定をダイアログボックスから変更可能にします](readme-ja/PreferenceManager-unit.md)
- [PresetManager の［プリセット1］と同じ設定一式を、Illustrator の環境設定へまとめて…](readme-ja/PresetManagerPreset1.md)


## その他

- [ドキュメントの不要な要素をまとめて削除](readme-ja/AiDocumentCleaner.md)
- [トンボ作成](readme-ja/AddTrimMark.md)
- [テキストの文字列でグラフィックスタイルを登録](readme-ja/RegisterGraphicStyleWithText.md)
- [選択オブジェクトを条件で絞り込み再選択](readme-ja/SmartSelectionFilter.md)
- [選択中とドキュメント全体のオブジェクト数を集計](readme-ja/SelectionInspector.md)
- [選択オブジェクトから新規レイヤー・アートボード・ドキュメントを作成](readme-ja/SelectionToNew.md)
- [シンボルへのリンクを解除して整理](readme-ja/AiBreakLink.md)
- [常に現在のアートボードに対して、日本式トンボを作成するIllustrator用スクリプトです](readme-ja/AddTrimMarkToCurrentArtboard.md)
- [開いているドキュメントのアクティブビューの回転角度（表示の回転）を取得](readme-ja/AiSmartRotateView.md)
- [常駐エンジンで動いている各種フローティングパレットをまとめて閉じるユーティリティ](readme-ja/CloseAllPalettes.md)
- [環境設定の「角度の制限」（Shiftキーを押したときの角度）と「キー増加」（矢印キーの移動量）の変更、](readme-ja/DirectPrefs.md)
- [選択したオブジェクトを 1 つずつ選び直しながら、それぞれにアピアランスの分割を適用します](readme-ja/ExpandAppearanceEachObject.md)
- [ダイアログの「スタイルを読み込み」ボタンで AI ファイルを指定し、そのファイル内のグラフィックスタイルを取…](readme-ja/ImportGraphicStyles.md)
- [ダイアログのラジオボタン（文字白抜き／枠のみ）でグラフィックスタイルを選択](readme-ja/ImportGraphicStyles-v2.md)
- [選択中／全体のパス統計をカウントし、常駐パレットで表示](readme-ja/PathInspector.md)
- [グラフィックスタイル／ブラシ／スウォッチ／シンボルの「名前」を、ダイアログで指定した「検索→置換」で一括変更…](readme-ja/RenameAssets.md)
- [選択したオブジェクトの線幅と矢印（始点／終点の形状・倍率・先端位置）をまとめて設定します](readme-ja/SetStrokeAndArrowheads.md)
- [アクティブなドキュメント上で、指定した .ai / .pdf（PDFはページ指定）をグリッド配置し、ポートフ…](readme-ja/SlideCollage.md)
- [基準日をもとに、指定した月数ぶんのカレンダー（月曜はじまり）をアートボード中心に作成します](readme-ja/SmartCalendarMaker.md)
- [Illustrator のアートボード／シンボル／レイヤー名を、接頭辞・接尾辞・名前ソース・検索置換を組み合…](readme-ja/SmartRenamer.md)
- [2つのオブジェクト（テキスト、パス、グループなど）を選択して実行すると、各オブジェクトの背面に左右2分割の背…](readme-ja/SmartTableMaker.md)
- [2つのオブジェクト（テキスト、パス、グループなど）を選択して実行すると、各オブジェクトの背面に2分割の背景を…](readme-ja/SplitBackgroundForTwo.md)
- [1つのオブジェクト（テキスト、パス、グループなど）を選択して実行すると、そのオブジェクトの外接矩形を左右また…](readme-ja/SplitForTwo.md)
- [選択した外枠の長方形を基準に、内部の縦罫／横罫を等間隔に再配置します](readme-ja/TableRuleAverager.md)
- [Illustrator の選択テキストや全体の文字情報を統計的に可視化](readme-ja/TextCountStats.md)
- [選択オブジェクトに合わせて、アクティブビューをズーム＆センタリングします](readme-ja/ZoomToSelection.md)
- [選択したオブジェクトから、新規レイヤー・新規アートボード・新規ドキュメントを作成するスクリプトです](readme-ja/new.md)
- [選択中のオブジェクトの見た目を「名称未設定」で登録する Illustrator 標準の挙動を回避し、固定名の…](readme-ja/register-temp-style.md)
- [アートボード／レイヤー／シンボル／グラフィックスタイルの名前を、検索置換とナンバリングでまとめて変更します](readme-ja/renamer.md)
- [選択したパスアイテムの線端を「丸型先端」に設定します](readme-ja/丸型先端にする.md)

