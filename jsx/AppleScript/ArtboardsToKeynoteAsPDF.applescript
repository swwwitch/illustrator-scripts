-- アートボードごとにPDFを書き出し、Keynoteのスライドへ画像として挿入する
-- クリップボードとキー操作を使わないので速い（貼られるのはPDF画像）

use AppleScript version "2.4"
use framework "Foundation"
use framework "AppKit"
use scripting additions

-- 設定ダイアログボックスとの受け渡し（メインスレッドで開くため、引数ではなくプロパティで渡す）
property dialogIsNewDocument : false
property dialogSlideCount : 0
property dialogSlideText : ""
property dialogScaleText : ""
property dialogButton : ""
property dialogArtboardCount : 0
property dialogArtboardMode : "current" -- "current"（現在のアートボードのみ）/ "all"（すべて）/ "range"（番号で範囲を指定）
property dialogArtboardFromText : ""
property dialogArtboardToText : ""
property dialogCurrentArtboard : 1
property dialogPlacementMode : "overlay" -- "append"（末尾にスライドを追加）/ "overlay"（既存スライドにオーバーレイ）
property dialogUseArtboardBounds : true -- true：アートボード全体を書き出す／false：アートワークの範囲だけを書き出す
-- ラジオボタンの切り替えで入力欄をディムするため、表示中の部品を控える
property dialogOverlayRadio : missing value
property dialogSlideRow : missing value
property dialogRangeRadio : missing value
property dialogRangeFields : missing value

tell application "Adobe Illustrator"
	if (count of documents) = 0 then
		display alert "Illustratorでドキュメントを開いてください。"
		return
	end if

	set artboardCount to count of artboards of current document
	-- 終了後に戻すため、作業中のアートボードを控える
	set originalArtboardIndex to do javascript "app.activeDocument.artboards.getActiveArtboardIndex();"
	-- スケールの初期値を決めるため、全アートボードの大きさを「幅,高さ;幅,高さ;…」（pt）で受け取る
	set artboardSizesText to do javascript "
		var sizes = [];
		var artboards = app.activeDocument.artboards;
		for (var i = 0; i < artboards.length; i++) {
			var rect = artboards[i].artboardRect;
			sizes.push((rect[2] - rect[0]) + ',' + (rect[1] - rect[3]));
		}
		sizes.join(';');
	"
end tell

tell application "Keynote Creator Studio"
	activate

	if (count of documents) = 0 then
		set keynoteDoc to make new document
		set isNewDocument to true
	else
		set keynoteDoc to front document
		set isNewDocument to false
	end if
	-- slide number は非表示スライドを数えない番号（ナビゲータの番号）、slide N は非表示スライドも含めた並び順
	-- 開始スライドは slide number で受け取るため、表示するスライドの並び順を控えて読み替える
	set skippedFlags to skipped of slides of keynoteDoc
	set visibleSlideIndexes to {}
	repeat with i from 1 to count of skippedFlags
		if not (item i of skippedFlags) then set end of visibleSlideIndexes to i
	end repeat
	set slideCount to count of visibleSlideIndexes
	-- 指定スライドの入力欄の初期値にするため、表示中のスライド番号を控える
	try
		set currentSlideNumber to slide number of current slide of keynoteDoc
	on error
		set currentSlideNumber to 1
	end try
	-- 画像をスライドの中央に置くため、スライドの大きさを控える
	set slideWidth to width of keynoteDoc
	set slideHeight to height of keynoteDoc
end tell

-- 貼り付け先：新規ドキュメントはスライド1から、既存ドキュメントは末尾に追加か指定スライドからを選ぶ
-- startSlideIndex 以降は既存スライド（非表示スライドは飛ばす）に貼り、足りなくなったら末尾に追加する
-- スケール（%）は100でアートボードと同じ大きさ（1pt＝1ポイント）。初期値は全アートボードがスライドに収まる最大の値
-- アートボードの範囲・スケール・書き出し範囲・配置方法・開始スライドは1つのダイアログボックスで受け取り、入力が不正なら値を残して開き直す
set dialogIsNewDocument to isNewDocument
set dialogArtboardCount to artboardCount
set dialogCurrentArtboard to (originalArtboardIndex as integer) + 1
set dialogArtboardFromText to "1"
set dialogArtboardToText to artboardCount as text
set dialogSlideCount to slideCount
set dialogSlideText to currentSlideNumber as text
set dialogScaleText to (fitScalePercent(artboardSizesText, slideWidth, slideHeight)) as text
repeat
	if (current application's NSThread's isMainThread()) as boolean then
		my showSettingsDialog:(missing value)
	else
		my performSelectorOnMainThread:"showSettingsDialog:" withObject:(missing value) waitUntilDone:true
	end if
	if dialogButton is "キャンセル" then return

	set errorText to ""
	if dialogArtboardMode is "range" then
		try
			set firstArtboard to dialogArtboardFromText as integer
			set lastArtboard to dialogArtboardToText as integer
			if firstArtboard < 1 or lastArtboard > artboardCount or firstArtboard > lastArtboard then error
		on error
			set errorText to "アートボードには1〜" & artboardCount & "の範囲を、小さい番号から入力してください。"
		end try
	else if dialogArtboardMode is "current" then
		set firstArtboard to dialogCurrentArtboard
		set lastArtboard to dialogCurrentArtboard
	else
		set firstArtboard to 1
		set lastArtboard to artboardCount
	end if
	try
		set scalePercent to dialogScaleText as real
		if scalePercent ≤ 0 then error
	on error
		set errorText to "スケールには0より大きい数字を入力してください。"
	end try
	if isNewDocument then
		set startSlideIndex to 1
	else if dialogPlacementMode is "overlay" then
		try
			set startSlideIndex to dialogSlideText as integer
			if startSlideIndex < 1 or startSlideIndex > slideCount then error
		on error
			set errorText to "開始スライドには1〜" & slideCount & "の数字を入力してください。"
		end try
	else
		set startSlideIndex to slideCount + 1
	end if
	if errorText is "" then exit repeat
	display alert errorText
end repeat
set scaleFactor to scalePercent / 100

-- 書き出し用の一時フォルダー
set exportRoot to do shell script "mktemp -d -t ArtboardsToKeynote"

set pasteCount to 0
set skippedArtboards to {}

repeat with artboardIndex from firstArtboard to lastArtboard

	-- アートボードごとに別のフォルダーへ書き出す（書き出し名はアートボード名で決まるため）
	set exportFolder to exportRoot & "/" & artboardIndex
	do shell script "mkdir -p " & quoted form of exportFolder

	tell application "Adobe Illustrator"
		set jsCode to "
			var doc = app.activeDocument;
			doc.artboards.setActiveArtboardIndex(" & (artboardIndex - 1) & ");
			app.executeMenuCommand('deselectall');
			doc.selectObjectsOnActiveArtboard();
			var count = doc.selection.length;
			app.executeMenuCommand('deselectall');
			var rect = doc.artboards[" & (artboardIndex - 1) & "].artboardRect;
			if (count > 0) {
				var useArtboardBounds = " & dialogUseArtboardBounds & ";
				var originalLayer = doc.activeLayer;
				var tempLayer = null;
				try {
					// PDFの範囲をアートボード全体にするときは、一時レイヤーにアートボードサイズの透明な長方形を描いてから書き出す
					if (useArtboardBounds) {
						tempLayer = doc.layers.add();
						var boundsRect = tempLayer.pathItems.rectangle(rect[1], rect[0], rect[2] - rect[0], rect[1] - rect[3]);
						boundsRect.filled = false;
						boundsRect.stroked = false;
					}
					var exportItem = new ExportForScreensItemToExport();
					exportItem.document = false;
					exportItem.artboards = '" & artboardIndex & "';
					doc.exportForScreens(new Folder('" & exportFolder & "'), ExportForScreensType.SE_PDF, new ExportForScreensPDFOptions(), exportItem, '');
				} finally {
					if (tempLayer) tempLayer.remove();
					doc.activeLayer = originalLayer;
				}
			}
			[count, rect[2] - rect[0], rect[1] - rect[3]].join(',');
		"

		set jsResult to do javascript jsCode
	end tell

	-- 結果は「オブジェクト数,アートボードの幅,高さ」（幅・高さはpt）
	set AppleScript's text item delimiters to ","
	set selectedCount to (text item 1 of jsResult) as integer
	set artboardWidth to (text item 2 of jsResult) as real
	set artboardHeight to (text item 3 of jsResult) as real
	set AppleScript's text item delimiters to ""

	set pdfPath to ""
	if selectedCount > 0 then
		set pdfPath to do shell script "find " & quoted form of exportFolder & " -name '*.pdf' | head -n 1"
	end if

	if pdfPath is "" then
		-- 空のアートボード、または書き出せなかったアートボードは飛ばす
		set end of skippedArtboards to artboardIndex
	else
		set targetSlideIndex to startSlideIndex + pasteCount
		set pasteCount to pasteCount + 1
		set pdfFile to POSIX file pdfPath as alias
		-- アートワークの範囲だけを書き出したときは、PDFのページの大きさ（pt）を画像の大きさにする
		if not dialogUseArtboardBounds then
			set pdfSize to (current application's NSImage's alloc()'s initWithContentsOfFile:pdfPath)'s |size|()
			set artboardWidth to (width of pdfSize) as real
			set artboardHeight to (height of pdfSize) as real
		end if

		tell application "Keynote Creator Studio"
			tell keynoteDoc
				if targetSlideIndex ≤ (count of visibleSlideIndexes) then
					set targetSlide to slide (item targetSlideIndex of visibleSlideIndexes)
				else
					set targetSlide to make new slide at end of slides
				end if
				if pasteCount = 1 then set firstTargetSlide to targetSlide

				-- PDFを画像として挿入し、書き出した範囲の大きさ（1pt＝1ポイント）にスケールを掛けてスライドの中央に置く
				set imageWidth to artboardWidth * scaleFactor
				set imageHeight to artboardHeight * scaleFactor
				tell targetSlide
					set pdfImage to make new image with properties {file:pdfFile}
					set width of pdfImage to imageWidth
					set height of pdfImage to imageHeight
					set position of pdfImage to {(slideWidth - imageWidth) / 2, (slideHeight - imageHeight) / 2}
				end tell
			end tell
		end tell
	end if
end repeat

do shell script "rm -rf " & quoted form of exportRoot

-- Illustratorの状態を戻す
tell application "Adobe Illustrator"
	do javascript "
		var doc = app.activeDocument;
		app.executeMenuCommand('deselectall');
		doc.artboards.setActiveArtboardIndex(" & originalArtboardIndex & ");
	"
end tell

tell application "Keynote Creator Studio"
	activate
	if pasteCount > 0 then set current slide of keynoteDoc to firstTargetSlide
end tell

if (count of skippedArtboards) = 0 then
	if dialogArtboardMode is "current" then
		display dialog "アートボード " & firstArtboard & " をKeynoteへ挿入しました。" buttons {"OK"} default button "OK"
	else if dialogArtboardMode is "range" then
		display dialog "アートボード " & firstArtboard & "〜" & lastArtboard & " をKeynoteへ挿入しました。" buttons {"OK"} default button "OK"
	else
		display dialog "全アートボードをKeynoteへ挿入しました。" buttons {"OK"} default button "OK"
	end if
else
	set AppleScript's text item delimiters to ", "
	set skippedText to skippedArtboards as text
	set AppleScript's text item delimiters to ""
	display dialog "Keynoteへ挿入しました。" & return & "挿入しなかったアートボード（空または書き出しに失敗） : " & skippedText buttons {"OK"} default button "OK"
end if

-- アートボードの範囲・スケール・書き出し範囲・配置方法（ラジオボタン）・開始スライドを並べたダイアログボックスを開く
-- 結果は dialogButton・dialogArtboardMode・dialogArtboardFromText・dialogArtboardToText・
-- dialogScaleText・dialogUseArtboardBounds・dialogPlacementMode・dialogSlideText に入る（NSAlertはメインスレッドで開く）
on showSettingsDialog:unusedArgument
	set settingsAlert to current application's NSAlert's alloc()'s init()
	settingsAlert's setMessageText:"IllustratorのアートボードをKeynoteへ挿入"
	set buttonTitles to {"挿入", "キャンセル"}
	repeat with buttonTitle in buttonTitles
		set lastButton to (settingsAlert's addButtonWithTitle:(buttonTitle as text))
	end repeat
	-- Escで［キャンセル］
	lastButton's setKeyEquivalent:(character id 27)

	-- 上から アートボード（現在のみ／すべて／範囲）／スケール／書き出し範囲／配置方法／開始スライド（新規ドキュメントは配置方法と開始スライドなし）
	if dialogIsNewDocument then
		set boundsY to 0
	else
		set boundsY to 86
	end if
	set scaleY to boundsY + 26
	set accessoryView to current application's NSView's alloc()'s initWithFrame:{{0, 0}, {300, scaleY + 104}}

	-- アートボードの範囲（同じ親・同じアクションのラジオボタンは自動で排他になる）
	set artboardLabel to current application's NSTextField's labelWithString:"アートボード :"
	artboardLabel's setAlignment:(current application's NSTextAlignmentRight)
	artboardLabel's setFrame:{{0, scaleY + 87}, {90, 17}}
	accessoryView's addSubview:artboardLabel
	set currentRadio to current application's NSButton's radioButtonWithTitle:("現在のアートボードのみ（" & dialogCurrentArtboard & "）") target:me action:"artboardRangeChanged:"
	currentRadio's setFrame:{{96, scaleY + 86}, {204, 18}}
	accessoryView's addSubview:currentRadio
	set allRadio to current application's NSButton's radioButtonWithTitle:"すべて" target:me action:"artboardRangeChanged:"
	allRadio's setFrame:{{96, scaleY + 60}, {204, 18}}
	accessoryView's addSubview:allRadio
	set rangeRadio to current application's NSButton's radioButtonWithTitle:"" target:me action:"artboardRangeChanged:"
	rangeRadio's setFrame:{{96, scaleY + 34}, {20, 18}}
	accessoryView's addSubview:rangeRadio
	set fromField to current application's NSTextField's alloc()'s initWithFrame:{{118, scaleY + 32}, {44, 22}}
	fromField's setStringValue:dialogArtboardFromText
	accessoryView's addSubview:fromField
	set rangeSeparator to current application's NSTextField's labelWithString:"〜"
	rangeSeparator's setFrame:{{166, scaleY + 35}, {16, 17}}
	accessoryView's addSubview:rangeSeparator
	set toField to current application's NSTextField's alloc()'s initWithFrame:{{184, scaleY + 32}, {44, 22}}
	toField's setStringValue:dialogArtboardToText
	accessoryView's addSubview:toField
	if dialogArtboardMode is "range" then
		rangeRadio's setState:(current application's NSControlStateValueOn)
	else if dialogArtboardMode is "current" then
		currentRadio's setState:(current application's NSControlStateValueOn)
	else
		allRadio's setState:(current application's NSControlStateValueOn)
	end if
	set dialogRangeRadio to rangeRadio
	set dialogRangeFields to {fromField, toField, rangeSeparator}
	my updateArtboardRangeEnabled()

	set scaleRow to my addFieldRow(accessoryView, scaleY, "スケール :", dialogScaleText, "%")
	fromField's setNextKeyView:toField
	toField's setNextKeyView:(inputField of scaleRow)
	(inputField of scaleRow)'s setNextKeyView:fromField

	-- オフにするとアートワークの範囲だけを書き出す（アートボード内の位置は保たれない）
	set boundsCheckbox to current application's NSButton's checkboxWithTitle:"アートボード全体を書き出す" target:(missing value) action:(missing value)
	boundsCheckbox's setFrame:{{96, boundsY + 2}, {204, 18}}
	if dialogUseArtboardBounds then boundsCheckbox's setState:(current application's NSControlStateValueOn)
	accessoryView's addSubview:boundsCheckbox

	if not dialogIsNewDocument then
		set appendRadio to current application's NSButton's radioButtonWithTitle:"末尾にスライドを追加して配置" target:me action:"placementModeChanged:"
		appendRadio's setFrame:{{96, 58}, {204, 18}}
		accessoryView's addSubview:appendRadio
		set overlayRadio to current application's NSButton's radioButtonWithTitle:"既存スライドにオーバーレイ" target:me action:"placementModeChanged:"
		overlayRadio's setFrame:{{96, 34}, {204, 18}}
		accessoryView's addSubview:overlayRadio
		set slideRow to my addFieldRow(accessoryView, 0, "開始スライド :", dialogSlideText, "（1〜" & dialogSlideCount & "）")
		if dialogPlacementMode is "overlay" then
			overlayRadio's setState:(current application's NSControlStateValueOn)
		else
			appendRadio's setState:(current application's NSControlStateValueOn)
		end if
		set dialogOverlayRadio to overlayRadio
		set dialogSlideRow to slideRow
		my updateSlideRowEnabled()
		(inputField of scaleRow)'s setNextKeyView:(inputField of slideRow)
		(inputField of slideRow)'s setNextKeyView:fromField
	end if
	settingsAlert's setAccessoryView:accessoryView
	settingsAlert's layout()
	(settingsAlert's |window|())'s setInitialFirstResponder:(inputField of scaleRow)

	-- osascript などから開くと前面のアプリの後ろに隠れるので、アクセサリーとして前面に出し、モーダルパネルのレベルに上げる
	set hostApp to current application's NSApplication's sharedApplication()
	hostApp's setActivationPolicy:(current application's NSApplicationActivationPolicyAccessory)
	hostApp's activateIgnoringOtherApps:true
	(settingsAlert's |window|())'s setLevel:(current application's NSModalPanelWindowLevel)
	set response to (settingsAlert's runModal()) as integer
	-- 1つ目のボタンが1000、以降1ずつ増える
	set dialogButton to item (response - 1000 + 1) of buttonTitles
	if ((rangeRadio's state()) as integer) = 1 then
		set dialogArtboardMode to "range"
	else if ((currentRadio's state()) as integer) = 1 then
		set dialogArtboardMode to "current"
	else
		set dialogArtboardMode to "all"
	end if
	set dialogArtboardFromText to (fromField's stringValue()) as text
	set dialogArtboardToText to (toField's stringValue()) as text
	set dialogScaleText to ((inputField of scaleRow)'s stringValue()) as text
	set dialogUseArtboardBounds to ((boundsCheckbox's state()) as integer) = 1
	if not dialogIsNewDocument then
		if ((overlayRadio's state()) as integer) = 1 then
			set dialogPlacementMode to "overlay"
		else
			set dialogPlacementMode to "append"
		end if
		set dialogSlideText to ((inputField of slideRow)'s stringValue()) as text
	end if
	set dialogOverlayRadio to missing value
	set dialogSlideRow to missing value
	set dialogRangeRadio to missing value
	set dialogRangeFields to missing value
end showSettingsDialog:

-- アートボードのラジオボタンを切り替えたとき、範囲の入力欄のディムを合わせる
on artboardRangeChanged:sender
	my updateArtboardRangeEnabled()
end artboardRangeChanged:

-- 範囲を選んだときだけ、アートボード番号の入力欄を有効にする
on updateArtboardRangeEnabled()
	set isRange to ((dialogRangeRadio's state()) as integer) = 1
	(item 1 of dialogRangeFields)'s setEnabled:isRange
	(item 2 of dialogRangeFields)'s setEnabled:isRange
	if isRange then
		(item 3 of dialogRangeFields)'s setTextColor:(current application's NSColor's labelColor())
	else
		(item 3 of dialogRangeFields)'s setTextColor:(current application's NSColor's disabledControlTextColor())
	end if
end updateArtboardRangeEnabled

-- 配置方法のラジオボタンを切り替えたとき、開始スライドの行のディムを合わせる
on placementModeChanged:sender
	my updateSlideRowEnabled()
end placementModeChanged:

-- ［既存スライドにオーバーレイ］のときだけ開始スライドの行を有効にする
on updateSlideRowEnabled()
	set isOverlay to ((dialogOverlayRadio's state()) as integer) = 1
	(inputField of dialogSlideRow)'s setEnabled:isOverlay
	if isOverlay then
		set rowTextColor to current application's NSColor's labelColor()
	else
		set rowTextColor to current application's NSColor's disabledControlTextColor()
	end if
	(labelField of dialogSlideRow)'s setTextColor:rowTextColor
	(unitField of dialogSlideRow)'s setTextColor:rowTextColor
end updateSlideRowEnabled

-- 「ラベル・入力欄・単位」の1行を parentView の rowY の位置に置き、3つの部品を返す
on addFieldRow(parentView, rowY, labelText, valueText, unitText)
	set labelField to current application's NSTextField's labelWithString:labelText
	labelField's setAlignment:(current application's NSTextAlignmentRight)
	labelField's setFrame:{{0, rowY + 3}, {90, 17}}
	parentView's addSubview:labelField

	set inputField to current application's NSTextField's alloc()'s initWithFrame:{{96, rowY}, {60, 22}}
	inputField's setStringValue:valueText
	parentView's addSubview:inputField

	set unitField to current application's NSTextField's labelWithString:unitText
	unitField's setFrame:{{162, rowY + 3}, {138, 17}}
	parentView's addSubview:unitField
	return {labelField:labelField, inputField:inputField, unitField:unitField}
end addFieldRow

-- 全アートボードがスライドに収まる最大のスケール（%、整数に切り捨て）を返す
-- artboardSizesText は「幅,高さ;幅,高さ;…」（pt）
on fitScalePercent(artboardSizesText, slideWidth, slideHeight)
	set fitScale to missing value
	set AppleScript's text item delimiters to ";"
	set sizeItems to text items of artboardSizesText
	set AppleScript's text item delimiters to ","
	repeat with sizeItem in sizeItems
		set artboardWidth to (text item 1 of sizeItem) as real
		set artboardHeight to (text item 2 of sizeItem) as real
		set widthScale to slideWidth / artboardWidth
		set heightScale to slideHeight / artboardHeight
		if heightScale < widthScale then set widthScale to heightScale
		if fitScale is missing value or widthScale < fitScale then set fitScale to widthScale
	end repeat
	set AppleScript's text item delimiters to ""
	set fitPercent to (fitScale * 100) div 1
	if fitPercent < 1 then set fitPercent to 1
	return fitPercent
end fitScalePercent
