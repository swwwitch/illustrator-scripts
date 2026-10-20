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

tell application "Adobe Illustrator"
	if (count of documents) = 0 then
		display alert "Illustratorでドキュメントを開いてください。"
		return
	end if

	set artboardCount to count of artboards of current document
	-- 終了後に戻すため、作業中のアートボードを控える
	set originalArtboardIndex to do javascript "app.activeDocument.artboards.getActiveArtboardIndex();"
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
	set slideCount to count of slides of keynoteDoc
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
-- startSlideIndex 以降は既存スライドに貼り、足りなくなったら末尾に追加する
-- スケール（%）は100でアートボードと同じ大きさ（1pt＝1ポイント）
-- 貼り付け先・開始スライド・スケールは1つのダイアログボックスで受け取り、入力が不正なら値を残して開き直す
set dialogIsNewDocument to isNewDocument
set dialogSlideCount to slideCount
set dialogSlideText to currentSlideNumber as text
set dialogScaleText to "100"
repeat
	if (current application's NSThread's isMainThread()) as boolean then
		my showSettingsDialog:(missing value)
	else
		my performSelectorOnMainThread:"showSettingsDialog:" withObject:(missing value) waitUntilDone:true
	end if
	if dialogButton is "キャンセル" then return

	set errorText to ""
	try
		set scalePercent to dialogScaleText as real
		if scalePercent ≤ 0 then error
	on error
		set errorText to "スケールには0より大きい数字を入力してください。"
	end try
	if dialogButton is "既存スライドに貼る" then
		try
			set startSlideIndex to dialogSlideText as integer
			if startSlideIndex < 1 or startSlideIndex > slideCount then error
		on error
			set errorText to "開始スライドには1〜" & slideCount & " の数字を入力してください。"
		end try
	else if dialogButton is "末尾に追加" then
		set startSlideIndex to slideCount + 1
	else
		set startSlideIndex to 1
	end if
	if errorText is "" then exit repeat
	display alert errorText
end repeat
set scaleFactor to scalePercent / 100


-- 書き出し用の一時フォルダー
set exportRoot to do shell script "mktemp -d -t ArtboardsToKeynote"

set pasteCount to 0
set skippedArtboards to {}

repeat with artboardIndex from 1 to artboardCount

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
				// PDFの範囲がアートボード全体になるよう、一時レイヤーにアートボードサイズの長方形を描いてから書き出す
				var originalLayer = doc.activeLayer;
				var tempLayer = doc.layers.add();
				try {
					var boundsRect = tempLayer.pathItems.rectangle(rect[1], rect[0], rect[2] - rect[0], rect[1] - rect[3]);
					boundsRect.filled = false;
					boundsRect.stroked = false;
					var exportItem = new ExportForScreensItemToExport();
					exportItem.document = false;
					exportItem.artboards = '" & artboardIndex & "';
					doc.exportForScreens(new Folder('" & exportFolder & "'), ExportForScreensType.SE_PDF, new ExportForScreensPDFOptions(), exportItem, '');
				} finally {
					tempLayer.remove();
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

		tell application "Keynote Creator Studio"
			tell keynoteDoc
				if targetSlideIndex ≤ (count of slides) then
					set targetSlide to slide targetSlideIndex
				else
					set targetSlide to make new slide at end of slides
				end if

				-- PDFを画像として挿入し、アートボードの大きさ（1pt＝1ポイント）にスケールを掛けてスライドの中央に置く
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
	if pasteCount > 0 then set current slide of keynoteDoc to slide startSlideIndex of keynoteDoc
end tell

if (count of skippedArtboards) = 0 then
	display dialog "全アートボードをKeynoteへ挿入しました。" buttons {"OK"} default button "OK"
else
	set AppleScript's text item delimiters to ", "
	set skippedText to skippedArtboards as text
	set AppleScript's text item delimiters to ""
	display dialog "Keynoteへ挿入しました。" & return & "空または書き出せなかったアートボード : " & skippedText buttons {"OK"} default button "OK"
end if

-- 貼り付け先のボタンと、開始スライド・スケールの入力欄を並べたダイアログボックスを開く
-- 結果は dialogButton・dialogSlideText・dialogScaleText に入る（NSAlertはメインスレッドで開く）
on showSettingsDialog:unusedArgument
	set settingsAlert to current application's NSAlert's alloc()'s init()
	settingsAlert's setMessageText:"アートボードをKeynoteへ挿入"
	if dialogIsNewDocument then
		set buttonTitles to {"OK", "キャンセル"}
		settingsAlert's setInformativeText:"新規ドキュメントのスライド1から挿入します。"
	else
		set buttonTitles to {"末尾に追加", "既存スライドに貼る", "キャンセル"}
		settingsAlert's setInformativeText:"［既存スライドに貼る］は開始スライドから順に貼り、足りなくなったら末尾に追加します。"
	end if
	repeat with buttonTitle in buttonTitles
		set lastButton to (settingsAlert's addButtonWithTitle:(buttonTitle as text))
	end repeat
	-- Escで［キャンセル］
	lastButton's setKeyEquivalent:(character id 27)

	-- 入力欄：既存ドキュメントでは開始スライドを上の行に置く
	if dialogIsNewDocument then
		set accessoryView to current application's NSView's alloc()'s initWithFrame:{{0, 0}, {260, 22}}
		set scaleField to my addFieldRow(accessoryView, 0, "スケール :", dialogScaleText, "%")
		set firstField to scaleField
	else
		set accessoryView to current application's NSView's alloc()'s initWithFrame:{{0, 0}, {260, 54}}
		set slideField to my addFieldRow(accessoryView, 32, "開始スライド :", dialogSlideText, "（1〜" & dialogSlideCount & "）")
		set scaleField to my addFieldRow(accessoryView, 0, "スケール :", dialogScaleText, "%")
		slideField's setNextKeyView:scaleField
		scaleField's setNextKeyView:slideField
		set firstField to slideField
	end if
	settingsAlert's setAccessoryView:accessoryView
	settingsAlert's layout()
	(settingsAlert's |window|())'s setInitialFirstResponder:firstField

	-- osascript などから開くと前面のアプリの後ろに隠れるので、アクセサリーとして前面に出し、モーダルパネルのレベルに上げる
	set hostApp to current application's NSApplication's sharedApplication()
	hostApp's setActivationPolicy:(current application's NSApplicationActivationPolicyAccessory)
	hostApp's activateIgnoringOtherApps:true
	(settingsAlert's |window|())'s setLevel:(current application's NSModalPanelWindowLevel)
	set response to (settingsAlert's runModal()) as integer
	-- 1つ目のボタンが1000、以降1ずつ増える
	set dialogButton to item (response - 1000 + 1) of buttonTitles
	if not dialogIsNewDocument then set dialogSlideText to (slideField's stringValue()) as text
	set dialogScaleText to (scaleField's stringValue()) as text
end showSettingsDialog:

-- 「ラベル・入力欄・単位」の1行を parentView の rowY の位置に置き、入力欄を返す
on addFieldRow(parentView, rowY, labelText, valueText, unitText)
	set labelField to current application's NSTextField's labelWithString:labelText
	labelField's setAlignment:(current application's NSTextAlignmentRight)
	labelField's setFrame:{{0, rowY + 3}, {90, 17}}
	parentView's addSubview:labelField

	set inputField to current application's NSTextField's alloc()'s initWithFrame:{{96, rowY}, {60, 22}}
	inputField's setStringValue:valueText
	parentView's addSubview:inputField

	set unitField to current application's NSTextField's labelWithString:unitText
	unitField's setFrame:{{162, rowY + 3}, {98, 17}}
	parentView's addSubview:unitField
	return inputField
end addFieldRow
