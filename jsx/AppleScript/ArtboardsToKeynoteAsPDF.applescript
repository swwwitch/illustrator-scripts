-- アートボードごとにPDFを書き出し、Keynoteのスライドへ画像として挿入する
-- クリップボードとキー操作を使わないので、ArtboardsToKeynote より速い（貼られるのはPDF画像）

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
if isNewDocument then
	set startSlideIndex to 1
else
	set modeAnswer to button returned of (display dialog "貼り付け先を選んでください。" buttons {"キャンセル", "既存スライドに貼る", "末尾に追加"} default button "末尾に追加" cancel button "キャンセル")
	if modeAnswer is "末尾に追加" then
		set startSlideIndex to slideCount + 1
	else
		repeat
			set startAnswer to text returned of (display dialog "何枚目のスライドから貼り付けますか？（1〜" & slideCount & "）" default answer (currentSlideNumber as text) buttons {"キャンセル", "OK"} default button "OK" cancel button "キャンセル")
			try
				set startSlideIndex to startAnswer as integer
				if startSlideIndex ≥ 1 and startSlideIndex ≤ slideCount then exit repeat
			end try
			display alert "1〜" & slideCount & " の数字を入力してください。"
		end repeat
	end if
end if

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

				-- PDFを画像として挿入し、アートボードの大きさ（1pt＝1ポイント）にしてスライドの中央に置く
				tell targetSlide
					set pdfImage to make new image with properties {file:pdfFile}
					set width of pdfImage to artboardWidth
					set height of pdfImage to artboardHeight
					set position of pdfImage to {(slideWidth - artboardWidth) / 2, (slideHeight - artboardHeight) / 2}
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
