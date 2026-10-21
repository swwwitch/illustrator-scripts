-- 作業中のアートボードをPDFで書き出し、Keynoteの表示中のスライドへ画像としてオーバーレイする
-- ダイアログボックスを出さない版（ArtboardsToKeynoteAsPDF の既定の動作を、スライドに収まる大きさで行う）

-- =========================================
-- 基本情報 / Basic info
-- =========================================
-- SCRIPT_NAME     : CurrentArtboardToKeynoteAsPDF
-- SCRIPT_VERSION  : v1.0.1
-- SCRIPT_AUTHOR   : Masahiro Takano (@swwwitch)
-- SCRIPT_RELEASED : 2026-10-21
-- SCRIPT_UPDATED  : 2026-10-22
--
-- SCRIPT_README_JA : https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CurrentArtboardToKeynoteAsPDF.md
-- SCRIPT_README_EN : https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CurrentArtboardToKeynoteAsPDF.md
--
-- Released under the MIT license
-- http://opensource.org/licenses/mit-license.php

use AppleScript version "2.4"
use scripting additions

tell application "Adobe Illustrator"
	if (count of documents) = 0 then
		display alert "Illustratorでドキュメントを開いてください。"
		return
	end if
end tell

tell application "Keynote Creator Studio"
	if (count of documents) = 0 then
		set keynoteDoc to make new document
	else
		set keynoteDoc to front document
	end if
	-- 表示中のスライドにオーバーレイする（取れないときはスライド1）
	-- slide number は非表示スライドを数えないので、番号ではなく参照で控える
	try
		set targetSlide to current slide of keynoteDoc
	on error
		set targetSlide to missing value
	end try
	-- 画像をスライドに収めて中央に置くため、スライドの大きさを控える
	set slideWidth to width of keynoteDoc
	set slideHeight to height of keynoteDoc
end tell

-- 書き出し用の一時フォルダー
set exportFolder to do shell script "mktemp -d -t CurrentArtboardToKeynote"

tell application "Adobe Illustrator"
	-- PDFの範囲をアートボード全体にするため、一時レイヤーにアートボードサイズの透明な長方形を描いてから書き出す
	set jsResult to do javascript "
		var doc = app.activeDocument;
		var artboardIndex = doc.artboards.getActiveArtboardIndex();
		app.executeMenuCommand('deselectall');
		doc.selectObjectsOnActiveArtboard();
		var count = doc.selection.length;
		app.executeMenuCommand('deselectall');
		var rect = doc.artboards[artboardIndex].artboardRect;
		if (count > 0) {
			var originalLayer = doc.activeLayer;
			var tempLayer = null;
			try {
				tempLayer = doc.layers.add();
				var boundsRect = tempLayer.pathItems.rectangle(rect[1], rect[0], rect[2] - rect[0], rect[1] - rect[3]);
				boundsRect.filled = false;
				boundsRect.stroked = false;
				var exportItem = new ExportForScreensItemToExport();
				exportItem.document = false;
				exportItem.artboards = String(artboardIndex + 1);
				doc.exportForScreens(new Folder('" & exportFolder & "'), ExportForScreensType.SE_PDF, new ExportForScreensPDFOptions(), exportItem, '');
			} finally {
				if (tempLayer) tempLayer.remove();
				doc.activeLayer = originalLayer;
			}
		}
		[count, rect[2] - rect[0], rect[1] - rect[3]].join(',');
	"
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
	do shell script "rm -rf " & quoted form of exportFolder
	display alert "アートボードが空か、書き出しに失敗しました。"
	return
end if
set pdfFile to POSIX file pdfPath as alias

-- スライドに収まる最大の大きさにする
set scaleFactor to slideWidth / artboardWidth
if slideHeight / artboardHeight < scaleFactor then set scaleFactor to slideHeight / artboardHeight
set imageWidth to artboardWidth * scaleFactor
set imageHeight to artboardHeight * scaleFactor

tell application "Keynote Creator Studio"
	tell keynoteDoc
		if targetSlide is missing value then
			if (count of slides) > 0 then
				set targetSlide to slide 1
			else
				set targetSlide to make new slide at end of slides
			end if
		end if
		tell targetSlide
			set pdfImage to make new image with properties {file:pdfFile}
			set width of pdfImage to imageWidth
			set height of pdfImage to imageHeight
			set position of pdfImage to {(slideWidth - imageWidth) / 2, (slideHeight - imageHeight) / 2}
		end tell
	end tell
end tell

do shell script "rm -rf " & quoted form of exportFolder

tell application "Keynote Creator Studio"
	activate
	set current slide of keynoteDoc to targetSlide
end tell
