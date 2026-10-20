-- 待ち時間（秒）：動作が不安定なときは長めにする
property copyTimeout : 3 -- コピーがクリップボードに入るまで待つ上限（実行時のダイアログボックスの初期値）
property activateTimeout : 2 -- Keynoteが前面に出るまで待つ上限
property pasteTimeout : 5 -- ペーストした内容がスライドに入るまで待つ上限

-- アートボードサイズの長方形（塗り・線なし）を一緒にコピーするか（実行時のダイアログボックスの初期値）
property copyArtboardRect : true

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

-- コピーを待つ上限の秒数を毎回確認する
repeat
	set timeoutAnswer to text returned of (display dialog "コピーを待つ上限（秒）" default answer (copyTimeout as text) buttons {"キャンセル", "OK"} default button "OK" cancel button "キャンセル")
	try
		set timeoutValue to timeoutAnswer as number
		if timeoutValue > 0 then
			set copyTimeout to timeoutValue
			exit repeat
		end if
	end try
	display alert "0より大きい数字を入力してください。"
end repeat

-- アートボードの範囲を残すための長方形を一緒にコピーするか毎回確認する
if copyArtboardRect then
	set rectDefaultButton to "コピーする"
else
	set rectDefaultButton to "コピーしない"
end if
set rectAnswer to button returned of (display dialog "アートボードサイズの長方形（塗り・線なし）も一緒にコピーしますか？" buttons {"キャンセル", "コピーしない", "コピーする"} default button rectDefaultButton cancel button "キャンセル")
set copyArtboardRect to (rectAnswer is "コピーする")

set pasteCount to 0

set skippedArtboards to {}
set unconfirmedArtboards to {}

repeat with artboardIndex from 1 to artboardCount

	-- 前回の内容が貼られないよう、コピー前にクリップボードを空にする
	set the clipboard to ""

	-- Illustratorで対象アートボードをコピー（app.copy() は黙って失敗するのでメニューコマンドで）
	tell application "Adobe Illustrator"
		activate

		set jsCode to "
			var doc = app.activeDocument;
			doc.artboards.setActiveArtboardIndex(" & (artboardIndex - 1) & ");
			app.executeMenuCommand('deselectall');
			doc.selectObjectsOnActiveArtboard();
			app.redraw();
			var count = doc.selection.length;
			if (count > 0) {
				var tempLayer = null;
				var originalLayer = doc.activeLayer;
				if (" & copyArtboardRect & ") {
					// 一時レイヤーにアートボードサイズの長方形を描いて選択に加える
					var rect = doc.artboards[" & (artboardIndex - 1) & "].artboardRect;
					tempLayer = doc.layers.add();
					var boundsRect = tempLayer.pathItems.rectangle(rect[1], rect[0], rect[2] - rect[0], rect[1] - rect[3]);
					boundsRect.filled = false;
					boundsRect.stroked = false;
					boundsRect.selected = true;
					app.redraw();
				}
				app.executeMenuCommand('copy');
				app.redraw();
				if (tempLayer) {
					tempLayer.remove();
					doc.activeLayer = originalLayer;
				}
			}
			count;
		"

		set selectedCount to (do javascript jsCode) as integer
	end tell

	if selectedCount = 0 then
		-- 空のアートボードは飛ばす
		set end of skippedArtboards to artboardIndex
	else
		-- クリップボードにIllustratorの内容が入るまで最大 copyTimeout 秒待つ
		set isCopied to false
		repeat (round (copyTimeout * 10) rounding up) times
			if my clipboardHasGraphics() then
				set isCopied to true
				exit repeat
			end if
			delay 0.1
		end repeat

		if not isCopied then
			set end of skippedArtboards to artboardIndex
		else
			-- Keynoteに対象スライドを用意（既存スライドがあればそれを使い、無ければ末尾に追加）
			set targetSlideIndex to startSlideIndex + pasteCount
			set pasteCount to pasteCount + 1
			tell application "Keynote Creator Studio"
				activate

				tell keynoteDoc
					if targetSlideIndex ≤ (count of slides) then
						set targetSlide to slide targetSlideIndex
					else
						set targetSlide to make new slide at end of slides
					end if

					set current slide to targetSlide
					-- ペーストできたかを数の増減で判定するため、貼る前の数を控える
					set itemCountBefore to count of iWork items of targetSlide
				end tell
			end tell

			-- Keynoteが前面に出たらすぐペーストする
			repeat (round (activateTimeout / 0.05) rounding up) times
				if frontmost of application "Keynote Creator Studio" then exit repeat
				delay 0.05
			end repeat

			tell application "System Events"
				keystroke "v" using command down
			end tell

			-- スライド上のオブジェクトが増えたら次へ進む
			set isPasted to false
			repeat (round (pasteTimeout / 0.05) rounding up) times
				tell application "Keynote Creator Studio"
					if (count of iWork items of targetSlide) > itemCountBefore then set isPasted to true
				end tell
				if isPasted then exit repeat
				delay 0.05
			end repeat
			if not isPasted then set end of unconfirmedArtboards to artboardIndex
		end if
	end if
end repeat

-- Illustratorの状態を戻す
tell application "Adobe Illustrator"
	do javascript "
		var doc = app.activeDocument;
		app.executeMenuCommand('deselectall');
		doc.artboards.setActiveArtboardIndex(" & originalArtboardIndex & ");
	"
end tell

tell application "Keynote Creator Studio" to activate

if (count of skippedArtboards) = 0 and (count of unconfirmedArtboards) = 0 then
	display dialog "全アートボードをKeynoteへコピーしました。" buttons {"OK"} default button "OK"
else
	set resultMessage to "Keynoteへコピーしました。"
	set AppleScript's text item delimiters to ", "
	if (count of skippedArtboards) > 0 then set resultMessage to resultMessage & return & "空またはコピーできなかったアートボード : " & (skippedArtboards as text)
	if (count of unconfirmedArtboards) > 0 then set resultMessage to resultMessage & return & "ペーストを確認できなかったアートボード : " & (unconfirmedArtboards as text)
	set AppleScript's text item delimiters to ""
	display dialog resultMessage buttons {"OK"} default button "OK"
end if

-- クリップボードに文字列以外（PDFなどの図形データ）が入っているか
on clipboardHasGraphics()
	try
		set infoList to clipboard info
	on error
		return false
	end try
	repeat with infoItem in infoList
		set itemClass to item 1 of infoItem
		if itemClass is not in {string, Unicode text, «class utf8», «class ut16»} then return true
	end repeat
	return false
end clipboardHasGraphics
