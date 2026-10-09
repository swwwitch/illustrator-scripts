(*
	SetSameGridColors
	［ファイル］→［ドキュメント設定...］で、透明グリッドのグリッドカラー2色を同じ色にそろえる。

	・gridHex が空なら、上側の色（gridColor1）を下側にもコピーする
	・gridHex に 16進数（例 "FFFFFF"）を入れると、2色ともその色にする
	・pressOK が true なら最後に［OK］を押す
	・色見本はアクションが無く System Events の click も効かないので、JXA の CGEvent で実クリックして
	  macOS 標準のカラーパネルを開き、16進数欄に書き込む。色見本への反映はパネルを閉じたとき

	動作確認：Illustrator 2026（30.8）日本語版 / macOS
	必要な許可：実行するアプリ（スクリプトエディタなど）にアクセシビリティの許可
*)

property gridHex : "" -- 例: "FFFFFF"。空なら上側の色に合わせる
property pressOK : true

tell application "Adobe Illustrator"
	activate
	if (count of documents) = 0 then
		display alert "ドキュメントが開かれていません。"
		return
	end if
end tell

tell application "System Events" to tell process "Adobe Illustrator"
	set frontmost to true
	-- ［ドキュメント設定］を開く（開いていなければ）
	if not (exists UI element "ドキュメント設定") then
		click menu item "ドキュメント設定..." of menu 1 of menu bar item "ファイル" of menu bar 1
	end if
	repeat 50 times
		if exists UI element "ドキュメント設定" then exit repeat
		delay 0.1
	end repeat
	if not (exists UI element "ドキュメント設定") then error "［ドキュメント設定］ダイアログボックスが開きませんでした。"
	set dlg to UI element "ドキュメント設定"
	set swatch1 to first UI element of dlg whose description is "gridColor1"
	set swatch2 to first UI element of dlg whose description is "gridColor2"
	set center1 to my centerOf(swatch1)
	set center2 to my centerOf(swatch2)
end tell

-- 上側の色見本：色を読み取る（gridHex 指定時はその色を書き込む）
set hexField to my openColorPanel(center1)
tell application "System Events"
	if gridHex is "" then
		set targetHex to value of hexField as text
	else
		set targetHex to gridHex
		my writeHex(hexField, targetHex)
	end if
end tell
my closeColorPanel()

-- 下側の色見本：同じ色を書き込む
set hexField to my openColorPanel(center2)
my writeHex(hexField, targetHex)
my closeColorPanel()

if pressOK then
	tell application "System Events" to tell process "Adobe Illustrator"
		click (first button of UI element "ドキュメント設定" whose description is "OK")
	end tell
end if

-- ===== 補助 =====

on centerOf(el)
	tell application "System Events"
		set {x, y} to value of attribute "AXPosition" of el
		set {w, h} to value of attribute "AXSize" of el
	end tell
	return {x + w div 2, y + h div 2}
end centerOf

(* 色見本は System Events の click では反応しないので、CGEvent で実クリックする *)
on realClick(pt)
	set js to "ObjC.import('CoreGraphics'); ObjC.import('Foundation');" & ¬
		"var p = $.CGPointMake(" & (item 1 of pt) & "," & (item 2 of pt) & ");" & ¬
		"[$.kCGEventMouseMoved, $.kCGEventLeftMouseDown, $.kCGEventLeftMouseUp].forEach(function (t) {" & ¬
		"  $.CGEventPost($.kCGHIDEventTap, $.CGEventCreateMouseEvent($(), t, p, $.kCGMouseButtonLeft));" & ¬
		"  $.NSThread.sleepForTimeInterval(0.08);" & ¬
		"});"
	run script js in "JavaScript"
end realClick

(* 色見本をクリックして「カラー」パネルを開き、RGBつまみの16進数欄を返す *)
on openColorPanel(pt)
	realClick(pt)
	tell application "System Events" to tell process "Adobe Illustrator"
		repeat 50 times
			if exists window "カラー" then exit repeat
			delay 0.1
		end repeat
		if not (exists window "カラー") then error "カラーパネルが開きませんでした。"
		set pnl to window "カラー"
		click (first button of toolbar 1 of pnl whose description is "カラーつまみ")
		delay 0.2
		set pb to pop up button 1 of splitter group 1 of pnl
		if (value of pb as text) is not "RGBつまみ" then
			click pb
			delay 0.2
			click menu item "RGBつまみ" of menu 1 of pb
			delay 0.2
		end if
		-- R・G・B 欄は3桁以内、16進数欄は6桁
		repeat with t in text fields of splitter group 1 of pnl
			if length of (value of t as text) is 6 then return contents of t
		end repeat
	end tell
	error "16進数カラー値の欄が見つかりませんでした。"
end openColorPanel

on writeHex(hexField, hexValue)
	tell application "System Events" to tell process "Adobe Illustrator"
		set focused of hexField to true
		set value of hexField to hexValue
		delay 0.1
		key code 36 -- return で確定
		delay 0.2
	end tell
end writeHex

(* パネルを閉じた時点で色見本に反映される *)
on closeColorPanel()
	tell application "System Events" to tell process "Adobe Illustrator"
		click (first button of window "カラー" whose description is "閉じるボタン")
		repeat 30 times
			if not (exists window "カラー") then exit repeat
			delay 0.1
		end repeat
	end tell
end closeColorPanel
