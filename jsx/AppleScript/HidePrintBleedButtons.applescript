(*
	HidePrintBleedButtons
	［環境設定］→［一般］の「裁ち落とし部分に「裁ち落としを印刷」生成 AI ボタンを表示」をオフにする。

	・設定キー enablePrintBleedWidget を書いてもカンバスのボタンは消えないので、
	  環境設定ダイアログボックスのチェックボックスをクリックする
	・項目名は「生成 AI」の間に半角スペースが入る
	・開いた直後は前のパネルの中身が残っていることがあるので、見出しが「一般」になるまで待つ

	動作確認：Illustrator 2026（30.8）日本語版 / macOS
	必要な許可：実行するアプリ（スクリプトエディタなど）にアクセシビリティの許可
*)

tell application id "com.adobe.illustrator" to activate
delay 0.3
tell application "System Events"
	set p to first process whose bundle identifier is "com.adobe.illustrator"
	repeat 50 times
		if exists menu bar 1 of p then exit repeat
		delay 0.1
	end repeat
	click menu item "一般..." of menu 1 of menu item "設定…" of menu 1 of menu bar item 2 of menu bar 1 of p
	repeat 50 times
		if exists UI element "環境設定" of p then exit repeat
		delay 0.1
	end repeat
	set w to UI element "環境設定" of p
	-- 開いた直後は前のパネルの中身が残っていることがあるので、見出しが「一般」になるまで待つ
	repeat 50 times
		if exists (first static text of w whose value is "一般") then exit repeat
		delay 0.1
	end repeat
	set cb to first checkbox of w whose description is "裁ち落とし部分に「裁ち落としを印刷」生成 AI ボタンを表示"
	if (value of cb as boolean) then click cb
	click (first button of w whose description is "OK")
end tell
