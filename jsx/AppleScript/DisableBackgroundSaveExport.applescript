(*
	DisableBackgroundSaveExport
	［環境設定］→［ファイル管理］の「バックグラウンドで保存」「バックグラウンドで書き出し」をオフにする。

	・JavaScript で enableBackgroundSave / enableBackgroundExport を書いても再起動まで効かないので、
	  環境設定ダイアログボックスのチェックボックスをクリックする
	・ダイアログボックスは window ではなく UI element "環境設定"（AXLayoutArea）として現れ、
	  チェックボックスの名前は title ではなく description に入っている
	・オンのものだけクリックし、最後に［OK］を押す

	動作確認：Illustrator 2026（30.8）日本語版 / macOS
	必要な許可：実行するアプリ（スクリプトエディタなど）にアクセシビリティの許可
*)

tell application id "com.adobe.illustrator" to activate
tell application "System Events"
	set p to first process whose bundle identifier is "com.adobe.illustrator"
	click menu item "ファイル管理..." of menu 1 of menu item "設定…" of menu 1 of menu bar item 2 of menu bar 1 of p
	repeat 30 times
		if exists UI element "環境設定" of p then exit repeat
		delay 0.1
	end repeat
	set w to UI element "環境設定" of p
	repeat with labelName in {"バックグラウンドで保存", "バックグラウンドで書き出し"}
		set cb to (first checkbox of w whose description is (labelName as text))
		if (value of cb as boolean) then click cb
	end repeat
	click (first button of w whose description is "OK")
end tell
