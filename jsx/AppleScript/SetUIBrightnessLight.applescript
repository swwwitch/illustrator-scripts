(*
	SetUIBrightnessLight
	［環境設定］→［ユーザーインターフェイス］の明るさを「明」にする。

	・JavaScript の setRealPreference('uiBrightness', …) は再起動まで画面に反映されないので、
	  環境設定ダイアログボックスのボタンを押して切り替える
	・明るさのボタンは description が「暗」「やや暗め」「やや明るめ」「明」、value が "Selected" / "Not Selected"
	・すでに「明」なら押さずに［OK］で閉じる

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
	click menu item "ユーザーインターフェイス..." of menu 1 of menu item "設定…" of menu 1 of menu bar item 2 of menu bar 1 of p
	repeat 50 times
		if exists UI element "環境設定" of p then exit repeat
		delay 0.1
	end repeat
	set w to UI element "環境設定" of p
	set brightnessButton to first button of w whose description is "明"
	if (value of brightnessButton as text) is not "Selected" then click brightnessButton
	click (first button of w whose description is "OK")
end tell
