(*
	ShowControlPanel
	［ウィンドウ］→［コントロール］を表示にそろえる。

	・app.executeMenuCommand('drover control palette plugin') は反転しかできないので、
	  メニュー項目の ✓（AXMenuItemMarkChar）を読んでから、必要なときだけクリックする
	・setControlPanel(false) にすると非表示にそろえる。戻り値は変更前の状態（表示なら true）
	・Illustrator が背面にあっても動く。JSX の実行中はメニューバーに触れないので、JSX からは呼べない

	動作確認：Illustrator 2026（30.8）日本語版 / macOS
	必要な許可：実行するアプリ（スクリプトエディタなど）にアクセシビリティの許可
*)

on setControlPanel(wantVisible)
	tell application "System Events"
		tell (first process whose bundle identifier starts with "com.adobe.illustrator")
			set mi to menu item "コントロール" of menu "ウィンドウ" of menu bar item "ウィンドウ" of menu bar 1
			set isVisible to (value of attribute "AXMenuItemMarkChar" of mi) is not missing value
			if isVisible is not wantVisible then click mi
			return isVisible -- 変更前の状態
		end tell
	end tell
end setControlPanel

setControlPanel(true) -- 表示に揃える
