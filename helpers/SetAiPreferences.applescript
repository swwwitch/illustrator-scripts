-- SetAiPreferences.app
-- Illustrator の環境設定を GUI スクリプティングで変更する。
--   ［ファイル管理］：［バックグラウンドで保存］［バックグラウンドで書き出し］［クラウドドキュメントを次の間隔で自動保存］を OFF
--   ［ユーザーインターフェイス］：［明るさ］を「明」
-- Changes Illustrator preferences via GUI scripting:
--   File Handling: turns off "Save in Background", "Export in Background" and cloud document autosave
--   User Interface: sets Brightness to Light
--
-- app.preferences.setBooleanPreference() で書いても再起動まで反映されないため、
-- ダイアログボックスのコントロールを直接操作する。
-- Writing the preference keys from JavaScript only takes effect after a restart,
-- so this operates the controls in the dialog instead.
--
-- 作り方 / Build:
--   osacompile -o /Applications/SetAiPreferences.app helpers/SetAiPreferences.applescript
--
-- このアプリに［システム設定］→［プライバシーとセキュリティ］→［アクセシビリティ］の許可が要る。
-- 作り直すと署名が変わるため、許可もやり直しになる。
-- The app needs Accessibility permission (System Settings > Privacy & Security > Accessibility).
-- Rebuilding changes its signature, so the permission must be granted again.

on run
	try
		-- ファイル管理 / File Handling
		set w to openPrefPane("ファイル管理...")
		tell application "System Events"
			-- クラウドドキュメントのラベルは末尾に「 :」と空白が付くので前方一致で探す
			-- Labels are matched by prefix; the cloud document one ends with " :" and spaces
			repeat with labelName in {"バックグラウンドで保存", "バックグラウンドで書き出し", "クラウドドキュメントを次の間隔で自動保存"}
				set cb to first checkbox of w whose description starts with (labelName as text)
				if (value of cb as boolean) then click cb
			end repeat
			click (first button of w whose description is "OK")
		end tell
		waitForPrefClosed()

		-- ユーザーインターフェイス / User Interface
		set w to openPrefPane("ユーザーインターフェイス...")
		tell application "System Events"
			set brightnessButton to first button of w whose description is "明"
			if (value of brightnessButton as text) is not "Selected" then click brightnessButton
			click (first button of w whose description is "OK")
		end tell
		waitForPrefClosed()
	on error errMsg number errNum
		display dialog "Illustrator の環境設定を変更できませんでした。" & return & return & errMsg & " (" & errNum & ")" buttons {"OK"} default button 1 with icon caution giving up after 10
	end try
end run

-- 環境設定の指定パネルを開き、ダイアログボックスを返す
-- ダイアログボックスは window ではなく UI element として現れる。ラベルは title ではなく description に入っている
-- Opens the given preferences pane and returns the dialog (a UI element, not a window; labels live in description)
on openPrefPane(menuItemName)
	tell application id "com.adobe.illustrator" to activate
	delay 0.3
	tell application "System Events"
		set p to first process whose bundle identifier is "com.adobe.illustrator"
		-- 前面に来るまではメニューバーが取れない / The menu bar is unavailable until Illustrator comes to the front
		repeat 50 times
			if exists menu bar 1 of p then exit repeat
			delay 0.1
		end repeat
		-- 直前に閉じたダイアログボックスの後はメニューが効かないことがあるので、開かなければもう一度押す
		-- The menu click can be ignored right after a dialog closes, so retry once
		repeat 2 times
			-- 「設定…」は1文字の「…」、パネル名はピリオド3つ / "設定…" uses an ellipsis, pane names use three periods
			click menu item menuItemName of menu 1 of menu item "設定…" of menu 1 of menu bar item 2 of menu bar 1 of p
			repeat 50 times
				if exists UI element "環境設定" of p then exit repeat
				delay 0.1
			end repeat
			if exists UI element "環境設定" of p then
				set w to UI element "環境設定" of p
				-- 開いた直後は前のパネルの中身が残っていることがあるので、見出しがパネル名になるまで待つ
				-- Right after opening, the previous pane's controls can linger; wait for this pane's heading
				set paneTitle to text 1 thru -4 of menuItemName
				repeat 50 times
					if exists (first static text of w whose value is paneTitle) then return w
					delay 0.1
				end repeat
				error "環境設定のパネルが切り替わりませんでした（" & paneTitle & "）。"
			end if
		end repeat
	end tell
	error "環境設定のダイアログボックスが開きませんでした（" & menuItemName & "）。"
end openPrefPane

-- 環境設定のダイアログボックスが閉じるまで待つ / Waits until the preferences dialog has closed
on waitForPrefClosed()
	tell application "System Events"
		set p to first process whose bundle identifier is "com.adobe.illustrator"
		repeat 50 times
			if not (exists UI element "環境設定" of p) then exit repeat
			delay 0.1
		end repeat
	end tell
	delay 0.5
end waitForPrefClosed
