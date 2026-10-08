-- SetAiMenuState.app
-- Illustrator のメニュー項目のチェック状態（✓）を GUI スクリプティングで読み、指定どおりに揃える。
--   ［表示］→［スマートガイド］［ポイントにスナップ］、［ウィンドウ］→［コントロール］など、
--   executeMenuCommand() では反転しかできず現在の状態がわからない項目に使う。
-- Reads the checkmark of Illustrator menu items via GUI scripting and sets them as requested.
--   For items such as View > Smart Guides or Window > Control, which executeMenuCommand() can only toggle blindly.
--
-- 受け渡し / Handoff:
--   スクリプトが /tmp/set_ai_menu_state_request.txt に1行1件で「メニューの道筋<TAB>操作」を書いてから、このアプリを起動する。
--     道筋はメニュー名を「>」でつなぐ（例: 表示>スマートガイド、表示>ガイド>ガイドをロック）
--     名前が入れ替わる項目は「オフのときの名前|オンのときの名前」と書く（例: 表示>ガイド>ガイドを表示|ガイドを隠す）
--     環境設定のチェックボックスは「環境設定>パネルのメニュー名>チェックボックスの名前」と書く（on / off のみ。日本語版のみ）
--       例: 環境設定>一般...>裁ち落とし部分に「裁ち落としを印刷」生成 AI ボタンを表示
--       JavaScript で設定キーを書いても画面に反映されない項目に使う。名前は前方一致で探す
--     操作は get（読むだけ）/ on / off / toggle
--   結果は /tmp/set_ai_menu_state_result.txt に1行1件で「道筋<TAB>変更前<TAB>変更後」を書く。
--     状態は on / off、項目が無ければ notfound、グレー表示なら disabled
--   Illustrator はスクリプトの実行中にメニューを読ませない（$.sleep や app.redraw() で待っても不可）。
--   読めるのはスクリプトが終わった後か、モーダルのダイアログボックスを表示している間だけ。
--     on / off / toggle：呼び出し元のスクリプトが終わってから切り替える（最大30秒待つ）
--     get：呼び出し元が「SetAiMenuState」という名前のダイアログボックスを表示して待ち、
--          このアプリが結果を書いてからそのボタンを押して閉じる
--   The script writes "menu path<TAB>action" per line to the request file, then launches this app.
--     The path joins menu names with ">"; the action is get / on / off / toggle.
--     An item whose label flips is written "name while off|name while on" (e.g. View>Guides>Show Guides|Hide Guides).
--     A checkbox in the preferences dialog is written "環境設定>pane menu name>checkbox name" (on / off only, Japanese UI only),
--       for settings whose key does not take effect when written from JavaScript. The name is matched by prefix.
--   The result file gets "path<TAB>before<TAB>after" per line (on / off / notfound / disabled).
--   Illustrator blocks menu access while a script runs ($.sleep and app.redraw() do not help);
--   menus are reachable only after the script ends or while a modal dialog is shown.
--     on / off / toggle: applied after the calling script ends (waits up to 30 seconds)
--     get: the caller shows a dialog titled "SetAiMenuState"; this app writes the result, then presses its button
--
-- 作り方 / Build:
--   osacompile -o /Applications/SetAiMenuState.app helpers/SetAiMenuState.applescript
--
-- このアプリに［システム設定］→［プライバシーとセキュリティ］→［アクセシビリティ］の許可が要る。
-- 作り直すと署名が変わるため、許可もやり直しになる。
-- The app needs Accessibility permission (System Settings > Privacy & Security > Accessibility).
-- Rebuilding changes its signature, so the permission must be granted again.

property requestPath : "/tmp/set_ai_menu_state_request.txt"
property resultPath : "/tmp/set_ai_menu_state_result.txt"
property waitDialogName : "SetAiMenuState"

on run
	-- 読んだら消す（前回の依頼を黙って実行しないように）。前回の結果も消しておく
	-- Remove after reading so a stale request is never replayed; clear the previous result too
	try
		set requestText to do shell script "/bin/cat " & quoted form of requestPath & "; /bin/rm -f " & quoted form of requestPath & " " & quoted form of resultPath
	on error
		return
	end try
	if requestText is "" then return

	set resultLines to {}
	try
		-- 読むだけなら前面に出さない / Reading alone does not bring Illustrator to the front
		set p to illustratorProcess(requestText contains (tab & "on") or requestText contains (tab & "off") or requestText contains (tab & "toggle"))
		repeat with requestLine in paragraphs of requestText
			set requestLine to requestLine as text
			if requestLine is not "" then set end of resultLines to applyRequest(p, requestLine)
		end repeat
	on error errMsg number errNum
		set end of resultLines to "error" & tab & errMsg & " (" & errNum & ")"
	end try

	writeResult(joinList(resultLines, linefeed))
	closeWaitDialog()
end run

-- 必要なら Illustrator を前面に出し、メニューバーが取れるまで待ってプロセスを返す
-- スクリプトの実行中はメニューバーが取れない（-1719）ので、呼び出し元のスクリプトが終わるのを待つ
-- Brings Illustrator to the front when asked, and returns its process once the menu bar is available.
-- The menu bar is unreachable while a script is running (-1719), so this waits for the calling script to finish.
on illustratorProcess(shouldActivate)
	if shouldActivate then tell application id "com.adobe.illustrator" to activate
	tell application "System Events"
		set p to first process whose bundle identifier is "com.adobe.illustrator"
		repeat 300 times
			if exists menu bar 1 of p then return p
			delay 0.1
		end repeat
	end tell
	error "Illustrator のメニューバーを取得できませんでした。"
end illustratorProcess

-- 1行分の依頼を処理し、結果の行を返す / Handles one request line and returns its result line
on applyRequest(p, requestLine)
	set {menuPath, action} to splitText(requestLine, tab)
	if action is "" then set action to "get"
	if menuPath starts with "環境設定>" then return menuPath & tab & applyPrefCheckbox(menuPath, action)

	set {mi, stateBefore} to findToggle(p, menuPath)
	if mi is missing value then return menuPath & tab & "notfound" & tab & "notfound"
	-- ダイアログボックスの表示中はメニューがグレーになるが、読むだけなら構わない
	-- Menus are greyed out while a dialog is shown, which does not matter for reading
	if action is "get" then return menuPath & tab & stateBefore & tab & stateBefore
	tell application "System Events" to set isEnabled to enabled of mi
	if not isEnabled then return menuPath & tab & stateBefore & tab & "disabled"

	set wanted to stateBefore
	if action is "on" then
		set wanted to "on"
	else if action is "off" then
		set wanted to "off"
	else if action is "toggle" then
		if stateBefore is "on" then
			set wanted to "off"
		else
			set wanted to "on"
		end if
	end if

	set stateAfter to stateBefore
	if wanted is not stateBefore then
		tell application "System Events" to click mi
		-- ✓ や項目名の更新を待つ / Wait for the checkmark or the item name to update
		repeat 20 times
			set stateAfter to item 2 of findToggle(p, menuPath)
			if stateAfter is wanted then exit repeat
			delay 0.05
		end repeat
	end if
	return menuPath & tab & stateBefore & tab & stateAfter
end applyRequest

-- 道筋からメニュー項目と状態を返す。{項目, "on"|"off"}、見つからなければ {missing value, "notfound"}
-- 末尾が「ガイドを表示|ガイドを隠す」のように「|」で区切られていれば、名前が入れ替わる項目として扱う。
-- 左がオフのときの名前、右がオンのときの名前で、どちらが出ているかで状態を決める
-- Returns {item, "on"|"off"} for a path, or {missing value, "notfound"}.
-- A last segment such as "Show Guides|Hide Guides" names an item whose label flips:
-- the left name appears while off, the right one while on.
on findToggle(p, menuPath)
	set names to splitText(menuPath, ">")
	set lastName to last item of names
	if lastName does not contain "|" then
		set mi to findMenuItem(p, menuPath)
		if mi is missing value then return {missing value, "notfound"}
		return {mi, itemState(mi)}
	end if

	set parentPath to joinList(items 1 thru -2 of names, ">")
	set {offName, onName} to splitText(lastName, "|")
	set mi to findMenuItem(p, parentPath & ">" & onName)
	if mi is not missing value then return {mi, "on"}
	set mi to findMenuItem(p, parentPath & ">" & offName)
	if mi is not missing value then return {mi, "off"}
	return {missing value, "notfound"}
end findToggle

-- 「表示>ガイド>ガイドをロック」のような道筋からメニュー項目を探す。無ければ missing value
-- Finds a menu item from a path such as "View>Guides>Lock Guides"; returns missing value when absent
on findMenuItem(p, menuPath)
	set names to splitText(menuPath, ">")
	tell application "System Events"
		try
			set mi to menu bar item (item 1 of names) of menu bar 1 of p
			repeat with k from 2 to count of names
				set mi to menu item (item k of names) of menu 1 of mi
			end repeat
			return mi
		on error
			return missing value
		end try
	end tell
end findMenuItem

-- ✓ が付いていれば "on"、なければ "off" / "on" when the item has a checkmark, otherwise "off"
on itemState(mi)
	tell application "System Events"
		set markChar to value of attribute "AXMenuItemMarkChar" of mi
	end tell
	if markChar is missing value or markChar is "" then return "off"
	return "on"
end itemState

-- 環境設定のチェックボックスを揃え、「変更前<TAB>変更後」を返す
-- Sets a checkbox in the preferences dialog and returns "before<TAB>after"
on applyPrefCheckbox(menuPath, action)
	set names to splitText(menuPath, ">")
	set paneMenuName to item 2 of names
	set checkboxName to item 3 of names
	if action is "get" then return "notfound" & tab & "notfound"

	set w to openPrefPane(paneMenuName)
	tell application "System Events"
		try
			set cb to first checkbox of w whose description starts with checkboxName
		on error
			click (first button of w whose description is "キャンセル")
			my waitForPrefClosed()
			return "notfound" & tab & "notfound"
		end try
		if (value of cb as boolean) then
			set stateBefore to "on"
		else
			set stateBefore to "off"
		end if
		set stateAfter to stateBefore
		if (action is "toggle") or (action is not stateBefore) then
			click cb
			if stateBefore is "on" then
				set stateAfter to "off"
			else
				set stateAfter to "on"
			end if
		end if
		click (first button of w whose description is "OK")
	end tell
	waitForPrefClosed()
	return stateBefore & tab & stateAfter
end applyPrefCheckbox

-- 環境設定の指定パネルを開き、ダイアログボックスを返す（SetAiPreferences.applescript と同じ）
-- ダイアログボックスは window ではなく UI element として現れる。ラベルは title ではなく description に入っている
-- Opens the given preferences pane and returns the dialog (same as SetAiPreferences.applescript)
on openPrefPane(menuItemName)
	tell application "System Events"
		set p to first process whose bundle identifier is "com.adobe.illustrator"
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

-- 呼び出し元が待機用のダイアログボックスを出していれば、ボタンを押して閉じる
-- ダイアログボックスは window ではなく UI element（AXLayoutArea）として現れる
-- Presses the button of the caller's waiting dialog, if any (it appears as a UI element, not a window)
on closeWaitDialog()
	try
		tell application "System Events"
			set p to first process whose bundle identifier is "com.adobe.illustrator"
			-- 表示が間に合っていないことがあるので少し待つ / The dialog may not be up yet, so wait briefly
			repeat 20 times
				if exists UI element waitDialogName of p then
					click (first button of UI element waitDialogName of p)
					exit repeat
				end if
				delay 0.1
			end repeat
		end tell
	end try
end closeWaitDialog

-- 一時ファイルに書いてから移す（書きかけを読まれないように）
-- Writes to a temp file and moves it into place so a half-written result is never read
on writeResult(resultText)
	set tempPath to resultPath & ".part"
	do shell script "/usr/bin/printf %s " & quoted form of resultText & " > " & quoted form of tempPath & " && /bin/mv -f " & quoted form of tempPath & " " & quoted form of resultPath
end writeResult

-- 区切り文字で分割する。2要素未満なら空文字で補う / Splits by a delimiter, padding to at least two items
on splitText(sourceText, delimiter)
	set savedDelims to AppleScript's text item delimiters
	set AppleScript's text item delimiters to delimiter
	set parts to text items of sourceText
	set AppleScript's text item delimiters to savedDelims
	if (count of parts) < 2 then set end of parts to ""
	return parts
end splitText

on joinList(itemList, delimiter)
	set savedDelims to AppleScript's text item delimiters
	set AppleScript's text item delimiters to delimiter
	set joined to itemList as text
	set AppleScript's text item delimiters to savedDelims
	return joined
end joinList
