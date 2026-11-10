-- CopyToIllustratorScripts.app
-- アイコンにドロップしたファイル（フォルダー可）を、Illustrator の Presets.localized/ja_JP/スクリプト に複製する。
-- Illustrator が複数入っているときは、複製先のバージョンをリストから選ぶ（既定は年付きで最も新しいもの）。
-- ダブルクリックで起動したときは、ファイルを選ぶダイアログボックスを出す。
-- 複製先の Illustrator が起動中なら、再起動するか確認する（通常の終了→同じバージョンを起動）。
-- Copies files (or folders) dropped on the icon into Illustrator's Presets.localized/ja_JP/Scripts folder.
-- When several Illustrator versions are installed, pick the target versions from a list
-- (the newest year-named version is preselected). Double-clicking opens a file chooser instead.
-- If a target Illustrator is running, offers to relaunch it so the Scripts menu picks up the new files.
--
-- 作り方 / Build:
--   osacompile -o helpers/CopyToIllustratorScripts.app helpers/CopyToIllustratorScripts.applescript
--
-- 複製先は root の持ち物なので、書き込めないときは管理者の権限でコピーする。
-- root のプロセスは ~/Downloads などを読めない（TCC）ため、いったん /private/tmp に写してから運ぶ。
-- Illustrator はどのバージョンも同じ Bundle ID なので、起動中かどうかはアプリの場所で見分ける。
-- The Scripts folders are owned by root, so copying asks for an administrator password when needed.
-- Files are staged in /private/tmp first because root processes cannot read ~/Downloads etc. (TCC).
-- All Illustrator versions share one bundle ID, so running instances are matched by app path.
--
-- Released under the MIT license
-- http://opensource.org/licenses/mit-license.php

use AppleScript version "2.4"
use scripting additions
use framework "AppKit"

property dialogTitle : "Illustratorスクリプトへ複製"
-- 再起動で終了を待つ上限（秒）。未保存のドキュメントの確認に答える時間を見込む
property quitTimeout : 300
property scriptsSubpath : "Presets.localized/ja_JP/スクリプト"

on open droppedItems
	copyToIllustrator(droppedItems)
end open

on run
	try
		set picked to choose file with prompt "複製するファイルを選んでください（複数可）" with multiple selections allowed
	on error number -128
		return
	end try
	copyToIllustrator(picked)
end run

-- ドロップされた項目を、選んだ Illustrator のスクリプトフォルダーへ複製する
on copyToIllustrator(itemList)
	-- 元のパス（フォルダーは末尾の / を外す）
	set sources to {}
	repeat with anItem in itemList
		set p to POSIX path of anItem
		if p ends with "/" and length of p > 1 then set p to text 1 thru -2 of p
		set end of sources to p
	end repeat
	if sources is {} then return

	-- 複製先の候補
	set candidates to listCandidates()
	if candidates is {} then
		display alert dialogTitle message "スクリプトフォルダー（" & scriptsSubpath & "）のある Illustrator が見つかりません。" as critical
		return
	end if
	set labels to {}
	repeat with c in candidates
		set end of labels to item 1 of c
	end repeat

	-- バージョンの選択（1本だけなら確認なし）
	if (count of candidates) is 1 then
		set chosenLabels to labels
	else
		set namesText to ""
		repeat with s in sources
			set namesText to namesText & return & "・" & baseName(s)
		end repeat
		activate
		set chosenLabels to choose from list labels with title dialogTitle with prompt "複製先の Illustrator を選んでください（複数可）" & namesText default items {defaultLabel(labels)} OK button name "複製" cancel button name "キャンセル" with multiple selections allowed
		if chosenLabels is false then return
	end if

	set targets to {}
	repeat with c in candidates
		if chosenLabels contains (item 1 of c) then set end of targets to {item 1 of c, item 2 of c, item 3 of c}
	end repeat

	-- 上書きの確認
	set existing to {}
	repeat with t in targets
		repeat with s in sources
			set destPath to (item 2 of t) & "/" & baseName(s)
			if pathExists(destPath) then set end of existing to (item 1 of t) & " : " & baseName(s)
		end repeat
	end repeat
	if existing is not {} then
		activate
		try
			display alert dialogTitle message "次のファイルはすでにあります。上書きしますか？" & return & joinList(existing, return) buttons {"キャンセル", "上書き"} default button "キャンセル" cancel button "キャンセル"
		on error number -128
			return
		end try
	end if

	-- root のプロセスは ~/Downloads などを読めない（TCC）ので、一時フォルダーに写してから運ぶ
	set stageDir to do shell script "/usr/bin/mktemp -d /private/tmp/ai-scripts-copy.XXXXXX"
	try
		set userCmds to {}
		repeat with s in sources
			set end of userCmds to "/usr/bin/ditto " & quoted form of (s as text) & " " & quoted form of (stageDir & "/" & baseName(s))
		end repeat
		do shell script joinList(userCmds, " && ")

		set copyCmds to {}
		set needAdmin to false
		repeat with t in targets
			if not isWritable(item 2 of t) then set needAdmin to true
			repeat with s in sources
				set end of copyCmds to "/usr/bin/ditto " & quoted form of (stageDir & "/" & baseName(s)) & " " & quoted form of ((item 2 of t) & "/" & baseName(s))
			end repeat
		end repeat
		if needAdmin then
			do shell script joinList(copyCmds, " && ") with administrator privileges
		else
			do shell script joinList(copyCmds, " && ")
		end if
	on error errMsg number errNum
		removeStage(stageDir)
		if errNum is not -128 then display alert dialogTitle message "複製できませんでした。" & return & errMsg as critical
		return
	end try
	removeStage(stageDir)

	set doneLabels to {}
	repeat with t in targets
		set end of doneLabels to item 1 of t
	end repeat
	display notification ((count of sources) as text) & " 件 → " & joinList(doneLabels, ", ") with title dialogTitle

	offerRelaunch(targets)
end copyToIllustrator

-- 複製先のうち起動中の Illustrator があれば、再起動するか確認して再起動する
on offerRelaunch(targets)
	set runningTargets to {}
	repeat with t in targets
		if runningAppsAt(item 3 of t) is not {} then set end of runningTargets to t
	end repeat
	if runningTargets is {} then return

	set names to {}
	repeat with t in runningTargets
		set end of names to "Illustrator " & (item 1 of t)
	end repeat
	activate
	try
		display alert dialogTitle message "複製しました。スクリプトメニューに出すには再起動が必要です。" & return & return & joinList(names, return) & return & return & "再起動しますか？" buttons {"あとで", "再起動"} default button "再起動" cancel button "あとで"
	on error number -128
		return
	end try

	repeat with t in runningTargets
		relaunchApp(item 3 of t, "Illustrator " & (item 1 of t))
	end repeat
end offerRelaunch

-- 指定したパスの .app として起動しているプロセスの配列（同じ Bundle ID の別バージョンと区別するためパスで照合）
on runningAppsAt(appPath)
	set res to {}
	set apps to current application's NSWorkspace's sharedWorkspace()'s runningApplications()
	repeat with a in apps
		set u to a's bundleURL()
		if u is not missing value then
			if ((u's |path|()) as text) is appPath then set end of res to a
		end if
	end repeat
	return res
end runningAppsAt

-- 通常の終了（未保存の確認あり）を送り、終わったら起動し直す
on relaunchApp(appPath, displayName)
	repeat with a in runningAppsAt(appPath)
		a's terminate()
	end repeat
	set waited to 0
	repeat while runningAppsAt(appPath) is not {}
		if waited ≥ quitTimeout then
			display alert dialogTitle message displayName & " が終了しなかったため、起動し直しませんでした。" as warning
			return
		end if
		delay 0.5
		set waited to waited + 0.5
	end repeat
	do shell script "/usr/bin/open " & quoted form of appPath
end relaunchApp

-- スクリプトフォルダーのある Illustrator を {表示名, フォルダーのパス} の配列で返す
on listCandidates()
	set sh to "setopt null_glob
for d in /Applications/Adobe\\ Illustrator*(/); do
  dest=\"$d/" & scriptsSubpath & "\"
  [[ -d \"$dest\" ]] || continue
  app=(\"$d\"/*.app(N[1]))
  ver=$(defaults read \"$app/Contents/Info\" CFBundleShortVersionString 2>/dev/null)
  name=${d:t}; name=${name#Adobe Illustrator}; name=${name# }; name=${${name#\\(}%\\)}
  [[ -z \"$name\" ]] && name=\"年なし\"
  print -r -- \"$name ($ver)\"$'\\t'\"$dest\"$'\\t'\"$app\"
done"
	set out to do shell script "/bin/zsh -c " & quoted form of sh
	set res to {}
	if out is "" then return res
	set AppleScript's text item delimiters to tab
	repeat with aLine in paragraphs of out
		set end of res to {text item 1 of aLine, text item 2 of aLine, text item 3 of aLine}
	end repeat
	set AppleScript's text item delimiters to ""
	return res
end listCandidates

-- 既定は年付きで最も新しいもの
on defaultLabel(labels)
	set best to item -1 of labels
	repeat with l in labels
		if (l as text) starts with "20" then set best to l as text
	end repeat
	return best
end defaultLabel

on baseName(p)
	set AppleScript's text item delimiters to "/"
	set n to text item -1 of p
	set AppleScript's text item delimiters to ""
	return n
end baseName

on pathExists(p)
	try
		do shell script "/bin/test -e " & quoted form of p
		return true
	on error
		return false
	end try
end pathExists

on isWritable(p)
	try
		do shell script "/bin/test -w " & quoted form of p
		return true
	on error
		return false
	end try
end isWritable

on removeStage(stageDir)
	if stageDir starts with "/private/tmp/ai-scripts-copy." then do shell script "/bin/rm -rf " & quoted form of stageDir
end removeStage

on joinList(theList, delim)
	set AppleScript's text item delimiters to delim
	set s to theList as text
	set AppleScript's text item delimiters to ""
	return s
end joinList
