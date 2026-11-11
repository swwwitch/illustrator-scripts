-- AiMenuDiff.scpt
-- Illustrator のメニュー項目をすべて書き出し、別のバージョンで書き出したものと比べて、
-- 追加・削除・改名・移動・ショートカットの変化を HTML で表示する。
-- Writes every Illustrator menu item to a snapshot file and compares it with a snapshot taken
-- in another version: added, removed, renamed, moved items and shortcut changes, shown as HTML.
--
-- 使い方 / Usage:
--   1. 比べたい2つのバージョンそれぞれのスクリプトメニューに入れる（CopyToIllustratorScripts.app にドロップ）。
--   2. 片方の Illustrator で［ファイル］>［スクリプト］>［AiMenuDiff］を実行 → スナップショットを保存。
--   3. もう片方でも実行 → 保存後に比較相手を選ぶと、レポートがブラウザーで開く。
--   スクリプトエディタや osascript で実行したときは、保存済みのスナップショットを2つ選んで比べる。
--   1. Put this script into the Scripts menu of both versions.
--   2. Run it in one version to save a snapshot; run it in the other and pick the first to compare.
--   Run outside Illustrator, it compares two saved snapshots.
--
-- 作り方 / Build:
--   osacompile -o helpers/AiMenuDiff.scpt helpers/AiMenuDiff.applescript
--
-- メニューは Illustrator のプロセス内の NSMenu から読むため、スクリプトメニューから実行する必要がある。
-- ドキュメントの［ウィンドウ］メニューの項目と、メールアドレスは書き出さない。
-- ショートカットは、その時点で選ばれているキーボードショートカットのセットの値。
-- The menu is read from NSMenu inside the Illustrator process, so it must run from the Scripts menu.
--
-- メニューを辿る処理は sttk3 の sttk3_menu_titles.scpt（© 2024 sttk3.com）を元にした。
-- Menu traversal is based on sttk3_menu_titles.scpt by sttk3.com.

use AppleScript version "2.4"
use scripting additions
use framework "Foundation"
use framework "AppKit"

property scriptTitle : "Illustratorメニュー比較"
property scriptVersion : "1.0.0"

-- パスの区切り（SetAiMenuState と同じ「>」）
property pathSeparator : ">"

-- スナップショットとレポートの置き場所。無ければ予備の場所を使う
property snapshotFolderPath : "~/sw Dropbox/takano masahiro/Dropbox-shared/AT-doc/menu-titles"
property fallbackFolderPath : "~/Documents/AiMenuTitles"
property snapshotExtension : "tsv"

-- 子孫を辿らない項目（中身が環境やユーザーごとに異なるもの）。区切りを変えたらここも合わせる
property excludeDescendants : {"Apple>Services", "File>Open Recent Files", "File>Scripts", "Type>Font", "Type>Recent Fonts", "Type>Size", "Type>Insert Special Character", "Type>Insert WhiteSpace Character", "Type>Insert Break Character", "View>Rotate View", "Window>Extensions", "Window>Graphic Style Libraries", "Window>Symbol Libraries", "Window>Swatch Libraries", "Window>Brush Libraries", "Apple>サービス", "ファイル>最近使用したファイルを開く", "ファイル>スクリプト", "書式>フォント", "書式>最近使用したフォント", "書式>サイズ", "書式>特殊文字を挿入", "書式>空白文字を挿入", "書式>分割文字を挿入", "表示>ビューを回転", "ウィンドウ>エクステンション", "ウィンドウ>グラフィックスタイルライブラリ", "ウィンドウ>シンボルライブラリ", "ウィンドウ>スウォッチライブラリ", "ウィンドウ>ブラシライブラリ"}

on run
	try
		set snapshotFolder to prepareSnapshotFolder()
		if isRunningInIllustrator() then
			runInIllustrator(snapshotFolder)
		else
			runStandalone(snapshotFolder)
		end if
	on error errorMessage number errorNumber
		if errorNumber is -128 then return
		display alert scriptTitle message errorMessage & " (" & errorNumber & ")" as critical
	end try
end run

-- ===== 実行の流れ =====

-- Illustrator 内：書き出してから、比較相手を選ぶ
on runInIllustrator(snapshotFolder)
	set captured to captureMenus(snapshotFolder)
	set currentName to snapshotName of captured
	set savedText to "保存しました：" & currentName & "（" & (itemCount of captured) & "項目）"
	set otherNames to listSnapshotNames(snapshotFolder, currentName)
	if otherNames is {} then
		set answer to display dialog savedText & return & return & "別のバージョンの Illustrator でも実行すると、2つを比較できます。" buttons {"フォルダーを開く", "OK"} default button "OK" with title scriptTitle
		if button returned of answer is "フォルダーを開く" then do shell script "open " & quoted form of snapshotFolder
		return
	end if
	set picked to choose from list otherNames with title scriptTitle with prompt savedText & return & "比較するバージョンを選んでください" default items {pickDefaultCounterpart(otherNames, currentName)} OK button name "比較" cancel button name "比較しない"
	if picked is false then return
	set otherName to item 1 of picked
	if compareNames(otherName, currentName) > 0 then
		compareAndOpen(snapshotFolder, currentName, otherName)
	else
		compareAndOpen(snapshotFolder, otherName, currentName)
	end if
end runInIllustrator

-- Illustrator の外：保存済みのスナップショットを2つ選ぶ
on runStandalone(snapshotFolder)
	set allNames to listSnapshotNames(snapshotFolder, "")
	if (count of allNames) < 2 then
		display alert scriptTitle message "比較するにはスナップショットが2つ必要です（いまは " & (count of allNames) & " 個）。" & return & return & "このスクリプトを Illustrator のスクリプトメニューに入れ、比べたいバージョンそれぞれで実行してください。" & return & return & snapshotFolder
		return
	end if
	set defaultItems to items -2 thru -1 of allNames
	repeat
		set picked to choose from list allNames with title scriptTitle with prompt "比較するスナップショットを2つ選んでください（⌘クリック）" default items defaultItems OK button name "比較" with multiple selections allowed
		if picked is false then return
		if (count of picked) is 2 then exit repeat
		display alert scriptTitle message "2つ選んでください。"
		set defaultItems to picked
	end repeat
	set firstName to item 1 of picked
	set secondName to item 2 of picked
	if compareNames(firstName, secondName) > 0 then
		compareAndOpen(snapshotFolder, secondName, firstName)
	else
		compareAndOpen(snapshotFolder, firstName, secondName)
	end if
end runStandalone

-- 2つを比べてレポートを書き、ブラウザーで開く
on compareAndOpen(snapshotFolder, oldName, newName)
	set reportFolder to snapshotFolder & "/diff"
	makeFolder(reportFolder)
	set reportPath to reportFolder & "/" & oldName & " - " & newName & ".html"
	writeReport(snapshotPathFor(snapshotFolder, oldName), snapshotPathFor(snapshotFolder, newName), reportPath)
	do shell script "open " & quoted form of reportPath
	return reportPath
end compareAndOpen

-- ===== 書き出し =====

-- 実行中の Illustrator のメニューをスナップショットに書く
on captureMenus(snapshotFolder)
	set mainBundle to current application's NSBundle's mainBundle()
	set appVersion to (mainBundle's objectForInfoDictionaryKey:"CFBundleShortVersionString") as text
	set appPath to mainBundle's bundlePath()
	set installFolder to appPath's stringByDeletingLastPathComponent()
	set folderName to (installFolder's lastPathComponent()) as text
	set channelName to ""
	if folderName contains "Beta" then set channelName to " Beta"
	if folderName contains "Prerelease" then set channelName to " Prerelease"
	set localeName to detectLocale(installFolder)
	set captureName to "Ai" & appVersion & channelName & " " & localeName

	set rows to collectMenuRows(current application's NSApp's mainMenu(), "")
	set metaPairs to {{"name", captureName}, {"app", appPath as text}, {"version", appVersion}, {"locale", localeName}, {"date", timestampText()}, {"tool", "AiMenuDiff " & scriptVersion}}
	set filePath to snapshotPathFor(snapshotFolder, captureName)
	writeSnapshot(filePath, metaPairs, rows)
	return {snapshotName:captureName, filePath:filePath, itemCount:count of rows}
end captureMenus

-- メニューの子孫を {itemPath, shortcut, flags} の並びにする（区切り線は除く）
on collectMenuRows(theMenu, parentPath)
	set rows to {}
	set itemList to (theMenu's itemArray()) as list
	repeat with itemRef in itemList
		set menuItem to contents of itemRef
		if not ((menuItem's isSeparatorItem()) as boolean) then
			set itemTitle to cleanTitle((menuItem's title()) as text)
			if itemTitle is not "" then
				if parentPath is "" then
					set itemPath to itemTitle
				else
					set itemPath to parentPath & pathSeparator & itemTitle
				end if
				set end of rows to {itemPath:itemPath, shortcut:formatShortcut(menuItem), flags:itemFlags(menuItem)}
				if itemPath is not in excludeDescendants and ((menuItem's hasSubmenu()) as boolean) then
					set rows to rows & collectMenuRows(menuItem's submenu(), itemPath)
				end if
			end if
		end if
	end repeat
	return rows
end collectMenuRows

-- 書き出さない名前は "" を返す。メールアドレスは消す
on cleanTitle(itemTitle)
	if itemTitle is "" then return "（名前なし）"
	-- 開いているドキュメントの名前（例：名称未設定-1 @ 100% (RGB/プレビュー)）
	if matchesPattern(itemTitle, " @ [0-9.,]+ ?%") then return ""
	if itemTitle does not contain "@" then return itemTitle
	set titleText to current application's NSString's stringWithString:itemTitle
	set titleText to replacePattern(titleText, "\\s*[(（][^()（）]*@[^()（）]*[)）]", "")
	set titleText to replacePattern(titleText, "\\S+@\\S+\\.\\S+", "(email)")
	return titleText as text
end cleanTitle

-- ショートカットを ⌃⌥⇧⌘ ＋キーの形にする
on formatShortcut(menuItem)
	set keyText to (menuItem's keyEquivalent()) as text
	if keyText is "" then return ""
	set maskValue to (menuItem's keyEquivalentModifierMask()) as integer
	set isUpper to false
	considering case
		if keyText is not (((current application's NSString's stringWithString:keyText)'s lowercaseString()) as text) then set isUpper to true
	end considering
	set modifierText to ""
	if (maskValue div 262144) mod 2 is 1 then set modifierText to modifierText & "⌃"
	if (maskValue div 524288) mod 2 is 1 then set modifierText to modifierText & "⌥"
	if (maskValue div 131072) mod 2 is 1 or isUpper then set modifierText to modifierText & "⇧"
	if (maskValue div 1048576) mod 2 is 1 then set modifierText to modifierText & "⌘"
	return modifierText & keyLabel(keyText)
end formatShortcut

on keyLabel(keyText)
	set code to id of character 1 of keyText
	if code is 13 or code is 3 then return "↩"
	if code is 9 then return "⇥"
	if code is 27 then return "⎋"
	if code is 8 or code is 127 then return "⌫"
	if code is 32 then return "Space"
	if code is 63232 then return "↑"
	if code is 63233 then return "↓"
	if code is 63234 then return "←"
	if code is 63235 then return "→"
	if code ≥ 63236 and code ≤ 63270 then return "F" & (code - 63235)
	if code is 63272 then return "⌦"
	if code is 63273 then return "↖"
	if code is 63275 then return "↘"
	if code is 63276 then return "⇞"
	if code is 63277 then return "⇟"
	return ((current application's NSString's stringWithString:keyText)'s uppercaseString()) as text
end keyLabel

on itemFlags(menuItem)
	set flagList to {}
	if (menuItem's isHidden()) as boolean then set end of flagList to "hidden"
	if (menuItem's isAlternate()) as boolean then set end of flagList to "alternate"
	return joinText(flagList, ",")
end itemFlags

-- 1行目からメタ情報（#キー<TAB>値）、続けて「パス<TAB>ショートカット<TAB>状態」。同じパスは「 #2」を付けて区別する
on writeSnapshot(filePath, metaPairs, rows)
	set lineList to {}
	repeat with pair in metaPairs
		set end of lineList to "#" & (item 1 of pair) & tab & (item 2 of pair)
	end repeat
	set seenCounts to current application's NSMutableDictionary's dictionary()
	repeat with rowRef in rows
		set itemPath to itemPath of rowRef
		set seen to (seenCounts's objectForKey:itemPath)
		if seen is missing value then
			(seenCounts's setObject:1 forKey:itemPath)
		else
			set seen to (seen as integer) + 1
			(seenCounts's setObject:seen forKey:itemPath)
			set itemPath to itemPath & " #" & seen
		end if
		set end of lineList to itemPath & tab & (shortcut of rowRef) & tab & (flags of rowRef)
	end repeat
	writeTextFile(filePath, joinText(lineList, linefeed) & linefeed)
end writeSnapshot

-- ===== 読み込みと比較 =====

-- {meta, paths, shortcuts, flags}（meta・shortcuts・flags は NSDictionary、paths は並び順のリスト）
on readSnapshot(filePath)
	set fileText to current application's NSString's stringWithContentsOfFile:filePath encoding:(current application's NSUTF8StringEncoding) |error|:(missing value)
	if fileText is missing value then error "スナップショットを読めません：" & filePath
	set metaDict to current application's NSMutableDictionary's dictionary()
	set shortcutDict to current application's NSMutableDictionary's dictionary()
	set flagDict to current application's NSMutableDictionary's dictionary()
	set pathList to {}
	set lineList to (fileText's componentsSeparatedByString:linefeed) as list
	repeat with lineRef in lineList
		set lineText to contents of lineRef
		if lineText is not "" then
			set fields to splitText(lineText, tab)
			if lineText starts with "#" then
				if (count of fields) ≥ 2 then (metaDict's setObject:(item 2 of fields) forKey:(text 2 thru -1 of (item 1 of fields)))
			else
				set itemPath to item 1 of fields
				set end of pathList to itemPath
				set shortcutText to ""
				set flagText to ""
				if (count of fields) ≥ 2 then set shortcutText to item 2 of fields
				if (count of fields) ≥ 3 then set flagText to item 3 of fields
				(shortcutDict's setObject:shortcutText forKey:itemPath)
				(flagDict's setObject:flagText forKey:itemPath)
			end if
		end if
	end repeat
	return {meta:metaDict, paths:pathList, shortcuts:shortcutDict, flags:flagDict}
end readSnapshot

-- 差分：added / removed / renamed（同じ親で1対1）/ moved（同じ名前で1対1）/ shortcutChanges / flagChanges
on compareSnapshots(oldSnap, newSnap)
	set oldSet to current application's NSSet's setWithArray:(paths of oldSnap)
	set newSet to current application's NSSet's setWithArray:(paths of newSnap)

	set removedPaths to {}
	set shortcutChanges to {}
	set flagChanges to {}
	repeat with pathRef in (paths of oldSnap)
		set itemPath to contents of pathRef
		if (newSet's containsObject:itemPath) as boolean then
			set oldShortcut to ((shortcuts of oldSnap)'s objectForKey:itemPath) as text
			set newShortcut to ((shortcuts of newSnap)'s objectForKey:itemPath) as text
			considering case
				if oldShortcut is not newShortcut then set end of shortcutChanges to {itemPath, oldShortcut, newShortcut}
			end considering
			set oldFlags to ((flags of oldSnap)'s objectForKey:itemPath) as text
			set newFlags to ((flags of newSnap)'s objectForKey:itemPath) as text
			if oldFlags is not newFlags then set end of flagChanges to {itemPath, oldFlags, newFlags}
		else
			set end of removedPaths to itemPath
		end if
	end repeat
	set addedPaths to {}
	repeat with pathRef in (paths of newSnap)
		set itemPath to contents of pathRef
		if not ((oldSet's containsObject:itemPath) as boolean) then set end of addedPaths to itemPath
	end repeat

	-- 同じ名前が削除側・追加側に1つずつ → 移動
	set movedPairs to pairUp(removedPaths, addedPaths, "leaf")
	set removedPaths to remainingPaths(removedPaths, movedPairs, 1)
	set addedPaths to remainingPaths(addedPaths, movedPairs, 2)
	-- 子が移った先のサブメニュー → 改名（名前が同じなら移動）
	set renamedPairs to {}
	set childPairs to pairByChildren(removedPaths, addedPaths, movedPairs)
	repeat with pairRef in childPairs
		if leafOf(item 1 of pairRef) is leafOf(item 2 of pairRef) then
			set end of movedPairs to contents of pairRef
		else
			set end of renamedPairs to contents of pairRef
		end if
	end repeat
	set removedPaths to remainingPaths(removedPaths, childPairs, 1)
	set addedPaths to remainingPaths(addedPaths, childPairs, 2)
	-- 同じ親の下で削除・追加が1つずつ → 改名
	set parentPairs to pairUp(removedPaths, addedPaths, "parent")
	set renamedPairs to renamedPairs & parentPairs
	set removedPaths to remainingPaths(removedPaths, parentPairs, 1)
	set addedPaths to remainingPaths(addedPaths, parentPairs, 2)

	-- 親の改名・移動に伴うだけの移動は伏せる
	set pairMap to current application's NSMutableDictionary's dictionary()
	repeat with pairRef in (movedPairs & renamedPairs)
		(pairMap's setObject:(item 2 of pairRef) forKey:(item 1 of pairRef))
	end repeat
	set shownMoves to {}
	set impliedCount to 0
	repeat with pairRef in movedPairs
		set mappedParent to (pairMap's objectForKey:parentOf(item 1 of pairRef))
		if mappedParent is not missing value and (mappedParent as text) is parentOf(item 2 of pairRef) then
			set impliedCount to impliedCount + 1
		else
			set end of shownMoves to contents of pairRef
		end if
	end repeat

	return {added:addedPaths, removed:removedPaths, renamed:renamedPairs, moved:shownMoves, impliedMoves:impliedCount, shortcutChanges:shortcutChanges, flagChanges:flagChanges}
end compareSnapshots

-- 移動した子の親どうしが、どちらも削除・追加の残りにあれば {旧, 新} の組にする（先に見つけた組を優先）
on pairByChildren(oldPaths, newPaths, movedPairs)
	set oldSet to current application's NSSet's setWithArray:oldPaths
	set newSet to current application's NSSet's setWithArray:newPaths
	set usedOld to current application's NSMutableSet's |set|()
	set usedNew to current application's NSMutableSet's |set|()
	set pairs to {}
	repeat with pairRef in movedPairs
		set oldParent to parentOf(item 1 of pairRef)
		set newParent to parentOf(item 2 of pairRef)
		if (oldSet's containsObject:oldParent) as boolean and (newSet's containsObject:newParent) as boolean then
			if not ((usedOld's containsObject:oldParent) as boolean) and not ((usedNew's containsObject:newParent) as boolean) then
				set end of pairs to {oldParent, newParent}
				(usedOld's addObject:oldParent)
				(usedNew's addObject:newParent)
			end if
		end if
	end repeat
	return pairs
end pairByChildren

-- keyKind（"leaf" か "parent"）が同じで、両側に1つずつしかないものを {旧, 新} の組にする
on pairUp(oldPaths, newPaths, keyKind)
	set oldGroups to groupPaths(oldPaths, keyKind)
	set newGroups to groupPaths(newPaths, keyKind)
	set pairs to {}
	repeat with pathRef in oldPaths
		set itemPath to contents of pathRef
		set groupKey to pathKey(itemPath, keyKind)
		set newGroup to (newGroups's objectForKey:groupKey)
		if newGroup is not missing value then
			if ((oldGroups's objectForKey:groupKey)'s |count|()) as integer is 1 and (newGroup's |count|()) as integer is 1 then
				set end of pairs to {itemPath, (newGroup's firstObject()) as text}
			end if
		end if
	end repeat
	return pairs
end pairUp

on groupPaths(pathList, keyKind)
	set groups to current application's NSMutableDictionary's dictionary()
	repeat with pathRef in pathList
		set groupKey to pathKey(contents of pathRef, keyKind)
		set group to (groups's objectForKey:groupKey)
		if group is missing value then
			set group to current application's NSMutableArray's array()
			(groups's setObject:group forKey:groupKey)
		end if
		(group's addObject:(contents of pathRef))
	end repeat
	return groups
end groupPaths

on pathKey(itemPath, keyKind)
	if keyKind is "leaf" then return leafOf(itemPath)
	return parentOf(itemPath)
end pathKey

-- 組になったもの（組の side 番目）を除いた残り
on remainingPaths(pathList, pairs, side)
	if pairs is {} then return pathList
	set pairedSet to current application's NSMutableSet's |set|()
	repeat with pairRef in pairs
		(pairedSet's addObject:(item side of pairRef))
	end repeat
	set remaining to {}
	repeat with pathRef in pathList
		if not ((pairedSet's containsObject:(contents of pathRef)) as boolean) then set end of remaining to contents of pathRef
	end repeat
	return remaining
end remainingPaths

on leafOf(itemPath)
	set parts to splitText(itemPath, pathSeparator)
	return item -1 of parts
end leafOf

on parentOf(itemPath)
	set parts to splitText(itemPath, pathSeparator)
	if (count of parts) is 1 then return ""
	return joinText(items 1 thru -2 of parts, pathSeparator)
end parentOf

-- ===== レポート =====

on writeReport(oldFile, newFile, reportPath)
	set oldSnap to readSnapshot(oldFile)
	set newSnap to readSnapshot(newFile)
	set diff to compareSnapshots(oldSnap, newSnap)
	set oldName to metaValue(oldSnap, "name")
	set newName to metaValue(newSnap, "name")

	set html to {"<!doctype html><html lang=\"ja\"><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">", ¬
		"<title>" & escapeHTML(oldName & " → " & newName) & "</title>", reportStyle(), "</head><body>", ¬
		"<h1>Illustratorのメニュー比較</h1>", ¬
		"<p class=\"lead\"><b>" & escapeHTML(oldName) & "</b> → <b>" & escapeHTML(newName) & "</b></p>"}

	if metaValue(oldSnap, "locale") is not metaValue(newSnap, "locale") then
		set end of html to "<p class=\"warn\">言語が異なります（" & escapeHTML(metaValue(oldSnap, "locale")) & " / " & escapeHTML(metaValue(newSnap, "locale")) & "）。メニュー名が訳語ごと異なるため、ほとんどの項目が追加・削除として出ます。</p>"
	end if

	-- 概要
	set end of html to "<table class=\"meta\"><tr><th></th><th>旧</th><th>新</th></tr>"
	set end of html to metaRow("名前", oldName, newName)
	set end of html to metaRow("項目数", (count of (paths of oldSnap)) as text, (count of (paths of newSnap)) as text)
	set end of html to metaRow("書き出した日時", metaValue(oldSnap, "date"), metaValue(newSnap, "date"))
	set end of html to metaRow("アプリ", metaValue(oldSnap, "app"), metaValue(newSnap, "app"))
	set end of html to "</table>"
	set end of html to "<ul class=\"summary\">" & ¬
		summaryItem("added", "追加", count of (added of diff)) & ¬
		summaryItem("removed", "削除", count of (removed of diff)) & ¬
		summaryItem("renamed", "名前の変更", count of (renamed of diff)) & ¬
		summaryItem("moved", "場所の変更", count of (moved of diff)) & ¬
		summaryItem("shortcuts", "ショートカットの変化", count of (shortcutChanges of diff)) & ¬
		summaryItem("flags", "表示状態の変化", count of (flagChanges of diff)) & "</ul>"
	set end of html to "<input id=\"filter\" type=\"search\" placeholder=\"絞り込み（メニュー名の一部）\" autocomplete=\"off\">"

	-- 追加・削除
	set end of html to sectionHead("added", "追加された項目", count of (added of diff), "")
	set end of html to pathListHTML(added of diff, newSnap)
	set end of html to sectionHead("removed", "削除された項目", count of (removed of diff), "")
	set end of html to pathListHTML(removed of diff, oldSnap)

	-- 改名・移動
	set end of html to sectionHead("renamed", "名前が変わったらしい項目", count of (renamed of diff), "同じメニューの中で、削除と追加が1つずつあったものを組にしています。")
	set end of html to pairTableHTML(renamed of diff)
	set movedNote to "同じ名前の項目が、削除側・追加側に1つずつあったものを組にしています。"
	if (impliedMoves of diff) > 0 then set movedNote to movedNote & "親メニューの名前・場所が変わっただけのもの（" & (impliedMoves of diff) & "件）は省いています。"
	set end of html to sectionHead("moved", "場所が変わったらしい項目", count of (moved of diff), movedNote)
	set end of html to pairTableHTML(moved of diff)

	-- ショートカット・状態
	set end of html to sectionHead("shortcuts", "ショートカットの変化", count of (shortcutChanges of diff), "書き出したときに選ばれていたキーボードショートカットのセットの値です。")
	set end of html to changeTableHTML(shortcutChanges of diff)
	set end of html to sectionHead("flags", "表示状態の変化", count of (flagChanges of diff), "hidden：隠し項目、alternate：⌥などを押している間だけ現れる項目")
	set end of html to changeTableHTML(flagChanges of diff)

	set end of html to "<p class=\"foot\">AiMenuDiff " & scriptVersion & "</p>" & reportScript() & "</body></html>"
	writeTextFile(reportPath, joinText(html, linefeed))
	return diff
end writeReport

on reportStyle()
	return "<style>
:root{--bg:#fff;--fg:#1d1d1f;--muted:#6e6e73;--line:#e3e3e8;--add:#1a7f37;--del:#c4302b;--chg:#9a6700;--chip:#f2f2f5}
@media (prefers-color-scheme:dark){:root{--bg:#1c1c1e;--fg:#f2f2f7;--muted:#98989f;--line:#38383c;--add:#56d364;--del:#ff7b72;--chg:#e3b341;--chip:#2c2c2e}}
body{background:var(--bg);color:var(--fg);font:14px/1.6 -apple-system,'Hiragino Sans',sans-serif;margin:0 auto;max-width:1000px;padding:24px 16px 48px}
h1{font-size:22px;margin:0 0 4px}h2{font-size:17px;margin:32px 0 6px;padding-bottom:4px;border-bottom:1px solid var(--line)}
h2 .n{color:var(--muted);font-weight:normal;margin-left:6px}.note,.lead{color:var(--muted);margin:4px 0 8px}.lead b{color:var(--fg)}
.warn{background:#fff4ce;color:#5c4400;padding:8px 12px;border-radius:6px}
table{border-collapse:collapse;width:100%}th,td{text-align:left;padding:4px 8px;border-bottom:1px solid var(--line);vertical-align:top;word-break:break-all}
th{color:var(--muted);font-weight:normal;white-space:nowrap}table.meta td{font-size:13px}
.summary{list-style:none;padding:0;display:flex;flex-wrap:wrap;gap:8px;margin:16px 0}
.summary a{display:block;background:var(--chip);border-radius:6px;padding:4px 10px;color:var(--fg);text-decoration:none}.summary b{margin-left:6px}
ul.paths{list-style:none;padding:0;margin:0}ul.paths li{padding:2px 0;border-bottom:1px solid var(--line)}
.p{color:var(--muted)}.k{color:var(--muted);font-size:12px;margin-left:8px}
.add{color:var(--add)}.del{color:var(--del)}.chg{color:var(--chg)}
#filter{width:100%;box-sizing:border-box;font:inherit;padding:6px 10px;border:1px solid var(--line);border-radius:6px;background:var(--bg);color:var(--fg)}
.none{color:var(--muted)}.foot{color:var(--muted);font-size:12px;margin-top:40px}
</style>"
end reportStyle

on reportScript()
	return "<script>
document.getElementById('filter').addEventListener('input',function(e){var q=e.target.value.toLowerCase();
document.querySelectorAll('[data-q]').forEach(function(el){el.style.display=(!q||el.getAttribute('data-q').toLowerCase().indexOf(q)>=0)?'':'none';});});
</script>"
end reportScript

on metaRow(label, oldValue, newValue)
	return "<tr><th>" & escapeHTML(label) & "</th><td>" & escapeHTML(oldValue) & "</td><td>" & escapeHTML(newValue) & "</td></tr>"
end metaRow

on summaryItem(anchorID, label, itemCount)
	return "<li><a href=\"#" & anchorID & "\">" & label & "<b>" & itemCount & "</b></a></li>"
end summaryItem

on sectionHead(anchorID, label, itemCount, noteText)
	set headText to "<h2 id=\"" & anchorID & "\">" & label & "<span class=\"n\">" & itemCount & "</span></h2>"
	if noteText is not "" then set headText to headText & "<p class=\"note\">" & escapeHTML(noteText) & "</p>"
	return headText
end sectionHead

-- 親を薄く、末端を通常の色で（コピーすると元のパスになる）
on pathHTML(itemPath)
	set parentText to parentOf(itemPath)
	if parentText is "" then return "<span class=\"leaf\">" & escapeHTML(itemPath) & "</span>"
	return "<span class=\"p\">" & escapeHTML(parentText & pathSeparator) & "</span><span class=\"leaf\">" & escapeHTML(leafOf(itemPath)) & "</span>"
end pathHTML

-- ショートカットと状態（hidden など）を添える
on pathListHTML(pathList, snapshot)
	if pathList is {} then return "<p class=\"none\">なし</p>"
	set itemsHTML to {"<ul class=\"paths\">"}
	repeat with pathRef in pathList
		set itemPath to contents of pathRef
		set shortcutText to ((shortcuts of snapshot)'s objectForKey:itemPath) as text
		set flagText to ((flags of snapshot)'s objectForKey:itemPath) as text
		set lineHTML to "<li data-q=\"" & escapeHTML(itemPath) & "\">" & pathHTML(itemPath)
		if shortcutText is not "" then set lineHTML to lineHTML & "<span class=\"k\">" & escapeHTML(shortcutText) & "</span>"
		if flagText is not "" then set lineHTML to lineHTML & "<span class=\"k\">（" & escapeHTML(flagText) & "）</span>"
		set end of itemsHTML to lineHTML & "</li>"
	end repeat
	set end of itemsHTML to "</ul>"
	return joinText(itemsHTML, linefeed)
end pathListHTML

on pairTableHTML(pairs)
	if pairs is {} then return "<p class=\"none\">なし</p>"
	set rowsHTML to {"<table><tr><th>旧</th><th>新</th></tr>"}
	repeat with pairRef in pairs
		set end of rowsHTML to "<tr data-q=\"" & escapeHTML((item 1 of pairRef) & " " & (item 2 of pairRef)) & "\"><td class=\"del\">" & pathHTML(item 1 of pairRef) & "</td><td class=\"add\">" & pathHTML(item 2 of pairRef) & "</td></tr>"
	end repeat
	set end of rowsHTML to "</table>"
	return joinText(rowsHTML, linefeed)
end pairTableHTML

on changeTableHTML(changes)
	if changes is {} then return "<p class=\"none\">なし</p>"
	set rowsHTML to {"<table><tr><th>項目</th><th>旧</th><th>新</th></tr>"}
	repeat with changeRef in changes
		set end of rowsHTML to "<tr data-q=\"" & escapeHTML(item 1 of changeRef) & "\"><td>" & pathHTML(item 1 of changeRef) & "</td><td class=\"del\">" & dashIfEmpty(item 2 of changeRef) & "</td><td class=\"chg\">" & dashIfEmpty(item 3 of changeRef) & "</td></tr>"
	end repeat
	set end of rowsHTML to "</table>"
	return joinText(rowsHTML, linefeed)
end changeTableHTML

on dashIfEmpty(valueText)
	if valueText is "" then return "—"
	return escapeHTML(valueText)
end dashIfEmpty

-- ===== ファイルと名前 =====

on prepareSnapshotFolder()
	set primaryPath to expandTilde(snapshotFolderPath)
	set fileManager to current application's NSFileManager's defaultManager()
	set parentPath to ((current application's NSString's stringWithString:primaryPath)'s stringByDeletingLastPathComponent()) as text
	if (fileManager's fileExistsAtPath:parentPath) as boolean then
		set folderPath to primaryPath
	else
		set folderPath to expandTilde(fallbackFolderPath)
	end if
	makeFolder(folderPath)
	return folderPath
end prepareSnapshotFolder

on makeFolder(folderPath)
	set fileManager to current application's NSFileManager's defaultManager()
	set {ok, theError} to fileManager's createDirectoryAtPath:folderPath withIntermediateDirectories:true attributes:(missing value) |error|:(reference)
	if not (ok as boolean) then error "フォルダーを作れません：" & folderPath & return & (theError's localizedDescription() as text)
end makeFolder

on snapshotPathFor(snapshotFolder, snapshotName)
	return snapshotFolder & "/" & snapshotName & "." & snapshotExtension
end snapshotPathFor

-- フォルダー内のスナップショット名（拡張子なし、バージョン順）。excludedName は除く
on listSnapshotNames(snapshotFolder, excludedName)
	set fileManager to current application's NSFileManager's defaultManager()
	set fileNames to fileManager's contentsOfDirectoryAtPath:snapshotFolder |error|:(missing value)
	if fileNames is missing value then return {}
	set baseNames to current application's NSMutableArray's array()
	repeat with fileRef in (fileNames as list)
		set fileName to (current application's NSString's stringWithString:(contents of fileRef))
		if (fileName's pathExtension() as text) is snapshotExtension then
			set baseName to (fileName's stringByDeletingPathExtension()) as text
			if baseName is not excludedName then (baseNames's addObject:baseName)
		end if
	end repeat
	return (baseNames's sortedArrayUsingSelector:"localizedStandardCompare:") as list
end listSnapshotNames

-- 既定の比較相手：同じ言語で、いまより古いうち最も新しいもの。無ければ最後のもの
on pickDefaultCounterpart(otherNames, currentName)
	set localeSuffix to " " & item -1 of splitText(currentName, " ")
	set candidate to missing value
	repeat with nameRef in otherNames
		set otherName to contents of nameRef
		if otherName ends with localeSuffix and compareNames(otherName, currentName) < 0 then set candidate to otherName
	end repeat
	if candidate is missing value then return item -1 of otherNames
	return candidate
end pickDefaultCounterpart

-- 数字を数として比べる（Ai9.0 < Ai30.8.2 < Ai31.0.0）
on compareNames(firstName, secondName)
	return ((current application's NSString's stringWithString:firstName)'s localizedStandardCompare:secondName) as integer
end compareNames

-- インストール先の言語フォルダー（ja_JP など）
on detectLocale(installFolder)
	set fileManager to current application's NSFileManager's defaultManager()
	set entries to fileManager's contentsOfDirectoryAtPath:installFolder |error|:(missing value)
	if entries is not missing value then
		repeat with entryRef in (entries as list)
			if matchesPattern(contents of entryRef, "^[a-z]{2}_[A-Z]{2}$") then return contents of entryRef
		end repeat
	end if
	return ((current application's NSLocale's preferredLanguages()'s firstObject()) as text)
end detectLocale

on isRunningInIllustrator()
	set bundleID to current application's NSBundle's mainBundle()'s bundleIdentifier()
	if bundleID is missing value then return false
	return (bundleID as text) starts with "com.adobe.illustrator"
end isRunningInIllustrator

on metaValue(snapshot, keyName)
	set value to ((meta of snapshot)'s objectForKey:keyName)
	if value is missing value then return ""
	return value as text
end metaValue

on timestampText()
	set formatter to current application's NSDateFormatter's new()
	formatter's setLocale:(current application's NSLocale's localeWithLocaleIdentifier:"en_US_POSIX")
	formatter's setDateFormat:"yyyy-MM-dd HH:mm"
	return (formatter's stringFromDate:(current application's NSDate's |date|())) as text
end timestampText

-- ===== 文字列 =====

on writeTextFile(filePath, textValue)
	set theString to current application's NSString's stringWithString:textValue
	set {ok, theError} to theString's writeToFile:filePath atomically:true encoding:(current application's NSUTF8StringEncoding) |error|:(reference)
	if not (ok as boolean) then error "書き込めません：" & filePath & return & (theError's localizedDescription() as text)
end writeTextFile

on expandTilde(pathText)
	return ((current application's NSString's stringWithString:pathText)'s stringByExpandingTildeInPath()) as text
end expandTilde

on matchesPattern(textValue, pattern)
	set theRange to (current application's NSString's stringWithString:textValue)'s rangeOfString:pattern options:(current application's NSRegularExpressionSearch)
	return (|length| of theRange) > 0
end matchesPattern

on replacePattern(theString, pattern, replacement)
	return theString's stringByReplacingOccurrencesOfString:pattern withString:replacement options:(current application's NSRegularExpressionSearch) range:{0, theString's |length|()}
end replacePattern

on escapeHTML(textValue)
	set textValue to replaceText(textValue as text, "&", "&amp;")
	set textValue to replaceText(textValue, "<", "&lt;")
	set textValue to replaceText(textValue, ">", "&gt;")
	return replaceText(textValue, "\"", "&quot;")
end escapeHTML

on replaceText(textValue, findText, replacement)
	if textValue does not contain findText then return textValue
	return joinText(splitText(textValue, findText), replacement)
end replaceText

on splitText(textValue, delimiter)
	set savedDelimiters to AppleScript's text item delimiters
	set AppleScript's text item delimiters to delimiter
	set parts to text items of textValue
	set AppleScript's text item delimiters to savedDelimiters
	return parts
end splitText

on joinText(textList, delimiter)
	set savedDelimiters to AppleScript's text item delimiters
	set AppleScript's text item delimiters to delimiter
	set joined to textList as text
	set AppleScript's text item delimiters to savedDelimiters
	return joined
end joinText
