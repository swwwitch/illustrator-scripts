-- OpenInFileViewer.app
-- Illustrator / InDesign のスクリプトから受け取ったフォルダーを開く。ファイルなら選択して表示する。
-- Path Finder が起動していれば Path Finder で、そうでなければ Finder で開く。
-- Opens a folder (or reveals a file) passed from an Illustrator / InDesign script:
-- in Path Finder when it is running, otherwise in Finder.
--
-- 受け渡し / Handoff:
--   スクリプトが /tmp/open_in_file_viewer_path.txt にフォルダーかファイルの絶対パスを書いてから、このアプリを起動する。
--   The script writes the absolute path of a folder or file to the temp file, then launches this app.
--
-- 作り方 / Build:
--   osacompile -o /Applications/OpenInFileViewer.app helpers/OpenInFileViewer.applescript
--
-- フォルダーはシェルだけで開くため、オートメーションの許可は要らない。
-- Path Finder でファイルを選択表示するときだけ Apple Events を使う（初回に許可を求められる）。
-- osascript で実行時にコンパイルするので、Path Finder が無い環境でも起動時に問われない。
-- Folders open with shell commands only, so no Automation permission is needed.
-- Revealing a file in Path Finder uses Apple Events (permission is asked once); osascript compiles it
-- at run time, so Macs without Path Finder are never asked to locate it.

on run
	set pathFile to "/tmp/open_in_file_viewer_path.txt"

	-- 読んだら消す（前回のパスを黙って開かないように）/ Remove after reading so a stale path is never reused
	try
		set folderPath to do shell script "/bin/cat " & quoted form of pathFile & "; /bin/rm -f " & quoted form of pathFile
	on error
		return
	end try
	if folderPath is "" then return

	set isFolder to true
	try
		do shell script "/bin/test -d " & quoted form of folderPath
	on error
		set isFolder to false
		try
			do shell script "/bin/test -e " & quoted form of folderPath
		on error
			return
		end try
	end try

	set pathFinderRunning to true
	try
		do shell script "/usr/bin/pgrep -x 'Path Finder'"
	on error
		set pathFinderRunning to false
	end try

	if pathFinderRunning then
		try
			if isFolder then
				do shell script "/usr/bin/open -a 'Path Finder' " & quoted form of folderPath
			else
				do shell script "/usr/bin/osascript -e 'on run argv' -e 'tell application \"Path Finder\"' -e 'activate' -e 'reveal (item 1 of argv)' -e 'end tell' -e 'end run' " & quoted form of folderPath
			end if
			return
		end try
	end if
	if isFolder then
		do shell script "/usr/bin/open -a Finder " & quoted form of folderPath
	else
		do shell script "/usr/bin/open -R " & quoted form of folderPath
	end if
end run
