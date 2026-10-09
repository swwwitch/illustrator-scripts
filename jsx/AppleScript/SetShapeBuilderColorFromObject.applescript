(*
	SetShapeBuilderColorFromObject
	［シェイプ形成ツールオプション］の「次のカラーを利用」を「オブジェクト」にする。

	・設定キー Planar/MergeTool/PaintFills（0：オブジェクト、1：スウォッチ）を読み、
	  すでに「オブジェクト」ならダイアログボックスを開かずに終わる
	・キーは setIntegerPreference で書いてもツールに反映されず［OK］で上書きされるので、GUI で切り替える
	・ツールを選んで return でオプションを開く（ドキュメントが開いていないと開かない）。
	  ダイアログボックスは AXLayoutArea として現れ、Illustrator が背面に回ると閉じる
	・ポップアップは値を書き込めず、開いた一覧の項目にもアクションが無いので、↑ と return で選ぶ
	・実行後は選択中のツールが［シェイプ形成ツール］になる

	動作確認：Illustrator 2026（30.8）日本語版 / macOS
	必要な許可：実行するアプリ（スクリプトエディタなど）にアクセシビリティの許可
*)

set dlgName to "シェイプ形成ツールオプション"
set targetName to "オブジェクト"

tell application "Adobe Illustrator"
	-- ドキュメントが無いとツールオプションは開かない
	if (count of documents) = 0 then
		display alert "ドキュメントが開かれていません。"
		return
	end if
	-- すでに「オブジェクト」ならダイアログボックスを開かずに終わる
	set paintFills to do javascript "app.preferences.getIntegerPreference('Planar/MergeTool/PaintFills');"
	if paintFills as integer is 0 then return
	activate
	do javascript "app.selectTool('Adobe Shape Builder Tool');"
end tell

tell application "System Events"
	tell process "Adobe Illustrator"
		set frontmost to true
		delay 0.3
		-- ツールを選んだ状態で return を押すとオプションが開く（1回目は空振りすることがある）
		repeat 3 times
			if exists UI element dlgName then exit repeat
			key code 36
			repeat 10 times
				delay 0.2
				if exists UI element dlgName then exit repeat
			end repeat
		end repeat
		if not (exists UI element dlgName) then
			display dialog "［シェイプ形成ツールオプション］を開けませんでした。" buttons {"OK"} default button 1
			return
		end if

		set w to UI element dlgName
		-- 2つ目のコンボボックスが「次のカラーを利用」（項目は オブジェクト／スウォッチ の順）
		set cb to combo box 2 of w
		if value of cb is not targetName then
			perform action "AXPress" of cb
			delay 0.5
			key code 126 -- ↑ で先頭の「オブジェクト」へ
			delay 0.2
			key code 36
			delay 0.5
		end if
		click (first button of w whose description is "OK")
	end tell
end tell
