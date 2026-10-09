(*
	SetUnits-mm-QH
	［環境設定］→［単位］を「一般 mm／線 mm／文字 Q／東アジア言語のオプション H」にそろえ、
	キー入力 0.1mm・角丸の半径 1mm・サイズ/行送り 1Q・ベースラインシフト 0.5H にする。

	・数値の設定は JavaScript で pt に換算して書き込む（1Q＝1H＝0.25mm）
	・単位は JavaScript で書いても再起動まで効かないので、目的の隣の値を書いてから
	  ［単位］の環境設定を開き、矢印キーで目的の単位に動かして［OK］で確定する

	対象：Illustrator 2020（24.2）以降の日本語版 / macOS（この版は実機で未確認）
	必要な許可：実行するアプリ（スクリプトエディタなど）にアクセシビリティの許可
*)

tell application id "com.adobe.illustrator"
	activate
	do javascript "
var p = app.preferences;
var mm = 72 / 25.4;
// 数値の設定（pt で保存）
p.setRealPreference('cursorKeyLength', 0.1 * mm);       // キー入力 0.1mm
p.setRealPreference('ovalRadius', 1 * mm);              // 角丸の半径 1mm
p.setRealPreference('text/sizeIncrement', 0.25 * mm);   // サイズ/行送り 1Q
p.setRealPreference('text/riseIncrement', 0.125 * mm);  // ベースラインシフト 0.5H
// 単位は目的の隣の値にしておき、あとで矢印キーで目的の単位に動かして確定する
p.setIntegerPreference('rulerType', 4);       // cm → ↑ で mm
p.setIntegerPreference('strokeUnits', 0);     // in → ↓ で mm
p.setIntegerPreference('text/units', 6);      // px → ↑ で Q/H
p.setIntegerPreference('text/asianunits', 6); // px → ↑ で Q/H
"
	ignoring application responses
		execute menu command menu command string "unitundoPref"
	end ignoring
end tell

tell application "System Events"
	delay 0.3
	key code 126 -- 一般: mm
	keystroke tab
	key code 125 -- 線: mm
	keystroke tab
	key code 126 -- 文字: Q/H
	keystroke tab
	key code 126 -- 東アジア言語のオプション: Q/H
	keystroke return
end tell
