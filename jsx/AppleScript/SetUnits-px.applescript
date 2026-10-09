(*
	SetUnits-px
	［環境設定］→［単位］をすべて px にそろえ、
	キー入力 1px・角丸の半径 1px・サイズ/行送り 1px・ベースラインシフト 0.1px にする。

	・数値の設定は JavaScript で書き込む（Illustrator では 1px＝1pt）
	・単位は JavaScript で書いても再起動まで効かないので、目的の隣の値を書いてから
	  ［単位］の環境設定を開き、矢印キーで目的の単位に動かして［OK］で確定する

	対象：Illustrator 2020（24.2）以降の日本語版 / macOS（この版は実機で未確認）
	必要な許可：実行するアプリ（スクリプトエディタなど）にアクセシビリティの許可
*)

tell application id "com.adobe.illustrator"
	activate
	do javascript "
var p = app.preferences;
// 数値の設定（pt で保存、1px = 1pt）
p.setRealPreference('cursorKeyLength', 1);      // キー入力 1px
p.setRealPreference('ovalRadius', 1);           // 角丸の半径 1px
p.setRealPreference('text/sizeIncrement', 1);   // サイズ/行送り 1px
p.setRealPreference('text/riseIncrement', 0.1); // ベースラインシフト 0.1px
// 単位は目的の隣の値にしておき、あとで矢印キーで目的の単位に動かして確定する
p.setIntegerPreference('rulerType', 2);       // pt → ↑ で px
p.setIntegerPreference('strokeUnits', 2);     // pt → ↑ で px
p.setIntegerPreference('text/units', 5);      // Q/H → ↓ で px
p.setIntegerPreference('text/asianunits', 5); // Q/H → ↓ で px
"
	ignoring application responses
		execute menu command menu command string "unitundoPref"
	end ignoring
end tell

tell application "System Events"
	delay 0.3
	key code 126 -- 一般: px
	keystroke tab
	key code 126 -- 線: px
	keystroke tab
	key code 125 -- 文字: px
	keystroke tab
	key code 125 -- 東アジア言語のオプション: px
	keystroke return
end tell
