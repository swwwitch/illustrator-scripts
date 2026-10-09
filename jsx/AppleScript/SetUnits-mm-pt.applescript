(*
	SetUnits-mm-pt
	［環境設定］→［単位］を「一般 mm／線 pt／文字 pt／東アジア言語のオプション pt」にそろえ、
	キー入力 0.1mm・角丸の半径 1mm・サイズ/行送り 1pt・ベースラインシフト 0.1pt にする。

	・数値の設定は JavaScript で pt に換算して書き込む
	・単位は JavaScript で書いても再起動まで効かないので、目的の隣の値を書いてから
	  ［単位］の環境設定を開き、矢印キーで目的の単位に動かして［OK］で確定する
	・SetUnits-mm-QH / SetUnits-px は同じ方法で単位の組み合わせを変えたもの

	対象：Illustrator 2020（24.2）以降の日本語版 / macOS（この版は実機で未確認）
	必要な許可：実行するアプリ（スクリプトエディタなど）にアクセシビリティの許可
*)

tell application id "com.adobe.illustrator"
	activate
	do javascript "
var p = app.preferences;
// 数値の設定（pt で保存）
p.setRealPreference('cursorKeyLength', 0.1 * 72 / 25.4); // キー入力 0.1mm
p.setRealPreference('ovalRadius', 1 * 72 / 25.4);        // 角丸の半径 1mm
p.setRealPreference('text/sizeIncrement', 1);            // サイズ/行送り 1pt
p.setRealPreference('text/riseIncrement', 0.1);          // ベースラインシフト 0.1pt
// 単位は目的の隣の値にしておき、あとで矢印キーで目的の単位に動かして確定する
p.setIntegerPreference('rulerType', 4);       // cm → ↑ で mm
p.setIntegerPreference('strokeUnits', 3);     // pc → ↑ で pt
p.setIntegerPreference('text/units', 0);      // in → ↑↑ で pt
p.setIntegerPreference('text/asianunits', 0); // in → ↑↑ で pt
"
	ignoring application responses
		execute menu command menu command string "unitundoPref"
	end ignoring
end tell

tell application "System Events"
	delay 0.3
	key code 126 -- 一般: mm
	keystroke tab
	key code 126 -- 線: pt
	keystroke tab
	key code 126
	key code 126 -- 文字: pt
	keystroke tab
	key code 126
	key code 126 -- 東アジア言語のオプション: pt
	keystroke return
end tell
