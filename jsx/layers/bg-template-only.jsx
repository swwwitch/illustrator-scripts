#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);


function act_setLayTmplAttr() {
    var actionSetName = "layer";
    var actionName = "template";

    // アクション定義テキスト / Action definition text
    var actionCode ='''
 /version 3
/name [ 5
	6c61796572
]
/isOpen 1
/actionCount 1
/action-1 {
	/name [ 8
		74656d706c617465
	]
	/keyIndex 0
	/colorIndex 0
	/isOpen 1
	/eventCount 1
	/event-1 {
		/useRulersIn1stQuadrant 0
		/internalName (ai_plugin_Layer)
		/localizedName [ 9
			e8a1a8e7a4ba203a20
		]
		/isOpen 1
		/isOn 1
		/hasDialog 1
		/showDialog 0
		/parameterCount 9
		/parameter-1 {
			/key 1836411236
			/showInPalette 4294967295
			/type (integer)
			/value 4
		}
		/parameter-2 {
			/key 1851878757
			/showInPalette 4294967295
			/type (ustring)
			/value [ 36
				e383ace382a4e383a4e383bce38391e3838de383abe382aae38397e382b7e383
				a7e383b3
			]
		}
		/parameter-3 {
			/key 1953329260
			/showInPalette 4294967295
			/type (boolean)
			/value 1
		}
		/parameter-4 {
			/key 1936224119
			/showInPalette 4294967295
			/type (boolean)
			/value 1
		}
		/parameter-5 {
			/key 1819239275
			/showInPalette 4294967295
			/type (boolean)
			/value 1
		}
		/parameter-6 {
			/key 1886549623
			/showInPalette 4294967295
			/type (boolean)
			/value 1
		}
		/parameter-7 {
			/key 1886547572
			/showInPalette 4294967295
			/type (boolean)
			/value 0
		}
		/parameter-8 {
			/key 1684630830
			/showInPalette 4294967295
			/type (boolean)
			/value 1
		}
		/parameter-9 {
			/key 1885564532
			/showInPalette 4294967295
			/type (unit real)
			/value 50.0
			/unit 592474723
		}
	}
}
''';
    try {
        var tempFile = new File(Folder.temp + "/temp_action.aia");
        tempFile.open("w");
        tempFile.write(actionCode);
        tempFile.close();

        // アクション読み込み / Load action
        app.loadAction(tempFile);
        // アクション実行 / Execute action
        app.doScript(actionName, actionSetName);
        // アクション削除 / Unload action
        app.unloadAction(actionSetName, "");
        // 一時ファイル削除 / Remove temporary file
        tempFile.remove();
    } catch (e) {
        alert("エラーが発生しました: " + e);
    }
}

act_setLayTmplAttr();