#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

現在のアートボードと同じ大きさの長方形を作成し、「bg-template」レイヤーに置いてテンプレート化したうえで最背面へ移動します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/bg-template.md

### Overview

Creates a rectangle the size of the current artboard, places it on a "bg-template" layer, marks that layer as a template and sends it to the back.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/bg-template.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "bg-template";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-29";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/bg-template.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/bg-template.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var BG_LAYER_NAME = "bg-template";         /* 背景を置くレイヤー名 / name of the background layer */
    var DEFAULT_CMYK = ["0", "0", "0", "45"];  /* CMYK の初期値 / default CMYK values */
    var DEFAULT_RGB = ["153", "153", "153"];   /* RGB の初期値 / default RGB values */
    var DEFAULT_HEX = "999999";                /* HEX の初期値 / default hex value */
    var DEFAULT_MARGIN = "0";                  /* マージンの初期値 / default margin */
    var DEFAULT_AS_TEMPLATE = true;            /* テンプレートレイヤーにするかの初期値 / template layer on by default */

    // =========================================
    // レイアウト / Layout
    // =========================================
    var PANEL_MARGINS = [15, 20, 15, 10];      /* パネル余白 [左,上,右,下] / panel margins */
    var COLOR_COLUMN_SPACING = 20;             /* CMYK 列と RGB 列の間隔 / gap between the CMYK and RGB columns */
    var CHANNEL_LABEL_WIDTH = 20;              /* チャンネル名の幅 / width of the channel labels */
    var CHANNEL_INPUT_WIDTH = 40;              /* チャンネル値の入力欄の幅 / width of the channel fields */

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コードに対応する表示ラベルと、1単位あたりのポイント数
       Unit code -> display label and points per unit */
    var UNITS = [
        { label: "in",    pointsPerUnit: 72 },                /* 0 */
        { label: "mm",    pointsPerUnit: 72 / 25.4 },         /* 1 */
        { label: "pt",    pointsPerUnit: 1 },                 /* 2 */
        { label: "pica",  pointsPerUnit: 12 },                /* 3 */
        { label: "cm",    pointsPerUnit: 72 / 2.54 },         /* 4 */
        { label: "Q",     pointsPerUnit: 72 / 25.4 * 0.25 },  /* 5 */
        { label: "px",    pointsPerUnit: 1 },                 /* 6 */
        { label: "ft/in", pointsPerUnit: 72 * 12 },           /* 7 */
        { label: "m",     pointsPerUnit: 72 / 25.4 * 1000 },  /* 8 */
        { label: "yd",    pointsPerUnit: 72 * 36 },           /* 9 */
        { label: "ft",    pointsPerUnit: 72 * 12 }            /* 10 */
    ];

    /* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
       Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        /* 級（Q）と歯（H）は同じ長さだが、文字サイズは「Q」、距離は「H」と呼び分ける */
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * UI言語を返す
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "背景色レイヤーを作成", en: "Create Background Color Layer" }
        },
        panel: {
            color: { ja: "カラー設定", en: "Color Settings" }
        },
        fieldLabel: {
            margin: { ja: "マージン", en: "Margin" }
        },
        checkbox: {
            template: { ja: "「テンプレート」レイヤーに", en: "Set as Template Layer" }
        },
        tooltip: {
            hex: {
                ja: "背景色を16進数で指定します（例: 999999）。RGB欄と連動します。",
                en: "Background color as a hex value, for example 999999. It is linked to the RGB fields."
            },
            margin: { ja: "アートボードの外側へ背景を広げる量です。", en: "How far the background extends past the artboard." },
            template: {
                ja: "作った背景をテンプレートレイヤーに置きます。印刷・書き出しの対象から外れ、ロックされます。",
                en: "Puts the background on a template layer: it is locked and left out of printing and export."
            }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            error: { ja: "エラーが発生しました: ", en: "An error occurred: " }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "panel.color" のようなパス
     * @returns {string} 表示言語のテキスト
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        return LABELS[labelPathKeys[0]][labelPathKeys[1]][uiLang];
    }

    /**
     * 文字列に言語別のコロンを付ける（日本語は全角、英語は半角）
     * @param {string} text - 項目名
     * @returns {string} コロン付きの項目名
     */
    function appendColon(text) {
        return text + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return appendColon(getLabel(labelPath));
    }

    // =========================================
    // 一時アクション / Temporary action
    // =========================================

    /**
     * アクティブレイヤーを一時アクションでテンプレートレイヤーにする（アクションはレイヤー名も書き換える）
     * @returns {void}
     */
    function applyTemplateByAction() {
        var actionSetName = "layer";
        var actionName = "template";

        /* アクション定義テキスト / Action definition text */
        var actionCode = [
            " /version 3",
            "/name [ 5",
            "	6c61796572",
            "]",
            "/isOpen 1",
            "/actionCount 1",
            "/action-1 {",
            "	/name [ 8",
            "		74656d706c617465",
            "	]",
            "	/keyIndex 0",
            "	/colorIndex 0",
            "	/isOpen 1",
            "	/eventCount 1",
            "	/event-1 {",
            "		/useRulersIn1stQuadrant 0",
            "		/internalName (ai_plugin_Layer)",
            "		/localizedName [ 9",
            "			e8a1a8e7a4ba203a20",
            "		]",
            "		/isOpen 1",
            "		/isOn 1",
            "		/hasDialog 1",
            "		/showDialog 0",
            "		/parameterCount 9",
            "		/parameter-1 {",
            "			/key 1836411236",
            "			/showInPalette 4294967295",
            "			/type (integer)",
            "			/value 4",
            "		}",
            "		/parameter-2 {",
            "			/key 1851878757",
            "			/showInPalette 4294967295",
            "			/type (ustring)",
            "			/value [ 36",
            "				e383ace382a4e383a4e383bce38391e3838de383abe382aae38397e382b7e383",
            "				a7e383b3",
            "			]",
            "		}",
            "		/parameter-3 {",
            "			/key 1953329260",
            "			/showInPalette 4294967295",
            "			/type (boolean)",
            "			/value 1",
            "		}",
            "		/parameter-4 {",
            "			/key 1936224119",
            "			/showInPalette 4294967295",
            "			/type (boolean)",
            "			/value 1",
            "		}",
            "		/parameter-5 {",
            "			/key 1819239275",
            "			/showInPalette 4294967295",
            "			/type (boolean)",
            "			/value 1",
            "		}",
            "		/parameter-6 {",
            "			/key 1886549623",
            "			/showInPalette 4294967295",
            "			/type (boolean)",
            "			/value 1",
            "		}",
            "		/parameter-7 {",
            "			/key 1886547572",
            "			/showInPalette 4294967295",
            "			/type (boolean)",
            "			/value 0",
            "		}",
            "		/parameter-8 {",
            "			/key 1684630830",
            "			/showInPalette 4294967295",
            "			/type (boolean)",
            "			/value 1",
            "		}",
            "		/parameter-9 {",
            "			/key 1885564532",
            "			/showInPalette 4294967295",
            "			/type (unit real)",
            "			/value 50.0",
            "			/unit 592474723",
            "		}",
            "	}",
            "}"
        ].join("\n");
        var tempFile = new File(Folder.temp + "/temp_action.aia");
        tempFile.open("w");
        tempFile.write(actionCode);
        tempFile.close();

        /* 読み込み時点でパース済みなので、ここで一時ファイルを消しておく / The action is parsed on load, so remove the temp file now */
        app.loadAction(tempFile);
        tempFile.remove();

        try {
            app.doScript(actionName, actionSetName);
        } catch (e) {
            alert(getLabel("alert.error") + e);
        } finally {
            app.unloadAction(actionSetName, "");
        }
    }

    // =========================================
    // カラー / Color
    // =========================================

    /**
     * 入力欄の値を 0〜255 の整数に丸める（数値でなければ 0）
     * @param {EditText} channelInput - RGB の入力欄
     * @returns {number} 0〜255 の整数
     */
    function readByteChannel(channelInput) {
        return Math.min(255, Math.max(0, parseInt(channelInput.text) || 0));
    }

    /**
     * 16進数の文字列（先頭の # は省略可）を RGB 値に分解する
     * @param {string} hexText - "999999" や "#999999"
     * @returns {number[]|null} [r, g, b]。6桁の16進数でなければ null
     */
    function parseHexColor(hexText) {
        var cleanHex = hexText.replace(/^#/, "");
        if (!/^[0-9A-Fa-f]{6}$/.test(cleanHex)) return null;
        return [
            parseInt(cleanHex.substr(0, 2), 16),
            parseInt(cleanHex.substr(2, 2), 16),
            parseInt(cleanHex.substr(4, 2), 16)
        ];
    }

    /**
     * RGB 欄の値から HEX 欄を更新する
     * @param {Object} colorInputs - buildColorPanel() が返す入力欄
     * @returns {void}
     */
    function updateHexFromRGB(colorInputs) {
        var hex = "#";
        var channels = [colorInputs.r, colorInputs.g, colorInputs.b];
        for (var i = 0; i < channels.length; i++) {
            hex += ("0" + readByteChannel(channels[i]).toString(16)).slice(-2);
        }
        colorInputs.hex.text = hex.toUpperCase();
    }

    /**
     * HEX 欄の値から RGB 欄を更新する（6桁の16進数のときだけ）
     * @param {Object} colorInputs - buildColorPanel() が返す入力欄
     * @returns {void}
     */
    function updateRGBFromHex(colorInputs) {
        var rgbValues = parseHexColor(colorInputs.hex.text);
        if (!rgbValues) return;
        colorInputs.r.text = rgbValues[0];
        colorInputs.g.text = rgbValues[1];
        colorInputs.b.text = rgbValues[2];
    }

    /**
     * 入力欄から塗りの色を作る（CMYK ドキュメントは CMYK、それ以外は HEX を優先して RGB）
     * @param {Object} colorInputs - buildColorPanel() が返す入力欄
     * @param {boolean} isCMYK - CMYK ドキュメントなら true
     * @returns {CMYKColor|RGBColor} 塗りの色
     */
    function buildFillColor(colorInputs, isCMYK) {
        if (isCMYK) {
            var cmykColor = new CMYKColor();
            cmykColor.cyan = parseFloat(colorInputs.c.text) || 0;
            cmykColor.magenta = parseFloat(colorInputs.m.text) || 0;
            cmykColor.yellow = parseFloat(colorInputs.y.text) || 0;
            cmykColor.black = parseFloat(colorInputs.k.text) || 0;
            return cmykColor;
        }
        var rgbColor = new RGBColor();
        var rgbValues = parseHexColor(colorInputs.hex.text);
        if (!rgbValues) {
            /* HEX が不正なら RGB 欄の値を使う / Fall back to the RGB fields when the hex is invalid */
            rgbValues = [
                parseFloat(colorInputs.r.text) || 0,
                parseFloat(colorInputs.g.text) || 0,
                parseFloat(colorInputs.b.text) || 0
            ];
        }
        rgbColor.red = rgbValues[0];
        rgbColor.green = rgbValues[1];
        rgbColor.blue = rgbValues[2];
        return rgbColor;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ↑↓キーで数値欄の値を増減する（shift で10刻み、option で0.1刻み）
     * @param {EditText} editText - 対象の数値欄
     * @param {Function} [onUpdate] - 値を変えたあとに呼ぶ処理
     * @returns {void}
     */
    function changeValueByArrowKey(editText, onUpdate) {
        editText.addEventListener("keydown", function (event) {
            if (event.keyName != "Up" && event.keyName != "Down") return;

            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var isUp = (event.keyName === "Up");
            if (keyboard.shiftKey) {
                /* 10刻み（10の倍数にスナップ）/ Step by 10, snapping to multiples of 10 */
                value = isUp ? Math.ceil((value + 1) / 10) * 10 : Math.floor((value - 1) / 10) * 10;
            } else if (keyboard.altKey) {
                /* 0.1刻み（小数第1位まで）/ Step by 0.1, one decimal place */
                value = Math.round((value + (isUp ? 0.1 : -0.1)) * 10) / 10;
            } else {
                value = Math.round(value + (isUp ? 1 : -1));
            }

            event.preventDefault();
            editText.text = value;
            if (onUpdate) onUpdate();
        });
    }

    /**
     * チャンネル名と数値欄の組を1行追加する
     * @param {Group} parent - 追加先の列
     * @param {string} channelName - "C" / "R" などのチャンネル名
     * @param {string} defaultValue - 初期値
     * @param {Function} [onUpdate] - 値が変わったときに呼ぶ処理
     * @returns {EditText} 追加した数値欄
     */
    function addChannelInput(parent, channelName, defaultValue, onUpdate) {
        var channelRow = parent.add("group");
        var channelLabel = channelRow.add("statictext", undefined, appendColon(channelName));
        channelLabel.preferredSize.width = CHANNEL_LABEL_WIDTH;
        var channelInput = channelRow.add("edittext", undefined, defaultValue);
        channelInput.characters = 4;
        channelInput.preferredSize.width = CHANNEL_INPUT_WIDTH;
        changeValueByArrowKey(channelInput, onUpdate);
        if (onUpdate) channelInput.onChanging = onUpdate;
        return channelInput;
    }

    /**
     * 入力欄をまとめて有効／無効にする
     * @param {EditText[]} inputs - 対象の入力欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setInputsEnabled(inputs, isEnabled) {
        for (var i = 0; i < inputs.length; i++) {
            inputs[i].enabled = isEnabled;
        }
    }

    /**
     * カラー設定パネル（左に CMYK、右に RGB と HEX）を作る
     * @param {Window} dlg - ダイアログ
     * @param {boolean} isCMYK - CMYK ドキュメントなら true（反対側の欄はディム表示）
     * @returns {Object} 入力欄 { c, m, y, k, r, g, b, hex }
     */
    function buildColorPanel(dlg, isCMYK) {
        var colorPanel = dlg.add("panel", undefined, getLabel("panel.color"));
        colorPanel.orientation = "row";
        colorPanel.alignChildren = ["fill", "top"];
        colorPanel.margins = PANEL_MARGINS;
        colorPanel.spacing = COLOR_COLUMN_SPACING;

        var cmykColumn = colorPanel.add("group");
        cmykColumn.orientation = "column";
        cmykColumn.alignChildren = ["left", "top"];

        var rgbColumn = colorPanel.add("group");
        rgbColumn.orientation = "column";
        rgbColumn.alignChildren = ["left", "top"];

        var colorInputs = {};
        var syncHex = function () { updateHexFromRGB(colorInputs); };

        colorInputs.c = addChannelInput(cmykColumn, "C", DEFAULT_CMYK[0]);
        colorInputs.m = addChannelInput(cmykColumn, "M", DEFAULT_CMYK[1]);
        colorInputs.y = addChannelInput(cmykColumn, "Y", DEFAULT_CMYK[2]);
        colorInputs.k = addChannelInput(cmykColumn, "K", DEFAULT_CMYK[3]);

        colorInputs.r = addChannelInput(rgbColumn, "R", DEFAULT_RGB[0], syncHex);
        colorInputs.g = addChannelInput(rgbColumn, "G", DEFAULT_RGB[1], syncHex);
        colorInputs.b = addChannelInput(rgbColumn, "B", DEFAULT_RGB[2], syncHex);

        /* RGB の下に HEX 欄 / Hex field under RGB */
        var hexRow = rgbColumn.add("group");
        hexRow.orientation = "row";
        var hexLabel = hexRow.add("statictext", undefined, "#");
        hexLabel.preferredSize.width = CHANNEL_LABEL_WIDTH;
        colorInputs.hex = hexRow.add("edittext", undefined, DEFAULT_HEX);
        colorInputs.hex.helpTip = getLabel("tooltip.hex");
        colorInputs.hex.characters = 7;
        colorInputs.hex.onChanging = function () { updateRGBFromHex(colorInputs); };

        /* カラーモードに合わない側の欄をディム表示 / Dim the fields that do not match the color mode */
        setInputsEnabled([colorInputs.c, colorInputs.m, colorInputs.y, colorInputs.k], isCMYK);
        setInputsEnabled([colorInputs.r, colorInputs.g, colorInputs.b, colorInputs.hex], !isCMYK);

        return colorInputs;
    }

    /**
     * 設定ダイアログを表示する
     * @param {boolean} isCMYK - CMYK ドキュメントなら true
     * @returns {Object|null} { colorInputs, margin, asTemplate }。キャンセル時は null
     */
    function showBackgroundDialog(isCMYK) {
        var dlg = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        dlg.orientation = "column";
        dlg.alignChildren = "left";

        var colorInputs = buildColorPanel(dlg, isCMYK);

        var marginRow = dlg.add("group");
        marginRow.orientation = "row";
        marginRow.add("statictext", undefined, labelText("fieldLabel.margin"));
        var marginInput = marginRow.add("edittext", undefined, DEFAULT_MARGIN);
        marginInput.helpTip = getLabel("tooltip.margin");
        marginInput.characters = 4;
        marginRow.add("statictext", undefined, getUnitInfo().label);
        changeValueByArrowKey(marginInput);

        var templateCheckbox = dlg.add("checkbox", undefined, getLabel("checkbox.template"));
        templateCheckbox.helpTip = getLabel("tooltip.template");
        templateCheckbox.value = DEFAULT_AS_TEMPLATE;

        var btnRowGroup = dlg.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = "center";
        btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        if (dlg.show() != 1) return null;

        return {
            colorInputs: colorInputs,
            margin: parseFloat(marginInput.text) || 0,
            asTemplate: templateCheckbox.value
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 背景レイヤーを作り直す（既存の同名レイヤーはロックを外して削除）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer} 新しい背景レイヤー
     */
    function recreateBackgroundLayer(doc) {
        var existingLayer = null;
        try {
            /* 見つからないと例外になる / getByName throws when the layer does not exist */
            existingLayer = doc.layers.getByName(BG_LAYER_NAME);
        } catch (e) {}

        if (existingLayer) {
            existingLayer.locked = false;
            existingLayer.remove();
        }

        var bgLayer = doc.layers.add();
        bgLayer.name = BG_LAYER_NAME;
        return bgLayer;
    }

    /**
     * メイン処理
     * @returns {void}
     */
    function main() {
        try {
            if (app.documents.length === 0) {
                alert(getLabel("alert.noDocument"));
                return;
            }

            var doc = app.activeDocument;
            var isCMYK = (doc.documentColorSpace === DocumentColorSpace.CMYK);

            var dialogResult = showBackgroundDialog(isCMYK);
            if (!dialogResult) return;

            var margin = dialogResult.margin;
            var artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect; /* [left, top, right, bottom] */

            var bgLayer = recreateBackgroundLayer(doc);

            /* マージンの分だけ外側へ広げた長方形 / Rectangle grown outward by the margin */
            var bgRect = bgLayer.pathItems.rectangle(
                artboardRect[1] + margin,
                artboardRect[0] - margin,
                (artboardRect[2] - artboardRect[0]) + (margin * 2),
                (artboardRect[1] - artboardRect[3]) + (margin * 2)
            );
            bgRect.fillColor = buildFillColor(dialogResult.colorInputs, isCMYK);
            bgRect.filled = true;
            bgRect.stroked = false;

            if (dialogResult.asTemplate) {
                applyTemplateByAction();
            }

            /* 最背面へ移動し、アクションで変わった名前を戻してロック / Send to back, restore the name the action changed, and lock */
            bgLayer.locked = false;
            bgLayer.zOrder(ZOrderMethod.SENDTOBACK);
            bgLayer.name = BG_LAYER_NAME;
            bgLayer.locked = true;

        } catch (e) {
            alert(getLabel("alert.error") + e);
        }
    }

    main();

})();
