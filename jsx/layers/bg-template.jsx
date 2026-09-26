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
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
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
            },
            stepUp: {
                ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepDown: {
                ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepUpInteger: { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
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

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ステップボタン（再利用パーツ） / Stepper buttons (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ローカライズより前）に貼る。
    //    識別子はすべて STEPPER_* / *Stepper* / *Stepped* の名前なので、既存の名前とはぶつからない
    // 2. コピー先の LABELS.tooltip に stepUp / stepDown / stepUpInteger / stepDownInteger を足す（このファイルの LABELS から写す）。
    //    getLabel() と uiLang はコピー先のものをそのまま使う
    // 3. 数値欄を addSteppedField() で作る。項目名・∧∨・入力欄がひと組で入り、↑↓キーも∧∨と同じ処理で増減する
    //      var widthInput = addSteppedField(parentPanel, {
    //          label: labelText(LABELS.fieldLabel.width), labelWidth: 60,
    //          text: "210 mm", characters: 8, step: 1, min: 1, unit: " mm",
    //          onStep: function (numberInput) { updatePreview(); }
    //      });
    //    値の種類は options で切り分ける:
    //      小数あり（幅・位置など）   … 指定なし（option＋クリックで0.1ずつ）
    //      整数・1以上（段数・個数など）… integer: true, min: 1（0・小数・負数は受け付けず、option＋クリックも1ずつ）
    //      整数・0以上（間隔の数など）  … integer: true, min: 0
    //      範囲つき（％など）           … min: 0, max: 100, unit: "%"
    // 4. 有効／無効は setSteppedFieldEnabled(widthInput, isEnabled)（∧∨のディム表示も切り替わる）。
    //    行・パネルなど親の enabled を切り替えたときは、そのあとで redrawSteppersIn(親) を呼んで∧∨を描き直す
    //    （∧∨は親をたどって無効を判定し、無効の間はクリックも↑↓キーも効かない）
    // 5. 値は parseFloat(widthInput.text) で読む（unit 付きの欄は「210 mm」の形で入っている）
    // 6. この欄に別の↑↓キー処理を付けない（↑↓キーが二重に効く）
    // 既存の edittext をそのまま使うときは、同じ行の group（spacing 0）に addStepper() → edittext の順で置き、
    // bindSteppedArrowKeys(edittext, stepperGroup) を呼ぶ
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    // -----------------------------------------
    // ステップボタンの寸法・増減量 / Stepper metrics and steps
    // -----------------------------------------
    var STEPPER_BUTTON_WIDTH   = 20;  /* ∧∨ボタンの幅 / button width */
    var STEPPER_BUTTON_HEIGHT  = 11;  /* ∧∨ボタン1つの高さ（2つ重ねた全体の高さは22） / button height (22 for the pair) */
    var STEPPER_CORNER_RADIUS  = 2;   /* 枠の角丸の半径（ScriptUIは円弧を描けないため短い線分で近似） / corner radius, approximated with segments */
    var STEPPER_FIELD_SPACING  = 3;   /* 項目名と∧∨の間隔 / spacing between the label and the stepper */
    var STEPPER_SIDE_MARGIN    = 3;   /* ∧∨の左に足す余白（右は入力欄に突き合わせる） / extra space left of the stepper */
    var STEPPER_SHIFT_MULTIPLE = 10;  /* shift＋クリックでそろえる倍数 / Shift-click snaps to multiples of this */
    var STEPPER_OPTION_STEP    = 0.1; /* option＋クリックの増減量 / Option-click step */

    // -----------------------------------------
    // ステップボタンの配色 / Stepper colors
    // -----------------------------------------
    /**
     * UIがダークテーマかどうかを判定する（Illustrator・InDesign の両方に対応）
     * @returns {boolean} ダークなら true。取得できない環境では false（明るいUI扱い）
     */
    function isDarkStepperUI() {
        try {
            if (app.preferences && app.preferences.getRealPreference) {
                return app.preferences.getRealPreference("uiBrightness") <= 0.5; /* Illustrator */
            }
            return app.generalPreferences.uiBrightnessPreference <= 0.5; /* InDesign */
        } catch (e) {
            return false;
        }
    }

    var STEPPER_UI_DARK           = isDarkStepperUI();
    /* UIの明るさは4段階あり、段階ごとに背景色が違う。どの段階でも背景に対する差で見せるよう、黒・白の半透明を重ねる。
       ダーク側は Illustrator 標準のスピナー（［グリッドに分割］）で実測、明るい側は最も明るい段階（背景 約0.94）から逆算
       UI brightness has four levels with different backgrounds, so colors are translucent overlays that follow the
       dialog background. Dark values are measured from Illustrator's own spinner; light values derived for the lightest level */
    var STEPPER_FILL_COLOR        = STEPPER_UI_DARK ? [0, 0, 0, 0.10]  : [1, 1, 1, 0.50];  /* 地 / background */
    var STEPPER_FRAME_COLOR       = STEPPER_UI_DARK ? [1, 1, 1, 0.07]  : [0, 0, 0, 0.10];  /* 枠線 / frame */
    var STEPPER_PRESSED_COLOR     = STEPPER_UI_DARK ? [1, 1, 1, 0.12]  : [0, 0, 0, 0.13];  /* 押下中 / pressed */
    var STEPPER_CHEVRON_COLOR     = STEPPER_UI_DARK ? [1, 1, 1, 1]     : [0, 0, 0, 0.70];  /* 山形の線 / chevron */
    var STEPPER_DIM_FILL_COLOR    = STEPPER_UI_DARK ? [1, 1, 1, 0.035] : [1, 1, 1, 0.30];  /* 無効時の地 / background when disabled */
    var STEPPER_DIM_FRAME_COLOR   = STEPPER_UI_DARK ? [1, 1, 1, 0.035] : [0, 0, 0, 0.05];  /* 無効時の枠線（ダークは地と同じで見せない） / frame when disabled */
    var STEPPER_DIM_CHEVRON_COLOR = STEPPER_UI_DARK ? [1, 1, 1, 0.20]  : [0, 0, 0, 0.25];  /* 無効時の山形 / chevron when disabled */

    // -----------------------------------------
    // 数値欄を作る（外から呼ぶ関数） / Public API
    // -----------------------------------------
    /**
     * 「項目名・∧∨・入力欄」をひと組にした数値欄を追加する。
     * ↑↓キーでも∧∨と同じように増減する。直接入力した値も、フォーカスが外れたときに
     * 整数化・下限・上限・単位（「20 mm」の形）へそろえ、数値でなければ直前の値に戻す
     * @param {Group|Panel} parent - 追加先
     * @param {Object} fieldOptions - label（コロン込みの項目名）/ labelWidth / text / characters /
     *     step / min / max / integer（true で整数のみ）/ unit / onStep
     * @returns {EditText} 入力欄（項目名は .fieldLabel、∧∨は .stepperGroup で参照できる）
     */
    function addSteppedField(parent, fieldOptions) {
        var fieldRowGroup = parent.add("group");
        fieldRowGroup.orientation = "row";
        fieldRowGroup.alignChildren = ["left", "center"];
        fieldRowGroup.spacing = STEPPER_FIELD_SPACING;

        var fieldLabel = fieldRowGroup.add("statictext", undefined, fieldOptions.label || "");
        if (fieldOptions.labelWidth) {
            fieldLabel.preferredSize.width = fieldOptions.labelWidth;
            fieldLabel.justify = "right";
        }

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = fieldRowGroup.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var numberInput;
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, fieldOptions);
        numberInput = stepperInputGroup.add("edittext", undefined, fieldOptions.text || "");
        numberInput.characters = fieldOptions.characters || 6;
        numberInput.fieldLabel = fieldLabel;
        numberInput.stepperGroup = stepperGroup;

        /* ↑↓キーも∧∨と同じ処理で増減する（増減量・下限・上限・単位・修飾キーをそろえる） / arrow keys share the stepper's logic */
        bindSteppedArrowKeys(numberInput, stepperGroup);

        /* 直接入力をそろえる。数値でなければ直前の値に戻す / normalize typed values; revert non-numbers */
        numberInput.lastValidText = numberInput.text;
        numberInput.onChange = function () {
            var value = parseFloat(numberInput.text);
            if (isNaN(value)) {
                numberInput.text = numberInput.lastValidText;
                return;
            }
            writeSteppedValue(numberInput, value, fieldOptions);
        };
        return numberInput;
    }

    /**
     * 数値欄の有効／無効を、項目名・∧∨ごとまとめて切り替える
     * @param {EditText} numberInput - addSteppedField() で作った入力欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setSteppedFieldEnabled(numberInput, isEnabled) {
        numberInput.enabled = isEnabled;
        numberInput.fieldLabel.enabled = isEnabled;
        numberInput.stepperGroup.enabled = isEnabled;
        /* ∧∨は自作描画なので、描き直してディム表示を切り替える / redraw the custom-drawn buttons to update the dimming */
        for (var i = 0; i < numberInput.stepperGroup.children.length; i++) {
            redrawStepperGroup(numberInput.stepperGroup.children[i]);
        }
    }

    /**
     * 入力欄の値を増減する∧∨ボタンを、隙間なく縦に積んで追加する
     * @param {Group|Panel} parent - 追加先
     * @param {Function} getNumberInput - 対象の入力欄を返す関数（入力欄を∧∨より後に作れるよう、クリック時に引く）
     * @param {Object} stepOptions - step（増減量）/ min / max / integer / unit（例 " mm"）/ onStep(numberInput)
     * @returns {Group} ∧∨をまとめた group（.stepBy(direction) で同じ増減を呼べる）
     */
    function addStepper(parent, getNumberInput, stepOptions) {
        var stepperGroup = parent.add("group");
        stepperGroup.orientation = "column";
        stepperGroup.spacing = 0; /* 2つのボタンをつなげて1つの枠に見せる / join the buttons into one frame */
        stepperGroup.margins = [STEPPER_SIDE_MARGIN, 0, 0, 0]; /* 右は入力欄に突き合わせる / butt against the field on the right */
        stepperGroup.alignment = ["left", "center"];

        /**
         * 入力欄の値を増減する（shift を押しながらなら STEPPER_SHIFT_MULTIPLE の倍数へ、option なら STEPPER_OPTION_STEP ずつ。下限・上限で止める）
         * @param {number} direction - 増やすなら 1、減らすなら -1
         * @returns {void}
         */
        function stepBy(direction) {
            var numberInput = getNumberInput();
            if (!isStepperEnabledInTree(numberInput)) return; /* 入力欄か親が無効の間は動かさない */
            var value = parseFloat(numberInput.text);
            if (isNaN(value)) value = 0;
            writeSteppedValue(numberInput, computeSteppedValue(value, direction, stepOptions), stepOptions);
            if (stepOptions.onStep) stepOptions.onStep(numberInput);
        }

        /* 整数の欄では option＋クリックの0.1刻みが効かないので、説明から外す / integer fields have no 0.1 step */
        var upTooltip = stepOptions.integer ? "tooltip.stepUpInteger" : "tooltip.stepUp";
        var downTooltip = stepOptions.integer ? "tooltip.stepDownInteger" : "tooltip.stepDown";
        makeStepperChevronButton(stepperGroup, "up", function () { stepBy(1); }).helpTip = getLabel(upTooltip);
        makeStepperChevronButton(stepperGroup, "down", function () { stepBy(-1); }).helpTip = getLabel(downTooltip);
        stepperGroup.stepBy = stepBy; /* ↑↓キーからも同じ処理で増減できるよう公開 / shared with the arrow keys */
        return stepperGroup;
    }

    /**
     * 入力欄の↑↓キーを、∧∨と同じ処理で増減させる。ほかのキーは素通し
     * @param {EditText} numberInput - 対象の入力欄
     * @param {Group} stepperGroup - addStepper() で作った∧∨
     * @returns {void}
     */
    function bindSteppedArrowKeys(numberInput, stepperGroup) {
        numberInput.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;
            stepperGroup.stepBy(event.keyName === "Up" ? 1 : -1);
            event.preventDefault(); /* カーソル移動を止める / keep the caret from moving */
        });
    }

    // -----------------------------------------
    // 値の計算 / Value helpers
    // -----------------------------------------
    /**
     * 押された修飾キーに応じて、1回分増減した値を返す
     * （shift なら STEPPER_SHIFT_MULTIPLE の倍数へ、option なら STEPPER_OPTION_STEP ずつ、それ以外は step の倍数へ（1.5→2、1.5→1）。
     * 整数の欄では option を無視して step の倍数へ）
     * @param {number} value - 元の値
     * @param {number} direction - 増やすなら 1、減らすなら -1
     * @param {Object} stepOptions - step（通常の増減量。省略時は 1）/ integer
     * @returns {number} 増減した値（下限・上限は未適用）
     */
    function computeSteppedValue(value, direction, stepOptions) {
        var keyState = ScriptUI.environment.keyboardState;
        if (keyState.shiftKey) return snapStepperToNextMultiple(value, STEPPER_SHIFT_MULTIPLE, direction);
        if (keyState.altKey && !stepOptions.integer) return value + direction * STEPPER_OPTION_STEP;
        return snapStepperToNextMultiple(value, stepOptions.step || 1, direction);
    }

    /**
     * 値を、指定した方向にある次の倍数へ移す（230→240、232→240、下げるときは 232→230、230→220）
     * @param {number} value - 元の値
     * @param {number} multiple - 倍数の単位（例 10）
     * @param {number} direction - 上げるなら 1、下げるなら -1
     * @returns {number} 移した値
     */
    function snapStepperToNextMultiple(value, multiple, direction) {
        /* 0.29 / 0.01 = 28.999… のような浮動小数の誤差で同じ値に戻らないよう、商を丸めてから切り捨て・切り上げる
           round the quotient first so float error (0.29 / 0.01 = 28.999…) does not step back to the same value */
        var quotient = Math.round(value / multiple * 1e6) / 1e6;
        if (direction > 0) return Math.round((Math.floor(quotient) + 1) * multiple * 1e6) / 1e6;
        return Math.round((Math.ceil(quotient) - 1) * multiple * 1e6) / 1e6;
    }

    /**
     * 値を下限・上限の範囲に収める
     * @param {number} value - 数値
     * @param {Object} rangeOptions - min / max（どちらも省略可）
     * @returns {number} 範囲に収めた値
     */
    function clampSteppedValue(value, rangeOptions) {
        if (rangeOptions.min !== undefined && value < rangeOptions.min) return rangeOptions.min;
        if (rangeOptions.max !== undefined && value > rangeOptions.max) return rangeOptions.max;
        return value;
    }

    /**
     * 値を整数化・下限・上限でそろえ、単位を付けて入力欄に書き込む（直前の正しい値としても控える）
     * @param {EditText} numberInput - 書き込む入力欄
     * @param {number} value - 数値
     * @param {Object} valueOptions - integer / min / max / unit（どれも省略可）
     * @returns {void}
     */
    function writeSteppedValue(numberInput, value, valueOptions) {
        numberInput.text = formatSteppedValue(value, valueOptions);
        numberInput.lastValidText = numberInput.text;
    }

    /**
     * 値を整数化・下限・上限でそろえ、丸めて単位を付けた表示用の文字列にする。
     * 整数化してから下限で止めるので、「整数・下限1」の欄に 0.4 が入っても 1 になる
     * @param {number} value - 数値
     * @param {Object} valueOptions - integer / min / max / unit（どれも省略可）
     * @returns {string} 入力欄に入れる文字列（例 "20 mm"）
     */
    function formatSteppedValue(value, valueOptions) {
        if (valueOptions.integer) value = Math.round(value);
        return formatStepperNumber(clampSteppedValue(value, valueOptions)) + (valueOptions.unit || "");
    }

    /**
     * 小数第2位で丸めた数値を文字列で返す
     * @param {number} value - 数値
     * @returns {string} 表示用の数値文字列
     */
    function formatStepperNumber(value) {
        return String(Math.round(value * 100) / 100);
    }

    // -----------------------------------------
    // ∧∨ボタンの描画 / Drawing
    // -----------------------------------------
    /**
     * 山形（∧／∨）の極小ボタンを作成する。
     * 上下2つを隙間なく積んで1つの枠に見えるよう、枠線は外側の辺だけ描き（上ボタンは上側、下ボタンは下側）、
     * 継ぎ目に線は引かない
     * @param {Group|Panel} parent - 追加先
     * @param {string} direction - "up" または "down"
     * @param {Function} onClickFn - クリック時の処理
     * @returns {Group} ボタンとして使う group
     */
    function makeStepperChevronButton(parent, direction, onClickFn) {
        var buttonWidth = STEPPER_BUTTON_WIDTH;
        var buttonHeight = STEPPER_BUTTON_HEIGHT;
        var isUp = (direction === "up");
        var chevronBox = parent.add("group");
        chevronBox.margins = 0;
        chevronBox.spacing = 0;
        chevronBox.preferredSize = [buttonWidth, buttonHeight];
        chevronBox.minimumSize = [buttonWidth, buttonHeight];
        chevronBox.maximumSize = [buttonWidth, buttonHeight];
        chevronBox.isPressed = false;
        chevronBox.isStepperButton = true; /* redrawSteppersIn() の目印 / marker for redrawSteppersIn() */

        chevronBox.onDraw = function () {
            var boxGraphics = chevronBox.graphics;
            /* 自作描画は自動でディムにならないため、無効なら薄い色で描く。親の無効化は子の enabled に出ないので親も見る
               Custom drawing is not dimmed automatically; the parent's state does not reach the child's enabled */
            var isDimmed = !isStepperEnabledInTree(chevronBox);

            /* 枠線の内側の地（押下中は押下色） / background inside the frame, pressed color while pressed */
            var fillColor = isDimmed ? STEPPER_DIM_FILL_COLOR : (chevronBox.isPressed ? STEPPER_PRESSED_COLOR : STEPPER_FILL_COLOR);
            boxGraphics.newPath();
            boxGraphics.rectPath(1, isUp ? 1 : 0, buttonWidth - 2, buttonHeight - 1);
            boxGraphics.fillPath(boxGraphics.newBrush(boxGraphics.BrushType.SOLID_COLOR, fillColor));

            drawStepperFrame(boxGraphics, buttonWidth, buttonHeight, isUp, isDimmed ? STEPPER_DIM_FRAME_COLOR : STEPPER_FRAME_COLOR);
            drawStepperChevron(boxGraphics, buttonWidth, buttonHeight, isUp, isDimmed ? STEPPER_DIM_CHEVRON_COLOR : STEPPER_CHEVRON_COLOR);
        };

        /**
         * 押下状態を変えて描き直す
         * @param {boolean} isPressed - 押下中なら true
         * @returns {void}
         */
        function repaint(isPressed) {
            if (chevronBox.isPressed === isPressed) return;
            chevronBox.isPressed = isPressed;
            redrawStepperGroup(chevronBox);
        }
        chevronBox.addEventListener("mousedown", function () {
            if (!isStepperEnabledInTree(chevronBox)) return;
            repaint(true);
            if (onClickFn) onClickFn();
        });
        chevronBox.addEventListener("mouseup", function () { repaint(false); });
        /* 押したまま外へ出たときも押下色を残さない / reset when the pointer leaves while pressed */
        chevronBox.addEventListener("mouseout", function () { repaint(false); });
        return chevronBox;
    }

    /**
     * 外側の辺だけの枠を描く（角は丸める）。継ぎ目側は開けておき、上下2つで1つの枠に見せる。
     * ScriptUI は円弧を描けないため、角丸は短い線分で近似する
     * @param {ScriptUIGraphics} boxGraphics - 描画先
     * @param {number} boxWidth - ボタンの幅
     * @param {number} boxHeight - ボタンの高さ
     * @param {boolean} isUp - 上のボタンなら true（上側に枠を描く）
     * @param {number[]} frameColor - [r, g, b, a]
     * @returns {void}
     */
    function drawStepperFrame(boxGraphics, boxWidth, boxHeight, isUp, frameColor) {
        var frameLeft = 0.5;
        var frameRight = boxWidth - 0.5;
        var outerY = isUp ? 0.5 : boxHeight - 0.5;
        var seamY = isUp ? boxHeight : 0;
        var towardSeam = isUp ? 1 : -1; /* 外側の辺から継ぎ目へ向かう向き / direction from the outer edge to the seam */
        var radius = STEPPER_CORNER_RADIUS;
        var arcSteps = 4; /* 角丸1つを何本の線分で近似するか / segments per corner */
        var angle, k;

        boxGraphics.newPath();
        boxGraphics.moveTo(frameLeft, seamY);
        /* 左の角丸 / left corner */
        for (k = 0; k <= arcSteps; k++) {
            angle = (Math.PI / 2) * k / arcSteps;
            boxGraphics.lineTo(frameLeft + radius - radius * Math.cos(angle), outerY + towardSeam * (radius - radius * Math.sin(angle)));
        }
        /* 右の角丸 / right corner */
        for (k = 0; k <= arcSteps; k++) {
            angle = (Math.PI / 2) * k / arcSteps;
            boxGraphics.lineTo(frameRight - radius + radius * Math.sin(angle), outerY + towardSeam * (radius - radius * Math.cos(angle)));
        }
        boxGraphics.lineTo(frameRight, seamY);
        boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, frameColor, 1));
    }

    /**
     * 山形（∧／∨）を描く。文字グリフの▲▼は上下で大きさやベースラインが揃わないため、線で描く
     * @param {ScriptUIGraphics} boxGraphics - 描画先
     * @param {number} boxWidth - ボタンの幅
     * @param {number} boxHeight - ボタンの高さ
     * @param {boolean} isUp - ∧なら true、∨なら false
     * @param {number[]} chevronColor - [r, g, b, a]
     * @returns {void}
     */
    function drawStepperChevron(boxGraphics, boxWidth, boxHeight, isUp, chevronColor) {
        var centerX = boxWidth / 2;
        var centerY = isUp ? boxHeight / 2 + 0.5 : boxHeight / 2 - 0.5; /* 継ぎ目から少し離す / nudged away from the seam */
        var halfWidth = 3.6; /* 山形の半幅（高さ1.8に対して開き約127°） / half width of the chevron */
        var tipOffsetY = isUp ? -1.8 : 1.8; /* 頂点の中心からのずれ（上向きは上、下向きは下） */
        boxGraphics.newPath();
        boxGraphics.moveTo(centerX - halfWidth, centerY - tipOffsetY);
        boxGraphics.lineTo(centerX, centerY + tipOffsetY);
        boxGraphics.lineTo(centerX + halfWidth, centerY - tipOffsetY);
        boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, chevronColor, 1.2));
    }

    /**
     * コントロールと、その親をたどってすべて有効かを返す（親の無効化は子の enabled に出ない）
     * @param {Object} control - 対象のコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isStepperEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (!node.enabled) return false;
        }
        return true;
    }

    /**
     * コンテナ以下にある∧∨ボタンをすべて描き直す。行やパネルの enabled を切り替えたあとに呼ぶ
     * @param {Object} container - 行・グループ・パネルなど
     * @returns {void}
     */
    function redrawSteppersIn(container) {
        if (!container.children) return;
        for (var i = 0; i < container.children.length; i++) {
            var child = container.children[i];
            if (child.isStepperButton) redrawStepperGroup(child);
            else redrawSteppersIn(child);
        }
    }

    /**
     * group の onDraw を呼び直す。group には notify() が無いため、隠して再表示して描き直させる
     * @param {Group} targetGroup - 描き直す group
     * @returns {void}
     */
    function redrawStepperGroup(targetGroup) {
        targetGroup.hide();
        targetGroup.show();
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ステップボタン（再利用パーツ）ここまで / End of the reusable stepper
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * チャンネル名と数値欄の組を1行追加する
     * @param {Group} parent - 追加先の列
     * @param {string} channelName - "C" / "R" などのチャンネル名
     * @param {string} defaultValue - 初期値
     * @param {Object} rangeOptions - ∧∨の min / max / integer
     * @param {Function} [onUpdate] - 値が変わったときに呼ぶ処理
     * @returns {EditText} 追加した数値欄（∧∨は .stepperGroup で参照できる）
     */
    function addChannelInput(parent, channelName, defaultValue, rangeOptions, onUpdate) {
        var channelRow = parent.add("group");
        var channelLabel = channelRow.add("statictext", undefined, appendColon(channelName));
        channelLabel.preferredSize.width = CHANNEL_LABEL_WIDTH;

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = channelRow.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;
        var channelStepper = addStepper(stepperInputGroup, function () { return channelInput; }, {
            step: 1, min: rangeOptions.min, max: rangeOptions.max, integer: rangeOptions.integer,
            onStep: function () { if (onUpdate) onUpdate(); }
        });
        var channelInput = stepperInputGroup.add("edittext", undefined, defaultValue);
        channelInput.characters = 4;
        channelInput.preferredSize.width = CHANNEL_INPUT_WIDTH;
        channelInput.stepperGroup = channelStepper;
        bindSteppedArrowKeys(channelInput, channelStepper);
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
            /* ∧∨も一緒に切り替えて描き直す（HEX 欄には∧∨が無い）/ toggle and redraw the stepper too (the hex field has none) */
            if (inputs[i].stepperGroup) {
                inputs[i].stepperGroup.enabled = isEnabled;
                redrawSteppersIn(inputs[i].stepperGroup);
            }
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
        /* CMYK は 0〜100%（小数可）、RGB は 0〜255 の整数 / CMYK 0-100 (decimals ok), RGB integers 0-255 */
        var cmykRange = { min: 0, max: 100 };
        var rgbRange = { min: 0, max: 255, integer: true };

        colorInputs.c = addChannelInput(cmykColumn, "C", DEFAULT_CMYK[0], cmykRange);
        colorInputs.m = addChannelInput(cmykColumn, "M", DEFAULT_CMYK[1], cmykRange);
        colorInputs.y = addChannelInput(cmykColumn, "Y", DEFAULT_CMYK[2], cmykRange);
        colorInputs.k = addChannelInput(cmykColumn, "K", DEFAULT_CMYK[3], cmykRange);

        colorInputs.r = addChannelInput(rgbColumn, "R", DEFAULT_RGB[0], rgbRange, syncHex);
        colorInputs.g = addChannelInput(rgbColumn, "G", DEFAULT_RGB[1], rgbRange, syncHex);
        colorInputs.b = addChannelInput(rgbColumn, "B", DEFAULT_RGB[2], rgbRange, syncHex);

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
        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var marginStepperGroup = marginRow.add("group");
        marginStepperGroup.orientation = "row";
        marginStepperGroup.alignChildren = ["left", "center"];
        marginStepperGroup.spacing = 0;
        marginStepperGroup.margins = 0;
        /* マージンは負の値も許す（アートボードより内側）/ negative margins are allowed */
        var marginStepper = addStepper(marginStepperGroup, function () { return marginInput; }, { step: 1 });
        var marginInput = marginStepperGroup.add("edittext", undefined, DEFAULT_MARGIN);
        marginInput.helpTip = getLabel("tooltip.margin");
        marginInput.characters = 4;
        bindSteppedArrowKeys(marginInput, marginStepper);
        marginRow.add("statictext", undefined, getUnitInfo().label);

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
