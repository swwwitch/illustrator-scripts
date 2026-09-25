#target illustrator
#targetengine "DirectPrefs"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

環境設定の「角度の制限」と「キー増加」の変更、およびガイド・グリッドの表示やロックの切り替えをパレットから行います。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DirectPrefs.md

### Overview

A palette for changing the constrain angle and the keyboard increment, and for toggling the display and lock state of guides and the grid.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DirectPrefs.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "DirectPrefs";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                             /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DirectPrefs.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DirectPrefs.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 角度の制限のプリセット（アイソメトリック作図の3軸＋0°）/ Constrain-angle presets (the three isometric axes plus 0°) */
    var CONSTRAIN_PRESETS = [0, 150, 90, 30];

    /* キー増加のプリセット（定規単位コードごと。値は各単位そのままの数値）
       / Keyboard-increment presets per ruler unit code (values are in that unit) */
    var KEY_INCREMENT_PRESETS = {
        "1": [1, 5, 10],   /* mm */
        "2": [1, 6, 12],   /* pt */
        "6": [0.1, 1, 8]   /* px */
    };

    /* 上表にない単位（in / cm / Q/H など）で使うプリセット / Presets used for units missing from the table (in, cm, Q/H, ...) */
    var KEY_INCREMENT_PRESETS_DEFAULT = [1, 5, 10];

    // =========================================
    // レイアウト / Layout
    // =========================================

    var PANEL_MARGINS = [16, 20, 16, 12];
    var PANEL_SPACING = 8;
    var INPUT_CHARS = 6;
    var PRESET_BUTTON_WIDTH = 48;
    var UNIT_LABEL_WIDTH = 32;

    /**
     * パネルの共通設定をまとめて適用する
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - 子どうしの間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * グループの共通設定をまとめて適用する（row は縦中央、column は左揃え）
     * @param {Group} targetGroup - 対象のグループ
     * @param {string} [orientation] - "row" / "column"（省略時は "column"）
     * @param {number} [spacing] - 子どうしの間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupGroup(targetGroup, orientation, spacing) {
        var groupOrientation = orientation || "column";
        targetGroup.orientation = groupOrientation;
        /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
        targetGroup.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
        targetGroup.alignment = "fill";
        targetGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * UI の表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILanguage();

    /* ラベル定義 / Label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "DirectPrefs", en: "DirectPrefs" }
        },
        panel: {
            constrain: { ja: "角度の制限", en: "Constrain Angle" },
            keyIncrement: { ja: "キー増加", en: "Keyboard Increment" },
            guide: { ja: "ガイド", en: "Guides" },
            grid: { ja: "グリッド", en: "Grid" }
        },
        fieldLabel: {
            constrain: { ja: "角度の制限", en: "Constrain angle" },
            keyIncrement: { ja: "キー増加", en: "Keyboard increment" }
        },
        button: {
            applyConstrain: { ja: "「角度の制限」の値を変更", en: "Change constrain angle" },
            applyKeyIncrement: { ja: "変更", en: "Change" },
            resetConstrain: { ja: "リセット", en: "Reset" },
            toggleVisibility: { ja: "表示・非表示", en: "Show/Hide" },
            toggleLock: { ja: "ロック・ロック解除", en: "Lock/Unlock" },
            snapToGrid: { ja: "グリッドにスナップ", en: "Snap to Grid" }
        },
        tooltip: {
            constrainInput: {
                ja: "ビューの回転角度が候補として入ります（ドキュメントがないときは現在の値）。［「角度の制限」の値を変更］で適用します。",
                en: "Prefilled with the view rotation (the current value when no document is open). Click \"Change constrain angle\" to apply it."
            },
            constrainPreset: { ja: "この角度を「角度の制限」にすぐ適用します。", en: "Applies this angle to the Constrain Angle preference right away." },
            resetConstrain: { ja: "角度の制限を0°に戻します。", en: "Resets the constrain angle to 0°." },
            keyIncrementPreset: { ja: "この値をキー増加にすぐ適用します。", en: "Applies this value to the keyboard increment right away." },
            keyIncrementInput: { ja: "定規の単位で入力します。［変更］で適用します。", en: "Entered in the ruler unit. Click Change to apply it." },
            snapToGrid: { ja: "「グリッドにスナップ」のオン・オフを切り替えます。", en: "Turns Snap to Grid on or off." }
        },
        status: {
            applied: { ja: "制限角度に適用しました。", en: "Applied to the constrain angle." },
            appliedKeyIncrement: { ja: "キー増加に適用しました。", en: "Applied to the keyboard increment." },
            resetConstrain: { ja: "制限角度を0°にリセットしました。", en: "Reset the constrain angle to 0°." },
            toggledGuideVisibility: { ja: "ガイドの表示を切り替えました。", en: "Toggled guide visibility." },
            toggledGuideLock: { ja: "ガイドのロックを切り替えました。", en: "Toggled the guide lock." },
            toggledGridVisibility: { ja: "グリッドの表示を切り替えました。", en: "Toggled grid visibility." },
            toggledSnapToGrid: { ja: "グリッドにスナップを切り替えました。", en: "Toggled snap to grid." }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            invalidAngle: { ja: "角度には数値を入力してください。", en: "Please enter a numeric angle." },
            invalidValue: { ja: "0以上の数値を入力してください。", en: "Please enter a number of 0 or greater." },
            error: { ja: "エラーが発生しました：", en: "An error occurred:" }
        }
    };

    /**
     * ドット区切りのパスで表示言語のラベルを取り出す
     * @param {string} labelPath - LABELS 内のパス（例: "button.resetConstrain"）
     * @returns {string} 表示言語のラベル
     */
    function getLabel(labelPath) {
        var pathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
        }
        return labelNode[uiLang];
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * ワーカーの結果文字列を表示用のステータス文にする（既知のコードは専用文、未知のものは汎用エラー）
     * @param {string} result - ワーカーから返った結果文字列
     * @returns {string} ステータス文
     */
    function statusFromResult(result) {
        if (result.indexOf("NODOC") !== -1) {
            return getLabel("alert.noDocument");
        }
        return getLabel("alert.error") + " " + result;
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位の定義を1か所に集約（コード／ラベル／pt換算係数／表示桁数）
       decimals：1pt 未満に潰れないように、大きい単位ほど桁数を増やす（in で 1mm ≒ 0.039）
       / Single source of unit definitions (code, label, pt factor, display decimals)
       decimals: larger units need more digits so small values do not collapse to 0 (1mm is 0.039in) */
    var UNITS = [
        { label: "in",    pointsPerUnit: 72,               decimals: 3 },  /* 0 */
        { label: "mm",    pointsPerUnit: 72 / 25.4,        decimals: 1 },  /* 1 */
        { label: "pt",    pointsPerUnit: 1,                decimals: 1 },  /* 2 */
        { label: "pica",  pointsPerUnit: 12,               decimals: 2 },  /* 3 */
        { label: "cm",    pointsPerUnit: 72 / 2.54,        decimals: 2 },  /* 4 */
        { label: "Q",     pointsPerUnit: 72 / 25.4 * 0.25, decimals: 1 },  /* 5 */
        { label: "px",    pointsPerUnit: 1,                decimals: 1 },  /* 6 */
        { label: "ft/in", pointsPerUnit: 72 * 12,          decimals: 4 },  /* 7 */
        { label: "m",     pointsPerUnit: 72 / 25.4 * 1000, decimals: 4 },  /* 8 */
        { label: "yd",    pointsPerUnit: 72 * 36,          decimals: 4 },  /* 9 */
        { label: "ft",    pointsPerUnit: 72 * 12,          decimals: 4 }   /* 10 */
    ];

    /* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
       Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 単位コードから単位定義を取得する（未対応コードは pt 相当）
     * @param {number} unitCode - 環境設定の単位コード
     * @returns {{label: string, pointsPerUnit: number, decimals: number}} 単位定義
     */
    function getUnitByCode(unitCode) {
        return UNITS[unitCode] || UNITS[2];
    }

    /**
     * 設定キーごとの単位情報を取得する
     * @param {string} prefKey - 環境設定キー（省略時は "rulerType"）
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        var unit = UNITS[unitCode] || UNITS[2];
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    /**
     * 単位コードに対応するキー増加のプリセットを取得する
     * @param {number} unitCode - 定規の単位コード
     * @returns {number[]} プリセットの値（その単位での数値）
     */
    function getKeyIncrementPresets(unitCode) {
        var presets = KEY_INCREMENT_PRESETS[String(unitCode)];
        return presets ? presets : KEY_INCREMENT_PRESETS_DEFAULT;
    }

    // =========================================
    // 角度の計算 / Angle helpers
    // =========================================

    /**
     * 角度を -180〜180 に正規化する
     * @param {number} angle - 角度（度）
     * @returns {number} 正規化した角度
     */
    function normalizeAngle(angle) {
        var normalized = angle % 360;
        if (normalized > 180) { normalized -= 360; }
        if (normalized < -180) { normalized += 360; }
        return normalized;
    }

    /**
     * 表示用に小数2桁へ丸める（atan2 由来の 30.00000001 のような桁あふれを抑える）
     * @param {number} angle - 角度（度）
     * @returns {number} 丸めた角度
     */
    function roundAngle(angle) {
        return Math.round(angle * 100) / 100;
    }

    // =========================================
    // メインエンジンへの委譲 / Delegation to the main engine
    // =========================================

    /**
     * メインエンジンでコードを実行する（常駐パレットの app は DOM 接続を失うため）。
     * 本文は encodeURIComponent + eval で送り、バックスラッシュ・多バイト文字を無傷で渡す
     * @param {string} code - 実行するコード
     * @param {Function} onResult - 結果文字列を受け取るコールバック（失敗時は "ERR:" で始まる）
     * @returns {void}
     */
    function runInMainEngine(code, onResult) {
        var bridgeTalk = new BridgeTalk();
        bridgeTalk.target = "illustrator";
        bridgeTalk.body = 'eval(decodeURIComponent("' + encodeURIComponent(code) + '"));';
        bridgeTalk.onResult = function (response) {
            onResult(String(response.body));
        };
        bridgeTalk.onError = function (response) {
            onResult("ERR:" + String(response.body));
        };
        bridgeTalk.send();
    }

    /**
     * ワーカーの本文を IIFE で包む
     * @param {string} workerCode - ワーカーの本文
     * @returns {string} IIFE で包んだコード
     */
    function wrapWorkerBody(workerCode) {
        return "(function(){" + workerCode + "})()";
    }

    /* ワーカー断片：「角度の制限」を読む。実際の拘束方向は constrain/sin・constrain/cos が持っているため、
       constrain/angle ではなくこの2つから角度を復元する（angle は書いても拘束に反映されない）
       / Worker fragment: read the constrain angle. The real constraint direction lives in constrain/sin and
       constrain/cos, so recover the angle from those instead of constrain/angle (writing `angle` alone has no effect) */
    var WORKER_READ_CONSTRAIN =
        "var constrainAngle=Math.atan2(app.preferences.getRealPreference('constrain/sin')," +
        "app.preferences.getRealPreference('constrain/cos'))*180/Math.PI;";

    /**
     * 「角度の制限」を書き込むワーカー断片を作る。angle は環境設定ダイアログの表示用で度、
     * sin・cos は実際の拘束方向でラジアン由来。angle だけでは拘束に効かないので3つとも書く
     * @param {number} angle - 書き込む角度（度）
     * @returns {string} ワーカー断片
     */
    function buildConstrainWriteCode(angle) {
        return "var radians=(" + angle + ")*Math.PI/180;" +
            "app.preferences.setRealPreference('constrain/angle'," + angle + ");" +
            "app.preferences.setRealPreference('constrain/sin',Math.sin(radians));" +
            "app.preferences.setRealPreference('constrain/cos',Math.cos(radians));";
    }

    /**
     * ビュー回転角度・制限角度・定規単位・キー増加(pt)を1回の委譲でまとめて取得する
     * （"OK:回転,制限,単位コード,キー増加"。回転はドキュメントが開いていなければ空）。
     * 環境設定はドキュメントがなくても読めるため、回転だけを条件付きにしている
     * @param {Function} onResult - 結果文字列を受け取るコールバック
     * @returns {void}
     */
    function fetchState(onResult) {
        runInMainEngine(wrapWorkerBody(
            "var viewRotation=(app.documents.length>0)?app.activeDocument.activeView.rotateAngle:'';" +
            WORKER_READ_CONSTRAIN +
            "var rulerUnitCode=app.preferences.getIntegerPreference('rulerType');" +
            "var keyIncrementPt=app.preferences.getRealPreference('cursorKeyLength');" +
            "return 'OK:'+viewRotation+','+constrainAngle+','+rulerUnitCode+','+keyIncrementPt;"
        ), onResult);
    }

    /**
     * 制限角度をメインエンジンで環境設定に適用する
     * @param {number} angle - 制限角度（度）
     * @param {Function} onResult - 結果文字列を受け取るコールバック
     * @returns {void}
     */
    function applyConstrainAngle(angle, onResult) {
        runInMainEngine(wrapWorkerBody(
            buildConstrainWriteCode(angle) +
            "return 'OK';"
        ), onResult);
    }

    /**
     * キー増加をメインエンジンで環境設定に適用する（値は pt）
     * @param {number} lengthPt - キー増加（pt）
     * @param {Function} onResult - 結果文字列を受け取るコールバック
     * @returns {void}
     */
    function applyKeyIncrement(lengthPt, onResult) {
        runInMainEngine(wrapWorkerBody(
            "app.preferences.setRealPreference('cursorKeyLength'," + lengthPt + ");" +
            "return 'OK';"
        ), onResult);
    }

    /**
     * メニューコマンドをメインエンジンで実行する（ガイド・グリッドの状態は取得できないため、メニューのトグルを呼ぶ）
     * @param {string} menuCommand - メニューコマンド名
     * @param {Function} onResult - 結果文字列を受け取るコールバック
     * @returns {void}
     */
    function runMenuCommand(menuCommand, onResult) {
        runInMainEngine(wrapWorkerBody(
            "if(app.documents.length===0){return 'ERR:NODOC';}" +
            "app.executeMenuCommand('" + menuCommand + "');" +
            "return 'OK';"
        ), onResult);
    }

    // =========================================
    // パレット / Palette
    // =========================================

    /**
     * パレットと各コントロールを作る
     * @returns {Object} パレット（palette）と各コントロールの参照
     */
    function buildPalette() {
        var prefsPalette = new Window("palette", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        prefsPalette.orientation = "column";
        prefsPalette.alignChildren = "fill";
        prefsPalette.margins = 16;
        prefsPalette.spacing = 12;

        /* 角度の制限を変更するパネル / Panel for changing the constrain angle */
        var constrainPanel = prefsPalette.add("panel", undefined, getLabel("panel.constrain"));
        setupPanel(constrainPanel, 6);

        /* 角度の制限（編集可。ビューの回転角度が候補値として入るが、反映は適用ボタンを押したときだけ）
           / Constrain angle (editable; seeded with the view rotation as a suggestion, but committed only on the Apply button) */
        var constrainInputGroup = constrainPanel.add("group");
        setupGroup(constrainInputGroup, "row");
        constrainInputGroup.add("statictext", undefined, labelText("fieldLabel.constrain"));
        var constrainInput = constrainInputGroup.add("edittext", undefined, "");
        constrainInput.characters = INPUT_CHARS;
        constrainInput.helpTip = getLabel("tooltip.constrainInput");
        constrainInputGroup.add("statictext", undefined, "°");

        /* プリセットボタン行（押すとその角度を即座に適用）/ Preset button row (clicking applies that angle immediately) */
        var constrainPresetGroup = constrainPanel.add("group");
        setupGroup(constrainPresetGroup, "row", 4);
        var constrainPresetButtons = [];
        for (var i = 0; i < CONSTRAIN_PRESETS.length; i++) {
            var constrainPresetButton = constrainPresetGroup.add("button", undefined, CONSTRAIN_PRESETS[i] + "°");
            constrainPresetButton.preferredSize.width = PRESET_BUTTON_WIDTH;
            constrainPresetButton.helpTip = getLabel("tooltip.constrainPreset");
            constrainPresetButtons.push({ angle: CONSTRAIN_PRESETS[i], button: constrainPresetButton });
        }

        /* ボタン行：適用とリセット / Button row: Apply and Reset */
        var constrainButtonGroup = constrainPanel.add("group");
        setupGroup(constrainButtonGroup, "row");
        constrainButtonGroup.alignment = "right";

        /* 角度の制限を0°に戻す / Reset the constrain angle to 0° */
        var resetConstrainButton = constrainButtonGroup.add("button", undefined, getLabel("button.resetConstrain"));
        resetConstrainButton.helpTip = getLabel("tooltip.resetConstrain");

        /* 「角度の制限」の値を変更ボタン / Change-constrain-angle button */
        var applyConstrainButton = constrainButtonGroup.add("button", undefined, getLabel("button.applyConstrain"));

        /* キー増加を変更するパネル / Panel for changing the keyboard increment */
        var keyIncrementPanel = prefsPalette.add("panel", undefined, getLabel("panel.keyIncrement"));
        setupPanel(keyIncrementPanel, 6);

        /* キー増加（定規単位で表示・入力。単位ラベルは rulerType に追従）
           / Keyboard increment (shown and entered in the ruler unit; the unit label follows rulerType) */
        var keyIncrementInputGroup = keyIncrementPanel.add("group");
        setupGroup(keyIncrementInputGroup, "row");
        keyIncrementInputGroup.add("statictext", undefined, labelText("fieldLabel.keyIncrement"));
        var keyIncrementInput = keyIncrementInputGroup.add("edittext", undefined, "");
        keyIncrementInput.characters = INPUT_CHARS;
        keyIncrementInput.helpTip = getLabel("tooltip.keyIncrementInput");
        var keyIncrementUnitLabel = keyIncrementInputGroup.add("statictext", undefined, "pt");
        keyIncrementUnitLabel.preferredSize.width = UNIT_LABEL_WIDTH;

        /* プリセットボタン行（押すとその値を即座に適用。単位が変わっても作り直さず、表示値だけ差し替える）
           / Preset button row (clicking applies that value immediately; the buttons persist and only their values follow the ruler unit) */
        var keyIncrementPresetGroup = keyIncrementPanel.add("group");
        setupGroup(keyIncrementPresetGroup, "row", 4);
        var keyIncrementPresetButtons = [];
        for (var j = 0; j < KEY_INCREMENT_PRESETS_DEFAULT.length; j++) {
            var keyIncrementPresetButton = keyIncrementPresetGroup.add("button", undefined, "");
            keyIncrementPresetButton.preferredSize.width = PRESET_BUTTON_WIDTH;
            keyIncrementPresetButton.helpTip = getLabel("tooltip.keyIncrementPreset");
            keyIncrementPresetButtons.push(keyIncrementPresetButton);
        }

        /* 「キー増加」の値を変更ボタン / Change-keyboard-increment button */
        var applyKeyIncrementButton = keyIncrementPanel.add("button", undefined, getLabel("button.applyKeyIncrement"));
        applyKeyIncrementButton.alignment = "right";

        /* ガイドのパネル / Guides panel */
        var guidePanel = prefsPalette.add("panel", undefined, getLabel("panel.guide"));
        setupPanel(guidePanel, 6);

        var guideButtonGroup = guidePanel.add("group");
        setupGroup(guideButtonGroup, "row", 4);

        /* ガイドの表示・非表示を切り替えるボタン / Button that toggles guide visibility */
        var toggleGuideVisibilityButton = guideButtonGroup.add("button", undefined, getLabel("button.toggleVisibility"));

        /* ガイドのロック・ロック解除を切り替えるボタン / Button that toggles the guide lock */
        var toggleGuideLockButton = guideButtonGroup.add("button", undefined, getLabel("button.toggleLock"));

        /* グリッドのパネル / Grid panel */
        var gridPanel = prefsPalette.add("panel", undefined, getLabel("panel.grid"));
        setupPanel(gridPanel, 6);

        var gridButtonGroup = gridPanel.add("group");
        setupGroup(gridButtonGroup, "row", 4);

        /* グリッドの表示・非表示を切り替えるボタン / Button that toggles grid visibility */
        var toggleGridVisibilityButton = gridButtonGroup.add("button", undefined, getLabel("button.toggleVisibility"));

        /* グリッドにスナップを切り替えるボタン / Button that toggles snap to grid */
        var snapToGridButton = gridButtonGroup.add("button", undefined, getLabel("button.snapToGrid"));
        snapToGridButton.helpTip = getLabel("tooltip.snapToGrid");

        /* ステータス表示 / Status line */
        var statusText = prefsPalette.add("statictext", undefined, "");
        statusText.alignment = "fill";

        return {
            palette: prefsPalette,
            constrainInput: constrainInput,
            constrainPresetButtons: constrainPresetButtons,
            resetConstrainButton: resetConstrainButton,
            applyConstrainButton: applyConstrainButton,
            keyIncrementInput: keyIncrementInput,
            keyIncrementUnitLabel: keyIncrementUnitLabel,
            keyIncrementPresetButtons: keyIncrementPresetButtons,
            applyKeyIncrementButton: applyKeyIncrementButton,
            toggleGuideVisibilityButton: toggleGuideVisibilityButton,
            toggleGuideLockButton: toggleGuideLockButton,
            toggleGridVisibilityButton: toggleGridVisibilityButton,
            snapToGridButton: snapToGridButton,
            statusText: statusText
        };
    }

    /**
     * パレットの表示更新と各コントロールの操作を結び付ける
     * @param {Object} paletteUI - buildPalette() が返したパレットとコントロールの参照
     * @returns {void}
     */
    function bindPaletteHandlers(paletteUI) {
        /* 現在の制限角度と定規単位（リセットボタンのディム判定・単位換算に使用）
           / Current constrain angle and ruler unit (used to dim the Reset button and to convert units) */
        var currentConstrain = 0;
        var currentUnit = getUnitByCode(2);

        /* キー増加のプリセットボタンが押された時点で読む値 / Values the keyboard-increment preset buttons read at click time */
        var keyIncrementPresetValues = KEY_INCREMENT_PRESETS_DEFAULT;

        /**
         * ワーカーの結果を受けるコールバックを作る。失敗ならステータスにエラーを出し、成功なら onSuccess を呼ぶ
         * @param {Function} onSuccess - 成功時の処理
         * @returns {Function} runInMainEngine に渡すコールバック
         */
        function handleWorkerResult(onSuccess) {
            return function (result) {
                if (result.indexOf("OK") === 0) {
                    onSuccess();
                } else {
                    paletteUI.statusText.text = statusFromResult(result);
                }
            };
        }

        /**
         * 適用済みの制限角度を状態に記録し、0°ならリセットボタンをディムする
         * @param {number} angle - 制限角度（度）
         * @returns {void}
         */
        function setConstrain(angle) {
            currentConstrain = angle;
            paletteUI.resetConstrainButton.enabled = (currentConstrain !== 0);
        }

        /**
         * 制限角度を適用して表示・状態を更新する
         * @param {number} angle - 制限角度（度）
         * @returns {void}
         */
        function commitConstrain(angle) {
            applyConstrainAngle(angle, handleWorkerResult(function () {
                setConstrain(angle);
                paletteUI.constrainInput.text = roundAngle(angle);
                paletteUI.statusText.text = getLabel("status.applied");
            }));
        }

        /**
         * キー増加(pt)を現在の定規単位の表示文字列に変換する
         * @param {number} lengthPt - キー増加（pt）
         * @returns {string} 表示用の文字列
         */
        function formatKeyIncrement(lengthPt) {
            return (lengthPt / currentUnit.pointsPerUnit).toFixed(currentUnit.decimals);
        }

        /**
         * キー増加を適用して表示を更新する（値は現在の定規単位）
         * @param {number} unitValue - 定規単位での値
         * @returns {void}
         */
        function commitKeyIncrement(unitValue) {
            var lengthPt = unitValue * currentUnit.pointsPerUnit;
            applyKeyIncrement(lengthPt, handleWorkerResult(function () {
                paletteUI.keyIncrementInput.text = formatKeyIncrement(lengthPt);
                paletteUI.statusText.text = getLabel("status.appliedKeyIncrement");
            }));
        }

        /**
         * 現在の定規単位に合わせてプリセットボタンの値とラベルを差し替える
         * @returns {void}
         */
        function updateKeyIncrementPresets() {
            keyIncrementPresetValues = getKeyIncrementPresets(currentUnit.code);
            for (var i = 0; i < paletteUI.keyIncrementPresetButtons.length; i++) {
                var hasValue = (i < keyIncrementPresetValues.length);
                paletteUI.keyIncrementPresetButtons[i].visible = hasValue;
                if (hasValue) {
                    paletteUI.keyIncrementPresetButtons[i].text = keyIncrementPresetValues[i] + currentUnit.label;
                }
            }
        }

        /**
         * ビュー回転角度・制限角度・定規単位・キー増加を取得し、各表示へ反映する
         * （制限角度の入力欄にはビューの回転角度を候補値として入れる。適用はボタン押下時のみ）
         * @returns {void}
         */
        function refreshState() {
            fetchState(function (result) {
                if (result.indexOf("OK:") !== 0) {
                    paletteUI.statusText.text = statusFromResult(result);
                    return;
                }
                var stateValues = result.substring(3).split(",");
                setConstrain(parseFloat(stateValues[1]));
                currentUnit = getUnitByCode(parseInt(stateValues[2], 10));
                paletteUI.keyIncrementUnitLabel.text = currentUnit.label;
                paletteUI.keyIncrementInput.text = formatKeyIncrement(parseFloat(stateValues[3]));
                updateKeyIncrementPresets();
                if (stateValues[0] === "") {
                    /* ドキュメントがなければ回転を取得できないので、現在の制限角度をそのまま表示
                       / Without a document there is no rotation to read, so show the current constrain angle as-is */
                    paletteUI.constrainInput.text = roundAngle(currentConstrain);
                    paletteUI.statusText.text = getLabel("alert.noDocument");
                } else {
                    paletteUI.constrainInput.text = roundAngle(normalizeAngle(parseFloat(stateValues[0])));
                    paletteUI.statusText.text = "";
                }
            });
        }

        /**
         * 角度の制限のプリセットボタンに適用処理を付ける（ES3にはブロックスコープがないため、クロージャを関数で切り出す）
         * @param {Object} constrainPreset - プリセットの角度（angle）とボタン（button）の対
         * @returns {void}
         */
        function bindConstrainPreset(constrainPreset) {
            constrainPreset.button.onClick = function () {
                commitConstrain(constrainPreset.angle);
            };
        }

        /**
         * キー増加のプリセットボタンに適用処理を付ける（押した時点の keyIncrementPresetValues を参照する）
         * @param {Button} presetButton - 対象のボタン
         * @param {number} presetIndex - プリセットの番号
         * @returns {void}
         */
        function bindKeyIncrementPreset(presetButton, presetIndex) {
            presetButton.onClick = function () {
                commitKeyIncrement(keyIncrementPresetValues[presetIndex]);
            };
        }

        /**
         * メニューコマンドをトグルボタンに割り当てる
         * @param {Button} toggleButton - 対象のボタン
         * @param {string} menuCommand - メニューコマンド名
         * @param {string} statusPath - 成功時に表示するステータスのラベルパス
         * @returns {void}
         */
        function bindMenuCommand(toggleButton, menuCommand, statusPath) {
            toggleButton.onClick = function () {
                runMenuCommand(menuCommand, handleWorkerResult(function () {
                    paletteUI.statusText.text = getLabel(statusPath);
                }));
            };
        }

        for (var i = 0; i < paletteUI.constrainPresetButtons.length; i++) {
            bindConstrainPreset(paletteUI.constrainPresetButtons[i]);
        }
        for (var j = 0; j < paletteUI.keyIncrementPresetButtons.length; j++) {
            bindKeyIncrementPreset(paletteUI.keyIncrementPresetButtons[j], j);
        }

        /* リセット：制限角度を0°に戻して入力欄と状態を更新 / Reset: set the constrain angle to 0° and refresh the field and state */
        paletteUI.resetConstrainButton.onClick = function () {
            applyConstrainAngle(0, handleWorkerResult(function () {
                setConstrain(0);
                paletteUI.constrainInput.text = "0";
                paletteUI.statusText.text = getLabel("status.resetConstrain");
            }));
        };

        /* 適用ボタン：入力値を検証してメインエンジンで適用 / Apply button: validate the input and apply in the main engine */
        paletteUI.applyConstrainButton.onClick = function () {
            var constrainAngle = parseFloat(paletteUI.constrainInput.text);
            if (isNaN(constrainAngle)) {
                paletteUI.statusText.text = getLabel("alert.invalidAngle");
                return;
            }
            commitConstrain(constrainAngle);
        };

        /* キー増加の適用ボタン：入力値（定規単位）を pt に換算して適用 / Keyboard-increment apply button: convert the entered value (ruler unit) to pt and apply */
        paletteUI.applyKeyIncrementButton.onClick = function () {
            var unitValue = parseFloat(paletteUI.keyIncrementInput.text);
            if (isNaN(unitValue) || unitValue < 0) {
                paletteUI.statusText.text = getLabel("alert.invalidValue");
                return;
            }
            commitKeyIncrement(unitValue);
        };

        bindMenuCommand(paletteUI.toggleGuideVisibilityButton, "showguide", "status.toggledGuideVisibility");
        bindMenuCommand(paletteUI.toggleGuideLockButton, "lockguide", "status.toggledGuideLock");
        bindMenuCommand(paletteUI.toggleGridVisibilityButton, "showgrid", "status.toggledGridVisibility");
        bindMenuCommand(paletteUI.snapToGridButton, "snapgrid", "status.toggledSnapToGrid");

        /* 初期表示時、およびパレットがアクティブになるたびに最新の状態を取得
           / Fetch the latest state on first show and whenever the palette becomes active */
        paletteUI.palette.onShow = refreshState;
        paletteUI.palette.onActivate = refreshState;
    }

    /* 常駐エンジンにパレット参照を保持するキー（CloseAllPalettes.jsx から閉じるため）
       Key holding the palette reference in the persistent engine (so CloseAllPalettes.jsx can close it) */
    var PALETTE_GLOBAL_KEY = "__directPrefsPalette";

    var paletteUI = buildPalette();
    bindPaletteHandlers(paletteUI);
    /* 閉じたら常駐エンジンの参照をクリア / Clear the persistent-engine reference on close */
    paletteUI.palette.onClose = function () { $.global[PALETTE_GLOBAL_KEY] = null; };
    paletteUI.palette.show();
    $.global[PALETTE_GLOBAL_KEY] = paletteUI.palette;

})();
