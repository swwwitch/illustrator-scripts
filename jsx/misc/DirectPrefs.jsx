#target illustrator
#targetengine "DirectPrefs"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

環境設定の「角度の制限」と「キー増加」の変更、およびガイド・グリッドの表示やロックの切り替えをパレットから行います。

詳細は README を参照してください。

### Overview

A palette for changing the constrain angle and the keyboard increment, and for toggling the display and lock state of guides and the grid.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "DirectPrefs";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DirectPrefs.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DirectPrefs.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    /* 言語判定 / Detect language */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

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
        label: {
            constrain: { ja: "角度の制限", en: "Constrain angle" },
            keyIncrement: { ja: "キー増加", en: "Keyboard increment" }
        },
        button: {
            apply: { ja: "「角度の制限」の値を変更", en: "Change constrain angle" },
            applyKeyIncrement: { ja: "変更", en: "Change" },
            resetConstrain: { ja: "リセット", en: "Reset" },
            toggleVisibility: { ja: "表示・非表示", en: "Show/Hide" },
            toggleLock: { ja: "ロック・ロック解除", en: "Lock/Unlock" },
            snapToGrid: { ja: "グリッドにスナップ", en: "Snap to Grid" }
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

    /* ネストしたラベルをドット区切りパスで取得 / Get a nested label by dot-separated path */
    function getLabel(path) {
        var parts = path.split(".");
        var node = LABELS;
        for (var i = 0; i < parts.length; i++) {
            node = node[parts[i]];
        }
        return node[currentLanguage];
    }

    /* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
    function labelText(path) {
        return getLabel(path) + (currentLanguage === "ja" ? "：" : ":");
    }

    /* 結果文字列を表示用ステータスへ変換（既知コードは専用文、未知は汎用エラー）
       / Convert a result string to status text (known codes get a specific message, unknown ones the generic error) */
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
        { code: 0,  label: "in",    factor: 72.0,                 decimals: 3 },
        { code: 1,  label: "mm",    factor: 72.0 / 25.4,          decimals: 1 },
        { code: 2,  label: "pt",    factor: 1.0,                  decimals: 1 },
        { code: 3,  label: "pica",  factor: 12.0,                 decimals: 2 },
        { code: 4,  label: "cm",    factor: 72.0 / 2.54,          decimals: 2 },
        { code: 5,  label: "Q/H",   factor: 72.0 / 25.4 * 0.25,   decimals: 1 },
        { code: 6,  label: "px",    factor: 1.0,                  decimals: 1 },
        { code: 7,  label: "ft/in", factor: 72.0 * 12.0,          decimals: 4 },
        { code: 8,  label: "m",     factor: 72.0 / 25.4 * 1000.0, decimals: 4 },
        { code: 9,  label: "yd",    factor: 72.0 * 36.0,          decimals: 4 },
        { code: 10, label: "ft",    factor: 72.0 * 12.0,          decimals: 4 }
    ];

    /* コードから単位定義を取得（未対応コードは pt 相当）/ Find a unit definition by code (unsupported codes fall back to pt) */
    function getUnitByCode(code) {
        for (var i = 0; i < UNITS.length; i++) {
            if (UNITS[i].code === code) { return UNITS[i]; }
        }
        return UNITS[2];
    }

    // =========================================
    // 角度の計算 / Angle helpers
    // =========================================

    /* 角度を -180〜180 に正規化 / Normalize an angle into the -180..180 range */
    function normalizeAngle(angle) {
        var a = angle % 360;
        if (a > 180) { a -= 360; }
        if (a < -180) { a += 360; }
        return a;
    }

    /* 表示用に小数2桁へ丸める（atan2 由来の 30.00000001 のような桁あふれを抑える）
       / Round to 2 decimals for display (suppresses float noise like 30.00000001 from atan2) */
    function roundAngle(angle) {
        return Math.round(angle * 100) / 100;
    }

    // =========================================
    // メインエンジンへの委譲 / Delegation to the main engine
    // =========================================

    /* メインエンジンでコードを実行する（常駐パレットの app は DOM 接続を失うため）。
       本文は encodeURIComponent + eval で送り、バックスラッシュ・多バイト文字を無傷で渡す。
       / Run code in the main engine (the palette's app loses DOM access).
       The body is sent via encodeURIComponent + eval so backslashes and multibyte chars survive intact. */
    function runInMainEngine(code, onResult) {
        var bridge = new BridgeTalk();
        bridge.target = "illustrator";
        bridge.body = 'eval(decodeURIComponent("' + encodeURIComponent(code) + '"));';
        bridge.onResult = function (response) {
            onResult(String(response.body));
        };
        bridge.onError = function (response) {
            onResult("ERR:" + String(response.body));
        };
        bridge.send();
    }

    /* worker 本文を IIFE で包む / Wrap a worker body in an IIFE */
    function workerBody(body) {
        return "(function(){" + body + "})()";
    }

    /* worker 断片：「角度の制限」を読む。実際の拘束方向は constrain/sin・constrain/cos が持っているため、
       constrain/angle ではなくこの2つから角度を復元する（angle は書いても拘束に反映されない）
       / Worker fragment: read the constrain angle. The real constraint direction lives in constrain/sin and
       constrain/cos, so recover the angle from those instead of constrain/angle (writing `angle` alone has no effect) */
    var W_CONSTRAIN_GET =
        "var c=Math.atan2(app.preferences.getRealPreference('constrain/sin')," +
        "app.preferences.getRealPreference('constrain/cos'))*180/Math.PI;";

    /* worker 断片を作る：「角度の制限」を書き込む。angle は環境設定ダイアログの表示用で度、
       sin・cos は実際の拘束方向でラジアン由来。angle だけでは拘束に効かないので3つとも書く
       / Build a worker fragment that writes the constrain angle. `angle` is the value shown in the
       Preferences dialog and is in degrees; sin and cos carry the real constraint direction and are
       derived from radians. Writing `angle` alone does not affect the constraint, so all three are written */
    function constrainSetter(deg) {
        return "var rad=(" + deg + ")*Math.PI/180;" +
            "app.preferences.setRealPreference('constrain/angle'," + deg + ");" +
            "app.preferences.setRealPreference('constrain/sin',Math.sin(rad));" +
            "app.preferences.setRealPreference('constrain/cos',Math.cos(rad));";
    }

    /* ビュー回転角度・制限角度・定規単位・キー増加(pt)を1回の委譲でまとめて取得
       （"OK:回転,制限,単位コード,キー増加"。回転はドキュメントが開いていなければ空）
       環境設定はドキュメントがなくても読めるため、回転だけを条件付きにしている。
       / Fetch the view rotation, constrain angle, ruler unit, and keyboard increment (pt) in a single delegation
       ("OK:rotation,constrain,unitCode,increment"; rotation is empty when no document is open).
       Preferences are readable without a document, so only the rotation is conditional. */
    function fetchState(onResult) {
        runInMainEngine(workerBody(
            "var r=(app.documents.length>0)?app.activeDocument.activeView.rotateAngle:'';" +
            W_CONSTRAIN_GET +
            "var u=app.preferences.getIntegerPreference('rulerType');" +
            "var k=app.preferences.getRealPreference('cursorKeyLength');" +
            "return 'OK:'+r+','+c+','+u+','+k;"
        ), onResult);
    }

    /* 制限角度をメインエンジンで環境設定に適用 / Apply the constrain angle to the preference in the main engine */
    function applyConstrainAngle(angle, onResult) {
        runInMainEngine(workerBody(
            constrainSetter(angle) +
            "return 'OK';"
        ), onResult);
    }

    /* 「角度の制限」を0°に戻す / Reset the constrain angle to 0° */
    function resetConstrain(onResult) {
        runInMainEngine(workerBody(
            constrainSetter(0) +
            "return 'OK';"
        ), onResult);
    }

    /* キー増加をメインエンジンで環境設定に適用（値は pt）/ Apply the keyboard increment to the preference in the main engine (value in pt) */
    function applyKeyIncrement(lengthPt, onResult) {
        runInMainEngine(workerBody(
            "app.preferences.setRealPreference('cursorKeyLength'," + lengthPt + ");" +
            "return 'OK';"
        ), onResult);
    }

    /* メニューコマンドをメインエンジンで実行（ガイド・グリッドの状態は取得できないため、メニューのトグルを呼ぶ）
       / Run a menu command in the main engine (guide/grid states are not readable, so the menu toggles are invoked) */
    function runMenuCommand(command, onResult) {
        runInMainEngine(workerBody(
            "if(app.documents.length===0){return 'ERR:NODOC';}" +
            "app.executeMenuCommand('" + command + "');" +
            "return 'OK';"
        ), onResult);
    }

    // =========================================
    // パレット / Palette
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

    /* 単位コードに対応するキー増加のプリセットを取得 / Get the keyboard-increment presets for a unit code */
    function getKeyIncrementPresets(code) {
        var presets = KEY_INCREMENT_PRESETS[String(code)];
        return presets ? presets : KEY_INCREMENT_PRESETS_DEFAULT;
    }

    /* パレットを作成して表示する（IIFEで即時実行）/ Build and show the palette (run immediately as an IIFE) */
    (function () {
        var palette = new Window("palette", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        palette.orientation = "column";
        palette.alignChildren = "fill";
        palette.margins = 16;
        palette.spacing = 12;

        // =========================================
        // レイアウト寸法 / Layout metrics
        // =========================================
        var PANEL_MARGINS = [16, 20, 16, 12];
        var PANEL_SPACING = 8;
        var INPUT_CHARS = 6;
        var PRESET_BUTTON_WIDTH = 48;
        var UNIT_LABEL_WIDTH = 32;

        /* パネルの共通設定 / Apply shared panel layout */
        function setupPanel(panel, spacing) {
            panel.orientation = "column";
            panel.alignChildren = ["fill", "top"];
            panel.alignment = "fill";
            panel.margins = PANEL_MARGINS;
            panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
        }

        /* グループの共通設定（row/column で整列を切り替え）/ Apply shared group layout (alignChildren switches by orientation) */
        function setupGroup(group, orientation, spacing) {
            var groupOrientation = orientation || "column";
            group.orientation = groupOrientation;
            /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
            group.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
            group.alignment = "fill";
            group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
        }

        /* 角度の制限を変更するパネル / Panel for changing the constrain angle */
        var constrainPanel = palette.add("panel", undefined, getLabel("panel.constrain"));
        setupPanel(constrainPanel, 6);

        /* 角度の制限（編集可。ビューの回転角度が候補値として入るが、反映は適用ボタンを押したときだけ）
           / Constrain angle (editable; seeded with the view rotation as a suggestion, but committed only on the Apply button) */
        var constrainGroup = constrainPanel.add("group");
        setupGroup(constrainGroup, "row");
        constrainGroup.add("statictext", undefined, labelText("label.constrain"));
        var constrainInput = constrainGroup.add("edittext", undefined, "");
        constrainInput.characters = INPUT_CHARS;
        constrainGroup.add("statictext", undefined, "°");

        /* プリセットボタン行（押すとその角度を即座に適用）/ Preset button row (clicking applies that angle immediately) */
        var presetGroup = constrainPanel.add("group");
        setupGroup(presetGroup, "row", 4);

        /* ボタン行：適用とリセット / Button row: Apply and Reset */
        var constrainButtonGroup = constrainPanel.add("group");
        setupGroup(constrainButtonGroup, "row");
        constrainButtonGroup.alignment = "right";

        /* 角度の制限を0°に戻す / Reset the constrain angle to 0° */
        var resetConstrainButton = constrainButtonGroup.add("button", undefined, getLabel("button.resetConstrain"));

        /* 「角度の制限」の値を変更ボタン / Change-constrain-angle button */
        var applyButton = constrainButtonGroup.add("button", undefined, getLabel("button.apply"));

        /* キー増加を変更するパネル / Panel for changing the keyboard increment */
        var keyIncrementPanel = palette.add("panel", undefined, getLabel("panel.keyIncrement"));
        setupPanel(keyIncrementPanel, 6);

        /* キー増加（定規単位で表示・入力。単位ラベルは rulerType に追従）
           / Keyboard increment (shown and entered in the ruler unit; the unit label follows rulerType) */
        var keyIncrementGroup = keyIncrementPanel.add("group");
        setupGroup(keyIncrementGroup, "row");
        keyIncrementGroup.add("statictext", undefined, labelText("label.keyIncrement"));
        var keyIncrementInput = keyIncrementGroup.add("edittext", undefined, "");
        keyIncrementInput.characters = INPUT_CHARS;
        var keyIncrementUnit = keyIncrementGroup.add("statictext", undefined, "pt");
        keyIncrementUnit.preferredSize.width = UNIT_LABEL_WIDTH;

        /* プリセットボタン行（押すとその値を即座に適用。表示値は定規単位に追従）
           / Preset button row (clicking applies that value immediately; the values follow the ruler unit) */
        var keyIncrementPresetGroup = keyIncrementPanel.add("group");
        setupGroup(keyIncrementPresetGroup, "row", 4);

        /* 「キー増加」の値を変更ボタン / Change-keyboard-increment button */
        var applyKeyIncrementButton = keyIncrementPanel.add("button", undefined, getLabel("button.applyKeyIncrement"));
        applyKeyIncrementButton.alignment = "right";

        /* ガイドのパネル / Guides panel */
        var guidePanel = palette.add("panel", undefined, getLabel("panel.guide"));
        setupPanel(guidePanel, 6);

        var guideButtonGroup = guidePanel.add("group");
        setupGroup(guideButtonGroup, "row", 4);

        /* ガイドの表示・非表示を切り替えるボタン / Button that toggles guide visibility */
        var toggleGuideVisibilityButton = guideButtonGroup.add("button", undefined, getLabel("button.toggleVisibility"));

        /* ガイドのロック・ロック解除を切り替えるボタン / Button that toggles the guide lock */
        var toggleGuideLockButton = guideButtonGroup.add("button", undefined, getLabel("button.toggleLock"));

        /* グリッドのパネル / Grid panel */
        var gridPanel = palette.add("panel", undefined, getLabel("panel.grid"));
        setupPanel(gridPanel, 6);

        var gridButtonGroup = gridPanel.add("group");
        setupGroup(gridButtonGroup, "row", 4);

        /* グリッドの表示・非表示を切り替えるボタン / Button that toggles grid visibility */
        var toggleGridVisibilityButton = gridButtonGroup.add("button", undefined, getLabel("button.toggleVisibility"));

        /* グリッドにスナップを切り替えるボタン / Button that toggles snap to grid */
        var snapToGridButton = gridButtonGroup.add("button", undefined, getLabel("button.snapToGrid"));

        /* ステータス表示 / Status line */
        var statusText = palette.add("statictext", undefined, "");
        statusText.alignment = "fill";

        /* 現在の制限角度と定規単位（リセットボタンのディム判定・単位換算に使用）
           / Current constrain angle and ruler unit (used to dim the Reset button and to convert units) */
        var currentConstrain = 0;
        var currentUnit = getUnitByCode(2);

        /* 適用済みの制限角度を状態に記録し、0°ならリセットボタンをディム
           / Record the applied constrain angle in state and dim the Reset button when it is 0° */
        function setConstrain(angle) {
            currentConstrain = angle;
            resetConstrainButton.enabled = (currentConstrain !== 0);
        }

        /* 制限角度を適用して表示・状態を更新する共通処理 / Shared routine that applies a constrain angle and updates the display and state */
        function commitConstrain(angle) {
            applyConstrainAngle(angle, function (result) {
                if (result.indexOf("OK") === 0) {
                    setConstrain(angle);
                    constrainInput.text = roundAngle(angle);
                    statusText.text = getLabel("status.applied");
                } else {
                    statusText.text = statusFromResult(result);
                }
            });
        }

        /* プリセットボタンを1つ作る（ES3にはブロックスコープがないため、クロージャを関数で切り出す）
           / Create one preset button (ES3 has no block scope, so the closure is captured in a function) */
        function addPresetButton(angle) {
            var button = presetGroup.add("button", undefined, angle + "°");
            button.preferredSize.width = PRESET_BUTTON_WIDTH;
            button.onClick = function () {
                commitConstrain(angle);
            };
        }
        for (var i = 0; i < CONSTRAIN_PRESETS.length; i++) {
            addPresetButton(CONSTRAIN_PRESETS[i]);
        }

        /* キー増加(pt)を現在の定規単位の表示文字列に変換 / Convert the keyboard increment (pt) to a display string in the current ruler unit */
        function formatKeyIncrement(lengthPt) {
            return (lengthPt / currentUnit.factor).toFixed(currentUnit.decimals);
        }

        /* キー増加を適用して表示を更新する共通処理（値は現在の定規単位）
           / Shared routine that applies the keyboard increment and updates the display (value in the current ruler unit) */
        function commitKeyIncrement(unitValue) {
            var lengthPt = unitValue * currentUnit.factor;
            applyKeyIncrement(lengthPt, function (result) {
                if (result.indexOf("OK") === 0) {
                    keyIncrementInput.text = formatKeyIncrement(lengthPt);
                    statusText.text = getLabel("status.appliedKeyIncrement");
                } else {
                    statusText.text = statusFromResult(result);
                }
            });
        }

        /* キー増加のプリセットボタン（単位が変わっても作り直さず、表示値だけ差し替える）
           / Keyboard-increment preset buttons (rebuilt values only; the buttons themselves persist across unit changes) */
        var keyIncrementPresetValues = KEY_INCREMENT_PRESETS_DEFAULT;
        var keyIncrementPresetButtons = [];

        /* プリセットボタンを1つ作る（押した時点の keyIncrementPresetValues を参照する）
           / Create one preset button (it reads keyIncrementPresetValues at click time) */
        function addKeyIncrementPresetButton(index) {
            var button = keyIncrementPresetGroup.add("button", undefined, "");
            button.preferredSize.width = PRESET_BUTTON_WIDTH;
            button.onClick = function () {
                commitKeyIncrement(keyIncrementPresetValues[index]);
            };
            keyIncrementPresetButtons.push(button);
        }
        for (var j = 0; j < KEY_INCREMENT_PRESETS_DEFAULT.length; j++) {
            addKeyIncrementPresetButton(j);
        }

        /* 現在の定規単位に合わせてプリセットボタンの値とラベルを差し替え
           / Swap the preset buttons' values and labels to match the current ruler unit */
        function updateKeyIncrementPresets() {
            keyIncrementPresetValues = getKeyIncrementPresets(currentUnit.code);
            for (var n = 0; n < keyIncrementPresetButtons.length; n++) {
                var hasValue = (n < keyIncrementPresetValues.length);
                keyIncrementPresetButtons[n].visible = hasValue;
                if (hasValue) {
                    keyIncrementPresetButtons[n].text = keyIncrementPresetValues[n] + currentUnit.label;
                }
            }
        }

        /* ビュー回転角度・制限角度・定規単位・キー増加を取得し、各表示へ反映
           （制限角度の入力欄にはビューの回転角度を候補値として入れる。適用はボタン押下時のみ）
           / Fetch the view rotation, constrain angle, ruler unit, and keyboard increment, and reflect them in the display
           (the constrain field is seeded with the view rotation; it is applied only on the button) */
        function refresh() {
            fetchState(function (result) {
                if (result.indexOf("OK:") !== 0) {
                    statusText.text = statusFromResult(result);
                    return;
                }
                var parts = result.substring(3).split(",");
                setConstrain(parseFloat(parts[1]));
                currentUnit = getUnitByCode(parseInt(parts[2], 10));
                keyIncrementUnit.text = currentUnit.label;
                keyIncrementInput.text = formatKeyIncrement(parseFloat(parts[3]));
                updateKeyIncrementPresets();
                if (parts[0] === "") {
                    /* ドキュメントがなければ回転を取得できないので、現在の制限角度をそのまま表示
                       / Without a document there is no rotation to read, so show the current constrain angle as-is */
                    constrainInput.text = roundAngle(currentConstrain);
                    statusText.text = getLabel("alert.noDocument");
                } else {
                    constrainInput.text = roundAngle(normalizeAngle(parseFloat(parts[0])));
                    statusText.text = "";
                }
            });
        }

        /* リセット：制限角度を0°に戻して入力欄と状態を更新 / Reset: set the constrain angle to 0° and refresh the field and state */
        resetConstrainButton.onClick = function () {
            resetConstrain(function (result) {
                if (result.indexOf("OK") === 0) {
                    setConstrain(0);
                    constrainInput.text = "0";
                    statusText.text = getLabel("status.resetConstrain");
                } else {
                    statusText.text = statusFromResult(result);
                }
            });
        };

        /* 適用ボタン：入力値を検証してメインエンジンで適用 / Apply button: validate the input and apply in the main engine */
        applyButton.onClick = function () {
            var constrainAngle = parseFloat(constrainInput.text);
            if (isNaN(constrainAngle)) {
                statusText.text = getLabel("alert.invalidAngle");
                return;
            }
            commitConstrain(constrainAngle);
        };

        /* キー増加の適用ボタン：入力値（定規単位）を pt に換算して適用 / Keyboard-increment apply button: convert the entered value (ruler unit) to pt and apply */
        applyKeyIncrementButton.onClick = function () {
            var unitValue = parseFloat(keyIncrementInput.text);
            if (isNaN(unitValue) || unitValue < 0) {
                statusText.text = getLabel("alert.invalidValue");
                return;
            }
            commitKeyIncrement(unitValue);
        };

        /* メニューコマンドをトグルボタンに割り当てる共通処理 / Shared routine that binds a menu command to a toggle button */
        function bindMenuCommand(button, command, statusPath) {
            button.onClick = function () {
                runMenuCommand(command, function (result) {
                    if (result.indexOf("OK") === 0) {
                        statusText.text = getLabel(statusPath);
                    } else {
                        statusText.text = statusFromResult(result);
                    }
                });
            };
        }

        bindMenuCommand(toggleGuideVisibilityButton, "showguide", "status.toggledGuideVisibility");
        bindMenuCommand(toggleGuideLockButton, "lockguide", "status.toggledGuideLock");
        bindMenuCommand(toggleGridVisibilityButton, "showgrid", "status.toggledGridVisibility");
        bindMenuCommand(snapToGridButton, "snapgrid", "status.toggledSnapToGrid");

        /* 初期表示時、およびパレットがアクティブになるたびに最新の状態を取得
           / Fetch the latest state on first show and whenever the palette becomes active */
        palette.onShow = refresh;
        palette.onActivate = refresh;

        palette.show();
    })();

})();
