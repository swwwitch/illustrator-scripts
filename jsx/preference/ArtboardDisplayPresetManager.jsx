#target illustrator
#targetengine "PresetManagerArtboardsPalette"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
 * ArtboardDisplayPresetManager.jsx
 * 更新日: 20260323
 *
 * 概要:
 * Illustrator のアートボード名表示とアートボード枠線の表示設定をまとめて切り替えるスクリプトです。
 * パレット型のコントロールパネルとして動作し、プリセット切り替え、アートボード名の表示、
 * ハイライトカラー、ストローク幅の変更を操作した瞬間に即時反映できます。
 * 下部には補助操作として［ビデオ定規］ボタンも備えています。
 *
 * Summary:
 * A utility script for changing Illustrator artboard display settings with an immediate-apply palette UI.
 * It works as a control panel where preset changes, artboard name visibility, highlight color,
 * and stroke width updates are applied instantly when you interact with the controls.
 * A Video Ruler button is also provided as an auxiliary action.
 *
 * 更新履歴:
 * v1.1 (20260323): ファイル名表記を ArtboardDisplayPresetManager.jsx に更新。
 * v1.0.0 (20260323): 初版。コード構成を constants / utility / UI / 値反映 / event に整理し、
 *                    即時反映のパレットUI、プリセット初期状態判定、ラベル参照の安全化、
 *                    ビデオ定規ボタン、ダイアログタイトルへのバージョン表示、
 *                    画面更新ハックの説明コメントを追加。
 *
 * Changelog:
 * v1.1 (20260323): Updated the internal file name reference to ArtboardDisplayPresetManager.jsx.
 * v1.0.0 (20260323): Initial release. Reorganized the code into constants / utility / UI /
 *                    reflection / event sections, added an immediate-apply palette UI,
 *                    initial preset-state detection, safer label lookup, a Video Ruler button,
 *                    version text in the palette title, and documentation for the zoom refresh hack.
 */

var SCRIPT_VERSION = "v1.1";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */

var LABELS = {
    dialogTitle: {
        ja: "アートボード名と枠線の設定",
        en: "Artboard Name & Border Settings"
    },
    OK: {
        ja: "閉じる",
        en: "Close"
    },
    Cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    VideoRuler: {
        ja: "ビデオ定規",
        en: "Video Ruler"
    },

    // プリセット / Preset
    presetDefault: {
        ja: "デフォルト",
        en: "Default"
    },
    presetEmphasis: {
        ja: "強調",
        en: "Emphasis"
    },
    presetLight: {
        ja: "ライト",
        en: "Light"
    },

    // アートボード / Artboard
    panelArtboardTitle: {
        ja: "アートボード",
        en: "Artboard"
    },
    cbShowArtboardName: {
        ja: "アートボード名を表示",
        en: "Show Artboard Name"
    },
    panelArtboardBorderTitle: {
        ja: "アートボードの枠線",
        en: "Artboard Border"
    },
    artboardStrokeColor: {
        ja: "ハイライトのカラー",
        en: "Highlight Color"
    },
    artboardStrokeWidth: {
        ja: "ストロークの幅",
        en: "Stroke Width"
    },
    artboardColorBlack: {
        ja: "ブラック",
        en: "Black"
    },
    artboardColorLightBlue: {
        ja: "ライトブルー",
        en: "Light Blue"
    },
    artboardColorRed: {
        ja: "サーモンピンク",
        en: "Light Red"
    },
    artboardColorGreen: {
        ja: "グリーン",
        en: "Green"
    },
    artboardColorBlue: {
        ja: "ミディアムブルー",
        en: "Medium Blue"
    },
    artboardColorCyan: {
        ja: "シアン",
        en: "Cyan"
    },
    artboardColorMagenta: {
        ja: "マゼンタ",
        en: "Magenta"
    },
    artboardColorYellow: {
        ja: "イエロー",
        en: "Yellow"
    },
    artboardColorGrey: {
        ja: "ライトグレー",
        en: "Light Gray"
    },
};

function L(key) {
    var entry = LABELS[key];
    if (!entry) return key; // fallback to key name if missing
    if (entry[lang]) return entry[lang];
    // fallback order: ja -> en -> key
    if (entry.ja) return entry.ja;
    if (entry.en) return entry.en;
    return key;
}

function main() {

    var prefs = app.preferences;

    // =========================================
    // Constants / 定数定義
    // =========================================
    var STROKE_COLOR_PRESETS = [
        { key: "LIGHT_BLUE", label: L("artboardColorLightBlue"), r: 0.29, g: 0.52, b: 1.0 },
        { key: "RED", label: L("artboardColorRed"), r: 1.0, g: 0.29, b: 0.29 },
        { key: "GREEN", label: L("artboardColorGreen"), r: 0.0, g: 0.65, b: 0.31 },
        { key: "BLUE", label: L("artboardColorBlue"), r: 0.0, g: 0.45, b: 0.78 },
        { key: "MAGENTA", label: L("artboardColorMagenta"), r: 1.0, g: 0.0, b: 1.0 },
        { key: "CYAN", label: L("artboardColorCyan"), r: 0.0, g: 1.0, b: 1.0 },
        { key: "GREY", label: L("artboardColorGrey"), r: 0.65, g: 0.65, b: 0.65 },
        { key: "BLACK", label: L("artboardColorBlack"), r: 0.0, g: 0.0, b: 0.0 },
        { key: "YELLOW", label: L("artboardColorYellow"), r: 1.0, g: 1.0, b: 0.0 }
    ];
    var STROKE_COLOR_INDEX = {
        LIGHT_BLUE: 0,
        RED: 1,
        GREEN: 2,
        BLUE: 3,
        MAGENTA: 4,
        CYAN: 5,
        GREY: 6,
        BLACK: 7,
        YELLOW: 8
    };

    // =========================================
    // Utility functions / ユーティリティ関数
    // =========================================

    function getReal(key, fb) {
        try { return prefs.getRealPreference(key); } catch (e) { return fb; }
    }

    function getBool(key, fb) {
        try { return prefs.getBooleanPreference(key); } catch (e) { return fb; }
    }

    function clamp(n, min, max) {
        return Math.max(min, Math.min(max, n));
    }


    function buildStrokeColorNames() {
        var names = [];
        for (var i = 0; i < STROKE_COLOR_PRESETS.length; i++) {
            names.push(STROKE_COLOR_PRESETS[i].label);
        }
        return names;
    }

    function findClosestStrokeColor(r, g, b) {
        var bestIdx = 0;
        var bestDist = Infinity;
        for (var i = 0; i < STROKE_COLOR_PRESETS.length; i++) {
            var p = STROKE_COLOR_PRESETS[i];
            var dist = Math.abs(p.r - r) + Math.abs(p.g - g) + Math.abs(p.b - b);
            if (dist < bestDist) {
                bestDist = dist;
                bestIdx = i;
            }
        }
        return bestIdx;
    }

    function getSelectedStrokeWidth() {
        for (var i = 0; i < rbStrokeWidths.length; i++) {
            if (rbStrokeWidths[i].value) return i + 1;
        }
        return 1;
    }

    function applyCurrentSettings() {
        prefs.setBooleanPreference("showArtboardLabelOnCanvas", cbShowArtboardName.value);
        var scIdx = ddStrokeColor.selection ? ddStrokeColor.selection.index : STROKE_COLOR_INDEX.BLACK;
        var scPreset = STROKE_COLOR_PRESETS[scIdx];
        prefs.setRealPreference("ArtboardBBColorRed", scPreset.r);
        prefs.setRealPreference("ArtboardBBColorGreen", scPreset.g);
        prefs.setRealPreference("ArtboardBBColorBlue", scPreset.b);
        prefs.setRealPreference("ArtboardBBWidth", getSelectedStrokeWidth());
        // Force canvas refresh after preference changes.
        // redraw() alone may not update the artboard highlight immediately.
        app.executeMenuCommand('zoomout');
        app.executeMenuCommand('zoomin');
    }

    function applyPreset(preset) {
        var i;
        if (preset === "default") {
            cbShowArtboardName.value = true;
            ddStrokeColor.selection = STROKE_COLOR_INDEX.BLACK;
            for (i = 0; i < rbStrokeWidths.length; i++) rbStrokeWidths[i].value = (i === 0);
        } else if (preset === "emphasis") {
            cbShowArtboardName.value = false;
            ddStrokeColor.selection = STROKE_COLOR_INDEX.RED;
            for (i = 0; i < rbStrokeWidths.length; i++) rbStrokeWidths[i].value = (i === 2);
        } else if (preset === "light") {
            cbShowArtboardName.value = false;
            ddStrokeColor.selection = STROKE_COLOR_INDEX.GREY;
            for (i = 0; i < rbStrokeWidths.length; i++) rbStrokeWidths[i].value = (i === 0);
        }
    }


    /*
  Build main palette / パレット生成
*/
    var dlg = new Window("palette", L("dialogTitle") + " " + SCRIPT_VERSION);
    var mainGroup = dlg.add("group");
    mainGroup.orientation = "column";
    mainGroup.alignChildren = "left";

    /* Preset radio buttons / プリセットラジオボタン */
    var presetRow = mainGroup.add("group");
    presetRow.orientation = "row";
    presetRow.alignment = "center";
    presetRow.margins = [0, 0, 0, 5];
    var rbPresetDefault = presetRow.add("radiobutton", undefined, L("presetDefault"));
    var rbPresetEmphasis = presetRow.add("radiobutton", undefined, L("presetEmphasis"));
    var rbPresetLight = presetRow.add("radiobutton", undefined, L("presetLight"));

    /*
  Artboard panel / ［アートボード］
*/
    var panelArtboard = mainGroup.add("panel", undefined, L("panelArtboardTitle"));
    panelArtboard.orientation = "column";
    panelArtboard.alignChildren = ["fill", "top"];
    panelArtboard.alignment = ["fill", "top"];
    panelArtboard.margins = [15, 20, 15, 10];

    var cbShowArtboardName = panelArtboard.add("checkbox", undefined, L("cbShowArtboardName"));
    cbShowArtboardName.helpTip = LABELS.cbShowArtboardName.ja + " / " + LABELS.cbShowArtboardName.en;

    // Artboard border panel / アートボードの枠線パネル
    var panelArtboardBorder = panelArtboard.add("panel", undefined, L("panelArtboardBorderTitle"));
    panelArtboardBorder.orientation = "column";
    panelArtboardBorder.alignChildren = ["fill", "top"];
    panelArtboardBorder.alignment = ["fill", "top"];
    panelArtboardBorder.margins = [15, 20, 15, 10];

    // Stroke color (dropdown) / ストロークのカラー（ドロップダウン）
    var strokeColorRow = panelArtboardBorder.add("group");
    strokeColorRow.orientation = "row";
    strokeColorRow.alignChildren = ["left", "center"];
    strokeColorRow.add("statictext", undefined, L("artboardStrokeColor") + "：");

    var ddStrokeColor = strokeColorRow.add("dropdownlist", undefined, buildStrokeColorNames());

    // Stroke width (1-4, radio buttons) / ストロークの幅（1〜4、ラジオボタン）
    var strokeWidthRow = panelArtboardBorder.add("group");
    strokeWidthRow.orientation = "row";
    strokeWidthRow.alignChildren = ["left", "center"];
    strokeWidthRow.add("statictext", undefined, L("artboardStrokeWidth") + "：");
    var rbStrokeWidth1 = strokeWidthRow.add("radiobutton", undefined, "1");
    var rbStrokeWidth2 = strokeWidthRow.add("radiobutton", undefined, "2");
    var rbStrokeWidth3 = strokeWidthRow.add("radiobutton", undefined, "3");
    var rbStrokeWidth4 = strokeWidthRow.add("radiobutton", undefined, "4");
    var rbStrokeWidths = [rbStrokeWidth1, rbStrokeWidth2, rbStrokeWidth3, rbStrokeWidth4];

    /* Bottom button row (VideoRuler / Close) / 下部ボタン行 */
    var outerGroup = mainGroup.add("group");
    outerGroup.orientation = "row";
    outerGroup.alignChildren = ["fill", "center"];
    outerGroup.alignment = ["fill", "bottom"];

    var leftGroup = outerGroup.add("group");
    leftGroup.orientation = "row";
    leftGroup.alignChildren = ["left", "center"];
    var btnVideoRuler = leftGroup.add("button", undefined, L("VideoRuler"));

    var spacer = outerGroup.add("group");
    spacer.alignment = ["fill", "fill"];

    var rightGroup = outerGroup.add("group");
    rightGroup.orientation = "row";
    rightGroup.alignChildren = ["right", "center"];
    rightGroup.spacing = 10;
    var btnClose = rightGroup.add("button", undefined, L("OK"), {
        name: "ok"
    });

    // =========================================
    // Reflect current values / 値反映
    // =========================================
    cbShowArtboardName.value = !!getBool("showArtboardLabelOnCanvas", false);
    var curSCR = getReal("ArtboardBBColorRed", 0.0);
    var curSCG = getReal("ArtboardBBColorGreen", 0.0);
    var curSCB = getReal("ArtboardBBColorBlue", 0.0);
    var closestIdx = findClosestStrokeColor(curSCR, curSCG, curSCB);
    ddStrokeColor.selection = closestIdx;
    if (ddStrokeColor.selection === null || ddStrokeColor.selection < 0) {
        ddStrokeColor.selection = STROKE_COLOR_INDEX.BLACK;
    }

    var curStrokeWidth = Math.round(getReal("ArtboardBBWidth", 1.0));
    var swIdx = clamp(curStrokeWidth, 1, 4) - 1;
    rbStrokeWidths[swIdx].value = true;

    // Preset initial state detection / プリセット初期状態の判定
    var isDefault = (
        cbShowArtboardName.value === true &&
        closestIdx === STROKE_COLOR_INDEX.BLACK &&
        curStrokeWidth === 1
    );

    var isEmphasis = (
        cbShowArtboardName.value === false &&
        closestIdx === STROKE_COLOR_INDEX.RED &&
        curStrokeWidth === 3
    );

    var isLight = (
        cbShowArtboardName.value === false &&
        closestIdx === STROKE_COLOR_INDEX.GREY &&
        curStrokeWidth === 1
    );

    if (isDefault) {
        rbPresetDefault.value = true;
    } else if (isEmphasis) {
        rbPresetEmphasis.value = true;
    } else if (isLight) {
        rbPresetLight.value = true;
    }

    // =========================================
    // Event wiring / イベント設定
    // =========================================

    rbPresetDefault.onClick = function () {
        applyPreset("default");
        applyCurrentSettings();
    };
    rbPresetEmphasis.onClick = function () {
        applyPreset("emphasis");
        applyCurrentSettings();
    };
    rbPresetLight.onClick = function () {
        applyPreset("light");
        applyCurrentSettings();
    };

    cbShowArtboardName.onClick = function () {
        applyCurrentSettings();
    };

    ddStrokeColor.onChange = function () {
        applyCurrentSettings();
    };

    rbStrokeWidth1.onClick = function () {
        applyCurrentSettings();
    };
    rbStrokeWidth2.onClick = function () {
        applyCurrentSettings();
    };
    rbStrokeWidth3.onClick = function () {
        applyCurrentSettings();
    };
    rbStrokeWidth4.onClick = function () {
        applyCurrentSettings();
    };


    btnVideoRuler.onClick = function () {
        app.executeMenuCommand('videoruler');
    };


    btnClose.onClick = function () {
        dlg.close();
    };

    dlg.center();
    dlg.show();
}

main();