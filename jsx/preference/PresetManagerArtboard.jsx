#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アートボード名表示とアートボード枠線の表示設定を、ダイアログでまとめて切り替えます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PresetManagerArtboard.md

### Overview

Switches the artboard-name and artboard-border display preferences from a dialog.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PresetManagerArtboard.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PresetManagerArtboard";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-23";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PresetManagerArtboard.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PresetManagerArtboard.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

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

    /**
     * ラベルを現在のUI言語で取得 / Return a label in the current UI language
     * @param {string} key - LABELS のキー
     * @returns {string} ラベル文字列。見つからないときはキー名を返す
     */
    function getLabel(key) {
        var entry = LABELS[key];
        if (!entry) return key;
        if (entry[uiLang]) return entry[uiLang];
        /* ja -> en -> キー名の順でフォールバック / Fall back ja -> en -> key name */
        if (entry.ja) return entry.ja;
        if (entry.en) return entry.en;
        return key;
    }

    /**
     * 項目名にコロンを付ける（日本語は全角、英語は半角）/ Append a colon; full-width in Japanese, half-width in English
     * @param {string} key - LABELS のキー
     * @returns {string} コロン付きのラベル文字列
     */
    function labelText(key) {
        return getLabel(key) + (uiLang === "ja" ? "：" : ": ");
    }

    // =========================================
    // UIレイアウトの共通設定 / Shared UI layout
    // =========================================

    var PANEL_MARGINS = [15, 20, 15, 10]; /* パネル余白 [左,上,右,下] / panel margins */
    var PRESET_ROW_BOTTOM_MARGIN = 5;     /* プリセット行の下余白 / preset row bottom margin */
    var BUTTON_ROW_TOP_MARGIN = 10;       /* ボタン行の上余白 / button row top margin */
    var BUTTON_SPACING = 10;              /* ボタン同士の間隔 / spacing between buttons */

    function main() {

        var prefs = app.preferences;

        // =========================================
        // Constants / 定数定義
        // =========================================
        var STROKE_COLOR_PRESETS = [
            { key: "LIGHT_BLUE", label: getLabel("artboardColorLightBlue"), r: 0.29, g: 0.52, b: 1.0 },
            { key: "RED", label: getLabel("artboardColorRed"), r: 1.0, g: 0.29, b: 0.29 },
            { key: "GREEN", label: getLabel("artboardColorGreen"), r: 0.0, g: 0.65, b: 0.31 },
            { key: "BLUE", label: getLabel("artboardColorBlue"), r: 0.0, g: 0.45, b: 0.78 },
            { key: "MAGENTA", label: getLabel("artboardColorMagenta"), r: 1.0, g: 0.0, b: 1.0 },
            { key: "CYAN", label: getLabel("artboardColorCyan"), r: 0.0, g: 1.0, b: 1.0 },
            { key: "GREY", label: getLabel("artboardColorGrey"), r: 0.65, g: 0.65, b: 0.65 },
            { key: "BLACK", label: getLabel("artboardColorBlack"), r: 0.0, g: 0.0, b: 0.0 },
            { key: "YELLOW", label: getLabel("artboardColorYellow"), r: 1.0, g: 1.0, b: 0.0 }
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
            forceScreenRefresh();
        }

        /**
         * 環境設定の変更後に画面を強制再描画（redraw だけではハイライトが更新されないため）
         * Force a redraw after preference changes (redraw alone leaves the highlight stale)
         * ドキュメントが開いていないときは何もしない / Does nothing when no document is open
         * @returns {void}
         */
        function forceScreenRefresh() {
            if (app.documents.length === 0) return;
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
      Build main dialog / ダイアログ生成
    */
        var dlg = new Window("dialog", getLabel("dialogTitle") + " " + SCRIPT_VERSION);
        var mainGroup = dlg.add("group");
        mainGroup.orientation = "column";
        mainGroup.alignChildren = "left";

        /* Preset radio buttons / プリセットラジオボタン */
        var presetRow = mainGroup.add("group");
        presetRow.orientation = "row";
        presetRow.alignment = "center";
        presetRow.margins = [0, 0, 0, PRESET_ROW_BOTTOM_MARGIN];
        var rbPresetDefault = presetRow.add("radiobutton", undefined, getLabel("presetDefault"));
        var rbPresetEmphasis = presetRow.add("radiobutton", undefined, getLabel("presetEmphasis"));
        var rbPresetLight = presetRow.add("radiobutton", undefined, getLabel("presetLight"));

        /*
      Artboard panel / ［アートボード］
    */
        var panelArtboard = mainGroup.add("panel", undefined, getLabel("panelArtboardTitle"));
        panelArtboard.orientation = "column";
        panelArtboard.alignChildren = ["fill", "top"];
        panelArtboard.alignment = ["fill", "top"];
        panelArtboard.margins = PANEL_MARGINS;

        var cbShowArtboardName = panelArtboard.add("checkbox", undefined, getLabel("cbShowArtboardName"));
        cbShowArtboardName.helpTip = LABELS.cbShowArtboardName.ja + " / " + LABELS.cbShowArtboardName.en;

        // Artboard border panel / アートボードの枠線パネル
        var panelArtboardBorder = panelArtboard.add("panel", undefined, getLabel("panelArtboardBorderTitle"));
        panelArtboardBorder.orientation = "column";
        panelArtboardBorder.alignChildren = ["fill", "top"];
        panelArtboardBorder.alignment = ["fill", "top"];
        panelArtboardBorder.margins = PANEL_MARGINS;

        // Stroke color (dropdown) / ストロークのカラー（ドロップダウン）
        var strokeColorRow = panelArtboardBorder.add("group");
        strokeColorRow.orientation = "row";
        strokeColorRow.alignChildren = ["left", "center"];
        strokeColorRow.add("statictext", undefined, labelText("artboardStrokeColor"));

        var ddStrokeColor = strokeColorRow.add("dropdownlist", undefined, buildStrokeColorNames());

        // Stroke width (1-4, radio buttons) / ストロークの幅（1〜4、ラジオボタン）
        var strokeWidthRow = panelArtboardBorder.add("group");
        strokeWidthRow.orientation = "row";
        strokeWidthRow.alignChildren = ["left", "center"];
        strokeWidthRow.add("statictext", undefined, labelText("artboardStrokeWidth"));
        var rbStrokeWidth1 = strokeWidthRow.add("radiobutton", undefined, "1");
        var rbStrokeWidth2 = strokeWidthRow.add("radiobutton", undefined, "2");
        var rbStrokeWidth3 = strokeWidthRow.add("radiobutton", undefined, "3");
        var rbStrokeWidth4 = strokeWidthRow.add("radiobutton", undefined, "4");
        var rbStrokeWidths = [rbStrokeWidth1, rbStrokeWidth2, rbStrokeWidth3, rbStrokeWidth4];

        /* 下部ボタン行（左：ビデオ定規／右：閉じる）/ Bottom button row (left: Video Ruler, right: Close) */
        var btnRowGroup = mainGroup.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.alignment = ["fill", "bottom"];

        /* 左側グループ / Left-side button group */
        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        var btnVideoRuler = btnLeftGroup.add("button", undefined, getLabel("VideoRuler"));

        /* スペーサー（伸縮）/ Spacer (stretchable) */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        /* 右側グループ / Right-side button group */
        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.spacing = BUTTON_SPACING;
        var btnOK = btnRightGroup.add("button", undefined, getLabel("OK"), { name: "ok" });

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

        /* 現在値がいずれかのプリセットと一致していればそのラジオを選ぶ / Select the preset radio that matches the current values */
        if (cbShowArtboardName.value === true && closestIdx === STROKE_COLOR_INDEX.BLACK && curStrokeWidth === 1) {
            rbPresetDefault.value = true;
        } else if (cbShowArtboardName.value === false && closestIdx === STROKE_COLOR_INDEX.RED && curStrokeWidth === 3) {
            rbPresetEmphasis.value = true;
        } else if (cbShowArtboardName.value === false && closestIdx === STROKE_COLOR_INDEX.GREY && curStrokeWidth === 1) {
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

        btnOK.onClick = function () {
            applyCurrentSettings();
            dlg.close();
        };

        dlg.center();
        dlg.show();
    }

    main();

})();
