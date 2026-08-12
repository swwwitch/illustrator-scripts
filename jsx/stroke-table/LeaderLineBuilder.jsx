#target illustrator
#targetengine "LeaderLineBuilderEngine"
#include "ColorPicker.jsx"

app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したパスやグループの外接矩形から、指定した角度の引き出し線を作成し、元のオブジェクトと置き換えます。
角度・斜線の方向・線のスタイル・先端マーカー・フチを、ダイアログでプレビューしながら調整できます。

詳細はREADMEを参照。

*/

/*

### Overview

Builds leader lines at a specified angle from the bounding box of the selected
paths or groups, replacing the original objects. Angle, diagonal direction, line
style, tip marker, and edge are adjusted with a live preview in the dialog.

See the README for details.

*/

(function () {

    // =========================================
    // 基本情報 / Basic info
    // =========================================
    var SCRIPT_NAME     = "LeaderLineBuilder";            /* スクリプト名 / script name */
    var SCRIPT_VERSION  = "v1.5.2";                       /* バージョン / version */
    var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
    var SCRIPT_RELEASED = "2026-03-06";                   /* 最初のリリース日 / first release date */
    var SCRIPT_UPDATED  = "2026-08-12";                   /* 更新日 / last updated */

    // README (Japanese)
    // https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/LeaderLineBuilder.md
    // README (English)
    // https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/LeaderLineBuilder.md
    var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n506df641d5c5"; /* 紹介記事 / article URL */

    // Released under the MIT license
    // http://opensource.org/licenses/mit-license.php

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 起動時の値。入力値が不正だったときのフォールバックにも使う */
    var DEFAULT_ANGLE         = 45;         /* 角度（度） */
    var DEFAULT_LINE_WIDTH_PT = 1;          /* 線幅（pt） */
    var DEFAULT_TIP_MARKER_SIZE_PT   = 3;          /* 線端の大きさ（pt） */
    var DEFAULT_TEXT_DIST_PT  = 72 / 25.4;  /* テキストとの距離（pt。1mm相当） */

    /* フチの太らせ方（いずれも線幅に対する倍率） */
    var EDGE_LINE_WIDTH_RATIO = 3;  /* 線のフチの線幅 */
    var EDGE_CIRCLE_EXPAND_RATIO    = 2;  /* 丸のフチの広がり */
    var EDGE_ARROW_EXPAND_RATIO  = 3;  /* 矢印のフチの広がり */

    // =========================================
    // レイアウト / Layout
    // =========================================
    var WINDOW_MARGINS    = 16;                /* ウィンドウ外周の余白 */
    var WINDOW_SPACING    = 12;                /* ウィンドウ内の要素間隔 */
    var PANEL_MARGINS     = [16, 20, 16, 12];  /* パネル余白 [左,上,右,下] */
    var PANEL_SPACING     = 6;                 /* パネル内の要素間隔 */
    var COLUMN_SPACING    = 12;                /* 2カラムの間隔 */
    var COLOR_CHIP_SIZE       = [20, 20];          /* カラースウォッチの大きさ */
    var ZOOM_SLIDER_WIDTH = 240;               /* ズームスライダーの幅 */
    var ZOOM_MIN          = 0.1;               /* ズームの下限（倍率） */
    var ZOOM_MAX          = 8;                 /* ズームの上限（倍率） */

    /**
     * ダイアログウィンドウに共通のレイアウトを適用する
     * @param {Window} win - 対象のウィンドウ
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(win, spacing) {
        win.orientation = "column";
        win.alignChildren = ["fill", "top"];
        win.margins = WINDOW_MARGINS;
        win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルに共通のレイアウトを適用する
     * @param {Panel} panel - 対象のパネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * 横並びグループに共通のレイアウトを適用する
     * @param {Group} group - 対象のグループ
     * @param {Array<string>} [alignment] - 子要素の整列（省略時は ["left","center"]）
     * @param {number} [spacing] - 要素間隔
     * @returns {void}
     */
    function setupRowGroup(group, alignment, spacing) {
        group.orientation = "row";
        group.alignChildren = alignment || ["left", "center"];
        if (typeof spacing === "number") group.spacing = spacing;
    }

    /**
     * 共通レイアウトを適用したラベル付きパネルを追加する
     * @param {Group|Panel|Window} parent - 追加先
     * @param {string} labelText - パネルのラベル
     * @param {number} [spacing] - パネル内の要素間隔
     * @returns {Panel} 追加したパネル
     */
    function addLabeledPanel(parent, labelText, spacing) {
        var panel = parent.add("panel", undefined, labelText);
        setupPanel(panel, spacing);
        return panel;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 設定キーとラジオボタンの対応から、指定キーだけを選択状態にする
     * @param {Array<object>} radioMap - key / radio を持つ要素の配列
     * @param {string} key - 選択するキー
     * @param {string} fallbackKey - key が対応表にないときに選択するキー
     * @returns {void}
     */
    function selectRadioByKey(radioMap, key, fallbackKey) {
        var i;
        var matched = false;
        for (i = 0; i < radioMap.length; i++) {
            if (radioMap[i].key === key) matched = true;
        }
        if (!matched) key = fallbackKey;
        for (i = 0; i < radioMap.length; i++) {
            radioMap[i].radio.value = (radioMap[i].key === key);
        }
    }

    /**
     * 設定キーとラジオボタンの対応から、選択中のキーを取得する
     * @param {Array<object>} radioMap - key / radio を持つ要素の配列
     * @param {string} fallbackKey - どれも選択されていないときに返すキー
     * @returns {string} 選択中のキー
     */
    function getSelectedRadioKey(radioMap, fallbackKey) {
        for (var i = 0; i < radioMap.length; i++) {
            if (radioMap[i].radio.value) return radioMap[i].key;
        }
        return fallbackKey;
    }

    /**
     * 現在のUI言語を取得する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        var locale = $.locale || "";
        return (locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "引き出し線ビルダー", en: "Leader Line Builder" }
        },
        panel: {
            applyScope: { ja: "適用範囲", en: "Apply Scope" },
            angle:      { ja: "角度", en: "Angle" },
            direction:  { ja: "斜線の方向", en: "Diagonal Direction" },
            lineStyle:  { ja: "線のスタイル", en: "Line Style" },
            tipMarker:  { ja: "先端マーカー", en: "Tip Marker" },
            edge:       { ja: "フチ", en: "Edge" },
            text:       { ja: "テキスト", en: "Text" }
        },
        radio: {
            applyScopeAll:          { ja: "すべて更新", en: "Update all" },
            applyScopeKeepDir:      { ja: "方向は維持", en: "Keep each direction" },
            diagDirUpperLeft:       { ja: "左上", en: "Upper Left" },
            diagDirLowerLeft:       { ja: "左下", en: "Lower Left" },
            diagDirUpperRight:      { ja: "右上", en: "Upper Right" },
            diagDirLowerRight:      { ja: "右下", en: "Lower Right" },
            lineColorBlack:         { ja: "黒", en: "Black" },
            lineColorWhite:         { ja: "白", en: "White" },
            colorCustom:            { ja: "指定", en: "Custom" },
            strokeCapButt:          { ja: "なし", en: "None" },
            strokeCapRound:         { ja: "丸型", en: "Round" },
            tipMarkerNone:          { ja: "なし", en: "None" },
            tipMarkerCircle:        { ja: "円", en: "Circle" },
            tipMarkerArrow:         { ja: "矢印", en: "Arrow" },
            tipMarkerFill:          { ja: "塗り", en: "Filled" },
            tipMarkerOutline:       { ja: "線のみ", en: "Outlined" }
        },
        checkbox: {
            groupEnabled: { ja: "グループ化", en: "Group items" },
            edgeEnabled:  { ja: "フチを付ける", en: "Add edge" },
            lightMode:    { ja: "簡易", en: "Defer redraw" }
        },
        fieldLabel: {
            lineWidth:     { ja: "線幅", en: "Line Width" },
            strokeCap:     { ja: "線端", en: "Line End" },
            tipMarkerSize: { ja: "大きさ", en: "Size" },
            textDist:      { ja: "テキストとの距離", en: "Distance to text" },
            zoom:          { ja: "ズーム", en: "Zoom" },
            targetMark:    { ja: "基準", en: "Target" }
        },
        tooltip: {
            applyScopeAll: {
                ja: "ダイアログの設定をすべて適用します。",
                en: "Apply every setting in this dialog."
            },
            applyScopeKeepDir: {
                ja: "方向だけは、選択中の引き出し線それぞれに保存された向きを使います。",
                en: "Keep the direction stored in each selected leader line."
            },
            angleInput: {
                ja: "0より大きく90未満。↑↓で1、Shift+↑↓で10、Option+↑↓で0.1ずつ変わります。",
                en: "Greater than 0 and less than 90. Arrow keys step by 1, Shift by 10, Option by 0.1."
            },
            lengthInput: {
                ja: "↑↓で0.1、Shift+↑↓で1、Option+↑↓で0.01ずつ変わります。単位は環境設定に従います。",
                en: "Arrow keys step by 0.1, Shift by 1, Option by 0.01. The unit follows your preferences."
            },
            diagDir: {
                ja: "引き出し線が伸びる向き。左上なら、対象の左上へ引き出します。",
                en: "Where the leader line runs. Upper Left draws toward the upper left of the object."
            },
            colorChip: {
                ja: "クリックするとカラーピッカーが開きます。",
                en: "Click to open the color picker."
            },
            strokeCap: {
                ja: "線そのものの端の形です。先端マーカーとは別の設定です。",
                en: "The shape of the stroke ends. This is separate from the tip marker."
            },
            tipMarkerStyle: {
                ja: "矢印は塗りのみです。",
                en: "Arrows are always filled."
            },
            tipMarkerSize: {
                ja: "円なら直径、矢印なら長さです。",
                en: "Diameter for a circle, length for an arrow."
            },
            groupEnabled: {
                ja: "線・先端マーカー・フチを1つのグループにまとめます。あとから設定を変えるにはグループ化が必要です。",
                en: "Group the line, tip marker, and edge together. Grouping is required to re-apply settings later."
            },
            edgeEnabled: {
                ja: "線と先端マーカーの背面に、一回り太らせた同じ形を敷きます。",
                en: "Place a thicker copy of the line and tip marker behind them."
            },
            textDist: {
                ja: "テキストを一緒に選択したときだけ有効です。",
                en: "Available only when a text frame is also selected."
            },
            lightMode: {
                ja: "ドラッグ中は再描画せず、離した時点で反映します。",
                en: "Skip redrawing while dragging; apply when the slider is released."
            }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok:     { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: {
                ja: "パスまたはグループを選択して実行してください。",
                en: "Please select a path or group and run the script."
            },
            noValidTargets: {
                ja: "引き出し線の基準にできるオブジェクトがありません。パス（2点以上）またはグループを選択してください。",
                en: "Nothing can be used as a leader line reference. Select a path with 2 or more points, or a group."
            },
            invalidAngle: {
                ja: "角度は0より大きく90未満の値を入力してください。",
                en: "Enter an angle greater than 0 and less than 90."
            }
        }
    };

    /**
     * LABELS を上から順にたどって現在のUI言語のラベルを返す
     * @param {...string} - LABELS をたどるキー
     * @returns {string} ラベル文字列（見つからない場合は空文字）
     */
    function getLabel() {
        var node = LABELS;
        for (var i = 0; i < arguments.length; i++) {
            if (node == null) break;
            node = node[arguments[i]];
        }
        return (node && node[uiLang] != null) ? node[uiLang] : "";
    }

    // =========================================
    // セッション記憶 / Session state
    // =========================================

    /* #targetengine が生きている間だけダイアログの設定を保持する */
    if (typeof $.global._leaderLineSettings === "undefined") {
        $.global._leaderLineSettings = {
            angle: DEFAULT_ANGLE,
            radioAngle: 45,
            applyScope: "all",
            hasUserSetApplyScope: false,
            diagDir: "upperLeft",
            hDir: "left",
            vDir: "up",
            tipMarkerType: "circle",
            tipMarkerStyle: "fill",
            tipMarkerSize: DEFAULT_TIP_MARKER_SIZE_PT,
            strokeCapType: "round",
            groupEnabled: true,
            whiteEdge: false,
            edgeColor: "white",
            edgeColorHex: "#ffcc00",
            dialogBounds: null, /* ダイアログ位置 [x, y] のみをセッション中に保存 */
            lineColor: "black",
            lineColorHex: "#ffcc00",
            lineWidth: DEFAULT_LINE_WIDTH_PT,
            textDist: DEFAULT_TEXT_DIST_PT
        };
    }
    var sessionSettings = $.global._leaderLineSettings;

    // =========================================
    // 単位ユーティリティ / Unit utilities
    // =========================================

    /* 単位コードと単位ラベルの対応 / Unit code to label */
    var UNIT_LABELS = {
        0: "in",
        1: "mm",
        2: "pt",
        3: "pica",
        4: "cm",
        6: "px",
        7: "ft/in",
        8: "m",
        9: "yd",
        10: "ft"
    };

    /**
     * 単位コードと設定キーから単位ラベルを返す
     * @param {number} code - 単位コード
     * @param {string} prefKey - 環境設定のキー
     * @returns {string} 単位ラベル
     */
    function getUnitLabel(code, prefKey) {
        if (code === 5) {
            /* 級数と歯は設定キーで使い分ける / Q for size, H for distance */
            var heightUnitKeys = {
                "text/asianunits": true,
                "rulerType": true,
                "strokeUnits": true
            };
            return heightUnitKeys[prefKey] ? "H" : "Q";
        }
        return UNIT_LABELS[code] || "pt";
    }

    /**
     * 設定キーから現在の単位ラベルを取得する
     * @param {string} prefKey - 環境設定のキー
     * @returns {string} 単位ラベル
     */
    function getCurrentUnitLabel(prefKey) {
        var unitCode = app.preferences.getIntegerPreference(prefKey);
        return getUnitLabel(unitCode, prefKey);
    }

    /**
     * 単位コードからpt換算係数を取得する
     * @param {number} code - 単位コード
     * @returns {number} pt換算係数
     */
    function getUnitToPtFactor(code) {
        switch (code) {
            case 0: return 72;          /* in */
            case 1: return 72 / 25.4;   /* mm */
            case 2: return 1;           /* pt */
            case 3: return 12;          /* pica */
            case 4: return 72 / 2.54;   /* cm */
            case 5: return 0.25;        /* Q/H */
            case 6: return 1;           /* px */
            case 7: return 72;          /* ft/in（数値入力はinch扱い） */
            case 8: return 72 / 0.0254; /* m */
            case 9: return 72 * 36;     /* yd */
            case 10: return 72 * 12;    /* ft */
            default: return 1;
        }
    }

    /**
     * 設定キーから現在のpt換算係数を取得する
     * @param {string} prefKey - 環境設定のキー
     * @returns {number} pt換算係数
     */
    function getCurrentUnitToPtFactor(prefKey) {
        return getUnitToPtFactor(app.preferences.getIntegerPreference(prefKey));
    }

    /**
     * 単位値をptに変換する
     * @param {string|number} value - 単位付きの数値
     * @param {string} prefKey - 環境設定のキー
     * @returns {number} pt値（数値でない場合は NaN）
     */
    function unitValueToPt(value, prefKey) {
        var parsed = parseFloat(value);
        if (isNaN(parsed)) return NaN;
        return parsed * getCurrentUnitToPtFactor(prefKey);
    }

    /**
     * ptを単位値に変換する
     * @param {string|number} valuePt - pt値
     * @param {string} prefKey - 環境設定のキー
     * @returns {number} 単位値（数値でない場合は NaN）
     */
    function ptValueToUnit(valuePt, prefKey) {
        var parsed = parseFloat(valuePt);
        if (isNaN(parsed)) return NaN;
        return parsed / getCurrentUnitToPtFactor(prefKey);
    }

    /**
     * 指定桁で四捨五入する
     * @param {number} value - 対象の値
     * @param {number} [digits] - 小数桁数
     * @returns {number} 丸めた値
     */
    function roundTo(value, digits) {
        var factor = Math.pow(10, digits || 0);
        return Math.round(value * factor) / factor;
    }

    /**
     * 数値を表示用の文字列に整形する
     * @param {number} value - 対象の値
     * @param {number} [digits] - 小数桁数
     * @returns {string} 整形した文字列（数値でない場合は空文字）
     */
    function formatNumber(value, digits) {
        if (isNaN(value)) return "";
        var text = String(roundTo(value, digits || 0));
        text = text.replace(/(\.\d*?)0+$/, "$1");
        /* digits>=1 のとき最低1桁の小数を保持（1 → 1.0） */
        if ((digits || 0) >= 1 && text.indexOf(".") === -1) text += ".0";
        return text;
    }

    /**
     * pt値を現在の単位の表示用文字列に整形する
     * @param {string|number} valuePt - pt値
     * @param {string} prefKey - 環境設定のキー
     * @returns {string} 整形した文字列
     */
    function formatPtInCurrentUnit(valuePt, prefKey) {
        return formatNumber(ptValueToUnit(valuePt, prefKey), 2);
    }

    /**
     * pt値を線幅系の入力欄に表示する文字列に整形する
     * @param {string|number} valuePt - pt値
     * @param {number} fallbackValue - 値が不正だったときに表示する値
     * @returns {string} 現在の単位に換算した文字列
     */
    function formatUnitInput(valuePt, fallbackValue) {
        var displayValue = parseFloat(formatPtInCurrentUnit(valuePt, "strokeUnits"));
        if (isNaN(displayValue) || displayValue <= 0) displayValue = fallbackValue;
        return formatNumber(displayValue, 2);
    }

    /**
     * 線幅系の入力欄の値を、セッション記憶用のpt値に変換する
     * @param {string} unitText - 単位付きの数値
     * @param {number} fallbackPt - 値が不正だったときに使うpt値
     * @returns {string|number} pt値
     */
    function parseUnitInput(unitText, fallbackPt) {
        var valuePt = unitValueToPt(unitText, "strokeUnits");
        return (!isNaN(valuePt) && valuePt > 0) ? formatNumber(valuePt, 4) : fallbackPt;
    }

    // =========================================
    // ズーム操作 / View zoom
    // =========================================

    /**
     * 現在のビューの表示倍率と中心を控える
     * @param {Document} doc - 対象ドキュメント
     * @returns {object} ビュー・倍率・中心を持つ状態オブジェクト
     */
    function captureViewState(doc) {
        var state = { view: null, zoom: null, center: null };
        try {
            state.view = doc.activeView;
            state.zoom = state.view.zoom;
            state.center = state.view.centerPoint;
        } catch (e) { }
        return state;
    }

    /**
     * 控えておいたビューの表示倍率と中心を戻す
     * @param {Document} doc - 対象ドキュメント
     * @param {object} state - captureViewState() の戻り値
     * @returns {void}
     */
    function restoreViewState(doc, state) {
        if (!state) return;
        try {
            var view = state.view || doc.activeView;
            if (view && state.zoom != null) view.zoom = state.zoom;
            if (view && state.center != null) view.centerPoint = state.center;
        } catch (e) { }
    }

    /**
     * ズームスライダーと「軽」チェックボックスをダイアログに追加する
     * 「軽」がONのときはドラッグ中に反映せず、離した時点でだけ倍率を適用する
     * @param {Window|Group|Panel} parent - 追加先
     * @param {Document} doc - 対象ドキュメント
     * @param {string} labelText - スライダーのラベル
     * @param {object} initialState - captureViewState() の戻り値
     * @param {object} [options] - 表示と挙動のオプション
     * @returns {object} スライダーなどの参照と操作関数をまとめたオブジェクト
     */
    function addZoomControls(parent, doc, labelText, initialState, options) {
        options = options || {};
        var minZoom = (typeof options.min === "number") ? options.min : ZOOM_MIN;
        var maxZoom = (typeof options.max === "number") ? options.max : ZOOM_MAX;
        var sliderWidth = (typeof options.sliderWidth === "number") ? options.sliderWidth : ZOOM_SLIDER_WIDTH;
        var doRedraw = (options.redraw !== false);
        var showLightMode = (options.lightMode !== false);
        var lightModeLabel = options.lightModeLabel || "Light mode";
        var lightModeDefault = (options.lightModeDefault === true);

        var zoomGroup = parent.add("group");
        setupRowGroup(zoomGroup, ["center", "center"]);
        zoomGroup.alignment = "center";
        if (options.margins) zoomGroup.margins = options.margins;

        zoomGroup.add("statictext", undefined, String(labelText || "Zoom"));

        var initialZoom = 1;
        try {
            if (initialState && initialState.zoom != null) initialZoom = Number(initialState.zoom);
            else initialZoom = Number(doc.activeView.zoom);
        } catch (e) { }
        if (!initialZoom || isNaN(initialZoom)) initialZoom = 1;

        var zoomSlider = zoomGroup.add("slider", undefined, initialZoom, minZoom, maxZoom);
        zoomSlider.preferredSize.width = sliderWidth;

        var lightModeCheckbox = null;
        if (showLightMode) {
            lightModeCheckbox = zoomGroup.add("checkbox", undefined, String(lightModeLabel));
            lightModeCheckbox.value = lightModeDefault;
            if (options.lightModeTip) lightModeCheckbox.helpTip = String(options.lightModeTip);
        }

        /**
         * 「軽」がONかどうかを返す
         * @returns {boolean} ONなら true
         */
        function isLightMode() {
            return !!(lightModeCheckbox && lightModeCheckbox.value);
        }

        /**
         * 表示倍率を適用する
         * @param {number} zoom - 表示倍率
         * @returns {void}
         */
        function applyZoom(zoom) {
            try {
                var view = (initialState && initialState.view) ? initialState.view : doc.activeView;
                if (!view) return;
                view.zoom = zoom;
                if (doRedraw) app.redraw();
            } catch (e) { }
        }

        zoomSlider.onChanging = function () {
            if (isLightMode()) return;
            applyZoom(Number(zoomSlider.value));
        };

        zoomSlider.onChange = function () {
            applyZoom(Number(zoomSlider.value));
        };

        if (lightModeCheckbox) {
            lightModeCheckbox.onClick = function () {
                applyZoom(Number(zoomSlider.value));
            };
        }

        return {
            slider: zoomSlider,
            lightModeCheckbox: lightModeCheckbox,
            applyZoom: applyZoom,
            restoreInitial: function () { restoreViewState(doc, initialState); }
        };
    }

    // =========================================
    // 引き出し線の方向 / Leader line direction
    // =========================================

    /* 斜線の方向名と、水平・垂直の向きの対応 */
    var DIAG_DIRECTIONS = [
        { name: "upperLeft",  hDir: "right", vDir: "down" },
        { name: "lowerLeft",  hDir: "right", vDir: "up" },
        { name: "upperRight", hDir: "left",  vDir: "down" },
        { name: "lowerRight", hDir: "left",  vDir: "up" }
    ];

    /**
     * 斜線の方向名から水平・垂直の向きを求める
     * @param {string} diagDir - 方向名（"upperLeft" など）
     * @returns {object} hDir / vDir を持つオブジェクト
     */
    function hvDirFromDiagDir(diagDir) {
        for (var i = 0; i < DIAG_DIRECTIONS.length; i++) {
            if (DIAG_DIRECTIONS[i].name === diagDir) {
                return { hDir: DIAG_DIRECTIONS[i].hDir, vDir: DIAG_DIRECTIONS[i].vDir };
            }
        }
        return { hDir: "right", vDir: "up" };
    }

    /**
     * 水平・垂直の向きから斜線の方向名を求める
     * @param {string} hDir - 水平方向（"left" / "right"）
     * @param {string} vDir - 垂直方向（"up" / "down"）
     * @returns {string|null} 方向名（該当しない場合は null）
     */
    function diagDirFromHVDir(hDir, vDir) {
        for (var i = 0; i < DIAG_DIRECTIONS.length; i++) {
            if (DIAG_DIRECTIONS[i].hDir === hDir && DIAG_DIRECTIONS[i].vDir === vDir) {
                return DIAG_DIRECTIONS[i].name;
            }
        }
        return null;
    }

    /**
     * 斜線の方向ラジオから水平・垂直の向きを求める
     * @param {object} ui - buildDialogUI() の戻り値
     * @returns {object} hDir / vDir を持つオブジェクト
     */
    function getDiagDirValues(ui) {
        return hvDirFromDiagDir(getSelectedRadioKey(ui.radioMaps.diagDir, "lowerLeft"));
    }

    /**
     * 引き出し線を1つだけ選んでいるとき、保存済みの方向をセッション記憶へ戻す
     * @param {Array<PathItem|GroupItem>} targetItems - 対象アイテム
     * @returns {void}
     */
    function restoreDirFromSelection(targetItems) {
        if (targetItems.length !== 1) return;
        var target = targetItems[0];
        if (!isLeaderLineGroup(target)) return;
        var storedDir = getLeaderLineStoredDir(target);
        if (!storedDir) return;
        var diagDir = diagDirFromHVDir(storedDir.hDir, storedDir.vDir);
        if (!diagDir) return;
        sessionSettings.diagDir = diagDir;
        sessionSettings.hDir = storedDir.hDir;
        sessionSettings.vDir = storedDir.vDir;
    }

    // =========================================
    // アイテム操作 / Item helpers
    // =========================================

    /**
     * アイテムを安全に削除する
     * @param {object} item - 対象アイテム
     * @returns {void}
     */
    function safeRemove(item) {
        if (!item) return;
        try {
            if (item.parent) item.remove();
        } catch (e) { }
    }

    /**
     * アイテムを安全に選択する
     * @param {object} item - 対象アイテム
     * @returns {void}
     */
    function safeSelect(item) {
        if (!item) return;
        try {
            item.selected = true;
        } catch (e) { }
    }

    /**
     * アイテムのメモ（note）を安全に設定する
     * @param {object} item - 対象アイテム
     * @param {string} note - 設定する文字列
     * @returns {void}
     */
    function safeSetNote(item, note) {
        if (!item) return;
        item.note = note;
    }

    /**
     * 編集可能なレイヤーをアクティブにする
     * 対象アイテムのレイヤーが使えればそこへ、なければ最初の編集可能レイヤーへ切り替える
     * @param {Document} doc - 対象ドキュメント
     * @param {object} item - 対象アイテム
     * @returns {void}
     */
    function switchToEditableLayer(doc, item) {
        try {
            var itemLayer = item.layer;
            if (itemLayer && !itemLayer.locked && itemLayer.visible) {
                doc.activeLayer = itemLayer;
                return;
            }
        } catch (e) { }
        for (var i = 0; i < doc.layers.length; i++) {
            var layer = doc.layers[i];
            if (!layer.locked && layer.visible) {
                doc.activeLayer = layer;
                return;
            }
        }
    }

    // =========================================
    // メモとタグ / Note and tag
    // =========================================

    /**
     * アイテムのタグに値を設定する（同名タグがあれば上書き）
     * @param {object} item - 対象アイテム
     * @param {string} name - タグ名
     * @param {string|number} value - 設定する値
     * @returns {void}
     */
    function setTagValue(item, name, value) {
        if (!item) return;
        try {
            for (var i = 0; i < item.tags.length; i++) {
                if (item.tags[i].name === name) {
                    item.tags[i].value = String(value);
                    return;
                }
            }
            var tag = item.tags.add();
            tag.name = name;
            tag.value = String(value);
        } catch (e) { }
    }

    /**
     * アイテムのタグから値を取得する
     * @param {object} item - 対象アイテム
     * @param {string} name - タグ名
     * @returns {string|null} タグの値（なければ null）
     */
    function getTagValue(item, name) {
        if (!item) return null;
        for (var i = 0; i < item.tags.length; i++) {
            if (item.tags[i].name === name) return item.tags[i].value;
        }
        return null;
    }

    /**
     * 引き出し線グループであることを示すメモ（note）を付ける
     * @param {GroupItem} item - 対象のグループ
     * @returns {void}
     */
    function setLeaderLineTag(item) {
        safeSetNote(item, "leader_line");
    }

    /**
     * 引き出し線の各パーツに再適用用のメモ（note）を付ける
     * @param {object} parts - assembleLeaderParts() の戻り値
     * @param {PathItem} edgePath - フチの線
     * @param {PathItem} mainPath - 本体の線
     * @param {PathItem} edgeTipMarker - フチの線端
     * @param {PathItem} mainTipMarker - 本体の線端
     * @param {boolean} hasEdge - フチがあるかどうか
     * @returns {void}
     */
    function tagLeaderParts(parts, edgePath, mainPath, edgeTipMarker, mainTipMarker, hasEdge) {
        if (hasEdge) {
            safeSetNote(parts.mainGroup, "leader_line_main");
            safeSetNote(parts.edgeGroup, "leader_line_edge");
        }

        safeSetNote(mainPath, "leader_line_main_path");
        safeSetNote(mainTipMarker, "leader_line_main_cap");
        safeSetNote(edgePath, "leader_line_edge_path");
        safeSetNote(edgeTipMarker, "leader_line_edge_cap");
    }

    /**
     * 再適用時に使う元オブジェクトの外接矩形をタグに保存する
     * @param {GroupItem} item - 対象のグループ
     * @param {object} targetMetrics - measureTargets() の要素
     * @returns {void}
     */
    function setLeaderLineBoundsTags(item, targetMetrics) {
        if (!item || !targetMetrics) return;
        setTagValue(item, "leader_line_x_left", formatNumber(targetMetrics.x_left, 4));
        setTagValue(item, "leader_line_x_right", formatNumber(targetMetrics.x_right, 4));
        setTagValue(item, "leader_line_y_top", formatNumber(targetMetrics.y_top, 4));
        setTagValue(item, "leader_line_y_bottom", formatNumber(targetMetrics.y_bottom, 4));
    }

    /**
     * 再適用時に使う斜線の方向をタグに保存する
     * @param {GroupItem} item - 対象のグループ
     * @param {string} hDir - 水平方向（"left" / "right"）
     * @param {string} vDir - 垂直方向（"up" / "down"）
     * @returns {void}
     */
    function setLeaderLineDirTags(item, hDir, vDir) {
        if (!item) return;
        setTagValue(item, "leader_line_hDir", hDir);
        setTagValue(item, "leader_line_vDir", vDir);
    }

    /**
     * タグに保存された斜線の方向を取得する
     * @param {GroupItem} item - 対象のグループ
     * @returns {object|null} hDir / vDir を持つオブジェクト（なければ null）
     */
    function getLeaderLineStoredDir(item) {
        var hDir = getTagValue(item, "leader_line_hDir");
        var vDir = getTagValue(item, "leader_line_vDir");
        if (!hDir || !vDir) return null;
        return { hDir: hDir, vDir: vDir };
    }

    /**
     * タグに保存された外接矩形を取得する
     * @param {GroupItem} item - 対象のグループ
     * @returns {object|null} x_left / x_right / y_top / y_bottom を持つ矩形（なければ null）
     */
    function getLeaderLineStoredBounds(item) {
        var xLeft = parseFloat(getTagValue(item, "leader_line_x_left"));
        var xRight = parseFloat(getTagValue(item, "leader_line_x_right"));
        var yTop = parseFloat(getTagValue(item, "leader_line_y_top"));
        var yBottom = parseFloat(getTagValue(item, "leader_line_y_bottom"));
        if (isNaN(xLeft) || isNaN(xRight) || isNaN(yTop) || isNaN(yBottom)) return null;
        return {
            x_left: xLeft,
            x_right: xRight,
            y_top: yTop,
            y_bottom: yBottom
        };
    }

    /**
     * 引き出し線グループかどうかを判定する
     * @param {object} item - 対象アイテム
     * @returns {boolean} 引き出し線グループなら true
     */
    function isLeaderLineGroup(item) {
        return item && item.typename === "GroupItem" && item.note === "leader_line";
    }

    /**
     * 引き出し線グループから線端（丸・矢印）のパスを取得する
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {PathItem|null} 線端のパス（なければ null）
     */
    function getLeaderLineTipMarker(groupItem) {
        for (var i = 0; i < groupItem.pathItems.length; i++) {
            var pathItem = groupItem.pathItems[i];
            if (pathItem.note === "leader_line_main_cap") return pathItem;
        }
        return null;
    }

    /**
     * 引き出し線グループから基準線（A線）を取得する
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {PathItem|null} 基準線（なければ null）
     */
    function getLeaderLineBasePath(groupItem) {
        var i, pathItem;

        /* 再適用時の基準線は常に A線のみとし、B/C/D は参照しない
           新構造なら note 付きの A線を最優先 */
        for (i = 0; i < groupItem.pathItems.length; i++) {
            pathItem = groupItem.pathItems[i];
            if (pathItem.note === "leader_line_main_path") return pathItem;
        }

        /* 旧構造用フォールバック：白でない stroked の open path を優先 */
        for (i = 0; i < groupItem.pathItems.length; i++) {
            pathItem = groupItem.pathItems[i];
            if (pathItem.stroked && !pathItem.closed && !isWhiteCMYKColor(pathItem.strokeColor)) return pathItem;
        }

        /* 最後のフォールバック：stroked の open path */
        for (i = 0; i < groupItem.pathItems.length; i++) {
            pathItem = groupItem.pathItems[i];
            if (pathItem.stroked && !pathItem.closed) return pathItem;
        }

        return null;
    }

    // =========================================
    // 外接矩形と線の属性 / Bounds and stroke
    // =========================================

    /**
     * GroupItem の参照用 PathItem を取得する
     * 通常は closed path を優先して返す（leader_line 再適用時の A線優先ロジックは getLeaderLineBasePath() 側で処理）
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {Array<PathItem>} 参照用のパス配列
     */
    function getGroupReferencePathItems(groupItem) {
        var closedPaths = [];
        var allPaths = [];
        for (var k = 0; k < groupItem.pathItems.length; k++) {
            var pathItem = groupItem.pathItems[k];
            allPaths.push(pathItem);
            if (pathItem.closed) closedPaths.push(pathItem);
        }
        return closedPaths.length ? closedPaths : allPaths;
    }

    /**
     * PathItem のアンカー座標から外接矩形を取得する
     * @param {PathItem} pathItem - 対象のパス
     * @returns {object} x_left / x_right / y_top / y_bottom を持つ矩形
     */
    function getPathAnchorBounds(pathItem) {
        var points = pathItem.pathPoints;
        if (!points || points.length === 0) {
            return {
                x_left: 0,
                x_right: 0,
                y_top: 0,
                y_bottom: 0
            };
        }

        var xMin = points[0].anchor[0], xMax = points[0].anchor[0];
        var yMin = points[0].anchor[1], yMax = points[0].anchor[1];
        for (var i = 1; i < points.length; i++) {
            var anchorX = points[i].anchor[0];
            var anchorY = points[i].anchor[1];
            if (anchorX < xMin) xMin = anchorX;
            if (anchorX > xMax) xMax = anchorX;
            if (anchorY < yMin) yMin = anchorY;
            if (anchorY > yMax) yMax = anchorY;
        }

        return {
            x_left: xMin,
            x_right: xMax,
            y_top: yMax,
            y_bottom: yMin
        };
    }

    /**
     * 線端パスの中心座標を取得する
     * @param {PathItem} capItem - 線端のパス
     * @returns {Array<number>|null} [x, y]（なければ null）
     */
    function getTipMarkerCenter(capItem) {
        if (!capItem) return null;
        var capBounds = capItem.geometricBounds; /* [left, top, right, bottom] */
        return [
            (capBounds[0] + capBounds[2]) / 2,
            (capBounds[1] + capBounds[3]) / 2
        ];
    }

    /**
     * 線端の中心を先端とみなして、引き出し線の外接矩形を求める
     * @param {PathItem} pathItem - 引き出し線のパス
     * @param {PathItem} capItem - 線端のパス
     * @returns {object} x_left / x_right / y_top / y_bottom を持つ矩形
     */
    function getLeaderLineBoundsWithTipMarker(pathItem, capItem) {
        var bounds = getPathAnchorBounds(pathItem);
        if (!capItem) return bounds;

        var points = pathItem.pathPoints;
        if (!points || points.length < 2) return bounds;

        var p0 = points[0].anchor;
        var p1 = points[1].anchor;
        var p2 = points[points.length - 1].anchor;
        var capCenter = getTipMarkerCenter(capItem);
        if (!capCenter) return bounds;

        /* 線端に丸がある場合、丸の中心（= 丸の半分位置）を期待する端として扱う */
        var centerX = capCenter[0];
        var centerY = capCenter[1];

        var distanceToStart = Math.abs(p0[0] - centerX) + Math.abs(p0[1] - centerY);
        var distanceToEnd = Math.abs(p2[0] - centerX) + Math.abs(p2[1] - centerY);

        var tipStart = [p0[0], p0[1]];
        var bend = [p1[0], p1[1]];
        var tipEnd = [p2[0], p2[1]];

        if (distanceToStart <= distanceToEnd) {
            tipStart = [centerX, centerY];
        } else {
            tipEnd = [centerX, centerY];
        }

        return {
            x_left: Math.min(tipStart[0], bend[0], tipEnd[0]),
            x_right: Math.max(tipStart[0], bend[0], tipEnd[0]),
            y_top: Math.max(tipStart[1], bend[1], tipEnd[1]),
            y_bottom: Math.min(tipStart[1], bend[1], tipEnd[1])
        };
    }

    /**
     * GroupItem の外接矩形を取得する
     * 通常はグループ全体、leader_line 再適用時は保存済み bounds を優先し、なければ旧構造から復元する
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {object} x_left / x_right / y_top / y_bottom を持つ矩形
     */
    function getGroupBounds(groupItem) {
        if (isLeaderLineGroup(groupItem)) {
            var storedBounds = getLeaderLineStoredBounds(groupItem);
            if (storedBounds) {
                return storedBounds;
            }

            var basePath = getLeaderLineBasePath(groupItem);
            if (basePath) {
                var mainTipMarker = getLeaderLineTipMarker(groupItem);
                return getLeaderLineBoundsWithTipMarker(basePath, mainTipMarker);
            }
        }

        var geoBounds = groupItem.geometricBounds; /* [left, top, right, bottom] */
        return {
            x_left: geoBounds[0],
            x_right: geoBounds[2],
            y_top: geoBounds[1],
            y_bottom: geoBounds[3]
        };
    }

    /**
     * 対象の線の属性を取得する
     * GroupItem は通常 closed path 優先、leader_line 再適用時は A線を優先する
     * @param {PathItem|GroupItem} item - 対象アイテム
     * @returns {object} stroked / strokeWidth / strokeColor を持つオブジェクト
     */
    function getStrokeInfo(item) {
        if (item.typename === "PathItem") {
            return {
                stroked: item.stroked,
                strokeWidth: item.stroked ? item.strokeWidth : DEFAULT_LINE_WIDTH_PT,
                strokeColor: item.stroked ? item.strokeColor : undefined
            };
        }

        /* 引き出し線グループの場合、A線（基準線）から取得 */
        if (isLeaderLineGroup(item)) {
            var leaderBasePath = getLeaderLineBasePath(item);
            if (leaderBasePath) {
                return {
                    stroked: leaderBasePath.stroked,
                    strokeWidth: leaderBasePath.stroked ? leaderBasePath.strokeWidth : DEFAULT_LINE_WIDTH_PT,
                    strokeColor: leaderBasePath.stroked ? leaderBasePath.strokeColor : undefined
                };
            }
        }

        /* GroupItemの場合、closed pathを優先して配下のPathItemを探す */
        var referencePaths = getGroupReferencePathItems(item);
        for (var k = 0; k < referencePaths.length; k++) {
            var pathItem = referencePaths[k];
            if (pathItem.stroked) {
                return {
                    stroked: true,
                    strokeWidth: pathItem.strokeWidth,
                    strokeColor: pathItem.strokeColor
                };
            }
        }
        return { stroked: false, strokeWidth: DEFAULT_LINE_WIDTH_PT, strokeColor: undefined };
    }

    /**
     * 各ターゲットの座標情報と線の属性を都度取得する
     * @param {Array<PathItem|GroupItem>} itemsToMeasure - 対象アイテムの配列
     * @returns {Array<object>} 外接矩形と線の属性をまとめた配列
     */
    function measureTargets(itemsToMeasure) {
        var data = [];
        for (var i = 0; i < itemsToMeasure.length; i++) {
            var item = itemsToMeasure[i];
            var x_left, x_right, y_bottom, y_top;

            if (item.typename === "PathItem") {
                /* PathItem：アンカーポイントから外接矩形を算出 */
                var pathBounds = getPathAnchorBounds(item);
                x_left = pathBounds.x_left;
                x_right = pathBounds.x_right;
                y_bottom = pathBounds.y_bottom;
                y_top = pathBounds.y_top;
            } else {
                /* GroupItem：通常はグループ全体、leader_line 再適用時は保存済み bounds を優先して使用 */
                var groupBounds = getGroupBounds(item);
                x_left = groupBounds.x_left;
                x_right = groupBounds.x_right;
                y_top = groupBounds.y_top;
                y_bottom = groupBounds.y_bottom;
            }

            var strokeInfo = getStrokeInfo(item);
            data.push({
                target: item,
                x_left: x_left,
                x_right: x_right,
                y_bottom: y_bottom,
                y_top: y_top,
                stroked: strokeInfo.stroked,
                strokeWidth: strokeInfo.strokeWidth,
                strokeColor: strokeInfo.strokeColor
            });
        }
        return data;
    }

    /**
     * テキストフレームの外接矩形を取得する
     * @param {TextFrame} textItem - 対象のテキストフレーム
     * @returns {object} x_left / x_right / y_top / y_bottom を持つ矩形
     */
    function getTextBounds(textItem) {
        var geoBounds = textItem.geometricBounds; /* [left, top, right, bottom] */
        return { x_left: geoBounds[0], x_right: geoBounds[2], y_top: geoBounds[1], y_bottom: geoBounds[3] };
    }

    // =========================================
    // カラー / Color
    // =========================================

    /**
     * HEXカラーを RGBColor に変換する
     * @param {string} hex - "#rrggbb" または "#rgb"
     * @returns {RGBColor|null} 変換したカラー（不正な場合は null）
     */
    function hexToRGBColor(hex) {
        hex = hex.replace(/^#/, "");
        if (hex.length === 3) {
            hex = hex.charAt(0) + hex.charAt(0) + hex.charAt(1) + hex.charAt(1) + hex.charAt(2) + hex.charAt(2);
        }
        var red = parseInt(hex.substring(0, 2), 16);
        var green = parseInt(hex.substring(2, 4), 16);
        var blue = parseInt(hex.substring(4, 6), 16);
        if (isNaN(red) || isNaN(green) || isNaN(blue)) return null;
        var color = new RGBColor();
        color.red = red;
        color.green = green;
        color.blue = blue;
        return color;
    }

    /**
     * ドキュメントのカラースペースがCMYKかどうかを判定する
     * @param {Document} doc - 対象ドキュメント
     * @returns {boolean} CMYKなら true
     */
    function isCMYKDocument(doc) {
        return doc.documentColorSpace === DocumentColorSpace.CMYK;
    }

    /**
     * ドキュメントのカラースペースに合わせた黒を作る
     * @param {Document} doc - 対象ドキュメント
     * @returns {CMYKColor|RGBColor} 黒のカラー
     */
    function createBlackColor(doc) {
        if (isCMYKDocument(doc)) {
            var cmykBlack = new CMYKColor();
            cmykBlack.cyan = 0;
            cmykBlack.magenta = 0;
            cmykBlack.yellow = 0;
            cmykBlack.black = 100;
            return cmykBlack;
        }
        var rgbBlack = new RGBColor();
        rgbBlack.red = 0;
        rgbBlack.green = 0;
        rgbBlack.blue = 0;
        return rgbBlack;
    }

    /**
     * ドキュメントのカラースペースに合わせた白を作る
     * @param {Document} doc - 対象ドキュメント
     * @returns {CMYKColor|RGBColor} 白のカラー
     */
    function createWhiteColor(doc) {
        if (isCMYKDocument(doc)) {
            var cmykWhite = new CMYKColor();
            cmykWhite.cyan = 0;
            cmykWhite.magenta = 0;
            cmykWhite.yellow = 0;
            cmykWhite.black = 0;
            return cmykWhite;
        }
        var rgbWhite = new RGBColor();
        rgbWhite.red = 255;
        rgbWhite.green = 255;
        rgbWhite.blue = 255;
        return rgbWhite;
    }

    /**
     * RGB値を CMYKColor に変換する
     * @param {number} red - 赤（0-255）
     * @param {number} green - 緑（0-255）
     * @param {number} blue - 青（0-255）
     * @returns {CMYKColor} 変換したカラー
     */
    function rgbToCMYKColor(red, green, blue) {
        var redRatio = Math.max(0, Math.min(255, red)) / 255;
        var greenRatio = Math.max(0, Math.min(255, green)) / 255;
        var blueRatio = Math.max(0, Math.min(255, blue)) / 255;
        var blackRatio = 1 - Math.max(redRatio, greenRatio, blueRatio);
        var cyanRatio = 0, magentaRatio = 0, yellowRatio = 0;

        if (blackRatio < 1) {
            cyanRatio = (1 - redRatio - blackRatio) / (1 - blackRatio);
            magentaRatio = (1 - greenRatio - blackRatio) / (1 - blackRatio);
            yellowRatio = (1 - blueRatio - blackRatio) / (1 - blackRatio);
        }

        var color = new CMYKColor();
        color.cyan = Math.round(cyanRatio * 100);
        color.magenta = Math.round(magentaRatio * 100);
        color.yellow = Math.round(yellowRatio * 100);
        color.black = Math.round(blackRatio * 100);
        return color;
    }

    /**
     * HEXカラーをドキュメントのカラースペースに合わせて変換する
     * @param {Document} doc - 対象ドキュメント
     * @param {string} hex - "#rrggbb" または "#rgb"
     * @returns {CMYKColor|RGBColor|null} 変換したカラー（不正な場合は null）
     */
    function hexToDocumentColor(doc, hex) {
        var rgb = hexToRGBColor(hex);
        if (!rgb) return null;
        if (isCMYKDocument(doc)) {
            return rgbToCMYKColor(rgb.red, rgb.green, rgb.blue);
        }
        return rgb;
    }

    /**
     * 白のCMYKカラーかどうかを判定する
     * @param {object} color - 対象のカラー
     * @returns {boolean} 白のCMYKカラーなら true
     */
    function isWhiteCMYKColor(color) {
        if (!color) return false;
        if (color.typename !== "CMYKColor") return false;
        return color.cyan === 0 && color.magenta === 0 && color.yellow === 0 && color.black === 0;
    }

    /**
     * スウォッチ表示を指定のHEX値に合わせて更新する
     * @param {Panel} colorChip - 対象のスウォッチ
     * @param {string} hex - "#rrggbb" または "#rgb"
     * @returns {void}
     */
    function updateColorChip(colorChip, hex) {
        if (!colorChip) return;
        var rgb = hexToRGBColor(String(hex || ""));
        if (!rgb) return;
        var graphics = colorChip.graphics;
        graphics.backgroundColor = graphics.newBrush(
            graphics.BrushType.SOLID_COLOR,
            [rgb.red / 255, rgb.green / 255, rgb.blue / 255]
        );
    }

    /**
     * 生成条件からフチのカラーを求める
     * @param {Document} doc - 対象ドキュメント
     * @param {object} options - readLeaderOptions() の戻り値
     * @returns {CMYKColor|RGBColor} フチのカラー
     */
    function resolveEdgeColor(doc, options) {
        if (options.edgeColorMode === "white") {
            return createWhiteColor(doc);
        }
        var documentColor = hexToDocumentColor(doc, options.edgeColorHex);
        if (documentColor) return documentColor;
        return createWhiteColor(doc);
    }

    /**
     * 生成条件から線のカラーを求める
     * @param {Document} doc - 対象ドキュメント
     * @param {object} options - readLeaderOptions() の戻り値
     * @param {object} targetMetrics - measureTargets() の要素
     * @returns {CMYKColor|RGBColor} 線のカラー
     */
    function resolveLineColor(doc, options, targetMetrics) {
        if (options.lineColorMode === "black") {
            return createBlackColor(doc);
        }
        if (options.lineColorMode === "white") {
            return createWhiteColor(doc);
        }
        /* その他：HEXカラーコードから現在のドキュメント色空間に合わせて変換 */
        var documentColor = hexToDocumentColor(doc, options.lineColorHex);
        if (documentColor) return documentColor;
        /* 無効な値の場合は元のオブジェクトの線色を使用 */
        if (targetMetrics.strokeColor) return targetMetrics.strokeColor;
        return createBlackColor(doc);
    }

    /**
     * 生成条件から線幅（pt）を求める
     * @param {object} options - readLeaderOptions() の戻り値
     * @param {object} targetMetrics - measureTargets() の要素
     * @returns {number} 線幅（pt）
     */
    function resolveLineWidth(options, targetMetrics) {
        var widthPt = options.lineWidthPt;
        if (!isNaN(widthPt) && widthPt > 0) return widthPt;
        return targetMetrics.strokeWidth;
    }

    /**
     * 引き出し線1本分の描画スタイルをまとめて求める
     * @param {Document} doc - 対象ドキュメント
     * @param {object} options - readLeaderOptions() の戻り値
     * @param {object} targetMetrics - measureTargets() の要素
     * @returns {object} lineColor / edgeColor / widthPt を持つオブジェクト
     */
    function resolveLeaderStyle(doc, options, targetMetrics) {
        return {
            lineColor: resolveLineColor(doc, options, targetMetrics),
            edgeColor: resolveEdgeColor(doc, options),
            widthPt: resolveLineWidth(options, targetMetrics)
        };
    }

    // =========================================
    // 引き出し線の座標計算 / Leader line geometry
    // =========================================

    /**
     * 引き出し線の座標を計算する
     * @param {object} targetMetrics - measureTargets() の要素
     * @param {number} angleRad - 斜線の角度（ラジアン）
     * @param {string} hDir - 水平方向（"left" / "right"）
     * @param {string} vDir - 垂直方向（"up" / "down"）
     * @returns {Array<Array<number>>} 3点の座標配列
     */
    function calcLeaderPoints(targetMetrics, angleRad, hDir, vDir) {
        var height = targetMetrics.y_top - targetMetrics.y_bottom;
        var offset = height / Math.tan(angleRad);
        var bendX;
        if (vDir === "down") {
            /* 水平線が上、斜線が下に向かう */
            if (hDir === "left") {
                bendX = targetMetrics.x_left + offset;
                return [[targetMetrics.x_left, targetMetrics.y_bottom], [bendX, targetMetrics.y_top], [targetMetrics.x_right, targetMetrics.y_top]];
            }
            bendX = targetMetrics.x_right - offset;
            return [[targetMetrics.x_left, targetMetrics.y_top], [bendX, targetMetrics.y_top], [targetMetrics.x_right, targetMetrics.y_bottom]];
        }
        /* 水平線が下、斜線が上に向かう */
        if (hDir === "left") {
            bendX = targetMetrics.x_left + offset;
            return [[targetMetrics.x_left, targetMetrics.y_top], [bendX, targetMetrics.y_bottom], [targetMetrics.x_right, targetMetrics.y_bottom]];
        }
        bendX = targetMetrics.x_right - offset;
        return [[targetMetrics.x_left, targetMetrics.y_bottom], [bendX, targetMetrics.y_bottom], [targetMetrics.x_right, targetMetrics.y_top]];
    }

    /**
     * テキスト整列時の引き出し線座標を計算する
     * オブジェクトとテキストの位置関係から方向を自動判定し、水平端をテキストの対応する端にそろえ、
     * 折れ点のY座標をテキスト端から指定距離だけ離す
     * @param {object} targetMetrics - measureTargets() の要素
     * @param {number} angleRad - 斜線の角度（ラジアン）
     * @param {object} textBounds - テキストの外接矩形
     * @param {number} textDistPt - テキストとの距離（pt）
     * @returns {object} points（3点の座標配列）／hDir／vDir を持つオブジェクト
     */
    function calcLeaderPointsWithText(targetMetrics, angleRad, textBounds, textDistPt) {
        var bendY, horizontalEndX, tipX, tipY, diffY, diffX, bendX;

        /* オブジェクトとテキストの中心座標を比較して方向を自動判定
           （Illustratorは Y 上向き正: 下にあるほど Y が小さい） */
        var objectCenterX = (targetMetrics.x_left + targetMetrics.x_right) / 2;
        var objectCenterY = (targetMetrics.y_top + targetMetrics.y_bottom) / 2;
        var textCenterX = (textBounds.x_left + textBounds.x_right) / 2;
        var textCenterY = (textBounds.y_top + textBounds.y_bottom) / 2;
        var autoHDir = (objectCenterX < textCenterX) ? "left" : "right";
        var autoVDir = (objectCenterY < textCenterY) ? "down" : "up";

        /* オブジェクトがテキストより下 → 折れ点はテキスト下端-textDistPt
           オブジェクトがテキストより上 → 折れ点はテキスト上端+textDistPt */
        if (autoVDir === "down") {
            bendY = textBounds.y_bottom - textDistPt;
        } else {
            bendY = textBounds.y_top + textDistPt;
        }

        tipY = (autoVDir === "down") ? targetMetrics.y_bottom : targetMetrics.y_top;
        diffY = Math.abs(bendY - tipY);
        diffX = diffY / Math.tan(angleRad);

        if (autoHDir === "left") {
            /* オブジェクトがテキストより左: 水平端をテキスト右端へ、先端はオブジェクト左 */
            horizontalEndX = textBounds.x_right;
            tipX = targetMetrics.x_left;
            bendX = tipX + diffX;
            return { points: [[tipX, tipY], [bendX, bendY], [horizontalEndX, bendY]], hDir: autoHDir, vDir: autoVDir };
        }
        /* オブジェクトがテキストより右: 水平端をテキスト左端へ、先端はオブジェクト右 */
        horizontalEndX = textBounds.x_left;
        tipX = targetMetrics.x_right;
        bendX = tipX - diffX;
        return { points: [[horizontalEndX, bendY], [bendX, bendY], [tipX, tipY]], hDir: autoHDir, vDir: autoVDir };
    }

    /**
     * 引き出し線の座標と、その座標が示す方向を求める
     * テキスト整列時は方向を自動判定するため、指定した方向とは異なる結果を返すことがある
     * @param {object} targetMetrics - measureTargets() の要素
     * @param {number} angleRad - 斜線の角度（ラジアン）
     * @param {string} hDir - 水平方向（"left" / "right"）
     * @param {string} vDir - 垂直方向（"up" / "down"）
     * @param {object} [textBounds] - テキストの外接矩形（テキスト整列時のみ）
     * @param {number} [textDistPt] - テキストとの距離（pt）
     * @returns {object} points（3点の座標配列）／hDir／vDir を持つオブジェクト
     */
    function calcLeaderGeometry(targetMetrics, angleRad, hDir, vDir, textBounds, textDistPt) {
        if (textBounds) {
            return calcLeaderPointsWithText(targetMetrics, angleRad, textBounds, textDistPt);
        }
        return { points: calcLeaderPoints(targetMetrics, angleRad, hDir, vDir), hDir: hDir, vDir: vDir };
    }

    /**
     * 斜線の先端座標を取得する
     * @param {Array<Array<number>>} points - 引き出し線の座標配列
     * @param {string} hDir - 水平方向（"left" / "right"）
     * @returns {Array<number>} 先端の座標 [x, y]
     */
    function getTipPoint(points, hDir) {
        return (hDir === "left") ? points[0] : points[2];
    }

    /**
     * 斜線の先端を指定長さだけ短縮する
     * @param {Array<Array<number>>} points - 引き出し線の座標配列（破壊的に更新）
     * @param {string} hDir - 水平方向（"left" / "right"）
     * @param {number} length - 短縮する長さ（pt）
     * @returns {void}
     */
    function shortenTip(points, hDir, length) {
        var tipIndex = (hDir === "left") ? 0 : 2;
        var tip = points[tipIndex];
        var bend = points[1];
        var dx = bend[0] - tip[0];
        var dy = bend[1] - tip[1];
        var distance = Math.sqrt(dx * dx + dy * dy);
        if (distance > 0) {
            points[tipIndex] = [tip[0] + dx / distance * length, tip[1] + dy / distance * length];
        }
    }

    /**
     * 矢印の3頂点を計算する
     * @param {Array<number>} tipPt - 先端の座標 [x, y]
     * @param {Array<number>} bendPt - 折れ点の座標 [x, y]
     * @param {number} arrowSize - 矢印の長さ（pt）
     * @returns {Array<Array<number>>|null} 先端・翼端・翼端の座標配列（長さが0の場合は null）
     */
    function calcArrowPoints(tipPt, bendPt, arrowSize) {
        /* tipPt → bendPt 方向のベクトル */
        var dx = bendPt[0] - tipPt[0];
        var dy = bendPt[1] - tipPt[1];
        var distance = Math.sqrt(dx * dx + dy * dy);
        if (distance === 0) return null;
        var unitX = dx / distance;
        var unitY = dy / distance;
        /* 矢印の2つの翼端を計算 */
        var halfWidth = arrowSize / 2;
        var backX = tipPt[0] + unitX * arrowSize;
        var backY = tipPt[1] + unitY * arrowSize;
        return [
            [tipPt[0], tipPt[1]],
            [backX + unitY * halfWidth, backY - unitX * halfWidth],
            [backX - unitY * halfWidth, backY + unitX * halfWidth]
        ];
    }

    // =========================================
    // 引き出し線の生成 / Leader line drawing
    // =========================================

    /**
     * 先端に円を追加する
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<number>} tipPt - 先端の座標 [x, y]
     * @param {object} style - resolveLeaderStyle() の戻り値
     * @param {number} diameter - 円の直径（pt）
     * @param {boolean} fillOnly - true なら塗り、false なら線で描く
     * @returns {PathItem} 追加した円
     */
    function addTipCircleMarker(doc, tipPt, style, diameter, fillOnly) {
        var radius = diameter / 2;
        var circle = doc.pathItems.ellipse(
            tipPt[1] + radius, tipPt[0] - radius, diameter, diameter
        );
        if (fillOnly) {
            /* ●（塗りのみ） */
            circle.filled = true;
            circle.fillColor = style.lineColor;
            circle.stroked = false;
        } else {
            /* ○（線のみ） */
            circle.filled = false;
            circle.stroked = true;
            circle.strokeWidth = style.widthPt;
            circle.strokeColor = style.lineColor;
        }
        return circle;
    }

    /**
     * 先端の円にフチを追加する（●／○ とも一回り大きい塗り円を背面に置く）
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<number>} tipPt - 先端の座標 [x, y]
     * @param {object} style - resolveLeaderStyle() の戻り値
     * @param {number} diameter - 本体の円の直径（pt）
     * @returns {PathItem} 追加したフチの円
     */
    function addTipCircleMarkerEdge(doc, tipPt, style, diameter) {
        var expandedDiameter = diameter + style.widthPt * EDGE_CIRCLE_EXPAND_RATIO;
        var expandedRadius = expandedDiameter / 2;
        var edgeCircle = doc.pathItems.ellipse(
            tipPt[1] + expandedRadius, tipPt[0] - expandedRadius, expandedDiameter, expandedDiameter
        );
        edgeCircle.filled = true;
        edgeCircle.fillColor = style.edgeColor;
        edgeCircle.stroked = false;
        return edgeCircle;
    }

    /**
     * 先端に矢印を追加する
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<number>} tipPt - 先端の座標 [x, y]
     * @param {Array<number>} bendPt - 折れ点の座標 [x, y]
     * @param {object} style - resolveLeaderStyle() の戻り値
     * @param {number} arrowSize - 矢印の長さ（pt）
     * @param {boolean} fillOnly - true なら塗り、false なら線で描く
     * @returns {PathItem|null} 追加した矢印（長さが0の場合は null）
     */
    function addTipArrow(doc, tipPt, bendPt, style, arrowSize, fillOnly) {
        var arrowPoints = calcArrowPoints(tipPt, bendPt, arrowSize);
        if (!arrowPoints) return null;

        var arrow = doc.pathItems.add();
        arrow.setEntirePath(arrowPoints);
        arrow.closed = true;
        if (fillOnly) {
            arrow.filled = true;
            arrow.fillColor = style.lineColor;
            arrow.stroked = false;
        } else {
            arrow.filled = false;
            arrow.stroked = true;
            arrow.strokeWidth = style.widthPt;
            arrow.strokeColor = style.lineColor;
        }
        return arrow;
    }

    /**
     * 先端の矢印にフチを追加する（本体と同じ重心を基準にスケールアップ）
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<number>} tipPt - 先端の座標 [x, y]
     * @param {Array<number>} bendPt - 折れ点の座標 [x, y]
     * @param {object} style - resolveLeaderStyle() の戻り値
     * @param {number} arrowSize - 本体の矢印の長さ（pt）
     * @returns {PathItem|null} 追加したフチの矢印（長さが0の場合は null）
     */
    function addTipArrowEdge(doc, tipPt, bendPt, style, arrowSize) {
        var arrowPoints = calcArrowPoints(tipPt, bendPt, arrowSize);
        if (!arrowPoints) return null;
        /* 重心を求める */
        var centerX = (arrowPoints[0][0] + arrowPoints[1][0] + arrowPoints[2][0]) / 3;
        var centerY = (arrowPoints[0][1] + arrowPoints[1][1] + arrowPoints[2][1]) / 3;
        /* 重心を基準にスケールアップ */
        var scale = (arrowSize + style.widthPt * EDGE_ARROW_EXPAND_RATIO) / arrowSize;
        var scaledPoints = [];
        for (var i = 0; i < arrowPoints.length; i++) {
            scaledPoints.push([
                centerX + (arrowPoints[i][0] - centerX) * scale,
                centerY + (arrowPoints[i][1] - centerY) * scale
            ]);
        }

        var edgeArrow = doc.pathItems.add();
        edgeArrow.setEntirePath(scaledPoints);
        edgeArrow.closed = true;
        edgeArrow.filled = true;
        edgeArrow.fillColor = style.edgeColor;
        edgeArrow.stroked = false;
        return edgeArrow;
    }

    /**
     * 引き出し線の各パーツをグループに収める
     * @param {GroupItem} container - 収める先のグループ
     * @param {PathItem} edgePath - フチの線
     * @param {PathItem} mainPath - 本体の線
     * @param {PathItem} edgeTipMarker - フチの線端
     * @param {PathItem} mainTipMarker - 本体の線端
     * @param {boolean} hasEdge - フチがあるかどうか
     * @returns {object} edgeGroup / mainGroup を持つオブジェクト
     */
    function assembleLeaderParts(container, edgePath, mainPath, edgeTipMarker, mainTipMarker, hasEdge) {
        var edgeGroup = null;
        var mainGroup = container;

        if (hasEdge) {
            edgeGroup = container.groupItems.add();
            mainGroup = container.groupItems.add();

            if (mainPath) mainPath.move(mainGroup, ElementPlacement.PLACEATEND);
            if (mainTipMarker) mainTipMarker.move(mainGroup, ElementPlacement.PLACEATEND);

            if (edgePath) edgePath.move(edgeGroup, ElementPlacement.PLACEATEND);
            if (edgeTipMarker) edgeTipMarker.move(edgeGroup, ElementPlacement.PLACEATEND);
        } else {
            if (mainPath) mainPath.move(container, ElementPlacement.PLACEATEND);
            if (mainTipMarker) mainTipMarker.move(container, ElementPlacement.PLACEATEND);
        }

        return {
            edgeGroup: edgeGroup,
            mainGroup: mainGroup
        };
    }

    /**
     * ダイアログから引き出し線の生成条件を読み取る
     * @param {object} ui - buildDialogUI() の戻り値
     * @param {Array<TextFrame>} textFrameTargets - 選択中のテキストフレーム
     * @returns {object} 生成条件をまとめたオブジェクト
     */
    function readLeaderOptions(ui, textFrameTargets) {
        var diagDirection = getDiagDirValues(ui);

        var markerSizePt = unitValueToPt(ui.tipMarkerSizeInput.text, "strokeUnits");
        if (isNaN(markerSizePt) || markerSizePt <= 0) markerSizePt = DEFAULT_TIP_MARKER_SIZE_PT;

        /* テキストを選択しているときだけテキスト整列を使う */
        var textBounds = null;
        var textDistPt = 0;
        if (textFrameTargets.length > 0) {
            textBounds = getTextBounds(textFrameTargets[0]);
            textDistPt = parseFloat(ui.textDistInput.text);
            if (isNaN(textDistPt) || textDistPt < 0) textDistPt = DEFAULT_TEXT_DIST_PT;
        }

        var options = {
            hDir: diagDirection.hDir,
            vDir: diagDirection.vDir,
            useStoredDir: ui.applyScopeKeepDirRadio.value,
            useCircleMarker: ui.tipMarkerCircleRadio.value,
            useArrowMarker: ui.tipMarkerArrowRadio.value,
            fillTipMarker: ui.tipMarkerFillRadio.value,
            hasEdge: ui.edgeEnabledCheck.value,
            roundStrokeCap: ui.strokeCapRoundRadio.value,
            markerSizePt: markerSizePt,
            textBounds: textBounds,
            textDistPt: textDistPt,
            lineColorMode: getSelectedRadioKey(ui.radioMaps.lineColor, "black"),
            lineColorHex: ui.lineColorHexHolder.text,
            edgeColorMode: getSelectedRadioKey(ui.radioMaps.edgeColor, "white"),
            edgeColorHex: ui.edgeColorHexHolder.text,
            lineWidthPt: unitValueToPt(ui.lineWidthInput.text, "strokeUnits")
        };
        options.hasTipMarker = options.useCircleMarker || options.useArrowMarker;
        return options;
    }

    /**
     * 引き出し線1本分のパーツ（フチ・本体・線端）を生成する
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem|GroupItem} targetItem - 元になったアイテム
     * @param {object} targetMetrics - measureTargets() の要素
     * @param {number} angleRad - 斜線の角度（ラジアン）
     * @param {object} options - readLeaderOptions() の戻り値
     * @param {Array<object>} [createdArtItems] - 生成したアイテムを追記する配列（後始末用）
     * @returns {object} 各パーツと、実際に使った方向を持つオブジェクト
     */
    function buildLeaderLineParts(doc, targetItem, targetMetrics, angleRad, options, createdArtItems) {
        createdArtItems = createdArtItems || [];

        /* 「斜線の方向以外」のとき、各オブジェクトの保存済み方向を使用 */
        var itemHDir = options.hDir;
        var itemVDir = options.vDir;
        if (options.useStoredDir) {
            var storedDir = getLeaderLineStoredDir(targetItem);
            if (storedDir) {
                itemHDir = storedDir.hDir;
                itemVDir = storedDir.vDir;
            }
        }

        var style = resolveLeaderStyle(doc, options, targetMetrics);
        /* テキスト整列時は方向が自動判定されるので、以降は geometry が返した方向を使う */
        var geometry = calcLeaderGeometry(targetMetrics, angleRad, itemHDir, itemVDir, options.textBounds, options.textDistPt);
        var points = geometry.points;

        /* 先端座標は短縮前に取得 */
        var tipPointBeforeShorten = options.hasTipMarker ? getTipPoint(points, geometry.hDir).slice(0) : null;
        var bendPt = options.useArrowMarker ? points[1].slice(0) : null;

        if (options.useCircleMarker && !options.fillTipMarker) {
            /* 円の半径 + 円の線幅の半分で短縮（線が円の縁に接する） */
            shortenTip(points, geometry.hDir, options.markerSizePt / 2 + style.widthPt / 2);
        }
        if (options.useArrowMarker) {
            /* 矢印の長さ分だけ短縮 */
            shortenTip(points, geometry.hDir, options.markerSizePt);
        }

        var parts = {
            edgePath: null,
            mainPath: null,
            edgeTipMarker: null,
            mainTipMarker: null,
            hDir: geometry.hDir,
            vDir: geometry.vDir
        };

        /* フチ（最背面） */
        if (options.hasEdge) {
            parts.edgePath = doc.pathItems.add();
            createdArtItems.push(parts.edgePath);
            parts.edgePath.setEntirePath(points);
            parts.edgePath.stroked = true;
            parts.edgePath.strokeWidth = style.widthPt * EDGE_LINE_WIDTH_RATIO;
            parts.edgePath.strokeColor = style.edgeColor;
            parts.edgePath.filled = false;
            parts.edgePath.strokeCap = options.roundStrokeCap ? StrokeCap.ROUNDENDCAP : StrokeCap.BUTTENDCAP;
        }

        parts.mainPath = doc.pathItems.add();
        createdArtItems.push(parts.mainPath);
        parts.mainPath.setEntirePath(points);
        parts.mainPath.stroked = true;
        parts.mainPath.strokeWidth = style.widthPt;
        parts.mainPath.strokeColor = style.lineColor;
        parts.mainPath.filled = false;
        parts.mainPath.strokeCap = options.roundStrokeCap ? StrokeCap.ROUNDENDCAP : StrokeCap.BUTTENDCAP;

        if (options.useCircleMarker) {
            if (options.hasEdge) parts.edgeTipMarker = addTipCircleMarkerEdge(doc, tipPointBeforeShorten, style, options.markerSizePt);
            parts.mainTipMarker = addTipCircleMarker(doc, tipPointBeforeShorten, style, options.markerSizePt, options.fillTipMarker);
        }
        if (options.useArrowMarker) {
            if (options.hasEdge) parts.edgeTipMarker = addTipArrowEdge(doc, tipPointBeforeShorten, bendPt, style, options.markerSizePt);
            parts.mainTipMarker = addTipArrow(doc, tipPointBeforeShorten, bendPt, style, options.markerSizePt, options.fillTipMarker);
        }
        if (parts.edgeTipMarker) createdArtItems.push(parts.edgeTipMarker);
        if (parts.mainTipMarker) createdArtItems.push(parts.mainTipMarker);

        return parts;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ↑↓キーで入力欄の値を増減できるようにする
     * @param {EditText} editText - 対象の入力欄
     * @param {object} [options] - 増減幅などのオプション
     * @returns {void}
     */
    function changeValueByArrowKey(editText, options) {
        options = options || {};
        var smallStep = (typeof options.smallStep === "number") ? options.smallStep : 1;
        var largeStep = (typeof options.largeStep === "number") ? options.largeStep : 10;
        var fineStep = (typeof options.fineStep === "number") ? options.fineStep : 0.1;
        var minValue = (typeof options.minValue === "number") ? options.minValue : 0;
        var digits = (typeof options.digits === "number") ? options.digits : 0;
        var onAfterChange = (typeof options.onAfterChange === "function") ? options.onAfterChange : null;

        editText.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;

            var value = parseFloat(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var step = smallStep;
            if (keyboard.shiftKey) {
                step = largeStep;
            } else if (keyboard.altKey) {
                step = fineStep;
            }

            if (keyboard.shiftKey) {
                /* Shift：largeStep の倍数にスナップ */
                if (event.keyName === "Up") {
                    var ceilValue = Math.ceil(value / step * (1 + 1e-9)) * step;
                    value = (ceilValue <= value) ? value + step : ceilValue;
                } else {
                    var floorValue = Math.floor(value / step * (1 - 1e-9)) * step;
                    value = (floorValue >= value) ? value - step : floorValue;
                    if (value < minValue) value = minValue;
                }
            } else if (event.keyName === "Up") {
                value += step;
            } else {
                value -= step;
                if (value < minValue) value = minValue;
            }

            editText.text = formatNumber(value, digits);
            event.preventDefault();
            if (onAfterChange) onAfterChange();
        });
    }

    /**
     * 適用範囲パネルを追加する
     * @param {Window|Group|Panel} parent - 追加先
     * @param {object} ui - コントロールを格納するオブジェクト
     * @returns {void}
     */
    function addApplyScopePanel(parent, ui) {
        ui.applyScopePanel = addLabeledPanel(parent, getLabel("panel", "applyScope"));
        setupRowGroup(ui.applyScopePanel, ["center", "center"]);
        ui.applyScopeAllRadio = ui.applyScopePanel.add("radiobutton", undefined, getLabel("radio", "applyScopeAll"));
        ui.applyScopeAllRadio.helpTip = getLabel("tooltip", "applyScopeAll");
        ui.applyScopeKeepDirRadio = ui.applyScopePanel.add("radiobutton", undefined, getLabel("radio", "applyScopeKeepDir"));
        ui.applyScopeKeepDirRadio.helpTip = getLabel("tooltip", "applyScopeKeepDir");
    }

    /**
     * 角度パネルを追加する
     * @param {Window|Group|Panel} parent - 追加先
     * @param {object} ui - コントロールを格納するオブジェクト
     * @returns {void}
     */
    function addAnglePanel(parent, ui) {
        ui.anglePanel = addLabeledPanel(parent, getLabel("panel", "angle"));

        ui.angleInputRow = ui.anglePanel.add("group");
        ui.angleInputRow.alignment = ["center", "top"];
        ui.angleInput = ui.angleInputRow.add("edittext", undefined, "");
        ui.angleInput.characters = 4;
        ui.angleInput.helpTip = getLabel("tooltip", "angleInput");
        ui.angleInputRow.add("statictext", undefined, "\u00B0");
        ui.angleInput.active = true;

        ui.anglePresetRow = ui.anglePanel.add("group");
        setupRowGroup(ui.anglePresetRow);
        ui.anglePreset30Radio = ui.anglePresetRow.add("radiobutton", undefined, "30\u00B0");
        ui.anglePreset45Radio = ui.anglePresetRow.add("radiobutton", undefined, "45\u00B0");
        ui.anglePreset60Radio = ui.anglePresetRow.add("radiobutton", undefined, "60\u00B0");
    }

    /**
     * 斜線の方向パネルを追加する
     * 左右2列のラジオの間に、対象を示す中央列を挟む
     * @param {Window|Group|Panel} parent - 追加先
     * @param {object} ui - コントロールを格納するオブジェクト
     * @returns {void}
     */
    function addDiagDirPanel(parent, ui) {
        ui.diagDirPanel = addLabeledPanel(parent, getLabel("panel", "direction"));

        ui.diagDirRow = ui.diagDirPanel.add("group");
        setupRowGroup(ui.diagDirRow, ["fill", "top"]);
        ui.diagDirRow.helpTip = getLabel("tooltip", "diagDir");

        ui.diagDirLeftColumn = ui.diagDirRow.add("group");
        ui.diagDirLeftColumn.orientation = "column";
        ui.diagDirLeftColumn.alignChildren = ["fill", "top"];
        ui.diagDirUpperLeftRadio = ui.diagDirLeftColumn.add("radiobutton", undefined, getLabel("radio", "diagDirUpperLeft"));
        ui.diagDirLeftColumn.add("statictext", undefined, " ");
        ui.diagDirLowerLeftRadio = ui.diagDirLeftColumn.add("radiobutton", undefined, getLabel("radio", "diagDirLowerLeft"));

        ui.diagDirCenterColumn = ui.diagDirRow.add("group");
        ui.diagDirCenterColumn.orientation = "column";
        ui.diagDirCenterColumn.alignChildren = ["center", "center"];
        ui.diagDirCenterColumn.add("statictext", undefined, " ");
        ui.diagDirCenterColumn.add("statictext", undefined, getLabel("fieldLabel", "targetMark"));

        ui.diagDirRightColumn = ui.diagDirRow.add("group");
        ui.diagDirRightColumn.orientation = "column";
        ui.diagDirRightColumn.alignChildren = ["fill", "top"];
        ui.diagDirUpperRightRadio = ui.diagDirRightColumn.add("radiobutton", undefined, getLabel("radio", "diagDirUpperRight"));
        ui.diagDirRightColumn.add("statictext", undefined, " ");
        ui.diagDirLowerRightRadio = ui.diagDirRightColumn.add("radiobutton", undefined, getLabel("radio", "diagDirLowerRight"));
    }

    /**
     * 線のスタイルパネル（色・線幅・線端の形状）を追加する
     * @param {Window|Group|Panel} parent - 追加先
     * @param {object} ui - コントロールを格納するオブジェクト
     * @param {string} strokeUnitLabel - 線幅の単位ラベル
     * @returns {void}
     */
    function addLineStylePanel(parent, ui, strokeUnitLabel) {
        ui.lineStylePanel = addLabeledPanel(parent, getLabel("panel", "lineStyle"));

        ui.lineColorRow = ui.lineStylePanel.add("group");
        setupRowGroup(ui.lineColorRow);
        ui.lineColorBlackRadio = ui.lineColorRow.add("radiobutton", undefined, getLabel("radio", "lineColorBlack"));
        ui.lineColorWhiteRadio = ui.lineColorRow.add("radiobutton", undefined, getLabel("radio", "lineColorWhite"));
        ui.lineColorCustomRadio = ui.lineColorRow.add("radiobutton", undefined, getLabel("radio", "colorCustom"));
        /* HEX値はコントロールではなく、text プロパティだけを持つ保持箱に入れる */
        ui.lineColorHexHolder = { text: sessionSettings.lineColorHex };
        ui.lineColorChip = ui.lineColorRow.add("panel", undefined, "");
        ui.lineColorChip.preferredSize = COLOR_CHIP_SIZE;
        ui.lineColorChip.helpTip = getLabel("tooltip", "colorChip");

        ui.lineWidthRow = ui.lineStylePanel.add("group");
        setupRowGroup(ui.lineWidthRow);
        ui.lineWidthRow.add("statictext", undefined, getLabel("fieldLabel", "lineWidth"));
        ui.lineWidthInput = ui.lineWidthRow.add("edittext", undefined, "");
        ui.lineWidthInput.characters = 3;
        ui.lineWidthInput.helpTip = getLabel("tooltip", "lengthInput");
        ui.lineWidthRow.add("statictext", undefined, strokeUnitLabel);

        ui.strokeCapRow = ui.lineStylePanel.add("group");
        setupRowGroup(ui.strokeCapRow);
        ui.strokeCapRow.helpTip = getLabel("tooltip", "strokeCap");
        ui.strokeCapRow.add("statictext", undefined, getLabel("fieldLabel", "strokeCap"));
        ui.strokeCapButtRadio = ui.strokeCapRow.add("radiobutton", undefined, getLabel("radio", "strokeCapButt"));
        ui.strokeCapRoundRadio = ui.strokeCapRow.add("radiobutton", undefined, getLabel("radio", "strokeCapRound"));
    }

    /**
     * 線端パネル（円・矢印の種類と大きさ）を追加する
     * @param {Window|Group|Panel} parent - 追加先
     * @param {object} ui - コントロールを格納するオブジェクト
     * @param {string} strokeUnitLabel - 大きさの単位ラベル
     * @returns {void}
     */
    function addTipMarkerPanel(parent, ui, strokeUnitLabel) {
        ui.tipMarkerPanel = addLabeledPanel(parent, getLabel("panel", "tipMarker"));

        ui.tipMarkerTypeRow = ui.tipMarkerPanel.add("group");
        setupRowGroup(ui.tipMarkerTypeRow);
        ui.tipMarkerNoneRadio = ui.tipMarkerTypeRow.add("radiobutton", undefined, getLabel("radio", "tipMarkerNone"));
        ui.tipMarkerCircleRadio = ui.tipMarkerTypeRow.add("radiobutton", undefined, getLabel("radio", "tipMarkerCircle"));
        ui.tipMarkerArrowRadio = ui.tipMarkerTypeRow.add("radiobutton", undefined, getLabel("radio", "tipMarkerArrow"));

        ui.tipMarkerStyleRow = ui.tipMarkerPanel.add("group");
        setupRowGroup(ui.tipMarkerStyleRow);
        ui.tipMarkerStyleRow.helpTip = getLabel("tooltip", "tipMarkerStyle");
        ui.tipMarkerFillRadio = ui.tipMarkerStyleRow.add("radiobutton", undefined, getLabel("radio", "tipMarkerFill"));
        ui.tipMarkerOutlineRadio = ui.tipMarkerStyleRow.add("radiobutton", undefined, getLabel("radio", "tipMarkerOutline"));

        ui.tipMarkerSizeRow = ui.tipMarkerPanel.add("group");
        setupRowGroup(ui.tipMarkerSizeRow);
        ui.tipMarkerSizeRow.add("statictext", undefined, getLabel("fieldLabel", "tipMarkerSize"));
        ui.tipMarkerSizeInput = ui.tipMarkerSizeRow.add("edittext", undefined, "");
        ui.tipMarkerSizeInput.characters = 3;
        ui.tipMarkerSizeInput.helpTip = getLabel("tooltip", "tipMarkerSize");
        ui.tipMarkerSizeRow.add("statictext", undefined, strokeUnitLabel);

        ui.groupItemsCheck = ui.tipMarkerPanel.add("checkbox", undefined, getLabel("checkbox", "groupEnabled"));
        ui.groupItemsCheck.alignment = "left";
        ui.groupItemsCheck.helpTip = getLabel("tooltip", "groupEnabled");
    }

    /**
     * テキストパネル（テキストとの距離）を追加する
     * @param {Window|Group|Panel} parent - 追加先
     * @param {object} ui - コントロールを格納するオブジェクト
     * @returns {void}
     */
    function addTextAlignPanel(parent, ui) {
        ui.textAlignPanel = addLabeledPanel(parent, getLabel("panel", "text"));

        ui.textDistRow = ui.textAlignPanel.add("group");
        setupRowGroup(ui.textDistRow);
        ui.textDistRow.add("statictext", undefined, getLabel("fieldLabel", "textDist"));
        ui.textDistInput = ui.textDistRow.add("edittext", undefined, "");
        ui.textDistInput.characters = 4;
        ui.textDistInput.helpTip = getLabel("tooltip", "textDist");
        ui.textDistRow.add("statictext", undefined, "pt");
    }

    /**
     * フチパネル（フチの有無と色）を追加する
     * @param {Window|Group|Panel} parent - 追加先
     * @param {object} ui - コントロールを格納するオブジェクト
     * @returns {void}
     */
    function addEdgePanel(parent, ui) {
        ui.edgePanel = addLabeledPanel(parent, getLabel("panel", "edge"));

        ui.edgeEnabledCheck = ui.edgePanel.add("checkbox", undefined, getLabel("checkbox", "edgeEnabled"));
        ui.edgeEnabledCheck.alignment = "left";
        ui.edgeEnabledCheck.helpTip = getLabel("tooltip", "edgeEnabled");

        ui.edgeColorRow = ui.edgePanel.add("group");
        setupRowGroup(ui.edgeColorRow);
        ui.edgeColorWhiteRadio = ui.edgeColorRow.add("radiobutton", undefined, getLabel("radio", "lineColorWhite"));
        ui.edgeColorCustomRadio = ui.edgeColorRow.add("radiobutton", undefined, getLabel("radio", "colorCustom"));
        ui.edgeColorHexHolder = { text: sessionSettings.edgeColorHex };
        ui.edgeColorChip = ui.edgeColorRow.add("panel", undefined, "");
        ui.edgeColorChip.preferredSize = COLOR_CHIP_SIZE;
        ui.edgeColorChip.helpTip = getLabel("tooltip", "colorChip");
    }

    /**
     * 設定キーとラジオボタンの対応表を作る
     * @param {object} ui - buildDialogUI() の戻り値
     * @returns {object} カテゴリごとの対応表
     */
    function buildRadioMaps(ui) {
        return {
            presetAngle: [
                { key: "30", radio: ui.anglePreset30Radio },
                { key: "45", radio: ui.anglePreset45Radio },
                { key: "60", radio: ui.anglePreset60Radio }
            ],
            applyScope: [
                { key: "all",             radio: ui.applyScopeAllRadio },
                { key: "exceptDirection", radio: ui.applyScopeKeepDirRadio }
            ],
            diagDir: [
                { key: "upperLeft",  radio: ui.diagDirUpperLeftRadio },
                { key: "lowerLeft",  radio: ui.diagDirLowerLeftRadio },
                { key: "upperRight", radio: ui.diagDirUpperRightRadio },
                { key: "lowerRight", radio: ui.diagDirLowerRightRadio }
            ],
            lineColor: [
                { key: "black", radio: ui.lineColorBlackRadio },
                { key: "white", radio: ui.lineColorWhiteRadio },
                { key: "other", radio: ui.lineColorCustomRadio }
            ],
            strokeCap: [
                { key: "none",  radio: ui.strokeCapButtRadio },
                { key: "round", radio: ui.strokeCapRoundRadio }
            ],
            tipMarkerType: [
                { key: "none",   radio: ui.tipMarkerNoneRadio },
                { key: "circle", radio: ui.tipMarkerCircleRadio },
                { key: "arrow",  radio: ui.tipMarkerArrowRadio }
            ],
            tipMarkerStyle: [
                { key: "fill",   radio: ui.tipMarkerFillRadio },
                { key: "stroke", radio: ui.tipMarkerOutlineRadio }
            ],
            edgeColor: [
                { key: "white", radio: ui.edgeColorWhiteRadio },
                { key: "other", radio: ui.edgeColorCustomRadio }
            ]
        };
    }

    /**
     * ダイアログを組み立てる
     * @returns {object} 各コントロールをまとめたオブジェクト
     */
    function buildDialogUI() {
        var ui = {};
        var strokeUnitLabel = getCurrentUnitLabel("strokeUnits");

        ui.dialogWindow = new Window("dialog", getLabel("dialog", "title") + " " + SCRIPT_VERSION);
        setupWindow(ui.dialogWindow);
        /* ダイアログ表示位置は復元する */
        if (sessionSettings.dialogBounds && sessionSettings.dialogBounds.length === 2) {
            ui.dialogWindow.location = sessionSettings.dialogBounds;
        }

        addApplyScopePanel(ui.dialogWindow, ui);

        ui.columnsRow = ui.dialogWindow.add("group");
        setupRowGroup(ui.columnsRow, ["fill", "top"], COLUMN_SPACING);

        ui.leftColumn = ui.columnsRow.add("group");
        ui.leftColumn.orientation = "column";
        ui.leftColumn.alignChildren = ["fill", "top"];

        addAnglePanel(ui.leftColumn, ui);
        addDiagDirPanel(ui.leftColumn, ui);

        ui.rightColumn = ui.columnsRow.add("group");
        ui.rightColumn.orientation = "column";
        ui.rightColumn.alignChildren = ["fill", "top"];

        addLineStylePanel(ui.rightColumn, ui, strokeUnitLabel);
        addTipMarkerPanel(ui.rightColumn, ui, strokeUnitLabel);
        addTextAlignPanel(ui.rightColumn, ui);
        addEdgePanel(ui.leftColumn, ui);

        ui.radioMaps = buildRadioMaps(ui);

        return ui;
    }

    /**
     * 前回の設定をダイアログへ反映する
     * @param {object} ui - buildDialogUI() の戻り値
     * @param {Array<PathItem|GroupItem>} targetItems - 対象アイテム
     * @param {Array<TextFrame>} textFrameTargets - 選択中のテキストフレーム
     * @returns {void}
     */
    function loadDialogSettingsToUI(ui, targetItems, textFrameTargets) {
        ui.angleInput.text = String(sessionSettings.angle);
        selectRadioByKey(ui.radioMaps.presetAngle, String(sessionSettings.radioAngle), "45");

        /* 複数選択で未設定のときは「斜線の方向以外」を初期選択にする */
        var applyScope = sessionSettings.applyScope;
        if (!sessionSettings.hasUserSetApplyScope && targetItems.length > 1) applyScope = "exceptDirection";
        selectRadioByKey(ui.radioMaps.applyScope, applyScope, "all");

        selectRadioByKey(ui.radioMaps.diagDir, sessionSettings.diagDir, "upperLeft");
        selectRadioByKey(ui.radioMaps.lineColor, sessionSettings.lineColor, "black");
        selectRadioByKey(ui.radioMaps.strokeCap, sessionSettings.strokeCapType, "round");
        selectRadioByKey(ui.radioMaps.tipMarkerType, sessionSettings.tipMarkerType, "none");
        selectRadioByKey(ui.radioMaps.tipMarkerStyle, sessionSettings.tipMarkerStyle, "fill");
        selectRadioByKey(ui.radioMaps.edgeColor, sessionSettings.edgeColor, "white");

        ui.lineColorHexHolder.text = sessionSettings.lineColorHex;
        updateColorChip(ui.lineColorChip, ui.lineColorHexHolder.text);
        ui.edgeColorHexHolder.text = sessionSettings.edgeColorHex;
        updateColorChip(ui.edgeColorChip, ui.edgeColorHexHolder.text);

        ui.lineWidthInput.text = formatUnitInput(sessionSettings.lineWidth, DEFAULT_LINE_WIDTH_PT);
        ui.tipMarkerSizeInput.text = formatUnitInput(sessionSettings.tipMarkerSize, DEFAULT_TIP_MARKER_SIZE_PT);

        ui.groupItemsCheck.value = !!sessionSettings.groupEnabled;
        ui.edgeEnabledCheck.value = !!sessionSettings.whiteEdge;

        var textDistValue = parseFloat(sessionSettings.textDist);
        if (isNaN(textDistValue) || textDistValue < 0) textDistValue = DEFAULT_TEXT_DIST_PT;
        ui.textDistInput.text = formatNumber(textDistValue, 2);

        ui.textAlignPanel.enabled = textFrameTargets.length > 0;

        updateDirPanelEnabled(ui);
        updateTipMarkerControlsEnabled(ui);
        updateEdgeColorEnabled(ui);
    }

    /**
     * ダイアログの設定をセッション記憶へ保存する
     * @param {object} ui - buildDialogUI() の戻り値
     * @returns {void}
     */
    function saveDialogSettingsFromUI(ui) {
        var diagDirection = getDiagDirValues(ui);

        sessionSettings.angle = ui.angleInput.text;
        sessionSettings.radioAngle = parseFloat(getSelectedRadioKey(ui.radioMaps.presetAngle, "45"));
        sessionSettings.applyScope = getSelectedRadioKey(ui.radioMaps.applyScope, "all");
        sessionSettings.diagDir = getSelectedRadioKey(ui.radioMaps.diagDir, "lowerRight");
        sessionSettings.hDir = diagDirection.hDir;
        sessionSettings.vDir = diagDirection.vDir;

        sessionSettings.tipMarkerType = getSelectedRadioKey(ui.radioMaps.tipMarkerType, "none");
        sessionSettings.tipMarkerStyle = getSelectedRadioKey(ui.radioMaps.tipMarkerStyle, "fill");
        sessionSettings.tipMarkerSize = parseUnitInput(ui.tipMarkerSizeInput.text, DEFAULT_TIP_MARKER_SIZE_PT);
        sessionSettings.strokeCapType = getSelectedRadioKey(ui.radioMaps.strokeCap, "round");

        sessionSettings.groupEnabled = ui.groupItemsCheck.value;
        sessionSettings.whiteEdge = ui.edgeEnabledCheck.value;
        sessionSettings.edgeColor = getSelectedRadioKey(ui.radioMaps.edgeColor, "white");
        sessionSettings.edgeColorHex = ui.edgeColorHexHolder.text;

        sessionSettings.lineColor = getSelectedRadioKey(ui.radioMaps.lineColor, "black");
        sessionSettings.lineColorHex = ui.lineColorHexHolder.text;
        sessionSettings.lineWidth = parseUnitInput(ui.lineWidthInput.text, DEFAULT_LINE_WIDTH_PT);

        var savedTextDist = parseFloat(ui.textDistInput.text);
        sessionSettings.textDist = (!isNaN(savedTextDist) && savedTextDist >= 0) ? formatNumber(savedTextDist, 4) : DEFAULT_TEXT_DIST_PT;

        rememberDialogLocation(ui);
    }

    /**
     * ズーム状態は保存しないが、ダイアログの表示位置は保存する
     * @param {object} ui - buildDialogUI() の戻り値
     * @returns {void}
     */
    function rememberDialogLocation(ui) {
        if (ui && ui.dialogWindow && ui.dialogWindow.location) {
            sessionSettings.dialogBounds = [ui.dialogWindow.location[0], ui.dialogWindow.location[1]];
        }
    }

    /**
     * 適用範囲に応じて斜線の方向パネルの有効・無効を切り替える
     * @param {object} ui - buildDialogUI() の戻り値
     * @returns {void}
     */
    function updateDirPanelEnabled(ui) {
        ui.diagDirPanel.enabled = ui.applyScopeAllRadio.value;
    }

    /**
     * 線端の選択に応じて関連コントロールの有効・無効を切り替える
     * @param {object} ui - buildDialogUI() の戻り値
     * @returns {void}
     */
    function updateTipMarkerControlsEnabled(ui) {
        var hasCap = ui.tipMarkerCircleRadio.value || ui.tipMarkerArrowRadio.value;
        if (ui.tipMarkerArrowRadio.value) {
            /* 矢印は塗りのみ */
            if (ui.tipMarkerOutlineRadio.value) {
                ui.tipMarkerOutlineRadio.value = false;
                ui.tipMarkerFillRadio.value = true;
            }
            ui.tipMarkerFillRadio.enabled = false;
            ui.tipMarkerOutlineRadio.enabled = false;
        } else {
            ui.tipMarkerFillRadio.enabled = hasCap;
            ui.tipMarkerOutlineRadio.enabled = hasCap;
        }
        ui.tipMarkerSizeInput.enabled = hasCap;
        /* 線端がなくても、フチONなら本体とフチの2本になるためグループ化が意味を持つ */
        ui.groupItemsCheck.enabled = hasCap || ui.edgeEnabledCheck.value;
    }

    /**
     * フチのON/OFFに応じてフチ色のコントロールの有効・無効を切り替える
     * @param {object} ui - buildDialogUI() の戻り値
     * @returns {void}
     */
    function updateEdgeColorEnabled(ui) {
        var edgeEnabled = ui.edgeEnabledCheck.value;
        ui.edgeColorRow.enabled = edgeEnabled;
        ui.edgeColorChip.enabled = edgeEnabled;
    }

    /**
     * カラーピッカーを開き、選ばれた色をHEX保持欄とスウォッチに反映する
     * @param {object} hexHolder - HEX値を保持するオブジェクト
     * @param {Panel} colorChip - 対象のスウォッチ
     * @returns {boolean} 色が選ばれたら true
     */
    function pickColor(hexHolder, colorChip) {
        var pickedHex = ColorPicker.show(hexHolder.text.replace(/^#/, ""));
        if (!pickedHex) return false;
        hexHolder.text = "#" + pickedHex;
        updateColorChip(colorChip, hexHolder.text);
        return true;
    }

    /**
     * 数値入力欄のイベントを登録する（↑↓キーと直接入力）
     * @param {object} ui - buildDialogUI() の戻り値
     * @param {function} onChange - 設定が変わったときに呼ぶ関数
     * @returns {void}
     */
    function bindNumberInputEvents(ui, onChange) {
        var afterChange = function () { onChange(); };

        changeValueByArrowKey(ui.angleInput, {
            smallStep: 1,
            largeStep: 10,
            fineStep: 0.1,
            minValue: 0,
            digits: 0,
            onAfterChange: afterChange
        });

        var lengthInputs = [ui.lineWidthInput, ui.tipMarkerSizeInput, ui.textDistInput];
        for (var i = 0; i < lengthInputs.length; i++) {
            changeValueByArrowKey(lengthInputs[i], {
                smallStep: 0.1,
                largeStep: 1,
                fineStep: 0.01,
                minValue: 0,
                digits: 2,
                onAfterChange: afterChange
            });
            lengthInputs[i].onChanging = afterChange;
        }

        ui.angleInput.onChanging = afterChange;
        ui.anglePreset30Radio.onClick = function () { ui.angleInput.text = "30"; onChange(); };
        ui.anglePreset45Radio.onClick = function () { ui.angleInput.text = "45"; onChange(); };
        ui.anglePreset60Radio.onClick = function () { ui.angleInput.text = "60"; onChange(); };
    }

    /**
     * 適用範囲と斜線の方向のイベントを登録する
     * @param {object} ui - buildDialogUI() の戻り値
     * @param {function} onChange - 設定が変わったときに呼ぶ関数
     * @returns {void}
     */
    function bindDiagDirEvents(ui, onChange) {
        /* 適用範囲はユーザーが明示的にクリックしたときだけ「手動設定済み」とみなす */
        var onScopeClick = function () {
            sessionSettings.hasUserSetApplyScope = true;
            updateDirPanelEnabled(ui);
            onChange();
        };
        ui.applyScopeAllRadio.onClick = onScopeClick;
        ui.applyScopeKeepDirRadio.onClick = onScopeClick;

        /* 方向ラジオは列をまたぐため自動排他が効かない。クリック時に自分だけを選択状態にする */
        for (var i = 0; i < ui.radioMaps.diagDir.length; i++) {
            (function (entry) {
                entry.radio.onClick = function () {
                    selectRadioByKey(ui.radioMaps.diagDir, entry.key, entry.key);
                    onChange();
                };
            })(ui.radioMaps.diagDir[i]);
        }
    }

    /**
     * 線のスタイル（色と線端の形状）のイベントを登録する
     * @param {object} ui - buildDialogUI() の戻り値
     * @param {function} onChange - 設定が変わったときに呼ぶ関数
     * @returns {void}
     */
    function bindLineStyleEvents(ui, onChange) {
        ui.lineColorChip.addEventListener("click", function () {
            if (!pickColor(ui.lineColorHexHolder, ui.lineColorChip)) return;
            selectRadioByKey(ui.radioMaps.lineColor, "other", "other");
            onChange();
        });
        ui.lineColorBlackRadio.onClick = function () { onChange(); };
        ui.lineColorWhiteRadio.onClick = function () { onChange(); };
        ui.lineColorCustomRadio.onClick = function () {
            pickColor(ui.lineColorHexHolder, ui.lineColorChip);
            onChange();
        };

        ui.strokeCapButtRadio.onClick = function () { onChange(); };
        ui.strokeCapRoundRadio.onClick = function () { onChange(); };
    }

    /**
     * 線端のイベントを登録する
     * @param {object} ui - buildDialogUI() の戻り値
     * @param {function} onChange - 設定が変わったときに呼ぶ関数
     * @returns {void}
     */
    function bindTipMarkerEvents(ui, onChange) {
        var onTypeClick = function () { updateTipMarkerControlsEnabled(ui); onChange(); };
        ui.tipMarkerNoneRadio.onClick = onTypeClick;
        ui.tipMarkerCircleRadio.onClick = onTypeClick;
        ui.tipMarkerArrowRadio.onClick = onTypeClick;

        ui.tipMarkerFillRadio.onClick = function () { onChange(); };
        ui.tipMarkerOutlineRadio.onClick = function () { onChange(); };
        ui.groupItemsCheck.onClick = function () { saveDialogSettingsFromUI(ui); };
    }

    /**
     * フチのイベントを登録する
     * @param {object} ui - buildDialogUI() の戻り値
     * @param {function} onChange - 設定が変わったときに呼ぶ関数
     * @returns {void}
     */
    function bindEdgeEvents(ui, onChange) {
        ui.edgeEnabledCheck.onClick = function () {
            updateEdgeColorEnabled(ui);
            updateTipMarkerControlsEnabled(ui);
            onChange();
        };
        ui.edgeColorChip.addEventListener("click", function () {
            if (!pickColor(ui.edgeColorHexHolder, ui.edgeColorChip)) return;
            selectRadioByKey(ui.radioMaps.edgeColor, "other", "other");
            onChange();
        });
        ui.edgeColorWhiteRadio.onClick = function () { onChange(); };
        ui.edgeColorCustomRadio.onClick = function () {
            pickColor(ui.edgeColorHexHolder, ui.edgeColorChip);
            onChange();
        };
    }

    /**
     * ダイアログのイベントを登録する
     * @param {object} ui - buildDialogUI() の戻り値
     * @param {function} onChange - 設定が変わったときに呼ぶ関数（プレビュー更新）
     * @returns {void}
     */
    function bindDialogEvents(ui, onChange) {
        bindNumberInputEvents(ui, onChange);
        bindDiagDirEvents(ui, onChange);
        bindLineStyleEvents(ui, onChange);
        bindTipMarkerEvents(ui, onChange);
        bindEdgeEvents(ui, onChange);

        ui.dialogWindow.onMove = function () {
            rememberDialogLocation(ui);
        };
        ui.dialogWindow.onClose = function () {
            saveDialogSettingsFromUI(ui);
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択から引き出し線の対象とテキストフレームを集める
     * PathItem は2点以上、GroupItem はそのまま対象とする
     * @param {object} selectedItems - ドキュメントの選択
     * @returns {object} leaderTargets / textTargets を持つオブジェクト
     */
    function collectTargets(selectedItems) {
        var leaderTargets = [];
        var textTargets = [];

        for (var i = 0; i < selectedItems.length; i++) {
            var itemType = selectedItems[i].typename;
            if (itemType === "PathItem" && selectedItems[i].pathPoints.length >= 2) {
                leaderTargets.push(selectedItems[i]);
            } else if (itemType === "GroupItem") {
                leaderTargets.push(selectedItems[i]);
            } else if (itemType === "TextFrame") {
                textTargets.push(selectedItems[i]);
            }
        }

        return { leaderTargets: leaderTargets, textTargets: textTargets };
    }

    /**
     * 確定した設定で引き出し線を生成し、元オブジェクトと置き換える
     * @param {Document} doc - 対象ドキュメント
     * @param {object} ui - buildDialogUI() の戻り値
     * @param {Array<PathItem|GroupItem>} targetItems - 対象アイテム
     * @param {Array<TextFrame>} textFrameTargets - 選択中のテキストフレーム
     * @param {number} angleRad - 斜線の角度（ラジアン）
     * @returns {void}
     */
    function applyLeaderLines(doc, ui, targetItems, textFrameTargets, angleRad) {
        var leaderOptions = readLeaderOptions(ui, textFrameTargets);
        var targetMetricsList = measureTargets(targetItems);
        var itemsToSelect = [];
        var i;
        doc.selection = null;

        for (i = 0; i < targetItems.length; i++) {
            var sourceTarget = targetItems[i];
            var targetMetrics = targetMetricsList[i];
            var createdArtItems = [];
            switchToEditableLayer(doc, sourceTarget);

            try {
                var parts = buildLeaderLineParts(doc, sourceTarget, targetMetrics, angleRad, leaderOptions, createdArtItems);

                if (ui.groupItemsCheck.value) {
                    var leaderGroup = doc.groupItems.add();
                    createdArtItems.push(leaderGroup);
                    var assembled = assembleLeaderParts(leaderGroup, parts.edgePath, parts.mainPath, parts.edgeTipMarker, parts.mainTipMarker, leaderOptions.hasEdge);
                    tagLeaderParts(assembled, parts.edgePath, parts.mainPath, parts.edgeTipMarker, parts.mainTipMarker, leaderOptions.hasEdge);
                    setLeaderLineTag(leaderGroup);
                    setLeaderLineBoundsTags(leaderGroup, targetMetrics);
                    setLeaderLineDirTags(leaderGroup, parts.hDir, parts.vDir);
                    itemsToSelect.push(leaderGroup);
                } else {
                    if (parts.edgePath) itemsToSelect.push(parts.edgePath);
                    if (parts.edgeTipMarker) itemsToSelect.push(parts.edgeTipMarker);
                    if (parts.mainPath) itemsToSelect.push(parts.mainPath);
                    if (parts.mainTipMarker) itemsToSelect.push(parts.mainTipMarker);
                }

                /* 置き換え成功時のみ元オブジェクトを削除 */
                sourceTarget.remove();
            } catch (err) {
                for (var k = createdArtItems.length - 1; k >= 0; k--) {
                    safeRemove(createdArtItems[k]);
                }
                throw err;
            }
        }

        for (i = 0; i < itemsToSelect.length; i++) {
            safeSelect(itemsToSelect[i]);
        }
    }

    /**
     * 選択オブジェクトから引き出し線を作成する
     * @returns {void}
     */
    function main() {
        /* ドキュメントが開かれているか確認 / Check for an open document */
        if (app.documents.length === 0) {
            alert(getLabel("alert", "noDocument"));
            return;
        }

        var doc = app.activeDocument;

        /* 選択オブジェクトのチェック / Check the selection */
        if (doc.selection.length === 0) {
            alert(getLabel("alert", "noSelection"));
            return;
        }

        var targets = collectTargets(doc.selection);
        var targetItems = targets.leaderTargets;
        var textFrameTargets = targets.textTargets;

        if (targetItems.length === 0) {
            alert(getLabel("alert", "noValidTargets"));
            return;
        }

        /* プレビュー用アイテムの配列（親グループ単位で管理） */
        var previewGroups = [];

        /**
         * プレビューを生成する
         * @param {number} angleDeg - 斜線の角度（度）
         * @param {object} ui - buildDialogUI() の戻り値
         * @returns {void}
         */
        function createPreview(angleDeg, ui) {
            var angleRad = angleDeg * Math.PI / 180;
            var targetMetricsList = measureTargets(targetItems);
            var options = readLeaderOptions(ui, textFrameTargets);

            for (var i = 0; i < targetMetricsList.length; i++) {
                switchToEditableLayer(doc, targetItems[i]);
                var parts = buildLeaderLineParts(doc, targetItems[i], targetMetricsList[i], angleRad, options);
                var previewGroup = doc.groupItems.add();
                assembleLeaderParts(previewGroup, parts.edgePath, parts.mainPath, parts.edgeTipMarker, parts.mainTipMarker, options.hasEdge);
                previewGroups.push(previewGroup);
            }
        }

        /**
         * プレビューを削除する
         * @returns {void}
         */
        function removePreview() {
            for (var i = previewGroups.length - 1; i >= 0; i--) {
                safeRemove(previewGroups[i]);
            }
            previewGroups = [];
        }

        /**
         * 対象オブジェクトの表示・非表示を切り替える
         * @param {boolean} visible - 表示するなら true
         * @returns {void}
         */
        function setTargetsVisible(visible) {
            for (var i = 0; i < targetItems.length; i++) {
                /* ロックされた対象では失敗するが、プレビューは続行する */
                try {
                    targetItems[i].hidden = !visible;
                } catch (e) { }
            }
        }

        var initialViewState = captureViewState(doc);

        /**
         * 入力値でプレビューを作り直す
         * @param {object} ui - buildDialogUI() の戻り値
         * @returns {void}
         */
        function updatePreview(ui) {
            removePreview();
            setTargetsVisible(true);

            var angleDeg = parseFloat(ui.angleInput.text);
            if (isNaN(angleDeg) || angleDeg <= 0 || angleDeg >= 90) {
                saveDialogSettingsFromUI(ui);
                app.redraw();
                return;
            }

            createPreview(angleDeg, ui);
            setTargetsVisible(false);
            saveDialogSettingsFromUI(ui);
            app.redraw();
        }

        var dialogUI = buildDialogUI();
        bindDialogEvents(dialogUI, function () { updatePreview(dialogUI); });
        restoreDirFromSelection(targetItems);
        loadDialogSettingsToUI(dialogUI, targetItems, textFrameTargets);

        var zoomControls = addZoomControls(dialogUI.dialogWindow, doc, getLabel("fieldLabel", "zoom"), initialViewState, {
            min: ZOOM_MIN,
            max: ZOOM_MAX,
            sliderWidth: ZOOM_SLIDER_WIDTH,
            margins: [0, 0, 0, 10],
            redraw: true,
            lightMode: true,
            lightModeLabel: getLabel("checkbox", "lightMode"),
            lightModeTip: getLabel("tooltip", "lightMode"),
            lightModeDefault: false
        });
        var buttonGroup = dialogUI.dialogWindow.add("group");
        buttonGroup.alignment = ["center", "top"];
        buttonGroup.add("button", undefined, getLabel("button", "cancel"), { name: "cancel" });
        buttonGroup.add("button", undefined, getLabel("button", "ok"), { name: "ok" });

        updatePreview(dialogUI);

        var dialogResult = dialogUI.dialogWindow.show();
        saveDialogSettingsFromUI(dialogUI);

        /* プレビューを削除して元のパスを復元 */
        removePreview();
        setTargetsVisible(true);
        app.redraw();

        if (dialogResult !== 1) {
            zoomControls.restoreInitial();
            return;
        }

        var angleDeg = parseFloat(dialogUI.angleInput.text);
        if (isNaN(angleDeg) || angleDeg <= 0 || angleDeg >= 90) {
            alert(getLabel("alert", "invalidAngle"));
            app.redraw();
            return;
        }
        /* 確定：引き出し線を生成 / Build the leader lines */
        applyLeaderLines(doc, dialogUI, targetItems, textFrameTargets, angleDeg * Math.PI / 180);
    }

    main();

}());
