#targetengine "SmartGridMakerEngine"
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

長方形の選択、またはアートボードを基準に、囲み罫とグリッドを一括生成します。
外枠・タイトルエリア・内側エリアの分割や線種、裁ち落とし対応のフレームを、プレビューを見ながら1つのダイアログで設定できます。

詳細は README を参照してください。

### Overview

Builds a border and a grid from a selected rectangle, or from the artboard.
The outer frame, the title area, the inner-area divisions, the line types, and a bleed-aware frame are all set in one dialog with a live preview.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartGridMaker";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.6.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-02-24";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-16";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartGridMaker.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartGridMaker.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n2b01f896c423"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 生成の設定 / Generation settings
    // =========================================

    /* 生成と既定値の設定（長さの既定値はmmで持ち、現在の定規単位に換算して使う）
       Generation and default values; lengths are in mm and converted to the current ruler unit */
    var GENERATION_SETTINGS = {
        bleedMm: 3,                /* 裁ち落とし幅（mm） / bleed width in mm */
        innerOffsetDivisor: 40,    /* 内側オフセット初期値＝(幅+高さ)/この値 / inner offset default divisor */
        titleSizeDivisor: 5,       /* タイトルエリア初期値＝外側エリアの幅または高さ/この値 / title size default divisor */
        maxGridCount: 100,         /* 列数・行数の上限 / maximum number of columns and rows */
        defaultMarginMm: 15,       /* マージンの初期値（mm） / default margin */
        defaultEdgeScaleMm: -5,    /* 辺の伸縮の初期値（mm） / default edge scale */
        defaultFrameWidthMm: 10,   /* フレームON時の既定幅（mm） / frame width applied when enabled */
        defaultRoundMm: 2          /* 角丸ON時の既定値（mm） / radius applied when rounding is enabled */
    };

    /* 生成物のグレーの濃度（CMYKのK%） / Gray tints of generated items (K in CMYK) */
    var GRAY_TINTS = {
        innerCell: 15,  /* 内側エリアのセル / inner area cells */
        titleBand: 30,  /* タイトル帯 / title band */
        frame: 50,      /* フレーム / frame */
        rule: 100       /* 罫線 / rules */
    };

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ダイアログとパネルの外観 / Dialog and panel appearance */
    var DIALOG_LAYOUT = {
        dialogOffsetX: 300,       /* ダイアログの表示位置オフセットX / dialog offset X */
        dialogOffsetY: 0,         /* ダイアログの表示位置オフセットY / dialog offset Y */
        dialogOpacity: 0.98,      /* ダイアログの不透明度 / dialog opacity */
        tabSize: [300, 460],      /* タブパネルの最小サイズ / minimum size of the tabbed panel */
        panelMargins: [15, 20, 15, 10], /* パネル余白 [左,上,右,下] / panel margins */
        viewLabelWidth: 58,       /* 画面表示タブのラベル幅 / label width in the Display tab */
        viewSliderWidth: 200,     /* 画面表示タブのスライダー幅 / slider width in the Display tab */
        viewButtonWidth: 220,     /* 表示コマンドボタンの幅 / view command button width */
        viewButtonHeight: 22,     /* 表示コマンドボタンの高さ / view command button height */
        toggleCheckboxWidth: 20   /* ラベルなしチェックボックスの幅（空文字ぶんの余白を抑える）
                                     width of the label-less checkboxes, to drop the phantom text space */
    };

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 実行環境のロケールからUI言語を判定します。
     *
     * @returns {string} "ja" または "en"。
     */
    function getCurrentLang() {
        return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* UI文言の定義 / UI string definitions */
    var LABELS = {
        dialog: {
            title: { ja: "囲み罫とグリッド", en: "Border and Grid" }
        },
        tab: {
            margin: { ja: "アートボード", en: "Artboard" },
            outer: { ja: "外枠", en: "Outer" },
            inner: { ja: "内側エリア", en: "Inner Area" },
            display: { ja: "画面表示", en: "Display" }
        },
        panel: {
            margin: { ja: "マージン", en: "Margin" },
            frame: { ja: "フレーム", en: "Frame" },
            outer: { ja: "外側エリア", en: "Outer Area" },
            strokeCap: { ja: "線端", en: "Line Caps" },
            titleArea: { ja: "タイトルエリア", en: "Title Area" },
            innerArea: { ja: "内側エリア", en: "Inner Area" },
            offset: { ja: "オフセット", en: "Offset" },
            columns: { ja: "列", en: "Columns" },
            rows: { ja: "行", en: "Rows" },
            lineType: { ja: "線の種類", en: "Line Type" },
            zoomPan: { ja: "ズームとパン", en: "Zoom & Pan" },
            viewCommands: { ja: "表示コマンド", en: "View Commands" }
        },
        checkbox: {
            keepOuter: { ja: "外枠を残す", en: "Keep outer frame" },
            edgeScale: { ja: "辺の伸縮", en: "Extend edges" },
            round: { ja: "角丸", en: "Round" },
            fill: { ja: "塗り", en: "Fill" },
            titleDivider: { ja: "仕切り線", en: "Divider" },
            dividerScale: { ja: "線の伸縮", en: "Extend divider" },
            bleed: { ja: "裁ち落とし", en: "Bleed" },
            link: { ja: "連動", en: "Link" },
            divider: { ja: "分割線", en: "Dividers" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        radio: {
            capButt: { ja: "なし", en: "Butt" },
            capRound: { ja: "丸型", en: "Round" },
            capProject: { ja: "突出", en: "Projecting" },
            lineSolid: { ja: "実線", en: "Solid" },
            lineDash: { ja: "破線", en: "Dashed" },
            lineDots: { ja: "点線", en: "Dotted" },
            top: { ja: "上", en: "Top" },
            bottom: { ja: "下", en: "Bottom" },
            left: { ja: "左", en: "Left" },
            right: { ja: "右", en: "Right" }
        },
        fieldLabel: {
            top: { ja: "上", en: "Top" },
            bottom: { ja: "下", en: "Bottom" },
            left: { ja: "左", en: "Left" },
            right: { ja: "右", en: "Right" },
            width: { ja: "幅", en: "Width" },
            titleSize: { ja: "幅／高さ", en: "Size" },
            columnCount: { ja: "列数", en: "Columns" },
            rowCount: { ja: "行数", en: "Rows" },
            spacing: { ja: "間隔", en: "Spacing" },
            zoom: { ja: "ズーム", en: "Zoom" },
            panX: { ja: "左右", en: "Pan L/R" },
            panY: { ja: "上下", en: "Pan U/D" }
        },
        button: {
            fitArtboard: { ja: "アートボードを全体表示", en: "Fit Artboard in Window" },
            actualSize: { ja: "100%表示", en: "Actual Size" },
            fitAll: { ja: "すべてのアートボードを全体表示", en: "Fit All in Window" },
            zoomOut10: { ja: "10%縮小", en: "Zoom Out 10%" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        tooltip: {
            keepOuter: {
                ja: "OFFにすると、基準にした長方形を削除します",
                en: "Turning this off removes the rectangle used as the base"
            },
            outerRound: {
                ja: "外枠の角を丸めます（ライブエフェクト）。タイトル帯の角にも同じ値を使います。\n辺の伸縮とは同時に使えません",
                en: "Rounds the corners of the outer frame (live effect); the title band reuses the same radius.\nCannot be combined with the edge extension"
            },
            edgeScale: {
                ja: "＋で各辺を両端から伸ばし、−で縮めます。0以外にすると外枠を4本の線に分解します",
                en: "Positive extends each edge at both ends, negative shortens it. Any value but zero splits the outer frame into four lines"
            },
            strokeCap: {
                ja: "辺を4本の線に分解したときだけ有効です",
                en: "Only available while the edges are split into four lines"
            },
            titleArea: {
                ja: "タイトル用の帯を、外枠の内側に作ります",
                en: "Adds a title band inside the outer frame"
            },
            titleSize: {
                ja: "上・下に置くときは高さ、左・右に置くときは幅です",
                en: "A height for a title at the top or bottom, a width for one at the left or right"
            },
            titleFill: {
                ja: "タイトル帯をK30%で塗ります",
                en: "Fills the title band with 30% black"
            },
            titleDivider: {
                ja: "タイトル帯と本文の境目に線を引きます",
                en: "Draws a line between the title band and the body"
            },
            dividerScale: {
                ja: "＋で仕切り線の両端を短く、−で長くします",
                en: "Positive shortens the divider at both ends, negative extends it"
            },
            frame: {
                ja: "アートボードの外周に太い帯を作ります（内側は穴あき）",
                en: "Adds a thick band around the artboard, with the inside cut out"
            },
            bleed: {
                ja: "フレームを裁ち落とし（3mm）まで広げます",
                en: "Extends the frame out to the bleed (3 mm)"
            },
            frameRound: {
                ja: "内側（穴）の角を丸めます",
                en: "Rounds the corners of the inner cutout"
            },
            offset: {
                ja: "外枠（タイトルエリアを除く）から内側エリアまでの距離",
                en: "Distance from the outer frame, excluding the title area, to the inner area"
            },
            spacing: {
                ja: "列／行が2以上のときに有効です",
                en: "Only available for two or more columns or rows"
            },
            innerFill: {
                ja: "各セルをK15%で塗ります。OFFのときは実行時に削除されます",
                en: "Fills each cell with 15% black; the cells are dropped on OK while this is off"
            },
            divider: {
                ja: "間隔の中央に線を1本ずつ引きます",
                en: "Draws one line at the centre of each gutter"
            },
            link: {
                ja: "上の値を下・左・右にも使います",
                en: "Uses the top value for the bottom, left and right as well"
            },
            numberInput: {
                ja: "↑↓で±1、Shift＋↑↓で±10、Option＋↑↓で±0.1",
                en: "Up/Down: ±1, Shift+Up/Down: ±10, Option+Up/Down: ±0.1"
            },
            viewSlider: {
                ja: "Optionキーを押しながらドラッグすると微調整できます",
                en: "Hold Option while dragging for fine adjustment"
            }
        },
        alert: {
            baseRectFailed: {
                ja: "アートボードを基準にする長方形を作成できませんでした。\nレイヤーのロックや非表示を解除してから実行してください。",
                en: "The rectangle used as the artboard base could not be created.\nUnlock the layer and make it visible, then run the script again."
            }
        }
    };

    /**
     * ラベル（ja/en の組）を現在の言語の文言にします。
     *
     * @param {Object} labelSet - { ja, en } 形式のラベル。
     * @returns {string} 現在の言語の文言。
     */
    function getLabel(labelSet) {
        return (labelSet && labelSet[uiLang]) || "";
    }

    /**
     * 入力欄の前に置く項目名に、言語別のコロンを付けます（日本語は全角、英語は半角）。
     *
     * チェックボックス・ラジオボタン・パネル名には付けません（項目名ではないため）。
     *
     * @param {Object} labelSet - { ja, en } 形式のラベル。
     * @returns {string} コロンを付けた文言。
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + ((uiLang === "ja") ? "：" : ":");
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 定規の単位コード→ラベルとpt換算係数（rulerTypeではコード5をHとして扱う）
       Ruler unit code to label and points factor (rulerType treats code 5 as H) */
    var RULER_UNITS = {
        0: { label: "in", factor: 72.0 },
        1: { label: "mm", factor: 72.0 / 25.4 },
        2: { label: "pt", factor: 1.0 },
        3: { label: "pica", factor: 12.0 },
        4: { label: "cm", factor: 72.0 / 2.54 },
        5: { label: "H", factor: 72.0 / 25.4 * 0.25 },   /* 0.25mm */
        6: { label: "px", factor: 1.0 },
        7: { label: "ft/in", factor: 72.0 * 12.0 },
        8: { label: "m", factor: 72.0 / 25.4 * 1000.0 },
        9: { label: "yd", factor: 72.0 * 36.0 },
        10: { label: "ft", factor: 72.0 * 12.0 }
    };

    /**
     * 現在の定規単位のラベルと換算係数を取得します（不明な単位は pt 扱い）。
     *
     * @returns {Object} {label, factor}（factor は単位→ptの換算係数）。
     */
    function getRulerUnit() {
        return RULER_UNITS[app.preferences.getIntegerPreference("rulerType")] || RULER_UNITS[2];
    }

    /**
     * mmをptに換算します。
     *
     * @param {number} mm - mmの値。
     * @returns {number} ptの値。
     */
    function mmToPt(mm) {
        return (72.0 / 25.4) * mm;
    }

    // =========================================
    // セッション状態 / Session state
    // =========================================
    /* Illustratorの起動中だけダイアログの値を保持する
       Dialog values are kept only while Illustrator is running */

    var SESSION_STATE_KEY = "__SmartGridMaker__";

    /**
     * 前回のダイアログ設定を読み込みます。
     *
     * @returns {Object} 保存されていた状態。未保存なら空のオブジェクト。
     */
    function loadSessionState() {
        return $.global[SESSION_STATE_KEY] || {};
    }

    /**
     * 現在のダイアログ設定を保存します。
     *
     * @param {Object} state - 保存する状態。
     * @returns {void}
     */
    function saveSessionState(state) {
        $.global[SESSION_STATE_KEY] = state;
    }

    /**
     * 入れ子のオブジェクトから、ドット区切りのパスで値を取り出します。
     *
     * @param {Object} source - 探索するオブジェクト。
     * @param {string} path - 例 "inner.grid.columns" のようなパス。
     * @returns {*} 見つかった値。存在しない場合は undefined。
     */
    function getStateValue(source, path) {
        var segments = path.split(".");
        var node = source;

        for (var i = 0; i < segments.length; i++) {
            if (node == null) return undefined;
            node = node[segments[i]];
        }
        return node;
    }

    /**
     * 入れ子のオブジェクトに、ドット区切りのパスで値を書き込みます（中間オブジェクトは自動作成）。
     *
     * @param {Object} target - 書き込み先のオブジェクト。
     * @param {string} path - 例 "inner.grid.columns" のようなパス。
     * @param {*} value - 書き込む値。
     * @returns {void}
     */
    function setStateValue(target, path, value) {
        var segments = path.split(".");
        var node = target;

        for (var i = 0; i < segments.length - 1; i++) {
            if (!node[segments[i]]) node[segments[i]] = {};
            node = node[segments[i]];
        }
        node[segments[segments.length - 1]] = value;
    }

    // =========================================
    // ズームとパン / ViewControl
    // =========================================
    /* 画面表示のズームとパン（左右・上下）をスライダーで操作する部品（単体で切り出して使える）
       Zoom and pan sliders for the Illustrator view; self-contained so it can be extracted

       - パンの基準はアクティブなアートボードの中心、可動範囲はアートボードの半分
         Pan is relative to the active artboard centre, within half the artboard size
       - ズームしてもパン量は保つ / Zooming keeps the pan offsets
       - Option(Alt)を押しながらドラッグすると移動量が1/10 / Option(Alt)-drag moves at one tenth speed

       使い方 / Usage:
         var viewControl = ViewControl.create(doc);
         viewControl.buildUI(group, { labelWidth: 58, sliderWidth: 200, zoomLabel: "ズーム：", panXLabel: "左右：", panYLabel: "上下：" });
         viewControl.restore(); // 元の表示に戻す / back to the original view */

    var ViewControl = (function () {

        var ZOOM_MIN_PERCENT = 10;     /* スライダーの最小倍率（%） / slider minimum */
        var ZOOM_MAX_PERCENT = 1600;   /* スライダーの最大倍率（%） / slider maximum */

        /**
         * 値を下限・上限の範囲に収めます。
         *
         * @param {number} value - 対象の値。
         * @param {number} minValue - 下限。
         * @param {number} maxValue - 上限。
         * @returns {number} 範囲に収めた値。
         */
        function clamp(value, minValue, maxValue) {
            return Math.min(Math.max(value, minValue), maxValue);
        }

        /**
         * Option(Alt)キーが押されているかを返します。
         *
         * @returns {boolean} 押されている場合は true。
         */
        function isAltKeyDown() {
            var keyboard = ScriptUI.environment.keyboardState;
            return !!(keyboard && keyboard.altKey);
        }

        /**
         * スライダーの値を処理に渡します。Option(Alt)を押しながらのドラッグは移動量を1/10にします。
         *
         * @param {Slider} slider - 対象のスライダー。
         * @param {Object} sliderState - {rawValue, effectiveValue}。呼び出し側がスライダーごとに保持します。
         * @param {Function} applyValue - 補正後の値を受け取る処理。
         * @returns {void}
         */
        function applySliderValue(slider, sliderState, applyValue) {
            var rawValue = Number(slider.value);
            if (isNaN(rawValue)) rawValue = 0;

            var effectiveValue = rawValue;
            if (isAltKeyDown() && sliderState.rawValue != null) {
                effectiveValue = sliderState.effectiveValue + (rawValue - sliderState.rawValue) * 0.1;
                slider.value = effectiveValue;
            }

            sliderState.rawValue = rawValue;
            sliderState.effectiveValue = effectiveValue;
            applyValue(effectiveValue);
        }

        /**
         * アクティブなアートボードの矩形を返します。
         *
         * @param {Document} doc - 対象のドキュメント。
         * @returns {number[]} [左, 上, 右, 下] の座標。
         */
        function getActiveArtboardRect(doc) {
            return doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
        }

        /**
         * パンの可動範囲を返します（アートボードの半分を 100〜50000pt に収めた値）。
         *
         * @param {Document} doc - 対象のドキュメント。
         * @returns {Object} {xMax, yMax}（pt）。
         */
        function getPanRange(doc) {
            var rect = getActiveArtboardRect(doc);
            return {
                xMax: clamp(Math.round(Math.abs(rect[2] - rect[0]) / 2), 100, 50000),
                yMax: clamp(Math.round(Math.abs(rect[1] - rect[3]) / 2), 100, 50000)
            };
        }

        /**
         * ドキュメントの表示を操作する ViewControl を作ります。
         *
         * @param {Document} doc - 対象のドキュメント。
         * @returns {Object} {zoomBy, restore, buildUI}
         */
        function create(doc) {
            var view = doc.views[0];
            var originalZoom = view.zoom;
            var originalCenter = view.centerPoint;

            var panX = 0; /* pt */
            var panY = 0; /* pt（UIでは下が正 / positive is down in the UI） */
            var zoomSlider = null;

            /**
             * 表示倍率と表示中心を反映します（中心はアートボード中心＋パン量）。
             *
             * @param {number|null} zoomFactor - 表示倍率（1＝100%）。null なら倍率は変えません。
             * @returns {void}
             */
            function applyView(zoomFactor) {
                try {
                    if (zoomFactor != null) view.zoom = clamp(zoomFactor, 0.0313, 640.0);

                    /* Illustratorは上が＋Y、UIは下が正なので引く / Illustrator's +Y is up, so subtract */
                    var rect = getActiveArtboardRect(doc);
                    view.centerPoint = [(rect[0] + rect[2]) / 2 + panX, (rect[1] + rect[3]) / 2 - panY];
                    app.redraw();

                    /* ズームスライダーを現在の表示倍率に追従させる / Keep the zoom slider in step */
                    if (zoomFactor != null && zoomSlider) zoomSlider.value = Math.round(view.zoom * 100);
                } catch (e) { }
            }

            /**
             * パン量を可動範囲に収めて反映します。
             *
             * @param {string} axis - "x" または "y"。
             * @param {number} value - パン量（pt）。
             * @returns {void}
             */
            function setPan(axis, value) {
                var panRange = getPanRange(doc);
                var panPt = Math.round(Number(value) || 0);

                if (axis === "x") panX = clamp(panPt, -panRange.xMax, panRange.xMax);
                else panY = clamp(panPt, -panRange.yMax, panRange.yMax);
                applyView(null);
            }

            /**
             * ラベル＋スライダーの1行を追加します。
             *
             * @param {Group} parent - 追加先のグループ。
             * @param {Object} options - buildUI() に渡された設定。
             * @param {string} labelText - 項目名。
             * @param {number} value - 初期値。
             * @param {number} minValue - 最小値。
             * @param {number} maxValue - 最大値。
             * @param {Function} applyValue - 補正後の値を受け取る処理。
             * @returns {Slider} 追加したスライダー。
             */
            function addSliderRow(parent, options, labelText, value, minValue, maxValue, applyValue) {
                var row = parent.add("group");
                row.orientation = "row";
                row.alignChildren = ["left", "center"];

                var label = row.add("statictext", undefined, labelText);
                label.preferredSize.width = options.labelWidth;

                var slider = row.add("slider", undefined, value, minValue, maxValue);
                slider.preferredSize.width = options.sliderWidth;
                if (options.sliderHelpTip) slider.helpTip = options.sliderHelpTip;

                var sliderState = { rawValue: null, effectiveValue: null };
                slider.onChanging = function () {
                    applySliderValue(slider, sliderState, applyValue);
                };
                return slider;
            }

            return {
                /**
                 * 現在の表示倍率に指定の倍率を掛けます（0.9 なら10%縮小）。
                 *
                 * @param {number} factor - 掛ける倍率。
                 * @returns {void}
                 */
                zoomBy: function (factor) {
                    applyView(view.zoom * factor);
                },

                /**
                 * ダイアログを開く前の表示倍率と表示中心に戻します。
                 *
                 * @returns {void}
                 */
                restore: function () {
                    try {
                        view.zoom = originalZoom;
                        view.centerPoint = originalCenter;
                        app.redraw();
                    } catch (e) { }
                    panX = 0;
                    panY = 0;
                },

                /**
                 * ズーム・左右・上下のスライダーを追加します。
                 *
                 * @param {Group} parent - 追加先のグループ。
                 * @param {Object} options - labelWidth / sliderWidth / zoomLabel / panXLabel / panYLabel / sliderHelpTip。
                 * @returns {void}
                 */
                buildUI: function (parent, options) {
                    var initialZoomPercent = Math.round(originalZoom * 100);
                    if (!(initialZoomPercent >= ZOOM_MIN_PERCENT)) initialZoomPercent = 100;

                    var panRange = getPanRange(doc);

                    zoomSlider = addSliderRow(parent, options, options.zoomLabel, initialZoomPercent, ZOOM_MIN_PERCENT, ZOOM_MAX_PERCENT, function (value) {
                        applyView(clamp(Math.round(value), ZOOM_MIN_PERCENT, ZOOM_MAX_PERCENT) / 100);
                    });
                    addSliderRow(parent, options, options.panXLabel, 0, -panRange.xMax, panRange.xMax, function (value) {
                        setPan("x", value);
                    });
                    addSliderRow(parent, options, options.panYLabel, 0, -panRange.yMax, panRange.yMax, function (value) {
                        setPan("y", value);
                    });
                }
            };
        }

        return { create: create };
    })();

    // =========================================
    // 生成物のタグ / Tags of generated items
    // =========================================
    /* 生成したオブジェクトは name と note の両方にタグを持たせ、後処理で見分けます
       Generated items carry the same tag in name and note so later passes can find them */

    var TAG_OUTER_EDGE = "__OuterEdge__";          /* 外枠の4辺 / the four outer edges */
    var TAG_OUTER_ROUND = "__OuterRoundPreview__"; /* 外枠の角丸プレビュー / outer round preview */
    var TAG_TITLE_FILL = "__TitleFill__";          /* タイトル帯の塗り / title band fill */
    var TAG_TITLE_DIVIDER = "__TitleDivider__";    /* タイトル帯の分割線 / title band divider */
    var TAG_INNER_FILL = "__InnerBoxFill__";       /* 内側エリアの塗り / inner area fill */
    var TAG_FRAME_FILL = "__FrameFill__";          /* フレーム / frame */

    /**
     * 生成したオブジェクトにタグを付けます（name と note の両方）。
     *
     * @param {PageItem} item - 対象のオブジェクト。
     * @param {string} tag - 付与するタグ文字列。
     * @returns {void}
     */
    function tagItem(item, tag) {
        item.name = tag;
        item.note = tag;
    }

    /**
     * 内部タグをレイヤーパネルから隠します（name をクリアし、判定用の note は残します）。
     *
     * @param {PageItem[]} items - 対象のオブジェクト。
     * @returns {void}
     */
    function clearTagNames(items) {
        for (var i = 0; i < items.length; i++) {
            /* パスファインダーでグループが置き換わり、参照が無効になっていることがある
               A pathfinder result may have replaced a group, leaving a stale reference */
            try { items[i].name = ""; } catch (e) { }
        }
    }

    // =========================================
    // オブジェクトの操作 / Item helpers
    // =========================================

    /**
     * K版のみのグレー（CMYK）を作成します。
     *
     * @param {number} blackPercent - K版の濃度（0〜100）。
     * @returns {CMYKColor} 生成したカラー。
     */
    function makeGrayColor(blackPercent) {
        var color = new CMYKColor();
        color.cyan = 0;
        color.magenta = 0;
        color.yellow = 0;
        color.black = blackPercent;
        return color;
    }

    /**
     * 線なし・グレーの塗りにします。
     *
     * @param {PathItem} item - 対象のパス。
     * @param {number} blackPercent - K版の濃度（0〜100）。
     * @returns {void}
     */
    function setGrayFill(item, blackPercent) {
        item.stroked = false;
        item.filled = true;
        item.fillColor = makeGrayColor(blackPercent);
    }

    /**
     * 塗りなし・指定の線にします。
     *
     * @param {PathItem} item - 対象のパス。
     * @param {Color} color - 線の色。
     * @param {number} width - 線幅（pt）。
     * @returns {void}
     */
    function setStroke(item, color, width) {
        item.filled = false;
        item.stroked = true;
        item.strokeColor = color;
        item.strokeWidth = width;
    }

    /**
     * 線と塗りの見た目をコピーします（角丸プレビュー用の複製に使用）。
     *
     * @param {PathItem} source - コピー元のパス。
     * @param {PathItem} target - コピー先のパス。
     * @returns {void}
     */
    function copyAppearance(source, target) {
        target.stroked = source.stroked;
        target.filled = source.filled;
        target.strokeColor = source.strokeColor;
        target.fillColor = source.fillColor;
        target.strokeWidth = source.strokeWidth;
    }

    /**
     * オブジェクトを最背面へ送ります（環境差で落ちることがあるため保護）。
     *
     * @param {PageItem} item - 対象のオブジェクト。
     * @returns {void}
     */
    function sendToBack(item) {
        try { item.zOrder(ZOrderMethod.SENDTOBACK); } catch (e) { }
    }

    /**
     * オブジェクトを削除します（すでに消えている場合は何もしません）。
     *
     * @param {PageItem} item - 対象のオブジェクト。
     * @returns {void}
     */
    function removeItem(item) {
        try { item.remove(); } catch (e) { }
    }

    /**
     * オブジェクトの配列をまとめて削除します。
     *
     * @param {PageItem[]} items - 対象のオブジェクト。
     * @returns {void}
     */
    function removeItems(items) {
        for (var i = 0; i < items.length; i++) {
            removeItem(items[i]);
        }
    }

    /**
     * 角丸のライブエフェクトを適用します。
     *
     * @param {PageItem} item - 対象のオブジェクト。
     * @param {number} radiusPt - 角丸の半径（pt）。0以下なら何もしません。
     * @returns {void}
     */
    function applyRoundCornersEffect(item, radiusPt) {
        if (!(radiusPt > 0)) return;
        var effectXml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius #value# "/></LiveEffect>';
        try { item.applyEffect(effectXml.replace('#value#', radiusPt)); } catch (e) { }
    }

    (function () {
        // =========================================
        // 準備 / Setup
        // =========================================
        if (app.documents.length === 0) return;
        var doc = app.activeDocument;
        var rulerUnit = getRulerUnit();

        /* 一部環境で StrokeCap が未定義になるため、最低限の定数を用意
           Provide the StrokeCap constants for hosts that do not expose them */
        if (typeof StrokeCap === "undefined") {
            StrokeCap = {
                BUTTENDCAP: 0,
                ROUNDENDCAP: 1,
                PROJECTINGENDCAP: 2
            };
        }

        /* 選択パスの元の見た目（キャンセル時に戻すため）
           The original look of the selected paths, restored on cancel */
        var originalAppearances = [];

        /* 基準の長方形（選択中のパス、またはアートボード基準の一時矩形）
           Base rectangles: the selected paths, or the temporary artboard rectangle */
        var baseRects = collectSelectedPaths();

        /* パスが選択されていなければ、アクティブなアートボードを基準にする
           With no path selected, the active artboard becomes the base */
        var isArtboardBased = (baseRects.length === 0);
        var artboardBaseRect = null;

        /* 生成したオブジェクト（プレビューと実行の両方） / Items generated for the preview or the final run */
        var generatedItems = [];

        /* 自動ONの判定に使う、直前の状態と手動操作のフラグ
           Previous states and manual-input flags used by the auto-on rules */
        var titleHadSize = false;           /* タイトルエリアにサイズがあったか / the title area had a size */
        var gridWasSplittable = false;      /* 列・行が分割可能だったか / the grid could carry dividers */
        var innerFillManuallySet = false;   /* 内側エリアの［塗り］を操作したか / inner fill was set by hand */
        var bleedManuallySet = false;       /* フレームの［裁ち落とし］を操作したか / bleed was set by hand */

        /* ダイアログのコントロール / Dialog controls */
        var viewControl = ViewControl.create(doc);
        var marginFields, innerOffsetFields;
        var frameCheckbox, frameWidthGroup, frameWidthInput, bleedCheckbox, frameRoundCheckbox, frameRoundInput;
        var keepOuterCheckbox, outerRoundCheckbox, outerRoundInput;
        var outerEdgeScaleCheckbox, outerEdgeScaleValueGroup, outerEdgeScaleInput;
        var strokeCapPanel, capButtRadio, capRoundRadio, capProjectRadio;
        var titleCheckbox, titleSizeInput;
        var titlePositionGroup, titleTopRadio, titleBottomRadio, titleLeftRadio, titleRightRadio;
        var titleOptionGroup, titleFillCheckbox, titleLineCheckbox;
        var titleEdgeScaleRow, titleEdgeScaleCheckbox, titleEdgeScaleInput;
        var columnCountInput, columnGutterInput, rowCountInput, rowGutterInput;
        var innerFillCheckbox, innerDividerCheckbox;
        var lineTypePanel, lineSolidRadio, lineDashRadio, lineDotsRadio;
        var previewCheckbox;

        // =========================================
        // 基準の長方形 / Base rectangles
        // =========================================

        /**
         * 選択からパスアイテムだけを取り出し、見た目を塗りなし・K100の1pt線にそろえます。
         *
         * 元の見た目は originalAppearances に控え、キャンセル時に戻せるようにします。
         * テキスト編集中は doc.selection が TextRange になるため、選択なしとして扱います。
         *
         * @returns {PathItem[]} 選択中のパスアイテム。
         */
        function collectSelectedPaths() {
            var selectedPaths = [];
            var selectedItems = doc.selection;

            /* 配列でなければ（テキスト編集中の TextRange など）選択なし扱い
               Anything but an array of items, such as a TextRange, counts as no selection */
            if (!selectedItems || !(selectedItems instanceof Array)) return selectedPaths;

            for (var i = 0; i < selectedItems.length; i++) {
                var item = selectedItems[i];
                if (!item || item.typename !== "PathItem") continue;

                originalAppearances.push({
                    item: item,
                    filled: item.filled,
                    fillColor: item.fillColor,
                    stroked: item.stroked,
                    strokeColor: item.strokeColor,
                    strokeWidth: item.strokeWidth
                });

                setStroke(item, makeGrayColor(GRAY_TINTS.rule), 1);
                selectedPaths.push(item);
            }
            return selectedPaths;
        }

        /**
         * 選択パスの見た目を、スクリプト実行前の状態に戻します（キャンセル時に使用）。
         *
         * @returns {void}
         */
        function restoreSelectedAppearances() {
            for (var i = 0; i < originalAppearances.length; i++) {
                var appearance = originalAppearances[i];
                try {
                    appearance.item.filled = appearance.filled;
                    appearance.item.fillColor = appearance.fillColor;
                    appearance.item.stroked = appearance.stroked;
                    appearance.item.strokeColor = appearance.strokeColor;
                    appearance.item.strokeWidth = appearance.strokeWidth;
                } catch (e) { }
            }
        }

        /**
         * アクティブなアートボードの矩形を返します。
         *
         * @returns {number[]} [左, 上, 右, 下] の座標。
         */
        function getActiveArtboardRect() {
            return doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
        }

        /**
         * アートボード基準の一時矩形を削除します（baseRects からも外します）。
         *
         * 削除後に参照へ触ると Error 45 になるため、参照を先に外します。
         *
         * @returns {void}
         */
        function removeArtboardBaseRect() {
            if (!artboardBaseRect) return;

            for (var i = baseRects.length - 1; i >= 0; i--) {
                if (baseRects[i] === artboardBaseRect) baseRects.splice(i, 1);
            }

            removeItem(artboardBaseRect);
            artboardBaseRect = null;
        }

        /**
         * アートボード基準の一時矩形を、マージンの分だけ内側に作り直します。
         *
         * マージンが大きすぎて領域が残らない場合は、マージンを 0 とみなします。
         * 裁ち落としはここでは適用しません（フレームだけに適用します）。
         *
         * @param {Object} marginPt - {top, right, bottom, left}（pt、0以上）。
         * @returns {void}
         */
        function rebuildArtboardBaseRect(marginPt) {
            var artboardRect = getActiveArtboardRect(); // [L, T, R, B]
            var left = artboardRect[0] + marginPt.left;
            var top = artboardRect[1] - marginPt.top;
            var width = (artboardRect[2] - marginPt.right) - left;
            var height = top - (artboardRect[3] + marginPt.bottom);

            if (!(width > 0) || !(height > 0)) {
                left = artboardRect[0];
                top = artboardRect[1];
                width = artboardRect[2] - artboardRect[0];
                height = artboardRect[1] - artboardRect[3];
            }

            removeArtboardBaseRect();

            /* ロックされたレイヤーなどで作成できなくても、ダイアログは続行する
               Keep the dialog running even if the rectangle cannot be created */
            try {
                artboardBaseRect = doc.activeLayer.pathItems.rectangle(top, left, width, height);
                setStroke(artboardBaseRect, makeGrayColor(GRAY_TINTS.rule), 1);
                baseRects.push(artboardBaseRect);
            } catch (e) { }
        }

        /**
         * 基準の長方形をすべて削除します（アートボード基準の一時矩形も含む）。
         *
         * @returns {void}
         */
        function removeBaseRects() {
            removeItems(baseRects);
            baseRects = [];
            artboardBaseRect = null;
        }

        /**
         * 基準の長方形の表示／非表示をまとめて切り替えます。
         *
         * @param {boolean} hidden - 非表示にする場合は true。
         * @returns {void}
         */
        function setBaseRectsHidden(hidden) {
            for (var i = 0; i < baseRects.length; i++) {
                baseRects[i].hidden = hidden;
            }
        }

        // =========================================
        // 生成物の管理 / Generated item tracking
        // =========================================

        /**
         * 生成したオブジェクトを追跡リストに登録して返します。
         *
         * 生成した直後に登録するのが重要です。見た目の設定などで例外が起きても、
         * すでに登録済みならプレビュー解除時に必ず削除できます
         * （登録前に中断すると、消せないオブジェクトがドキュメントに残ります）。
         *
         * @param {PageItem} item - 登録するオブジェクト。
         * @returns {PageItem} 受け取ったオブジェクトをそのまま返します。
         */
        function trackGeneratedItem(item) {
            generatedItems.push(item);
            return item;
        }

        /**
         * 生成したオブジェクトをすべて削除します。
         *
         * @returns {void}
         */
        function removeGeneratedItems() {
            removeItems(generatedItems);
            generatedItems = [];
        }

        // =========================================
        // 入力値の変換 / Input conversion
        // =========================================

        /**
         * 入力欄の文字列を、現在の単位からpt値に換算します。
         *
         * @param {string} text - 入力欄の文字列。
         * @returns {number} pt値。数値として読めない場合は 0。
         */
        function toPt(text) {
            var value = parseFloat(text);
            return isNaN(value) ? 0 : (value * rulerUnit.factor);
        }

        /**
         * 入力欄の文字列を、0以上のpt値に換算します（マージンなど負を許さない項目用）。
         *
         * @param {string} text - 入力欄の文字列。
         * @returns {number} 0以上のpt値。
         */
        function toPositivePt(text) {
            var valuePt = toPt(text);
            return (valuePt > 0) ? valuePt : 0;
        }

        /**
         * 入力欄の文字列を、1以上・上限以下の整数（列数・行数）に変換します。
         *
         * 上限を設けないと、入力のたびに膨大な数のセルを生成してIllustratorが止まります。
         *
         * @param {string} text - 入力欄の文字列。
         * @returns {number} 1〜maxGridCount の整数。
         */
        function toCount(text) {
            var value = parseInt(text, 10);
            if (isNaN(value) || value < 1) return 1;
            return Math.min(value, GENERATION_SETTINGS.maxGridCount);
        }

        /**
         * 列数・行数の入力欄を、1〜上限の範囲に補正します（空欄のあいだは触りません）。
         *
         * @param {EditText} input - 対象の入力欄。
         * @returns {void}
         */
        function snapCountInput(input) {
            if (input.text === "") return;

            var countText = String(toCount(input.text));
            if (countText !== input.text) input.text = countText;
        }

        /**
         * mmで持っている既定値を、現在の定規単位の文字列にします。
         *
         * @param {number} valueMm - mmでの既定値。
         * @returns {string} 現在の定規単位での既定値（小数第1位まで）。
         */
        function defaultValueText(valueMm) {
            return String(Math.round(mmToPt(valueMm) / rulerUnit.factor * 10) / 10);
        }

        /**
         * チェックがONのときだけ、入力欄の値を0以上のpt値で返します。
         *
         * @param {Checkbox} checkbox - 有効／無効を決めるチェックボックス。
         * @param {EditText} input - 値の入力欄。
         * @returns {number} pt値。チェックがOFFなら 0。
         */
        function readCheckedPt(checkbox, input) {
            return checkbox.value ? toPositivePt(input.text) : 0;
        }

        /**
         * 入力欄に正の数値が入っているかを返します。
         *
         * @param {EditText} input - 対象の入力欄。
         * @returns {boolean} 正の数値なら true。
         */
        function hasPositiveValue(input) {
            return (parseFloat(input.text) > 0);
        }

        /**
         * 入力欄が空欄または0なら、既定値を入れます（チェックをONにしたときに使用）。
         *
         * @param {EditText} input - 対象の入力欄。
         * @param {string} defaultText - 入れる既定値。
         * @returns {void}
         */
        function fillDefaultIfZero(input, defaultText) {
            var value = parseFloat(input.text);
            if (isNaN(value) || value === 0) input.text = defaultText;
        }

        /**
         * 内側エリアのオフセット初期値を求めます（外形の (幅+高さ)/40 を10単位に丸めた値）。
         *
         * @returns {number} 現在の定規単位でのオフセット初期値。
         */
        function calcDefaultInnerOffset() {
            if (baseRects.length === 0) return 0;

            var bounds = baseRects[0].geometricBounds; // [L, T, R, B]
            var widthPt = bounds[2] - bounds[0];
            var heightPt = bounds[1] - bounds[3];
            if (!(widthPt > 0) || !(heightPt > 0)) return 0;

            var offsetValue = (widthPt + heightPt) / GENERATION_SETTINGS.innerOffsetDivisor / rulerUnit.factor;

            /* 10以上なら10単位、それ未満は0.1単位に丸める
               （cmやinchのように値が小さくなる単位で0にならないようにする）
               Round to the nearest ten, or to one decimal for units that give small numbers */
            if (offsetValue >= 10) return Math.round(offsetValue / 10) * 10;
            return Math.round(offsetValue * 10) / 10;
        }

        /**
         * タイトルエリアのサイズ初期値を求めます（帯が伸びる方向の長さ / titleSizeDivisor）。
         *
         * 上下に置くときは高さ、左右に置くときは幅を基準にします
         * （高さで決めてしまうと、縦長の長方形で幅を超えてタイトルが作れなくなります）。
         *
         * @returns {number} 現在の定規単位でのサイズ初期値。
         */
        function calcDefaultTitleSize() {
            if (baseRects.length === 0) return 10;

            var bounds = baseRects[0].geometricBounds; // [L, T, R, B]
            var positionKey = getTitlePositionKey();
            var basePt = (positionKey === "top" || positionKey === "bottom")
                ? Math.abs(bounds[1] - bounds[3])
                : Math.abs(bounds[2] - bounds[0]);

            var titleSize = basePt / rulerUnit.factor / GENERATION_SETTINGS.titleSizeDivisor;
            if (!(titleSize > 0)) return 10;

            return Math.max(Math.round(titleSize * 10) / 10, 0.1);
        }

        // =========================================
        // UIの部品 / UI building blocks
        // =========================================

        /**
         * タブを追加し、その中身を並べる縦組みのグループを返します。
         *
         * @param {TabbedPanel} tabbedPanel - 追加先のタブパネル。
         * @param {Object} labelSet - タブ名のラベル。
         * @returns {Group} 中身を追加するグループ。
         */
        function addTabColumn(tabbedPanel, labelSet) {
            var tab = tabbedPanel.add("tab", undefined, getLabel(labelSet));
            tab.orientation = "column";
            tab.alignChildren = ["fill", "top"];
            tab.spacing = 8;
            tab.margins = 10;

            var column = tab.add("group");
            column.orientation = "column";
            column.alignChildren = ["fill", "top"];
            column.spacing = 8;
            return column;
        }

        /**
         * パネル名の後ろに付ける単位を、言語別の括弧で返します（日本語は全角）。
         *
         * @returns {string} 例）"（mm）" / " (mm)"
         */
        function unitSuffix() {
            return (uiLang === "ja") ? ("（" + rulerUnit.label + "）") : (" (" + rulerUnit.label + ")");
        }

        /**
         * 入力欄にツールチップを設定します（↑↓キーの説明を必ず添えます）。
         *
         * @param {EditText} input - 対象の入力欄。
         * @param {Object} [labelSet] - 入力欄ごとの説明。省略時は↑↓キーの説明だけ。
         * @returns {void}
         */
        function setInputHelpTip(input, labelSet) {
            input.helpTip = (labelSet ? getLabel(labelSet) + "\n" : "") + getLabel(LABELS.tooltip.numberInput);
        }

        /**
         * 縦組みのパネルを追加します。
         *
         * @param {Group|Panel} parent - 追加先。
         * @param {string} title - パネル名。
         * @param {number} [spacing] - パネル内の要素間隔。省略時は既定値。
         * @returns {Panel} 追加したパネル。
         */
        function addPanel(parent, title, spacing) {
            var panel = parent.add("panel", undefined, title);
            panel.orientation = "column";
            panel.alignChildren = ["fill", "top"];
            panel.margins = DIALOG_LAYOUT.panelMargins;
            if (spacing) panel.spacing = spacing;
            return panel;
        }

        /**
         * ラジオボタンを横に並べるパネルを追加します。
         *
         * @param {Group|Panel} parent - 追加先。
         * @param {string} title - パネル名。
         * @returns {Panel} 追加したパネル。
         */
        function addRadioPanel(parent, title) {
            var panel = parent.add("panel", undefined, title);
            panel.orientation = "row";
            panel.alignChildren = ["left", "center"];
            panel.margins = DIALOG_LAYOUT.panelMargins;
            return panel;
        }

        /**
         * 左揃えの1行グループを追加します。
         *
         * @param {Group|Panel} parent - 追加先。
         * @param {number} [spacing] - 要素間隔。省略時は既定値。
         * @returns {Group} 追加したグループ。
         */
        function addRow(parent, spacing) {
            var row = parent.add("group");
            row.orientation = "row";
            row.alignChildren = ["left", "center"];
            if (spacing != null) row.spacing = spacing;
            return row;
        }

        /**
         * パネルを非表示にし、レイアウト上の高さも潰します（長方形スタート時に使用）。
         *
         * @param {Panel} panel - 対象のパネル。
         * @param {boolean} collapsed - 畳む場合は true。
         * @returns {void}
         */
        function setPanelCollapsed(panel, collapsed) {
            panel.visible = !collapsed;
            panel.minimumSize.height = 0;
            panel.maximumSize.height = collapsed ? 0 : 10000;
        }

        /**
         * 複数のコントロールに同じイベントハンドラーを割り当てます。
         *
         * @param {Object[]} controls - 対象のコントロール。
         * @param {string} eventName - "onClick" などのハンドラー名。
         * @param {Function} handler - 割り当てる関数。
         * @returns {void}
         */
        function bindAll(controls, eventName, handler) {
            for (var i = 0; i < controls.length; i++) {
                controls[i][eventName] = handler;
            }
        }

        /**
         * 「ラベルなしチェックボックス＋項目名：入力欄 単位」の1行を作ります
         * （タイトルエリアとフレームの有効／無効で共用）。
         *
         * @param {Panel|Group} parent - 追加先。
         * @param {Object} labelSet - 項目名のラベル。
         * @param {Object} [inputTooltipSet] - 入力欄のツールチップ。
         * @returns {Object} {checkbox, valueGroup, input}
         */
        function addToggleValueRow(parent, labelSet, inputTooltipSet) {
            var row = addRow(parent, 0);

            var checkbox = row.add("checkbox", undefined, "");
            /* ラベルがないので、空文字ぶんの幅が入らないよう抑える
               No label, so cap the width to avoid the phantom text space */
            checkbox.preferredSize.width = DIALOG_LAYOUT.toggleCheckboxWidth;

            var valueGroup = addRow(row, 4);
            valueGroup.margins = 0;
            valueGroup.add("statictext", undefined, labelText(labelSet));

            var input = valueGroup.add("edittext", undefined, "0");
            input.characters = 4;
            valueGroup.add("statictext", undefined, rulerUnit.label);
            changeValueByArrowKey(input, false);
            setInputHelpTip(input, inputTooltipSet);

            return { checkbox: checkbox, valueGroup: valueGroup, input: input };
        }

        /**
         * 「チェックボックス＋入力欄＋単位」の1行を作ります（角丸・辺の伸縮で共用）。
         *
         * @param {Panel|Group} parent - 追加先。
         * @param {Object} labelSet - チェックボックスのラベル。
         * @param {string} initialText - 入力欄の初期値。
         * @param {boolean} allowNegative - 負の値を許す場合は true。
         * @returns {Object} {row, checkbox, valueGroup, input}
         */
        function addCheckboxValueRow(parent, labelSet, initialText, allowNegative) {
            var row = addRow(parent);
            var checkbox = row.add("checkbox", undefined, getLabel(labelSet));

            var valueGroup = addRow(row);
            var input = valueGroup.add("edittext", undefined, initialText);
            input.characters = 4;
            valueGroup.add("statictext", undefined, rulerUnit.label);
            changeValueByArrowKey(input, allowNegative);
            setInputHelpTip(input);

            return { row: row, checkbox: checkbox, valueGroup: valueGroup, input: input };
        }

        /**
         * 上／左＋連動＋右／下 の3段組の入力欄を作ります（マージンと内側エリアのオフセットで共用）。
         *
         * ［連動］がONのあいだは、上の値を他の3つへコピーし、3つをディム表示にします。
         *
         * @param {Panel} parent - 追加先のパネル。
         * @param {string} initialText - 入力欄の初期値。
         * @param {number} characters - 入力欄の文字数。
         * @param {string} unitLabel - 上下の入力欄に付ける単位。空文字なら付けません。
         * @param {Object} [inputTooltipSet] - 入力欄のツールチップ。
         * @returns {Object} {top, bottom, left, right, linkCheckbox, applyLinkState}
         */
        function buildLinkedQuadUI(parent, initialText, characters, unitLabel, inputTooltipSet) {
            var inputs = {};
            var fieldGroups = {};
            var followerKeys = ["bottom", "left", "right"];
            var isSyncing = false;

            /**
             * 中央寄せの1段を追加します。
             *
             * @param {number} [spacing] - 要素間隔。
             * @returns {Group} 追加したグループ。
             */
            function addCenteredRow(spacing) {
                var row = parent.add("group");
                row.orientation = "row";
                row.alignChildren = ["center", "center"];
                row.alignment = ["fill", "top"];
                if (spacing) row.spacing = spacing;
                return row;
            }

            /**
             * 「項目名：入力欄（＋単位）」のひとまとまりを追加します。
             *
             * @param {Group} row - 追加先の行。
             * @param {string} positionKey - "top" / "bottom" / "left" / "right"。
             * @param {boolean} withUnit - 単位を付ける場合は true。
             * @returns {void}
             */
            function addField(row, positionKey, withUnit) {
                var fieldGroup = addRow(row);
                fieldGroup.add("statictext", undefined, labelText(LABELS.fieldLabel[positionKey]));

                var input = fieldGroup.add("edittext", undefined, initialText);
                input.characters = characters;
                changeValueByArrowKey(input, false);
                setInputHelpTip(input, inputTooltipSet);

                if (withUnit && unitLabel) fieldGroup.add("statictext", undefined, unitLabel);

                inputs[positionKey] = input;
                fieldGroups[positionKey] = fieldGroup;
            }

            /**
             * ［連動］の状態を反映します（連動ONなら上の値を他へコピーし、3つをディム表示）。
             *
             * @returns {void}
             */
            function applyLinkState() {
                var linked = linkCheckbox.value;

                /* 値の書き込みで onChanging が走らないようにフラグで抑える
                   The flag keeps our own writes from triggering onChanging */
                isSyncing = true;
                for (var i = 0; i < followerKeys.length; i++) {
                    fieldGroups[followerKeys[i]].enabled = !linked;
                    if (linked) inputs[followerKeys[i]].text = inputs.top.text;
                }
                isSyncing = false;
            }

            // 1段目：上（中央寄せ）
            addField(addCenteredRow(), "top", true);

            // 2段目：左 ＋ 連動（中央）＋ 右
            var middleRow = addCenteredRow(12);
            addField(middleRow, "left", false);
            var linkCheckbox = middleRow.add("checkbox", undefined, getLabel(LABELS.checkbox.link));
            linkCheckbox.value = true;
            linkCheckbox.helpTip = getLabel(LABELS.tooltip.link);
            addField(middleRow, "right", false);

            // 3段目：下（中央寄せ）
            addField(addCenteredRow(), "bottom", true);

            inputs.top.onChanging = function () {
                if (isSyncing) return;
                if (linkCheckbox.value) applyLinkState();
                requestPreview();
            };

            bindAll([inputs.bottom, inputs.left, inputs.right], "onChanging", function () {
                if (isSyncing) return;
                requestPreview();
            });

            linkCheckbox.onClick = function () {
                applyLinkState();
                requestPreview();
            };

            applyLinkState();

            return {
                top: inputs.top,
                bottom: inputs.bottom,
                left: inputs.left,
                right: inputs.right,
                linkCheckbox: linkCheckbox,
                applyLinkState: applyLinkState
            };
        }

        /**
         * 「◯数」＋「間隔」の1行パネルを作ります（列／行で共用）。
         *
         * @param {Group} parent - 追加先。
         * @param {Object} titleLabelSet - パネル名のラベル。
         * @param {Object} countLabelSet - 個数の項目名のラベル。
         * @returns {Object} {countInput, gutterInput}
         */
        function addGridCountPanel(parent, titleLabelSet, countLabelSet) {
            var panel = addPanel(parent, getLabel(titleLabelSet), 8);
            var row = addRow(panel);

            row.add("statictext", undefined, labelText(countLabelSet));
            var countInput = row.add("edittext", undefined, "1");
            countInput.characters = 3;
            changeValueByArrowKey(countInput, false);
            setInputHelpTip(countInput);

            row.add("statictext", undefined, labelText(LABELS.fieldLabel.spacing));
            var gutterInput = row.add("edittext", undefined, "0");
            gutterInput.characters = 4;
            changeValueByArrowKey(gutterInput, false);
            setInputHelpTip(gutterInput, LABELS.tooltip.spacing);
            row.add("statictext", undefined, rulerUnit.label);

            return { countInput: countInput, gutterInput: gutterInput };
        }

        /**
         * 表示コマンド用の小さめボタンを1つ追加します。
         *
         * @param {Group} parent - 追加先のグループ。
         * @param {Object} labelSet - ボタンのラベル。
         * @param {Function} action - クリック時に実行する処理。
         * @returns {void}
         */
        function addViewCommandButton(parent, labelSet, action) {
            var button = parent.add("button", undefined, getLabel(labelSet));
            button.alignment = "left";
            button.preferredSize = [DIALOG_LAYOUT.viewButtonWidth, DIALOG_LAYOUT.viewButtonHeight];
            button.minimumSize = [DIALOG_LAYOUT.viewButtonWidth, DIALOG_LAYOUT.viewButtonHeight];
            button.maximumSize = [DIALOG_LAYOUT.viewButtonWidth, DIALOG_LAYOUT.viewButtonHeight];
            button.onClick = action;
        }

        /**
         * ↑↓キーでの値の増減を入力欄に割り当てます（shift=±10でスナップ、option=±0.1）。
         *
         * @param {EditText} editText - 対象の入力欄。
         * @param {boolean} allowNegative - 負の値を許す場合は true。
         * @returns {void}
         */
        function changeValueByArrowKey(editText, allowNegative) {
            editText.addEventListener("keydown", function (event) {
                if (!event || (event.keyName !== "Up" && event.keyName !== "Down")) return;

                var value = Number(editText.text);
                if (isNaN(value)) return;

                /* 先にキーの既定動作を止める（値を書き換えたあとでは間に合わない環境がある）
                   Cancel the default first: some hosts apply it before we finish */
                try { event.preventDefault(); } catch (e) { }

                var keyboard = ScriptUI.environment.keyboardState;
                var isUp = (event.keyName === "Up");
                var isFineStep = !!keyboard.altKey;

                if (keyboard.shiftKey) {
                    /* Shiftキー押下時は10の倍数にスナップ / Snap to the nearest ten */
                    value = isUp ? Math.ceil((value + 1) / 10) * 10 : Math.floor((value - 1) / 10) * 10;
                } else {
                    var delta = isFineStep ? 0.1 : 1;
                    value = isUp ? (value + delta) : (value - delta);
                }

                if (!allowNegative && value < 0) value = 0;

                /* optionキー押下時は小数第1位まで、それ以外は整数に丸める
                   Round to one decimal with option held, otherwise to an integer */
                value = isFineStep ? (Math.round(value * 10) / 10) : Math.round(value);

                editText.text = String(value);

                // keydownでtextを書き換えた場合、onChangingが発火しないことがあるため明示的に呼ぶ
                try {
                    if (typeof editText.onChanging === "function") editText.onChanging();
                } catch (e) { }
            });
        }

        // =========================================
        // コントロールの有効／無効 / Enabled states
        // =========================================

        /**
         * 辺の伸縮の値を返します（チェックがOFFのときは 0 とみなします）。
         *
         * @returns {number} 現在の単位での辺の伸縮量。
         */
        function getOuterEdgeScaleValue() {
            if (!outerEdgeScaleCheckbox.value) return 0;
            var value = parseFloat(outerEdgeScaleInput.text);
            return isNaN(value) ? 0 : value;
        }

        /**
         * 線端パネルの有効／無効を反映します。
         *
         * 線端は「4辺に分解するとき」だけ意味を持ちます（＝外枠を残す＋辺の伸縮≠0）。
         *
         * @returns {void}
         */
        function applyStrokeCapPanelEnabledState() {
            strokeCapPanel.enabled = (keepOuterCheckbox.value && getOuterEdgeScaleValue() !== 0);
        }

        /**
         * 外側エリアの［辺の伸縮］［角丸］［線端］の有効／無効を反映します。
         *
         * ［角丸］のチェック自体は、タイトルエリアの角丸も参照する値のため常に操作できます。
         *
         * @returns {void}
         */
        function applyOuterAreaEnabledState() {
            outerEdgeScaleCheckbox.enabled = keepOuterCheckbox.value;
            outerEdgeScaleValueGroup.enabled = (keepOuterCheckbox.value && outerEdgeScaleCheckbox.value);

            outerRoundInput.enabled = outerRoundCheckbox.value;
            if (!outerRoundCheckbox.value) outerRoundInput.text = "0";

            applyStrokeCapPanelEnabledState();
        }

        /**
         * タイトルエリアの［辺の伸縮］の入力欄の有効／無効を反映します。
         *
         * @returns {void}
         */
        function applyTitleEdgeScaleEnabledState() {
            titleEdgeScaleInput.enabled = titleEdgeScaleCheckbox.value;
            if (!titleEdgeScaleCheckbox.value) titleEdgeScaleInput.text = "0";
        }

        /**
         * タイトルエリアの各コントロールの有効／無効を反映します。
         *
         * 幅／高さはディムせず常に入力できます（チェックを外しても値は残します）。
         *
         * @returns {void}
         */
        function applyTitleAreaEnabledState() {
            var areaEnabled = titleCheckbox.value;
            titleOptionGroup.enabled = areaEnabled;

            /* 位置・塗り・仕切り線は「有効」かつ「サイズ>0」のときだけ
               The position, fill and divider need both the checkbox and a size */
            var usable = (areaEnabled && hasPositiveValue(titleSizeInput));
            titlePositionGroup.enabled = usable;
            titleFillCheckbox.enabled = usable;
            titleLineCheckbox.enabled = usable;

            if (!usable) {
                titleFillCheckbox.value = false;
                titleLineCheckbox.value = false;
            } else if (!titleHadSize) {
                /* 仕切り線：サイズが 0→>0 になった瞬間だけ自動ON（ユーザーは後からOFF可）
                   The divider is auto-enabled only on the zero-to-positive transition */
                titleLineCheckbox.value = true;
            }

            titleHadSize = usable;

            /* ［線の伸縮］は仕切り線がONのときだけ操作できる（自動ONの結果を見てから判定する）
               The divider scale follows the divider checkbox, after the auto-on rule above */
            titleEdgeScaleRow.enabled = (usable && titleLineCheckbox.value);
            if (!titleEdgeScaleRow.enabled) {
                titleEdgeScaleCheckbox.value = false;
                applyTitleEdgeScaleEnabledState();
            }
        }

        /**
         * フレーム関連UIの有効／無効を、基準の種類とフレーム幅に応じて反映します。
         *
         * フレームはアートボード基準のときだけ使えるため、長方形スタート時は値ごとリセットします。
         *
         * @returns {void}
         */
        function applyFrameEnabledState() {
            if (!isArtboardBased) {
                frameCheckbox.value = false;
                bleedCheckbox.value = false;
                frameRoundCheckbox.value = false;
                frameWidthInput.text = "0";
                frameRoundInput.text = "0";

                frameWidthGroup.enabled = false;
                bleedCheckbox.enabled = false;
                frameRoundCheckbox.enabled = false;
                frameRoundInput.enabled = false;
                return;
            }

            /* 幅はディムせず常に入力可。OFFのときはチェックを外すだけで値は残す
               Never dim the width; turning the checkbox off keeps the value as-is */
            frameWidthGroup.enabled = true;

            var usable = (frameCheckbox.value && hasPositiveValue(frameWidthInput));
            bleedCheckbox.enabled = usable;
            frameRoundCheckbox.enabled = usable;
            if (!frameCheckbox.value) {
                bleedCheckbox.value = false;
                frameRoundCheckbox.value = false;
            }
            frameRoundInput.enabled = (usable && frameRoundCheckbox.value);
        }

        /**
         * 列数・行数から分割線を引けるかを判定します。
         *
         * @param {number} columnCount - 列数。
         * @param {number} rowCount - 行数。
         * @returns {boolean} 分割線を引ける場合は true。
         */
        function isGridSplittable(columnCount, rowCount) {
            return (columnCount > 1 || rowCount > 1);
        }

        /**
         * ［分割線］と［線の種類］パネルの有効／無効を反映します。
         *
         * @param {number} columnCount - 列数。
         * @param {number} rowCount - 行数。
         * @param {boolean} allowAutoOn - 1/1から分割可能になった瞬間に自動ONしてよい場合は true。
         * @returns {void}
         */
        function applyInnerDividerEnabledState(columnCount, rowCount, allowAutoOn) {
            var splittable = isGridSplittable(columnCount, rowCount);
            innerDividerCheckbox.enabled = splittable;

            if (!splittable) {
                // 分割できないなら分割線は不要
                innerDividerCheckbox.value = false;
            } else if (allowAutoOn && !gridWasSplittable) {
                // 1/1 から分割可能になった瞬間だけ自動ON
                innerDividerCheckbox.value = true;
            }

            gridWasSplittable = splittable;
            lineTypePanel.enabled = (splittable && innerDividerCheckbox.value);
        }

        /**
         * 他の入力値に依存するコントロールを更新します（collectOptions の副作用をまとめたもの）。
         *
         * @param {Object} options - collectOptions() が組み立てた生成条件。
         * @returns {void}
         */
        function syncDependentControls(options) {
            /* ガターは列／行が2以上のときだけ入力可 / Gutters are editable for 2+ columns or rows only */
            columnGutterInput.enabled = (options.columnCount > 1);
            rowGutterInput.enabled = (options.rowCount > 1);

            /* ガターが入ったら塗りを自動ON（手動操作があれば尊重）
               Turn the fill on once a gutter is set, unless the user set it manually */
            if (!innerFillManuallySet && (options.columnGutterPt !== 0 || options.rowGutterPt !== 0)) {
                innerFillCheckbox.value = true;
            }

            applyInnerDividerEnabledState(options.columnCount, options.rowCount, false);
            applyStrokeCapPanelEnabledState();
        }

        // =========================================
        // パネルの組み立て / Panel builders
        // =========================================

        /**
         * ［マージン］パネルを組み立てます（アートボード基準のときだけ使います）。
         *
         * @param {Group} parent - 追加先のグループ。
         * @returns {void}
         */
        function buildMarginPanel(parent) {
            var marginPanel = addPanel(parent, getLabel(LABELS.panel.margin), 10);
            marginFields = buildLinkedQuadUI(marginPanel, defaultValueText(GENERATION_SETTINGS.defaultMarginMm), 4, rulerUnit.label);

            /* アートボード基準のときだけ表示・操作できる
               The panel is shown and enabled for artboard-based runs only */
            marginPanel.enabled = isArtboardBased;
            setPanelCollapsed(marginPanel, !isArtboardBased);
        }

        /**
         * ［フレーム］パネルを組み立てます（アートボード基準のときだけ使います）。
         *
         * @param {Group} parent - 追加先のグループ。
         * @returns {void}
         */
        function buildFramePanel(parent) {
            var framePanel = addPanel(parent, getLabel(LABELS.panel.frame), 10);
            setPanelCollapsed(framePanel, !isArtboardBased);

            var widthRow = addToggleValueRow(framePanel, LABELS.fieldLabel.width);
            frameCheckbox = widthRow.checkbox;
            frameCheckbox.helpTip = getLabel(LABELS.tooltip.frame);
            frameWidthGroup = widthRow.valueGroup;
            frameWidthInput = widthRow.input;

            /* 裁ち落とし（表示は常時。長方形スタート時は enabled で制御）
               The bleed row is always visible; rectangle-based runs disable it instead */
            bleedCheckbox = addRow(framePanel).add("checkbox", undefined, getLabel(LABELS.checkbox.bleed));
            bleedCheckbox.helpTip = getLabel(LABELS.tooltip.bleed);

            var roundRow = addCheckboxValueRow(framePanel, LABELS.checkbox.round, "0", false);
            frameRoundCheckbox = roundRow.checkbox;
            frameRoundCheckbox.helpTip = getLabel(LABELS.tooltip.frameRound);
            frameRoundInput = roundRow.input;

            frameCheckbox.onClick = function () {
                if (frameCheckbox.value) {
                    /* 幅が0なら既定値を入れ、裁ち落としも自動でON（手動操作があれば尊重）
                       Fill in the default width at zero and turn the bleed on, unless the user set it */
                    fillDefaultIfZero(frameWidthInput, defaultValueText(GENERATION_SETTINGS.defaultFrameWidthMm));
                    if (!bleedManuallySet) bleedCheckbox.value = true;
                }

                applyFrameEnabledState();
                requestPreview();
            };

            frameWidthInput.onChanging = function () {
                /* OFFのときは入力だけ受け付け、プレビューには反映しない
                   While unchecked the field just stores the value; nothing is previewed */
                if (!frameCheckbox.value) return;

                /* 幅が0→>0になったら裁ち落としを自動ON（手動操作があれば尊重）
                   Turn the bleed on once a width is set, unless the user set it manually */
                var hasWidth = hasPositiveValue(frameWidthInput);
                bleedCheckbox.enabled = hasWidth;
                if (!hasWidth) bleedCheckbox.value = false;
                else if (!bleedManuallySet) bleedCheckbox.value = true;

                frameRoundCheckbox.enabled = hasWidth;
                if (!hasWidth) {
                    frameRoundCheckbox.value = false;
                    frameRoundInput.enabled = false;
                }
                requestPreview();
            };

            bleedCheckbox.onClick = function () {
                /* ユーザーが操作したら以後は自動ONしない / Stop auto-enabling once the user decides */
                bleedManuallySet = true;
                requestPreview();
            };

            frameRoundCheckbox.onClick = function () {
                frameRoundInput.enabled = frameRoundCheckbox.value;

                /* ONにしたとき、0なら既定値を入れる / Fill in the default radius when enabled at zero */
                if (frameRoundCheckbox.value) fillDefaultIfZero(frameRoundInput, defaultValueText(GENERATION_SETTINGS.defaultRoundMm));
                requestPreview();
            };

            frameRoundInput.onChanging = requestPreview;

            /* 初期反映（フレーム幅0なら裁ち落とし・角丸はディム）
               Initial state: a zero width dims both the bleed and the rounding */
            applyFrameEnabledState();
        }

        /**
         * ［外側エリア］パネルを組み立てます。
         *
         * @param {Group} parent - 追加先のグループ。
         * @returns {void}
         */
        function buildOuterPanel(parent) {
            var outerPanel = addPanel(parent, getLabel(LABELS.panel.outer), 10);

            keepOuterCheckbox = outerPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.keepOuter));
            keepOuterCheckbox.value = true;
            keepOuterCheckbox.helpTip = getLabel(LABELS.tooltip.keepOuter);

            var roundRow = addCheckboxValueRow(outerPanel, LABELS.checkbox.round, "0", false);
            outerRoundCheckbox = roundRow.checkbox;
            outerRoundCheckbox.helpTip = getLabel(LABELS.tooltip.outerRound);
            outerRoundInput = roundRow.input;

            var edgeScaleRow = addCheckboxValueRow(outerPanel, LABELS.checkbox.edgeScale, defaultValueText(GENERATION_SETTINGS.defaultEdgeScaleMm), true);
            outerEdgeScaleCheckbox = edgeScaleRow.checkbox;
            outerEdgeScaleCheckbox.value = true;
            outerEdgeScaleCheckbox.helpTip = getLabel(LABELS.tooltip.edgeScale);
            outerEdgeScaleValueGroup = edgeScaleRow.valueGroup;
            outerEdgeScaleInput = edgeScaleRow.input;
            outerEdgeScaleInput.active = true;

            buildStrokeCapPanel(outerPanel);

            keepOuterCheckbox.onClick = function () {
                applyOuterAreaEnabledState();
                requestPreview();
            };

            outerRoundCheckbox.onClick = function () {
                if (outerRoundCheckbox.value) {
                    /* ONにしたとき、0なら既定値を入れる / Fill in the default radius when enabled at zero */
                    fillDefaultIfZero(outerRoundInput, defaultValueText(GENERATION_SETTINGS.defaultRoundMm));

                    /* 角丸と辺の伸縮は同時に使えないので、ONにした側を残して他方をOFFにする
                       The radius and the edge scale cannot be combined, so turning one on turns the other off */
                    outerEdgeScaleCheckbox.value = false;
                }

                applyOuterAreaEnabledState();
                requestPreview();
            };

            outerEdgeScaleCheckbox.onClick = function () {
                if (outerEdgeScaleCheckbox.value) outerRoundCheckbox.value = false;

                applyOuterAreaEnabledState();
                requestPreview();
            };

            outerRoundInput.onChanging = requestPreview;

            outerEdgeScaleInput.onChanging = function () {
                if (!outerEdgeScaleCheckbox.value) return;
                applyStrokeCapPanelEnabledState();
                requestPreview();
            };

            applyOuterAreaEnabledState();
        }

        /**
         * ［線端］パネルを組み立てます。
         *
         * @param {Panel} parent - 追加先のパネル。
         * @returns {void}
         */
        function buildStrokeCapPanel(parent) {
            strokeCapPanel = addRadioPanel(parent, getLabel(LABELS.panel.strokeCap));
            strokeCapPanel.helpTip = getLabel(LABELS.tooltip.strokeCap);

            capButtRadio = strokeCapPanel.add("radiobutton", undefined, getLabel(LABELS.radio.capButt));
            capRoundRadio = strokeCapPanel.add("radiobutton", undefined, getLabel(LABELS.radio.capRound));
            capProjectRadio = strokeCapPanel.add("radiobutton", undefined, getLabel(LABELS.radio.capProject));

            /* 初期値は基準の長方形の線端を優先 / Prefer the base rectangle's cap */
            var currentCap = (baseRects.length > 0 && baseRects[0].stroked) ? baseRects[0].strokeCap : null;
            if (currentCap === StrokeCap.ROUNDENDCAP) capRoundRadio.value = true;
            else if (currentCap === StrokeCap.PROJECTINGENDCAP) capProjectRadio.value = true;
            else capButtRadio.value = true;

            bindAll([capButtRadio, capRoundRadio, capProjectRadio], "onClick", requestPreview);
        }

        /**
         * ［タイトルエリア］パネルを組み立てます。
         *
         * @param {Group} parent - 追加先のグループ。
         * @returns {void}
         */
        function buildTitlePanel(parent) {
            var titlePanel = addPanel(parent, getLabel(LABELS.panel.titleArea), 10);

            // 有効 ＋ 幅／高さ（1行）
            var sizeRow = addToggleValueRow(titlePanel, LABELS.fieldLabel.titleSize, LABELS.tooltip.titleSize);
            titleCheckbox = sizeRow.checkbox;
            titleCheckbox.helpTip = getLabel(LABELS.tooltip.titleArea);
            titleSizeInput = sizeRow.input;

            // 位置（上／下／左／右）
            titlePositionGroup = addRow(titlePanel);
            titleTopRadio = titlePositionGroup.add("radiobutton", undefined, getLabel(LABELS.radio.top));
            titleBottomRadio = titlePositionGroup.add("radiobutton", undefined, getLabel(LABELS.radio.bottom));
            titleLeftRadio = titlePositionGroup.add("radiobutton", undefined, getLabel(LABELS.radio.left));
            titleRightRadio = titlePositionGroup.add("radiobutton", undefined, getLabel(LABELS.radio.right));
            titleTopRadio.value = true;

            // 塗り／線
            titleOptionGroup = addRow(titlePanel);
            titleFillCheckbox = titleOptionGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.fill));
            titleFillCheckbox.helpTip = getLabel(LABELS.tooltip.titleFill);

            titleLineCheckbox = titleOptionGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.titleDivider));
            titleLineCheckbox.value = true;
            titleLineCheckbox.helpTip = getLabel(LABELS.tooltip.titleDivider);

            // 仕切り線の伸縮（両端の詰め量）
            var edgeScaleRow = addCheckboxValueRow(titlePanel, LABELS.checkbox.dividerScale, "0", true);
            titleEdgeScaleRow = edgeScaleRow.row;
            titleEdgeScaleCheckbox = edgeScaleRow.checkbox;
            titleEdgeScaleCheckbox.helpTip = getLabel(LABELS.tooltip.dividerScale);
            titleEdgeScaleInput = edgeScaleRow.input;

            titleCheckbox.onClick = function () {
                /* ONにしたとき、サイズが0ならデフォルト値を入れる
                   Fill in the default size when enabled at zero */
                if (titleCheckbox.value && !hasPositiveValue(titleSizeInput)) {
                    titleSizeInput.text = String(calcDefaultTitleSize());
                }

                applyTitleAreaEnabledState();
                requestPreview();
            };

            titleSizeInput.onChanging = function () {
                /* OFFのときは入力だけ受け付け、プレビューには反映しない
                   While unchecked the field just stores the value; nothing is previewed */
                if (!titleCheckbox.value) return;

                applyTitleAreaEnabledState();
                requestPreview();
            };

            bindAll([titleTopRadio, titleBottomRadio, titleLeftRadio, titleRightRadio], "onClick", requestPreview);

            titleFillCheckbox.onClick = requestPreview;

            titleLineCheckbox.onClick = function () {
                applyTitleAreaEnabledState();
                requestPreview();
            };

            titleEdgeScaleCheckbox.onClick = function () {
                applyTitleEdgeScaleEnabledState();
                requestPreview();
            };

            titleEdgeScaleInput.onChanging = function () {
                if (!titleEdgeScaleCheckbox.value) return;
                requestPreview();
            };

            applyTitleEdgeScaleEnabledState();
            applyTitleAreaEnabledState();
        }

        /**
         * ［内側エリア］パネルを組み立てます（オフセット／列・行／線の種類）。
         *
         * @param {Group} parent - 追加先のグループ。
         * @returns {void}
         */
        function buildInnerPanel(parent) {
            var innerPanel = addPanel(parent, getLabel(LABELS.panel.innerArea));

            /* 単位はパネル名に入れているので入力欄には付けない
               The unit is in the panel title, so the fields carry none */
            var offsetPanel = addPanel(innerPanel, getLabel(LABELS.panel.offset) + unitSuffix());
            innerOffsetFields = buildLinkedQuadUI(offsetPanel, String(calcDefaultInnerOffset()), 3, "", LABELS.tooltip.offset);

            buildInnerGridPanels(innerPanel);
            buildLineTypePanel(innerPanel);

            /* 初期状態：列／行が1/1なら分割線はディム（OFF）
               Initially the dividers are dimmed while the grid is 1 by 1 */
            applyInnerDividerEnabledState(toCount(columnCountInput.text), toCount(rowCountInput.text), false);
        }

        /**
         * 内側エリアの列・行と、塗り／分割線のオプションを組み立てます。
         *
         * @param {Panel} parent - 追加先のパネル。
         * @returns {void}
         */
        function buildInnerGridPanels(parent) {
            var gridWrapper = parent.add("group");
            gridWrapper.orientation = "row";
            gridWrapper.alignChildren = ["left", "top"];
            gridWrapper.alignment = ["fill", "top"];

            var gridColumn = gridWrapper.add("group");
            gridColumn.orientation = "column";
            gridColumn.alignChildren = ["left", "top"];
            gridColumn.alignment = ["left", "top"];
            gridColumn.spacing = 12;

            var columnFields = addGridCountPanel(gridColumn, LABELS.panel.columns, LABELS.fieldLabel.columnCount);
            columnCountInput = columnFields.countInput;
            columnGutterInput = columnFields.gutterInput;

            var rowFields = addGridCountPanel(gridColumn, LABELS.panel.rows, LABELS.fieldLabel.rowCount);
            rowCountInput = rowFields.countInput;
            rowGutterInput = rowFields.gutterInput;

            // 塗り・分割線のオプション（中央寄せ）
            var optionWrapper = gridColumn.add("group");
            optionWrapper.orientation = "row";
            optionWrapper.alignChildren = ["center", "center"];
            optionWrapper.alignment = ["fill", "top"];

            var optionGroup = addRow(optionWrapper);
            optionGroup.alignment = ["center", "center"];

            innerFillCheckbox = optionGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.fill));
            innerFillCheckbox.helpTip = getLabel(LABELS.tooltip.innerFill);

            innerDividerCheckbox = optionGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.divider));
            innerDividerCheckbox.helpTip = getLabel(LABELS.tooltip.divider);

            bindAll([columnCountInput, rowCountInput], "onChanging", function () {
                /* 空欄のあいだは書き換えない（打ち直すたびに「1」が残って桁が増えるため）
                   Leave a blank field alone; rewriting it would prepend a 1 to the next digit */
                snapCountInput(columnCountInput);
                snapCountInput(rowCountInput);

                var columnCount = toCount(columnCountInput.text);
                var rowCount = toCount(rowCountInput.text);

                columnGutterInput.enabled = (columnCount > 1);
                rowGutterInput.enabled = (rowCount > 1);
                applyInnerDividerEnabledState(columnCount, rowCount, true);
                requestPreview();
            });

            bindAll([columnGutterInput, rowGutterInput], "onChanging", function () {
                /* ガターが入ったら塗りを自動ON（手動操作があれば尊重）
                   Turn the fill on once a gutter is set, unless the user set it manually */
                var gutter = parseFloat(this.text);
                if (!innerFillManuallySet && !isNaN(gutter) && gutter !== 0) {
                    innerFillCheckbox.value = true;
                }
                requestPreview();
            });

            innerFillCheckbox.onClick = function () {
                /* ユーザーが操作したら以後は自動ONしない / Stop auto-enabling once the user decides */
                innerFillManuallySet = true;
                requestPreview();
            };
        }

        /**
         * 内側エリアの［線の種類］パネルを組み立てます。
         *
         * @param {Panel} parent - 追加先のパネル。
         * @returns {void}
         */
        function buildLineTypePanel(parent) {
            lineTypePanel = addRadioPanel(parent, getLabel(LABELS.panel.lineType));

            lineSolidRadio = lineTypePanel.add("radiobutton", undefined, getLabel(LABELS.radio.lineSolid));
            lineDashRadio = lineTypePanel.add("radiobutton", undefined, getLabel(LABELS.radio.lineDash));
            lineDotsRadio = lineTypePanel.add("radiobutton", undefined, getLabel(LABELS.radio.lineDots));
            lineSolidRadio.value = true;

            bindAll([lineSolidRadio, lineDashRadio, lineDotsRadio], "onClick", requestPreview);

            innerDividerCheckbox.onClick = function () {
                lineTypePanel.enabled = (innerDividerCheckbox.enabled && innerDividerCheckbox.value);
                requestPreview();
            };
        }

        /**
         * ［画面表示］タブを組み立てます（ズームとパン、表示コマンド）。
         *
         * @param {Group} parent - 追加先のグループ。
         * @returns {void}
         */
        function buildDisplayTab(parent) {
            var zoomPanPanel = addPanel(parent, getLabel(LABELS.panel.zoomPan), 10);

            var zoomPanGroup = zoomPanPanel.add("group");
            zoomPanGroup.orientation = "column";
            zoomPanGroup.alignChildren = "left";
            zoomPanGroup.spacing = 8;

            viewControl.buildUI(zoomPanGroup, {
                labelWidth: DIALOG_LAYOUT.viewLabelWidth,
                sliderWidth: DIALOG_LAYOUT.viewSliderWidth,
                zoomLabel: labelText(LABELS.fieldLabel.zoom),
                panXLabel: labelText(LABELS.fieldLabel.panX),
                panYLabel: labelText(LABELS.fieldLabel.panY),
                sliderHelpTip: getLabel(LABELS.tooltip.viewSlider)
            });

            var viewCommandPanel = addPanel(parent, getLabel(LABELS.panel.viewCommands), 10);

            var viewCommandGroup = viewCommandPanel.add("group");
            viewCommandGroup.orientation = "column";
            viewCommandGroup.alignChildren = ["left", "top"];
            viewCommandGroup.alignment = ["fill", "top"];
            viewCommandGroup.spacing = 6;

            addViewCommandButton(viewCommandGroup, LABELS.button.fitArtboard, function () {
                app.executeMenuCommand("fitin");
            });
            addViewCommandButton(viewCommandGroup, LABELS.button.actualSize, function () {
                app.executeMenuCommand("actualsize");
            });
            addViewCommandButton(viewCommandGroup, LABELS.button.fitAll, function () {
                app.executeMenuCommand("fitall");
            });
            addViewCommandButton(viewCommandGroup, LABELS.button.zoomOut10, function () {
                /* 現在の表示倍率を10%縮小（スライダーも追従）
                   Shrink the current zoom by 10%; the slider follows */
                viewControl.zoomBy(0.9);
            });
        }

        // =========================================
        // セッションの保存と復元 / Save and restore the session
        // =========================================
        /* 保存と復元はこの定義表を共有します（項目の追加・変更はここだけ）
           Save and restore share these tables, so a field is defined in one place */

        /**
         * セッションに保存する入力欄・チェックボックスの一覧を返します。
         *
         * @returns {Array} [コントロール, 保存先のパス] の配列。
         */
        function getSessionControls() {
            return [
                [marginFields.top, "margin.top"],
                [marginFields.bottom, "margin.bottom"],
                [marginFields.left, "margin.left"],
                [marginFields.right, "margin.right"],
                [marginFields.linkCheckbox, "margin.link"],
                [frameCheckbox, "frame.enabled"],
                [frameWidthInput, "frame.width"],
                [bleedCheckbox, "frame.bleed"],
                [frameRoundCheckbox, "frame.round.enabled"],
                [frameRoundInput, "frame.round.value"],
                [keepOuterCheckbox, "outer.keepOuter"],
                [outerRoundCheckbox, "outer.round.enabled"],
                [outerRoundInput, "outer.round.value"],
                [outerEdgeScaleCheckbox, "outer.edgeScale.enabled"],
                [outerEdgeScaleInput, "outer.edgeScale.value"],
                [titleCheckbox, "title.enabled"],
                [titleSizeInput, "title.size"],
                [titleFillCheckbox, "title.fill"],
                [titleLineCheckbox, "title.line"],
                [titleEdgeScaleCheckbox, "title.edgeScale.enabled"],
                [titleEdgeScaleInput, "title.edgeScale.value"],
                [innerOffsetFields.top, "inner.offset.top"],
                [innerOffsetFields.bottom, "inner.offset.bottom"],
                [innerOffsetFields.left, "inner.offset.left"],
                [innerOffsetFields.right, "inner.offset.right"],
                [innerOffsetFields.linkCheckbox, "inner.offset.link"],
                [columnCountInput, "inner.grid.columns"],
                [columnGutterInput, "inner.grid.columnGutter"],
                [rowCountInput, "inner.grid.rows"],
                [rowGutterInput, "inner.grid.rowGutter"],
                [innerFillCheckbox, "inner.grid.fill"],
                [innerDividerCheckbox, "inner.grid.divider"],
                [previewCheckbox, "preview"]
            ];
        }

        /**
         * セッションに保存するラジオボタン群の一覧を返します。
         *
         * @returns {Array} [保存先のパス, 保存値→コントロールの対応] の配列。
         */
        function getSessionRadioGroups() {
            return [
                ["outer.strokeCap", { butt: capButtRadio, round: capRoundRadio, project: capProjectRadio }],
                ["title.position", { top: titleTopRadio, bottom: titleBottomRadio, left: titleLeftRadio, right: titleRightRadio }],
                ["inner.grid.lineType", { solid: lineSolidRadio, dash: lineDashRadio, dots: lineDotsRadio }]
            ];
        }

        /**
         * ラジオボタン群から、選択中の保存値を返します。
         *
         * @param {Object} radiosByKey - 保存値→コントロールの対応。
         * @returns {string|undefined} 選択中の保存値。どれも選択されていなければ undefined。
         */
        function readRadioKey(radiosByKey) {
            for (var key in radiosByKey) {
                if (!radiosByKey.hasOwnProperty(key)) continue;
                if (radiosByKey[key].value) return key;
            }
            return undefined;
        }

        /**
         * 前回のダイアログ設定を復元します（Illustratorの起動中のみ有効）。
         *
         * @returns {void}
         */
        function restoreDialogState() {
            var state = loadSessionState();
            var controls = getSessionControls();
            var radioGroups = getSessionRadioGroups();
            var i, value;

            for (i = 0; i < controls.length; i++) {
                value = getStateValue(state, controls[i][1]);
                if (typeof value === "undefined") continue;

                if (controls[i][0].type === "edittext") controls[i][0].text = String(value);
                else controls[i][0].value = !!value;
            }

            for (i = 0; i < radioGroups.length; i++) {
                var radio = radioGroups[i][1][getStateValue(state, radioGroups[i][0])];
                if (radio) radio.value = true;
            }

            /* 手動操作のフラグも復元する（復元直後に自動ONで上書きされないように）
               Restore the manual-input flags too, so the auto-on rules do not override them */
            innerFillManuallySet = !!getStateValue(state, "manual.innerFill");
            bleedManuallySet = !!getStateValue(state, "manual.bleed");

            /* 角丸と辺の伸縮は同時にONにできない（角丸を優先）
               The radius and the edge scale cannot both be on; the radius wins */
            if (outerRoundCheckbox.value) outerEdgeScaleCheckbox.value = false;

            /* 復元は「サイズが0→>0になった瞬間」ではないため、
               タイトルの［線］が自動ONで上書きされないように直前の状態をそろえる
               A restore is not a zero-to-positive transition, so seed the previous state */
            titleHadSize = (titleCheckbox.value && hasPositiveValue(titleSizeInput));

            /* 他のコントロールに依存する有効／無効を反映し直す
               Re-apply the enabled states that depend on other controls */
            marginFields.applyLinkState();
            innerOffsetFields.applyLinkState();
            applyOuterAreaEnabledState();
            applyTitleEdgeScaleEnabledState();
            applyTitleAreaEnabledState();
            applyFrameEnabledState();
            applyInnerDividerEnabledState(toCount(columnCountInput.text), toCount(rowCountInput.text), false);
        }

        /**
         * 現在のダイアログ設定をセッションに保存します。
         *
         * @returns {void}
         */
        function saveDialogState() {
            var state = {};
            var controls = getSessionControls();
            var radioGroups = getSessionRadioGroups();
            var i;

            for (i = 0; i < controls.length; i++) {
                setStateValue(state, controls[i][1],
                    (controls[i][0].type === "edittext") ? controls[i][0].text : controls[i][0].value);
            }

            for (i = 0; i < radioGroups.length; i++) {
                setStateValue(state, radioGroups[i][0], readRadioKey(radioGroups[i][1]));
            }

            /* 手動操作のフラグ（次回の自動ONを抑えるために保存）
               The manual-input flags, saved so the auto-on rules stay suppressed next time */
            setStateValue(state, "manual.innerFill", innerFillManuallySet);
            setStateValue(state, "manual.bleed", bleedManuallySet);

            saveSessionState(state);
        }

        // =========================================
        // プレビューと生成 / Preview and generation
        // =========================================
        // collectOptions()        : UIを読み、pt単位の生成条件にまとめる
        // generateFromOptions()   : 生成条件からオブジェクトを作る
        // rebuildGeneratedItems() : 前回分を消して生成し、再描画する（唯一の入口）
        // -----------------------------------------

        /**
         * プレビューがONのときだけプレビューを更新します。
         *
         * @returns {void}
         */
        function requestPreview() {
            if (previewCheckbox.value) rebuildGeneratedItems(false);
        }

        /**
         * プレビューまたは最終結果を作り直して再描画します。
         *
         * @param {boolean} isFinal - 実行（OK）時の最終生成なら true。
         * @returns {void}
         */
        function rebuildGeneratedItems(isFinal) {
            removeGeneratedItems();
            generateFromOptions(collectOptions(), isFinal);
            app.redraw();
        }

        /**
         * プレビューを消して元の状態に戻します。
         *
         * @returns {void}
         */
        function clearPreview() {
            removeGeneratedItems();

            /* アートボード基準の一時矩形を先に破棄（baseRectsに残っていると無効参照になる）
               Remove the temporary artboard rectangle first to avoid stale references */
            removeArtboardBaseRect();

            /* 角丸プレビューなどで隠した元オブジェクトを表示に戻す
               Show the originals that the preview had hidden */
            setBaseRectsHidden(false);

            app.redraw();
        }

        /**
         * UIの入力値を読み、pt単位の生成条件にまとめます。
         *
         * @returns {Object} 生成条件。
         */
        function collectOptions() {
            var columnCount = toCount(columnCountInput.text);
            var rowCount = toCount(rowCountInput.text);
            var bleedPt = bleedCheckbox.value ? mmToPt(GENERATION_SETTINGS.bleedMm) : 0;

            var options = {
                /* マージン（アートボード基準のみ） / Margins, artboard-based runs only */
                marginPt: {
                    top: toPositivePt(marginFields.top.text),
                    right: toPositivePt(marginFields.right.text),
                    bottom: toPositivePt(marginFields.bottom.text),
                    left: toPositivePt(marginFields.left.text)
                },

                /* フレーム：幅に裁ち落としを加算し、アートボードも裁ち落としぶん広げる
                   Frame: the bleed widens both the frame and the artboard bounds */
                bleedPt: bleedPt,
                framePt: (frameCheckbox.value ? toPositivePt(frameWidthInput.text) : 0) + bleedPt,
                frameRoundPt: readCheckedPt(frameRoundCheckbox, frameRoundInput),

                /* 外側エリア：辺の伸縮と角丸（角丸はタイトル帯にも使う）
                   Outer area: the edge scale and the radius, which the title band reuses */
                outerEdgeScalePt: getOuterEdgeScaleValue() * rulerUnit.factor,
                outerRoundPt: readCheckedPt(outerRoundCheckbox, outerRoundInput),

                /* タイトルエリア：仕切り線の詰め量は入力値の符号を反転（＋で両端が短くなる）
                   Title area: the divider inset is the negated input (positive shortens the line) */
                titleSizePt: titleCheckbox.value ? toPositivePt(titleSizeInput.text) : 0,
                titleDividerInsetPt: (titleLineCheckbox.value && titleEdgeScaleCheckbox.value) ? -toPt(titleEdgeScaleInput.text) : 0,

                /* 内側エリア：オフセット / Inner area offsets */
                innerOffsetPt: {
                    top: toPositivePt(innerOffsetFields.top.text),
                    right: toPositivePt(innerOffsetFields.right.text),
                    bottom: toPositivePt(innerOffsetFields.bottom.text),
                    left: toPositivePt(innerOffsetFields.left.text)
                },

                /* 内側エリア：列・行とガター（列／行が1のときガターは0扱い）
                   Columns, rows and gutters; gutters are ignored for a single column or row */
                columnCount: columnCount,
                rowCount: rowCount,
                columnGutterPt: (columnCount > 1) ? toPositivePt(columnGutterInput.text) : 0,
                rowGutterPt: (rowCount > 1) ? toPositivePt(rowGutterInput.text) : 0
            };

            syncDependentControls(options);
            return options;
        }

        /**
         * 生成条件からオブジェクトを作成します（プレビュー・実行の共通処理）。
         *
         * @param {Object} options - collectOptions() が返す生成条件。
         * @param {boolean} isFinal - 実行（OK）時の最終生成なら true。
         * @returns {void}
         */
        function generateFromOptions(options, isFinal) {
            if (isArtboardBased) rebuildArtboardBaseRect(options.marginPt);

            var splitEdges = (options.outerEdgeScalePt !== 0);
            var i;

            if (splitEdges) {
                /* 4辺に分解するため、元の長方形は常に隠す
                   Always hide the original rectangle while the edges are split */
                setBaseRectsHidden(true);
                if (keepOuterCheckbox.value) {
                    for (i = 0; i < baseRects.length; i++) {
                        createOuterEdgeLines(baseRects[i], options.outerEdgeScalePt);
                    }
                }
            } else {
                /* 外枠の表示は「外枠を残す」に従う
                   Show the original frame according to the keep-outer checkbox */
                setBaseRectsHidden(!keepOuterCheckbox.value);
            }

            /* フレーム（アートボード基準） / Frame, based on the artboard */
            if (options.framePt > 0 && baseRects.length > 0) {
                createFrame(baseRects[0].layer, getActiveArtboardBounds(options.bleedPt), options.framePt, options.frameRoundPt);
            }

            if (options.titleSizePt > 0) {
                for (i = 0; i < baseRects.length; i++) {
                    createTitleArea(baseRects[i], options);
                }
            }

            if (!splitEdges) applyOuterAreaRound(options.outerRoundPt, isFinal);

            for (i = 0; i < baseRects.length; i++) {
                createInnerArea(baseRects[i], options);
            }
        }

        /**
         * アクティブなアートボードの矩形を返します（裁ち落としのぶん外側に広げます）。
         *
         * @param {number} bleedPt - 裁ち落とし幅（pt）。
         * @returns {number[]} [左, 上, 右, 下] の座標。
         */
        function getActiveArtboardBounds(bleedPt) {
            var rect = getActiveArtboardRect(); // [L, T, R, B]
            return [rect[0] - bleedPt, rect[1] + bleedPt, rect[2] + bleedPt, rect[3] - bleedPt];
        }

        /**
         * 長方形の各辺を、伸縮させた直線として生成します（4辺に分解）。
         *
         * @param {PathItem} baseRect - 基準の長方形。
         * @param {number} edgeScalePt - 伸縮量（pt。正で伸ばし、負で縮める）。
         * @returns {void}
         */
        function createOuterEdgeLines(baseRect, edgeScalePt) {
            var points = baseRect.pathPoints;
            var edgeCount = baseRect.closed ? points.length : points.length - 1;
            var scaleAmount = Math.abs(edgeScalePt);

            /* 正なら両端を外へ、負なら内へ動かす / Positive extends the ends, negative pulls them in */
            var direction = (edgeScalePt >= 0) ? -1 : 1;

            for (var i = 0; i < edgeCount; i++) {
                var startAnchor = points[i].anchor;
                var endAnchor = points[(i + 1) % points.length].anchor;

                var dx = endAnchor[0] - startAnchor[0];
                var dy = endAnchor[1] - startAnchor[1];
                var edgeLength = Math.sqrt(dx * dx + dy * dy);

                /* 長さ0の辺（重なったアンカー）は計算できないので飛ばす
                   A zero-length edge (duplicated anchors) cannot be scaled */
                if (!(edgeLength > 0)) continue;

                /* 縮めるとき、辺が縮小量の2倍以下なら線が反転するので生成しない
                   While shrinking, an edge shorter than twice the amount would turn inside out */
                if (edgeScalePt < 0 && edgeLength <= scaleAmount * 2) continue;

                var ratio = scaleAmount / edgeLength;
                var edgeLine = trackGeneratedItem(baseRect.layer.pathItems.add());
                edgeLine.setEntirePath([
                    [startAnchor[0] + dx * ratio * direction, startAnchor[1] + dy * ratio * direction],
                    [endAnchor[0] - dx * ratio * direction, endAnchor[1] - dy * ratio * direction]
                ]);

                setStroke(edgeLine, baseRect.strokeColor, baseRect.strokeWidth);
                edgeLine.strokeCap = getSelectedStrokeCap();

                // 外枠（4辺）として識別できるようタグ付け
                tagItem(edgeLine, TAG_OUTER_EDGE);
            }
        }

        /**
         * 外側エリアの角丸を適用します（辺の伸縮OFFのときだけ）。
         *
         * プレビューでは元の長方形を隠し、同じ位置に角丸用の一時矩形を作ります。
         * 実行時は元の長方形にライブエフェクトを適用します。
         *
         * @param {number} radiusPt - 角丸の半径（pt）。
         * @param {boolean} isFinal - 実行（OK）時の最終生成なら true。
         * @returns {void}
         */
        function applyOuterAreaRound(radiusPt, isFinal) {
            if (!keepOuterCheckbox.value || outerEdgeScaleCheckbox.value || !(radiusPt > 0)) return;

            for (var i = 0; i < baseRects.length; i++) {
                var baseRect = baseRects[i];

                if (isFinal) {
                    applyRoundCornersEffect(baseRect, radiusPt);
                    baseRect.hidden = false;
                    continue;
                }

                var bounds = baseRect.geometricBounds; // [L, T, R, B]
                var width = bounds[2] - bounds[0];
                var height = bounds[1] - bounds[3];
                if (!(width > 0) || !(height > 0)) continue;

                baseRect.hidden = true;

                var previewRect = trackGeneratedItem(baseRect.layer.pathItems.rectangle(bounds[1], bounds[0], width, height));
                copyAppearance(baseRect, previewRect);
                tagItem(previewRect, TAG_OUTER_ROUND);
                applyRoundCornersEffect(previewRect, radiusPt);
            }
        }

        /**
         * フレーム（外側はグレー、内側は透明の穴）を作成します。
         *
         * 穴あきはグループに Live Pathfinder Exclude を適用して作ります。
         *
         * @param {Layer} layer - 作成先のレイヤー。
         * @param {number[]} bounds - 基準領域 [左, 上, 右, 下]。
         * @param {number} framePt - フレームの幅（pt）。
         * @param {number} roundPt - 穴側の角丸の半径（pt）。
         * @returns {void}
         */
        function createFrame(layer, bounds, framePt, roundPt) {
            var left = bounds[0], top = bounds[1], right = bounds[2], bottom = bounds[3];
            var width = right - left;
            var height = top - bottom;

            /* 内側は最終的に穴になる矩形 / The inner rectangle becomes the hole */
            var holeWidth = width - framePt * 2;
            var holeHeight = height - framePt * 2;
            if (!(holeWidth > 0) || !(holeHeight > 0)) return;

            /* パスファインダーはメニューコマンド経由のため、失敗しても続行する
               The pathfinder runs as a menu command, so keep going if it fails */
            try {
                var outerRect = trackGeneratedItem(layer.pathItems.rectangle(top, left, width, height));
                setGrayFill(outerRect, GRAY_TINTS.frame);

                var holeRect = trackGeneratedItem(layer.pathItems.rectangle(top - framePt, left + framePt, holeWidth, holeHeight));
                /* 内側は一時的な塗り（最終的に Exclude の結果で穴になる）
                   A temporary fill; the Exclude result turns it into a hole */
                setGrayFill(holeRect, 0);

                // 角丸は「内側の長方形」に適用する（穴側を丸める）
                applyRoundCornersEffect(holeRect, roundPt);

                var frameGroup = trackGeneratedItem(layer.groupItems.add());
                outerRect.move(frameGroup, ElementPlacement.PLACEATEND);
                holeRect.move(frameGroup, ElementPlacement.PLACEATEND);

                var frameItem = applyPathfinderExclude(frameGroup);
                if (frameItem !== frameGroup) trackGeneratedItem(frameItem);

                tagItem(frameItem, TAG_FRAME_FILL);
                sendToBack(frameItem);
            } catch (e) { }
        }

        /**
         * グループに Live Pathfinder Exclude を適用し、結果のオブジェクトを返します。
         *
         * 選択を一時的に置き換えるため、実行後に元の選択へ戻します。
         *
         * @param {GroupItem} group - 対象のグループ。
         * @returns {PageItem} Exclude の結果（取得できない場合は元のグループ）。
         */
        function applyPathfinderExclude(group) {
            var previousSelection = doc.selection;
            var resultItem = group;

            try {
                doc.selection = null;
                group.selected = true;
                app.executeMenuCommand('Live Pathfinder Exclude');

                /* 結果は selection の先頭に入る / The result lands at the head of the selection */
                if (doc.selection.length > 0) resultItem = doc.selection[0];
            } catch (e) { }

            try { doc.selection = previousSelection; } catch (e) { }

            return resultItem;
        }

        /**
         * 選択中のタイトルの位置を返します。
         *
         * @returns {string} "top" / "bottom" / "left" / "right"。
         */
        function getTitlePositionKey() {
            if (titleRightRadio.value) return "right";
            if (titleBottomRadio.value) return "bottom";
            if (titleLeftRadio.value) return "left";
            return "top";
        }

        /**
         * タイトルエリアの配置を計算します（位置による分岐をここに集約）。
         *
         * @param {number[]} bounds - 基準領域 [左, 上, 右, 下]。
         * @param {number} sizePt - タイトルエリアの幅／高さ（pt）。
         * @param {number} dividerInsetPt - 仕切り線の両端の詰め量（pt。＋で短く／−で長く）。
         * @returns {Object|null} band（帯の矩形）／divider（仕切り線）／inner（帯を除いた領域）。
         *                        領域が成立しない場合は null。
         */
        function calcTitleAreaLayout(bounds, sizePt, dividerInsetPt) {
            var left = bounds[0], top = bounds[1], right = bounds[2], bottom = bounds[3];
            var width = right - left;
            var height = top - bottom;
            if (!(width > 0) || !(height > 0)) return null;

            var positionKey = getTitlePositionKey();
            var isHorizontal = (positionKey === "top" || positionKey === "bottom");
            if (sizePt >= (isHorizontal ? height : width)) return null;

            /* 横並び（上／下）：帯は横いっぱい、仕切り線は水平
               Top or bottom: a full-width band with a horizontal divider */
            if (isHorizontal) {
                var dividerY = (positionKey === "top") ? (top - sizePt) : (bottom + sizePt);

                var startX = left + dividerInsetPt;
                var endX = right - dividerInsetPt;
                if (startX >= endX) { startX = left; endX = right; }

                return {
                    band: {
                        top: (positionKey === "top") ? top : (bottom + sizePt),
                        left: left,
                        width: width,
                        height: sizePt
                    },
                    divider: [[startX, dividerY], [endX, dividerY]],
                    inner: (positionKey === "top") ? [left, dividerY, right, bottom] : [left, top, right, dividerY]
                };
            }

            /* 縦並び（左／右）：帯は縦いっぱい、仕切り線は垂直
               Left or right: a full-height band with a vertical divider */
            var dividerX = (positionKey === "left") ? (left + sizePt) : (right - sizePt);

            var startY = top - dividerInsetPt;
            var endY = bottom + dividerInsetPt;
            if (startY <= endY) { startY = top; endY = bottom; }

            return {
                band: {
                    top: top,
                    left: (positionKey === "left") ? left : (right - sizePt),
                    width: sizePt,
                    height: height
                },
                divider: [[dividerX, startY], [dividerX, endY]],
                inner: (positionKey === "left") ? [dividerX, top, right, bottom] : [left, top, dividerX, bottom]
            };
        }

        /**
         * タイトル帯の塗りと、本文との仕切り線を作成します。
         *
         * @param {PathItem} baseRect - 基準の長方形。
         * @param {Object} options - collectOptions() が返す生成条件。
         * @returns {void}
         */
        function createTitleArea(baseRect, options) {
            var layout = calcTitleAreaLayout(baseRect.geometricBounds, options.titleSizePt, options.titleDividerInsetPt);
            if (!layout) return;

            var layer = baseRect.layer;

            if (titleFillCheckbox.value) {
                var bandRect = trackGeneratedItem(layer.pathItems.rectangle(
                    layout.band.top, layout.band.left, layout.band.width, layout.band.height));
                setGrayFill(bandRect, GRAY_TINTS.titleBand);
                tagItem(bandRect, TAG_TITLE_FILL);

                /* 外側エリアの角丸値で、位置に応じた2角だけ角丸にする
                   Round the two corners that match the title position, with the outer radius */
                roundCornersOnSide(bandRect, getTitlePositionKey(), options.outerRoundPt);

                /* 背面へ（他の罫線や要素の下に敷く） / Send behind the rules */
                sendToBack(bandRect);
            }

            if (titleLineCheckbox.value) {
                /* 他の生成物と同じレイヤーに作る（doc.activeLayer は使わない）
                   Create it on the same layer as the other generated items */
                var dividerLine = trackGeneratedItem(layer.pathItems.add());
                dividerLine.setEntirePath(layout.divider);
                setStroke(dividerLine, makeGrayColor(GRAY_TINTS.rule), 1);
                tagItem(dividerLine, TAG_TITLE_DIVIDER);
            }
        }

        /**
         * 長方形の、指定した辺に接する2角だけを角丸にします（同じパスを書き換えます）。
         *
         * @param {PathItem} rect - 対象の長方形（閉じたパス）。
         * @param {string} sideKey - "top" / "bottom" / "left" / "right"。
         * @param {number} radiusPt - 角丸の半径（pt）。0以下なら何もしません。
         * @returns {void}
         */
        function roundCornersOnSide(rect, sideKey, radiusPt) {
            var bounds = rect.geometricBounds; // [L, T, R, B]
            var left = bounds[0], top = bounds[1], right = bounds[2], bottom = bounds[3];

            /* 半径は辺の半分までに抑える / Cap the radius at half the shorter side */
            var radius = Math.min(radiusPt, (right - left) / 2, (top - bottom) / 2);
            if (!(radius > 0)) return;

            /* 四分円をベジェで近似するハンドル長 / Bezier handle length for a quarter circle */
            var handleLength = radius * 0.5522847498307936;
            var POINT_TYPE = (typeof PointType !== "undefined") ? PointType : { CORNER: 0, SMOOTH: 1 };

            var roundTopLeft = (sideKey === "top" || sideKey === "left");
            var roundTopRight = (sideKey === "top" || sideKey === "right");
            var roundBottomRight = (sideKey === "bottom" || sideKey === "right");
            var roundBottomLeft = (sideKey === "bottom" || sideKey === "left");

            var cornerPoints = [];

            /**
             * アンカーと左右のハンドルを1点ぶん記録します（ハンドル省略時はアンカーと同じ位置）。
             *
             * @param {number[]} anchor - アンカーの座標 [x, y]。
             * @param {number[]} [leftDirection] - 左方向ハンドルの座標。
             * @param {number[]} [rightDirection] - 右方向ハンドルの座標。
             * @returns {void}
             */
            function addPoint(anchor, leftDirection, rightDirection) {
                cornerPoints.push({
                    anchor: anchor,
                    leftDirection: leftDirection || anchor,
                    rightDirection: rightDirection || anchor,
                    pointType: (leftDirection || rightDirection) ? POINT_TYPE.SMOOTH : POINT_TYPE.CORNER
                });
            }

            // 始点：左上（上辺側）/ Start at the top-left corner, on the top edge
            if (roundTopLeft) addPoint([left + radius, top], [left + radius - handleLength, top], null);
            else addPoint([left, top], null, null);

            // 右上 / Top-right
            if (roundTopRight) {
                addPoint([right - radius, top], null, [right - radius + handleLength, top]);
                addPoint([right, top - radius], [right, top - radius + handleLength], null);
            } else {
                addPoint([right, top], null, null);
            }

            // 右下 / Bottom-right
            if (roundBottomRight) {
                addPoint([right, bottom + radius], null, [right, bottom + radius - handleLength]);
                addPoint([right - radius, bottom], [right - radius + handleLength, bottom], null);
            } else {
                addPoint([right, bottom], null, null);
            }

            // 左下 / Bottom-left
            if (roundBottomLeft) {
                addPoint([left + radius, bottom], null, [left + radius - handleLength, bottom]);
                addPoint([left, bottom + radius], [left, bottom + radius - handleLength], null);
            } else {
                addPoint([left, bottom], null, null);
            }

            // 終点：左上（左辺側）/ Close at the top-left corner, on the left edge
            if (roundTopLeft) addPoint([left, top - radius], null, [left, top - radius + handleLength]);

            var anchors = [];
            for (var i = 0; i < cornerPoints.length; i++) {
                anchors.push(cornerPoints[i].anchor);
            }
            rect.setEntirePath(anchors);
            rect.closed = true;

            var pathPoints = rect.pathPoints;
            for (var j = 0; j < cornerPoints.length; j++) {
                pathPoints[j].leftDirection = cornerPoints[j].leftDirection;
                pathPoints[j].rightDirection = cornerPoints[j].rightDirection;
                pathPoints[j].pointType = cornerPoints[j].pointType;
            }
        }

        /**
         * タイトル領域を除いた「内側エリア」の計算領域を返します。
         *
         * @param {PathItem} baseRect - 基準の長方形。
         * @param {number} titleSizePt - タイトルエリアの幅／高さ（pt）。
         * @returns {number[]|null} [左, 上, 右, 下] の座標。成立しない場合は null。
         */
        function getInnerAreaBounds(baseRect, titleSizePt) {
            var bounds = baseRect.geometricBounds; // [L, T, R, B]
            if (!((bounds[2] - bounds[0]) > 0 && (bounds[1] - bounds[3]) > 0)) return null;

            if (!(titleSizePt > 0)) return bounds;

            /* タイトルが入らないサイズのときは帯を作らないので、内側エリアは外形いっぱい
               When the title does not fit, no band is drawn, so the inner area keeps the full bounds */
            var layout = calcTitleAreaLayout(bounds, titleSizePt, 0);
            return layout ? layout.inner : bounds;
        }

        /**
         * 内側エリアのグリッド配置を計算します。
         *
         * @param {number[]} bounds - 基準領域 [左, 上, 右, 下]。
         * @param {Object} options - collectOptions() が返す生成条件。
         * @returns {Object|null} セル配置。成立しない場合は null。
         */
        function calcInnerGrid(bounds, options) {
            var offset = options.innerOffsetPt;

            /* オフセットが0でも内側エリアは描画する / The inner area is drawn even at zero offset */
            var width = (bounds[2] - bounds[0]) - (offset.left + offset.right);
            var height = (bounds[1] - bounds[3]) - (offset.top + offset.bottom);
            if (!(width > 0) || !(height > 0)) return null;

            /* ガターを除いた1セルの大きさ / Cell size with the gutters removed */
            var cellWidth = (width - options.columnGutterPt * (options.columnCount - 1)) / options.columnCount;
            var cellHeight = (height - options.rowGutterPt * (options.rowCount - 1)) / options.rowCount;
            if (!(cellWidth > 0) || !(cellHeight > 0)) return null;

            return {
                left: bounds[0] + offset.left,
                top: bounds[1] - offset.top,
                width: width,
                height: height,
                columnCount: options.columnCount,
                rowCount: options.rowCount,
                columnGutter: options.columnGutterPt,
                rowGutter: options.rowGutterPt,
                cellWidth: cellWidth,
                cellHeight: cellHeight
            };
        }

        /**
         * 内側エリアのセル（塗り）と分割線を作成します。
         *
         * @param {PathItem} baseRect - 基準の長方形（レイヤーと線の色の参照元）。
         * @param {Object} options - collectOptions() が返す生成条件。
         * @returns {void}
         */
        function createInnerArea(baseRect, options) {
            var areaBounds = getInnerAreaBounds(baseRect, options.titleSizePt);
            if (!areaBounds) return;

            var grid = calcInnerGrid(areaBounds, options);
            if (!grid) return;

            var layer = baseRect.layer;

            /* セルの塗りは［塗り］がONのときだけ（プレビューと実行結果を一致させる）
               The cells are only drawn while the fill is on, so the preview matches the result */
            if (innerFillCheckbox.value) createInnerCellFills(layer, grid);

            /* 分割線OFF、または分割できない構成ならここまで
               Stop here when the dividers are off or the grid cannot carry them */
            if (!innerDividerCheckbox.value || !isGridSplittable(grid.columnCount, grid.rowCount)) return;

            /* 分割線の色は基準の長方形から取る（線がなければ K100）
               Take the divider colour from the base rectangle, falling back to K100 */
            createInnerDividers(layer, grid, baseRect.stroked ? baseRect.strokeColor : makeGrayColor(GRAY_TINTS.rule));
        }

        /**
         * 内側エリアのセル（塗り）を作成します。
         *
         * @param {Layer} layer - 作成先のレイヤー。
         * @param {Object} grid - calcInnerGrid() が返すセル配置。
         * @returns {void}
         */
        function createInnerCellFills(layer, grid) {
            for (var i = 0; i < grid.rowCount; i++) {
                var cellTop = grid.top - (grid.cellHeight + grid.rowGutter) * i;

                for (var j = 0; j < grid.columnCount; j++) {
                    var cellLeft = grid.left + (grid.cellWidth + grid.columnGutter) * j;

                    var cellRect = trackGeneratedItem(layer.pathItems.rectangle(cellTop, cellLeft, grid.cellWidth, grid.cellHeight));
                    setGrayFill(cellRect, GRAY_TINTS.innerCell);
                    tagItem(cellRect, TAG_INNER_FILL);

                    /* 背面へ（罫線などのパスの下に敷く） / Send behind the rules */
                    sendToBack(cellRect);
                }
            }
        }

        /**
         * 内側エリアの分割線を、各ガターの中心に1本ずつ作成します。
         *
         * @param {Layer} layer - 作成先のレイヤー。
         * @param {Object} grid - calcInnerGrid() が返すセル配置。
         * @param {Color} strokeColor - 分割線の色。
         * @returns {void}
         */
        function createInnerDividers(layer, grid, strokeColor) {
            var bottomY = grid.top - grid.height;
            var rightX = grid.left + grid.width;

            /**
             * 分割線を1本作成します。
             *
             * @param {number[][]} points - [[x1, y1], [x2, y2]] 形式の始点と終点。
             * @returns {void}
             */
            function addDivider(points) {
                var dividerLine = trackGeneratedItem(layer.pathItems.add());
                dividerLine.setEntirePath(points);
                setStroke(dividerLine, strokeColor, 1);
                applyInnerLineStyle(dividerLine);
            }

            // 列の分割線 / Column dividers
            for (var i = 1; i < grid.columnCount; i++) {
                var gutterCenterX = grid.left + (grid.cellWidth * i) + (grid.columnGutter * (i - 1)) + (grid.columnGutter / 2);
                addDivider([[gutterCenterX, grid.top], [gutterCenterX, bottomY]]);
            }

            // 行の分割線 / Row dividers
            for (var j = 1; j < grid.rowCount; j++) {
                var gutterCenterY = grid.top - (grid.cellHeight * j) - (grid.rowGutter * (j - 1)) - (grid.rowGutter / 2);
                addDivider([[grid.left, gutterCenterY], [rightX, gutterCenterY]]);
            }
        }

        /**
         * UIで選択された線端を返します。
         *
         * @returns {StrokeCap} 選択中の線端。
         */
        function getSelectedStrokeCap() {
            if (capRoundRadio.value) return StrokeCap.ROUNDENDCAP;
            if (capProjectRadio.value) return StrokeCap.PROJECTINGENDCAP;
            return StrokeCap.BUTTENDCAP; // 線端なし
        }

        /**
         * 内側の分割線に線種（実線・点線・ドット点線）を適用します。
         *
         * @param {PathItem} dividerLine - 対象の分割線。
         * @returns {void}
         */
        function applyInnerLineStyle(dividerLine) {
            if (lineDotsRadio.value) {
                /* ドット点線：線端を丸型にして strokeDashes=[0, 線幅*2]
                   Dotted line: round caps with strokeDashes = [0, width * 2] */
                dividerLine.strokeWidth = 2;
                dividerLine.strokeCap = StrokeCap.ROUNDENDCAP;
                dividerLine.strokeDashes = [0, dividerLine.strokeWidth * 2];
                try { dividerLine.strokeJoin = StrokeJoin.ROUNDENDJOIN; } catch (e) { }
                return;
            }

            dividerLine.strokeWidth = 1;
            // 点線（ダッシュ）は破線パターン、実線は破線なし
            dividerLine.strokeDashes = lineDashRadio.value ? [4, 2] : [];
            // 線端は外枠の線端設定に合わせる
            dividerLine.strokeCap = getSelectedStrokeCap();
        }

        /**
         * 実行（OK）後の後処理です。
         *
         * 基準の長方形と内側エリアの塗りの扱いを確定し、内部タグをレイヤーパネルから隠します。
         *
         * @returns {void}
         */
        function finalizeGeneratedItems() {
            /* 外枠を残し、4辺に分解していないときだけ元の長方形を表示する。
               それ以外は削除する（非表示のまま残すと、実行のたびに見えない残骸がたまる）
               Show the original only when it is kept and not replaced by the four edges */
            if (keepOuterCheckbox.value && getOuterEdgeScaleValue() === 0) setBaseRectsHidden(false);
            else removeBaseRects();

            /* 内部タグは note だけに残し、レイヤーパネルに出る name はクリアする
               Keep the tag in note only and clear the name shown in the Layers panel */
            clearTagNames(generatedItems);

            doc.selection = null;
        }

        // =========================================
        // ダイアログの組み立てと実行 / Build the dialog and run
        // =========================================

        /* アートボード基準のときは、先に基準の矩形を作る（初期値の計算にも使う）。
           ロックや非表示のレイヤーでは作成できないので、その場合は理由を知らせて終了する
           Create the base rectangle first; a locked or hidden layer makes it impossible, so report and stop */
        if (isArtboardBased) {
            rebuildArtboardBaseRect({ top: 0, right: 0, bottom: 0, left: 0 });
            if (!artboardBaseRect) {
                alert(getLabel(LABELS.alert.baseRectFailed));
                return;
            }
        }

        var dialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];
        dialog.spacing = 20;
        dialog.margins = 16;
        dialog.opacity = DIALOG_LAYOUT.dialogOpacity;

        /* 表示位置をずらす（生成物が隠れないように） / Shift the dialog so it does not cover the artwork */
        dialog.onShow = function () {
            dialog.location = [
                dialog.location[0] + DIALOG_LAYOUT.dialogOffsetX,
                dialog.location[1] + DIALOG_LAYOUT.dialogOffsetY
            ];
        };

        /* 4タブ：マージン（＋フレーム）／外側エリア（＋タイトルエリア）／内側エリア／画面表示
           Four tabs: margin and frame, outer and title, inner area, display */
        var settingsTabs = dialog.add("tabbedpanel");
        settingsTabs.alignChildren = ["fill", "top"];
        settingsTabs.alignment = ["fill", "top"];
        settingsTabs.margins = [5, 20, 0, 0];

        /* tabbedpanel は内容量に応じて自動で高さが伸びないことがあるため、最低サイズを与える
           A tabbedpanel does not always grow with its content, so give it a minimum size */
        settingsTabs.minimumSize = DIALOG_LAYOUT.tabSize;
        settingsTabs.preferredSize = DIALOG_LAYOUT.tabSize;

        var marginTabColumn = addTabColumn(settingsTabs, LABELS.tab.margin);
        var outerTabColumn = addTabColumn(settingsTabs, LABELS.tab.outer);
        var innerTabColumn = addTabColumn(settingsTabs, LABELS.tab.inner);
        var displayTabColumn = addTabColumn(settingsTabs, LABELS.tab.display);

        buildMarginPanel(marginTabColumn);
        buildFramePanel(marginTabColumn);
        buildOuterPanel(outerTabColumn);
        buildTitlePanel(outerTabColumn);
        buildInnerPanel(innerTabColumn);
        buildDisplayTab(displayTabColumn);

        /* 長方形スタート時（アートボード基準でない場合）は左タブ全体を非表示
           Hide the whole left tab for rectangle-based runs */
        if (!isArtboardBased) {
            marginTabColumn.parent.visible = false;
            marginTabColumn.parent.enabled = false;
            settingsTabs.selection = outerTabColumn.parent;
        }

        // =========================================
        // ボタンエリア（左：プレビュー ／ 右：ボタン） / Footer: preview on the left, buttons on the right
        // =========================================
        var btnRowGroup = dialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignChildren = ["left", "center"];
        btnRowGroup.alignment = ["fill", "center"];

        previewCheckbox = btnRowGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.preview));
        previewCheckbox.value = true; /* 最初からプレビューON / preview starts enabled */
        previewCheckbox.alignment = "left";

        /* スペーサー（右側のボタンを押し出す） / Spacer that pushes the buttons to the right */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignment = ["right", "center"];
        btnRightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        btnRightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        previewCheckbox.onClick = function () {
            if (previewCheckbox.value) rebuildGeneratedItems(false);
            else clearPreview();
        };

        restoreDialogState();

        // レイアウト確定（tabbedpanel の内容が潰れるのを防ぐ）
        dialog.layout.layout(true);
        dialog.layout.resize();

        /* プレビューがOFFなら描画しない。UIの依存関係だけは collectOptions() でそろえる
           Draw nothing while the preview is off; collectOptions() still syncs the dependent controls */
        if (previewCheckbox.value) rebuildGeneratedItems(false);
        else collectOptions();

        var dialogResult = dialog.show();
        saveDialogState();

        /* キャンセル時：生成物を削除し、選択パスの見た目と画面表示を元に戻して終了
           On cancel: drop the generated items and restore both the appearance and the view */
        if (dialogResult !== 1) {
            clearPreview();
            restoreSelectedAppearances();
            viewControl.restore();
            return;
        }

        /* 最終生成（この結果はヒストリーに残す） / Final generation, kept in the history */
        rebuildGeneratedItems(true);
        finalizeGeneratedItems();

    })();

})();
