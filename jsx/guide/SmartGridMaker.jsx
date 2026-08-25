#targetengine "SmartGridMakerEngine"
#target illustrator

/*

### 概要

長方形の選択、またはアートボードを基準に、囲み罫とグリッドを一括生成します。
外側エリア・タイトルエリア・内側エリアの分割や線種、裁ち落とし対応のフレームを、プレビューを見ながら1つのダイアログで設定できます。

詳細は README を参照してください。

### Overview

Generates an enclosing rule and a grid from a selected rectangle, or from the artboard.
The outer area, title area, inner-area divisions, stroke styles, and a bleed-aware frame are all set in one dialog with a live preview.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartGridMaker";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.6.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-02-24";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-07-31";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartGridMaker.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartGridMaker.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n2b01f896c423"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    /* 生成と既定値の設定 / Generation and default values */
    var GRID_CONFIG = {
        bleedMm: 3,               /* 裁ち落とし幅（mm） / bleed width in mm */
        innerOffsetDivisor: 40,   /* 内側オフセット初期値＝(幅+高さ)/この値 / inner offset default divisor */
        titleSizeDivisor: 5,      /* タイトルエリア初期値＝外側エリア高/この値 / title size default divisor */
        defaultMargin: "15",      /* マージンの初期値 / default margin */
        defaultEdgeScale: "-5",   /* 辺の伸縮の初期値 / default edge scale */
        defaultFrameWidth: "10",  /* フレームON時の既定幅 / frame width applied when enabled */
        defaultRound: "2"         /* 角丸ON時の既定値 / radius applied when rounding is enabled */
    };

    /* 生成物の塗り濃度（CMYKのK%） / Fill tints of generated items (K in CMYK) */
    var FILL_TINTS = {
        innerBox: 15,   /* 内側エリア / inner area */
        titleBand: 30,  /* タイトル帯 / title band */
        frame: 50,      /* フレーム / frame */
        line: 100       /* 罫線 / rules */
    };

    /* ダイアログとパネルの外観 / Dialog and panel appearance */
    var UI_CONFIG = {
        dialogOffsetX: 300,       /* ダイアログの表示位置オフセットX / dialog offset X */
        dialogOffsetY: 0,         /* ダイアログの表示位置オフセットY / dialog offset Y */
        dialogOpacity: 0.98,      /* ダイアログの不透明度 / dialog opacity */
        tabSize: [300, 460],      /* タブパネルの最小サイズ / minimum size of the tabbed panel */
        panelMargins: [15, 20, 15, 10], /* パネル余白 [左,上,右,下] / panel margins */
        viewLabelWidth: 58,       /* 画面表示タブのラベル幅 / label width in the Display tab */
        viewSliderWidth: 200,     /* 画面表示タブのスライダー幅 / slider width in the Display tab */
        viewButtonWidth: 190,     /* 表示コマンドボタンの幅 / view command button width */
        viewButtonHeight: 22,     /* 表示コマンドボタンの高さ / view command button height */
        toggleCheckboxWidth: 20   /* ラベルなしチェックボックスの幅（空文字ぶんの余白を抑える）
                                     width of the label-less checkboxes, to drop the phantom text space */
    };

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 実行環境のロケールからUI言語を判定 / Detect UI language from the host locale */
    function getCurrentLang() {
        return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var lang = getCurrentLang();

    /* UI文言の定義 / UI string definitions */
    var LABELS = {
        dialog: {
            title: { ja: "囲み罫とグリッド", en: "Frame and Grid" }
        },
        panel: {
            margin: { ja: "マージン", en: "Margin" },
            outer: { ja: "外側エリア", en: "Outer Area" },
            cap: { ja: "線端", en: "Line Caps" },
            line: { ja: "線", en: "Line" },
            titleArea: { ja: "タイトルエリア", en: "Title Area" },
            frame: { ja: "フレーム", en: "Frame" },
            innerArea: { ja: "内側エリア", en: "Inner Area" },
            offset: { ja: "オフセット", en: "Offset" },
            columns: { ja: "列", en: "Columns" },
            rows: { ja: "行", en: "Rows" },
            lineType: { ja: "線の種類", en: "Line Type" },
            display: { ja: "画面表示", en: "Display" },
            zoomPan: { ja: "ズームとパン", en: "Pan & Zoom" }
        },
        checkbox: {
            keepOuter: { ja: "外枠を残す", en: "Keep outer frame" },
            edgeScale: { ja: "辺の伸縮", en: "Edge scale" },
            outerRound: { ja: "角丸", en: "Round" },
            fill: { ja: "塗り", en: "Fill" },
            bleed: { ja: "裁ち落とし", en: "Bleed" },
            frameRound: { ja: "角丸", en: "Round" },
            link: { ja: "連動", en: "Link" },
            divider: { ja: "分割線", en: "Dividers" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        cap: {
            butt: { ja: "なし", en: "Butt" },
            round: { ja: "丸型", en: "Round" },
            project: { ja: "突出", en: "Project" }
        },
        lineType: {
            solid: { ja: "実線", en: "Solid" },
            dash: { ja: "点線", en: "Dash" },
            dots: { ja: "ドット点線", en: "Dots" }
        },
        position: {
            top: { ja: "上", en: "Top" },
            bottom: { ja: "下", en: "Bottom" },
            left: { ja: "左", en: "Left" },
            right: { ja: "右", en: "Right" }
        },
        field: {
            width: { ja: "幅", en: "Width" },
            titleSize: { ja: "幅／高さ", en: "Size" },
            columnCount: { ja: "列数", en: "Count" },
            rowCount: { ja: "行数", en: "Count" },
            spacing: { ja: "間隔", en: "Spacing" }
        },
        view: {
            zoom: { ja: "ズーム", en: "Zoom" },
            panLR: { ja: "左右", en: "Pan L/R" },
            panUD: { ja: "上下", en: "Pan U/D" },
            unavailable: { ja: "（画面表示は利用できません）", en: "(ViewControl unavailable)" }
        },
        button: {
            fitIn: { ja: "アートボード全体表示", en: "Fit Artboard" },
            actualSize: { ja: "100%表示", en: "Actual Size" },
            fitAll: { ja: "全アートボード全体表示", en: "Fit All" },
            zoomOut10: { ja: "10%縮小", en: "Zoom Out 10%" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "実行", en: "OK" }
        }
    };

    /* 入力欄の前に置く項目名の区切り記号（日本語は全角コロン）
       Separator appended to a field label placed before its input */
    var FIELD_LABEL_SEPARATOR = (lang === "ja") ? "：" : ":";

    /* ドット区切りのパスでLABELSから文言を取得 / Look up a label by dot-separated path */
    function L(labelPath) {
        var pathSegments = String(labelPath).split(".");
        var labelNode = LABELS;

        for (var i = 0; i < pathSegments.length; i++) {
            if (!labelNode || labelNode[pathSegments[i]] == null) return labelPath;
            labelNode = labelNode[pathSegments[i]];
        }

        if (labelNode[lang] != null) return labelNode[lang];
        if (labelNode.ja != null) return labelNode.ja;
        return labelPath;
    }

    /**
     * 入力欄の前に置く項目名を、区切り記号付きで取得します。
     *
     * チェックボックス・ラジオボタン・パネル名には付けません（項目名ではないため）。
     *
     * @param {string} labelPath - LABELS のドット区切りパス。
     * @returns {string} 区切り記号を付けた文言。
     */
    function fieldLabel(labelPath) {
        return L(labelPath) + FIELD_LABEL_SEPARATOR;
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コード→ラベル（rulerTypeではコード5をHとして扱う）
       Unit code to label (rulerType treats code 5 as H) */
    var UNIT_LABELS = {
        0: "in",
        1: "mm",
        2: "pt",
        3: "pica",
        4: "cm",
        5: "H",
        6: "px",
        7: "ft/in",
        8: "m",
        9: "yd",
        10: "ft"
    };

    /* 単位コード→pt換算係数 / Unit code to points factor */
    var UNIT_FACTORS = {
        0: 72.0,                 /* in */
        1: 72.0 / 25.4,          /* mm */
        2: 1.0,                  /* pt */
        3: 12.0,                 /* pica */
        4: 72.0 / 2.54,          /* cm */
        5: 72.0 / 25.4 * 0.25,   /* H（0.25mm） */
        6: 1.0,                  /* px */
        7: 72.0 * 12.0,          /* ft/in */
        8: 72.0 / 25.4 * 1000.0, /* m */
        9: 72.0 * 36.0,          /* yd */
        10: 72.0 * 12.0          /* ft */
    };

    /* 現在の定規単位コードを取得 / Get the current ruler unit code */
    function getCurrentRulerUnitCode() {
        try {
            return app.preferences.getIntegerPreference("rulerType");
        } catch (_) {
            return 2; // pt
        }
    }

    /* 現在の定規単位のラベルを取得 / Get the current ruler unit label */
    function getCurrentRulerUnitLabel() {
        return UNIT_LABELS[getCurrentRulerUnitCode()] || "pt";
    }

    /* 現在の定規単位→ptの換算係数を取得 / Get the ruler unit to points factor */
    function getCurrentRulerPtFactor() {
        var factor = UNIT_FACTORS[getCurrentRulerUnitCode()];
        return factor ? factor : 1.0;
    }

    /* mmをptに換算 / Convert millimeters to points */
    function mmToPt(mm) {
        return (72.0 / 25.4) * mm;
    }

    // =========================================
    // セッション状態 / Session state
    // =========================================
    /* Illustratorの起動中だけダイアログの値を保持する
       Dialog values are kept only while Illustrator is running */

    var ENGINE_STATE_KEY = "__SmartGridMaker__";
    $.global[ENGINE_STATE_KEY] = $.global[ENGINE_STATE_KEY] || {};

    /* 前回のダイアログ設定を読み込む / Load the previous dialog state */
    function loadSessionState() {
        return $.global[ENGINE_STATE_KEY] || {};
    }

    /* 現在のダイアログ設定を保存する / Save the current dialog state */
    function saveSessionState(state) {
        $.global[ENGINE_STATE_KEY] = state || {};
    }

    /* =========================================
     * ViewControl util (extractable)
     * Zoom + Pan (L/R, U/D) for Illustrator view
     *
     * UI (slider only):
     *   ズーム <===== slider =====>
     *   左右   <===== slider =====>
     *   上下   <===== slider =====>
     *
     * - Base center: active artboard center
     * - Pan offsets: relative to base center
     * - Zoom keeps current pan offsets
     * - Hold Option(Alt) while dragging => 1/10 speed (delta*0.1)
     * - Pan range: dynamic (half of artboard W/H), safety clamped
     *
     * How to use:
     *   var viewControl = ViewControl.create(doc, L);   // L optional
     *   viewControl.buildUI(zoomPanGroup, { labelWidth:58, sliderWidth:200 });
     *   // Cancel:
     *   viewControl.restore();
     * ========================================= */

    var ViewControl = (function () {

        /* 値を下限・上限に収める / Clamp a value into range */
        function clamp(value, minValue, maxValue) {
            if (value < minValue) return minValue;
            if (value > maxValue) return maxValue;
            return value;
        }

        function getAltKey() {
            try { return !!(ScriptUI.environment.keyboardState && ScriptUI.environment.keyboardState.altKey); } catch (_) { }
            return false;
        }

        /* Option(Alt)を押しながらのドラッグを1/10の移動量にする
           Option(Alt) held while dragging moves the slider at one tenth speed
           sliderState: {rawValue, effectiveValue} を呼び出し側が保持する */
        function applySliderAltFine(slider, sliderState, applyFn) {
            try {
                if (!slider || !sliderState || typeof applyFn !== "function") return;

                var rawValue = Number(slider.value);
                if (isNaN(rawValue)) rawValue = 0;

                if (sliderState.rawValue == null || sliderState.effectiveValue == null) {
                    sliderState.rawValue = rawValue;
                    sliderState.effectiveValue = rawValue;
                }

                if (!getAltKey()) {
                    sliderState.rawValue = rawValue;
                    sliderState.effectiveValue = rawValue;
                    applyFn(rawValue);
                    return;
                }

                var delta = rawValue - Number(sliderState.rawValue);
                var effectiveValue = Number(sliderState.effectiveValue) + delta * 0.1;

                sliderState.rawValue = rawValue;
                sliderState.effectiveValue = effectiveValue;
                try { slider.value = effectiveValue; } catch (_) { }
                applyFn(effectiveValue);
            } catch (_) { }
        }

        function getActiveArtboardCenter(doc, view) {
            try {
                var artboardIndex = doc.artboards.getActiveArtboardIndex();
                var rect = doc.artboards[artboardIndex].artboardRect; // [L, T, R, B]
                return [rect[0] + (rect[2] - rect[0]) / 2, rect[1] + (rect[3] - rect[1]) / 2];
            } catch (_) { }
            try { return (view && view.centerPoint) ? view.centerPoint : [0, 0]; } catch (_) { }
            return [0, 0];
        }

        /* パンの可動範囲＝アートボードの半分（安全側にクランプ）
           Pan range is half the artboard size, clamped for safety */
        function getPanRangePt(doc) {
            try {
                var artboardIndex = doc.artboards.getActiveArtboardIndex();
                var rect = doc.artboards[artboardIndex].artboardRect;
                var artboardWidth = Math.abs(rect[2] - rect[0]);
                var artboardHeight = Math.abs(rect[1] - rect[3]);
                var xMax = Math.round(artboardWidth / 2);
                var yMax = Math.round(artboardHeight / 2);
                if (!xMax || xMax < 100) xMax = 100;
                if (!yMax || yMax < 100) yMax = 100;
                if (xMax > 50000) xMax = 50000;
                if (yMax > 50000) yMax = 50000;
                return { xMax: xMax, yMax: yMax };
            } catch (_) { }
            return { xMax: 2000, yMax: 2000 };
        }

        /* Illustratorが受け付けるズーム倍率の範囲に収める
           Keep the zoom factor inside the range Illustrator accepts */
        function clampZoomFactor(zoomFactor) {
            return clamp(zoomFactor, 0.0313, 640.0);
        }

        function create(doc, localizeFn) {
            var viewControl = {};

            viewControl.doc = doc;
            viewControl.view = null;
            viewControl.localizeFn = (typeof localizeFn === "function") ? localizeFn : null;

            viewControl.originalZoom = null;
            viewControl.originalCenter = null;

            viewControl.panX = 0; // pt
            viewControl.panY = 0; // pt (UI positive => down)
            viewControl.panRange = { xMax: 2000, yMax: 2000 };

            viewControl.zoomSlider = null;
            viewControl.panXSlider = null;
            viewControl.panYSlider = null;

            try {
                viewControl.view = doc.views[0];
                viewControl.originalZoom = viewControl.view.zoom;
                viewControl.originalCenter = viewControl.view.centerPoint;
            } catch (_) { }

            /* ローカライズ関数が渡されていればそれを使い、無ければ fallback
               Use the supplied localizer, falling back to the given text */
            viewControl.localize = function (labelPath, fallbackText) {
                try { if (viewControl.localizeFn) return viewControl.localizeFn(labelPath); } catch (_) { }
                return fallbackText || labelPath;
            };

            viewControl.refreshPanRange = function () {
                try { viewControl.panRange = getPanRangePt(viewControl.doc); } catch (_) { }
                return viewControl.panRange;
            };

            /* アートボード中心＋パン量を表示中心にする / Centre on the artboard plus the pan offsets */
            viewControl.applyCenter = function () {
                try {
                    if (!viewControl.view) return;

                    var center = getActiveArtboardCenter(viewControl.doc, viewControl.view);
                    var centerX = center[0] + Number(viewControl.panX || 0);
                    // Illustrator: +Y is up. UI: positive is down => subtract
                    var centerY = center[1] - Number(viewControl.panY || 0);

                    viewControl.view.centerPoint = [centerX, centerY];
                    app.redraw();
                } catch (_) { }
            };

            viewControl.setZoomPercent = function (percent, zoomMin, zoomMax) {
                try {
                    if (!viewControl.view) return;

                    var clampedPercent = Number(percent);
                    if (isNaN(clampedPercent)) return;

                    clampedPercent = clamp(Math.round(clampedPercent), zoomMin, zoomMax);
                    viewControl.view.zoom = clampZoomFactor(clampedPercent / 100.0);
                    viewControl.applyCenter(); // keep pan
                    syncZoomSlider();
                } catch (_) { }
            };

            /* ズームスライダーを現在の表示倍率に追従させる
               Keep the zoom slider in step with the current view */
            function syncZoomSlider() {
                try {
                    if (!viewControl.zoomSlider || !viewControl.view) return;
                    viewControl.zoomSlider.value = Math.round(Number(viewControl.view.zoom) * 100);
                } catch (_) { }
            }

            /**
             * 現在の表示倍率を指定の倍率で変更します（0.9 なら10%縮小）。
             *
             * @param {number} factor - 掛ける倍率。
             * @returns {void}
             */
            viewControl.zoomBy = function (factor) {
                try {
                    if (!viewControl.view) return;

                    var currentZoom = Number(viewControl.view.zoom);
                    var scale = Number(factor);
                    if (isNaN(currentZoom) || !(currentZoom > 0) || isNaN(scale) || !(scale > 0)) return;

                    viewControl.view.zoom = clampZoomFactor(currentZoom * scale);
                    viewControl.applyCenter(); // keep pan
                    syncZoomSlider();
                } catch (_) { }
            };

            /* setPanX / setPanY の共通処理：可動範囲に収めて表示中心を更新
               Shared by setPanX / setPanY: clamp to the current range and re-centre */
            function setPan(panProperty, value, rangeProperty) {
                try {
                    var panPt = Number(value);
                    if (isNaN(panPt)) panPt = 0;

                    var limit = viewControl.refreshPanRange()[rangeProperty];
                    viewControl[panProperty] = clamp(Math.round(panPt), -limit, limit);
                    viewControl.applyCenter();
                } catch (_) { }
            }

            viewControl.setPanX = function (value) { setPan("panX", value, "xMax"); };
            viewControl.setPanY = function (value) { setPan("panY", value, "yMax"); };

            viewControl.restore = function () {
                try {
                    if (viewControl.view && viewControl.originalZoom != null && viewControl.originalCenter != null) {
                        viewControl.view.zoom = viewControl.originalZoom;
                        viewControl.view.centerPoint = viewControl.originalCenter;
                        app.redraw();
                    }
                } catch (_) { }
                viewControl.panX = 0; viewControl.panY = 0;
            };

            viewControl.buildUI = function (parent, options) {
                options = options || {};
                var labelWidth = (typeof options.labelWidth === "number") ? options.labelWidth : 58;
                var sliderWidth = (typeof options.sliderWidth === "number") ? options.sliderWidth : 200;

                var zoomMin = (typeof options.zoomMin === "number") ? options.zoomMin : 10;
                var zoomMax = (typeof options.zoomMax === "number") ? options.zoomMax : 1600;

                /* 項目名の区切り記号（ViewControl 単体で使えるようロケールは内部で判定）
                   Field label separator, resolved locally so ViewControl stays self-contained */
                var separator = ($.locale && $.locale.indexOf("ja") === 0) ? "：" : ":";

                // labels (prefer L())
                var labelZoom = viewControl.localize("view.zoom", "ズーム") + separator;
                var labelPanX = viewControl.localize("view.panLR", "左右") + separator;
                var labelPanY = viewControl.localize("view.panUD", "上下") + separator;

                // initial zoom percent from the current view
                var initialZoomPercent = 100;
                try {
                    if (viewControl.originalZoom != null) initialZoomPercent = Math.round(Number(viewControl.originalZoom) * 100);
                } catch (_) { }
                if (!initialZoomPercent || initialZoomPercent < zoomMin) initialZoomPercent = 100;

                // pan ranges (UI bounds decided at build time)
                var panRange = viewControl.refreshPanRange();

                /* ラベル＋スライダーの1行を追加（applyFn は補正後の値を受け取る）
                   Label + slider on one row; applyFn receives the alt-fine adjusted value */
                function addSliderRow(labelText, value, minValue, maxValue, applyFn) {
                    var row = parent.add("group");
                    row.orientation = "row";
                    row.alignChildren = ["left", "center"];

                    var label = row.add("statictext", undefined, labelText);
                    var slider = row.add("slider", undefined, value, minValue, maxValue);
                    try {
                        label.preferredSize.width = labelWidth;
                        slider.preferredSize.width = sliderWidth;
                    } catch (_) { }

                    var sliderState = { rawValue: null, effectiveValue: null };
                    slider.onChanging = function () {
                        applySliderAltFine(this, sliderState, applyFn);
                    };
                    return slider;
                }

                viewControl.zoomSlider = addSliderRow(labelZoom, initialZoomPercent, zoomMin, zoomMax, function (value) {
                    viewControl.setZoomPercent(value, zoomMin, zoomMax);
                });
                viewControl.panXSlider = addSliderRow(labelPanX, 0, -panRange.xMax, panRange.xMax, function (value) {
                    viewControl.setPanX(value);
                });
                viewControl.panYSlider = addSliderRow(labelPanY, 0, -panRange.yMax, panRange.yMax, function (value) {
                    viewControl.setPanY(value);
                });

                return viewControl;
            };

            return viewControl;
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
        try {
            item.name = tag;
            item.note = tag;
        } catch (_) { }
    }

    /**
     * オブジェクトが指定のタグを持つかどうかを判定します。
     *
     * @param {PageItem} item - 判定するオブジェクト。
     * @param {string} tag - 探すタグ文字列。
     * @returns {boolean} タグを持つ場合は true。
     */
    function hasTag(item, tag) {
        /* 削除済みの参照に触ると例外になるためまとめて保護
           A removed reference throws, so guard the whole lookup */
        try {
            return (item.note === tag || item.name === tag);
        } catch (_) { }
        return false;
    }

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
     * 線と塗りの見た目をコピーします（角丸プレビュー用の複製に使用）。
     *
     * @param {PathItem} source - コピー元のパス。
     * @param {PathItem} target - コピー先のパス。
     * @returns {void}
     */
    function copyAppearance(source, target) {
        try {
            target.stroked = source.stroked;
            target.filled = source.filled;
            target.strokeColor = source.strokeColor;
            target.fillColor = source.fillColor;
            target.strokeWidth = source.strokeWidth;
        } catch (_) { }
    }

    /**
     * オブジェクトを最背面へ送ります（環境差で落ちることがあるため保護）。
     *
     * @param {PageItem} item - 対象のオブジェクト。
     * @returns {void}
     */
    function sendToBack(item) {
        try { item.zOrder(ZOrderMethod.SENDTOBACK); } catch (_) { }
    }

    /**
     * オブジェクトを削除します（すでに消えている場合は何もしません）。
     *
     * @param {PageItem} item - 対象のオブジェクト。
     * @returns {void}
     */
    function removeItem(item) {
        try { item.remove(); } catch (_) { }
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
     * 内部タグをレイヤーパネルから隠します（name をクリアし、判定用の note は残します）。
     *
     * @param {PageItem[]} items - 対象のオブジェクト。
     * @returns {void}
     */
    function clearTagNames(items) {
        for (var i = 0; i < items.length; i++) {
            try { items[i].name = ""; } catch (_) { }
        }
    }

    (function () {
        // =========================================
        // 準備とチェック / Setup
        // =========================================
        if (app.documents.length === 0) return;
        var doc = app.activeDocument;
        var selection = doc.selection;

        /* ズーム・パン用のViewControl / ViewControl instance for zoom and pan */
        var viewCtl = null;
        try { viewCtl = ViewControl.create(doc, L); } catch (_) { viewCtl = null; }

        /* 選択からパスアイテムだけを抽出 / Collect the path items from the selection */
        var targetItems = [];
        for (var i = 0; i < selection.length; i++) {
            if (selection[i].typename === "PathItem") {
                targetItems.push(selection[i]);
            }
        }

        /* 基準にする長方形の見た目を統一（塗りなし／黒1pt）
           Normalize the base rectangles to no fill and a 1pt black stroke */
        for (var i = 0; i < targetItems.length; i++) {
            try {
                targetItems[i].filled = false;
                targetItems[i].stroked = true;
                targetItems[i].strokeWidth = 1;
                targetItems[i].strokeColor = makeGrayColor(FILL_TINTS.line);
            } catch (_) { }
        }

        /* 選択がない場合は、現在のアートボードを基準にする
           With nothing selected, the active artboard becomes the base */
        var _usingArtboardBase = false;
        var _artboardBaseRect = null;

        /* 裁ち落とし（アートボード基準のみ） / Bleed, artboard-based runs only */
        var _bleedEnabled = false;

        /* アートボード基準の一時矩形を破棄する（targetItems からも外す）
           remove後に触ると Error 45 になるため、参照を必ず先に外す
           Discard the temporary artboard rectangle; drop the reference first to avoid Error 45 */
        function cleanupArtboardBaseRect() {
            if (!_artboardBaseRect) return;

            for (var i = targetItems.length - 1; i >= 0; i--) {
                if (targetItems[i] === _artboardBaseRect) targetItems.splice(i, 1);
            }

            removeItem(_artboardBaseRect);
            _artboardBaseRect = null;
        }

        // アートボード基準の一時矩形を「マージン」分だけ内側に作り直す
        function rebuildArtboardBaseRect(marginTopPt, marginRightPt, marginBottomPt, marginLeftPt) {
            if (!_usingArtboardBase) return;
            try {
                var artboardIndex = doc.artboards.getActiveArtboardIndex();
                var artboardRect = doc.artboards[artboardIndex].artboardRect; // [L, T, R, B]
                var left = artboardRect[0], top = artboardRect[1], right = artboardRect[2], bottom = artboardRect[3];

                // ※裁ち落とし（bleed）はここでは適用しない（フレームのみで適用）
                var marginTop = (marginTopPt > 0) ? marginTopPt : 0;
                var marginRight = (marginRightPt > 0) ? marginRightPt : 0;
                var marginBottom = (marginBottomPt > 0) ? marginBottomPt : 0;
                var marginLeft = (marginLeftPt > 0) ? marginLeftPt : 0;

                var insetLeft = left + marginLeft;
                var insetTop = top - marginTop;
                var insetRight = right - marginRight;
                var insetBottom = bottom + marginBottom;

                var width = insetRight - insetLeft;
                var height = insetTop - insetBottom;
                if (!(width > 0) || !(height > 0)) {
                    // マージンが大きすぎる場合は 0 扱い
                    insetLeft = left; insetTop = top; insetRight = right; insetBottom = bottom;
                    width = insetRight - insetLeft;
                    height = insetTop - insetBottom;
                }

                // 既存の一時矩形があれば安全に差し替える（targetItemsからも外す）
                cleanupArtboardBaseRect();

                // アートボード基準の外枠罫線設定：1pt / 黒（ガイドにはしない）
                _artboardBaseRect = doc.activeLayer.pathItems.rectangle(insetTop, insetLeft, width, height);
                _artboardBaseRect.stroked = true;
                _artboardBaseRect.filled = false;
                _artboardBaseRect.strokeColor = makeGrayColor(FILL_TINTS.line);
                _artboardBaseRect.strokeWidth = 1;

                targetItems.push(_artboardBaseRect);
            } catch (_) { }
        }

        /* パスアイテムが選択されていなければアートボード基準に切り替える
           Fall back to the artboard when no path item is selected
           ※実体の基準矩形はダイアログ生成直前に rebuildArtboardBaseRect() で作成する
             The actual base rectangle is created just before the dialog is built */
        if (targetItems.length === 0) {
            _usingArtboardBase = true;
        }

        /* 一部環境で StrokeCap が未定義になるため、最低限の定数を用意
           Provide the StrokeCap constants for hosts that do not expose them */
        if (typeof StrokeCap === "undefined") {
            StrokeCap = {
                BUTTENDCAP: 0,
                ROUNDENDCAP: 1,
                PROJECTINGENDCAP: 2
            };
        }

        /**
         * 内側エリアのオフセット初期値を求めます（外形の (幅+高さ)/40 を10単位に丸めた値）。
         *
         * @returns {number} 現在の定規単位でのオフセット初期値。
         */
        function calcDefaultInnerOffset() {
            if (targetItems.length === 0) return 0;

            var bounds = targetItems[0].geometricBounds; // [L, T, R, B]
            var widthPt = bounds[2] - bounds[0];
            var heightPt = bounds[1] - bounds[3];
            if (!(widthPt > 0) || !(heightPt > 0)) return 0;

            var offsetPt = (widthPt + heightPt) / GRID_CONFIG.innerOffsetDivisor;
            var offsetValue = offsetPt / getCurrentRulerPtFactor();
            if (offsetValue < 0) return 0;

            /* 10単位に丸める（例：150.3→150, 155→160） / Round to the nearest ten */
            return Math.round(offsetValue / 10) * 10;
        }

        /* プレビューで生成した一時オブジェクト / Temporary items created for the preview */
        var tempPreviewItems = [];

        /**
         * 生成した一時オブジェクトを追跡リストに登録して返します。
         *
         * 生成した直後に登録するのが重要です。見た目の設定などで例外が起きても、
         * すでに登録済みならプレビュー解除時に必ず削除できます
         * （登録前に中断すると、消せないオブジェクトがドキュメントに残ります）。
         *
         * @param {PageItem} item - 登録するオブジェクト。
         * @returns {PageItem} 受け取ったオブジェクトをそのまま返します。
         */
        function trackTempItem(item) {
            tempPreviewItems.push(item);
            return item;
        }

        /**
         * 長方形の指定した2角だけを角丸にします（同じパスを書き換えます）。
         *
         * @param {PathItem} rect - 対象の長方形（閉じたパス）。
         * @param {string} cornerPair - "TOP" / "BOTTOM" / "LEFT" / "RIGHT"。
         * @param {number} radiusPt - 角丸の半径（pt）。
         * @returns {boolean} 角丸にした場合は true。
         */
        function roundRectCornerPairInPlace(rect, cornerPair, radiusPt) {
            if (!rect || rect.typename !== "PathItem" || !rect.closed) return false;

            var requestedRadius = Number(radiusPt);
            if (isNaN(requestedRadius) || requestedRadius <= 0) return false;

            var bounds = rect.geometricBounds; // [L,T,R,B]
            var left = bounds[0], top = bounds[1], right = bounds[2], bottom = bounds[3];
            var width = right - left, height = top - bottom;
            if (!(width > 0) || !(height > 0)) return false;

            /* 半径は辺の半分までに抑える / Cap the radius at half the shorter side */
            var radius = Math.min(requestedRadius, width / 2, height / 2);
            if (!(radius > 0)) return false;

            /* 四分円をベジェで近似するハンドル長 / Bezier handle length for a quarter circle */
            var handleLength = radius * 0.5522847498307936;
            var POINT_TYPE = (typeof PointType !== "undefined") ? PointType : { CORNER: 0, SMOOTH: 1 };

            var roundTopLeft = (cornerPair === "TOP" || cornerPair === "LEFT");
            var roundTopRight = (cornerPair === "TOP" || cornerPair === "RIGHT");
            var roundBottomRight = (cornerPair === "BOTTOM" || cornerPair === "RIGHT");
            var roundBottomLeft = (cornerPair === "BOTTOM" || cornerPair === "LEFT");

            var cornerPoints = [];

            /**
             * アンカーと左右のハンドルを1点ぶん記録します（ハンドル省略時はアンカーと同じ位置）。
             *
             * @param {number} anchorX - アンカーのX座標。
             * @param {number} anchorY - アンカーのY座標。
             * @param {number} leftX - 左方向ハンドルのX座標（null でアンカーと同じ）。
             * @param {number} leftY - 左方向ハンドルのY座標。
             * @param {number} rightX - 右方向ハンドルのX座標（null でアンカーと同じ）。
             * @param {number} rightY - 右方向ハンドルのY座標。
             * @param {number} pointType - PointType.CORNER または PointType.SMOOTH。
             * @returns {void}
             */
            function addPoint(anchorX, anchorY, leftX, leftY, rightX, rightY, pointType) {
                cornerPoints.push({
                    anchor: [anchorX, anchorY],
                    leftDirection: (leftX == null) ? [anchorX, anchorY] : [leftX, leftY],
                    rightDirection: (rightX == null) ? [anchorX, anchorY] : [rightX, rightY],
                    pointType: pointType
                });
            }

            // 始点：左上（上辺側）/ Start at the top-left corner, on the top edge
            if (roundTopLeft) {
                addPoint(left + radius, top, left + radius - handleLength, top, null, null, POINT_TYPE.SMOOTH);
            } else {
                addPoint(left, top, null, null, null, null, POINT_TYPE.CORNER);
            }

            // 右上 / Top-right
            if (roundTopRight) {
                addPoint(right - radius, top, null, null, right - radius + handleLength, top, POINT_TYPE.SMOOTH);
                addPoint(right, top - radius, right, top - radius + handleLength, null, null, POINT_TYPE.SMOOTH);
            } else {
                addPoint(right, top, null, null, null, null, POINT_TYPE.CORNER);
            }

            // 右下 / Bottom-right
            if (roundBottomRight) {
                addPoint(right, bottom + radius, null, null, right, bottom + radius - handleLength, POINT_TYPE.SMOOTH);
                addPoint(right - radius, bottom, right - radius + handleLength, bottom, null, null, POINT_TYPE.SMOOTH);
            } else {
                addPoint(right, bottom, null, null, null, null, POINT_TYPE.CORNER);
            }

            // 左下 / Bottom-left
            if (roundBottomLeft) {
                addPoint(left + radius, bottom, null, null, left + radius - handleLength, bottom, POINT_TYPE.SMOOTH);
                addPoint(left, bottom + radius, left, bottom + radius - handleLength, null, null, POINT_TYPE.SMOOTH);
            } else {
                addPoint(left, bottom, null, null, null, null, POINT_TYPE.CORNER);
            }

            // 終点：左上（左辺側）/ Close at the top-left corner, on the left edge
            if (roundTopLeft) {
                addPoint(left, top - radius, null, null, left, top - radius + handleLength, POINT_TYPE.SMOOTH);
            }

            var anchors = [];
            for (var i = 0; i < cornerPoints.length; i++) {
                anchors.push(cornerPoints[i].anchor);
            }
            rect.setEntirePath(anchors);
            rect.closed = true;

            var pathPoints = rect.pathPoints;
            var pointCount = Math.min(pathPoints.length, cornerPoints.length);
            for (var j = 0; j < pointCount; j++) {
                pathPoints[j].leftDirection = cornerPoints[j].leftDirection;
                pathPoints[j].rightDirection = cornerPoints[j].rightDirection;
                try { pathPoints[j].pointType = cornerPoints[j].pointType; } catch (_) { }
            }
            return true;
        }

        /* タイトルエリアの生成物か（タグ定数で判定） / Is this item part of the title area? */
        function isTitleBandItem(item) {
            if (!item) return false;
            return hasTag(item, TAG_TITLE_FILL) || hasTag(item, TAG_TITLE_DIVIDER);
        }

        /* タイトルの位置キー→角丸にする2角 / Title position to the pair of rounded corners */
        var TITLE_CORNER_PAIRS = {
            top: "TOP",
            bottom: "BOTTOM",
            left: "LEFT",
            right: "RIGHT"
        };

        /**
         * タイトルエリアの角丸を、位置に応じた2角だけに適用します。
         *
         * @param {PathItem} titleRect - タイトル帯の矩形。
         * @param {string} titlePosKey - "top" / "bottom" / "left" / "right"。
         * @param {number} radiusPt - 角丸の半径（pt）。
         * @returns {boolean} 適用した場合は true。
         */
        function applyTitleAreaCornerRounding(titleRect, titlePosKey, radiusPt) {
            if (!titleRect || titleRect.typename !== "PathItem") return false;
            if (!isTitleBandItem(titleRect)) return false;
            if (!(radiusPt > 0)) return false;

            var cornerPair = TITLE_CORNER_PAIRS[titlePosKey];
            if (!cornerPair) return false;

            return roundRectCornerPairInPlace(titleRect, cornerPair, radiusPt);
        }

        // アートボード基準の場合はここで基準矩形を作成（プレビュー生成物と同じ扱いにする）
        if (_usingArtboardBase) {
            rebuildArtboardBaseRect(0, 0, 0, 0);
        }

        // --- ダイアログ作成 ---
        var dialog = new Window("dialog", L("dialog.title") + " " + SCRIPT_VERSION);
        // --- ダイアログ位置・透明度設定（値は UI_CONFIG） ---
        function shiftDialogPosition(dialog, offsetX, offsetY) {
            dialog.onShow = function () {
                var currentX = dialog.location[0];
                var currentY = dialog.location[1];
                dialog.location = [currentX + offsetX, currentY + offsetY];
            };
        }

        function setDialogOpacity(dialog, opacityValue) {
            try {
                dialog.opacity = opacityValue;
            } catch (_) { }
        }

        setDialogOpacity(dialog, UI_CONFIG.dialogOpacity);
        shiftDialogPosition(dialog, UI_CONFIG.dialogOffsetX, UI_CONFIG.dialogOffsetY);
        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];
        dialog.spacing = 20;
        dialog.margins = 16;

        // 3タブレイアウト
        // 左：マージン / フレーム
        // 中央：外側エリア / タイトルエリア
        // 右：内側エリア
        var tabPanel = dialog.add("tabbedpanel");
        tabPanel.alignChildren = ["fill", "top"];
        tabPanel.alignment = ["fill", "top"];
        tabPanel.margins = [5, 20, 0, 0];
        // tabbedpanel は内容量に応じて自動で高さが伸びないことがあるため、最低サイズを与える
        try {
            tabPanel.minimumSize = UI_CONFIG.tabSize;
            tabPanel.preferredSize = UI_CONFIG.tabSize;
        } catch (_) { }

        // Tabs
        var tabMargin = tabPanel.add("tab", undefined, L("panel.margin"));
        var tabOuter = tabPanel.add("tab", undefined, L("panel.outer"));
        var tabInner = tabPanel.add("tab", undefined, L("panel.innerArea"));
        var tabDisplay = tabPanel.add("tab", undefined, L("panel.display"));

        // Left tab (Margin / Frame)
        tabMargin.orientation = "column";
        tabMargin.alignChildren = ["fill", "top"];
        tabMargin.spacing = 8;
        tabMargin.margins = 10;
        var marginTabColumn = tabMargin.add("group");
        marginTabColumn.orientation = "column";
        marginTabColumn.alignChildren = ["fill", "top"];
        marginTabColumn.spacing = 8;

        // Middle tab (Outer / Title)
        tabOuter.orientation = "column";
        tabOuter.alignChildren = ["fill", "top"];
        tabOuter.spacing = 8;
        tabOuter.margins = 10;
        var outerTabColumn = tabOuter.add("group");
        outerTabColumn.orientation = "column";
        outerTabColumn.alignChildren = ["fill", "top"];
        outerTabColumn.spacing = 8;

        // Right tab (Inner)
        tabInner.orientation = "column";
        tabInner.alignChildren = ["fill", "top"];
        tabInner.spacing = 8;
        tabInner.margins = 10;
        var innerTabColumn = tabInner.add("group");
        innerTabColumn.orientation = "column";
        innerTabColumn.alignChildren = ["fill", "top"];
        innerTabColumn.spacing = 8;

        // Display tab
        tabDisplay.orientation = "column";
        tabDisplay.alignChildren = ["fill", "top"];
        tabDisplay.spacing = 8;
        tabDisplay.margins = 10;

        var displayTabColumn = tabDisplay.add("group");
        displayTabColumn.orientation = "column";
        displayTabColumn.alignChildren = ["fill", "top"];
        displayTabColumn.spacing = 8;

        // =========================================
        // UIの組み立て / UI builders
        // =========================================

        /**
         * パネルを非表示にし、レイアウト上の高さも潰します（長方形スタート時に使用）。
         *
         * @param {Panel} panel - 対象のパネル。
         * @param {boolean} collapsed - 畳む場合は true。
         * @returns {void}
         */
        function setPanelCollapsed(panel, collapsed) {
            try {
                panel.visible = !collapsed;
                panel.minimumSize.height = 0;
                panel.maximumSize.height = collapsed ? 0 : 10000;
            } catch (_) { }
        }

        // --- Margin UI handles ---
        var marginPanel;
        var editArtboardMarginTop, editArtboardMarginBottom, editArtboardMarginLeft, editArtboardMarginRight;
        var chkArtboardMarginLink;
        var _syncingArtboardMargins = false;
        var applyArtboardMarginLinkState;

        // --- Outer UI handles ---
        var outerPanel;
        var chkKeepOuter;
        var chkOuterEdgeScale;
        var outerEdgeScaleGroup;
        var editOuterEdgeScale;
        var chkOuterRound;
        var editOuterRound;
        var applyOuterEdgeScaleEnabledState;
        var applyStrokeCapPanelEnabledState;
        var applyOuterRoundEnabledState;
        var strokeCapPanel;
        var rbCapButt, rbCapRound, rbCapProject;

        // --- Title UI handles ---
        var titlePanel, titleSizeGroup, titlePositionGroup, titleOptionGroup, titleEdgeScaleRow;
        var chkTitleEnable, editTitleSize;
        var rbTitleTop, rbTitleBottom, rbTitleLeft, rbTitleRight;
        var chkTitleFill, chkTitleLine;
        var chkTitleEdgeScale, editTitleEdgeScale;
        var applyTitleAreaEnabledState, applyTitleEdgeScaleEnabledState;
        var getTitlePosKey, maybeApplyTitleAreaRound;

        // --- Frame UI handles ---
        var framePanel, frameWidthGroup;
        var chkFrameEnable, editFrameWidth, chkBleed, chkFrameRound, editFrameRound;
        var applyFrameEnabledState;

        // --- Inner UI handles ---
        var innerPanel;
        var editInnerOffsetTop, editInnerOffsetBottom, editInnerOffsetLeft, editInnerOffsetRight;
        var chkInnerOffsetLink;
        var _syncingInnerOffsets = false;
        var applyInnerOffsetLinkState;

        var editInnerColumns, editInnerRows;
        var editColGutter, editRowGutter;
        var chkInnerFill, chkInnerDivider;
        var innerLineTypePanel;
        var rbInnerLineSolid, rbInnerLineDash, rbInnerLineDotDash;

        /**
         * 上／左＋連動＋右／下 の3段組の入力UIを作ります（マージンと内側オフセットで共用）。
         *
         * @param {Panel|Group} parent - 追加先。
         * @param {Object} options - value: 初期値／characters: 入力欄の文字数／unitLabel: 上下に付ける単位（省略可）。
         * @returns {Object} {top, bottom, left, right, link, applyLinkState}
         */
        function buildLinkedQuadUI(parent, options) {
            var fields = {};

            /* 中央寄せの1段を追加 / Add one centred row */
            function addRow(spacing) {
                var row = parent.add("group");
                row.orientation = "row";
                row.alignChildren = ["center", "center"];
                row.alignment = ["fill", "top"];
                if (spacing) row.spacing = spacing;
                return row;
            }

            /* ラベル＋入力欄（＋単位）のひとまとまりを追加 / Add one labelled input */
            function addField(row, posKey, withUnit) {
                var group = row.add("group");
                group.orientation = "row";
                group.alignChildren = ["left", "center"];
                group.add("statictext", undefined, fieldLabel("position." + posKey));

                var field = group.add("edittext", undefined, options.value);
                field.characters = options.characters;
                changeValueByArrowKey(field, false);

                if (withUnit && options.unitLabel) group.add("statictext", undefined, options.unitLabel);

                fields[posKey] = field;
                return group;
            }

            // 1段目：上（中央寄せ）
            addField(addRow(), "top", true);

            // 2段目：左 ＋ 連動（中央）＋ 右
            var midRow = addRow(12);
            var leftGroup = addField(midRow, "left", false);
            var link = midRow.add("checkbox", undefined, L("checkbox.link"));
            link.value = true;
            var rightGroup = addField(midRow, "right", false);

            // 3段目：下（中央寄せ）
            var bottomGroup = addField(addRow(), "bottom", true);

            return {
                top: fields.top,
                bottom: fields.bottom,
                left: fields.left,
                right: fields.right,
                link: link,

                /* ［連動］の状態を反映（連動ONなら上の値を他へコピー）
                   Reflect the link state; while linked the top value drives the others */
                applyLinkState: function () {
                    var linked = !!link.value;

                    /* 連動ONのときは下・左・右をディム表示 / Dim the other three while linked */
                    bottomGroup.enabled = !linked;
                    leftGroup.enabled = !linked;
                    rightGroup.enabled = !linked;
                    if (!linked) return;

                    syncLinkedFields(fields.top, [fields.bottom, fields.left, fields.right]);
                }
            };
        }

        /**
         * 「◯数」＋「間隔」の1行パネルを作ります（列／行で共用）。
         *
         * @param {Group} parent - 追加先。
         * @param {string} titlePath - パネル名のラベルパス。
         * @param {string} countPath - 個数ラベルのラベルパス。
         * @returns {Object} {count, gutter}
         */
        function addCountSpacingPanel(parent, titlePath, countPath) {
            var panel = parent.add("panel", undefined, L(titlePath));
            panel.orientation = "column";
            panel.alignChildren = ["fill", "top"];
            panel.margins = UI_CONFIG.panelMargins;
            panel.spacing = 8;

            var row = panel.add("group");
            row.orientation = "row";
            row.alignChildren = ["left", "center"];

            row.add("statictext", undefined, fieldLabel(countPath));
            var count = row.add("edittext", undefined, "1");
            count.characters = 3;
            changeValueByArrowKey(count, false);

            row.add("statictext", undefined, fieldLabel("field.spacing"));
            var gutter = row.add("edittext", undefined, "0");
            gutter.characters = 4;
            changeValueByArrowKey(gutter, false);
            row.add("statictext", undefined, getCurrentRulerUnitLabel());

            return { count: count, gutter: gutter };
        }

        // Builder: MarginUI
        function buildMarginUI(parent) {
            // マージン（アートボード基準のときだけ有効）
            marginPanel = parent.add("panel", undefined, L("panel.margin"));
            marginPanel.orientation = "column";
            marginPanel.alignChildren = ["fill", "top"];
            marginPanel.margins = UI_CONFIG.panelMargins;
            marginPanel.spacing = 10;

            // マージン入力：3段組（上 / 左＋連動＋右 / 下）
            var marginUI = buildLinkedQuadUI(marginPanel, {
                value: GRID_CONFIG.defaultMargin,
                characters: 4,
                unitLabel: getCurrentRulerUnitLabel()
            });

            editArtboardMarginTop = marginUI.top;
            editArtboardMarginBottom = marginUI.bottom;
            editArtboardMarginLeft = marginUI.left;
            editArtboardMarginRight = marginUI.right;
            chkArtboardMarginLink = marginUI.link;

            /* ［連動］の反映（同期中は onChanging を無視させる）
               Apply the link state; the flag makes onChanging ignore our own writes */
            applyArtboardMarginLinkState = function () {
                _syncingArtboardMargins = true;
                marginUI.applyLinkState();
                _syncingArtboardMargins = false;
            };

            /* アートボード基準のときだけ表示・操作できる
               The panel is shown and enabled for artboard-based runs only */
            marginPanel.enabled = _usingArtboardBase;
            setPanelCollapsed(marginPanel, !_usingArtboardBase);

            applyArtboardMarginLinkState();
        }

        // Builder: OuterUI
        function buildOuterUI(parent) {
            // 外枠パネル
            outerPanel = parent.add("panel", undefined, L("panel.outer"));
            outerPanel.orientation = "column";
            outerPanel.alignChildren = ["fill", "top"];
            outerPanel.margins = UI_CONFIG.panelMargins;
            outerPanel.spacing = 10;

            // 外枠を残す
            chkKeepOuter = outerPanel.add("checkbox", undefined, L("checkbox.keepOuter"));
            chkKeepOuter.value = true;

            // 外側エリア：角丸
            var outerRoundRow = outerPanel.add("group");
            outerRoundRow.orientation = "row";
            outerRoundRow.alignChildren = ["left", "center"];

            chkOuterRound = outerRoundRow.add("checkbox", undefined, L("checkbox.outerRound"));
            chkOuterRound.value = false;

            editOuterRound = outerRoundRow.add("edittext", undefined, "0");
            editOuterRound.characters = 4;
            changeValueByArrowKey(editOuterRound, false);

            outerRoundRow.add("statictext", undefined, getCurrentRulerUnitLabel());

            // 辺の長さ調整（1行）
            var outerEdgeScaleRow = outerPanel.add("group");
            outerEdgeScaleRow.orientation = "row";
            outerEdgeScaleRow.alignChildren = ["left", "center"];

            chkOuterEdgeScale = outerEdgeScaleRow.add("checkbox", undefined, L("checkbox.edgeScale"));
            chkOuterEdgeScale.value = true;

            // 値入力（チェックがOFFのときだけディム表示）
            outerEdgeScaleGroup = outerEdgeScaleRow.add("group");
            outerEdgeScaleGroup.orientation = "row";
            outerEdgeScaleGroup.alignChildren = ["left", "center"];

            editOuterEdgeScale = outerEdgeScaleGroup.add("edittext", undefined, GRID_CONFIG.defaultEdgeScale);
            editOuterEdgeScale.characters = 4;
            outerEdgeScaleGroup.add("statictext", undefined, getCurrentRulerUnitLabel());
            editOuterEdgeScale.active = true;
            changeValueByArrowKey(editOuterEdgeScale, true);

            /**
             * ［辺の伸縮］の有効／無効を反映します。
             *
             * @returns {void}
             */
            applyOuterEdgeScaleEnabledState = function () {
                chkOuterEdgeScale.enabled = chkKeepOuter.value;
                outerEdgeScaleGroup.enabled = (chkKeepOuter.value && chkOuterEdgeScale.value);
            };
            applyOuterEdgeScaleEnabledState();

            /**
             * ［角丸］の入力欄の有効／無効を反映します。
             *
             * タイトルエリアの角丸が参照する値のため、チェック自体は常に操作できます。
             *
             * @returns {void}
             */
            applyOuterRoundEnabledState = function () {
                editOuterRound.enabled = chkOuterRound.value;
                if (!chkOuterRound.value) editOuterRound.text = "0";
            };
            applyOuterRoundEnabledState();

            chkOuterRound.onClick = function () {
                if (chkOuterRound.value) {
                    /* ONにしたとき、0なら既定値を入れる / Fill in the default radius when enabled at zero */
                    var radius = parseFloat(editOuterRound.text);
                    if (isNaN(radius) || radius === 0) editOuterRound.text = GRID_CONFIG.defaultRound;

                    /* 角丸と辺の伸縮は同時に使えないので、ONにした側を残して他方をOFFにする
                       The radius and the edge scale cannot be combined, so turning one on turns the other off */
                    chkOuterEdgeScale.value = false;
                }

                applyOuterEdgeScaleEnabledState();
                applyOuterRoundEnabledState();
                applyStrokeCapPanelEnabledState();
                requestPreview();
            };

            editOuterRound.onChanging = requestPreview;

            buildOuterCapUI(outerPanel);
        }

        /* 外側エリア：線端（ストロークキャップ） / Outer area line caps */
        function buildOuterCapUI(parent) {
            strokeCapPanel = parent.add("panel", undefined, L("panel.cap"));
            strokeCapPanel.orientation = "row";
            strokeCapPanel.alignChildren = ["left", "center"];
            strokeCapPanel.margins = UI_CONFIG.panelMargins;

            rbCapButt = strokeCapPanel.add("radiobutton", undefined, L("cap.butt"));
            rbCapRound = strokeCapPanel.add("radiobutton", undefined, L("cap.round"));
            rbCapProject = strokeCapPanel.add("radiobutton", undefined, L("cap.project"));

            /* 初期値は選択オブジェクトの線端を優先 / Prefer the selected object's cap */
            var currentCap = null;
            try {
                if (targetItems.length > 0 && targetItems[0].stroked) currentCap = targetItems[0].strokeCap;
            } catch (_) { }

            if (currentCap === StrokeCap.ROUNDENDCAP) rbCapRound.value = true;
            else if (currentCap === StrokeCap.PROJECTINGENDCAP) rbCapProject.value = true;
            else rbCapButt.value = true;

            /* 線端は「4辺に分解するとき」だけ意味を持つ（＝外枠を残す＋辺の伸縮≠0）
               The caps only matter while the edges are split into four lines */
            applyStrokeCapPanelEnabledState = function () {
                strokeCapPanel.enabled = (!!chkKeepOuter.value && !!chkOuterEdgeScale.value && getEffectiveOuterEdgeScale() !== 0);
            };
            applyStrokeCapPanelEnabledState();
        }

        // Builder: InnerUI
        function buildInnerUI(parent) {
            // inner box パネル
            innerPanel = parent.add("panel", undefined, L("panel.innerArea"));
            innerPanel.orientation = "column";
            innerPanel.alignChildren = ["fill", "top"];
            innerPanel.margins = UI_CONFIG.panelMargins;

            buildInnerOffsetUI(innerPanel);
            buildInnerGridUI(innerPanel);
            buildInnerLineTypeUI(innerPanel);
        }

        /* 内側エリア：オフセット入力（3段組） / Inner area offsets, three rows */
        function buildInnerOffsetUI(parent) {
            var offsetPanel = parent.add("panel", undefined, L("panel.offset") + "（" + getCurrentRulerUnitLabel() + "）");
            offsetPanel.orientation = "column";
            offsetPanel.alignChildren = ["fill", "top"];
            offsetPanel.margins = UI_CONFIG.panelMargins;

            /* 単位はパネル名に入れているので入力欄には付けない
               The unit is in the panel title, so the fields carry none */
            var offsetUI = buildLinkedQuadUI(offsetPanel, {
                value: String(calcDefaultInnerOffset()),
                characters: 3
            });

            editInnerOffsetTop = offsetUI.top;
            editInnerOffsetBottom = offsetUI.bottom;
            editInnerOffsetLeft = offsetUI.left;
            editInnerOffsetRight = offsetUI.right;
            chkInnerOffsetLink = offsetUI.link;

            applyInnerOffsetLinkState = function () {
                _syncingInnerOffsets = true;
                offsetUI.applyLinkState();
                _syncingInnerOffsets = false;
            };

            applyInnerOffsetLinkState();
        }

        /* 内側エリア：列・行と行オプション / Inner area columns, rows and their options */
        function buildInnerGridUI(parent) {
            var gridWrapper = parent.add("group");
            gridWrapper.orientation = "row";
            gridWrapper.alignChildren = ["left", "top"];
            gridWrapper.alignment = ["fill", "top"];

            var gridColumnGroup = gridWrapper.add("group");
            gridColumnGroup.orientation = "column";
            gridColumnGroup.alignChildren = ["left", "top"];
            gridColumnGroup.alignment = ["left", "top"];
            gridColumnGroup.spacing = 12;

            var columnFields = addCountSpacingPanel(gridColumnGroup, "panel.columns", "field.columnCount");
            editInnerColumns = columnFields.count;
            editColGutter = columnFields.gutter;
            editColGutter.enabled = (toCount(editInnerColumns.text) > 1);

            var rowFields = addCountSpacingPanel(gridColumnGroup, "panel.rows", "field.rowCount");
            editInnerRows = rowFields.count;
            editRowGutter = rowFields.gutter;

            // 塗り・分割線のオプション（中央寄せ）
            var gridOptionsWrapper = gridColumnGroup.add("group");
            gridOptionsWrapper.orientation = "row";
            gridOptionsWrapper.alignChildren = ["center", "center"];
            gridOptionsWrapper.alignment = ["fill", "top"];

            var gridOptionsGroup = gridOptionsWrapper.add("group");
            gridOptionsGroup.orientation = "row";
            gridOptionsGroup.alignChildren = ["left", "center"];
            gridOptionsGroup.alignment = ["center", "center"];

            chkInnerFill = gridOptionsGroup.add("checkbox", undefined, L("checkbox.fill"));
            chkInnerFill.value = false;

            chkInnerDivider = gridOptionsGroup.add("checkbox", undefined, L("checkbox.divider"));
            chkInnerDivider.value = false;
        }

        /* 内側エリア：分割線の線種 / Inner area divider line type */
        function buildInnerLineTypeUI(parent) {
            innerLineTypePanel = parent.add("panel", undefined, L("panel.lineType"));
            innerLineTypePanel.orientation = "row";
            innerLineTypePanel.alignChildren = ["left", "center"];
            innerLineTypePanel.margins = UI_CONFIG.panelMargins;

            rbInnerLineSolid = innerLineTypePanel.add("radiobutton", undefined, L("lineType.solid"));
            rbInnerLineDash = innerLineTypePanel.add("radiobutton", undefined, L("lineType.dash"));
            rbInnerLineDotDash = innerLineTypePanel.add("radiobutton", undefined, L("lineType.dots"));
            rbInnerLineSolid.value = true;
        }

        // タイトルエリアのデフォルト高さ：外側エリア高 / titleSizeDivisor（現在の rulerType 単位）
        function calcDefaultTitleSize() {
            try {
                var bounds = null;
                if (targetItems && targetItems.length > 0) {
                    bounds = targetItems[0].geometricBounds; // [L, T, R, B]
                } else {
                    var artboardIndex = doc.artboards.getActiveArtboardIndex();
                    bounds = doc.artboards[artboardIndex].artboardRect; // [L, T, R, B]
                }

                var heightPt = Math.abs(bounds[1] - bounds[3]);
                if (!(heightPt > 0)) return 10;

                var factor = getCurrentRulerPtFactor(); // unit -> pt
                if (!factor || factor === 0) factor = 1;

                var divisor = GRID_CONFIG.titleSizeDivisor;
                if (!divisor || divisor <= 0) divisor = 5;

                var titleSize = (heightPt / factor) / divisor;
                if (!(titleSize > 0)) return 10;

                // 0.1 単位で丸め（整数に近ければ整数化）
                var rounded = Math.round(titleSize * 10) / 10;
                if (Math.abs(rounded - Math.round(rounded)) < 1e-6) rounded = Math.round(rounded);
                if (rounded < 0.1) rounded = 0.1;
                return rounded;
            } catch (_) { }
            return 10;
        }

        /* タイトルエリアのサイズが入力されているか / Does the title area have a size? */
        function titleHasSize() {
            return (parseFloat(editTitleSize.text) > 0);
        }

        // 0→>0 の瞬間だけ［線］を自動ONにするためのフラグ（ユーザーは後からOFF可）
        var _prevTitleHasSize = false;

        /* タイトルエリアUIを組み立てる / Build the title area panel */
        function buildTitleUI(parent) {
            // タイトルエリア
            titlePanel = parent.add("panel", undefined, L("panel.titleArea"));
            titlePanel.orientation = "column";
            titlePanel.alignChildren = ["fill", "top"];
            titlePanel.margins = UI_CONFIG.panelMargins;
            titlePanel.spacing = 10;

            // 有効 ＋ 幅／高さ（1行）
            var titleEnableRow = titlePanel.add("group");
            titleEnableRow.orientation = "row";
            titleEnableRow.alignChildren = ["left", "center"];
            titleEnableRow.spacing = 0;

            chkTitleEnable = titleEnableRow.add("checkbox", undefined, "");
            chkTitleEnable.value = false; // デフォルトOFF
            /* ラベルがないので、空文字ぶんの幅が入らないよう抑える
               No label, so cap the width to avoid the phantom text space */
            try { chkTitleEnable.preferredSize.width = UI_CONFIG.toggleCheckboxWidth; } catch (_) { }

            // 幅／高さ（有効チェックの右側）
            titleSizeGroup = titleEnableRow.add("group");
            titleSizeGroup.orientation = "row";
            titleSizeGroup.alignChildren = ["left", "center"];
            titleSizeGroup.margins = 0;
            titleSizeGroup.spacing = 4;

            titleSizeGroup.add("statictext", undefined, fieldLabel("field.titleSize"));
            editTitleSize = titleSizeGroup.add("edittext", undefined, "0");
            editTitleSize.characters = 4;
            titleSizeGroup.add("statictext", undefined, getCurrentRulerUnitLabel());
            changeValueByArrowKey(editTitleSize, false);

            // 位置（上/下/左/右）
            titlePositionGroup = titlePanel.add("group");
            titlePositionGroup.orientation = "row";
            titlePositionGroup.alignChildren = ["left", "center"];

            rbTitleTop = titlePositionGroup.add("radiobutton", undefined, L("position.top"));
            rbTitleBottom = titlePositionGroup.add("radiobutton", undefined, L("position.bottom"));
            rbTitleLeft = titlePositionGroup.add("radiobutton", undefined, L("position.left"));
            rbTitleRight = titlePositionGroup.add("radiobutton", undefined, L("position.right"));

            // デフォルト：上
            rbTitleTop.value = true;

            // 塗り／線
            titleOptionGroup = titlePanel.add("group");
            titleOptionGroup.orientation = "row";
            titleOptionGroup.alignChildren = ["left", "center"];

            chkTitleFill = titleOptionGroup.add("checkbox", undefined, L("checkbox.fill"));
            chkTitleFill.value = false;

            chkTitleLine = titleOptionGroup.add("checkbox", undefined, L("panel.line"));
            chkTitleLine.value = true;

            // タイトルエリア：辺の伸縮
            titleEdgeScaleRow = titlePanel.add("group");
            titleEdgeScaleRow.orientation = "row";
            titleEdgeScaleRow.alignChildren = ["left", "center"];

            chkTitleEdgeScale = titleEdgeScaleRow.add("checkbox", undefined, L("checkbox.edgeScale"));
            chkTitleEdgeScale.value = false;

            editTitleEdgeScale = titleEdgeScaleRow.add("edittext", undefined, "0");
            editTitleEdgeScale.characters = 4;
            changeValueByArrowKey(editTitleEdgeScale, true);

            titleEdgeScaleRow.add("statictext", undefined, getCurrentRulerUnitLabel());

            bindTitleUI();
        }

        /* タイトルエリアの状態反映とイベントを設定する / Wire up the title area state and events */
        function bindTitleUI() {
            /* 選択中のタイトル位置キーを返す / The selected title position */
            getTitlePosKey = function () {
                if (rbTitleRight.value) return "right";
                if (rbTitleBottom.value) return "bottom";
                if (rbTitleLeft.value) return "left";
                return "top";
            };

            /* タイトル帯に、外側エリアの角丸値で2角だけ角丸を適用する
               Round the two matching corners of the title band with the outer radius */
            maybeApplyTitleAreaRound = function (titleRect) {
                try {
                    return !!applyTitleAreaCornerRounding(titleRect, getTitlePosKey(), getOuterRoundPt());
                } catch (_) { }
                return false;
            };

            /* ［辺の伸縮］の入力欄の有効／無効を反映 / Reflect the title edge scale checkbox */
            applyTitleEdgeScaleEnabledState = function () {
                editTitleEdgeScale.enabled = !!chkTitleEdgeScale.value;
                if (!chkTitleEdgeScale.value) editTitleEdgeScale.text = "0";
            };
            applyTitleEdgeScaleEnabledState();

            chkTitleEdgeScale.onClick = function () {
                applyTitleEdgeScaleEnabledState();
                requestPreview();
            };

            editTitleEdgeScale.onChanging = function () {
                if (!chkTitleEdgeScale.value) return;
                requestPreview();
            };

            /**
             * タイトルエリアの各コントロールの有効／無効を反映します。
             *
             * @returns {void}
             */
            applyTitleAreaEnabledState = function () {
                var areaEnabled = !!chkTitleEnable.value;

                // 無効なら 0 扱い（=タイトル生成なし）
                if (!areaEnabled) {
                    chkTitleFill.value = false;
                    chkTitleLine.value = false;
                }

                /* 幅／高さはディムせず常に入力可。OFFのときはチェックを外すだけで値は残す
                   Never dim the size; turning the checkbox off keeps the value as-is */
                titleSizeGroup.enabled = true;
                titleOptionGroup.enabled = areaEnabled;

                // ［辺の伸縮］は「タイトル有効」かつ「線ON」のときのみ操作可能
                titleEdgeScaleRow.enabled = (areaEnabled && !!chkTitleLine.value);
                if (!titleEdgeScaleRow.enabled) {
                    chkTitleEdgeScale.value = false;
                    applyTitleEdgeScaleEnabledState();
                }

                // 位置・塗り・線は「有効」かつ「サイズ>0」のときのみ
                var usable = (areaEnabled && titleHasSize());
                titlePositionGroup.enabled = usable;

                chkTitleFill.enabled = usable;
                if (!usable) chkTitleFill.value = false;

                // 線：サイズが 0→>0 になった瞬間だけ自動ON
                if (usable && !_prevTitleHasSize) chkTitleLine.value = true;
                chkTitleLine.enabled = usable;
                if (!usable) chkTitleLine.value = false;

                _prevTitleHasSize = usable;
            };

            // 初期反映
            applyTitleAreaEnabledState();

            // 位置変更でプレビュー更新
            bindAll([rbTitleTop, rbTitleBottom, rbTitleLeft, rbTitleRight], "onClick", function () {
                requestPreview();
            });

            // サイズ変更
            editTitleSize.onChanging = function () {
                /* OFFのときは入力だけ受け付け、プレビューには反映しない
                   While unchecked the field just stores the value; nothing is previewed */
                if (!chkTitleEnable.value) return;

                applyTitleAreaEnabledState();
                requestPreview();
            };

            // 有効切替
            chkTitleEnable.onClick = function () {
                /* ONにしたとき、現在値が0ならデフォルト値を入れる
                   Fill in the default size when enabled at zero */
                if (chkTitleEnable.value && !titleHasSize()) {
                    editTitleSize.text = String(calcDefaultTitleSize());
                }

                applyTitleAreaEnabledState();
                requestPreview();
            };
        }

        /* フレームUIを組み立てる（アートボード基準のときだけ使う）
           Build the frame panel; used for artboard-based runs only */
        function buildFrameUI(parent) {
            framePanel = parent.add("panel", undefined, L("panel.frame"));
            framePanel.orientation = "column";
            framePanel.alignChildren = ["fill", "top"];
            framePanel.margins = UI_CONFIG.panelMargins;
            framePanel.spacing = 10;

            var frameRow = framePanel.add("group");
            frameRow.orientation = "row";
            frameRow.alignChildren = ["left", "center"];
            frameRow.spacing = 0;

            chkFrameEnable = frameRow.add("checkbox", undefined, "");
            chkFrameEnable.value = false; /* デフォルトOFF / off by default */
            /* ラベルがないので、空文字ぶんの幅が入らないよう抑える
               No label, so cap the width to avoid the phantom text space */
            try { chkFrameEnable.preferredSize.width = UI_CONFIG.toggleCheckboxWidth; } catch (_) { }

            chkFrameEnable.onClick = function () {
                if (chkFrameEnable.value) {
                    /* 幅が0なら既定値を入れ、裁ち落としも自動でON（手動操作があれば尊重）
                       Fill in the default width at zero and turn the bleed on, unless the user set it */
                    var width = parseFloat(editFrameWidth.text);
                    if (isNaN(width) || width === 0) editFrameWidth.text = GRID_CONFIG.defaultFrameWidth;
                    if (_usingArtboardBase && !_bleedManuallySet) chkBleed.value = true;
                }

                applyFrameEnabledState();
                requestPreview();
            };

            /* 幅（ラベル・入力欄・単位をまとめて配置）
               Width; the label, field and unit share one group */
            frameWidthGroup = frameRow.add("group");
            frameWidthGroup.orientation = "row";
            frameWidthGroup.alignChildren = ["left", "center"];
            frameWidthGroup.margins = 0;
            frameWidthGroup.spacing = 4;

            frameWidthGroup.add("statictext", undefined, fieldLabel("field.width"));
            editFrameWidth = frameWidthGroup.add("edittext", undefined, "0");
            editFrameWidth.characters = 4;
            frameWidthGroup.add("statictext", undefined, getCurrentRulerUnitLabel());
            changeValueByArrowKey(editFrameWidth, false);

            /**
             * フレーム関連UIの有効／無効を、基準の種類とフレーム幅に応じて反映します。
             *
             * フレームはアートボード基準のときだけ使えるため、長方形スタート時はパネルごと畳みます。
             *
             * @returns {void}
             */
            applyFrameEnabledState = function () {
                setPanelCollapsed(framePanel, !_usingArtboardBase);

                if (!_usingArtboardBase) {
                    /* 長方形スタート時はフレームを使わないので、値も状態もリセット
                       Frames are unavailable for rectangle-based runs, so reset everything */
                    chkFrameEnable.value = false;
                    chkBleed.value = false;
                    chkFrameRound.value = false;
                    editFrameWidth.text = "0";
                    editFrameRound.text = "0";

                    frameWidthGroup.enabled = false;
                    chkBleed.enabled = false;
                    chkFrameRound.enabled = false;
                    editFrameRound.enabled = false;
                    return;
                }

                var enabled = !!chkFrameEnable.value;
                /* 幅はディムせず常に入力可。OFFのときはチェックを外すだけで値は残す
                   Never dim the width; turning the checkbox off keeps the value as-is */
                frameWidthGroup.enabled = true;

                var width = parseFloat(editFrameWidth.text);
                var hasWidth = (!isNaN(width) && width > 0);

                chkBleed.enabled = (enabled && hasWidth);
                chkFrameRound.enabled = (enabled && hasWidth);
                if (!enabled) {
                    chkBleed.value = false;
                    chkFrameRound.value = false;
                }
                editFrameRound.enabled = (enabled && hasWidth && chkFrameRound.value);
            };

            /* 裁ち落とし（表示は常時。長方形スタート時は enabled で制御）
               The bleed row is always visible; rectangle-based runs disable it instead */
            var bleedRow = framePanel.add("group");
            bleedRow.orientation = "row";
            bleedRow.alignChildren = ["left", "center"];

            chkBleed = bleedRow.add("checkbox", undefined, L("checkbox.bleed"));
            chkBleed.value = false;

            /* フレームの角丸 / Rounded corners of the frame */
            var frameRoundRow = framePanel.add("group");
            frameRoundRow.orientation = "row";
            frameRoundRow.alignChildren = ["left", "center"];

            chkFrameRound = frameRoundRow.add("checkbox", undefined, L("checkbox.frameRound"));
            chkFrameRound.value = false;

            editFrameRound = frameRoundRow.add("edittext", undefined, "0");
            editFrameRound.characters = 4;
            changeValueByArrowKey(editFrameRound, false);
            frameRoundRow.add("statictext", undefined, getCurrentRulerUnitLabel());

            /* 初期反映（フレーム幅0なら裁ち落とし・角丸はディム）
               Initial state: a zero width dims both the bleed and the rounding */
            applyFrameEnabledState();
        }

        // =========================================
        // UIの組み立て（タブごと） / Build the panels, tab by tab
        // =========================================
        buildMarginUI(marginTabColumn);
        buildFrameUI(marginTabColumn);
        buildOuterUI(outerTabColumn);
        buildTitleUI(outerTabColumn);
        buildInnerUI(innerTabColumn);
        buildDisplayUI(displayTabColumn);

        /* 長方形スタート時（アートボード基準でない場合）は左タブ全体を非表示
           Hide the whole left tab for rectangle-based runs */
        if (!_usingArtboardBase) {
            tabMargin.visible = false;
            tabMargin.enabled = false;
            tabPanel.selection = tabOuter;
        }

        /**
         * ［画面表示］タブを組み立てます（ズームとパン、表示コマンド）。
         *
         * @param {Group} parent - 追加先のグループ。
         * @returns {void}
         */
        function buildDisplayUI(parent) {
            var zoomPanPanel = parent.add("panel", undefined, L("panel.zoomPan"));
            zoomPanPanel.orientation = "column";
            zoomPanPanel.alignChildren = ["fill", "top"];
            zoomPanPanel.margins = UI_CONFIG.panelMargins;
            zoomPanPanel.spacing = 10;

            var zoomPanGroup = zoomPanPanel.add("group");
            zoomPanGroup.orientation = "column";
            zoomPanGroup.alignChildren = "left";
            zoomPanGroup.spacing = 8;

            if (viewCtl && typeof viewCtl.buildUI === "function") {
                viewCtl.buildUI(zoomPanGroup, {
                    labelWidth: UI_CONFIG.viewLabelWidth,
                    sliderWidth: UI_CONFIG.viewSliderWidth
                });
            } else {
                zoomPanGroup.add("statictext", undefined, L("view.unavailable"));
            }

            var viewCommandPanel = parent.add("panel", undefined, L("panel.display"));
            viewCommandPanel.orientation = "column";
            viewCommandPanel.alignChildren = ["fill", "top"];
            viewCommandPanel.margins = UI_CONFIG.panelMargins;
            viewCommandPanel.spacing = 10;

            var viewCommandGroup = viewCommandPanel.add("group");
            viewCommandGroup.orientation = "column";
            viewCommandGroup.alignChildren = ["left", "top"];
            viewCommandGroup.alignment = ["fill", "top"];
            viewCommandGroup.spacing = 6;

            /* 表示コマンド [ラベル, 実行する処理] / View commands: [label, action] */
            var commands = [
                [L("button.fitIn"), runMenuCommand("fitin")],
                [L("button.actualSize"), runMenuCommand("actualsize")],
                [L("button.fitAll"), runMenuCommand("fitall")],
                [L("button.zoomOut10"), function () {
                    /* 現在の表示倍率を10%縮小（スライダーも追従）
                       Shrink the current zoom by 10%; the slider follows */
                    if (viewCtl && typeof viewCtl.zoomBy === "function") viewCtl.zoomBy(0.9);
                }]
            ];

            for (var i = 0; i < commands.length; i++) {
                addViewCommandButton(viewCommandGroup, commands[i][0], commands[i][1]);
            }
        }

        /**
         * メニューコマンドを実行する処理を作って返します。
         *
         * @param {string} menuCommand - 実行するメニューコマンド名。
         * @returns {Function} クリック時に実行する処理。
         */
        function runMenuCommand(menuCommand) {
            return function () {
                try { app.executeMenuCommand(menuCommand); } catch (_) { }
            };
        }

        /**
         * 表示コマンド用の小さめボタンを1つ追加します。
         *
         * @param {Group} parent - 追加先のグループ。
         * @param {string} label - ボタンのラベル。
         * @param {Function} action - クリック時に実行する処理。
         * @returns {Button} 追加したボタン。
         */
        function addViewCommandButton(parent, label, action) {
            var button = parent.add("button", undefined, label);
            button.alignment = "left";

            /* サイズ系プロパティは環境差で落ちることがあるため保護
               Size properties can throw on some hosts, so guard them */
            try {
                button.preferredSize = [UI_CONFIG.viewButtonWidth, UI_CONFIG.viewButtonHeight];
                button.minimumSize = [UI_CONFIG.viewButtonWidth, UI_CONFIG.viewButtonHeight];
                button.maximumSize = [UI_CONFIG.viewButtonWidth, UI_CONFIG.viewButtonHeight];
            } catch (_) { }

            button.onClick = action;
            return button;
        }

        /* 手動操作を優先するためのフラグ（自動ONを抑制）
           Flags that keep manual input from being overwritten by the auto-on rules */
        var _innerFillManuallySet = false; /* 内側エリアの［塗り］ / inner area fill */
        var _bleedManuallySet = false;   /* フレームの［裁ち落とし］ / frame bleed */

        // 列/行の分割が可能になった瞬間だけ「分割線」を自動ONにするためのフラグ
        var _prevGridSplittable = isGridSplittable(toCount(editInnerColumns.text), toCount(editInnerRows.text));

        /* 列数・行数から分割線を引けるかを判定 / Can the grid carry dividers? */
        function isGridSplittable(colCount, rowCount) {
            return (colCount > 1 || rowCount > 1);
        }

        /**
         * ［分割線］と［線の種類］パネルの有効／無効を反映します。
         *
         * @param {number} colCount - 列数。
         * @param {number} rowCount - 行数。
         * @param {boolean} allowAutoOn - 1/1から分割可能になった瞬間に自動ONしてよい場合は true。
         * @returns {void}
         */
        function applyInnerDividerEnabledState(colCount, rowCount, allowAutoOn) {
            var splittable = isGridSplittable(colCount, rowCount);
            chkInnerDivider.enabled = splittable;

            if (!splittable) {
                // 分割できないなら分割線は不要
                chkInnerDivider.value = false;
            } else if (allowAutoOn && !_prevGridSplittable) {
                // 1/1 から分割可能になった瞬間だけ自動ON
                chkInnerDivider.value = true;
            }

            _prevGridSplittable = splittable;

            // 線種パネルも連動
            innerLineTypePanel.enabled = (splittable && chkInnerDivider.value);
        }

        // 初期状態：列/行が1/1なら分割線はディム（OFF）、分割可能ならON/OFFに従う
        applyInnerDividerEnabledState(toCount(editInnerColumns.text), toCount(editInnerRows.text), false);

        // =========================================
        // 下段（左：プレビュー ／ 右：ボタン） / Footer: preview on the left, buttons on the right
        // =========================================
        var footerRow = dialog.add("group");
        footerRow.orientation = "row";
        footerRow.alignChildren = ["left", "center"];
        footerRow.alignment = ["fill", "center"];

        var chkPreview = footerRow.add("checkbox", undefined, L("checkbox.preview"));
        chkPreview.value = true; /* 最初からプレビューON / preview starts enabled */
        chkPreview.alignment = "left";

        /* スペーサー（右側のボタンを押し出す） / Spacer that pushes the buttons to the right */
        var spacer = footerRow.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var buttonGroup = footerRow.add("group");
        buttonGroup.alignment = ["right", "center"];
        var cancelButton = buttonGroup.add("button", undefined, L("button.cancel"), { name: "cancel" });
        /* ［実行］は name:"ok" により dialog.show() が 1 を返すので、参照は保持しない
           The OK button is identified by name:"ok", so no reference is needed */
        buttonGroup.add("button", undefined, L("button.ok"), { name: "ok" });

        /**
         * ダイアログを開く前の表示状態に戻します。
         *
         * @returns {void}
         */
        function restoreView() {
            if (viewCtl && typeof viewCtl.restore === "function") viewCtl.restore();
        }

        /* キャンセル・クローズのどちらでも表示状態を戻す
           Both cancel and close restore the original view */
        cancelButton.onClick = function () {
            restoreView();
            clearPreview();
            dialog.close(0);
        };

        dialog.onClose = function () {
            restoreView();
            return true;
        };

        // =========================================
        // セッションの保存と復元 / Save and restore the session
        // =========================================

        /**
         * 入れ子のオブジェクトから、ドット区切りのパスで値を取り出します。
         *
         * @param {Object} source - 探索するオブジェクト。
         * @param {string} path - 例 "inner.grid.cols" のようなパス。
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
         * @param {string} path - 例 "inner.grid.cols" のようなパス。
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

        /**
         * 保存値を新形式（入れ子）→旧形式（フラット）の順に探します。
         *
         * @param {Object} state - 保存されていた状態。
         * @param {string} path - 新形式のパス。
         * @param {string} legacyKey - 旧形式のキー。
         * @returns {*} 見つかった値。存在しない場合は undefined。
         */
        function readState(state, path, legacyKey) {
            var value = getStateValue(state, path);
            if (typeof value !== "undefined") return value;
            return legacyKey ? state[legacyKey] : undefined;
        }

        /* 保存と復元はこの定義表を共有します（項目の追加・変更はここだけ）
           Save and restore share these tables, so a field is defined in one place */

        /* 入力欄 [コントロール, 新形式のパス, 旧形式のキー]
           Text fields: [control, nested path, legacy key] */
        var SESSION_TEXT_FIELDS = [
            [editArtboardMarginTop, "margin.top", "marginTop"],
            [editArtboardMarginBottom, "margin.bottom", "marginBottom"],
            [editArtboardMarginLeft, "margin.left", "marginLeft"],
            [editArtboardMarginRight, "margin.right", "marginRight"],
            [editOuterEdgeScale, "outer.lenVal", "lenVal"],
            [editOuterRound, "outer.round.val", "outerRoundVal"],
            [editTitleSize, "title.size", "titleSize"],
            [editTitleEdgeScale, "title.edgeScale.val", "titleEdgeScaleVal"],
            [editFrameWidth, "frame.width", "frameWidth"],
            [editFrameRound, "frame.round.val", "frameRoundVal"],
            [editInnerOffsetTop, "inner.offset.top", "offTop"],
            [editInnerOffsetBottom, "inner.offset.bottom", "offBottom"],
            [editInnerOffsetLeft, "inner.offset.left", "offLeft"],
            [editInnerOffsetRight, "inner.offset.right", "offRight"],
            [editInnerColumns, "inner.grid.cols", "cols"],
            [editInnerRows, "inner.grid.rows", "rows"],
            [editColGutter, "inner.grid.gutter.col", "colGutter"],
            [editRowGutter, "inner.grid.gutter.row", "rowGutter"]
        ];

        /* チェックボックス [コントロール, 新形式のパス, 旧形式のキー] / Checkboxes */
        var SESSION_CHECKBOXES = [
            [chkPreview, "preview", "preview"],
            [chkArtboardMarginLink, "margin.link", "marginLink"],
            [chkKeepOuter, "outer.keepOuter", "keepOuter"],
            [chkOuterEdgeScale, "outer.enableLen", "enableLen"],
            [chkOuterRound, "outer.round.enable", "outerRoundEnable"],
            [chkTitleEnable, "title.enable", "titleEnable"],
            [chkTitleFill, "title.fill", "titleFill"],
            [chkTitleLine, "title.line", "titleLine"],
            [chkTitleEdgeScale, "title.edgeScale.enable", "titleEdgeScaleEnable"],
            [chkFrameEnable, "frame.enable", "frameEnable"],
            [chkBleed, "frame.bleed", "bleed"],
            [chkFrameRound, "frame.round.enable", "frameRound"],
            [chkInnerOffsetLink, "inner.link", "innerLink"],
            [chkInnerFill, "inner.grid.fill", "rowFill"],
            [chkInnerDivider, "inner.grid.divider", "rowDivider"]
        ];

        /* ラジオボタン [新形式のパス, 旧形式のキー, 保存値→コントロール, 既定の保存値]
           Radio buttons: [nested path, legacy key, stored value to control, default] */
        var SESSION_RADIO_GROUPS = [
            ["outer.cap", "cap", { butt: rbCapButt, round: rbCapRound, project: rbCapProject }, "butt"],
            ["title.pos", "titlePos", { top: rbTitleTop, bottom: rbTitleBottom, left: rbTitleLeft, right: rbTitleRight }, "top"],
            ["inner.grid.lineType", "innerLine", { solid: rbInnerLineSolid, dash: rbInnerLineDash, dotdash: rbInnerLineDotDash }, "solid"]
        ];

        /**
         * ラジオボタン群から、選択中の保存値を返します。
         *
         * @param {Object} controlsByKey - 保存値→コントロールの対応。
         * @param {string} fallbackKey - どれも選択されていない場合の値。
         * @returns {string} 選択中の保存値。
         */
        function readRadioKey(controlsByKey, fallbackKey) {
            for (var key in controlsByKey) {
                if (!controlsByKey.hasOwnProperty(key)) continue;
                if (controlsByKey[key].value) return key;
            }
            return fallbackKey;
        }

        /**
         * 前回のダイアログ設定を復元します（Illustratorの起動中のみ有効）。
         *
         * @returns {void}
         */
        function restoreUIState() {
            var state = loadSessionState();
            if (!state) return;

            var i, value;

            for (i = 0; i < SESSION_TEXT_FIELDS.length; i++) {
                value = readState(state, SESSION_TEXT_FIELDS[i][1], SESSION_TEXT_FIELDS[i][2]);
                if (typeof value !== "undefined") SESSION_TEXT_FIELDS[i][0].text = String(value);
            }

            for (i = 0; i < SESSION_CHECKBOXES.length; i++) {
                value = readState(state, SESSION_CHECKBOXES[i][1], SESSION_CHECKBOXES[i][2]);
                if (typeof value !== "undefined") SESSION_CHECKBOXES[i][0].value = !!value;
            }

            for (i = 0; i < SESSION_RADIO_GROUPS.length; i++) {
                value = readState(state, SESSION_RADIO_GROUPS[i][0], SESSION_RADIO_GROUPS[i][1]);
                var button = SESSION_RADIO_GROUPS[i][2][value];
                if (button) button.value = true;
            }

            /* 角丸と辺の伸縮は同時にONにできない（角丸を優先）
               The radius and the edge scale cannot both be on; the radius wins */
            if (chkOuterRound.value) chkOuterEdgeScale.value = false;

            /* 復元は「サイズが0→>0になった瞬間」ではないため、
               タイトルの［線］が自動ONで上書きされないように直前の状態をそろえる
               A restore is not a zero-to-positive transition, so seed the previous state
               and keep the saved title line value from being auto-enabled */
            _prevTitleHasSize = (!!chkTitleEnable.value && titleHasSize());

            /* 他のコントロールに依存する有効／無効を反映し直す
               Re-apply the enabled states that depend on other controls */
            applyArtboardMarginLinkState();
            applyInnerOffsetLinkState();
            applyOuterEdgeScaleEnabledState();
            applyStrokeCapPanelEnabledState();
            applyOuterRoundEnabledState();
            applyTitleEdgeScaleEnabledState();
            applyTitleAreaEnabledState();
            applyFrameEnabledState();
            applyInnerDividerEnabledState(toCount(editInnerColumns.text), toCount(editInnerRows.text), false);
        }
        restoreUIState();

        // =========================================
        // イベント処理 / Event handlers
        // =========================================

        /**
         * プレビューがONのときだけプレビューを更新します。
         *
         * @returns {void}
         */
        function requestPreview() {
            if (chkPreview.value) updatePreview(false);
        }

        /**
         * 4方向の入力欄に同じ値を書き込みます（連動用）。
         *
         * @param {EditText} source - 値の入力元。
         * @param {EditText[]} targets - 反映先の入力欄。
         * @returns {void}
         */
        function syncLinkedFields(source, targets) {
            for (var i = 0; i < targets.length; i++) {
                targets[i].text = source.text;
            }
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

        /* 外側エリア / Outer area */
        editOuterEdgeScale.onChanging = function () {
            if (!chkOuterEdgeScale.value) return;
            applyStrokeCapPanelEnabledState();
            requestPreview();
        };

        chkKeepOuter.onClick = function () {
            applyOuterEdgeScaleEnabledState();
            applyStrokeCapPanelEnabledState();
            applyOuterRoundEnabledState();
            requestPreview();
        };

        chkOuterEdgeScale.onClick = function () {
            /* 角丸と辺の伸縮は同時に使えないので、ONにした側を残して他方をOFFにする
               The radius and the edge scale cannot be combined, so turning one on turns the other off */
            if (chkOuterEdgeScale.value) chkOuterRound.value = false;

            applyOuterEdgeScaleEnabledState();
            applyStrokeCapPanelEnabledState();
            applyOuterRoundEnabledState();
            requestPreview();
        };

        bindAll([rbCapButt, rbCapRound, rbCapProject], "onClick", requestPreview);

        /* マージン（アートボード基準のみ） / Margins, artboard-based runs only */
        editArtboardMarginTop.onChanging = function () {
            if (!_usingArtboardBase || _syncingArtboardMargins) return;

            if (chkArtboardMarginLink.value) {
                _syncingArtboardMargins = true;
                syncLinkedFields(editArtboardMarginTop, [editArtboardMarginBottom, editArtboardMarginLeft, editArtboardMarginRight]);
                _syncingArtboardMargins = false;
            }
            requestPreview();
        };

        bindAll([editArtboardMarginBottom, editArtboardMarginLeft, editArtboardMarginRight], "onChanging", function () {
            if (!_usingArtboardBase || _syncingArtboardMargins) return;
            requestPreview();
        });

        chkArtboardMarginLink.onClick = function () {
            applyArtboardMarginLinkState();
            requestPreview();
        };

        /* タイトルエリア / Title area */
        chkTitleFill.onClick = requestPreview;

        chkTitleLine.onClick = function () {
            applyTitleAreaEnabledState();
            requestPreview();
        };

        /* フレーム / Frame */
        chkBleed.onClick = function () {
            /* ユーザーが操作したら以後は自動ONしない / Stop auto-enabling once the user decides */
            _bleedManuallySet = true;
            requestPreview();
        };

        editFrameWidth.onChanging = function () {
            /* OFFのときは入力だけ受け付け、プレビューには反映しない
               While unchecked the field just stores the value; nothing is previewed */
            if (!chkFrameEnable.value) return;

            var width = parseFloat(editFrameWidth.text);
            var hasWidth = (!isNaN(width) && width > 0);

            /* 幅が0→>0になったら裁ち落としを自動ON（手動操作があれば尊重）
               Turn the bleed on once a width is set, unless the user set it manually */
            chkBleed.enabled = (_usingArtboardBase && hasWidth);
            if (!chkBleed.enabled) chkBleed.value = false;
            else if (!_bleedManuallySet) chkBleed.value = true;

            chkFrameRound.enabled = hasWidth;
            if (!hasWidth) {
                chkFrameRound.value = false;
                editFrameRound.enabled = false;
            }
            requestPreview();
        };

        chkFrameRound.onClick = function () {
            editFrameRound.enabled = chkFrameRound.value;

            /* ONにしたとき、0なら既定値を入れる / Fill in the default radius when enabled at zero */
            if (chkFrameRound.value) {
                var radius = parseFloat(editFrameRound.text);
                if (isNaN(radius) || radius === 0) editFrameRound.text = GRID_CONFIG.defaultRound;
            }
            requestPreview();
        };

        editFrameRound.onChanging = requestPreview;

        /* 内側エリア：オフセット / Inner area offsets */
        editInnerOffsetTop.onChanging = function () {
            if (_syncingInnerOffsets) return;

            if (chkInnerOffsetLink.value) {
                _syncingInnerOffsets = true;
                syncLinkedFields(editInnerOffsetTop, [editInnerOffsetBottom, editInnerOffsetLeft, editInnerOffsetRight]);
                _syncingInnerOffsets = false;
            }
            requestPreview();
        };

        bindAll([editInnerOffsetBottom, editInnerOffsetLeft, editInnerOffsetRight], "onChanging", function () {
            if (_syncingInnerOffsets) return;
            requestPreview();
        });

        chkInnerOffsetLink.onClick = function () {
            applyInnerOffsetLinkState();
            requestPreview();
        };

        /* 内側エリア：列・行 / Inner area columns and rows */
        bindAll([editInnerColumns, editInnerRows], "onChanging", function () {
            var columns = toCount(editInnerColumns.text);
            var rows = toCount(editInnerRows.text);

            /* 1未満や空欄は1に補正 / Snap blank or sub-one input back to 1 */
            if (String(columns) !== editInnerColumns.text) editInnerColumns.text = String(columns);
            if (String(rows) !== editInnerRows.text) editInnerRows.text = String(rows);

            editColGutter.enabled = (columns > 1);
            applyInnerDividerEnabledState(columns, rows, true);
            requestPreview();
        });

        bindAll([editColGutter, editRowGutter], "onChanging", function () {
            /* ガターが入ったら塗りを自動ON（手動操作があれば尊重）
               Turn the fill on once a gutter is set, unless the user set it manually */
            var gutter = parseFloat(this.text);
            if (!_innerFillManuallySet && !isNaN(gutter) && gutter !== 0) {
                chkInnerFill.value = true;
            }
            requestPreview();
        });

        chkInnerFill.onClick = function () {
            /* ユーザーが操作したら以後は自動ONしない / Stop auto-enabling once the user decides */
            _innerFillManuallySet = true;
            requestPreview();
        };

        chkInnerDivider.onClick = function () {
            innerLineTypePanel.enabled = (chkInnerDivider.enabled && chkInnerDivider.value);
            requestPreview();
        };

        bindAll([rbInnerLineSolid, rbInnerLineDash, rbInnerLineDotDash], "onClick", requestPreview);

        /* プレビュー切り替え / Preview toggle */
        chkPreview.onClick = function () {
            if (chkPreview.value) updatePreview(false);
            else clearPreview();
        };

        // レイアウト確定（tabbedpanel の内容が潰れるのを防ぐ）
        try { dialog.layout.layout(true); } catch (_) { }
        try { dialog.layout.resize(); } catch (_) { }
        updatePreview(false);

        // --- ダイアログ表示 ---
        var result = dialog.show();

        persistUIState();

        // キャンセル時：プレビュー生成物を削除して終了
        if (result != 1) {
            restoreView();
            clearPreview();
            return;
        }

        // =========================================
        // 実行後の後処理 / Post-processing after OK
        // =========================================

        /* 最終生成（この結果はヒストリーに残す） / Final generation, kept in the history */
        updatePreview(true);

        applyOuterFrameResult();
        var keptFills = applyInnerFillResult();

        /* タイトル帯とフレームは残し、選択対象からだけ外す
           Keep the title band and the frame, but drop them from the selection list */
        var keptBands = takeTaggedItems(TAG_TITLE_FILL).concat(takeTaggedItems(TAG_FRAME_FILL));

        /* 内部タグは note だけに残し、レイヤーパネルに出る name はクリアする
           Keep the tag in note only and clear the name shown in the Layers panel */
        clearTagNames(keptBands);
        clearTagNames(keptFills);
        clearTagNames(tempPreviewItems);

        /* アートボード基準の矩形は残す（キャンセル時のみ clearPreview() が破棄する）
           The artboard-based rectangle stays; only a cancel removes it in clearPreview() */
        try { doc.selection = null; } catch (_) { }

        // =========================================
        // 実行後の後処理の実装 / Post-processing helpers
        // =========================================

        /**
         * 基準の長方形をすべて削除します（アートボード基準の一時矩形も含む）。
         *
         * 非表示のまま残すと目に見えない残骸になるため、対象は必ず削除して参照も外します。
         *
         * @returns {void}
         */
        function removeTargets() {
            removeItems(targetItems);
            targetItems = [];
            _artboardBaseRect = null;
        }

        /**
         * tempPreviewItems から指定タグのオブジェクトを取り出します（配列からは外します）。
         *
         * @param {string} tag - 対象のタグ文字列。
         * @returns {PageItem[]} 取り出したオブジェクトの配列。
         */
        function takeTaggedItems(tag) {
            var found = [];
            for (var i = tempPreviewItems.length - 1; i >= 0; i--) {
                if (!hasTag(tempPreviewItems[i], tag)) continue;
                found.push(tempPreviewItems[i]);
                tempPreviewItems.splice(i, 1);
            }
            return found;
        }

        /**
         * 実行後の外枠の扱いを確定します（元の長方形の削除／表示と、不要な4辺線の削除）。
         *
         * @returns {void}
         */
        function applyOuterFrameResult() {
            if (getEffectiveOuterEdgeScale() !== 0) {
                /* 4辺線が外形を置き換えたので、元の長方形は削除する
                   （残すと非表示のまま毎回蓄積してしまう）
                   The edge lines replace the outline, so remove the rectangle;
                   keeping it would leave an invisible leftover on every run */
                removeTargets();
            } else if (chkKeepOuter.value) {
                setTargetsHidden(false);
            } else {
                /* 外枠を残さないので削除 / The outer frame is not kept, so remove it */
                removeTargets();
            }

            /* ［外枠を残す］がOFFなら、生成した4辺線も削除
               Drop the generated edges when the outer frame is not kept */
            if (chkKeepOuter.value) return;
            removeItems(takeTaggedItems(TAG_OUTER_EDGE));
        }

        /**
         * 実行後の内側エリアの塗りを確定します（塗りONなら通常オブジェクトとして残し、OFFなら削除）。
         *
         * @returns {PageItem[]} 残した塗り矩形。塗りOFFのときは空配列。
         */
        function applyInnerFillResult() {
            var fills = takeTaggedItems(TAG_INNER_FILL);

            /* 塗りOFFならまとめて削除 / Drop them all when the fill is off */
            if (!chkInnerFill.value) {
                removeItems(fills);
                return [];
            }

            /* 通常オブジェクトとして保持（ガイド属性は解除）
               Keep them as regular objects and clear the guide attribute */
            for (var i = 0; i < fills.length; i++) {
                try {
                    fills[i].guides = false;
                    fills[i].stroked = false;
                    fills[i].filled = true;
                    fills[i].fillColor = makeGrayColor(FILL_TINTS.innerBox);
                } catch (_) { }
            }
            return fills;
        }

        // =========================================
        // プレビューと生成 / Preview and generation
        // =========================================
        // collectOptions()      : UIを読み、pt単位の生成条件にまとめる
        // generateFromOptions() : 生成条件からオブジェクトを作る
        // updatePreview()       : 前回分を消して生成し、再描画する（唯一の入口）
        // -----------------------------------------

        /**
         * タイトル帯の仕切り線の「両端の詰め量」を返します。
         *
         * 入力値は外側エリアと同じ向き（＋で長く／−で短く）なので、
         * 詰め量に変換するため符号を反転します（詰め量＋で線が短くなる）。
         *
         * @param {number} factor - 単位→ptの換算係数。
         * @returns {number} 詰め量（pt）。
         */
        function readTitleDividerInsetPt(factor) {
            if (!chkTitleLine.value || !chkTitleEdgeScale.value) return 0;
            return -toPt(editTitleEdgeScale.text, factor);
        }

        /**
         * フレーム幅を返します（裁ち落としONなら加算）。
         *
         * @param {number} factor - 単位→ptの換算係数。
         * @returns {number} フレーム幅（pt）。
         */
        function readFramePt(factor) {
            var framePt = chkFrameEnable.value ? toPositivePt(editFrameWidth.text, factor) : 0;
            if (_bleedEnabled) framePt += mmToPt(GRID_CONFIG.bleedMm);
            return framePt;
        }

        /**
         * 他の入力値に依存するコントロールを更新します（collectOptions の副作用をまとめたもの）。
         *
         * @param {Object} options - collectOptions() が組み立てた生成条件。
         * @returns {void}
         */
        function syncDependentControls(options) {
            /* ガターは列／行が2以上のときだけ入力可 / Gutters are editable for 2+ columns or rows only */
            editColGutter.enabled = (options.colCount > 1);
            editRowGutter.enabled = (options.rowCount > 1);

            /* ガターが入ったら塗りを自動ON（手動操作があれば尊重）
               Turn the fill on once a gutter is set, unless the user set it manually */
            if (!_innerFillManuallySet && (options.colGutterPt !== 0 || options.rowGutterPt !== 0)) {
                chkInnerFill.value = true;
            }

            applyInnerDividerEnabledState(options.colCount, options.rowCount, false);
            applyStrokeCapPanelEnabledState();
        }

        /* UIの入力値を読み、pt単位の生成条件にまとめる
           Read the dialog and build the generation options in points */
        function collectOptions() {
            var factor = getCurrentRulerPtFactor();

            /* 裁ち落としの状態を先に確定させる（framePt と getActiveArtboardBounds() の両方が参照する）
               Resolve the bleed first: both framePt and getActiveArtboardBounds() read it */
            _bleedEnabled = !!chkBleed.value;

            var colCount = toCount(editInnerColumns.text);
            var rowCount = toCount(editInnerRows.text);

            var options = {
                factor: factor,

                /* 外側エリア：辺の伸縮 / Outer area: edge scale */
                outerEdgeScalePt: getEffectiveOuterEdgeScale() * factor,

                /* タイトルエリア / Title area */
                titleSizePt: chkTitleEnable.value ? toPositivePt(editTitleSize.text, factor) : 0,
                titleDividerInsetPt: readTitleDividerInsetPt(factor),

                /* フレーム / Frame */
                framePt: readFramePt(factor),

                /* 内側エリア：オフセット / Inner area offsets */
                offTopPt: toPt(editInnerOffsetTop.text, factor),
                offBottomPt: toPt(editInnerOffsetBottom.text, factor),
                offLeftPt: toPt(editInnerOffsetLeft.text, factor),
                offRightPt: toPt(editInnerOffsetRight.text, factor),

                /* 内側エリア：列・行とガター（列／行が1のときガターは0扱い）
                   Columns, rows and gutters; gutters are ignored for a single column or row */
                colCount: colCount,
                rowCount: rowCount,
                colGutterPt: (colCount > 1) ? toPositivePt(editColGutter.text, factor) : 0,
                rowGutterPt: (rowCount > 1) ? toPositivePt(editRowGutter.text, factor) : 0
            };

            syncDependentControls(options);
            return options;
        }

        /* アートボード基準の一時矩形を、現在のマージンで作り直す
           Rebuild the temporary artboard-based rectangle from the current margins */
        function rebuildBaseRectFromMargins(factor) {
            if (!_usingArtboardBase) return;

            var top = toPositivePt(editArtboardMarginTop.text, factor);
            var bottom = toPositivePt(editArtboardMarginBottom.text, factor);
            var left = toPositivePt(editArtboardMarginLeft.text, factor);
            var right = toPositivePt(editArtboardMarginRight.text, factor);

            /* 裁ち落としは collectOptions() で確定済み / The bleed is already resolved in collectOptions() */
            rebuildArtboardBaseRect(top, right, bottom, left);
        }

        /* タイトル帯とその分割線を全対象に作成
           Create the title band and its divider for every target */
        function createTitleParts(options) {
            if (options.titleSizePt <= 0) return;
            for (var i = 0; i < targetItems.length; i++) {
                createTitleFill(targetItems[i], options.titleSizePt);
                createTitleDivider(targetItems[i], options.titleSizePt, options.titleDividerInsetPt);
            }
        }

        /* 内側エリアを全対象に作成 / Create the inner area for every target */
        function createInnerParts(options) {
            for (var i = 0; i < targetItems.length; i++) {
                var innerBounds = getInnerAreaBounds(targetItems[i], options.titleSizePt);
                if (!innerBounds) continue;
                createInnerArea(targetItems[i], innerBounds, options);
            }
        }

        /* 外側エリアの角丸（辺の伸縮OFFのときだけ）
           Rounded corners for the outer area, only while the edge scale is off */
        function applyOuterAreaRound(isFinal) {
            if (!chkKeepOuter.value || chkOuterEdgeScale.value) return;

            var radiusPt = getOuterRoundPt();
            if (!(radiusPt > 0)) return;

            for (var i = 0; i < targetItems.length; i++) {
                var sourceRect = targetItems[i];
                if (!sourceRect || sourceRect.typename !== "PathItem") continue;

                if (isFinal) {
                    /* 実行時：元の外枠へライブエフェクトを適用
                       On OK: apply the live effect to the original frame */
                    applyRoundCornersEffect(sourceRect, radiusPt);
                    sourceRect.hidden = false;
                    continue;
                }

                /* プレビュー時：元を隠し、同じ位置に角丸用の一時矩形を作る
                   On preview: hide the original and round a temporary copy instead */
                var bounds = sourceRect.geometricBounds; // [L, T, R, B]
                var width = bounds[2] - bounds[0];
                var height = bounds[1] - bounds[3];
                if (!(width > 0) || !(height > 0)) continue;

                sourceRect.hidden = true;

                var previewRect = trackTempItem(sourceRect.layer.pathItems.rectangle(bounds[1], bounds[0], width, height));
                copyAppearance(sourceRect, previewRect);
                tagItem(previewRect, TAG_OUTER_ROUND);
                applyRoundCornersEffect(previewRect, radiusPt);
            }
        }

        /* 生成条件からオブジェクトを作成する（プレビュー・実行の共通処理）
           Create the objects from the options, shared by preview and final run */
        function generateFromOptions(options, isFinal) {
            rebuildBaseRectFromMargins(options.factor);

            /* 線端パネルの有効／無効は applyStrokeCapPanelEnabledState() が唯一の管理者
               （collectOptions() から呼ばれる）ため、ここでは触らない
               applyStrokeCapPanelEnabledState() owns the cap panel state, so don't touch it here */
            var splitEdges = (options.outerEdgeScalePt !== 0);

            if (splitEdges) {
                /* 4辺に分解するため、元の長方形は常に隠す
                   Always hide the original rectangle while the edges are split */
                setTargetsHidden(true);
                if (chkKeepOuter.value) {
                    for (var i = 0; i < targetItems.length; i++) {
                        createOuterEdgeLines(targetItems[i], options.outerEdgeScalePt);
                    }
                }
            } else {
                /* 外枠の表示は「外枠を残す」に従う
                   Show the original frame according to the keep-outer checkbox */
                setTargetsHidden(!chkKeepOuter.value);
            }

            /* フレーム（アートボード基準） / Frame, based on the artboard */
            if (options.framePt > 0) {
                var artboardBounds = getActiveArtboardBounds();
                if (artboardBounds) createFrameFill(targetItems[0], options.framePt, artboardBounds);
            }

            createTitleParts(options);
            if (!splitEdges) applyOuterAreaRound(isFinal);
            createInnerParts(options);
        }

        /* プレビューまたは最終結果を作り直して再描画する
           Rebuild the preview or the final result, then redraw */
        function updatePreview(isFinal) {
            removeTempItems();
            generateFromOptions(collectOptions(), !!isFinal);
            app.redraw();
        }

        /**
         * 現在のダイアログ設定をセッションに保存します。
         *
         * @returns {void}
         */
        function persistUIState() {
            var state = {};
            var i;

            try {
                for (i = 0; i < SESSION_TEXT_FIELDS.length; i++) {
                    setStateValue(state, SESSION_TEXT_FIELDS[i][1], SESSION_TEXT_FIELDS[i][0].text);
                }

                for (i = 0; i < SESSION_CHECKBOXES.length; i++) {
                    setStateValue(state, SESSION_CHECKBOXES[i][1], !!SESSION_CHECKBOXES[i][0].value);
                }

                for (i = 0; i < SESSION_RADIO_GROUPS.length; i++) {
                    var group = SESSION_RADIO_GROUPS[i];
                    setStateValue(state, group[0], readRadioKey(group[2], group[3]));
                }
            } catch (_) { }

            saveSessionState(state);
        }

        /**
         * 入力欄の文字列を、現在の単位からpt値に換算します。
         *
         * @param {string} text - 入力欄の文字列。
         * @param {number} factor - 単位→ptの換算係数。
         * @returns {number} pt値。数値として読めない場合は 0。
         */
        function toPt(text, factor) {
            var value = parseFloat(text);
            return isNaN(value) ? 0 : (value * factor);
        }

        /**
         * 入力欄の文字列を、0以上のpt値に換算します（マージンなど負を許さない項目用）。
         *
         * @param {string} text - 入力欄の文字列。
         * @param {number} factor - 単位→ptの換算係数。
         * @returns {number} 0以上のpt値。
         */
        function toPositivePt(text, factor) {
            var valuePt = toPt(text, factor);
            return (valuePt > 0) ? valuePt : 0;
        }

        /**
         * 入力欄の文字列を、1以上の整数（列数・行数）に変換します。
         *
         * @param {string} text - 入力欄の文字列。
         * @returns {number} 1以上の整数。
         */
        function toCount(text) {
            var value = parseInt(text, 10);
            return (isNaN(value) || value < 1) ? 1 : value;
        }

        /**
         * アクティブなアートボードの矩形を返します（裁ち落としONなら外側に広げます）。
         *
         * @returns {number[]|null} [左, 上, 右, 下] の座標。取得できない場合は null。
         */
        function getActiveArtboardBounds() {
            try {
                var artboardIndex = doc.artboards.getActiveArtboardIndex();
                var rect = doc.artboards[artboardIndex].artboardRect; // [L,T,R,B]
                if (!_bleedEnabled) return [rect[0], rect[1], rect[2], rect[3]];

                var bleedPt = mmToPt(GRID_CONFIG.bleedMm);
                return [rect[0] - bleedPt, rect[1] + bleedPt, rect[2] + bleedPt, rect[3] - bleedPt];
            } catch (_) {
                return null;
            }
        }

        /**
         * 辺の伸縮の値を返します（チェックがOFFのときは 0 とみなします）。
         *
         * @returns {number} 現在の単位での辺の伸縮量。
         */
        function getEffectiveOuterEdgeScale() {
            if (!chkOuterEdgeScale || !chkOuterEdgeScale.value) return 0;
            var value = parseFloat(editOuterEdgeScale.text);
            return isNaN(value) ? 0 : value;
        }

        /**
         * 対象オブジェクトの表示／非表示をまとめて切り替えます。
         *
         * @param {boolean} hidden - 非表示にする場合は true。
         * @returns {void}
         */
        function setTargetsHidden(hidden) {
            for (var i = 0; i < targetItems.length; i++) {
                try { targetItems[i].hidden = hidden; } catch (_) { }
            }
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
                try { event.preventDefault(); } catch (_) { }

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
                } catch (_) { }
            });
        }

        // UIで選択された線端を取得
        function getSelectedStrokeCap() {
            if (rbCapRound.value) return StrokeCap.ROUNDENDCAP;
            if (rbCapProject.value) return StrokeCap.PROJECTINGENDCAP;
            return StrokeCap.BUTTENDCAP; // 線端なし
        }

        /* ドット点線を適用（線端を丸型にして strokeDashes=[0, 線幅*2]）
           Dotted line: round caps with strokeDashes = [0, width * 2] */
        function applyDotDash(pathItem) {
            pathItem.stroked = true;
            pathItem.strokeCap = StrokeCap.ROUNDENDCAP;
            try { pathItem.strokeJoin = StrokeJoin.ROUNDENDJOIN; } catch (_) { }
            pathItem.strokeDashes = [0, pathItem.strokeWidth * 2];
        }

        /**
         * 内側の区切り線に線種（実線・点線・ドット点線）を適用します。
         *
         * @param {PathItem} pathItem - 対象の分割線。
         * @returns {void}
         */
        function applyInnerLineStyle(pathItem) {
            try {
                // ドット点線
                if (rbInnerLineDotDash.value) {
                    pathItem.strokeWidth = 2;
                    applyDotDash(pathItem);
                    return;
                }

                pathItem.strokeWidth = 1;
                // 点線（ダッシュ）は破線パターン、実線は破線なし
                pathItem.strokeDashes = rbInnerLineDash.value ? [4, 2] : [];
                // 線端は外枠の線端設定に合わせる
                pathItem.strokeCap = getSelectedStrokeCap();
            } catch (_) { }
        }

        /**
         * 角丸のライブエフェクトを適用します。
         *
         * @param {PageItem} item - 対象のオブジェクト。
         * @param {number} radiusPt - 角丸の半径（pt）。0以下なら何もしません。
         * @returns {void}
         */
        function applyRoundCornersEffect(item, radiusPt) {
            if (!item || !(radiusPt > 0)) return;
            var effectXml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius #value# "/></LiveEffect>';
            try { item.applyEffect(effectXml.replace('#value#', radiusPt)); } catch (_) { }
        }

        // フレームを作成して tempPreviewItems に登録（外枠基準 or 指定bounds）
        // 外：FILL_TINTS.frame / 内：透明の穴あき（グループ＋Live Pathfinder Exclude）
        function createFrameFill(pathItem, framePt, baseBounds) {
            try {
                if (!framePt || framePt <= 0) return;

                var bounds = (baseBounds && baseBounds.length === 4) ? baseBounds : pathItem.geometricBounds; // [L,T,R,B]
                var left = bounds[0], top = bounds[1], right = bounds[2], bottom = bounds[3];
                var width = right - left;
                var height = top - bottom;
                if (!(width > 0) || !(height > 0)) return;

                /* 内側は最終的に穴になる矩形 / The inner rectangle becomes the hole */
                var holeLeft = left + framePt;
                var holeTop = top - framePt;
                var holeRight = right - framePt;
                var holeBottom = bottom + framePt;
                var holeWidth = holeRight - holeLeft;
                var holeHeight = holeTop - holeBottom;
                if (!(holeWidth > 0) || !(holeHeight > 0)) return;

                var layer = pathItem.layer;

                // 外側・内側の矩形を作成
                var outerRect = trackTempItem(layer.pathItems.rectangle(top, left, width, height));
                outerRect.stroked = false;
                outerRect.filled = true;
                outerRect.fillColor = makeGrayColor(FILL_TINTS.frame);

                var innerRect = trackTempItem(layer.pathItems.rectangle(holeTop, holeLeft, holeWidth, holeHeight));
                innerRect.stroked = false;
                innerRect.filled = true;
                // 内側は一時的な塗り（色は最終的に Exclude の結果で穴になる）
                innerRect.fillColor = makeGrayColor(0);

                // 角丸は「内側の長方形」に適用する（穴側を丸める）
                applyRoundCornersEffect(innerRect, getFrameRoundPt());

                // グループ化して穴あきにする
                var group = trackTempItem(layer.groupItems.add());
                outerRect.move(group, ElementPlacement.PLACEATEND);
                innerRect.move(group, ElementPlacement.PLACEATEND);

                var resultItem = applyPathfinderExclude(group);
                if (resultItem !== group) trackTempItem(resultItem);
                tagItem(resultItem, TAG_FRAME_FILL);
                sendToBack(resultItem);
            } catch (_) { }
        }

        /**
         * 外側エリアの角丸半径を返します（OFFや未入力なら 0）。
         *
         * タイトルエリアの角丸もこの値を参照します。
         *
         * @returns {number} 角丸の半径（pt）。
         */
        function getOuterRoundPt() {
            if (!chkOuterRound.value) return 0;
            return toPositivePt(editOuterRound.text, getCurrentRulerPtFactor());
        }

        /**
         * フレームの角丸半径を返します（OFFや未入力なら 0）。
         *
         * @returns {number} 角丸の半径（pt）。
         */
        function getFrameRoundPt() {
            if (!chkFrameRound.value) return 0;
            return toPositivePt(editFrameRound.text, getCurrentRulerPtFactor());
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
            var prevSel = null;
            try { prevSel = doc.selection; } catch (_) { }

            var resultItem = null;
            try {
                doc.selection = null;
                group.selected = true;
                app.executeMenuCommand('Live Pathfinder Exclude');

                /* 結果は selection の先頭に入る / The result lands at the head of the selection */
                if (doc.selection && doc.selection.length > 0) resultItem = doc.selection[0];
            } catch (_) { }

            try { doc.selection = prevSel; } catch (_) { }

            return resultItem ? resultItem : group;
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

            var positionKey = getTitlePosKey();
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

            return {
                band: {
                    top: top,
                    left: (positionKey === "left") ? left : (right - sizePt),
                    width: sizePt,
                    height: height
                },
                divider: [[dividerX, top - dividerInsetPt], [dividerX, bottom + dividerInsetPt]],
                inner: (positionKey === "left") ? [dividerX, top, right, bottom] : [left, top, dividerX, bottom]
            };
        }

        /* タイトル帯：塗り矩形を作成して tempPreviewItems に登録
           The title band fill */
        function createTitleFill(pathItem, sizePt) {
            try {
                if (!chkTitleFill.value || !(sizePt > 0)) return;

                var layout = calcTitleAreaLayout(pathItem.geometricBounds, sizePt, 0);
                if (!layout) return;

                var bandRect = trackTempItem(pathItem.layer.pathItems.rectangle(
                    layout.band.top, layout.band.left, layout.band.width, layout.band.height));
                bandRect.stroked = false;
                bandRect.filled = true;
                bandRect.fillColor = makeGrayColor(FILL_TINTS.titleBand);

                tagItem(bandRect, TAG_TITLE_FILL);
                maybeApplyTitleAreaRound(bandRect);

                /* 背面へ（他の罫線や要素の下に敷く） / Send behind the rules */
                sendToBack(bandRect);
            } catch (_) { }
        }

        /* タイトル帯と本文の仕切り線（titleDividerInsetPt：＋で両端を短く／−で伸ばす）
           The divider between the band and the body */
        function createTitleDivider(target, titleSizePt, titleDividerInsetPt) {
            try {
                if (!target || target.typename !== "PathItem") return;
                if (!(titleSizePt > 0) || !chkTitleLine.value) return;

                var layout = calcTitleAreaLayout(target.geometricBounds, titleSizePt, titleDividerInsetPt);
                if (!layout) return;

                /* 他の生成物と同じレイヤーに作る（doc.activeLayer は使わない）
                   Create it on the same layer as the other generated items */
                var dividerLine = trackTempItem(target.layer.pathItems.add());
                dividerLine.setEntirePath(layout.divider);
                dividerLine.stroked = true;
                dividerLine.filled = false;
                dividerLine.strokeColor = makeGrayColor(FILL_TINTS.line);
                dividerLine.strokeWidth = 1;

                tagItem(dividerLine, TAG_TITLE_DIVIDER);
            } catch (_) { }
        }

        /* タイトル領域を除いた「内側エリア」の計算領域を返す
           The inner area bounds with the title band removed: [L, T, R, B] or null */
        function getInnerAreaBounds(pathItem, titleSizePt) {
            try {
                var bounds = pathItem.geometricBounds; // [L, T, R, B]

                if (!(titleSizePt > 0)) {
                    var hasArea = ((bounds[2] - bounds[0]) > 0 && (bounds[1] - bounds[3]) > 0);
                    return hasArea ? bounds : null;
                }

                var layout = calcTitleAreaLayout(bounds, titleSizePt, 0);
                return layout ? layout.inner : null;
            } catch (_) {
                return null;
            }
        }

        /**
         * 内側エリアのグリッド配置を計算します。
         *
         * @param {number[]} bounds - 基準領域 [左, 上, 右, 下]。
         * @param {Object} options - collectOptions() が返す生成条件。
         * @returns {Object|null} セル配置。成立しない場合は null。
         */
        function calcInnerGrid(bounds, options) {
            /* オフセットが0でも内側エリアは描画する / The inner area is drawn even at zero offset */
            var offTop = (options.offTopPt > 0) ? options.offTopPt : 0;
            var offBottom = (options.offBottomPt > 0) ? options.offBottomPt : 0;
            var offLeft = (options.offLeftPt > 0) ? options.offLeftPt : 0;
            var offRight = (options.offRightPt > 0) ? options.offRightPt : 0;

            var width = (bounds[2] - bounds[0]) - (offLeft + offRight);
            var height = (bounds[1] - bounds[3]) - (offTop + offBottom);
            if (!(width > 0) || !(height > 0)) return null;

            var cols = toCount(options.colCount);
            var rows = toCount(options.rowCount);
            var gutterX = (options.colGutterPt > 0) ? options.colGutterPt : 0;
            var gutterY = (options.rowGutterPt > 0) ? options.rowGutterPt : 0;

            /* ガターを除いた1セルの大きさ / Cell size with the gutters removed */
            var cellW = (width - gutterX * (cols - 1)) / cols;
            var cellH = (height - gutterY * (rows - 1)) / rows;
            if (!(cellW > 0) || !(cellH > 0)) return null;

            return {
                left: bounds[0] + offLeft,
                top: bounds[1] - offTop,
                width: width,
                height: height,
                cols: cols,
                rows: rows,
                gutterX: gutterX,
                gutterY: gutterY,
                cellW: cellW,
                cellH: cellH
            };
        }

        /**
         * 内側エリアのセル（塗り）を作成し、tempPreviewItems に登録します。
         *
         * @param {Layer} layer - 作成先のレイヤー。
         * @param {Object} grid - calcInnerGrid() が返すセル配置。
         * @returns {void}
         */
        function createInnerCellFills(layer, grid) {
            for (var row = 0; row < grid.rows; row++) {
                var cellTop = grid.top - (grid.cellH + grid.gutterY) * row;

                for (var col = 0; col < grid.cols; col++) {
                    var cellLeft = grid.left + (grid.cellW + grid.gutterX) * col;

                    var cellRect = trackTempItem(layer.pathItems.rectangle(cellTop, cellLeft, grid.cellW, grid.cellH));
                    cellRect.stroked = false;
                    cellRect.filled = true;
                    cellRect.fillColor = makeGrayColor(FILL_TINTS.innerBox);
                    tagItem(cellRect, TAG_INNER_FILL);

                    /* 背面へ（罫線などのパスの下に敷く） / Send behind the rules */
                    sendToBack(cellRect);
                }
            }
        }

        /**
         * 分割線の見た目を基準オブジェクトから取り出します（線がなければ K100 / 1pt）。
         *
         * @param {PathItem} pathItem - 基準の長方形。
         * @returns {Object} {color, width}
         */
        function getDividerAppearance(pathItem) {
            var color = null;
            var width = 0;

            try {
                if (pathItem.stroked) {
                    color = pathItem.strokeColor;
                    width = pathItem.strokeWidth;
                }
            } catch (_) { }

            return {
                color: color ? color : makeGrayColor(FILL_TINTS.line),
                width: width ? width : 1
            };
        }

        /**
         * 内側エリアの分割線を作成し、tempPreviewItems に登録します。
         *
         * @param {Layer} layer - 作成先のレイヤー。
         * @param {Object} grid - calcInnerGrid() が返すセル配置。
         * @param {Object} appearance - getDividerAppearance() が返す線の見た目。
         * @returns {void}
         */
        function createInnerDividers(layer, grid, appearance) {
            var bottomY = grid.top - grid.height;
            var rightX = grid.left + grid.width;

            /**
             * 分割線を1本作成します。
             *
             * @param {number[][]} points - [[x1, y1], [x2, y2]] 形式の始点と終点。
             * @returns {void}
             */
            function addDivider(points) {
                var dividerLine = trackTempItem(layer.pathItems.add());
                dividerLine.setEntirePath(points);
                dividerLine.stroked = true;
                dividerLine.filled = false;
                dividerLine.strokeColor = appearance.color;
                dividerLine.strokeWidth = appearance.width;
                applyInnerLineStyle(dividerLine);
            }

            /* 列の分割線：各ガターの中心に1本ずつ
               Column dividers, one at the centre of each gutter */
            for (var col = 1; col < grid.cols; col++) {
                var gutterCenterX = grid.left + (grid.cellW * col) + (grid.gutterX * (col - 1)) + (grid.gutterX / 2);
                addDivider([[gutterCenterX, grid.top], [gutterCenterX, bottomY]]);
            }

            /* 行の分割線：各ガターの中心に1本ずつ
               Row dividers, one at the centre of each gutter */
            for (var row = 1; row < grid.rows; row++) {
                var gutterCenterY = grid.top - (grid.cellH * row) - (grid.gutterY * (row - 1)) - (grid.gutterY / 2);
                addDivider([[grid.left, gutterCenterY], [rightX, gutterCenterY]]);
            }
        }

        /**
         * 内側エリアのセル（塗り）と分割線を作成します。
         *
         * @param {PathItem} pathItem - 基準の長方形（レイヤーと線の見た目の参照元）。
         * @param {number[]|null} baseBounds - 基準領域 [左, 上, 右, 下]。null なら pathItem の外形を使います。
         * @param {Object} options - collectOptions() が返す生成条件。
         * @returns {void}
         */
        function createInnerArea(pathItem, baseBounds, options) {
            /* 前提：軸に平行な長方形 / Assumes an axis-aligned rectangle: [左, 上, 右, 下] */
            var bounds = (baseBounds && baseBounds.length === 4) ? baseBounds : pathItem.geometricBounds;

            var grid = calcInnerGrid(bounds, options);
            if (!grid) return;

            var layer = pathItem.layer;
            createInnerCellFills(layer, grid);

            /* 分割線OFF、または分割できない構成ならここまで
               Stop here when the dividers are off or the grid cannot carry them */
            if (!chkInnerDivider.value || !isGridSplittable(grid.cols, grid.rows)) return;

            createInnerDividers(layer, grid, getDividerAppearance(pathItem));
        }

        /* プレビューを消して元の状態に戻す / Drop the preview and restore the original state */
        function clearPreview() {
            removeTempItems();

            /* アートボード基準の一時矩形を先に破棄（targetItemsに残っていると無効参照になる）
               Remove the temporary artboard rectangle first to avoid stale references */
            cleanupArtboardBaseRect();

            /* 角丸プレビューなどで隠した元オブジェクトを表示に戻す
               Show the originals that the preview had hidden */
            setTargetsHidden(false);

            try { applyStrokeCapPanelEnabledState(); } catch (_) { }

            app.redraw();
        }

        /* プレビューで作った一時アイテムを削除する / Drop the items created for the preview */
        function removeTempItems() {
            removeItems(tempPreviewItems);
            tempPreviewItems = [];
        }

        /**
         * 長方形の各辺を、伸縮させた直線として生成します（4辺に分解）。
         *
         * @param {PathItem} pathItem - 基準の長方形。
         * @param {number} edgeScalePt - 伸縮量（pt。正で伸ばし、負で縮める）。
         * @returns {void}
         */
        function createOuterEdgeLines(pathItem, edgeScalePt) {
            var points = pathItem.pathPoints;
            var isClosed = pathItem.closed;
            var edgeCount = isClosed ? points.length : points.length - 1;

            for (var i = 0; i < edgeCount; i++) {
                var startAnchor = points[i].anchor;
                var endAnchor = points[(i + 1) % points.length].anchor;

                var dx = endAnchor[0] - startAnchor[0];
                var dy = endAnchor[1] - startAnchor[1];
                var edgeLength = Math.sqrt(dx * dx + dy * dy);

                var scaleAmount = Math.abs(edgeScalePt);

                /* ※既存動作を踏襲：正の値のとき、辺が伸縮量の2倍以下なら生成しない
                   Kept as-is from the previous behaviour: skip short edges when the value is positive */
                if (edgeScalePt > 0 && edgeLength <= scaleAmount * 2) continue;

                var ratio = scaleAmount / edgeLength;

                /* 正なら両端を外へ、負なら内へ動かす / Positive extends the ends, negative pulls them in */
                var direction = (edgeScalePt >= 0) ? -1 : 1;
                var startX = startAnchor[0] + dx * ratio * direction;
                var startY = startAnchor[1] + dy * ratio * direction;
                var endX = endAnchor[0] - dx * ratio * direction;
                var endY = endAnchor[1] - dy * ratio * direction;

                var edgeLine = trackTempItem(pathItem.layer.pathItems.add());
                edgeLine.setEntirePath([[startX, startY], [endX, endY]]);

                edgeLine.stroked = true;
                edgeLine.filled = false;
                edgeLine.strokeColor = pathItem.strokeColor;
                edgeLine.strokeWidth = pathItem.strokeWidth;
                edgeLine.strokeCap = getSelectedStrokeCap();

                // 外枠（4辺）として識別できるようタグ付け
                tagItem(edgeLine, TAG_OUTER_EDGE);
            }
        }

    })();

})();
