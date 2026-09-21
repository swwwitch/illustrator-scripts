#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトにIllustratorの「自由変形」ライブ効果を適用します。
台形・平行四辺形・三角形・対角線の全18プリセットをアイコンから選び、調整可能なプリセットでは変形量と強度を設定できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartFreeDistort.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n15a7ae196a23

### Overview

Applies Illustrator's Free Distort live effect to the selected objects.
Eighteen presets — trapezoids, parallelograms, triangles and diagonals — are picked from icons, and the adjustable ones expose an amount and a strength setting.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartFreeDistort.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartFreeDistort";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.5.5";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-24";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-18";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartFreeDistort.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartFreeDistort.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n15a7ae196a23"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================
    /* 自由変形の座標は 0〜1 の正規化値で扱うため、定規単位には依存しない。
       Free Distort coordinates are normalized to 0-1, so the ruler unit is irrelevant. */
    var DISTORT_CONFIG = {
        amountMinPercent: 0,       /* 変形量スライダーの最小値（%） / slider minimum (percent) */
        amountMaxPercent: 49,      /* 変形量スライダーの最大値（%） / slider maximum (percent) */
        amountDefaultPercent: 20,  /* 変形量スライダーの初期値（%） / slider default (percent) */
        previewThrottleMs: 150,    /* プレビュー更新の最小間隔（ミリ秒） / minimum preview interval in ms */
        shearBaseFactor: 2         /* シアー量の基準倍率 / base multiplier for the shear amount */
    };

    /* 強度倍率。台形・平行四辺形で共通（平行四辺形はさらに shearBaseFactor を掛ける）
       Strength factors shared by trapezoid and parallelogram (the latter also uses shearBaseFactor) */
    var STRENGTH_FACTORS = {
        weak: 0.25,
        normal: 0.5,
        strong: 1.0
    };

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 実行環境のロケールから UI 言語を判定する。
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILanguage();

    /* UI 文言の定義 / UI string definitions */
    var LABELS = {
        dialog: {
            title: { ja: "スマート パスの自由変形", en: "Smart Free Distort" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: {
                ja: "オブジェクトを選択してください。",
                en: "Please select one or more objects."
            },
            noTarget: {
                ja: "選択内容には効果を適用できません。文字の一部ではなく、オブジェクトを選択してください。",
                en: "The selection cannot take an effect. Select whole objects rather than a range of text."
            },
            partialApply: {
                ja: "{count}件に適用した時点で中断しました。適用前に戻すには取り消してください。",
                en: "Stopped after applying to {count} object(s). Undo to return to the state before the run."
            }
        },
        panel: {
            trapezoid: { ja: "台形", en: "Trapezoid" },
            parallelogram: { ja: "平行四辺形", en: "Parallelogram" },
            triangle: { ja: "三角形", en: "Triangle" },
            diagonal: { ja: "対角線", en: "Diagonal" },
            adjust: { ja: "変形の調整", en: "Adjust" },
            amount: { ja: "変形量", en: "Amount" },
            strength: { ja: "強度", en: "Strength" }
        },
        trapezoid: {
            narrowTop: { ja: "上辺を狭く", en: "Narrow Top" },
            wideTop: { ja: "上辺を広く", en: "Wide Top" },
            narrowBottom: { ja: "下辺を狭く", en: "Narrow Bottom" },
            wideBottom: { ja: "下辺を広く", en: "Wide Bottom" }
        },
        shear: {
            axisHorizontal: { ja: "横にずらす", en: "Shear horizontally" },
            axisVertical: { ja: "縦にずらす", en: "Shear vertically" },
            anchorTopLeft: { ja: "左上", en: "Top Left" },
            anchorTopRight: { ja: "右上", en: "Top Right" },
            anchorBottomLeft: { ja: "左下", en: "Bottom Left" },
            anchorBottomRight: { ja: "右下", en: "Bottom Right" },
            tipFormat: {
                ja: "{axis}（固定：{anchor}）",
                en: "{axis} (fixed: {anchor})"
            }
        },
        triangle: {
            bottomLeft: { ja: "直角を左下に", en: "Right angle at bottom left" },
            bottomRight: { ja: "直角を右下に", en: "Right angle at bottom right" },
            topLeft: { ja: "直角を左上に", en: "Right angle at top left" },
            topRight: { ja: "直角を右上に", en: "Right angle at top right" }
        },
        diagonal: {
            backslash: { ja: "左上から右下へ", en: "Top left to bottom right" },
            slash: { ja: "右上から左下へ", en: "Top right to bottom left" }
        },
        strength: {
            weak: { ja: "弱", en: "Weak" },
            normal: { ja: "標準", en: "Normal" },
            strong: { ja: "強", en: "Strong" }
        },
        readout: {
            amount: { ja: "{amount}%", en: "{amount}%" },
            amountWithEffective: {
                ja: "{amount}%（実効 {effective}%）",
                en: "{amount}% (effective {effective}%)"
            }
        },
        checkbox: {
            preview: { ja: "プレビュー", en: "Preview" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        tooltip: {
            amount: {
                ja: "変形の大きさ（{min}〜{max}%）。［強度］の倍率が掛かります。",
                en: "How far the shape is distorted ({min}-{max}%). The Strength factor is applied on top."
            },
            strength: {
                ja: "変形量の{percent}%を適用します。",
                en: "Applies {percent}% of the amount."
            },
            preview: {
                ja: "チェックしている間、設定の変更をカンバスに反映します。",
                en: "While checked, every change is applied to the canvas."
            }
        }
    };

    /**
     * ドット区切りのパスで LABELS から文言を取得する。
     * @param {string} labelPath - "panel.amount" のようなドット区切りのキー
     * @returns {string} 現在の UI 言語の文言（見つからないときは labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var pathSegments = String(labelPath).split(".");
        var labelNode = LABELS;

        for (var i = 0; i < pathSegments.length; i++) {
            if (!labelNode || labelNode[pathSegments[i]] == null) return labelPath;
            labelNode = labelNode[pathSegments[i]];
        }

        if (labelNode[uiLang] != null) return labelNode[uiLang];
        if (labelNode.ja != null) return labelNode.ja;
        if (labelNode.en != null) return labelNode.en;
        return labelPath;
    }

    /**
     * テンプレート中の {name} を値で置き換える。UI 文言と自由変形 XML の両方で使う。
     * @param {string} template - {name} 形式のプレースホルダを含む文字列
     * @param {Object} values - プレースホルダ名をキーにした値のテーブル
     * @returns {string} 置換後の文字列（対応する値が無いプレースホルダはそのまま残る）
     */
    function fillPlaceholders(template, values) {
        return String(template).replace(/\{(\w+)\}/g, function (matched, placeholderName) {
            return (values[placeholderName] == null) ? matched : String(values[placeholderName]);
        });
    }

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 8;                  /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
    var BUTTON_ROW_TOP_MARGIN = 10;          /* ボタンエリアの上余白 / top margin of the button row */
    var AMOUNT_SLIDER_WIDTH = 220;           /* 変形量スライダーの幅（px） / amount slider width in pixels */

    /**
     * ウィンドウに共通のレイアウトを適用する。
     * @param {Window} win - 対象のダイアログウィンドウ
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(win, spacing) {
        win.orientation = "column";
        win.alignChildren = "fill";
        win.margins = WINDOW_MARGINS;
        win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルに共通のレイアウトを適用する。
     * @param {Panel} panel - 対象のパネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * 横並びグループに共通のレイアウトを適用する。
     * @param {Group} group - 対象のグループ
     * @param {(string|Array<string>)} [alignment] - 親の中での配置（省略時は "left"）
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(group, alignment, spacing) {
        group.orientation = "row";
        group.alignment = alignment || "left";
        group.alignChildren = ["left", "center"];
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* 変形プリセットのアイコンボタンの外観 / Appearance of the preset icon buttons */
    var ICON_BUTTON_SIZE = 40;          /* ボタンの一辺（px） / button side in pixels */
    var ICON_BUTTON_PADDING = 5;        /* ボタン枠とタイルの余白（px） / gap between the button edge and the tile */
    var ICON_GRID_COLUMNS = 2;          /* アイコンを並べる既定の列数 / default icons per row */
    var TRAPEZOID_ICON_COLUMNS = 1;     /* 台形は1列に縦積み / the trapezoid presets stack in one column */
    var ICON_GRID_SPACING = 4;          /* アイコン同士の間隔（px） / gap between icons */

    var ICON_TILE_LINE_WIDTH = 2;       /* タイルの枠線の太さ（px） / width of the tile frame */
    var ICON_SHAPE_SHRINK = 0.94;       /* 図形の縮小率（枠に触れさせない） / shrink so the shape clears the frame */
    var ICON_SHAPE_LINE_WIDTH = 1;      /* 図形の輪郭線の太さ（px） / outline width of a filled shape */
    var ICON_DIAGONAL_LINE_WIDTH = 3.5; /* 対角線の太さ（px） / stroke width of the diagonal presets */
    var ICON_SELECTED_LINE_WIDTH = 2;   /* 選択枠の太さ（px） / width of the selection frame */

    /* アイコンに描く変形量のサンプル値。変形の向きが 40px 角でも読み取れるよう、
       既定（変形量20%×強度［標準］）よりも誇張して描いている。形の傾向を示すための
       見本であって、適用結果の寸法そのものではない。
       Sample amounts used when drawing the icons. They are deliberately exaggerated compared
       with the default settings (20 percent at normal strength) so the direction of each
       distortion stays readable at 40 pixels. The icon shows the kind of distortion, not its
       exact magnitude. */
    var ICON_SAMPLE_TRAPEZOID_AMOUNT = 0.20;
    var ICON_SAMPLE_SHEAR_AMOUNT =
        ICON_SAMPLE_TRAPEZOID_AMOUNT * DISTORT_CONFIG.shearBaseFactor;

    /* この面積を下回る図形は線として描く（対角線プリセット用）
       Shapes below this area are drawn as a line; this covers the diagonal presets */
    var DEGENERATE_AREA_THRESHOLD = 0.0001;

    /* アイコンの配色（RGBA、各0〜1） / Icon colors, RGBA in the 0-1 range */
    var ICON_COLORS = {
        tileFill:   [0.91, 0.91, 0.91, 1],  /* 変形前の領域 / the area before distortion */
        tileBorder: [0.76, 0.76, 0.76, 1],  /* タイルの枠線 / the tile frame */
        distorted:  [0.09, 0.07, 0.06, 1],  /* 変形後の形 / the distorted shape */
        selected:   [0.20, 0.55, 0.95, 1]   /* 選択枠 / the selection frame */
    };

    // =========================================
    // プリセット定義 / Preset definitions
    // =========================================

    /* 変形前の4隅。順序は [TL, TR, BL, BR] で、自由変形XMLとアイコン描画の両方で使う。
       The corners before distortion, in [TL, TR, BL, BR] order, shared by the Free Distort
       XML and the icon drawing. */
    var SOURCE_CORNERS = [[0, 0], [1, 0], [0, 1], [1, 1]];

    /* makeCorners は変形後の4隅を [TL, TR, BL, BR] の順で返す。
       引数 trapezoidAmount / shearAmount は強度倍率を掛けたあとの変形量。
       makeCorners returns the destination corners in [TL, TR, BL, BR] order.
       Its arguments are the amounts after the strength factor has been applied.
       UIの並び順・ツールチップ・可変/固定・座標をこのテーブル1か所で管理する。
       Order, tooltip, adjustable flag and geometry all live in this single table. */
    var DISTORT_PRESETS = [
        { presetKey: "trapNarrowTop", tipText: getLabel("trapezoid.narrowTop"), groupKey: "trapezoid", adjustable: true,
          makeCorners: function (trapezoidAmount) {
              return [[trapezoidAmount, 0], [1 - trapezoidAmount, 0], [0, 1], [1, 1]];
          } },

        { presetKey: "trapWideTop", tipText: getLabel("trapezoid.wideTop"), groupKey: "trapezoid", adjustable: true,
          makeCorners: function (trapezoidAmount) {
              return [[-trapezoidAmount, 0], [1 + trapezoidAmount, 0], [0, 1], [1, 1]];
          } },

        { presetKey: "trapNarrowBottom", tipText: getLabel("trapezoid.narrowBottom"), groupKey: "trapezoid", adjustable: true,
          makeCorners: function (trapezoidAmount) {
              return [[0, 0], [1, 0], [trapezoidAmount, 1], [1 - trapezoidAmount, 1]];
          } },

        { presetKey: "trapWideBottom", tipText: getLabel("trapezoid.wideBottom"), groupKey: "trapezoid", adjustable: true,
          makeCorners: function (trapezoidAmount) {
              return [[0, 0], [1, 0], [-trapezoidAmount, 1], [1 + trapezoidAmount, 1]];
          } },

        /* 平行四辺形は 4基準点 × 2軸 の8パターンを後段で生成して追加する
           The eight parallelogram presets are generated further below */

        /* 三角形：名前の角と対角の頂点をたたむ。並び順はアイコングリッドの
           左上→右上→左下→右下に対応させる。
           Triangle: fold the corner opposite to the name. The order matches the icon
           grid, filled top-left, top-right, bottom-left, bottom-right. */
        { presetKey: "triTopLeft", tipText: getLabel("triangle.topLeft"), groupKey: "triangle", adjustable: false,
          makeCorners: function () { return [[0, 0], [1, 0], [0, 1], [0, 1]]; } },

        { presetKey: "triTopRight", tipText: getLabel("triangle.topRight"), groupKey: "triangle", adjustable: false,
          makeCorners: function () { return [[0, 0], [1, 0], [1, 1], [1, 1]]; } },

        { presetKey: "triBottomLeft", tipText: getLabel("triangle.bottomLeft"), groupKey: "triangle", adjustable: false,
          makeCorners: function () { return [[0, 0], [0, 0], [0, 1], [1, 1]]; } },

        { presetKey: "triBottomRight", tipText: getLabel("triangle.bottomRight"), groupKey: "triangle", adjustable: false,
          makeCorners: function () { return [[1, 0], [1, 0], [0, 1], [1, 1]]; } },

        /* 対角線：上辺と下辺をそれぞれ1点につぶす / Diagonal: collapse the top and bottom edges */
        { presetKey: "diagonalBackslash", tipText: getLabel("diagonal.backslash"), groupKey: "diagonal", adjustable: false,
          makeCorners: function () { return [[0, 0], [0, 0], [1, 1], [1, 1]]; } },

        { presetKey: "diagonalSlash", tipText: getLabel("diagonal.slash"), groupKey: "diagonal", adjustable: false,
          makeCorners: function () { return [[1, 0], [1, 0], [0, 1], [0, 1]]; } }
    ];

    /* 平行四辺形の基準点。x=0 が左、y=0 が上。
       Anchor corners for the parallelogram presets; x=0 is left, y=0 is top. */
    var SHEAR_ANCHORS = [
        { anchorKey: "TopLeft",     x: 0, y: 0 },
        { anchorKey: "TopRight",    x: 1, y: 0 },
        { anchorKey: "BottomLeft",  x: 0, y: 1 },
        { anchorKey: "BottomRight", x: 1, y: 1 }
    ];

    /* シアーの方向 / Shear axes */
    var SHEAR_AXES = [
        { axisKey: "Horizontal", labelPath: "shear.axisHorizontal" },
        { axisKey: "Vertical",   labelPath: "shear.axisVertical" }
    ];

    /**
     * 基準点の角を固定し、その角を含む辺を残したまま、反対側の辺を基準点から
     * 遠ざかる向きにずらした4隅を返す。4基準点 × 2軸 の8パターンが一意に決まる。
     * @param {Object} anchor - SHEAR_ANCHORS の要素（x / y は 0 または 1）
     * @param {string} axisKey - "Horizontal"（左右）または "Vertical"（上下）
     * @param {number} shearAmount - ずらす量（0〜1の正規化値）
     * @returns {Array<Array<number>>} 変形後の4隅 [TL, TR, BL, BR]
     */
    function makeShearCorners(anchor, axisKey, shearAmount) {
        var corners = [[0, 0], [1, 0], [0, 1], [1, 1]]; /* [TL, TR, BL, BR] */
        var i;

        if (axisKey === "Horizontal") {
            /* 基準点の反対側の横辺が、左右にずれる / the opposite horizontal edge slides sideways */
            var movingRow = 1 - anchor.y;
            var horizontalShift = (anchor.x === 0) ? shearAmount : -shearAmount;

            for (i = 0; i < corners.length; i++) {
                if (corners[i][1] === movingRow) corners[i][0] += horizontalShift;
            }
        } else {
            /* 基準点の反対側の縦辺が、上下にずれる / the opposite vertical edge slides up or down */
            var movingColumn = 1 - anchor.x;
            var verticalShift = (anchor.y === 0) ? shearAmount : -shearAmount;

            for (i = 0; i < corners.length; i++) {
                if (corners[i][0] === movingColumn) corners[i][1] += verticalShift;
            }
        }
        return corners;
    }

    /* 8パターンを生成してテーブルに追加する。軸ごとに4基準点を並べるので、
       2列のアイコングリッドでは上2行が左右、下2行が上下になる。
       Generate the eight presets; anchors are grouped by axis, so in the two-column icon
       grid the top two rows are the horizontal shears and the bottom two the vertical ones. */
    (function () {
        for (var axisIndex = 0; axisIndex < SHEAR_AXES.length; axisIndex++) {
            for (var anchorIndex = 0; anchorIndex < SHEAR_ANCHORS.length; anchorIndex++) {
                var axis = SHEAR_AXES[axisIndex];
                var anchor = SHEAR_ANCHORS[anchorIndex];

                DISTORT_PRESETS.push({
                    presetKey: "shear" + axis.axisKey + anchor.anchorKey,
                    tipText: fillPlaceholders(getLabel("shear.tipFormat"), {
                        axis: getLabel(axis.labelPath),
                        anchor: getLabel("shear.anchor" + anchor.anchorKey)
                    }),
                    groupKey: "parallelogram",
                    adjustable: true,
                    /* クロージャで軸と基準点を閉じ込める / capture the axis and anchor in a closure */
                    makeCorners: (function (capturedAnchor, capturedAxisKey) {
                        return function (trapezoidAmount, shearAmount) {
                            return makeShearCorners(capturedAnchor, capturedAxisKey, shearAmount);
                        };
                    })(anchor, axis.axisKey)
                });
            }
        }
    })();

    var DEFAULT_PRESET_INDEX = 0; /* 初期選択プリセット / initially selected preset */

    /* パネルに割り当てるプリセットグループ。値は DISTORT_PRESETS の groupKey と
       LABELS.panel のキーを兼ねる。
       Preset groups assigned to panels; each value doubles as a groupKey in
       DISTORT_PRESETS and as a key in LABELS.panel. */
    var TRAPEZOID_GROUP_KEY = "trapezoid";
    var SHEAR_GROUP_KEY = "parallelogram";
    var FIXED_GROUP_KEYS = ["triangle", "diagonal"];

    (function () {

        // =========================================
        // 事前チェック / Preflight checks
        // =========================================
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        if (!doc.selection || doc.selection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        // =========================================
        // 適用対象の収集 / Collecting the target objects
        // =========================================

        /**
         * 選択のうち、ライブ効果を適用できるオブジェクトだけを抽出する。
         * TextRange などは applyEffect を持たないため除外する。
         * @returns {Array<PageItem>} 効果を適用できる選択オブジェクト
         */
        function getSelectedEffectTargets() {
            var selectedItems = doc.selection;
            var targetItems = [];

            if (!selectedItems || selectedItems.length == null) return targetItems;

            for (var i = 0; i < selectedItems.length; i++) {
                if (selectedItems[i] && selectedItems[i].applyEffect != undefined) {
                    targetItems.push(selectedItems[i]);
                }
            }
            return targetItems;
        }

        /**
         * 保持しておいた参照のうち、まだ有効なものだけを返す。
         * Undo をまたぐと参照が無効になっていることがあるため、プロパティ参照で確かめる。
         * @param {Array<PageItem>} items - 検査するオブジェクトの配列
         * @returns {Array<PageItem>} まだ有効なオブジェクトだけの配列
         */
        function filterLiveItems(items) {
            var liveItems = [];

            for (var i = 0; i < items.length; i++) {
                try {
                    if (items[i].typename) liveItems.push(items[i]);
                } catch (error) {
                    /* すでに無効な参照は捨てる / drop references that are already dead */
                }
            }
            return liveItems;
        }

        /* ダイアログを開いた時点の適用対象。モーダルダイアログの表示中は選択を変えられないため、
           プレビューでも最終適用でも、常にこの集合が正しい適用対象になる。
           The targets captured when the dialog opens. The selection cannot change while a modal
           dialog is up, so this set stays the correct target for both preview and final apply. */
        var initialTargets = getSelectedEffectTargets();

        /**
         * 適用対象を解決する。プレビュー解除の Undo で選択が変わっても、
         * ダイアログを開いた時点の対象を優先する。
         * @returns {Array<PageItem>} 効果を適用する対象
         */
        function resolveEffectTargets() {
            var liveTargets = filterLiveItems(initialTargets);
            return (liveTargets.length > 0) ? liveTargets : getSelectedEffectTargets();
        }

        // =========================================
        // ライブ効果の適用 / Live effect application
        // =========================================

        /**
         * 変形後の4隅から自由変形の XML を組み立てる。
         * プレースホルダは位置ではなく名前で解決する（src0h と src0v の取り違えを防ぐ）。
         * @param {Array<Array<number>>} destCorners - 変形後の4隅 [TL, TR, BL, BR]
         * @returns {string} applyEffect に渡す LiveEffect XML
         */
        function buildFreeDistortXML(destCorners) {
            var cornerValues = {};

            for (var i = 0; i < 4; i++) {
                /* Y軸は Illustrator の座標系に合わせて符号を反転 / flip Y for the Illustrator axis */
                cornerValues["src" + i + "h"] = SOURCE_CORNERS[i][0];
                cornerValues["src" + i + "v"] = -SOURCE_CORNERS[i][1];
                cornerValues["dst" + i + "h"] = destCorners[i][0];
                cornerValues["dst" + i + "v"] = -destCorners[i][1];
            }

            var xmlTemplate = '<LiveEffect name="Adobe Free Distort"><Dict data="' +
                'R src0h {src0h} R src0v {src0v} R src1h {src1h} R src1v {src1v} ' +
                'R src2h {src2h} R src2v {src2v} R src3h {src3h} R src3v {src3v} ' +
                'R dst0h {dst0h} R dst0v {dst0v} R dst1h {dst1h} R dst1v {dst1v} ' +
                'R dst2h {dst2h} R dst2v {dst2v} R dst3h {dst3h} R dst3v {dst3v} ' +
                '"/></LiveEffect>';

            return fillPlaceholders(xmlTemplate, cornerValues);
        }

        /**
         * 各オブジェクトへ効果を適用する。例外で中断しても、呼び出し側は
         * onApplied で受け取った件数から何件まで適用済みかを把握できる。
         * @param {Array<PageItem>} targetItems - 適用対象
         * @param {Array<Array<number>>} destCorners - 変形後の4隅 [TL, TR, BL, BR]
         * @param {function(number):void} [onApplied] - 1件適用するたびに、それまでの適用件数で呼ばれる
         * @returns {void}
         */
        function applyFreeDistort(targetItems, destCorners, onApplied) {
            var effectXML = buildFreeDistortXML(destCorners);

            for (var i = 0; i < targetItems.length; i++) {
                targetItems[i].applyEffect(effectXML);
                if (onApplied) onApplied(i + 1);
            }
        }

        /**
         * 例外の内容をユーザーに通知する。
         * @param {Error} error - 捕捉した例外
         * @returns {void}
         */
        function showErrorAlert(error) {
            alert(SCRIPT_NAME + ": " + (error && error.message ? error.message : String(error)));
        }

        // =========================================
        // Undo境界の目印 / Undo boundary marker
        // =========================================

        /* applyEffect がオブジェクトごとに別々のアンドゥステップになるか、まとめて1ステップに
           なるかは環境によって変わる。件数ぶん undo すると戻りすぎ、1回だけだと戻し足りない。
           そこで適用前に空レイヤーを1枚足しておき、それが消えるまで undo する。
           上限は「適用件数＋目印レイヤー自身のステップ数」で、目印が消えた時点で必ず止まるため、
           ユーザー自身の操作まで巻き戻すことは決してない。
           Whether applyEffect produces one undo step per object or a single coalesced step
           varies by environment: undoing once can leave effects behind, and undoing once per
           object can eat the user's own edits. So add an empty layer before applying and undo
           until it disappears. The cap covers the applications plus the marker's own steps, and
           the loop stops the moment the marker is gone - never further back. */
        var UNDO_MARK_LAYER_NAME = "__" + SCRIPT_NAME + "_undo_mark__";

        /* 目印レイヤー自体が消費しうるアンドゥステップ数（追加とリネームが分かれる場合がある）
           Undo steps the marker layer itself can take: adding and renaming may be separate */
        var UNDO_MARK_EXTRA_STEPS = 2;

        var undoMarkBaseLayerCount = -1;
        var undoMarkLayer = null;

        /**
         * 目印レイヤーを追加する。
         * @returns {boolean} 追加できたら true、できなければ false
         */
        function beginUndoMark() {
            try {
                undoMarkBaseLayerCount = doc.layers.length;
                undoMarkLayer = doc.layers.add();
                undoMarkLayer.name = UNDO_MARK_LAYER_NAME;
                return true;
            } catch (error) {
                undoMarkBaseLayerCount = -1;
                undoMarkLayer = null;
                return false;
            }
        }

        /**
         * 目印がまだ残っているかを判定する。名前ではなくレイヤー数で見るのは、
         * レイヤー追加とリネームが別のアンドゥステップに分かれる場合があるため。
         * @returns {boolean} 目印が残っていれば true
         */
        function hasUndoMark() {
            return undoMarkBaseLayerCount >= 0 && doc.layers.length > undoMarkBaseLayerCount;
        }

        /**
         * 目印を直接取り除く。undo で消えなかった場合の後始末に使う。
         * 追加時に控えた参照で消すので、リネームが巻き戻っていても取り違えない。
         * @returns {void}
         */
        function removeUndoMark() {
            try {
                if (undoMarkLayer) undoMarkLayer.remove();
                else doc.layers.getByName(UNDO_MARK_LAYER_NAME).remove();
            } catch (error) {
                /* すでに無ければ何もしない / nothing to do when it is already gone */
            }
            undoMarkBaseLayerCount = -1;
            undoMarkLayer = null;
        }

        /**
         * 目印が消えるまで undo する。
         * @param {number} appliedCount - 目印を置いたあとに適用した件数
         * @returns {void}
         */
        function undoToMark(appliedCount) {
            var stepLimit = appliedCount + UNDO_MARK_EXTRA_STEPS;
            var undoneSteps = 0;

            while (undoneSteps < stepLimit && hasUndoMark()) {
                try {
                    app.undo();
                } catch (error) {
                    /* アンドゥ履歴が尽きた、または拒否された。目印は下で直接消す
                       The undo stack is exhausted or refused; the marker is removed below */
                    break;
                }
                undoneSteps++;
            }
            if (hasUndoMark()) removeUndoMark();
            undoMarkBaseLayerCount = -1;
            undoMarkLayer = null;
        }

        // =========================================
        // 変形量の計算 / Amount calculation
        // =========================================

        /**
         * 強度キーから倍率を取得する。
         * @param {string} strengthKey - "weak" / "normal" / "strong"
         * @returns {number} 強度倍率（未知のキーは標準扱い）
         */
        function getStrengthFactor(strengthKey) {
            var factor = STRENGTH_FACTORS[strengthKey];
            return (factor == null) ? STRENGTH_FACTORS.normal : factor;
        }

        /**
         * スライダーの変形量と強度から、実際に座標へ渡す変形量を求める。
         * @param {number} amountRatio - 変形量（0.00〜0.49）
         * @param {string} strengthKey - 強度キー
         * @returns {{trapezoidAmount: number, shearAmount: number}} 台形・シアーそれぞれの変形量
         */
        function getAppliedAmounts(amountRatio, strengthKey) {
            var trapezoidAmount = amountRatio * getStrengthFactor(strengthKey);

            return {
                trapezoidAmount: trapezoidAmount,
                shearAmount: trapezoidAmount * DISTORT_CONFIG.shearBaseFactor
            };
        }

        /**
         * プリセットから変形後の4隅座標を生成する。固定変形では変形量・強度を使わない。
         * @param {Object} preset - DISTORT_PRESETS の要素
         * @param {number} amountRatio - 変形量（0.00〜0.49）
         * @param {string} strengthKey - 強度キー
         * @returns {Array<Array<number>>} 変形後の4隅 [TL, TR, BL, BR]
         */
        function makeDestCorners(preset, amountRatio, strengthKey) {
            if (!preset.adjustable) return preset.makeCorners();

            var amounts = getAppliedAmounts(amountRatio, strengthKey);
            return preset.makeCorners(amounts.trapezoidAmount, amounts.shearAmount);
        }

        /**
         * 数値表示に出す実効変形量を求める。スライダーの値に強度倍率が掛かり、
         * 平行四辺形ではさらにシアーの基準倍率が掛かるため、表示と実際がずれないようにする。
         * @param {Object} preset - DISTORT_PRESETS の要素
         * @param {number} amountRatio - 変形量（0.00〜0.49）
         * @param {string} strengthKey - 強度キー
         * @returns {number} 実際に座標へ渡る変形量（固定変形では 0）
         */
        function getEffectiveAmountRatio(preset, amountRatio, strengthKey) {
            if (!preset.adjustable) return 0;

            var amounts = getAppliedAmounts(amountRatio, strengthKey);
            return (preset.groupKey === SHEAR_GROUP_KEY)
                ? amounts.shearAmount
                : amounts.trapezoidAmount;
        }

        // =========================================
        // 選択の判定 / Selection inspection
        // =========================================

        /**
         * オブジェクト内にテキストが含まれるかを再帰的に判定する（グループの入れ子に対応）。
         * @param {PageItem} item - 検査するオブジェクト
         * @returns {boolean} テキストを含んでいれば true
         */
        function itemContainsText(item) {
            if (!item || !item.typename) return false;
            if (item.typename === "TextFrame") return true;
            if (!item.pageItems) return false;

            /* pageItems へのアクセスはDOM越しで重いため、参照と件数を控えてから回す
               Touching pageItems crosses the DOM bridge, so cache the collection and its length */
            var childItems = item.pageItems;
            var childCount = childItems.length;

            for (var i = 0; i < childCount; i++) {
                if (itemContainsText(childItems[i])) return true;
            }
            return false;
        }

        /**
         * 選択全体にテキストが含まれるかを判定する。
         * @param {Array<PageItem>} selectedItems - 選択オブジェクト
         * @returns {boolean} いずれかにテキストが含まれていれば true
         */
        function selectionContainsText(selectedItems) {
            if (!selectedItems || selectedItems.length == null) return false;

            for (var i = 0; i < selectedItems.length; i++) {
                if (itemContainsText(selectedItems[i])) return true;
            }
            return false;
        }

        // =========================================
        // アイコンの描画 / Icon drawing
        // =========================================

        /**
         * [TL, TR, BL, BR] を外周をたどる順に並べ替える。
         * @param {Array<Array<number>>} corners - 4隅 [TL, TR, BL, BR]
         * @returns {Array<Array<number>>} 外周順に並んだ4点
         */
        function toPerimeterOrder(corners) {
            return [corners[0], corners[1], corners[3], corners[2]];
        }

        /**
         * 点群を囲む矩形を求める。
         * @param {Array<Array<number>>} points - 点の配列
         * @returns {{minX: number, minY: number, maxX: number, maxY: number}} 囲み矩形
         */
        function getPointsBounds(points) {
            var bounds = { minX: points[0][0], minY: points[0][1],
                           maxX: points[0][0], maxY: points[0][1] };

            for (var i = 1; i < points.length; i++) {
                bounds.minX = Math.min(bounds.minX, points[i][0]);
                bounds.minY = Math.min(bounds.minY, points[i][1]);
                bounds.maxX = Math.max(bounds.maxX, points[i][0]);
                bounds.maxY = Math.max(bounds.maxY, points[i][1]);
            }
            return bounds;
        }

        /**
         * 点群を (0.5, 0.5) を中心に縮小する。図形がタイル枠に触れないようにするため。
         * @param {Array<Array<number>>} points - 点の配列
         * @param {number} factor - 縮小率
         * @returns {Array<Array<number>>} 縮小後の点の配列
         */
        function shrinkAboutCenter(points, factor) {
            var shrunkPoints = [];

            for (var i = 0; i < points.length; i++) {
                shrunkPoints.push([
                    0.5 + (points[i][0] - 0.5) * factor,
                    0.5 + (points[i][1] - 0.5) * factor
                ]);
            }
            return shrunkPoints;
        }

        /**
         * 多角形の面積を靴ひも公式で求める。対角線プリセットのように面積が0のとき、
         * 塗りではなく線で描くべきかを判定するために使う。
         * @param {Array<Array<number>>} points - 外周順に並んだ点の配列
         * @returns {number} 面積
         */
        function getPolygonArea(points) {
            var doubledArea = 0;

            for (var i = 0; i < points.length; i++) {
                var nextPoint = points[(i + 1) % points.length];
                doubledArea += points[i][0] * nextPoint[1] - nextPoint[0] * points[i][1];
            }
            return Math.abs(doubledArea) / 2;
        }

        /**
         * 変形前の正方形と変形後の図形の両方が収まるよう、描画領域に合わせて
         * 拡大率と位置を決め、正規化座標をピクセル座標へ移す関数を返す。
         * 平行四辺形や台形は 0〜1 の外へはみ出すため、この収め込みが必要になる。
         * @param {Array<Array<number>>} points - 収めたい点すべて（変形前後の合成）
         * @param {number} areaLeft - 描画領域の左端（px）
         * @param {number} areaTop - 描画領域の上端（px）
         * @param {number} areaSide - 描画領域の一辺（px）
         * @returns {function(Array<number>):Array<number>} 正規化座標をピクセル座標に変換する関数
         */
        function makeIconPointMapper(points, areaLeft, areaTop, areaSide) {
            var bounds = getPointsBounds(points);
            var width = bounds.maxX - bounds.minX;
            var height = bounds.maxY - bounds.minY;
            var scale = areaSide / Math.max(width, height);
            var offsetX = areaLeft + (areaSide - width * scale) / 2 - bounds.minX * scale;
            var offsetY = areaTop + (areaSide - height * scale) / 2 - bounds.minY * scale;

            return function (point) {
                return [offsetX + point[0] * scale, offsetY + point[1] * scale];
            };
        }

        /**
         * 多角形のパスを引く。
         * @param {Object} graphics - ScriptUI の ScriptUIGraphics
         * @param {Array<Array<number>>} points - 外周順に並んだ点の配列
         * @param {function(Array<number>):Array<number>} toPixel - 座標変換関数
         * @returns {void}
         */
        function tracePolygon(graphics, points, toPixel) {
            graphics.newPath();

            for (var i = 0; i < points.length; i++) {
                var pixelPoint = toPixel(points[i]);
                if (i === 0) graphics.moveTo(pixelPoint[0], pixelPoint[1]);
                else graphics.lineTo(pixelPoint[0], pixelPoint[1]);
            }
            graphics.closePath();
        }

        /**
         * アイコンに描く変形後の形を求める。
         * @param {Object} preset - DISTORT_PRESETS の要素
         * @returns {Array<Array<number>>} 外周順・縮小済みの点の配列
         */
        function getPresetIconShape(preset) {
            var corners = preset.adjustable
                ? preset.makeCorners(ICON_SAMPLE_TRAPEZOID_AMOUNT, ICON_SAMPLE_SHEAR_AMOUNT)
                : preset.makeCorners();

            return shrinkAboutCenter(toPerimeterOrder(corners), ICON_SHAPE_SHRINK);
        }

        /**
         * 変形前の領域をタイルとして塗り、枠線を回す。
         * @param {Object} graphics - ScriptUI の ScriptUIGraphics
         * @param {Array<Array<number>>} tilePoints - タイルの4点（外周順）
         * @param {function(Array<number>):Array<number>} toPixel - 座標変換関数
         * @returns {void}
         */
        function drawIconTile(graphics, tilePoints, toPixel) {
            tracePolygon(graphics, tilePoints, toPixel);
            graphics.fillPath(graphics.newBrush(
                graphics.BrushType.SOLID_COLOR, ICON_COLORS.tileFill));
            graphics.strokePath(graphics.newPen(
                graphics.PenType.SOLID_COLOR, ICON_COLORS.tileBorder, ICON_TILE_LINE_WIDTH));
        }

        /**
         * 変形後の形を描く。面積を持たない対角線プリセットだけは太い線で描く。
         * @param {Object} graphics - ScriptUI の ScriptUIGraphics
         * @param {Array<Array<number>>} shapePoints - 変形後の4点（外周順）
         * @param {function(Array<number>):Array<number>} toPixel - 座標変換関数
         * @returns {void}
         */
        function drawIconShape(graphics, shapePoints, toPixel) {
            tracePolygon(graphics, shapePoints, toPixel);

            if (getPolygonArea(shapePoints) < DEGENERATE_AREA_THRESHOLD) {
                /* 対角線は面積が0で塗りが出ないため、太い線として描く
                   The diagonals have no area to fill, so draw them as a thick stroke */
                graphics.strokePath(graphics.newPen(
                    graphics.PenType.SOLID_COLOR, ICON_COLORS.distorted, ICON_DIAGONAL_LINE_WIDTH));
                return;
            }

            /* 塗りだけだと輪郭が粗く出るため、同色の細線を重ねる
               Fill alone leaves ragged edges, so overlay a thin outline in the same color */
            graphics.fillPath(graphics.newBrush(
                graphics.BrushType.SOLID_COLOR, ICON_COLORS.distorted));
            graphics.strokePath(graphics.newPen(
                graphics.PenType.SOLID_COLOR, ICON_COLORS.distorted, ICON_SHAPE_LINE_WIDTH));
        }

        /**
         * 選択中であることを示す枠を、ボタン全体に回す。
         * @param {Object} graphics - ScriptUI の ScriptUIGraphics
         * @param {number} buttonWidth - ボタンの幅（px）
         * @param {number} buttonHeight - ボタンの高さ（px）
         * @returns {void}
         */
        function drawIconSelectionFrame(graphics, buttonWidth, buttonHeight) {
            graphics.newPath();
            graphics.rectPath(1, 1, buttonWidth - 2, buttonHeight - 2);
            graphics.strokePath(graphics.newPen(
                graphics.PenType.SOLID_COLOR, ICON_COLORS.selected, ICON_SELECTED_LINE_WIDTH));
        }

        /**
         * アイコンボタンを描画する。変形後の4隅を makeCorners から直接得るため、
         * アイコンの形と実際に適用される効果の向きは必ず一致する（大きさは誇張。
         * ICON_SAMPLE_TRAPEZOID_AMOUNT のコメントを参照）。
         * this は描画対象のボタン。
         * @returns {void}
         */
        function drawPresetIcon() {
            var graphics = this.graphics;

            var areaSide = Math.min(this.size.width, this.size.height) - ICON_BUTTON_PADDING * 2;
            var areaLeft = (this.size.width - areaSide) / 2;
            var areaTop = (this.size.height - areaSide) / 2;

            var sourceSquare = toPerimeterOrder(SOURCE_CORNERS);
            var distortedShape = getPresetIconShape(this.presetRef);
            var toPixel = makeIconPointMapper(
                sourceSquare.concat(distortedShape), areaLeft, areaTop, areaSide);

            drawIconTile(graphics, sourceSquare, toPixel);
            drawIconShape(graphics, distortedShape, toPixel);

            if (this.value) drawIconSelectionFrame(graphics, this.size.width, this.size.height);
        }

        // =========================================
        // ダイアログの構築 / Dialog construction
        // =========================================

        /**
         * プリセット1つぶんのアイコンボタンを追加する。
         * @param {Group} parentGroup - 追加先のグループ
         * @param {Object} preset - DISTORT_PRESETS の要素
         * @returns {Button} 追加したアイコンボタン
         */
        function addPresetIconButton(parentGroup, preset) {
            var iconButton = parentGroup.add("iconbutton", undefined, undefined,
                { style: "toolbutton", toggle: true });

            iconButton.preferredSize = [ICON_BUTTON_SIZE, ICON_BUTTON_SIZE];
            iconButton.helpTip = preset.tipText;
            iconButton.presetRef = preset;
            iconButton.onDraw = drawPresetIcon;

            return iconButton;
        }

        /**
         * 指定グループのプリセットを、専用パネル内のアイコングリッドとして並べる。
         * @param {Group} parentGroup - 追加先のグループ
         * @param {string} groupKey - DISTORT_PRESETS の groupKey（LABELS.panel のキーも兼ねる）
         * @param {Array<Button>} presetControls - 添字を DISTORT_PRESETS に合わせて埋める配列
         * @param {number} [columns] - 折り返す列数（省略時は ICON_GRID_COLUMNS）
         * @returns {Panel} 追加したパネル
         */
        function addPresetIconPanel(parentGroup, groupKey, presetControls, columns) {
            var groupPanel = parentGroup.add("panel", undefined, getLabel("panel." + groupKey));
            setupPanel(groupPanel);
            groupPanel.alignChildren = ["center", "top"];
            groupPanel.spacing = ICON_GRID_SPACING;

            var iconsPerRow = (typeof columns === "number") ? columns : ICON_GRID_COLUMNS;
            var currentRow = null;
            var iconsInRow = 0;

            for (var i = 0; i < DISTORT_PRESETS.length; i++) {
                if (DISTORT_PRESETS[i].groupKey !== groupKey) continue;

                if (!currentRow || iconsInRow >= iconsPerRow) {
                    currentRow = groupPanel.add("group");
                    setupRow(currentRow, "center", ICON_GRID_SPACING);
                    iconsInRow = 0;
                }

                presetControls[i] = addPresetIconButton(currentRow, DISTORT_PRESETS[i]);
                iconsInRow++;
            }
            return groupPanel;
        }

        /**
         * ［変形の調整］パネルに、変形量スライダーと強度ラジオを組み立てる。
         * @param {Window} parentWindow - 追加先のダイアログ
         * @returns {Object} 調整パネルと各コントロールの参照
         */
        function buildAdjustPanel(parentWindow) {
            var adjustPanel = parentWindow.add("panel", undefined, getLabel("panel.adjust"));
            setupPanel(adjustPanel);

            var amountPanel = adjustPanel.add("panel", undefined, getLabel("panel.amount"));
            setupPanel(amountPanel);

            var amountSlider = amountPanel.add("slider", undefined,
                DISTORT_CONFIG.amountDefaultPercent,
                DISTORT_CONFIG.amountMinPercent,
                DISTORT_CONFIG.amountMaxPercent);
            /* 範囲はスライダーの設定から起こすので、設定を変えても説明がずれない
               The range comes from the slider's own settings, so the tip cannot drift */
            amountSlider.helpTip = fillPlaceholders(getLabel("tooltip.amount"), {
                min: DISTORT_CONFIG.amountMinPercent,
                max: DISTORT_CONFIG.amountMaxPercent
            });
            amountSlider.preferredSize.width = AMOUNT_SLIDER_WIDTH;

            /* 空文字で作ると幅が確保されず、あとから入れた文字が切れることがある。
               最長の表示でいったん作り、表示前に updateAmountReadout() が書き換える。
               An empty static text reserves no width, so build it at its widest and let
               updateAmountReadout() replace the text before the dialog is shown. */
            var amountReadout = amountPanel.add("statictext", undefined,
                fillPlaceholders(getLabel("readout.amountWithEffective"),
                    { amount: 100, effective: 100 }));
            amountReadout.alignment = ["center", "center"];

            var strengthPanel = adjustPanel.add("panel", undefined, getLabel("panel.strength"));
            setupPanel(strengthPanel);

            /* ラジオを内側のグループにまとめ、パネルの左右中央に置く
               Group the radios so the row can be centered inside the panel */
            var strengthRadioRow = strengthPanel.add("group");
            setupRow(strengthRadioRow, "center");

            return {
                adjustPanel: adjustPanel,
                amountSlider: amountSlider,
                amountReadout: amountReadout,
                weakRadio: addStrengthRadio(strengthRadioRow, "weak"),
                normalRadio: addStrengthRadio(strengthRadioRow, "normal"),
                strongRadio: addStrengthRadio(strengthRadioRow, "strong")
            };
        }

        /**
         * 強度ラジオを1つ追加する。ツールチップの倍率は STRENGTH_FACTORS から起こす。
         * @param {Group} parentGroup - 追加先のグループ
         * @param {string} strengthKey - "weak" / "normal" / "strong"
         * @returns {RadioButton} 追加したラジオボタン
         */
        function addStrengthRadio(parentGroup, strengthKey) {
            var radio = parentGroup.add("radiobutton", undefined, getLabel("strength." + strengthKey));

            radio.helpTip = fillPlaceholders(getLabel("tooltip.strength"), {
                percent: Math.round(getStrengthFactor(strengthKey) * 100)
            });
            return radio;
        }

        /**
         * ボタンエリアを左右分割で組み立てる。
         * 左：プレビュー、中央：伸縮スペーサー、右：キャンセル / OK。
         * @param {Window} parentWindow - 追加先のダイアログ
         * @returns {Checkbox} プレビューのチェックボックス
         */
        function buildFooterRow(parentWindow) {
            /* メイングループ（横並び） / Main group (horizontal layout) */
            var btnRowGroup = parentWindow.add("group");
            btnRowGroup.orientation = "row";
            btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
            btnRowGroup.alignment = ["fill", "bottom"];

            /* 左側グループ / Left-side button group
               プレビューは押しっぱなしの切り替えなので、ボタンではなくチェックボックスにしている
               Preview is a sticky toggle, so it stays a checkbox rather than a button */
            var btnLeftGroup = btnRowGroup.add("group");
            btnLeftGroup.alignChildren = ["left", "center"];
            var previewCheckbox = btnLeftGroup.add("checkbox", undefined, getLabel("checkbox.preview"));
            previewCheckbox.helpTip = getLabel("tooltip.preview");

            /* スペーサー（伸縮） / Spacer (stretchable) */
            var spacer = btnRowGroup.add("group");
            spacer.alignment = ["fill", "fill"];
            spacer.minimumSize.width = 0;

            /* 右側グループ / Right-side button group */
            var btnRightGroup = btnRowGroup.add("group");
            btnRightGroup.alignChildren = ["right", "center"];
            btnRightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
            btnRightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

            return previewCheckbox;
        }

        /**
         * ダイアログ本体を組み立てて各コントロールを返す。
         * @param {boolean} preferWeakStrength - 強度の初期値を［弱］にするか
         * @returns {Object} ダイアログと各コントロールの参照
         */
        function buildDialog(preferWeakStrength) {
            var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
            setupWindow(dialog);

            var presetColumnsGroup = dialog.add("group");
            setupRow(presetColumnsGroup, "fill", COLUMN_SPACING);
            presetColumnsGroup.alignChildren = ["fill", "top"];

            /* 添字は DISTORT_PRESETS と対応させる / indices match DISTORT_PRESETS */
            var presetControls = [];

            /* 名前では変形の向きが伝わらないため、すべてアイコンで形を見せる
               Names cannot convey the direction of a distortion, so every preset is an icon */

            /* --- 1カラム目：台形（1列×4行） / Column 1: trapezoid, one column of four --- */
            addPresetIconPanel(presetColumnsGroup, TRAPEZOID_GROUP_KEY, presetControls,
                TRAPEZOID_ICON_COLUMNS);

            /* --- 2カラム目：平行四辺形（2列×4行、上2行が左右・下2行が上下）
                   Column 2: parallelogram, 2 by 4 with horizontal on top and vertical below --- */
            addPresetIconPanel(presetColumnsGroup, SHEAR_GROUP_KEY, presetControls);

            /* --- 3カラム目：三角形と対角線を縦に積む
                   Column 3: the triangle and diagonal panels stacked --- */
            var fixedPresetColumn = presetColumnsGroup.add("group");
            fixedPresetColumn.orientation = "column";
            fixedPresetColumn.alignChildren = ["fill", "top"];
            fixedPresetColumn.spacing = PANEL_SPACING;

            for (var i = 0; i < FIXED_GROUP_KEYS.length; i++) {
                addPresetIconPanel(fixedPresetColumn, FIXED_GROUP_KEYS[i], presetControls);
            }

            /* --- 変形の調整：プリセットの3カラムの下に、全幅で置く
                   Adjust: placed full width below the preset columns --- */
            var amountControls = buildAdjustPanel(dialog);
            var previewCheckbox = buildFooterRow(dialog);

            /* --- 初期状態 / Initial state --- */
            presetControls[DEFAULT_PRESET_INDEX].value = true;
            previewCheckbox.value = false;
            /* テキストを含む選択では、崩れを抑えるため［弱］を初期値にする
               Default to Weak when text is selected, to keep the shapes readable */
            amountControls.weakRadio.value = !!preferWeakStrength;
            amountControls.normalRadio.value = !preferWeakStrength;
            amountControls.strongRadio.value = false;

            return {
                dialog: dialog,
                presetControls: presetControls,
                amountControls: amountControls,
                previewCheckbox: previewCheckbox
            };
        }

        // =========================================
        // ダイアログの制御 / Dialog interaction
        // =========================================

        /**
         * ダイアログの現在値を読み出す関数群を作る。
         * @param {Object} dialogUI - buildDialog の戻り値
         * @returns {Object} 現在値を返す関数群
         */
        function createSettingsReaders(dialogUI) {
            var presetControls = dialogUI.presetControls;
            var amountControls = dialogUI.amountControls;

            /**
             * 選択中のプリセットを返す。
             * @returns {Object} DISTORT_PRESETS の要素（未選択時は既定のプリセット）
             */
            function getSelectedPreset() {
                for (var i = 0; i < presetControls.length; i++) {
                    if (presetControls[i].value) return DISTORT_PRESETS[i];
                }
                return DISTORT_PRESETS[DEFAULT_PRESET_INDEX];
            }

            /**
             * スライダー値を整数（%）に丸めて返す。表示・適用・差分判定で同じ値を使うため。
             * @returns {number} 変形量（%）
             */
            function getAmountPercent() {
                return Math.round(amountControls.amountSlider.value);
            }

            /**
             * 実際の変形量を返す。
             * @returns {number} 変形量（0.00〜0.49）
             */
            function getAmountRatio() {
                return getAmountPercent() / 100;
            }

            /**
             * 選択中の強度キーを返す。
             * @returns {string} "weak" / "normal" / "strong"
             */
            function getStrengthKey() {
                if (amountControls.weakRadio.value) return "weak";
                if (amountControls.strongRadio.value) return "strong";
                return "normal";
            }

            return {
                getSelectedPreset: getSelectedPreset,
                getAmountPercent: getAmountPercent,
                getAmountRatio: getAmountRatio,
                getStrengthKey: getStrengthKey
            };
        }

        /**
         * プレビューの適用と巻き戻しを受け持つコントローラーを作る。
         * @param {Object} dialogUI - buildDialog の戻り値
         * @param {Object} readers - createSettingsReaders の戻り値
         * @returns {{clear: function():void, update: function(boolean):void}} プレビュー操作
         */
        function createPreviewController(dialogUI, readers) {
            var previewCheckbox = dialogUI.previewCheckbox;

            /* 取り消すべきプレビューが何件ぶん残っているか。undoToMark の上限に使う。
               How many applications of the pending preview remain; used as the undo cap. */
            var previewAppliedCount = 0;
            var lastPreviewUpdateTime = 0;
            var lastPreviewSignature = null;

            /**
             * 同一条件の再適用を避けるための署名を作る。チェックボックスの状態は
             * updatePreview 側で先に弾いているため、署名には含めない。
             * @returns {string} プリセット・変形量・強度を連結した署名
             */
            function buildPreviewSignature() {
                return [
                    readers.getSelectedPreset().presetKey,
                    readers.getAmountPercent(),
                    readers.getStrengthKey()
                ].join("|");
            }

            /**
             * プレビューを取り消す。目印レイヤーが消えるまで undo する。
             * @returns {void}
             */
            function clearPreview() {
                if (previewAppliedCount > 0 || hasUndoMark()) {
                    undoToMark(previewAppliedCount);
                    app.redraw();
                }
                previewAppliedCount = 0;
                lastPreviewSignature = null;
            }

            /**
             * プレビューを適用する。
             * @param {boolean} ignoreThrottle - true なら更新の間引きを無視する
             * @returns {void}
             */
            function updatePreview(ignoreThrottle) {
                if (!previewCheckbox.value) return;

                var signature = buildPreviewSignature();
                if (signature === lastPreviewSignature) return;

                if (!ignoreThrottle &&
                    (new Date().getTime() - lastPreviewUpdateTime) < DISTORT_CONFIG.previewThrottleMs) {
                    return;
                }

                clearPreview();

                /* 直前の undo で選択が変わっていても、ダイアログを開いた時点の対象に戻す
                   The undo just above can change the selection, so resolve the targets again */
                var targetItems = resolveEffectTargets();
                if (targetItems.length === 0) return;

                /* 目印を置けないと安全に巻き戻せないため、プレビュー自体を諦める
                   Without a marker the preview cannot be reverted safely, so skip it entirely */
                if (!beginUndoMark()) {
                    previewCheckbox.value = false;
                    return;
                }

                var destCorners = makeDestCorners(
                    readers.getSelectedPreset(), readers.getAmountRatio(), readers.getStrengthKey());

                try {
                    applyFreeDistort(targetItems, destCorners,
                        function (appliedCount) { previewAppliedCount = appliedCount; });
                } catch (error) {
                    /* 適用済みのぶんだけ巻き戻す / roll back exactly what was applied */
                    clearPreview();
                    previewCheckbox.value = false;
                    showErrorAlert(error);
                    return;
                }

                app.redraw();
                /* 描画が終わってから計り始める。失敗して何も描いていない回は
                   間引きの起点にしない。
                   Start the window once the redraw is done, so an attempt that drew
                   nothing never consumes it. */
                lastPreviewUpdateTime = new Date().getTime();
                lastPreviewSignature = signature;
            }

            return { clear: clearPreview, update: updatePreview };
        }

        /**
         * ダイアログのイベントを結び、初期表示を整える。
         * @param {Object} dialogUI - buildDialog の戻り値
         * @param {Object} readers - createSettingsReaders の戻り値
         * @param {Object} preview - createPreviewController の戻り値
         * @returns {void}
         */
        function bindDialogEvents(dialogUI, readers, preview) {
            var presetControls = dialogUI.presetControls;
            var amountControls = dialogUI.amountControls;
            var previewCheckbox = dialogUI.previewCheckbox;

            /**
             * 変形量の数値表示を更新する。調整できるプリセットでは、［強度］の倍率を
             * 掛けたあとの実効値を併記する。
             * @returns {void}
             */
            function updateAmountReadout() {
                var preset = readers.getSelectedPreset();
                var amountPercent = readers.getAmountPercent();

                if (!preset.adjustable) {
                    amountControls.amountReadout.text =
                        fillPlaceholders(getLabel("readout.amount"), { amount: amountPercent });
                    return;
                }

                var effectiveRatio = getEffectiveAmountRatio(
                    preset, readers.getAmountRatio(), readers.getStrengthKey());

                amountControls.amountReadout.text =
                    fillPlaceholders(getLabel("readout.amountWithEffective"), {
                        amount: amountPercent,
                        effective: Math.round(effectiveRatio * 100)
                    });
            }

            /**
             * 固定変形では［変形の調整］をディム表示にする（子コントロールもまとめて無効化）。
             * @returns {void}
             */
            function updateAdjustPanelEnabled() {
                amountControls.adjustPanel.enabled = readers.getSelectedPreset().adjustable;
            }

            /**
             * 押されたプリセットだけを選択状態にする。プリセットは複数パネルに分かれており、
             * さらにラジオとトグルボタンが混在するため、ScriptUI の排他選択は使えない。
             * @param {Button} clickedControl - 押されたアイコンボタン
             * @returns {void}
             */
            function selectPresetControl(clickedControl) {
                for (var k = 0; k < presetControls.length; k++) {
                    presetControls[k].value = (presetControls[k] === clickedControl);
                }
                /* アイコンボタンは onDraw で選択枠を描くので、再描画させる
                   The icon buttons draw their selection frame in onDraw, so force a repaint */
                if (dialogUI.dialog.update) dialogUI.dialog.update();
            }

            amountControls.amountSlider.onChanging = function () {
                updateAmountReadout();
                preview.update(false);
            };

            amountControls.amountSlider.onChange = function () {
                updateAmountReadout();
                preview.update(true);
            };

            previewCheckbox.onClick = function () {
                if (previewCheckbox.value) preview.update(true);
                else preview.clear();
            };

            /* 強度を変えると実効値が変わるので、数値表示も更新する
               The effective amount changes with the strength, so refresh the readout too */
            amountControls.weakRadio.onClick =
            amountControls.normalRadio.onClick =
            amountControls.strongRadio.onClick = function () {
                updateAmountReadout();
                preview.update(true);
            };

            for (var i = 0; i < presetControls.length; i++) {
                presetControls[i].onClick = function () {
                    /* トグルボタンは再クリックで解除されてしまうため、選択状態を維持する
                       A toggle button would clear itself on a second click, so keep it selected */
                    selectPresetControl(this);
                    updateAdjustPanelEnabled();
                    updateAmountReadout();
                    preview.update(true);
                };
            }

            updateAmountReadout();
            updateAdjustPanelEnabled();
        }

        /**
         * ダイアログを表示し、確定した設定を返す。
         * @returns {?Object} プリセット・変形量・強度（キャンセル時は null）
         */
        function showDistortDialog() {
            var dialogUI = buildDialog(selectionContainsText(doc.selection));
            var readers = createSettingsReaders(dialogUI);
            var preview = createPreviewController(dialogUI, readers);

            bindDialogEvents(dialogUI, readers, preview);

            var isAccepted = (dialogUI.dialog.show() == 1);

            /* キャンセル時も OK 時も、いったんプレビューを完全に戻してから抜ける
               Either way, revert the preview completely before leaving */
            preview.clear();
            if (!isAccepted) return null;

            return {
                preset: readers.getSelectedPreset(),
                amountRatio: readers.getAmountRatio(),
                strengthKey: readers.getStrengthKey()
            };
        }

        // =========================================
        // 実行 / Run
        // =========================================

        /* 効果を適用できる対象が無いなら、ダイアログを出す前に知らせる
           Tell the user before opening the dialog when nothing can take an effect */
        if (initialTargets.length === 0) {
            alert(getLabel("alert.noTarget"));
            return;
        }

        var distortSettings = showDistortDialog();
        if (!distortSettings) return;

        var targetItems = resolveEffectTargets();
        if (targetItems.length === 0) {
            alert(getLabel("alert.noTarget"));
            return;
        }

        var destCorners = makeDestCorners(
            distortSettings.preset,
            distortSettings.amountRatio,
            distortSettings.strengthKey);

        /* ここでは目印レイヤーを使わない。成功時に目印を消す操作がアンドゥ履歴の最後に残り、
           ユーザーが Undo したときに空レイヤーが復活してしまうため。
           代わりに、途中で失敗したら何件まで適用したかを伝える。
           No marker layer here: removing it on success would sit at the top of the undo stack,
           so the user's first undo would resurrect an empty layer. Report how far the run got
           instead, and leave the undo to the user. */
        var appliedCount = 0;

        try {
            applyFreeDistort(targetItems, destCorners,
                function (count) { appliedCount = count; });
        } catch (error) {
            showErrorAlert(error);
            if (appliedCount > 0) {
                alert(fillPlaceholders(getLabel("alert.partialApply"), { count: appliedCount }));
            }
        }

    })();

})();
