#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した複数オブジェクトを左から順に見て、隣り合うオブジェクト同士のアキの中央に記号を配置します。
記号は9種類から選べ、高さ・幅・線幅・位置をダイアログで調整できます。

詳細は README を参照してください。

### Overview

Scans the selected objects from left to right and places a mark in the middle of the gap between each adjacent pair.
Nine marks are available, with height, width, stroke weight and position set from the dialog.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "RightMarkPlacer";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-28";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-01";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RightMarkPlacer.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RightMarkPlacer.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nebac730ec187"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    /* 記号の色（CMYK）/ Mark color (CMYK) */
    var MARK_COLOR_CMYK = [0, 0, 0, 100];

    /* 高さ（％）の上限 / Maximum height percentage */
    var MAX_HEIGHT_PERCENT = 200;

    /* 線幅の下限（pt）/ Minimum stroke width in points */
    var MIN_STROKE_WIDTH_PT = 0.25;

    /* ▶ の凹みの上限（幅に対する割合）/ Maximum inset as a ratio of the width */
    var MAX_INSET_RATIO = 0.8;

    /* ▶ の角丸半径（幅と高さの小さい方に対する割合）/ Rounded-corner radius as a ratio of the smaller side */
    var TRI_CORNER_RADIUS_RATIO = 0.12;

    /* ＿\ の斜線の角度（度）/ Slash angle in degrees */
    var SLASH_ANGLE_DEFAULT = 35;
    var SLASH_ANGLE_MAX = 89;

    /* ➡ の矢じりの天地を求める、線幅に対する倍率 / Height of the solid arrowhead as a multiple of the stroke width */
    var ARROW3_HEIGHT_TO_STROKE_RATIO = 3;

    /* → ➡ の矢じりの奥行きが、幅に占める割合の上限 / Maximum arrowhead depth as a ratio of the width */
    var MAX_ARROW_HEAD_RATIO = 0.9;

    /* 幅を自動計算するときの、アキに対する割合 / Ratio of the gap used when the width is calculated automatically */
    var AUTO_WIDTH_GAP_RATIO = 0.7;
    var AUTO_WIDTH_GAP_RATIO_SMALL = 0.35;

    /* 山形の幅を自動計算するときの、高さに対する割合 / Ratio of the height used when a chevron width is calculated automatically */
    var AUTO_WIDTH_CHEVRON_HEIGHT_RATIO = 0.5;

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 実行環境の言語を判定します。
     *
     * @returns {string} 日本語環境なら "ja"、それ以外は "en"。
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var lang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            /* ＋ や × のような無方向の形状と［左右逆］があるため、向きを含めない名前にしています / Kept direction-neutral: some shapes have no direction and the mark can be mirrored */
            title: { ja: "オブジェクト間に記号を配置", en: "Place Marks Between Objects" }
        },
        panel: {
            shape: { ja: "形状", en: "Shape" },
            cap: { ja: "先端", en: "End Style" },
            options: { ja: "オプション", en: "Options" },
            adjust: { ja: "位置調整", en: "Position" }
        },
        radio: {
            arrowSlash: { ja: "＿\\", en: "─\\" },
            capNone: { ja: "なし", en: "None" },
            capRound: { ja: "丸型", en: "Round" }
        },
        checkbox: {
            mirrorHorizontal: { ja: "左右逆", en: "Mirror horizontally" },
            alignVerticalCenter: { ja: "天地を水平に", en: "Keep top and bottom edges horizontal" },
            roundCorners: { ja: "角丸", en: "Rounded corners" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        /* 入力欄の前に置くラベル。区切りの「：」まで含める / Labels placed before a field, including the trailing colon */
        label: {
            height: { ja: "高さ：", en: "Height:" },
            width: { ja: "幅：", en: "Width:" },
            gap: { ja: "間隔：", en: "Gap:" },
            inset: { ja: "凹み：", en: "Inset:" },
            stroke: { ja: "線幅：", en: "Stroke:" },
            angle: { ja: "角度：", en: "Angle:" },
            adjustX: { ja: "左右：", en: "Horizontal:" },
            adjustY: { ja: "上下：", en: "Vertical:" }
        },
        tooltip: {
            arrow3: { ja: "矢じりの天地は線幅の{0}倍になります。", en: "The arrowhead is {0}× the stroke width tall." },
            inset: { ja: "幅の{0}%が上限です", en: "Max {0}% of width" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            openDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            selectTwoObjects: { ja: "オブジェクトを2つ以上選択してください。", en: "Please select two or more objects." },
            lockedLayer: {
                ja: "作業レイヤーがロックまたは非表示です。解除してから実行してください。",
                en: "The active layer is locked or hidden. Please unlock and show it, then run again."
            },
            positiveNumber: { ja: "正の数値を入力してください。", en: "Please enter a positive value." },
            maxHeight: { ja: "{0}% 以下の値を入力してください。", en: "Please enter a value of {0}% or less." },
            noGap: {
                ja: "隣り合うオブジェクト間に作成できるアキがありません。",
                en: "There is no usable gap between adjacent objects."
            },
            strokePositive: {
                ja: "線幅は {0} {1} 以上の値を入力してください。",
                en: "Please enter a stroke width of {0} {1} or greater."
            },
            widthPositive: { ja: "幅は 0 以上の値を入力してください。", en: "Please enter a width value of 0 or greater." },
            invalidValue: { ja: "入力値を確認してください。", en: "Please check the input values." }
        },
        log: {
            measureTextBounds: { ja: "テキストの計測用アウトライン化", en: "Outline text for measurement" },
            removeMeasurementCopy: { ja: "計測用複製の削除", en: "Remove measurement copy" },
            applyRoundCorners: { ja: "角丸効果の適用", en: "Apply rounded-corners effect" },
            restoreSelection: { ja: "選択状態の復元", en: "Restore selection" },
            removePreviewItem: { ja: "プレビューの削除", en: "Remove preview item" },
            mergeSolidArrow: { ja: "➡ の合成", en: "Merge the solid arrow" },
            mirrorItem: { ja: "左右の反転", en: "Mirror the item horizontally" },
            layoutDialog: { ja: "ダイアログのレイアウト", en: "Lay out the dialog" }
        }
    };

    /**
     * ドットパスで指定したラベルを、現在の言語で取得します。
     *
     * @param {string} labelPath - ラベルのドットパス（例 "panel.shape"）。
     * @returns {string} 現在の言語のラベル。見つからない場合は英語、それもなければ labelPath。
     */
    function getLabel(labelPath) {
        var pathParts = labelPath.split(".");
        var entry = LABELS;
        for (var i = 0; i < pathParts.length; i++) {
            if (!entry) break;
            entry = entry[pathParts[i]];
        }
        if (entry && entry[lang]) return entry[lang];
        if (entry && entry.en) return entry.en;
        return labelPath;
    }

    /**
     * ラベル中の {0} {1} … を、与えた値で置き換えます。
     *
     * @param {string} labelPath - ラベルのドットパス。
     * @param {array} [values] - 差し込む値の配列。
     * @returns {string} 置き換え後の文言。
     */
    function formatLabel(labelPath, values) {
        var text = getLabel(labelPath);
        if (!values) return text;
        for (var i = 0; i < values.length; i++) {
            text = text.replace("{" + i + "}", String(values[i]));
        }
        return text;
    }

    /**
     * コンテキスト付きでエラーを $.writeln に出力します。
     *
     * @param {string} context - どの処理で起きたかを示す文言。
     * @param {object} errorObject - エラーオブジェクトまたはメッセージ。
     * @returns {void}
     */
    function logScriptError(context, errorObject) {
        $.writeln("[" + SCRIPT_NAME + " " + SCRIPT_VERSION + "] " + context + ": " + errorObject);
    }

    /**
     * 処理を実行し、失敗した場合はログを出して続行します。
     *
     * @param {function} operation - 実行する処理。
     * @param {string} context - 失敗時にログへ出す文言。
     * @returns {boolean} 成功したら true、失敗したら false。
     */
    function runSafely(operation, context) {
        try {
            operation();
            return true;
        } catch (e) {
            logScriptError(context, e);
            return false;
        }
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コードと単位ラベルの対応 / Mapping from unit code to unit label */
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
     * 単位コードから単位ラベルを求めます。
     *
     * @param {number} unitCode - 環境設定の単位コード。
     * @param {string} preferenceKey - 環境設定のキー（Q と H の判別に使用）。
     * @returns {string} 単位ラベル。
     */
    function getUnitLabel(unitCode, preferenceKey) {
        if (unitCode === 5) {
            /* 級（Q）は文字サイズ、歯（H）は送りや罫線に使う / Q is for font size, H for spacing and rules */
            var usesHaLabel = {
                "rulerType": true,
                "strokeUnits": true
            };
            return usesHaLabel[preferenceKey] ? "H" : "Q";
        }
        return UNIT_LABELS[unitCode] || "pt";
    }

    /**
     * 単位コードから pt への換算係数を求めます。
     *
     * @param {number} unitCode - 環境設定の単位コード。
     * @returns {number} 1単位あたりの pt 数。
     */
    function getPtFactorFromUnitCode(unitCode) {
        switch (unitCode) {
            case 0: return 72.0;                        // in
            case 1: return 72.0 / 25.4;                 // mm
            case 2: return 1.0;                         // pt
            case 3: return 12.0;                        // pica
            case 4: return 72.0 / 2.54;                 // cm
            case 5: return 72.0 / 25.4 * 0.25;          // Q or H
            case 6: return 1.0;                         // px
            case 7: return 72.0 * 12.0;                 // ft/in
            case 8: return 72.0 / 25.4 * 1000.0;        // m
            case 9: return 72.0 * 36.0;                 // yd
            case 10: return 72.0 * 12.0;                // ft
            default: return 1.0;
        }
    }

    /**
     * 環境設定から単位情報（コード・ラベル・換算係数）を取得します。
     *
     * @param {string} preferenceKey - 環境設定のキー（"rulerType" など）。
     * @returns {object} code / label / factor を持つオブジェクト。
     */
    function getPreferenceUnitInfo(preferenceKey) {
        var unitCode = app.preferences.getIntegerPreference(preferenceKey);
        return {
            code: unitCode,
            label: getUnitLabel(unitCode, preferenceKey),
            factor: getPtFactorFromUnitCode(unitCode)
        };
    }

    var rulerUnitInfo = getPreferenceUnitInfo("rulerType");
    var strokeUnitInfo = getPreferenceUnitInfo("strokeUnits");

    /**
     * 単位の数値を pt に換算します。
     *
     * @param {number} value - 単位付きの数値。
     * @param {object} unitInfo - 単位情報。
     * @returns {number} pt 値。
     */
    function convertValueToPt(value, unitInfo) {
        return value * unitInfo.factor;
    }

    /**
     * pt を単位の数値に換算します。
     *
     * @param {number} valuePt - pt 値。
     * @param {object} unitInfo - 単位情報。
     * @returns {number} 単位付きの数値。
     */
    function convertPtToUnitValue(valuePt, unitInfo) {
        return valuePt / unitInfo.factor;
    }

    /* 表示桁数 / Decimal places used for display */
    var DISPLAY_DECIMALS = 2;          /* 既定の桁数 / default decimal places */
    var DISPLAY_DECIMALS_MAX = 5;      /* 桁数の上限 / maximum decimal places */
    var DISPLAY_DECIMALS_PT_FACTOR = 3; /* この換算係数までは既定の桁数 / units up to this factor keep the default */

    /**
     * pt 値を小数点以下2桁へ丸めます。
     *
     * @param {number} value - 丸める値。
     * @returns {number} 丸めた値。
     */
    function roundDisplayValue(value) {
        return Math.round(value * 100) / 100;
    }

    /**
     * 単位に応じた表示桁数を求めます。
     * inch のように 1単位が大きい単位では、2桁だと pt 換算で精度が足りないため桁数を増やします。
     *
     * @param {object} unitInfo - 単位情報。
     * @returns {number} 小数点以下の桁数。
     */
    function getDisplayDecimals(unitInfo) {
        if (!unitInfo || !(unitInfo.factor > DISPLAY_DECIMALS_PT_FACTOR)) return DISPLAY_DECIMALS;
        var extraDigits = Math.ceil(Math.log(unitInfo.factor) / Math.LN10);
        return Math.min(DISPLAY_DECIMALS + extraDigits, DISPLAY_DECIMALS_MAX);
    }

    /**
     * 単位に合わせた桁数で丸めます。
     *
     * @param {number} value - 単位付きの数値。
     * @param {object} unitInfo - 単位情報。
     * @returns {number} 丸めた値。
     */
    function roundDisplayValueForUnit(value, unitInfo) {
        var scale = Math.pow(10, getDisplayDecimals(unitInfo));
        return Math.round(value * scale) / scale;
    }

    /**
     * 矢印キーで増減する量を、単位に合わせて求めます。
     * pt・px・mm・Q・H は 1単位、inch のように 1単位が大きい単位はより細かい刻みにします。
     *
     * @param {object} [unitInfo] - 単位情報。％や度など単位のない入力欄では省略します。
     * @returns {number} 増減量。
     */
    function getArrowKeyStep(unitInfo) {
        if (!unitInfo || !(unitInfo.factor > 0)) return 1;
        var exponent = Math.max(0, Math.round(Math.log(unitInfo.factor) / Math.LN10));
        return Math.pow(10, -exponent);
    }

    /**
     * pt 値を単位に換算して入力欄へ表示します。
     *
     * @param {object} editText - 対象の edittext。
     * @param {number} valuePt - pt 値。
     * @param {object} unitInfo - 単位情報。
     * @returns {void}
     */
    function setFieldFromPt(editText, valuePt, unitInfo) {
        editText.text = String(roundDisplayValueForUnit(convertPtToUnitValue(valuePt, unitInfo), unitInfo));
    }

    /**
     * 入力欄の値を pt として読み取ります。
     *
     * @param {object} editText - 対象の edittext。
     * @param {object} unitInfo - 単位情報。
     * @param {boolean} allowNegative - 負の値を許可するなら true。
     * @returns {number} pt 値。数値でない場合と、許可していない負の値の場合は NaN。
     */
    function parseFieldToPt(editText, unitInfo, allowNegative) {
        var value = parseFloat(editText.text);
        if (isNaN(value)) return NaN;
        /* 負の値を黙って 0 にせず、呼び出し元でエラーとして扱えるようにする / Report it instead of silently clamping to zero */
        if (!allowNegative && value < 0) return NaN;
        return convertValueToPt(value, unitInfo);
    }

    // =========================================
    // UIレイアウトの共通設定 / Shared UI layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
    var FIELD_ROW_SPACING = 8;               /* ラベルと入力欄の間隔 / gap between a label and its field */
    var LABEL_COLUMN_WIDTH = 60;             /* ラベル列の幅 / width of the label column */

    /**
     * ウィンドウに共通のレイアウトを適用します。
     *
     * @param {object} win - 対象の Window。
     * @param {number} [spacing] - 要素間隔。省略時は WINDOW_SPACING。
     * @returns {void}
     */
    function setupWindow(win, spacing) {
        win.orientation = "column";
        win.alignChildren = "fill";
        win.margins = WINDOW_MARGINS;
        win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルに共通のレイアウトを適用します。
     *
     * @param {object} panel - 対象の Panel。
     * @param {number} [spacing] - 要素間隔。省略時は PANEL_SPACING。
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
     * 縦並びのグループ（カラム）に共通のレイアウトを適用します。
     *
     * @param {object} group - 対象の Group。
     * @param {number} [spacing] - 要素間隔。省略時は PANEL_SPACING。
     * @returns {void}
     */
    function setupColumn(group, spacing) {
        group.orientation = "column";
        group.alignChildren = ["fill", "top"];
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * 横並びのグループ（入力行・ボタン列など）に共通のレイアウトを適用します。
     *
     * @param {object} group - 対象の Group。
     * @param {string} [alignment] - グループ自体の配置。省略時は "left"。
     * @param {number} [spacing] - 要素間隔。省略時は PANEL_SPACING。
     * @returns {void}
     */
    function setupRow(group, alignment, spacing) {
        group.orientation = "row";
        group.alignChildren = ["left", "center"];
        group.alignment = alignment || "left";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    (function () {

        if (app.documents.length === 0) {
            alert(getLabel("alert.openDocument"));
            return;
        }

        var activeDocument = app.activeDocument;

        /* ロック・非表示のレイヤーには作成できないので、先に知らせる / Nothing can be created on a locked or hidden layer */
        if (activeDocument.activeLayer.locked || !activeDocument.activeLayer.visible) {
            alert(getLabel("alert.lockedLayer"));
            return;
        }

        var selectedItems = [];
        for (var selectionIndex = 0; selectionIndex < activeDocument.selection.length; selectionIndex++) {
            selectedItems.push(activeDocument.selection[selectionIndex]);
        }

        if (selectedItems.length < 2) {
            alert(getLabel("alert.selectTwoObjects"));
            return;
        }

        // =========================================
        // 汎用ユーティリティ / Generic utilities
        // =========================================

        /**
         * ExtendScript のコレクションを通常の配列へコピーします。
         *
         * @param {object} collection - selection などのコレクション。
         * @returns {array} コピーした配列。
         */
        function collectionToArray(collection) {
            var items = [];
            if (!collection) return items;
            for (var i = 0; i < collection.length; i++) {
                items.push(collection[i]);
            }
            return items;
        }

        /**
         * アイテムを削除します（失敗しても続行）。
         *
         * @param {object} item - 対象のアイテム。null なら何もしません。
         * @param {string} context - 失敗時にログへ出す文言。
         * @returns {void}
         */
        function removeItemSafely(item, context) {
            if (!item) return;
            runSafely(function () {
                item.remove();
            }, context);
        }

        /**
         * アイテムを選択状態にします（失敗しても続行）。
         *
         * @param {object} item - 対象のアイテム。
         * @param {string} context - 失敗時にログへ出す文言。
         * @returns {void}
         */
        function selectItemSafely(item, context) {
            runSafely(function () {
                item.selected = true;
            }, context);
        }

        /**
         * 選択状態をまとめて復元します。
         *
         * @param {array} items - 選択し直すアイテムの配列。
         * @returns {void}
         */
        function restoreSelection(items) {
            activeDocument.selection = null;
            for (var i = 0; i < items.length; i++) {
                selectItemSafely(items[i], getLabel("log.restoreSelection"));
            }
        }

        // =========================================
        // 選択オブジェクトの計測 / Measuring the selection
        // =========================================

        /*
         * 同じアイテムを何度も計測しないためのキャッシュ。
         * モーダルダイアログの表示中は選択オブジェクトが変化しないため、閉じるまで保持します。
         * テキストの計測はアウトライン化を伴うので、プレビューの更新ごとに測り直すと重くなります。
         *
         * Cache so each item is measured only once.
         * The selection cannot change while the modal dialog is up, so it is kept until the dialog closes.
         * Measuring text requires outlining it, which is too slow to repeat on every preview refresh.
         */
        var measurementBoundsCache = [];

        /**
         * バウンズ配列を複製します（呼び出し側での書き換えを防ぐため）。
         *
         * @param {array} bounds - geometricBounds。
         * @returns {array} 複製したバウンズ。
         */
        function copyBounds(bounds) {
            return [bounds[0], bounds[1], bounds[2], bounds[3]];
        }

        /**
         * テキストをアウトライン化した状態のバウンズを求めます。
         * 複製をアウトライン化して計測し、計測用の複製は必ず削除します。
         *
         * @param {object} textFrame - 対象の TextFrame。
         * @returns {array} バウンズ。失敗した場合は元のバウンズ。
         */
        function measureOutlinedTextBounds(textFrame) {
            var duplicatedText = null;
            var outlinedText = null;
            try {
                duplicatedText = textFrame.duplicate();
                outlinedText = duplicatedText.createOutline();
                return copyBounds(outlinedText.geometricBounds);
            } catch (e) {
                logScriptError(getLabel("log.measureTextBounds"), e);
                return copyBounds(textFrame.geometricBounds);
            } finally {
                /* createOutline 後は複製が消費されるため、両方を安全に片付ける / The duplicate is consumed by createOutline, so clean up both defensively */
                removeItemSafely(outlinedText, getLabel("log.removeMeasurementCopy"));
                removeItemSafely(duplicatedText, getLabel("log.removeMeasurementCopy"));
            }
        }

        /**
         * 計測用のバウンズを取得します（テキストはアウトライン化した形状で計測）。
         *
         * @param {object} item - 対象のアイテム。
         * @returns {array} バウンズ。
         */
        function getItemMeasurementBounds(item) {
            for (var i = 0; i < measurementBoundsCache.length; i++) {
                if (measurementBoundsCache[i].item === item) {
                    return copyBounds(measurementBoundsCache[i].bounds);
                }
            }

            var bounds = (item.typename === "TextFrame")
                ? measureOutlinedTextBounds(item)
                : copyBounds(item.geometricBounds);

            measurementBoundsCache.push({ item: item, bounds: bounds });
            return copyBounds(bounds);
        }

        /**
         * アイテムの中心の X 座標を求めます。
         *
         * @param {object} item - 対象のアイテム。
         * @returns {number} 中心の X 座標。
         */
        function getItemCenterX(item) {
            var bounds = getItemMeasurementBounds(item);
            return (bounds[0] + bounds[2]) / 2;
        }

        /**
         * 選択オブジェクトを左から右の順に並べ替えます。
         *
         * @returns {array} 並べ替えたアイテムの配列。
         */
        function getSortedSelectionItems() {
            var items = collectionToArray(selectedItems);
            items.sort(function (leftItem, rightItem) {
                return getItemCenterX(leftItem) - getItemCenterX(rightItem);
            });
            return items;
        }

        // =========================================
        // 形状定義 / Shape definitions
        // =========================================

        /* 形状ごとの初期値と、使用するコントロールの有無 / Per-shape defaults and which controls apply */
        var SHAPE_CONFIG = {
            tri: {
                radioKey: "triangleRadio",
                enableCapPanel: false,
                forceFillOnly: true,
                enableStrokeInput: false,
                requirePositiveStroke: false,
                enableHeightInput: true,
                enableMirror: true,
                enableGap: false,
                enableInset: true,
                enableRoundCorners: true,
                enableAngle: false,
                defaultHeightPercent: 30,
                defaultGap: -1,
                defaultInsetPt: 0,
                defaultStrokePt: 0.3,
                calcDefaultWidth: function (heightPercent, minTotalHeight) {
                    return roundDisplayValue(minTotalHeight * (heightPercent / 100));
                }
            },
            arrow: {
                radioKey: "arrowRadio",
                enableCapPanel: true,
                forceFillOnly: false,
                enableStrokeInput: true,
                requirePositiveStroke: true,
                enableHeightInput: true,
                enableMirror: true,
                enableGap: false,
                enableInset: false,
                enableRoundCorners: false,
                enableAngle: false,
                defaultHeightPercent: 30,
                defaultGap: -1,
                defaultInsetPt: 0,
                defaultStrokePt: 0.6,
                calcDefaultWidth: function (heightPercent, minTotalHeight, minGapWidth) {
                    return roundDisplayValue(minGapWidth * AUTO_WIDTH_GAP_RATIO);
                }
            },
            arrow3: {
                radioKey: "arrow3Radio",
                enableCapPanel: false,
                forceFillOnly: false,
                enableStrokeInput: true,
                requirePositiveStroke: true,
                /* 天地は線幅から決まるため、高さ（％）は使いません / The height comes from the stroke width, so the percentage is unused */
                enableHeightInput: false,
                enableMirror: true,
                enableGap: false,
                enableInset: false,
                enableRoundCorners: false,
                enableAngle: false,
                defaultHeightPercent: undefined,
                defaultGap: -1,
                defaultInsetPt: 0,
                defaultStrokePt: 3.6,
                calcDefaultWidth: function (heightPercent, minTotalHeight, minGapWidth) {
                    return roundDisplayValue(minGapWidth * AUTO_WIDTH_GAP_RATIO);
                }
            },
            arrowSlash: {
                radioKey: "arrowSlashRadio",
                enableCapPanel: true,
                forceFillOnly: false,
                enableStrokeInput: true,
                requirePositiveStroke: true,
                enableHeightInput: true,
                enableMirror: true,
                enableGap: false,
                enableInset: false,
                enableRoundCorners: false,
                enableAngle: true,
                defaultHeightPercent: 50,
                defaultGap: -1,
                defaultInsetPt: 0,
                defaultStrokePt: 0.6,
                defaultAngle: SLASH_ANGLE_DEFAULT,
                calcDefaultWidth: function (heightPercent, minTotalHeight, minGapWidth) {
                    return roundDisplayValue(minGapWidth * AUTO_WIDTH_GAP_RATIO);
                }
            },
            chevron: {
                radioKey: "chevronRadio",
                enableCapPanel: true,
                forceFillOnly: false,
                enableStrokeInput: true,
                requirePositiveStroke: true,
                enableHeightInput: true,
                enableMirror: true,
                enableGap: false,
                enableInset: false,
                enableRoundCorners: false,
                enableAngle: false,
                defaultHeightPercent: 50,
                defaultGap: -1,
                defaultInsetPt: 0,
                defaultStrokePt: 0.6,
                calcDefaultWidth: function (heightPercent, minTotalHeight) {
                    return roundDisplayValue(minTotalHeight * (heightPercent / 100) * AUTO_WIDTH_CHEVRON_HEIGHT_RATIO);
                }
            },
            dash: {
                radioKey: "dashRadio",
                enableCapPanel: true,
                forceFillOnly: false,
                enableStrokeInput: true,
                requirePositiveStroke: true,
                enableHeightInput: false,
                enableMirror: false,
                enableGap: false,
                enableInset: false,
                enableRoundCorners: false,
                enableAngle: false,
                defaultHeightPercent: undefined,
                defaultGap: -1,
                defaultInsetPt: 0,
                defaultStrokePt: 0.6,
                calcDefaultWidth: function (heightPercent, minTotalHeight, minGapWidth) {
                    return roundDisplayValue(minGapWidth * AUTO_WIDTH_GAP_RATIO);
                }
            },
            plus: {
                radioKey: "plusRadio",
                enableCapPanel: true,
                forceFillOnly: false,
                enableStrokeInput: true,
                requirePositiveStroke: true,
                enableHeightInput: false,
                enableMirror: false,
                enableGap: false,
                enableInset: false,
                enableRoundCorners: false,
                enableAngle: false,
                defaultHeightPercent: undefined,
                defaultGap: -1,
                defaultInsetPt: 0,
                defaultStrokePt: 0.6,
                calcDefaultWidth: function (heightPercent, minTotalHeight, minGapWidth) {
                    return roundDisplayValue(minGapWidth * AUTO_WIDTH_GAP_RATIO_SMALL);
                }
            },
            multiply: {
                radioKey: "multiplyRadio",
                enableCapPanel: true,
                forceFillOnly: false,
                enableStrokeInput: true,
                requirePositiveStroke: true,
                enableHeightInput: false,
                enableMirror: false,
                enableGap: false,
                enableInset: false,
                enableRoundCorners: false,
                enableAngle: false,
                defaultHeightPercent: undefined,
                defaultGap: -1,
                defaultInsetPt: 0,
                defaultStrokePt: 0.6,
                calcDefaultWidth: function (heightPercent, minTotalHeight, minGapWidth) {
                    return roundDisplayValue(minGapWidth * AUTO_WIDTH_GAP_RATIO_SMALL);
                }
            },
            chevron2: {
                radioKey: "chevronDoubleRadio",
                enableCapPanel: true,
                forceFillOnly: false,
                enableStrokeInput: true,
                requirePositiveStroke: true,
                enableHeightInput: true,
                enableMirror: true,
                enableGap: true,
                enableInset: false,
                enableRoundCorners: false,
                enableAngle: false,
                defaultHeightPercent: 50,
                defaultGap: 0,
                defaultInsetPt: 0,
                defaultStrokePt: 0.6,
                /* 幅は山形1つ分。全体の幅は「幅×2＋間隔」になります / The width is per chevron; the pair spans width × 2 + gap */
                calcDefaultWidth: function (heightPercent, minTotalHeight) {
                    return roundDisplayValue(minTotalHeight * (heightPercent / 100) * AUTO_WIDTH_CHEVRON_HEIGHT_RATIO);
                }
            }
        };

        // =========================================
        // ダイアログ / Dialog
        // =========================================

        /**
         * ラベル・入力欄・単位を1行にまとめて追加します。
         *
         * @param {object} parentPanel - 追加先のパネル。
         * @param {string} labelPath - ラベルのドットパス。
         * @param {string} initialText - 入力欄の初期値。
         * @param {string} unitText - 入力欄の右に置く単位表記。
         * @returns {object} row / label / field / unitLabel を持つオブジェクト。
         */
        function addLabeledField(parentPanel, labelPath, initialText, unitText) {
            var row = parentPanel.add("group");
            setupRow(row, "left", FIELD_ROW_SPACING);

            var label = row.add("statictext", undefined, getLabel(labelPath));
            label.preferredSize = [LABEL_COLUMN_WIDTH, -1];
            label.justify = "right";

            var field = row.add("edittext", undefined, initialText);
            field.characters = 4;

            var unitLabel = row.add("statictext", undefined, unitText);

            return { row: row, label: label, field: field, unitLabel: unitLabel };
        }

        /**
         * ［形状］パネルを作成します。
         *
         * @param {object} parent - 追加先のコンテナ。
         * @returns {object} パネルと各コントロール。
         */
        function buildShapePanel(parent) {
            var panel = parent.add("panel", undefined, getLabel("panel.shape"));
            setupPanel(panel, 6);

            var triangleRadio = panel.add("radiobutton", undefined, "▶");
            var chevronRadio = panel.add("radiobutton", undefined, ">");
            var chevronDoubleRadio = panel.add("radiobutton", undefined, ">>");
            var dashRadio = panel.add("radiobutton", undefined, "─");
            var arrowRadio = panel.add("radiobutton", undefined, "→");
            var arrow3Radio = panel.add("radiobutton", undefined, "➡");
            arrow3Radio.helpTip = formatLabel("tooltip.arrow3", [ARROW3_HEIGHT_TO_STROKE_RATIO]);
            var arrowSlashRadio = panel.add("radiobutton", undefined, getLabel("radio.arrowSlash"));
            var plusRadio = panel.add("radiobutton", undefined, "＋");
            var multiplyRadio = panel.add("radiobutton", undefined, "×");

            var mirrorRow = panel.add("group");
            setupRow(mirrorRow, "left");
            mirrorRow.margins = [0, 6, 0, 0];

            var mirrorCheckbox = mirrorRow.add("checkbox", undefined, getLabel("checkbox.mirrorHorizontal"));
            mirrorCheckbox.value = false;

            triangleRadio.value = true;

            return {
                panel: panel,
                triangleRadio: triangleRadio,
                chevronRadio: chevronRadio,
                chevronDoubleRadio: chevronDoubleRadio,
                arrowRadio: arrowRadio,
                arrow3Radio: arrow3Radio,
                arrowSlashRadio: arrowSlashRadio,
                dashRadio: dashRadio,
                plusRadio: plusRadio,
                multiplyRadio: multiplyRadio,
                mirrorCheckbox: mirrorCheckbox
            };
        }

        /**
         * ［先端］パネルを作成します。
         *
         * @param {object} parent - 追加先のコンテナ。
         * @returns {object} パネルと各コントロール。
         */
        function buildCapPanel(parent) {
            var panel = parent.add("panel", undefined, getLabel("panel.cap"));
            setupPanel(panel, 6);

            var capNoneRadio = panel.add("radiobutton", undefined, getLabel("radio.capNone"));
            var capRoundRadio = panel.add("radiobutton", undefined, getLabel("radio.capRound"));
            capNoneRadio.value = true;

            return {
                panel: panel,
                capNoneRadio: capNoneRadio,
                capRoundRadio: capRoundRadio
            };
        }

        /**
         * ［オプション］パネルを作成します。
         *
         * @param {object} parent - 追加先のコンテナ。
         * @returns {object} パネルと各コントロール。
         */
        function buildOptionsPanel(parent) {
            var panel = parent.add("panel", undefined, getLabel("panel.options"));
            setupPanel(panel, FIELD_ROW_SPACING);

            var heightRow = addLabeledField(panel, "label.height", "20", "%");
            heightRow.field.active = true;

            var widthRow = addLabeledField(panel, "label.width", "", rulerUnitInfo.label);

            var insetRow = addLabeledField(panel, "label.inset", "0", rulerUnitInfo.label);
            insetRow.field.helpTip = formatLabel("tooltip.inset", [roundDisplayValue(MAX_INSET_RATIO * 100)]);
            insetRow.field.enabled = false;

            var gapRow = addLabeledField(panel, "label.gap", String(roundDisplayValueForUnit(convertPtToUnitValue(-1, rulerUnitInfo), rulerUnitInfo)), rulerUnitInfo.label);
            gapRow.field.enabled = false;

            var strokeRow = addLabeledField(panel, "label.stroke", String(roundDisplayValueForUnit(convertPtToUnitValue(0.3, strokeUnitInfo), strokeUnitInfo)), strokeUnitInfo.label);

            var angleRow = addLabeledField(panel, "label.angle", "0", "°");

            var alignCenterRow = panel.add("group");
            setupRow(alignCenterRow, "left");
            alignCenterRow.margins = [40, 10, 0, 0];

            var alignVerticalCenterCheckbox = alignCenterRow.add("checkbox", undefined, getLabel("checkbox.alignVerticalCenter"));
            alignVerticalCenterCheckbox.value = false;

            var roundCornersRow = panel.add("group");
            setupRow(roundCornersRow, "left");
            roundCornersRow.margins = [40, 0, 0, 0];

            var roundCornersCheckbox = roundCornersRow.add("checkbox", undefined, getLabel("checkbox.roundCorners"));
            roundCornersCheckbox.value = true;

            return {
                panel: panel,
                heightRow: heightRow.row,
                heightField: heightRow.field,
                widthField: widthRow.field,
                insetRow: insetRow.row,
                insetField: insetRow.field,
                gapRow: gapRow.row,
                gapField: gapRow.field,
                strokeRow: strokeRow.row,
                strokeField: strokeRow.field,
                angleLabel: angleRow.label,
                angleField: angleRow.field,
                angleUnitLabel: angleRow.unitLabel,
                alignVerticalCenterCheckbox: alignVerticalCenterCheckbox,
                roundCornersCheckbox: roundCornersCheckbox
            };
        }

        /**
         * ［位置調整］パネルを作成します。
         *
         * @param {object} parent - 追加先のコンテナ。
         * @returns {object} パネルと各コントロール。
         */
        function buildAdjustPanel(parent) {
            var panel = parent.add("panel", undefined, getLabel("panel.adjust"));
            setupPanel(panel, FIELD_ROW_SPACING);

            var horizontalRow = addLabeledField(panel, "label.adjustX", "0", rulerUnitInfo.label);
            var verticalRow = addLabeledField(panel, "label.adjustY", "0", rulerUnitInfo.label);

            return {
                panel: panel,
                offsetXField: horizontalRow.field,
                offsetYField: verticalRow.field
            };
        }

        /**
         * プレビューチェックボックスとボタンの行を作成します。
         *
         * @param {object} dialog - 追加先のダイアログ。
         * @returns {object} 行と各コントロール。
         */
        function buildButtonRow(dialog) {
            var buttonRow = dialog.add("group");
            setupRow(buttonRow, "fill", FIELD_ROW_SPACING);
            buttonRow.margins = [0, 10, 0, 0];

            var previewCheckbox = buttonRow.add("checkbox", undefined, getLabel("checkbox.preview"));
            previewCheckbox.value = false;

            /* 空グループでボタンを右端へ押しやる / Empty group pushes the buttons to the right edge */
            var buttonSpacer = buttonRow.add("group");
            buttonSpacer.alignment = ["fill", "center"];

            var cancelButton = buttonRow.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
            var okButton = buttonRow.add("button", undefined, getLabel("button.ok"), { name: "ok" });

            return {
                buttonRow: buttonRow,
                previewCheckbox: previewCheckbox,
                buttonSpacer: buttonSpacer,
                cancelButton: cancelButton,
                okButton: okButton
            };
        }

        /**
         * ダイアログの中身を組み立て、コントロールをまとめて返します。
         *
         * @param {object} dialog - 対象のダイアログ。
         * @returns {object} すべてのコントロールを持つオブジェクト。
         */
        function buildDialogControls(dialog) {
            var mainRow = dialog.add("group");
            setupRow(mainRow, "fill", COLUMN_SPACING);
            mainRow.alignChildren = ["fill", "top"];

            var leftColumn = mainRow.add("group");
            setupColumn(leftColumn, 10);

            var rightColumn = mainRow.add("group");
            setupColumn(rightColumn, 10);

            var shapeUI = buildShapePanel(leftColumn);
            var capUI = buildCapPanel(leftColumn);
            var optionsUI = buildOptionsPanel(rightColumn);
            var adjustUI = buildAdjustPanel(rightColumn);
            var buttonUI = buildButtonRow(dialog);

            return {
                shapePanel: shapeUI.panel,
                capPanel: capUI.panel,
                optionsPanel: optionsUI.panel,
                adjustPanel: adjustUI.panel,
                buttonRow: buttonUI.buttonRow,
                triangleRadio: shapeUI.triangleRadio,
                chevronRadio: shapeUI.chevronRadio,
                chevronDoubleRadio: shapeUI.chevronDoubleRadio,
                arrowRadio: shapeUI.arrowRadio,
                arrow3Radio: shapeUI.arrow3Radio,
                arrowSlashRadio: shapeUI.arrowSlashRadio,
                dashRadio: shapeUI.dashRadio,
                plusRadio: shapeUI.plusRadio,
                multiplyRadio: shapeUI.multiplyRadio,
                mirrorCheckbox: shapeUI.mirrorCheckbox,
                capNoneRadio: capUI.capNoneRadio,
                capRoundRadio: capUI.capRoundRadio,
                heightRow: optionsUI.heightRow,
                insetRow: optionsUI.insetRow,
                heightField: optionsUI.heightField,
                widthField: optionsUI.widthField,
                insetField: optionsUI.insetField,
                gapRow: optionsUI.gapRow,
                gapField: optionsUI.gapField,
                strokeRow: optionsUI.strokeRow,
                strokeField: optionsUI.strokeField,
                angleLabel: optionsUI.angleLabel,
                angleField: optionsUI.angleField,
                angleUnitLabel: optionsUI.angleUnitLabel,
                alignVerticalCenterCheckbox: optionsUI.alignVerticalCenterCheckbox,
                roundCornersCheckbox: optionsUI.roundCornersCheckbox,
                offsetXField: adjustUI.offsetXField,
                offsetYField: adjustUI.offsetYField,
                previewCheckbox: buttonUI.previewCheckbox,
                cancelButton: buttonUI.cancelButton,
                okButton: buttonUI.okButton
            };
        }

        /* どの組にもアキがなければ記号を作成できないので、ダイアログを出す前に知らせる / Fail early when no pair has a usable gap */
        if (!measureNarrowestGap(null)) {
            alert(getLabel("alert.noGap"));
            return;
        }

        var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(dialog);

        var controls = buildDialogControls(dialog);

        /* 幅を手入力したかどうか（手入力後は自動計算しない）/ Whether the width was typed in (auto-calculation stops once it is) */
        var widthManuallySet = false;

        /* プレビューで作成したアイテム / Items created for the preview */
        var previewItems = [];

        // =========================================
        // 幅の自動計算 / Automatic width
        // =========================================

        /**
         * 選択オブジェクトを走査し、最も狭いアキと最も低い合計高さを求めます。
         *
         * @param {number} heightPercent - 高さ（％）。高さ入力を使わない形状では NaN でも構いません。
         * @returns {object} minGapWidth / minTotalHeight を持つオブジェクト。求められない場合は null。
         */
        function measureNarrowestGap(heightPercent) {
            var items = getSortedSelectionItems();
            if (items.length < 2) return null;

            var minGapWidth = null;
            var minTotalHeight = null;

            for (var i = 0; i < items.length - 1; i++) {
                var placement = computePlacementBetweenItems(items[i], items[i + 1], heightPercent, 0, 0, 0);
                if (!placement) continue;

                var gapWidth = placement.gapRight - placement.gapLeft;
                if (minGapWidth === null || gapWidth < minGapWidth) {
                    minGapWidth = gapWidth;
                }
                if (minTotalHeight === null || placement.totalHeight < minTotalHeight) {
                    minTotalHeight = placement.totalHeight;
                }
            }

            if (minGapWidth === null || minTotalHeight === null) return null;

            return { minGapWidth: minGapWidth, minTotalHeight: minTotalHeight };
        }

        /**
         * 現在の形状と高さから、幅の初期値（pt）を求めます。
         *
         * @returns {number} 幅（pt）。求められない場合は null。
         */
        function computeAutoWidthPt() {
            var shapeConfig = SHAPE_CONFIG[getSelectedShapeKey()];
            if (!shapeConfig || !shapeConfig.calcDefaultWidth) return null;

            var heightPercent = parseFloat(controls.heightField.text);
            if (shapeConfig.enableHeightInput && (isNaN(heightPercent) || heightPercent <= 0)) return null;

            var gapInfo = measureNarrowestGap(heightPercent);
            if (!gapInfo) return null;

            var autoWidthPt = shapeConfig.calcDefaultWidth(heightPercent, gapInfo.minTotalHeight, gapInfo.minGapWidth);
            if (autoWidthPt === null || typeof autoWidthPt === "undefined") return null;

            return autoWidthPt;
        }

        /**
         * 幅の入力欄に、自動計算した初期値を表示します。
         *
         * @returns {void}
         */
        function applyAutoWidthToField() {
            var autoWidthPt = computeAutoWidthPt();
            if (autoWidthPt === null) return;
            setFieldFromPt(controls.widthField, autoWidthPt, rulerUnitInfo);
        }

        /**
         * 幅を自動計算に戻します。
         *
         * @returns {void}
         */
        function resetWidthToAuto() {
            widthManuallySet = false;
            applyAutoWidthToField();
        }

        /**
         * 実際に使われる幅（pt）を求めます。手入力があればその値を優先します。
         *
         * @returns {number} 幅（pt）。求められない場合は 0。
         */
        function getEffectiveWidthPt() {
            var typedWidthPt = parseFieldToPt(controls.widthField, rulerUnitInfo, false);
            if (!isNaN(typedWidthPt) && typedWidthPt > 0) return typedWidthPt;

            var autoWidthPt = computeAutoWidthPt();
            return (autoWidthPt && autoWidthPt > 0) ? autoWidthPt : 0;
        }

        // =========================================
        // 形状の切り替え / Shape switching
        // =========================================

        /**
         * 選択中の形状キーを返します。
         *
         * @returns {string} SHAPE_CONFIG のキー。
         */
        function getSelectedShapeKey() {
            for (var shapeKey in SHAPE_CONFIG) {
                if (!SHAPE_CONFIG.hasOwnProperty(shapeKey)) continue;
                var radio = controls[SHAPE_CONFIG[shapeKey].radioKey];
                if (radio && radio.value) return shapeKey;
            }
            return "tri";
        }

        /**
         * 形状のラジオボタンをまとめて返します。
         *
         * @returns {array} ラジオボタンの配列。
         */
        function getShapeRadioControls() {
            var radios = [];
            for (var shapeKey in SHAPE_CONFIG) {
                if (!SHAPE_CONFIG.hasOwnProperty(shapeKey)) continue;
                radios.push(controls[SHAPE_CONFIG[shapeKey].radioKey]);
            }
            return radios;
        }

        /**
         * 形状に応じて、各入力欄の初期値を設定します。
         *
         * @param {object} shapeConfig - 形状の設定。
         * @returns {void}
         */
        function applyShapeFieldDefaults(shapeConfig) {
            if (shapeConfig.enableHeightInput && shapeConfig.defaultHeightPercent !== undefined) {
                controls.heightField.text = String(shapeConfig.defaultHeightPercent);
            }
            if (!widthManuallySet) {
                applyAutoWidthToField();
            }
            setFieldFromPt(controls.gapField, shapeConfig.defaultGap, rulerUnitInfo);
            setFieldFromPt(controls.insetField, shapeConfig.defaultInsetPt || 0, rulerUnitInfo);
        }

        /**
         * 形状に応じて、角度欄の初期値と有効・無効を切り替えます。
         * 高さや幅と同じく、形状を切り替えるたびにその形状の初期値へ戻します。
         *
         * @param {object} shapeConfig - 形状の設定。
         * @returns {void}
         */
        function applyAngleFieldState(shapeConfig) {
            var isAngleEnabled = !!shapeConfig.enableAngle;
            controls.angleField.enabled = isAngleEnabled;
            controls.angleLabel.enabled = isAngleEnabled;
            controls.angleUnitLabel.enabled = isAngleEnabled;

            var defaultAngle = (typeof shapeConfig.defaultAngle !== "undefined") ? shapeConfig.defaultAngle : 0;
            controls.angleField.text = String(defaultAngle);
        }

        /**
         * 形状に応じて、線幅欄の初期値を入れ直します。
         * 形状ごとに適切な線幅が違うため、切り替えるたびにその形状の初期値へ戻します。
         *
         * @param {object} shapeConfig - 形状の設定。
         * @returns {void}
         */
        function applyStrokeFieldState(shapeConfig) {
            setFieldFromPt(controls.strokeField, shapeConfig.defaultStrokePt || 0, strokeUnitInfo);
        }

        /**
         * 形状に応じて、各コントロールの有効・無効を切り替えます。
         *
         * @param {string} shapeKey - 形状キー。
         * @param {object} shapeConfig - 形状の設定。
         * @returns {void}
         */
        function applyShapeEnabledStates(shapeKey, shapeConfig) {
            controls.capPanel.enabled = !!shapeConfig.enableCapPanel;
            controls.heightRow.enabled = !!shapeConfig.enableHeightInput;
            /* ラベルと単位表記もまとめてグレーにするため、行ごと切り替える / Toggle the whole row so the label and unit gray out too */
            controls.strokeRow.enabled = !!shapeConfig.enableStrokeInput;
            controls.strokeField.enabled = !!shapeConfig.enableStrokeInput;
            controls.gapRow.enabled = !!shapeConfig.enableGap;
            controls.gapField.enabled = !!shapeConfig.enableGap;
            controls.insetRow.enabled = !!shapeConfig.enableInset;
            controls.insetField.enabled = !!shapeConfig.enableInset;

            /* 天地を水平にできるのは > と >> だけ / Only the chevrons can keep their top and bottom edges horizontal */
            var isChevronShape = (shapeKey === "chevron" || shapeKey === "chevron2");
            controls.alignVerticalCenterCheckbox.enabled = isChevronShape;
            if (!isChevronShape) controls.alignVerticalCenterCheckbox.value = false;

            controls.mirrorCheckbox.enabled = !!shapeConfig.enableMirror;
            if (!shapeConfig.enableMirror) controls.mirrorCheckbox.value = false;

            controls.roundCornersCheckbox.enabled = !!shapeConfig.enableRoundCorners;
            if (!shapeConfig.enableRoundCorners) controls.roundCornersCheckbox.value = false;

            if (!shapeConfig.enableCapPanel) {
                controls.capNoneRadio.value = true;
                controls.capRoundRadio.value = false;
            }
        }

        /**
         * 選択中の形状に合わせて、ダイアログ全体の初期値と状態を更新します。
         *
         * @returns {void}
         */
        function applyShapeDefaults() {
            var shapeKey = getSelectedShapeKey();
            var shapeConfig = SHAPE_CONFIG[shapeKey];

            applyShapeFieldDefaults(shapeConfig);
            applyAngleFieldState(shapeConfig);
            applyStrokeFieldState(shapeConfig);
            applyShapeEnabledStates(shapeKey, shapeConfig);
        }

        // =========================================
        // 入力値の読み取り / Reading the input values
        // =========================================

        /**
         * 入力エラーを知らせ、対象の入力欄にフォーカスを移します。
         *
         * @param {object} field - 対象の edittext。
         * @param {string} alertPath - 表示するメッセージのドットパス。
         * @param {boolean} showAlert - 警告を表示するなら true。
         * @param {array} [values] - メッセージに差し込む値。
         * @returns {object} 呼び出し元がそのまま返せるよう null を返します。
         */
        function reportInvalidValue(field, alertPath, showAlert, values) {
            if (showAlert) {
                alert(formatLabel(alertPath, values));
                field.active = true;
            }
            return null;
        }

        /**
         * ダイアログの入力値をまとめて読み取り、検証します。
         *
         * @param {boolean} showAlert - 不正な値のときに警告を表示するなら true。
         * @returns {object} 検証済みの入力値。不正な場合は null。
         */
        function readInputValues(showAlert) {
            var shapeKey = getSelectedShapeKey();
            var shapeConfig = SHAPE_CONFIG[shapeKey];
            var heightPercent = null;
            var widthPt = parseFieldToPt(controls.widthField, rulerUnitInfo, false);
            var insetPt = shapeConfig.enableInset ? parseFieldToPt(controls.insetField, rulerUnitInfo, false) : 0;
            var offsetX = parseFieldToPt(controls.offsetXField, rulerUnitInfo, true);
            var offsetY = parseFieldToPt(controls.offsetYField, rulerUnitInfo, true);
            var strokeWidthPt = parseFieldToPt(controls.strokeField, strokeUnitInfo, false);
            var angleDeg = shapeConfig.enableAngle ? parseFloat(controls.angleField.text) : 0;

            if (shapeConfig.enableInset) {
                if (isNaN(insetPt) || insetPt < 0) {
                    return reportInvalidValue(controls.insetField, "alert.invalidValue", showAlert);
                }
                /* 幅が明示されている場合は、その割合まで凹みを抑える / Clamp the inset when the width is set explicitly */
                if (widthPt > 0) {
                    insetPt = Math.min(insetPt, widthPt * MAX_INSET_RATIO);
                }
            }

            if (shapeConfig.enableHeightInput) {
                heightPercent = parseFloat(controls.heightField.text);
                if (isNaN(heightPercent) || heightPercent <= 0) {
                    return reportInvalidValue(controls.heightField, "alert.positiveNumber", showAlert);
                }
                if (heightPercent > MAX_HEIGHT_PERCENT) {
                    return reportInvalidValue(controls.heightField, "alert.maxHeight", showAlert, [MAX_HEIGHT_PERCENT]);
                }
            }

            if (isNaN(widthPt) || widthPt < 0) {
                return reportInvalidValue(controls.widthField, "alert.widthPositive", showAlert);
            }
            if (isNaN(offsetX)) {
                return reportInvalidValue(controls.offsetXField, "alert.invalidValue", showAlert);
            }
            if (isNaN(offsetY)) {
                return reportInvalidValue(controls.offsetYField, "alert.invalidValue", showAlert);
            }
            if (isNaN(strokeWidthPt)) {
                return reportInvalidValue(controls.strokeField, "alert.invalidValue", showAlert);
            }
            /* 塗りだけの形状は線幅を使わないので、下限を求めない / Fill-only shapes ignore the stroke width, so the minimum does not apply */
            if (shapeConfig.requirePositiveStroke && strokeWidthPt < MIN_STROKE_WIDTH_PT) {
                var minStrokeText = roundDisplayValueForUnit(convertPtToUnitValue(MIN_STROKE_WIDTH_PT, strokeUnitInfo), strokeUnitInfo);
                return reportInvalidValue(controls.strokeField, "alert.strokePositive", showAlert, [minStrokeText, strokeUnitInfo.label]);
            }
            if (shapeConfig.enableAngle && isNaN(angleDeg)) {
                return reportInvalidValue(controls.angleField, "alert.invalidValue", showAlert);
            }

            return {
                shapeKey: shapeKey,
                shapeConfig: shapeConfig,
                heightPercent: heightPercent,
                widthPt: widthPt,
                insetPt: insetPt,
                offsetX: offsetX,
                offsetY: offsetY,
                strokeWidthPt: strokeWidthPt,
                angleDeg: angleDeg
            };
        }

        // =========================================
        // パス作成のヘルパー / Path creation helpers
        // =========================================

        /**
         * 記号の色を作成します。
         *
         * @returns {object} CMYKColor。
         */
        function createMarkColor() {
            var color = new CMYKColor();
            color.cyan = MARK_COLOR_CMYK[0];
            color.magenta = MARK_COLOR_CMYK[1];
            color.yellow = MARK_COLOR_CMYK[2];
            color.black = MARK_COLOR_CMYK[3];
            return color;
        }

        /**
         * 線の先端・角の形状を、［先端］パネルの選択に合わせて設定します。
         *
         * @param {object} path - 対象のパス。
         * @param {object} options - cap / join を適用するか、丸型でないときに明示的に指定するか。
         * @returns {void}
         */
        function applyStrokeEnds(path, options) {
            if (controls.capRoundRadio.value) {
                if (options.cap) path.strokeCap = StrokeCap.ROUNDENDCAP;
                if (options.join) path.strokeJoin = StrokeJoin.ROUNDENDJOIN;
                return;
            }
            if (options.setFlatEnds) {
                if (options.cap) path.strokeCap = StrokeCap.BUTTENDCAP;
                if (options.join) path.strokeJoin = StrokeJoin.MITERENDJOIN;
            }
        }

        /**
         * 線だけのパスを作成します。
         *
         * @param {array} points - アンカーポイントの配列。
         * @param {number} strokeWidthPt - 線幅（pt）。
         * @param {object} [endOptions] - applyStrokeEnds に渡すオプション。
         * @param {boolean} [closed] - 閉じたパスにするなら true。
         * @returns {object} 作成した PathItem。
         */
        function createStrokedPath(points, strokeWidthPt, endOptions, closed) {
            var path = activeDocument.activeLayer.pathItems.add();
            path.setEntirePath(points);
            path.closed = !!closed;
            path.filled = false;
            path.stroked = true;
            path.strokeWidth = strokeWidthPt;
            path.strokeColor = createMarkColor();
            if (endOptions) applyStrokeEnds(path, endOptions);
            return path;
        }

        /**
         * 塗りだけのパスを作成します。
         *
         * @param {array} points - アンカーポイントの配列。
         * @returns {object} 作成した PathItem。
         */
        function createFilledPath(points) {
            var path = activeDocument.activeLayer.pathItems.add();
            path.setEntirePath(points);
            path.closed = true;
            path.filled = true;
            path.fillColor = createMarkColor();
            path.stroked = false;
            return path;
        }

        /**
         * 複数のアイテムを1つのグループにまとめます。
         *
         * @param {array} items - まとめるアイテムの配列。
         * @returns {object} 作成した GroupItem。
         */
        function groupItems(items) {
            var group = activeDocument.activeLayer.groupItems.add();
            for (var i = 0; i < items.length; i++) {
                items[i].move(group, ElementPlacement.INSIDE);
            }
            return group;
        }

        /**
         * 角丸のライブエフェクトを適用します。
         *
         * @param {object} item - 対象のアイテム。
         * @param {number} radiusPt - 角丸の半径（pt）。
         * @returns {void}
         */
        function applyRoundCornersEffect(item, radiusPt) {
            if (!item || radiusPt <= 0) return;
            runSafely(function () {
                item.applyEffect('<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + radiusPt + ' "/></LiveEffect>');
            }, getLabel("log.applyRoundCorners"));
        }

        // =========================================
        // 形状の作成 / Shape creation
        // =========================================

        /**
         * ▶ の角丸半径を求めます。
         *
         * @param {number} width - 幅（pt）。
         * @param {number} height - 高さ（pt）。
         * @param {number} insetPt - 凹み（pt）。
         * @returns {number} 角丸の半径（pt）。
         */
        function getTriangleCornerRadius(width, height, insetPt) {
            var radius = Math.min(width, height) * TRI_CORNER_RADIUS_RATIO;
            if (insetPt > 0) {
                radius = Math.min(radius, insetPt * 0.45);
            }
            return Math.max(1, radius);
        }

        /**
         * ▶（三角形）を作成します。凹みを指定すると左辺がへこんだ形になります。
         *
         * @param {number} centerX - 中心の X 座標。
         * @param {number} centerY - 中心の Y 座標。
         * @param {number} width - 幅（pt）。
         * @param {number} height - 高さ（pt）。
         * @param {number} insetPt - 凹み（pt）。
         * @param {number} strokeWidthPt - 線幅（pt）。
         * @param {boolean} forceFillOnly - 塗りのみにするなら true。
         * @returns {object} 作成した PathItem。
         */
        function createTriangleShape(centerX, centerY, width, height, insetPt, strokeWidthPt, forceFillOnly) {
            var inset = Math.max(0, Math.min(insetPt || 0, width * MAX_INSET_RATIO));
            var leftX = centerX - width / 2;
            var rightX = centerX + width / 2;
            var halfHeight = height / 2;

            var points = (inset > 0)
                ? [[rightX, centerY], [leftX, centerY + halfHeight], [leftX + inset, centerY], [leftX, centerY - halfHeight]]
                : [[rightX, centerY], [leftX, centerY + halfHeight], [leftX, centerY - halfHeight]];

            var path = createFilledPath(points);

            if (!forceFillOnly && strokeWidthPt > 0) {
                path.stroked = true;
                path.strokeWidth = strokeWidthPt;
                path.strokeColor = createMarkColor();
                path.strokeJoin = StrokeJoin.ROUNDENDJOIN;
            }

            if (controls.roundCornersCheckbox.value) {
                applyRoundCornersEffect(path, getTriangleCornerRadius(width, height, inset));
            }
            return path;
        }

        /**
         * 矢じりの奥行き（水平方向の長さ）を求めます。
         * 高さの半分を基本としつつ、矢じりが幅からはみ出さないよう上限を設けます。
         *
         * @param {number} width - 幅（pt）。
         * @param {number} height - 高さ（pt）。
         * @returns {number} 矢じりの奥行き（pt）。
         */
        function getArrowHeadDepth(width, height) {
            var headDepth = height / 2;
            if (width > 0) headDepth = Math.min(headDepth, width * MAX_ARROW_HEAD_RATIO);
            return headDepth;
        }

        /**
         * →（線の矢印）を作成します。
         *
         * @param {number} centerX - 中心の X 座標。
         * @param {number} centerY - 中心の Y 座標。
         * @param {number} width - 幅（pt）。
         * @param {number} height - 高さ（pt）。
         * @param {number} strokeWidthPt - 線幅（pt）。
         * @returns {object} 作成した GroupItem。
         */
        function createArrowShape(centerX, centerY, width, height, strokeWidthPt) {
            var headHalfHeight = height / 2;
            var headDepth = getArrowHeadDepth(width, height);
            var tipX = centerX + width / 2;

            var shaft = createStrokedPath([
                [centerX - width / 2, centerY],
                [tipX, centerY]
            ], strokeWidthPt, { cap: true });

            var head = createStrokedPath([
                [tipX - headDepth, centerY + headHalfHeight],
                [tipX, centerY],
                [tipX - headDepth, centerY - headHalfHeight]
            ], strokeWidthPt, { cap: true, join: true });

            return groupItems([shaft, head]);
        }

        /**
         * ➡ の軸の太さを求めます。線幅をそのまま使います。
         *
         * @param {number} strokeWidthPt - 線幅（pt）。
         * @returns {number} 軸の太さ（pt）。
         */
        function getSolidArrowShaftWidth(strokeWidthPt) {
            return (strokeWidthPt > 0) ? strokeWidthPt : 1;
        }

        /**
         * ➡ の高さ（矢じりの天地）を線幅から求めます。
         * 高さ（％）は使わず、線幅を変えると軸と矢じりが一緒に拡大縮小します。
         *
         * @param {number} strokeWidthPt - 線幅（pt）。
         * @returns {number} 高さ（pt）。
         */
        function getSolidArrowHeight(strokeWidthPt) {
            return getSolidArrowShaftWidth(strokeWidthPt) * ARROW3_HEIGHT_TO_STROKE_RATIO;
        }

        /**
         * ➡（塗りの矢印）の軸と矢じりを作成します。
         * 高さは線幅から決まるため、引数では受け取りません。
         *
         * @param {number} centerX - 中心の X 座標。
         * @param {number} centerY - 中心の Y 座標。
         * @param {number} width - 幅（pt）。
         * @param {number} strokeWidthPt - 線幅（pt）。
         * @returns {object} shaft / head を持つオブジェクト。
         */
        function createSolidArrowParts(centerX, centerY, width, strokeWidthPt) {
            var shaftWidthPt = getSolidArrowShaftWidth(strokeWidthPt);
            var height = getSolidArrowHeight(strokeWidthPt);
            var headHalfHeight = height / 2;
            var headDepth = getArrowHeadDepth(width, height);
            var tipX = centerX + width / 2;

            var shaft = createStrokedPath([
                [centerX - width / 2, centerY],
                [tipX - headDepth, centerY]
            ], shaftWidthPt);

            var head = createFilledPath([
                [tipX - headDepth, centerY - headHalfHeight],
                [tipX, centerY],
                [tipX - headDepth, centerY + headHalfHeight]
            ]);

            return { shaft: shaft, head: head };
        }

        /**
         * ➡ の軸と矢じりを、アウトライン化と合体で1つのパスにまとめます。
         * 選択状態を使うメニューコマンドを呼ぶため、呼び出し側（createMarks）で選択を退避・復元します。
         * 合成できなかった場合も、軸と矢じりを取り残さないよう1つのグループにまとめて返します。
         *
         * @param {object} parts - createSolidArrowParts の戻り値。
         * @returns {object} 合成後のアイテム。
         */
        function mergeSolidArrowParts(parts) {
            if (!parts || !parts.shaft || !parts.head) return parts;

            var outlinedShaft = parts.shaft;

            try {
                activeDocument.selection = [parts.shaft];
                app.executeMenuCommand("Live Outline Stroke");
                if (activeDocument.selection && activeDocument.selection.length > 0) {
                    outlinedShaft = activeDocument.selection[0];
                }

                activeDocument.selection = [outlinedShaft, parts.head];
                app.executeMenuCommand("group");
                app.executeMenuCommand("Live Pathfinder Add");

                /* 1つにまとまったときだけ成功とみなす / Treat it as merged only when a single item is left */
                if (activeDocument.selection && activeDocument.selection.length === 1) {
                    return activeDocument.selection[0];
                }
            } catch (e) {
                logScriptError(getLabel("log.mergeSolidArrow"), e);
            }

            /* メニューコマンドが効かなかった場合のフォールバック / Fallback when the menu commands did not take effect */
            var fallbackGroup = null;
            runSafely(function () {
                fallbackGroup = groupItems([outlinedShaft, parts.head]);
            }, getLabel("log.mergeSolidArrow"));

            return fallbackGroup || outlinedShaft;
        }

        /**
         * ＿\（横線＋斜線）を作成します。
         *
         * @param {number} centerX - 中心の X 座標。
         * @param {number} centerY - 中心の Y 座標。
         * @param {number} width - 幅（pt）。
         * @param {number} height - 高さ（pt）。
         * @param {number} strokeWidthPt - 線幅（pt）。
         * @param {number} angleDeg - 斜線の角度（度）。
         * @returns {object} 作成した PathItem。
         */
        function createArrowSlashShape(centerX, centerY, width, height, strokeWidthPt, angleDeg) {
            var lineWidthPt = (strokeWidthPt > 0) ? strokeWidthPt : 1;
            var rise = height / 2;
            var leftX = centerX - width / 2;
            var tipX = centerX + width / 2;

            var slashAngle = angleDeg;
            if (isNaN(slashAngle) || slashAngle <= 0) slashAngle = SLASH_ANGLE_DEFAULT;
            if (slashAngle >= SLASH_ANGLE_MAX) slashAngle = SLASH_ANGLE_MAX;

            var slashDx = rise / Math.tan(slashAngle * Math.PI / 180);
            if (!isFinite(slashDx) || slashDx <= 0) slashDx = rise;

            var slashTopX = tipX - slashDx;
            if (slashTopX <= leftX) {
                slashTopX = leftX + Math.max(lineWidthPt, width * 0.15);
            }

            return createStrokedPath([
                [leftX, centerY],
                [tipX, centerY],
                [slashTopX, centerY + rise]
            ], lineWidthPt, { cap: true, join: true, setFlatEnds: true });
        }

        /**
         * 天地を水平にした > の腕を1本作成します。
         *
         * @param {number} startX - 左端の X 座標。
         * @param {number} startY - 左端の Y 座標。
         * @param {number} tipX - 先端の X 座標。
         * @param {number} tipY - 先端の Y 座標。
         * @param {number} thickness - 腕の太さ（pt）。
         * @returns {object} 作成した PathItem。
         */
        function createFlatChevronArm(startX, startY, tipX, tipY, thickness) {
            return createFilledPath([
                [startX, startY],
                [startX + thickness, startY],
                [tipX, tipY],
                [tipX - thickness, tipY]
            ]);
        }

        /**
         * >> の2つの山形の間隔（pt）を読み取ります。
         *
         * @returns {number} 間隔（pt）。数値でない場合は 0。
         */
        function getChevronGapPt() {
            var gapPt = parseFieldToPt(controls.gapField, rulerUnitInfo, true);
            return isNaN(gapPt) ? 0 : gapPt;
        }

        /**
         * 天地を水平にした >（塗りの山形）を作成します。
         *
         * @param {number} centerX - 中心の X 座標。
         * @param {number} centerY - 中心の Y 座標。
         * @param {number} height - 高さ（pt）。
         * @param {number} widthPt - 幅（pt）。0以下なら高さと同じ幅にします。
         * @param {number} strokeWidthPt - 腕の太さ（pt）。
         * @returns {object} 作成した GroupItem。
         */
        function createFlatChevronShape(centerX, centerY, height, widthPt, strokeWidthPt) {
            var fullWidth = (widthPt > 0) ? widthPt : height;
            var thickness = (strokeWidthPt > 0) ? strokeWidthPt : 1;
            var halfHeight = height / 2;
            var leftX = centerX - fullWidth / 2;
            var tipX = centerX + fullWidth / 2;

            var topArm = createFlatChevronArm(leftX, centerY + halfHeight, tipX, centerY, thickness);
            var bottomArm = createFlatChevronArm(leftX, centerY - halfHeight, tipX, centerY, thickness);

            return groupItems([topArm, bottomArm]);
        }

        /**
         * 天地を水平にした >>（塗りの二重山形）を作成します。
         *
         * @param {number} centerX - 中心の X 座標。
         * @param {number} centerY - 中心の Y 座標。
         * @param {number} height - 高さ（pt）。
         * @param {number} widthPt - 幅（pt）。
         * @param {number} strokeWidthPt - 腕の太さ（pt）。
         * @returns {object} 作成した GroupItem。
         */
        function createFlatChevronDoubleShape(centerX, centerY, height, widthPt, strokeWidthPt) {
            var fullWidth = (widthPt > 0) ? widthPt : height;
            var gapPt = getChevronGapPt();

            var leftShape = createFlatChevronShape(centerX - (fullWidth + gapPt) / 2, centerY, height, fullWidth, strokeWidthPt);
            var rightShape = createFlatChevronShape(centerX + (fullWidth + gapPt) / 2, centerY, height, fullWidth, strokeWidthPt);

            return groupItems([leftShape, rightShape]);
        }

        /**
         * >（線の山形）を1つ作成します。
         *
         * @param {number} startX - 左端の X 座標。
         * @param {number} centerY - 中心の Y 座標。
         * @param {number} chevronWidth - 山形の幅（pt）。
         * @param {number} chevronHeight - 中心から上下それぞれの高さ（pt）。
         * @param {number} strokeWidthPt - 線幅（pt）。
         * @returns {object} 作成した PathItem。
         */
        function createChevronStroke(startX, centerY, chevronWidth, chevronHeight, strokeWidthPt) {
            return createStrokedPath([
                [startX, centerY + chevronHeight],
                [startX + chevronWidth, centerY],
                [startX, centerY - chevronHeight]
            ], strokeWidthPt, { cap: true, join: true });
        }

        /**
         * >（山形）を作成します。［天地を水平に］がONなら塗りの形状にします。
         *
         * @param {number} centerX - 中心の X 座標。
         * @param {number} centerY - 中心の Y 座標。
         * @param {number} height - 高さ（pt）。
         * @param {number} widthPt - 幅（pt）。0以下なら高さと同じ幅にします。
         * @param {number} strokeWidthPt - 線幅（pt）。
         * @returns {object} 作成したアイテム。
         */
        function createChevronShape(centerX, centerY, height, widthPt, strokeWidthPt) {
            if (controls.alignVerticalCenterCheckbox.value) {
                return createFlatChevronShape(centerX, centerY, height, widthPt, strokeWidthPt);
            }

            var chevronHeight = height / 2;
            var chevronWidth = (widthPt > 0) ? widthPt : height;

            return createChevronStroke(centerX - chevronWidth / 2, centerY, chevronWidth, chevronHeight, strokeWidthPt);
        }

        /**
         * >>（二重の山形）を作成します。［天地を水平に］がONなら塗りの形状にします。
         *
         * @param {number} centerX - 中心の X 座標。
         * @param {number} centerY - 中心の Y 座標。
         * @param {number} height - 高さ（pt）。
         * @param {number} widthPt - 山形1つ分の幅（pt）。0以下なら高さと同じ幅にします。
         * @param {number} strokeWidthPt - 線幅（pt）。
         * @returns {object} 作成したアイテム。
         */
        function createChevronDoubleShape(centerX, centerY, height, widthPt, strokeWidthPt) {
            if (controls.alignVerticalCenterCheckbox.value) {
                return createFlatChevronDoubleShape(centerX, centerY, height, widthPt, strokeWidthPt);
            }

            var chevronHeight = height / 2;
            var chevronWidth = (widthPt > 0) ? widthPt : height;
            var gapPt = getChevronGapPt();
            var leftStartX = centerX - (chevronWidth * 2 + gapPt) / 2;

            var leftChevron = createChevronStroke(leftStartX, centerY, chevronWidth, chevronHeight, strokeWidthPt);
            var rightChevron = createChevronStroke(leftStartX + chevronWidth + gapPt, centerY, chevronWidth, chevronHeight, strokeWidthPt);

            return groupItems([leftChevron, rightChevron]);
        }

        /**
         * ─（横線）を作成します。
         *
         * @param {number} centerX - 中心の X 座標。
         * @param {number} centerY - 中心の Y 座標。
         * @param {number} width - 幅（pt）。
         * @param {number} strokeWidthPt - 線幅（pt）。
         * @returns {object} 作成した PathItem。
         */
        function createDashShape(centerX, centerY, width, strokeWidthPt) {
            var halfWidth = width / 2;
            return createStrokedPath([
                [centerX - halfWidth, centerY],
                [centerX + halfWidth, centerY]
            ], strokeWidthPt, { cap: true });
        }

        /**
         * ＋（十字）を作成します。
         *
         * @param {number} centerX - 中心の X 座標。
         * @param {number} centerY - 中心の Y 座標。
         * @param {number} width - 幅（pt）。
         * @param {number} strokeWidthPt - 線幅（pt）。
         * @returns {object} 作成した GroupItem。
         */
        function createPlusShape(centerX, centerY, width, strokeWidthPt) {
            var halfWidth = width / 2;

            var horizontalLine = createStrokedPath([
                [centerX - halfWidth, centerY],
                [centerX + halfWidth, centerY]
            ], strokeWidthPt, { cap: true });

            var verticalLine = createStrokedPath([
                [centerX, centerY + halfWidth],
                [centerX, centerY - halfWidth]
            ], strokeWidthPt, { cap: true });

            return groupItems([horizontalLine, verticalLine]);
        }

        /**
         * ×（斜め十字）を作成します。＋ を45°回転させた形です。
         *
         * ほかの形状と違い、幅は実寸ではなく ＋ と同じ腕の長さを表します。
         * 実寸で揃えると ＋ より大きく見えるため、意図的にこの基準にしています。
         *
         * Unlike the other shapes, the width here is the arm length shared with the plus sign, not the drawn width.
         * Matching the drawn width would make it look larger than the plus sign, so this is intentional.
         *
         * @param {number} centerX - 中心の X 座標。
         * @param {number} centerY - 中心の Y 座標。
         * @param {number} width - 幅（pt）。描画幅は width × 0.707 になります。
         * @param {number} strokeWidthPt - 線幅（pt）。
         * @returns {object} 作成した GroupItem。
         */
        function createMultiplyShape(centerX, centerY, width, strokeWidthPt) {
            var diagonalOffset = (width / 2) * Math.SQRT2 / 2;

            var fallingLine = createStrokedPath([
                [centerX - diagonalOffset, centerY + diagonalOffset],
                [centerX + diagonalOffset, centerY - diagonalOffset]
            ], strokeWidthPt, { cap: true });

            var risingLine = createStrokedPath([
                [centerX - diagonalOffset, centerY - diagonalOffset],
                [centerX + diagonalOffset, centerY + diagonalOffset]
            ], strokeWidthPt, { cap: true });

            return groupItems([fallingLine, risingLine]);
        }

        /* 形状キーごとの作成関数 / Builder per shape key */
        var SHAPE_BUILDERS = {
            tri: function (placement, values) {
                return createTriangleShape(placement.centerX, placement.centerY, placement.width, placement.height, values.insetPt, values.strokeWidthPt, !!values.shapeConfig.forceFillOnly);
            },
            arrow: function (placement, values) {
                return createArrowShape(placement.centerX, placement.centerY, placement.arrowWidth, placement.height, values.strokeWidthPt);
            },
            arrow3: function (placement, values) {
                return mergeSolidArrowParts(createSolidArrowParts(placement.centerX, placement.centerY, placement.arrowWidth, values.strokeWidthPt));
            },
            arrowSlash: function (placement, values) {
                return createArrowSlashShape(placement.centerX, placement.centerY, placement.arrowWidth, placement.height, values.strokeWidthPt, values.angleDeg);
            },
            chevron: function (placement, values) {
                return createChevronShape(placement.centerX, placement.centerY, placement.height, values.widthPt, values.strokeWidthPt);
            },
            chevron2: function (placement, values) {
                return createChevronDoubleShape(placement.centerX, placement.centerY, placement.height, values.widthPt, values.strokeWidthPt);
            },
            dash: function (placement, values) {
                return createDashShape(placement.centerX, placement.centerY, placement.width, values.strokeWidthPt);
            },
            plus: function (placement, values) {
                return createPlusShape(placement.centerX, placement.centerY, placement.width, values.strokeWidthPt);
            },
            multiply: function (placement, values) {
                return createMultiplyShape(placement.centerX, placement.centerY, placement.width, values.strokeWidthPt);
            }
        };

        // =========================================
        // 配置と作成 / Placement and creation
        // =========================================

        /**
         * 隣り合う2つのオブジェクトから、記号を置く位置とサイズを求めます。
         *
         * @param {object} leftItem - 左側のアイテム。
         * @param {object} rightItem - 右側のアイテム。
         * @param {number} heightPercent - 高さ（％）。null なら合計高さをそのまま使います。
         * @param {number} widthPt - 幅（pt）。0以下なら高さから決めます。
         * @param {number} offsetX - 左右の位置調整（pt）。
         * @param {number} offsetY - 上下の位置調整（pt）。
         * @returns {object} 配置情報。アキがない場合は null。
         */
        function computePlacementBetweenItems(leftItem, rightItem, heightPercent, widthPt, offsetX, offsetY) {
            var firstBounds = getItemMeasurementBounds(leftItem);
            var secondBounds = getItemMeasurementBounds(rightItem);

            var leftBounds = (firstBounds[0] < secondBounds[0]) ? firstBounds : secondBounds;
            var rightBounds = (firstBounds[0] < secondBounds[0]) ? secondBounds : firstBounds;

            var gapLeft = leftBounds[2];
            var gapRight = rightBounds[0];
            if (gapRight - gapLeft <= 0) return null;

            var top = Math.max(firstBounds[1], secondBounds[1]);
            var bottom = Math.min(firstBounds[3], secondBounds[3]);
            var totalHeight = top - bottom;
            var usesFullHeight = (heightPercent === null || typeof heightPercent === "undefined");
            var markHeight = usesFullHeight ? totalHeight : (totalHeight * (heightPercent / 100));

            return {
                gapLeft: gapLeft,
                gapRight: gapRight,
                centerX: (gapLeft + gapRight) / 2 + (offsetX || 0),
                centerY: (top + bottom) / 2 + (offsetY || 0),
                totalHeight: totalHeight,
                height: markHeight,
                width: (widthPt > 0) ? widthPt : markHeight,
                /* 矢印系は幅未指定のとき、アキに対する割合で決める / Arrow shapes fall back to a ratio of the gap when no width is set */
                arrowWidth: (widthPt > 0) ? widthPt : ((gapRight - gapLeft) * AUTO_WIDTH_GAP_RATIO)
            };
        }

        /**
         * ［左右逆］がONの場合に、作成したアイテムを左右反転します。
         *
         * @param {object} item - 対象のアイテム。
         * @returns {object} 反転後のアイテム。
         */
        function applyMirrorIfNeeded(item) {
            if (!item || !controls.mirrorCheckbox.value) return item;

            try {
                item.resize(-100, 100, true, true, true, true, 100, Transformation.CENTER);
            } catch (e) {
                /* 中心を基準にできない場合は、左上基準で反転してから幅の分だけ右へ戻す / Mirror from the top-left, then shift right by its own width */
                logScriptError(getLabel("log.mirrorItem"), e);
                runSafely(function () {
                    var bounds = item.geometricBounds;
                    var itemWidth = bounds[2] - bounds[0];
                    item.resize(-100, 100, true, true, true, true, 100, Transformation.TOPLEFT);
                    item.translate(itemWidth, 0);
                }, getLabel("log.mirrorItem"));
            }
            return item;
        }

        /**
         * 隣り合う2つのオブジェクトの間に記号を1つ作成します。
         *
         * @param {object} leftItem - 左側のアイテム。
         * @param {object} rightItem - 右側のアイテム。
         * @param {object} values - 検証済みの入力値。
         * @returns {object} 作成したアイテム。アキがない場合は null。
         */
        function createMarkBetweenItems(leftItem, rightItem, values) {
            var placement = computePlacementBetweenItems(leftItem, rightItem, values.heightPercent, values.widthPt, values.offsetX, values.offsetY);
            if (!placement) return null;

            var buildShape = SHAPE_BUILDERS[values.shapeKey] || SHAPE_BUILDERS.tri;
            return applyMirrorIfNeeded(buildShape(placement, values));
        }

        /**
         * 選択オブジェクトの、隣り合うすべての組に記号を作成します。
         *
         * @param {object} values - 検証済みの入力値。
         * @returns {array} 作成したアイテムの配列。
         */
        function createMarks(values) {
            var items = getSortedSelectionItems();
            var createdItems = [];

            /* ➡ だけは選択状態を使うメニューコマンドを呼ぶため、ここで1回だけ退避・復元する / Only the solid arrow drives menu commands, so save and restore the selection once */
            var previousSelection = (values.shapeKey === "arrow3")
                ? collectionToArray(activeDocument.selection)
                : null;

            try {
                for (var i = 0; i < items.length - 1; i++) {
                    var markItem = createMarkBetweenItems(items[i], items[i + 1], values);
                    if (markItem) createdItems.push(markItem);
                }
            } finally {
                if (previousSelection) restoreSelection(previousSelection);
            }
            return createdItems;
        }

        // =========================================
        // プレビュー / Preview
        // =========================================

        /**
         * プレビューで作成したアイテムを削除します。
         *
         * @returns {void}
         */
        function removePreview() {
            if (previewItems.length === 0) return;

            for (var i = 0; i < previewItems.length; i++) {
                removeItemSafely(previewItems[i], getLabel("log.removePreviewItem"));
            }
            previewItems = [];
            app.redraw();
        }

        /**
         * プレビューを作り直します。
         *
         * @returns {void}
         */
        function updatePreview() {
            removePreview();
            if (!controls.previewCheckbox.value) return;

            var values = readInputValues(false);
            if (!values) return;

            previewItems = createMarks(values);
            app.redraw();
        }

        // =========================================
        // イベント / Events
        // =========================================

        /**
         * 矢印キーで入力値を増減できるようにします。
         * 基本の増減量は単位に合わせて決まり、Shift はその10倍、Option（Alt）は10分の1で増減します。
         *
         * @param {object} editText - 対象の edittext。
         * @param {boolean} allowNegative - 負の値を許可するなら true。
         * @param {object} [unitInfo] - 単位情報。％や度など単位のない入力欄では省略します。
         * @returns {void}
         */
        function enableArrowKeyInput(editText, allowNegative, unitInfo) {
            editText.addEventListener("keydown", function (event) {
                var value = Number(editText.text);
                if (isNaN(value)) return;

                var isUp = (event.keyName === "Up");
                var isDown = (event.keyName === "Down");
                if (!isUp && !isDown) return;

                var step = getArrowKeyStep(unitInfo);
                var keyboard = ScriptUI.environment.keyboardState;
                if (keyboard.shiftKey) {
                    /* 増減量10個分の倍数へ丸めながら増減 / Snap to multiples of ten steps while stepping */
                    var coarseStep = step * 10;
                    value = isUp
                        ? Math.ceil((value + step) / coarseStep) * coarseStep
                        : Math.floor((value - step) / coarseStep) * coarseStep;
                } else {
                    var fineStep = keyboard.altKey ? step / 10 : step;
                    value = isUp ? value + fineStep : value - fineStep;
                    /* 増減量の刻みに揃える / Snap to the step grid */
                    value = Math.round(value / fineStep) * fineStep;
                }
                event.preventDefault();

                if (!allowNegative && value < 0) value = 0;

                editText.text = String(roundDisplayValueForUnit(value, unitInfo));

                if (editText === controls.heightField && !widthManuallySet) {
                    applyAutoWidthToField();
                }
                if (editText === controls.widthField) {
                    widthManuallySet = true;
                }
                updatePreview();
            });
        }

        /* ％と度は単位を持たないため、増減量は1固定 / Percent and degrees have no unit, so they step by one */
        enableArrowKeyInput(controls.heightField, false);
        enableArrowKeyInput(controls.angleField, false);
        enableArrowKeyInput(controls.widthField, false, rulerUnitInfo);
        enableArrowKeyInput(controls.insetField, false, rulerUnitInfo);
        enableArrowKeyInput(controls.strokeField, false, strokeUnitInfo);
        enableArrowKeyInput(controls.gapField, true, rulerUnitInfo);
        enableArrowKeyInput(controls.offsetXField, true, rulerUnitInfo);
        enableArrowKeyInput(controls.offsetYField, true, rulerUnitInfo);

        controls.heightField.onChanging = function () {
            if (!widthManuallySet) applyAutoWidthToField();
            updatePreview();
        };

        controls.widthField.onChanging = function () {
            if (controls.widthField.text === "") {
                resetWidthToAuto();
            } else {
                widthManuallySet = true;
            }
            updatePreview();
        };

        controls.insetField.onChanging = function () {
            /* 凹みが幅の割合を超えないよう、入力中に抑える / Clamp the inset to a ratio of the width while typing */
            var insetValue = parseFloat(controls.insetField.text);
            if (!isNaN(insetValue) && insetValue >= 0) {
                var effectiveWidthPt = getEffectiveWidthPt();
                var maxInsetPt = effectiveWidthPt * MAX_INSET_RATIO;
                if (effectiveWidthPt > 0 && convertValueToPt(insetValue, rulerUnitInfo) > maxInsetPt) {
                    setFieldFromPt(controls.insetField, maxInsetPt, rulerUnitInfo);
                }
            }
            updatePreview();
        };

        controls.gapField.onChanging = updatePreview;
        controls.strokeField.onChanging = updatePreview;
        controls.angleField.onChanging = updatePreview;
        controls.offsetXField.onChanging = updatePreview;
        controls.offsetYField.onChanging = updatePreview;

        controls.capNoneRadio.onClick = updatePreview;
        controls.capRoundRadio.onClick = updatePreview;
        controls.mirrorCheckbox.onClick = updatePreview;
        controls.alignVerticalCenterCheckbox.onClick = updatePreview;
        controls.roundCornersCheckbox.onClick = updatePreview;
        controls.previewCheckbox.onClick = updatePreview;

        /**
         * 形状を切り替えたときに、初期値と位置調整をリセットします。
         *
         * @returns {void}
         */
        function handleShapeChange() {
            widthManuallySet = false;
            controls.offsetXField.text = "0";
            controls.offsetYField.text = "0";
            applyShapeDefaults();
            updatePreview();
        }

        var shapeRadios = getShapeRadioControls();
        for (var radioIndex = 0; radioIndex < shapeRadios.length; radioIndex++) {
            shapeRadios[radioIndex].onClick = handleShapeChange;
        }

        /* キーボードショートカット / Keyboard shortcuts */
        var SHORTCUT_ACTIONS = {
            F: function () {
                if (!controls.capNoneRadio.enabled) return false;
                controls.capNoneRadio.value = true;
                controls.capRoundRadio.value = false;
                return true;
            },
            R: function () {
                if (!controls.capRoundRadio.enabled) return false;
                controls.capRoundRadio.value = true;
                controls.capNoneRadio.value = false;
                return true;
            },
            V: function () {
                if (!controls.mirrorCheckbox.enabled) return false;
                controls.mirrorCheckbox.value = !controls.mirrorCheckbox.value;
                return true;
            }
        };

        dialog.addEventListener("keydown", function (event) {
            /* Cmd+V などの修飾キー付きの入力は横取りしない / Do not swallow modified keystrokes such as Cmd+V */
            var keyboard = ScriptUI.environment.keyboardState;
            if (keyboard.metaKey || keyboard.ctrlKey || keyboard.altKey || keyboard.shiftKey) return;

            var action = SHORTCUT_ACTIONS[event.keyName];
            if (!action || !action()) return;
            updatePreview();
            event.preventDefault();
        });

        controls.okButton.onClick = function () {
            var values = readInputValues(true);
            if (!values) return;

            /* プレビューがあれば、それをそのまま結果として確定する / When a preview exists, keep it as the result */
            if (previewItems.length > 0) {
                activeDocument.selection = previewItems;
                previewItems = [];
                dialog.close(1);
                return;
            }

            dialog.close(1);

            var createdItems = createMarks(values);
            if (createdItems.length === 0) {
                alert(getLabel("alert.noGap"));
                return;
            }
            activeDocument.selection = createdItems;
        };

        controls.cancelButton.onClick = function () {
            removePreview();
            dialog.close(0);
        };

        dialog.onClose = function () {
            removePreview();
        };

        applyShapeDefaults();

        /* 表示直後にレイアウトを再計算して描画欠けを防ぐ / Recalculate layout on show to avoid partial rendering */
        runSafely(function () {
            dialog.layout.layout(true);
            dialog.layout.resize();
        }, getLabel("log.layoutDialog"));

        dialog.show();

    })();

})();
