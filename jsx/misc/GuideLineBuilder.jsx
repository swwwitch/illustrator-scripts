#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);
#targetengine "ExtendLinesEngine"

/*

### 概要

選択オブジェクト（グループ／複合パス／テキストを含む）から直線セグメントを拾い、描画範囲いっぱいに延長した補助線を描画します。
Bezier曲線から円を推定する［円弧から円］、アンカーポイントに円・正方形を置く機能もあります。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GuideLineBuilder.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nd801b9b0367f

### Overview

Collects the straight segments of the selection — groups, compound paths and text included — and draws each one as a construction line extended across the drawing area.
It can also estimate circles from Bézier curves and place circles or squares on the anchor points.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GuideLineBuilder.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "GuideLineBuilder";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-02-27";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GuideLineBuilder.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GuideLineBuilder.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nd801b9b0367f"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function() {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 生成物の目印（再実行時に削除してよいものを見分ける）/ Marker that flags generated items */
    var SCRIPT_MARKER = "__ExtendLines__";

    /* 補助線を置くレイヤー名 / Layer name for construction lines */
    var LINE_LAYER_NAME = "_construction_guide";

    /* アンカー図形を置くレイヤー名 / Layer name for anchor shapes */
    var ANCHOR_LAYER_NAME = "_construction_anchorpoint";

    /* プレビュー用レイヤー名のもと / Base name of the preview layer */
    var PREVIEW_LAYER_BASE_NAME = "__ExtendLines_Preview";

    /* 線幅の初期値（mm）/ Default stroke width (mm) */
    var DEFAULT_STROKE_WIDTH_MM = 0.1;

    /* アンカー図形の大きさの初期値（mm）/ Default size of anchor shapes (mm) */
    var DEFAULT_ANCHOR_SIZE_MM = 1;

    /* 選択がアートボードの外にあるときの描画範囲の倍率 / Draw area scale when the selection is off the artboard */
    var OFF_ARTBOARD_SCALE = 4;

    /* アンカー図形のブルー [R, G, B] / Blue used for anchor shapes */
    var ANCHOR_BLUE_RGB = [78, 128, 255];

    /* 座標比較の許容値（pt）/ Tolerance for comparing coordinates (pt) */
    var POSITION_TOLERANCE = 0.001;

    /* ズームスライダーの範囲 / Range of the zoom slider */
    var ZOOM_MIN = 0.1;
    var ZOOM_MAX = 16;

    // =========================================
    // レイアウト / Layout
    // =========================================

    var WINDOW_MARGINS          = 15;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING          = 10;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS           = [15, 20, 15, 10];   /* パネル余白 [左,上,右,下] / panel margins */
    var NESTED_PANEL_MARGINS    = [15, 10, 15, 10];   /* 入れ子パネルの余白 [左,上,右,下] / nested panel margins */
    var PANEL_SPACING           = 6;                  /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING          = 10;                 /* 2カラムの間隔 / column spacing */
    var ROW_SPACING             = 6;                  /* 行内の要素間隔 / spacing within a row */
    var RADIO_SPACING           = 10;                 /* 横並びラジオの間隔 / spacing between radios */
    var BUTTON_SPACING          = 10;                 /* ボタンの間隔 / spacing between buttons */
    var BUTTON_ROW_TOP_MARGIN   = 8;                  /* ボタン行の上余白 / top margin of the button row */
    var NUMBER_INPUT_CHARACTERS = 6;                  /* 数値入力欄の文字数 / numeric field width */
    var ZOOM_SLIDER_WIDTH       = 240;                /* ズームスライダーの幅 / zoom slider width */
    var ZOOM_ROW_MARGINS        = [0, 0, 0, 10];      /* ズームの行の余白 / zoom row margins */

    /**
     * ウィンドウの共通レイアウトを設定する
     * @param {Window} targetWindow - 対象のウィンドウ
     * @returns {void}
     */
    function setupWindow(targetWindow) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = ["left", "top"];
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = WINDOW_SPACING;
    }

    /**
     * パネルを追加し、共通レイアウトを設定する
     * @param {object} parentContainer - 追加先のウィンドウまたはグループ
     * @param {object} [titleSet] - ja/en を持つタイトル（省略時はタイトルなしの入れ子パネル）
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parentContainer, titleSet) {
        var newPanel = parentContainer.add("panel", undefined, titleSet ? getLabel(titleSet) : "");
        newPanel.orientation = "column";
        newPanel.alignChildren = ["left", "top"];
        newPanel.alignment = ["fill", "top"];
        newPanel.margins = titleSet ? PANEL_MARGINS : NESTED_PANEL_MARGINS;
        newPanel.spacing = PANEL_SPACING;
        return newPanel;
    }

    /**
     * 行グループを追加し、共通レイアウトを設定する（横位置と天地を対で指定し、子は伸ばさない）
     * @param {object} parentContainer - 追加先のウィンドウまたはグループ
     * @param {number} [spacing] - 要素間隔（省略時は ROW_SPACING）
     * @returns {Group} 追加した行グループ
     */
    function addRow(parentContainer, spacing) {
        var rowGroup = parentContainer.add("group");
        rowGroup.orientation = "row";
        rowGroup.alignment = ["left", "center"];
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.spacing = (typeof spacing === "number") ? spacing : ROW_SPACING;
        return rowGroup;
    }

    /**
     * 数値入力欄に↑↓キーでの増減を付ける（shift 併用で10刻み、option 併用で0.1刻み）
     * @param {EditText} editText - 対象の入力欄
     * @param {boolean} allowNegative - マイナスの値を許可するか
     * @param {function} [onChange] - 値を変えたあとに呼ぶ処理
     * @returns {void}
     */
    function changeValueByArrowKey(editText, allowNegative, onChange) {
        editText.addEventListener("keydown", function(event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;

            var currentValue = Number(editText.text);
            if (isNaN(currentValue)) return;

            var keyboardState = ScriptUI.environment.keyboardState;
            var stepSign = (event.keyName === "Up") ? 1 : -1;
            var nextValue;

            if (keyboardState.shiftKey) {
                /* shift 併用時は10の倍数にスナップ / Snap to multiples of 10 with shift */
                nextValue = (stepSign > 0) ? Math.ceil((currentValue + 1) / 10) * 10 : Math.floor((currentValue - 1) / 10) * 10;
            } else if (keyboardState.altKey) {
                /* option 併用時は0.1刻み / Step by 0.1 with option */
                nextValue = Math.round((currentValue + stepSign * 0.1) * 10) / 10;
            } else {
                nextValue = Math.round(currentValue + stepSign);
            }

            if (!allowNegative && nextValue < 0) nextValue = 0;

            editText.text = String(nextValue);
            event.preventDefault();
            if (typeof onChange === "function") onChange();
        });
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    var uiLang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "補助線の描画", en: "Construction Lines" }
        },
        panel: {
            lines: { ja: "補助線を描画", en: "Draw Construction Lines" },
            anchorShapes: { ja: "アンカーポイントに図形", en: "Shapes on Anchor Points" },
            options: { ja: "オプション", en: "Options" }
        },
        checkbox: {
            straight: { ja: "直線", en: "Straight segments" },
            horizontal: { ja: "水平線", en: "Horizontal lines" },
            vertical: { ja: "垂直線", en: "Vertical lines" },
            diagonal: { ja: "斜線", en: "Diagonal lines" },
            arcToCircle: { ja: "円弧から円", en: "Create circles from arcs" },
            group: { ja: "グループ化", en: "Group output" },
            separateLayer: { ja: "別レイヤーに", en: "Use separate layer" },
            guide: { ja: "ガイド化", en: "Convert to guides" },
            dedup: { ja: "線のダブりを削除", en: "Remove duplicate lines" },
            preview: { ja: "プレビュー", en: "Preview" },
            lightMode: { ja: "軽量モード", en: "Light Preview" }
        },
        radio: {
            anchorNone: { ja: "なし", en: "None" },
            anchorCircle: { ja: "円", en: "Circle" },
            anchorSquare: { ja: "正方形", en: "Square" },
            colorBlack: { ja: "黒", en: "Black" },
            colorBlue: { ja: "ブルー", en: "Blue" },
            arcIgnore: { ja: "無視", en: "Ignore" },
            arcChord: { ja: "弦", en: "Chord" },
            arcExtendChord: { ja: "弦を延長", en: "Extend chord" }
        },
        fieldLabel: {
            strokeWidth: { ja: "線幅", en: "Stroke Width" },
            anchorSize: { ja: "大きさ", en: "Size" },
            zoom: { ja: "ズーム", en: "Zoom" }
        },
        tooltip: {
            arcFallback: { ja: "完全な円弧以外の場合", en: "For segments that are not true circular arcs" },
            straight: {
                ja: "選択したパスの直線セグメントを、描画範囲いっぱいまで延長して描きます",
                en: "Extends each straight segment of the selection across the drawing area"
            },
            soloDirection: {
                ja: "option キーを押しながらクリックすると、この向きだけを残します",
                en: "Option-click to keep only this direction"
            },
            arcToCircle: {
                ja: "円弧とみなせる曲線セグメントから、その円を描きます",
                en: "Draws the full circle of each curved segment that is a circular arc"
            },
            arcIgnore: { ja: "完全な円弧でないセグメントには何も描きません", en: "Draws nothing for segments that are not true arcs" },
            arcChord: {
                ja: "完全な円弧でないセグメントは、両端を結ぶ線分を描きます",
                en: "Draws the chord between the end points of segments that are not true arcs"
            },
            arcExtendChord: {
                ja: "完全な円弧でないセグメントは、両端を結ぶ直線を描画範囲いっぱいまで延長して描きます",
                en: "Extends the chord of segments that are not true arcs across the drawing area"
            },
            group: {
                ja: "補助線とアンカーポイントの図形を、それぞれグループにまとめます",
                en: "Groups the construction lines and the anchor shapes separately"
            },
            separateLayer: {
                ja: "補助線は「_construction_guide」、図形は「_construction_anchorpoint」レイヤーに描き、前回このスクリプトで描いたものは置き換えます",
                en: "Draws lines on \"_construction_guide\" and shapes on \"_construction_anchorpoint\", replacing what this script drew there before"
            },
            guide: {
                ja: "補助線と円をガイドにします（アンカーポイントの図形はガイドにしません）",
                en: "Turns the lines and circles into guides (anchor shapes stay as paths)"
            },
            dedup: {
                ja: "延長した補助線が同じ位置に重なるときは1本だけ描きます",
                en: "Draws only one line when extended lines fall in the same place"
            },
            zoom: {
                ja: "作業中の表示倍率を変えます。option キーでアートボードの中央、shift キーで選択の中央に合わせます",
                en: "Changes the view zoom. Hold Option to center on the artboard, or Shift to center on the selection"
            },
            lightMode: {
                ja: "ズームのスライダーを離したときだけ表示倍率を変えます",
                en: "Applies the zoom only when you release the slider"
            }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        itemName: {
            history: { ja: "補助線の描画", en: "Construction Lines" },
            lineGroup: { ja: "補助線", en: "Construction Lines" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "オブジェクトを選択してください。", en: "Please select objects." },
            noValidPath: { ja: "有効なパスが見つかりません。", en: "No valid paths were found." },
            error: { ja: "エラー: ", en: "Error: " }
        }
    };

    /**
     * ラベル（ja/en）を現在の UI 言語の文字列にする
     * @param {object} labelSet - ja/en を持つラベル
     * @returns {string} 現在の言語の文字列
     */
    function getLabel(labelSet) {
        return (labelSet && labelSet[uiLang]) || "";
    }

    /**
     * 項目名にコロンを付ける（日本語は全角、英語は半角）
     * @param {object} labelSet - ja/en を持つラベル
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // 定数 / Constants
    // =========================================

    /* アンカーポイントに置く図形 / Shapes drawn on anchor points */
    var ANCHOR_SHAPE = {
        NONE: "NONE",
        CIRCLE: "CIRCLE",
        SQUARE: "SQUARE"
    };

    /* アンカー図形の色 / Color of anchor shapes */
    var ANCHOR_COLOR = {
        BLACK: "BLACK",
        BLUE: "BLUE"
    };

    /* 完全な円弧として扱えないカーブの処理 / Fallback for curves that are not true arcs */
    var ARC_FALLBACK = {
        IGNORE: "IGNORE",
        CHORD: "CHORD",
        EXTEND: "EXTEND"
    };

    /* 直線セグメントの向き / Direction of a straight segment */
    var SEGMENT_DIRECTION = {
        HORIZONTAL: "HORIZONTAL",
        VERTICAL: "VERTICAL",
        DIAGONAL: "DIAGONAL"
    };

    // =========================================
    // セッション記憶 / Session state
    // =========================================

    /* 常駐エンジンが生きているあいだだけ保持する（Illustrator終了で破棄）/ Kept while the engine lives */
    var SESSION_GLOBAL_KEY = "__GuideLineBuilderSession";

    /**
     * セッション記憶を取り出す（無ければ作る）
     * @returns {object} セッション記憶
     */
    function getSession() {
        if (!$.global[SESSION_GLOBAL_KEY]) $.global[SESSION_GLOBAL_KEY] = {};
        return $.global[SESSION_GLOBAL_KEY];
    }

    // =========================================
    // ユーティリティ / Utilities
    // =========================================

    /**
     * エラーを ExtendScript コンソールへ書き出す
     * @param {Error} err - 捕まえた例外
     * @param {string} context - 発生箇所を示す文字列
     * @returns {void}
     */
    function logError(err, context) {
        $.writeln("[" + SCRIPT_NAME + "] " + context + ": " + err);
    }

    /**
     * mm を pt に変換する
     * @param {number} millimeters - ミリメートル値
     * @returns {number} ポイント値
     */
    function mmToPt(millimeters) {
        return millimeters * 72.0 / 25.4;
    }

    /**
     * 2点がほぼ同じ位置かどうかを判定する
     * @param {Array<number>} pointA - 座標 [x, y]
     * @param {Array<number>} pointB - 座標 [x, y]
     * @returns {boolean} ほぼ同じ位置なら true
     */
    function isSamePosition(pointA, pointB) {
        return Math.abs(pointA[0] - pointB[0]) < POSITION_TOLERANCE &&
            Math.abs(pointA[1] - pointB[1]) < POSITION_TOLERANCE;
    }

    /**
     * 座標を丸めて重複判定用のキーにする
     * @param {Array<number>} point - 座標 [x, y]
     * @returns {string} 重複判定キー
     */
    function makePointKey(point) {
        return Number(point[0]).toFixed(3) + "," + Number(point[1]).toFixed(3);
    }

    /**
     * 2点から線分の重複判定キーを作る（向きは無視）
     * @param {Array<number>} pointA - 座標 [x, y]
     * @param {Array<number>} pointB - 座標 [x, y]
     * @returns {string} 重複判定キー
     */
    function makeLineKey(pointA, pointB) {
        var keyA = makePointKey(pointA);
        var keyB = makePointKey(pointB);
        return (keyA < keyB) ? (keyA + "|" + keyB) : (keyB + "|" + keyA);
    }

    /**
     * 入力文字列を数値に直す（全角小数点・カンマ区切りを吸収する）
     * @param {string} text - 入力欄の文字列
     * @returns {number} 数値（読めない場合は NaN）
     */
    function parseNumberInput(text) {
        var normalized = String(text).replace(/\s+/g, "").replace(/，/g, ",").replace(/．/g, ".");
        /* "1,5" は小数点、"1,000" は桁区切りとみなす / Treat "1,5" as a decimal and "1,000" as a separator */
        normalized = (normalized.indexOf(",") >= 0 && normalized.indexOf(".") < 0) ?
            normalized.replace(/,/g, ".") :
            normalized.replace(/,/g, "");
        return Number(normalized);
    }

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

    /**
     * 入力欄の文字列を pt に直す（読めない場合は直前の有効値を返す）
     * @param {EditText} editText - 対象の入力欄
     * @param {string} prefKey - 単位を決める環境設定キー（strokeUnits / rulerType）
     * @param {object} lastValidCache - 直前の有効値を持つ `{ value: number }`
     * @returns {number} ポイント値
     */
    function readLengthAsPt(editText, prefKey, lastValidCache) {
        var pointsPerUnit = getUnitInfo(prefKey).pointsPerUnit;
        var inputValue = parseNumberInput(editText.text);
        if (!(inputValue > 0)) return lastValidCache.value;

        var lengthPt = inputValue * pointsPerUnit;
        if (!(lengthPt > 0)) return lastValidCache.value;

        lastValidCache.value = lengthPt;
        return lengthPt;
    }

    // =========================================
    // レイヤー / Layers
    // =========================================

    /**
     * 名前でレイヤーを探す
     * @param {Document} doc - 対象のドキュメント
     * @param {string} layerName - レイヤー名
     * @returns {Layer|null} 見つかったレイヤー（無ければ null）
     */
    function findLayerByName(doc, layerName) {
        try { return doc.layers.getByName(layerName); } catch (e) { }
        return null;
    }

    /**
     * 既存と重ならないレイヤー名を作る
     * @param {Document} doc - 対象のドキュメント
     * @param {string} baseName - もとにする名前
     * @returns {string} 重複しないレイヤー名
     */
    function createUniqueLayerName(doc, baseName) {
        var candidateName = baseName;
        var suffixNumber = 2;
        while (findLayerByName(doc, candidateName)) {
            candidateName = baseName + "_" + suffixNumber;
            suffixNumber++;
        }
        return candidateName;
    }

    /**
     * 名前が一致するレイヤーがあれば削除する
     * @param {Document} doc - 対象のドキュメント
     * @param {string} layerName - レイヤー名
     * @returns {void}
     */
    function removeLayerIfExists(doc, layerName) {
        var targetLayer = findLayerByName(doc, layerName);
        if (!targetLayer) return;
        /* ロックなどで削除できないことがある / Removal can fail, e.g. on a locked layer */
        try { targetLayer.remove(); } catch (e) { logError(e, "removeLayerIfExists"); }
    }

    /**
     * アイテムがこのスクリプトの生成物かどうかを判定する
     * @param {PageItem} item - 対象のアイテム
     * @returns {boolean} 生成物なら true
     */
    function isGeneratedItem(item) {
        /* note / name を読めないアイテムは生成物ではないとみなす / Treat unreadable items as not generated */
        try {
            if (item.note === SCRIPT_MARKER) return true;
            return String(item.name || "").indexOf(SCRIPT_MARKER) === 0;
        } catch (e) { }
        return false;
    }

    /**
     * このスクリプトが生成したオブジェクトだけをレイヤーから削除する
     * @param {Layer} layer - 対象のレイヤー
     * @returns {void}
     */
    function clearGeneratedItemsInLayer(layer) {
        if (!layer) return;
        /* ロックされたアイテムなどは削除できない / Locked items cannot be removed */
        try {
            for (var i = layer.pageItems.length - 1; i >= 0; i--) {
                var layerItem = layer.pageItems[i];
                if (layerItem && isGeneratedItem(layerItem)) layerItem.remove();
            }
        } catch (e) { logError(e, "clearGeneratedItemsInLayer"); }
    }

    /**
     * 選択がすべて指定レイヤー上にあるかどうかを判定する
     * @param {Array} selectedItems - 選択アイテム
     * @param {string} layerName - レイヤー名
     * @returns {boolean} すべてそのレイヤー上なら true
     */
    function isSelectionOnLayer(selectedItems, layerName) {
        if (!selectedItems || selectedItems.length === 0) return false;
        for (var i = 0; i < selectedItems.length; i++) {
            var itemLayer = null;
            /* layer を持たないアイテムがある / Some items have no layer */
            try { itemLayer = selectedItems[i].layer; } catch (e) { }
            if (!itemLayer || itemLayer.name !== layerName) return false;
        }
        return true;
    }

    /**
     * 補助線の描画先レイヤーを用意する（再実行時は自分の生成物だけ消す）
     * @param {Document} doc - 対象のドキュメント
     * @param {Array} selectedItems - 選択アイテム
     * @returns {Layer} 描画先のレイヤー
     */
    function prepareLineLayer(doc, selectedItems) {
        var lineLayer;

        if (isSelectionOnLayer(selectedItems, LINE_LAYER_NAME)) {
            /* 選択自体が対象レイヤー上にあるので、消さずに退避して新しく作る / Keep the selection by backing up the layer */
            var existingLayer = findLayerByName(doc, LINE_LAYER_NAME);
            if (existingLayer) existingLayer.name = createUniqueLayerName(doc, LINE_LAYER_NAME + "_backup");
            lineLayer = doc.layers.add();
            lineLayer.name = LINE_LAYER_NAME;
        } else {
            lineLayer = findLayerByName(doc, LINE_LAYER_NAME);
            if (lineLayer) {
                clearGeneratedItemsInLayer(lineLayer);
            } else {
                lineLayer = doc.layers.add();
                lineLayer.name = LINE_LAYER_NAME;
            }
        }

        lineLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
        return lineLayer;
    }

    /**
     * アンカー図形の描画先レイヤーを用意する（図形を作らないときも前回の生成物は消す）
     * @param {Document} doc - 対象のドキュメント
     * @param {boolean} needsLayer - 今回アンカー図形を作るかどうか
     * @returns {Layer|null} 描画先のレイヤー（作らない場合は null）
     */
    function prepareAnchorLayer(doc, needsLayer) {
        var anchorLayer = findLayerByName(doc, ANCHOR_LAYER_NAME);
        if (anchorLayer) clearGeneratedItemsInLayer(anchorLayer);

        if (!needsLayer) return null;

        if (!anchorLayer) {
            anchorLayer = doc.layers.add();
            anchorLayer.name = ANCHOR_LAYER_NAME;
        }
        anchorLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
        return anchorLayer;
    }

    /**
     * 目印付きのグループを追加する
     * @param {Layer|GroupItem} parentContainer - 追加先のレイヤーまたはグループ
     * @param {string} groupName - グループ名
     * @returns {GroupItem} 追加したグループ
     */
    function addMarkedGroup(parentContainer, groupName) {
        var markedGroup = parentContainer.groupItems.add();
        markedGroup.name = groupName;
        markedGroup.note = SCRIPT_MARKER;
        return markedGroup;
    }

    // =========================================
    // パスの収集 / Collecting paths
    // =========================================

    /**
     * グループや複合パスの中から再帰的にパスアイテムを集める
     * @param {Array} items - 走査するアイテム
     * @param {Array<PathItem>} collected - 集めたパスの入れ物
     * @returns {void}
     */
    function extractPathItems(items, collected) {
        for (var i = 0; i < items.length; i++) {
            var pageItem = items[i];
            if (pageItem.typename === "PathItem") {
                collected.push(pageItem);
            } else if (pageItem.typename === "CompoundPathItem") {
                extractPathItems(pageItem.pathItems, collected);
            } else if (pageItem.typename === "GroupItem") {
                extractPathItems(pageItem.pageItems, collected);
            }
        }
    }

    /**
     * 選択内のテキストを複製してアウトライン化する（元のテキストは触らない）
     * @param {Array} items - 選択アイテム
     * @returns {Array<PageItem>} 一時的に作ったアウトラインの入れ物
     */
    function outlineTextFromSelection(items) {
        var outlineRoots = [];

        /**
         * アイテムをたどってテキストだけアウトライン化する
         * @param {PageItem} pageItem - 対象のアイテム
         * @returns {void}
         */
        function outlineTextInItem(pageItem) {
            if (!pageItem) return;

            if (pageItem.typename === "TextFrame") {
                try {
                    /* 同じレイヤーの末尾に複製し、複製だけをアウトライン化する / Duplicate first, outline the copy only */
                    var textCopy = pageItem.duplicate(pageItem.layer, ElementPlacement.PLACEATEND);
                    var outlineGroup = textCopy.createOutline();
                    /* createOutline() は複製を消費するので、ふつうは例外になる / createOutline() consumes the copy, so this usually throws */
                    try { textCopy.remove(); } catch (e) { }
                    if (outlineGroup) outlineRoots.push(outlineGroup);
                } catch (e) {
                    /* 1つ失敗しても全体は止めない / Keep going even if one frame fails */
                    logError(e, "outlineTextFromSelection/createOutline");
                }
                return;
            }

            if (pageItem.typename === "GroupItem") {
                for (var i = 0; i < pageItem.pageItems.length; i++) outlineTextInItem(pageItem.pageItems[i]);
            }
        }

        for (var i = 0; i < items.length; i++) outlineTextInItem(items[i]);

        return outlineRoots;
    }

    /**
     * 一時的に作ったアウトラインを削除する
     * @param {Array<PageItem>} outlineRoots - 削除するアイテム
     * @returns {void}
     */
    function cleanupTempOutlines(outlineRoots) {
        if (!outlineRoots) return;
        for (var i = outlineRoots.length - 1; i >= 0; i--) {
            /* すでに消えているものは飛ばす / Skip items that are already gone */
            try { outlineRoots[i].remove(); } catch (e) { }
        }
    }

    // =========================================
    // ジオメトリ / Geometry
    // =========================================

    /**
     * 選択アイテム全体の外接バウンディングを求める
     * @param {Array} items - 対象のアイテム
     * @returns {Array<number>|null} [左, 上, 右, 下]（求められない場合は null）
     */
    function getUnionBounds(items) {
        var unionBounds = null;
        for (var i = 0; i < items.length; i++) {
            /* 空白だけの文字グループなどは geometricBounds で例外になる / geometricBounds throws for e.g. blank text groups */
            try {
                var itemBounds = items[i].geometricBounds;
                if (!unionBounds) {
                    unionBounds = [itemBounds[0], itemBounds[1], itemBounds[2], itemBounds[3]];
                } else {
                    if (itemBounds[0] < unionBounds[0]) unionBounds[0] = itemBounds[0];
                    if (itemBounds[1] > unionBounds[1]) unionBounds[1] = itemBounds[1];
                    if (itemBounds[2] > unionBounds[2]) unionBounds[2] = itemBounds[2];
                    if (itemBounds[3] < unionBounds[3]) unionBounds[3] = itemBounds[3];
                }
            } catch (e) { logError(e, "getUnionBounds/item[" + i + "]"); }
        }
        return unionBounds;
    }

    /**
     * @typedef {object} DrawArea
     * @property {number} left - 左端
     * @property {number} top - 上端
     * @property {number} right - 右端
     * @property {number} bottom - 下端
     */

    /**
     * 補助線を伸ばす範囲を求める（通常はアートボード、選択がアートボードの外なら選択中心の矩形）
     * @param {Document} doc - 対象のドキュメント
     * @param {Array} selectedItems - 選択アイテム
     * @returns {DrawArea} 描画範囲
     */
    function getDrawArea(doc, selectedItems) {
        var artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
        var artboardArea = {
            left: artboardRect[0],
            top: artboardRect[1],
            right: artboardRect[2],
            bottom: artboardRect[3]
        };

        var selectionBounds = getUnionBounds(selectedItems);
        if (!selectionBounds) return artboardArea;

        var selectionLeft = selectionBounds[0], selectionTop = selectionBounds[1];
        var selectionRight = selectionBounds[2], selectionBottom = selectionBounds[3];

        /* Illustrator 座標は上が大きい / In Illustrator coordinates, top is the larger Y */
        var intersects = !(selectionRight < artboardArea.left || selectionLeft > artboardArea.right ||
            selectionTop < artboardArea.bottom || selectionBottom > artboardArea.top);
        if (intersects) return artboardArea;

        /* 選択がアートボードと全く重ならないときは、選択を中心にした矩形を使う / Use a rect around the selection instead */
        var areaWidth = Math.max(selectionRight - selectionLeft, 1) * OFF_ARTBOARD_SCALE;
        var areaHeight = Math.max(selectionTop - selectionBottom, 1) * OFF_ARTBOARD_SCALE;
        var centerX = (selectionLeft + selectionRight) / 2;
        var centerY = (selectionTop + selectionBottom) / 2;

        return {
            left: centerX - areaWidth / 2,
            top: centerY + areaHeight / 2,
            right: centerX + areaWidth / 2,
            bottom: centerY - areaHeight / 2
        };
    }

    /**
     * 2つのアンカーポイント間が直線（ハンドルが出ていない）かどうかを判定する
     * @param {PathPoint} startPoint - 始点のアンカーポイント
     * @param {PathPoint} endPoint - 終点のアンカーポイント
     * @returns {boolean} 直線なら true
     */
    function isStraightSegment(startPoint, endPoint) {
        return isSamePosition(startPoint.rightDirection, startPoint.anchor) &&
            isSamePosition(endPoint.leftDirection, endPoint.anchor);
    }

    /**
     * 直線セグメントの向きを分類する
     * @param {Array<number>} start - 始点の座標 [x, y]
     * @param {Array<number>} end - 終点の座標 [x, y]
     * @returns {string} SEGMENT_DIRECTION のいずれか
     */
    function classifySegmentDirection(start, end) {
        if (Math.abs(start[1] - end[1]) < POSITION_TOLERANCE) return SEGMENT_DIRECTION.HORIZONTAL;
        if (Math.abs(start[0] - end[0]) < POSITION_TOLERANCE) return SEGMENT_DIRECTION.VERTICAL;
        return SEGMENT_DIRECTION.DIAGONAL;
    }

    /**
     * ベクトルの長さを求める
     * @param {Array<number>} vector - ベクトル [x, y]
     * @returns {number} 長さ
     */
    function vectorLength(vector) {
        return Math.sqrt(vector[0] * vector[0] + vector[1] * vector[1]);
    }

    /**
     * ベクトルの差を求める
     * @param {Array<number>} pointA - 座標 [x, y]
     * @param {Array<number>} pointB - 座標 [x, y]
     * @returns {Array<number>} pointA - pointB
     */
    function subtractVector(pointA, pointB) {
        return [pointA[0] - pointB[0], pointA[1] - pointB[1]];
    }

    /**
     * ベクトルの内積を求める
     * @param {Array<number>} vectorA - ベクトル [x, y]
     * @param {Array<number>} vectorB - ベクトル [x, y]
     * @returns {number} 内積
     */
    function dotProduct(vectorA, vectorB) {
        return vectorA[0] * vectorB[0] + vectorA[1] * vectorB[1];
    }

    /**
     * ベクトルを90度回す
     * @param {Array<number>} vector - ベクトル [x, y]
     * @returns {Array<number>} 直交するベクトル
     */
    function rotate90(vector) {
        return [-vector[1], vector[0]];
    }

    /**
     * 点と方向で表した2直線の交点を求める
     * @param {Array<number>} pointA - 直線1が通る点 [x, y]
     * @param {Array<number>} directionA - 直線1の方向 [x, y]
     * @param {Array<number>} pointB - 直線2が通る点 [x, y]
     * @param {Array<number>} directionB - 直線2の方向 [x, y]
     * @returns {Array<number>|null} 交点の座標（平行な場合は null）
     */
    function intersectLines(pointA, directionA, pointB, directionB) {
        var denominator = directionA[0] * directionB[1] - directionA[1] * directionB[0];
        if (Math.abs(denominator) < 1e-9) return null;

        var deltaX = pointB[0] - pointA[0];
        var deltaY = pointB[1] - pointA[1];
        var ratio = (deltaX * directionB[1] - deltaY * directionB[0]) / denominator;

        return [pointA[0] + directionA[0] * ratio, pointA[1] + directionA[1] * ratio];
    }

    /**
     * 3次ベジェ曲線上の点を求める
     * @param {Array<number>} start - 始点 [x, y]
     * @param {Array<number>} startHandle - 始点のハンドル [x, y]
     * @param {Array<number>} endHandle - 終点のハンドル [x, y]
     * @param {Array<number>} end - 終点 [x, y]
     * @param {number} t - パラメーター（0〜1）
     * @returns {Array<number>} 曲線上の座標 [x, y]
     */
    function cubicBezierPoint(start, startHandle, endHandle, end, t) {
        var u = 1 - t;
        var uuu = u * u * u;
        var uut = 3 * u * u * t;
        var utt = 3 * u * t * t;
        var ttt = t * t * t;

        return [
            uuu * start[0] + uut * startHandle[0] + utt * endHandle[0] + ttt * end[0],
            uuu * start[1] + uut * startHandle[1] + utt * endHandle[1] + ttt * end[1]
        ];
    }

    /**
     * 円弧の両端の接線から円の中心を求める
     * @param {Array<number>} start - 始点 [x, y]
     * @param {Array<number>} startHandle - 始点のハンドル [x, y]
     * @param {Array<number>} end - 終点 [x, y]
     * @param {Array<number>} endHandle - 終点のハンドル [x, y]
     * @returns {Array<number>|null} 中心の座標（求められない場合は null）
     */
    function getArcCenter(start, startHandle, end, endHandle) {
        var startTangent = subtractVector(startHandle, start);
        var endTangent = subtractVector(end, endHandle);
        return intersectLines(start, rotate90(startTangent), end, rotate90(endTangent));
    }

    /**
     * 曲線セグメントを円弧とみなしてよいかを判定する（両端と中点が同じ円上にあり、接線が半径と直交する）
     * @param {Array<number>} start - 始点 [x, y]
     * @param {Array<number>} startHandle - 始点のハンドル [x, y]
     * @param {Array<number>} endHandle - 終点のハンドル [x, y]
     * @param {Array<number>} end - 終点 [x, y]
     * @param {Array<number>} center - 推定した中心 [x, y]
     * @param {number} radius - 推定した半径
     * @returns {boolean} 円弧とみなせるなら true
     */
    function isApproxCircularArc(start, startHandle, endHandle, end, center, radius) {
        if (!center || !(radius > 0)) return false;

        /* 半径の1%、ただし最低 0.2pt を許容値にする / Tolerance: 1% of the radius, at least 0.2pt */
        var tolerance = Math.max(0.2, radius * 0.01);

        if (Math.abs(vectorLength(subtractVector(start, center)) - radius) > tolerance) return false;
        if (Math.abs(vectorLength(subtractVector(end, center)) - radius) > tolerance) return false;

        var midPoint = cubicBezierPoint(start, startHandle, endHandle, end, 0.5);
        if (Math.abs(vectorLength(subtractVector(midPoint, center)) - radius) > tolerance) return false;

        var startTangent = subtractVector(startHandle, start);
        var endTangent = subtractVector(end, endHandle);
        if (vectorLength(startTangent) < 1e-6 || vectorLength(endTangent) < 1e-6) return false;

        if (Math.abs(dotProduct(subtractVector(start, center), startTangent)) > tolerance * vectorLength(startTangent)) return false;
        if (Math.abs(dotProduct(subtractVector(end, center), endTangent)) > tolerance * vectorLength(endTangent)) return false;

        return true;
    }

    // =========================================
    // 描画 / Drawing
    // =========================================

    /**
     * K100 のカラーを作る
     * @returns {CMYKColor} 黒100%のカラー
     */
    function createBlackColor() {
        var blackColor = new CMYKColor();
        blackColor.cyan = 0;
        blackColor.magenta = 0;
        blackColor.yellow = 0;
        blackColor.black = 100;
        return blackColor;
    }

    /**
     * 補助線のスタイルを適用する（塗りなし、ガイド化またはK100の線）
     * @param {PathItem} pathItem - 対象のパス
     * @param {object} settings - ダイアログの設定
     * @returns {void}
     */
    function applyLineStyle(pathItem, settings) {
        pathItem.note = SCRIPT_MARKER;
        pathItem.filled = false;
        pathItem.fillColor = new NoColor();

        if (settings.guide) {
            pathItem.stroked = false;
            pathItem.guides = true;
            return;
        }

        pathItem.stroked = true;
        pathItem.strokeColor = createBlackColor();
        pathItem.strokeWidth = settings.strokeWidthPt;
    }

    /**
     * アンカー図形のスタイルを適用する（塗りのみ、線なし。ガイド化はしない）
     * @param {PathItem} pathItem - 対象のパス
     * @param {string} anchorColor - ANCHOR_COLOR のいずれか
     * @returns {void}
     */
    function applyAnchorShapeStyle(pathItem, anchorColor) {
        pathItem.note = SCRIPT_MARKER;
        pathItem.filled = true;

        if (anchorColor === ANCHOR_COLOR.BLUE) {
            var blueColor = new RGBColor();
            blueColor.red = ANCHOR_BLUE_RGB[0];
            blueColor.green = ANCHOR_BLUE_RGB[1];
            blueColor.blue = ANCHOR_BLUE_RGB[2];
            pathItem.fillColor = blueColor;
        } else {
            pathItem.fillColor = createBlackColor();
        }

        pathItem.stroked = false;
        pathItem.strokeColor = new NoColor();
        pathItem.guides = false;
    }

    /**
     * アンカーポイントの位置に図形（円または正方形）を作る
     * @param {Array<PathItem>} pathItems - 対象のパス
     * @param {Layer|GroupItem} container - 図形の追加先
     * @param {object} settings - ダイアログの設定
     * @returns {void}
     */
    function createAnchorShapes(pathItems, container, settings) {
        var shapeSize = settings.anchorSizePt;
        var halfSize = shapeSize / 2;
        var placedPointKeys = {};

        for (var i = 0; i < pathItems.length; i++) {
            var pathPoints = pathItems[i].pathPoints;
            if (!pathPoints) continue;

            for (var j = 0; j < pathPoints.length; j++) {
                var anchorPosition = pathPoints[j].anchor;
                if (!anchorPosition) continue;

                /* 同じ位置に重ねて作らない / Do not stack shapes on the same position */
                var pointKey = makePointKey(anchorPosition);
                if (placedPointKeys[pointKey]) continue;
                placedPointKeys[pointKey] = true;

                var shapeTop = anchorPosition[1] + halfSize;
                var shapeLeft = anchorPosition[0] - halfSize;
                var anchorShapeItem = (settings.anchorShape === ANCHOR_SHAPE.CIRCLE) ?
                    container.pathItems.ellipse(shapeTop, shapeLeft, shapeSize, shapeSize) :
                    container.pathItems.rectangle(shapeTop, shapeLeft, shapeSize, shapeSize);

                anchorShapeItem.closed = true;
                applyAnchorShapeStyle(anchorShapeItem, settings.anchorColor);
            }
        }
    }

    /**
     * 2点を結ぶ開いたパス（弦など）を描いて補助線のスタイルを適用する
     * @param {Layer|GroupItem} container - 線の追加先
     * @param {Array<number>} start - 始点 [x, y]
     * @param {Array<number>} end - 終点 [x, y]
     * @param {object} settings - ダイアログの設定
     * @returns {void}
     */
    function drawOpenLine(container, start, end, settings) {
        var openLine = container.pathItems.add();
        openLine.setEntirePath([start, end]);
        openLine.closed = false;
        applyLineStyle(openLine, settings);
    }

    /**
     * 2点を通る直線と描画範囲との交点を求める
     * @param {Array<number>} start - 始点 [x, y]
     * @param {Array<number>} end - 終点 [x, y]
     * @param {DrawArea} area - 描画範囲
     * @returns {Array<Array<number>>} 交点の配列
     */
    function getDrawAreaIntersections(start, end, area) {
        var intersections = [];

        if (Math.abs(start[0] - end[0]) < POSITION_TOLERANCE) {
            /* 垂直線 / Vertical line */
            intersections.push([start[0], area.top]);
            intersections.push([start[0], area.bottom]);
            return intersections;
        }

        if (Math.abs(start[1] - end[1]) < POSITION_TOLERANCE) {
            /* 水平線 / Horizontal line */
            intersections.push([area.left, start[1]]);
            intersections.push([area.right, start[1]]);
            return intersections;
        }

        var slope = (end[1] - start[1]) / (end[0] - start[0]);
        var intercept = start[1] - slope * start[0];

        var yAtLeft = slope * area.left + intercept;
        if (yAtLeft <= area.top + POSITION_TOLERANCE && yAtLeft >= area.bottom - POSITION_TOLERANCE) {
            intersections.push([area.left, yAtLeft]);
        }

        var yAtRight = slope * area.right + intercept;
        if (yAtRight <= area.top + POSITION_TOLERANCE && yAtRight >= area.bottom - POSITION_TOLERANCE) {
            intersections.push([area.right, yAtRight]);
        }

        var xAtTop = (area.top - intercept) / slope;
        if (xAtTop >= area.left - POSITION_TOLERANCE && xAtTop <= area.right + POSITION_TOLERANCE) {
            intersections.push([xAtTop, area.top]);
        }

        var xAtBottom = (area.bottom - intercept) / slope;
        if (xAtBottom >= area.left - POSITION_TOLERANCE && xAtBottom <= area.right + POSITION_TOLERANCE) {
            intersections.push([xAtBottom, area.bottom]);
        }

        return intersections;
    }

    /**
     * 2点を通る直線を描画範囲いっぱいまで延長して描く
     * @param {Layer|GroupItem} container - 線の追加先
     * @param {Array<number>} start - 始点 [x, y]
     * @param {Array<number>} end - 終点 [x, y]
     * @param {DrawArea} area - 描画範囲
     * @param {object} settings - ダイアログの設定
     * @param {object} dedupMap - 重複判定用のキー置き場
     * @returns {void}
     */
    function drawLineAcrossDrawArea(container, start, end, area, settings, dedupMap) {
        var intersections = getDrawAreaIntersections(start, end, area);
        if (intersections.length < 2) return;

        /* 同じ位置の交点（角をかすめた場合など）は1つにまとめる / Collapse duplicate hits, e.g. at a corner */
        var lineStart = intersections[0];
        var lineEnd = null;
        for (var i = 1; i < intersections.length; i++) {
            if (!isSamePosition(intersections[i], lineStart)) {
                lineEnd = intersections[i];
                break;
            }
        }
        if (!lineEnd) return;

        if (settings.dedup) {
            var lineKey = makeLineKey(lineStart, lineEnd);
            if (dedupMap[lineKey]) return;
            dedupMap[lineKey] = true;
        }

        drawOpenLine(container, lineStart, lineEnd, settings);
    }

    /**
     * パスのセグメント数を返す（3点以上の閉じたパスは最後の点から最初の点へのセグメントも数える）
     * @param {PathItem} pathItem - 対象のパス
     * @returns {number} セグメント数
     */
    function getSegmentCount(pathItem) {
        var pointCount = pathItem.pathPoints.length;
        return (pathItem.closed && pointCount >= 3) ? pointCount : pointCount - 1;
    }

    /**
     * パスの中から曲線セグメント（どちらかにハンドルが出ているセグメント）を集める
     * @param {PathItem} pathItem - 対象のパス
     * @returns {Array<object>} `{ start, end }` の配列
     */
    function getCurvedSegments(pathItem) {
        var pathPoints = pathItem.pathPoints;
        var segmentCount = getSegmentCount(pathItem);
        var segments = [];

        for (var i = 0; i < segmentCount; i++) {
            var startPoint = pathPoints[i];
            var endPoint = pathPoints[(i + 1) % pathPoints.length];
            if (!isStraightSegment(startPoint, endPoint)) {
                segments.push({ start: startPoint, end: endPoint });
            }
        }

        return segments;
    }

    /**
     * 曲線セグメントから円を推定して描く（円弧とみなせない場合は設定に従う）
     * @param {PathItem} pathItem - 対象のパス
     * @param {Layer|GroupItem} container - 図形の追加先
     * @param {object} settings - ダイアログの設定
     * @param {DrawArea} area - 描画範囲
     * @param {object} dedupMap - 重複判定用のキー置き場
     * @returns {void}
     */
    function createCirclesFromArcPath(pathItem, container, settings, area, dedupMap) {
        var segments = getCurvedSegments(pathItem);

        for (var i = 0; i < segments.length; i++) {
            var start = segments[i].start.anchor;
            var end = segments[i].end.anchor;
            var startHandle = segments[i].start.rightDirection;
            var endHandle = segments[i].end.leftDirection;

            var center = getArcCenter(start, startHandle, end, endHandle);
            if (!center) continue;

            var radius = vectorLength(subtractVector(start, center));
            if (!(radius > 0)) continue;

            if (!isApproxCircularArc(start, startHandle, endHandle, end, center, radius)) {
                if (settings.arcFallback === ARC_FALLBACK.CHORD) {
                    drawOpenLine(container, start, end, settings);
                } else if (settings.arcFallback === ARC_FALLBACK.EXTEND) {
                    drawLineAcrossDrawArea(container, start, end, area, settings, dedupMap);
                }
                /* IGNORE は何もしない / IGNORE draws nothing */
                continue;
            }

            var circle = container.pathItems.ellipse(
                center[1] + radius,
                center[0] - radius,
                radius * 2,
                radius * 2
            );
            applyLineStyle(circle, settings);
        }
    }

    /**
     * 直線セグメントの向きが描画対象に選ばれているかを判定する
     * @param {string} direction - SEGMENT_DIRECTION のいずれか
     * @param {object} settings - ダイアログの設定
     * @returns {boolean} 描画する向きなら true
     */
    function isDirectionEnabled(direction, settings) {
        if (direction === SEGMENT_DIRECTION.HORIZONTAL) return settings.horizontal;
        if (direction === SEGMENT_DIRECTION.VERTICAL) return settings.vertical;
        return settings.diagonal;
    }

    /**
     * パスの直線セグメントを描画範囲いっぱいまで延長して描く
     * @param {PathItem} pathItem - 対象のパス
     * @param {Layer|GroupItem} container - 線の追加先
     * @param {object} settings - ダイアログの設定
     * @param {DrawArea} area - 描画範囲
     * @param {object} dedupMap - 重複判定用のキー置き場
     * @returns {void}
     */
    function drawExtensionsFromPath(pathItem, container, settings, area, dedupMap) {
        var pathPoints = pathItem.pathPoints;
        if (!pathPoints || pathPoints.length < 2) return;

        var segmentCount = getSegmentCount(pathItem);

        for (var i = 0; i < segmentCount; i++) {
            var startPoint = pathPoints[i];
            var endPoint = pathPoints[(i + 1) % pathPoints.length];
            if (!isStraightSegment(startPoint, endPoint)) continue;

            var start = startPoint.anchor;
            var end = endPoint.anchor;

            /* 2点が重なっているゴミパスは飛ばす / Skip degenerate segments */
            if (isSamePosition(start, end)) continue;
            if (!isDirectionEnabled(classifySegmentDirection(start, end), settings)) continue;

            drawLineAcrossDrawArea(container, start, end, area, settings, dedupMap);
        }
    }

    /**
     * 設定に従って、アンカー図形・円・補助線をまとめて描く
     * @param {Array<PathItem>} targetPaths - 対象のパス
     * @param {object} settings - ダイアログの設定
     * @param {DrawArea} area - 描画範囲
     * @param {Layer|GroupItem} lineContainer - 線と円の追加先
     * @param {Layer|GroupItem} anchorContainer - アンカー図形の追加先
     * @returns {void}
     */
    function drawConstructionItems(targetPaths, settings, area, lineContainer, anchorContainer) {
        var dedupMap = {};
        var i;

        if (settings.anchorShape !== ANCHOR_SHAPE.NONE) {
            createAnchorShapes(targetPaths, anchorContainer, settings);
        }

        if (settings.arcToCircle) {
            for (i = 0; i < targetPaths.length; i++) {
                try {
                    createCirclesFromArcPath(targetPaths[i], lineContainer, settings, area, dedupMap);
                } catch (e) { logError(e, "createCirclesFromArcPath"); }
            }
        }

        if (settings.straightLines) {
            for (i = 0; i < targetPaths.length; i++) {
                drawExtensionsFromPath(targetPaths[i], lineContainer, settings, area, dedupMap);
            }
        }
    }

    /**
     * @typedef {object} OutputContainers
     * @property {Layer|GroupItem} lineContainer - 線と円の追加先
     * @property {Layer|GroupItem} anchorContainer - アンカー図形の追加先
     */

    /**
     * 設定に従って描画先のレイヤーとグループを用意する
     * @param {Document} doc - 対象のドキュメント
     * @param {Array} selectedItems - 選択アイテム
     * @param {Layer} activeLayer - もとのアクティブレイヤー
     * @param {object} settings - ダイアログの設定
     * @returns {OutputContainers} 描画先
     */
    function prepareOutputContainers(doc, selectedItems, activeLayer, settings) {
        var lineLayer = activeLayer;
        var anchorLayer = activeLayer;
        var needsAnchors = (settings.anchorShape !== ANCHOR_SHAPE.NONE);

        if (settings.separateLayer) {
            lineLayer = prepareLineLayer(doc, selectedItems);
            anchorLayer = prepareAnchorLayer(doc, needsAnchors) || lineLayer;
        }

        var outputContainers = { lineContainer: lineLayer, anchorContainer: anchorLayer };
        if (!settings.group) return outputContainers;

        /* 空のグループを残さないよう、描くものがある側だけグループを作る / Only group what will actually be drawn */
        if (settings.straightLines || settings.arcToCircle) {
            outputContainers.lineContainer = addMarkedGroup(lineLayer, SCRIPT_MARKER + "_" + getLabel(LABELS.itemName.lineGroup));
        }
        if (needsAnchors) {
            outputContainers.anchorContainer = addMarkedGroup(anchorLayer, SCRIPT_MARKER + "_AnchorShapes");
        }

        return outputContainers;
    }

    // =========================================
    // 画面ズーム / Zoom controls
    // =========================================

    /**
     * @typedef {object} ViewState
     * @property {View} view - 対象のビュー
     * @property {number} zoom - 倍率
     * @property {Array<number>} center - 中心の座標 [x, y]
     */

    /**
     * 現在の表示状態を控える
     * @param {Document} doc - 対象のドキュメント
     * @returns {ViewState} 表示状態
     */
    function captureViewState(doc) {
        var viewState = { view: null, zoom: null, center: null };
        /* ビューが取れないことがある / The view may be unavailable */
        try {
            viewState.view = doc.activeView;
            viewState.zoom = viewState.view.zoom;
            viewState.center = viewState.view.centerPoint;
        } catch (e) { logError(e, "captureViewState"); }
        return viewState;
    }

    /**
     * 控えておいた表示状態に戻す
     * @param {ViewState} viewState - 表示状態
     * @returns {void}
     */
    function restoreViewState(viewState) {
        if (!viewState || !viewState.view) return;
        /* ビューが閉じられていることがある / The view may be gone */
        try {
            if (viewState.zoom != null) viewState.view.zoom = viewState.zoom;
            if (viewState.center != null) viewState.view.centerPoint = viewState.center;
        } catch (e) { logError(e, "restoreViewState"); }
    }

    /**
     * ズームの中心を修飾キーに応じて移す（option でアートボード中心、shift で選択中心）
     * @param {Document} doc - 対象のドキュメント
     * @param {View} targetView - 対象のビュー
     * @returns {boolean} メニューコマンドでズームまで処理した場合は true
     */
    function moveZoomCenterByModifier(doc, targetView) {
        var keyboardState = ScriptUI.environment.keyboardState;
        if (!keyboardState) return false;

        var hasSelection = !!(doc.selection && doc.selection.length > 0);

        /* option + shift：選択をウィンドウにフィット / Fit the selection in the window */
        if (keyboardState.altKey && keyboardState.shiftKey && hasSelection) {
            try {
                app.executeMenuCommand("fitinwindow");
                return true;
            } catch (e) { logError(e, "moveZoomCenterByModifier/fitinwindow"); }
            return false;
        }

        /* option：アートボード中心 / Center on the artboard */
        if (keyboardState.altKey) {
            var artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
            targetView.centerPoint = [(artboardRect[0] + artboardRect[2]) / 2, (artboardRect[1] + artboardRect[3]) / 2];
            return false;
        }

        /* shift：選択中心 / Center on the selection */
        if (keyboardState.shiftKey && hasSelection) {
            var selectionBounds = getUnionBounds(doc.selection);
            if (selectionBounds) {
                targetView.centerPoint = [(selectionBounds[0] + selectionBounds[2]) / 2, (selectionBounds[1] + selectionBounds[3]) / 2];
            }
        }

        return false;
    }

    /**
     * ズームのスライダーと軽量モードのチェックボックスを追加する
     * @param {Window} parentWindow - 追加先のウィンドウ
     * @param {Document} doc - 対象のドキュメント
     * @param {ViewState} initialState - 開いた時点の表示状態
     * @returns {{restoreInitial: Function}} 開いた時点の表示に戻す関数
     */
    function addZoomControls(parentWindow, doc, initialState) {
        var zoomRow = parentWindow.add("group");
        zoomRow.orientation = "row";
        zoomRow.alignChildren = ["center", "center"];
        zoomRow.alignment = "center";
        zoomRow.margins = ZOOM_ROW_MARGINS;

        zoomRow.add("statictext", undefined, labelText(LABELS.fieldLabel.zoom));

        var initialZoom = Number(initialState && initialState.zoom);
        if (!initialZoom || isNaN(initialZoom)) initialZoom = 1;

        var zoomSlider = zoomRow.add("slider", undefined, initialZoom, ZOOM_MIN, ZOOM_MAX);
        zoomSlider.preferredSize.width = ZOOM_SLIDER_WIDTH;
        zoomSlider.helpTip = getLabel(LABELS.tooltip.zoom);

        var lightModeCheckbox = addCheckbox(zoomRow, LABELS.checkbox.lightMode, false, LABELS.tooltip.lightMode);

        /**
         * スライダーの値を画面に反映する
         * @returns {void}
         */
        function applyZoom() {
            var targetView = (initialState && initialState.view) ? initialState.view : doc.activeView;
            if (!targetView) return;

            /* ビューが閉じられている・メニューコマンドが失敗するときは何もしない / Ignore a missing view or a failed menu command */
            try {
                if (moveZoomCenterByModifier(doc, targetView)) return;
                targetView.zoom = Number(zoomSlider.value);
                app.redraw();
            } catch (e) { logError(e, "addZoomControls/applyZoom"); }
        }

        /* 軽量モードではドラッグ中に再描画せず、離したときだけ反映する / Light mode applies on release only */
        zoomSlider.onChanging = function() {
            if (lightModeCheckbox.value) return;
            applyZoom();
        };
        zoomSlider.onChange = applyZoom;
        lightModeCheckbox.onClick = applyZoom;

        return {
            restoreInitial: function() { restoreViewState(initialState); }
        };
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * チェックボックスを追加する
     * @param {object} parentContainer - 追加先のウィンドウ・パネル・グループ
     * @param {object} labelSet - ja/en を持つラベル
     * @param {boolean} initialValue - 初期値
     * @param {object} [tooltipSet] - ja/en を持つ tooltip
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addCheckbox(parentContainer, labelSet, initialValue, tooltipSet) {
        var newCheckbox = parentContainer.add("checkbox", undefined, getLabel(labelSet));
        newCheckbox.value = initialValue;
        if (tooltipSet) newCheckbox.helpTip = getLabel(tooltipSet);
        return newCheckbox;
    }

    /**
     * ラジオボタンを追加する
     * @param {object} parentContainer - 追加先のパネル・グループ
     * @param {object} labelSet - ja/en を持つラベル
     * @param {object} [tooltipSet] - ja/en を持つ tooltip
     * @returns {RadioButton} 追加したラジオボタン
     */
    function addRadio(parentContainer, labelSet, tooltipSet) {
        var newRadio = parentContainer.add("radiobutton", undefined, getLabel(labelSet));
        if (tooltipSet) newRadio.helpTip = getLabel(tooltipSet);
        return newRadio;
    }

    /**
     * 長さの入力行（項目名・入力欄・単位）を追加する
     * @param {object} parentContainer - 追加先のパネル
     * @param {object} labelSet - 項目名のラベル
     * @param {string} prefKey - 単位を決める環境設定キー（strokeUnits / rulerType）
     * @param {number} initialPt - 初期値（pt）
     * @returns {{row: Group, input: EditText}} 追加した行と入力欄
     */
    function addLengthRow(parentContainer, labelSet, prefKey, initialPt) {
        var lengthRow = addRow(parentContainer);
        lengthRow.add("statictext", undefined, labelText(labelSet));

        var lengthUnit = getUnitInfo(prefKey);
        var lengthInput = lengthRow.add("edittext", undefined, (initialPt / lengthUnit.pointsPerUnit).toFixed(3));
        lengthInput.characters = NUMBER_INPUT_CHARACTERS;
        lengthRow.add("statictext", undefined, lengthUnit.label);
        return { row: lengthRow, input: lengthInput };
    }

    /**
     * 設定ダイアログを組み立てる（イベントは showDialog() で結線する）
     * @param {Document} doc - 対象のドキュメント
     * @param {object} session - セッション記憶（前回のダイアログ位置）
     * @returns {object} ダイアログと各コントロール
     */
    function buildDialog(doc, session) {
        var mainDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(mainDialog);

        /* 直前の表示位置を引き継ぐ（使えない値なら既定の位置のまま） / Restore the last dialog position if it is still valid */
        if (session.dialogLocation) {
            try { mainDialog.location = session.dialogLocation; } catch (e) { }
        }

        // --- 2カラム / Two columns ---
        var columnsGroup = mainDialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.alignment = ["fill", "top"];
        columnsGroup.spacing = COLUMN_SPACING;

        // --- 左カラム：補助線を描画 / Left column: construction lines ---
        var linesPanel = addPanel(columnsGroup, LABELS.panel.lines);
        var straightCheckbox = addCheckbox(linesPanel, LABELS.checkbox.straight, true, LABELS.tooltip.straight);

        var directionPanel = addPanel(linesPanel);
        var horizontalCheckbox = addCheckbox(directionPanel, LABELS.checkbox.horizontal, true, LABELS.tooltip.soloDirection);
        var verticalCheckbox = addCheckbox(directionPanel, LABELS.checkbox.vertical, true, LABELS.tooltip.soloDirection);
        var diagonalCheckbox = addCheckbox(directionPanel, LABELS.checkbox.diagonal, true, LABELS.tooltip.soloDirection);

        var arcToCircleCheckbox = addCheckbox(linesPanel, LABELS.checkbox.arcToCircle, true, LABELS.tooltip.arcToCircle);

        var arcFallbackPanel = addPanel(linesPanel);
        arcFallbackPanel.helpTip = getLabel(LABELS.tooltip.arcFallback);
        var arcIgnoreRadio = addRadio(arcFallbackPanel, LABELS.radio.arcIgnore, LABELS.tooltip.arcIgnore);
        var arcChordRadio = addRadio(arcFallbackPanel, LABELS.radio.arcChord, LABELS.tooltip.arcChord);
        var arcExtendRadio = addRadio(arcFallbackPanel, LABELS.radio.arcExtendChord, LABELS.tooltip.arcExtendChord);
        arcIgnoreRadio.value = true;

        var strokeWidthCache = { value: mmToPt(DEFAULT_STROKE_WIDTH_MM) };
        var strokeWidthField = addLengthRow(linesPanel, LABELS.fieldLabel.strokeWidth, "strokeUnits", strokeWidthCache.value);
        strokeWidthField.row.margins = [0, 6, 0, 0];

        // --- 右カラム / Right column ---
        var rightColumn = columnsGroup.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = ["fill", "top"];
        rightColumn.alignment = ["fill", "top"];
        rightColumn.spacing = COLUMN_SPACING;

        // --- アンカーポイントに図形 / Shapes on anchor points ---
        var anchorPanel = addPanel(rightColumn, LABELS.panel.anchorShapes);

        var anchorShapeRow = addRow(anchorPanel, RADIO_SPACING);
        var anchorNoneRadio = addRadio(anchorShapeRow, LABELS.radio.anchorNone);
        var anchorCircleRadio = addRadio(anchorShapeRow, LABELS.radio.anchorCircle);
        var anchorSquareRadio = addRadio(anchorShapeRow, LABELS.radio.anchorSquare);
        anchorNoneRadio.value = true;

        var anchorSizeCache = { value: mmToPt(DEFAULT_ANCHOR_SIZE_MM) };
        var anchorSizeField = addLengthRow(anchorPanel, LABELS.fieldLabel.anchorSize, "rulerType", anchorSizeCache.value);

        var anchorColorRow = addRow(anchorPanel, RADIO_SPACING);
        var anchorBlackRadio = addRadio(anchorColorRow, LABELS.radio.colorBlack);
        var anchorBlueRadio = addRadio(anchorColorRow, LABELS.radio.colorBlue);
        anchorBlackRadio.value = true;

        // --- オプション / Options ---
        var optionsPanel = addPanel(rightColumn, LABELS.panel.options);
        var groupCheckbox = addCheckbox(optionsPanel, LABELS.checkbox.group, true, LABELS.tooltip.group);
        var separateLayerCheckbox = addCheckbox(optionsPanel, LABELS.checkbox.separateLayer, true, LABELS.tooltip.separateLayer);
        var guideCheckbox = addCheckbox(optionsPanel, LABELS.checkbox.guide, false, LABELS.tooltip.guide);
        var dedupCheckbox = addCheckbox(optionsPanel, LABELS.checkbox.dedup, true, LABELS.tooltip.dedup);

        // --- 画面ズーム / Zoom controls ---
        var zoomControls = addZoomControls(mainDialog, doc, captureViewState(doc));

        mainDialog.add("panel", undefined, undefined);

        // --- ボタンエリア / Button row ---
        var btnRowGroup = mainDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        var previewCheckbox = addCheckbox(btnLeftGroup, LABELS.checkbox.preview, true);

        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.spacing = BUTTON_SPACING;
        var btnCancel = btnRightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = btnRightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        return {
            dialog: mainDialog,
            straightCheckbox: straightCheckbox,
            directionPanel: directionPanel,
            horizontalCheckbox: horizontalCheckbox,
            verticalCheckbox: verticalCheckbox,
            diagonalCheckbox: diagonalCheckbox,
            arcToCircleCheckbox: arcToCircleCheckbox,
            arcFallbackPanel: arcFallbackPanel,
            arcIgnoreRadio: arcIgnoreRadio,
            arcChordRadio: arcChordRadio,
            arcExtendRadio: arcExtendRadio,
            strokeWidthInput: strokeWidthField.input,
            strokeWidthCache: strokeWidthCache,
            anchorNoneRadio: anchorNoneRadio,
            anchorCircleRadio: anchorCircleRadio,
            anchorSquareRadio: anchorSquareRadio,
            anchorSizeRow: anchorSizeField.row,
            anchorSizeInput: anchorSizeField.input,
            anchorSizeCache: anchorSizeCache,
            anchorColorRow: anchorColorRow,
            anchorBlackRadio: anchorBlackRadio,
            anchorBlueRadio: anchorBlueRadio,
            groupCheckbox: groupCheckbox,
            separateLayerCheckbox: separateLayerCheckbox,
            guideCheckbox: guideCheckbox,
            dedupCheckbox: dedupCheckbox,
            zoomControls: zoomControls,
            previewCheckbox: previewCheckbox,
            btnCancel: btnCancel,
            btnOK: btnOK
        };
    }

    /**
     * 設定ダイアログを表示して設定を取得する
     * @param {Document} doc - 対象のドキュメント
     * @param {Array} selectedItems - 選択アイテム
     * @param {Array<PathItem>} targetPaths - 対象のパス
     * @returns {object|null} ダイアログの設定（キャンセル時は null）
     */
    function showDialog(doc, selectedItems, targetPaths) {
        var session = getSession();
        var dialogControls = buildDialog(doc, session);

        var mainDialog = dialogControls.dialog;
        var straightCheckbox = dialogControls.straightCheckbox;
        var horizontalCheckbox = dialogControls.horizontalCheckbox;
        var verticalCheckbox = dialogControls.verticalCheckbox;
        var diagonalCheckbox = dialogControls.diagonalCheckbox;
        var arcToCircleCheckbox = dialogControls.arcToCircleCheckbox;
        var arcChordRadio = dialogControls.arcChordRadio;
        var arcExtendRadio = dialogControls.arcExtendRadio;
        var anchorNoneRadio = dialogControls.anchorNoneRadio;
        var anchorCircleRadio = dialogControls.anchorCircleRadio;
        var anchorSquareRadio = dialogControls.anchorSquareRadio;
        var anchorBlueRadio = dialogControls.anchorBlueRadio;
        var groupCheckbox = dialogControls.groupCheckbox;
        var separateLayerCheckbox = dialogControls.separateLayerCheckbox;
        var guideCheckbox = dialogControls.guideCheckbox;
        var dedupCheckbox = dialogControls.dedupCheckbox;
        var previewCheckbox = dialogControls.previewCheckbox;
        var zoomControls = dialogControls.zoomControls;

        // =========================================
        // ダイアログの状態 / Dialog behavior
        // =========================================

        var previewLayerName = createUniqueLayerName(doc, PREVIEW_LAYER_BASE_NAME);
        var isAccepted = false;

        /**
         * ダイアログの入力内容を設定オブジェクトにまとめる
         * @returns {object} ダイアログの設定
         */
        function getUISettings() {
            return {
                straightLines: straightCheckbox.value,
                horizontal: horizontalCheckbox.value,
                vertical: verticalCheckbox.value,
                diagonal: diagonalCheckbox.value,
                arcToCircle: arcToCircleCheckbox.value,
                arcFallback: arcChordRadio.value ? ARC_FALLBACK.CHORD :
                    (arcExtendRadio.value ? ARC_FALLBACK.EXTEND : ARC_FALLBACK.IGNORE),
                group: groupCheckbox.value,
                separateLayer: separateLayerCheckbox.value,
                guide: guideCheckbox.value,
                dedup: dedupCheckbox.value,
                anchorShape: anchorCircleRadio.value ? ANCHOR_SHAPE.CIRCLE :
                    (anchorSquareRadio.value ? ANCHOR_SHAPE.SQUARE : ANCHOR_SHAPE.NONE),
                anchorColor: anchorBlueRadio.value ? ANCHOR_COLOR.BLUE : ANCHOR_COLOR.BLACK,
                anchorSizePt: readLengthAsPt(dialogControls.anchorSizeInput, "rulerType", dialogControls.anchorSizeCache),
                strokeWidthPt: readLengthAsPt(dialogControls.strokeWidthInput, "strokeUnits", dialogControls.strokeWidthCache)
            };
        }

        /**
         * 選べない組み合わせのコントロールをディムする
         * @returns {void}
         */
        function updateEnabledState() {
            dialogControls.directionPanel.enabled = straightCheckbox.value;
            dialogControls.arcFallbackPanel.enabled = arcToCircleCheckbox.value;

            /* 線を1本も描かないなら、ガイド化とダブり削除は効かない / Both options are moot without any line */
            var drawsLines = (straightCheckbox.value || arcToCircleCheckbox.value);
            guideCheckbox.enabled = drawsLines;
            dedupCheckbox.enabled = drawsLines;

            var drawsAnchors = !anchorNoneRadio.value;
            dialogControls.anchorSizeRow.enabled = drawsAnchors;
            dialogControls.anchorColorRow.enabled = drawsAnchors;
        }

        /**
         * プレビューを消す
         * @returns {void}
         */
        function clearPreview() {
            removeLayerIfExists(doc, previewLayerName);
            app.redraw();
        }

        /**
         * 現在の設定でプレビューを描き直す（ユーザーのオブジェクトには触らない）
         * @returns {void}
         */
        function refreshPreview() {
            if (!previewCheckbox.value) return;

            removeLayerIfExists(doc, previewLayerName);

            var previewLayer = doc.layers.add();
            previewLayer.name = previewLayerName;
            previewLayer.zOrder(ZOrderMethod.BRINGTOFRONT);

            var previewGroup = addMarkedGroup(previewLayer, SCRIPT_MARKER + "_PREVIEW");
            var settings = getUISettings();

            /* 本処理と同じグループ構造でプレビューする / Mirror the grouping used by the real run */
            var lineContainer = previewGroup;
            var anchorContainer = previewGroup;
            if (settings.group) {
                lineContainer = addMarkedGroup(previewGroup, SCRIPT_MARKER + "_" + getLabel(LABELS.itemName.lineGroup));
                anchorContainer = addMarkedGroup(previewGroup, SCRIPT_MARKER + "_AnchorShapes");
            }

            drawConstructionItems(targetPaths, settings, getDrawArea(doc, selectedItems), lineContainer, anchorContainer);
            app.redraw();
        }

        /**
         * 設定変更のハンドラーを作る（ディム状態を更新してプレビューを描き直す）
         * @param {function} [beforeRefresh] - プレビュー更新の前に行う処理
         * @returns {function} onClick に割り当てるハンドラー
         */
        function onSettingChanged(beforeRefresh) {
            return function() {
                if (typeof beforeRefresh === "function") beforeRefresh();
                updateEnabledState();
                refreshPreview();
            };
        }

        /**
         * option + クリックで、その向きだけを残す
         * @param {Checkbox} clickedCheckbox - クリックされたチェックボックス
         * @param {Checkbox} otherA - もう一方のチェックボックス
         * @param {Checkbox} otherB - さらにもう一方のチェックボックス
         * @returns {void}
         */
        function soloDirectionOnAltClick(clickedCheckbox, otherA, otherB) {
            var keyboardState = ScriptUI.environment.keyboardState;
            if (!keyboardState || !keyboardState.altKey) return;
            if (!clickedCheckbox.value) return;

            otherA.value = false;
            otherB.value = false;
        }

        var settingControls = [
            straightCheckbox, arcToCircleCheckbox, groupCheckbox, separateLayerCheckbox, guideCheckbox, dedupCheckbox,
            dialogControls.arcIgnoreRadio, arcChordRadio, arcExtendRadio,
            anchorNoneRadio, anchorCircleRadio, anchorSquareRadio, dialogControls.anchorBlackRadio, anchorBlueRadio
        ];
        for (var i = 0; i < settingControls.length; i++) {
            settingControls[i].onClick = onSettingChanged();
        }

        horizontalCheckbox.onClick = onSettingChanged(function() {
            soloDirectionOnAltClick(horizontalCheckbox, verticalCheckbox, diagonalCheckbox);
        });
        verticalCheckbox.onClick = onSettingChanged(function() {
            soloDirectionOnAltClick(verticalCheckbox, horizontalCheckbox, diagonalCheckbox);
        });
        diagonalCheckbox.onClick = onSettingChanged(function() {
            soloDirectionOnAltClick(diagonalCheckbox, horizontalCheckbox, verticalCheckbox);
        });

        previewCheckbox.onClick = function() {
            if (previewCheckbox.value) refreshPreview();
            else clearPreview();
        };

        dialogControls.strokeWidthInput.onChanging = refreshPreview;
        dialogControls.anchorSizeInput.onChanging = refreshPreview;
        changeValueByArrowKey(dialogControls.strokeWidthInput, false, refreshPreview);
        changeValueByArrowKey(dialogControls.anchorSizeInput, false, refreshPreview);

        dialogControls.btnOK.onClick = function() {
            isAccepted = true;
            mainDialog.close(1);
        };

        dialogControls.btnCancel.onClick = function() {
            zoomControls.restoreInitial();
            mainDialog.close(0);
        };

        /* タイトルバーの×もキャンセルと同じ扱いにする / Closing from the title bar cancels too */
        mainDialog.onClose = function() {
            if (!isAccepted) zoomControls.restoreInitial();
            return true;
        };

        updateEnabledState();
        /* 既定でONなので、ダイアログを開く前に描いておく / Preview is on by default, so draw it up front */
        refreshPreview();
        var dialogResult = mainDialog.show();

        clearPreview();
        /* 閉じたあとの位置が読めないときは記憶しない / Skip saving if the closed dialog has no readable location */
        try { session.dialogLocation = [mainDialog.location[0], mainDialog.location[1]]; } catch (e) { }

        return (dialogResult === 1) ? getUISettings() : null;
    }

    // =========================================
    // エントリーポイント / Entry point
    // =========================================

    /**
     * 補助線を描く本体の処理（suspendHistory から呼ばれる）
     * @returns {void}
     */
    function mainImpl() {
        var doc = app.activeDocument;
        var selectedItems = doc.selection;

        if (!selectedItems || selectedItems.length === 0) {
            alert(getLabel(LABELS.alert.noSelection));
            return;
        }

        var targetPaths = [];
        extractPathItems(selectedItems, targetPaths);

        /* テキストは複製をアウトライン化してからパスを拾う / Outline a copy of the text to read its paths */
        var outlineRoots = outlineTextFromSelection(selectedItems);
        if (outlineRoots.length > 0) extractPathItems(outlineRoots, targetPaths);

        try {
            if (targetPaths.length === 0) {
                alert(getLabel(LABELS.alert.noValidPath));
                return;
            }

            /* プレビューがレイヤーを増減するので、先にアクティブレイヤーを控える / Preview adds and removes layers */
            var activeLayer = doc.activeLayer;

            var settings = showDialog(doc, selectedItems, targetPaths);
            if (!settings) return;

            var containers = prepareOutputContainers(doc, selectedItems, activeLayer, settings);
            drawConstructionItems(
                targetPaths,
                settings,
                getDrawArea(doc, selectedItems),
                containers.lineContainer,
                containers.anchorContainer
            );
        } finally {
            /* 失敗しても一時アウトラインは必ず片付ける / Always clean up the temporary outlines */
            cleanupTempOutlines(outlineRoots);
        }
    }

    /**
     * ドキュメントを確認して本体を実行する（suspendHistory が使える環境では取り消し1ステップにまとめる）
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        try {
            var doc = app.activeDocument;
            if (typeof doc.suspendHistory === "function") {
                doc.suspendHistory(getLabel(LABELS.itemName.history), "mainImpl()");
            } else {
                mainImpl();
            }
        } catch (e) {
            logError(e, "main");
            alert(getLabel(LABELS.alert.error) + e);
        }
    }

    main();

})();
