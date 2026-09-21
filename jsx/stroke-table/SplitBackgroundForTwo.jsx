#target illustrator
#targetengine "SplitBackgroundForTwoEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

2つのオブジェクトを選択して実行すると、背面に2分割の背景を作成します。
左右・上下のどちらで分けるかは位置関係から自動で判別し、大きさや分割位置はプレビューを見ながら調整できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SplitBackgroundForTwo.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n1b7b8759e53b

### Overview

Creates a two-part background behind two selected objects.
Whether it splits left/right or top/bottom follows from how the objects are placed, and the size and split position are adjusted with a live preview.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SplitBackgroundForTwo.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SplitBackgroundForTwo";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.9.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-01-24";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SplitBackgroundForTwo.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SplitBackgroundForTwo.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n1b7b8759e53b"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    var DEFAULT_SIZE_PERCENT = 200;        /* サイズ（%）の初期値 / initial size (%) */
    var DEFAULT_STROKE_PT = 1;             /* 線幅の初期値（pt）/ initial stroke width (pt) */
    var DEFAULT_CORNER_PT = 0;             /* 角丸の初期値（pt）/ initial corner radius (pt) */
    var FIRST_FILL_RGB = [220, 220, 220];  /* 左（上）側の塗り / fill of the left (top) side */
    var SECOND_FILL_RGB = [128, 128, 128]; /* 右（下）側の塗り / fill of the right (bottom) side */
    var LINE_CMYK = [0, 0, 0, 100];        /* 全体の枠と区切り線の色（K100）/ color of the frame and divider (K100) */

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_MARGINS = 18;               /* ダイアログ外周の余白 / dialog margins */
    var DIALOG_OFFSET_X = 300;             /* 表示位置の横のずらし量 / horizontal dialog offset */
    var DIALOG_OFFSET_Y = 0;               /* 表示位置の縦のずらし量 / vertical dialog offset */
    var DIALOG_OPACITY = 0.98;             /* ダイアログの不透明度 / dialog opacity */
    var PANEL_MARGINS = [15, 20, 15, 10];  /* パネル余白 [左,上,右,下] / panel margins */
    var COLUMN_SPACING = 12;               /* 2カラムの間隔 / column gutter */
    var BALANCE_SPACING = 6;               /* 幅の入力欄とスライダーの間隔 / gap between the width field and slider */
    var SIZE_INPUT_CHARS = 6;              /* サイズ・幅の入力欄の文字数 / characters for the size and width fields */
    var LENGTH_INPUT_CHARS = 4;            /* 線幅・角丸の入力欄の文字数 / characters for the stroke and corner fields */
    var BALANCE_SLIDER_SIZE = [220, 20];   /* 幅スライダーの大きさ / width slider size */

    /**
     * 共通レイアウトのパネルを追加する
     * @param {Group|Panel|Window} parent - 追加先
     * @param {string} titleText - パネルのタイトル
     * @param {string} [childAlignment] - 子の横方向の揃え（既定は "left"）
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parent, titleText, childAlignment) {
        var panel = parent.add("panel", undefined, titleText);
        panel.orientation = "column";
        panel.alignChildren = [childAlignment || "left", "top"];
        panel.margins = PANEL_MARGINS;
        return panel;
    }

    /**
     * 横並びの行グループを追加する
     * @param {Group|Panel|Window} parent - 追加先
     * @returns {Group} 追加した行グループ
     */
    function addRow(parent) {
        var rowGroup = parent.add("group");
        rowGroup.orientation = "row";
        rowGroup.alignChildren = ["left", "center"];
        return rowGroup;
    }

    /**
     * ツールチップ付きのチェックボックスを追加する
     * @param {Group|Panel|Window} parent - 追加先
     * @param {string} labelKey - ラベルの LABELS キー
     * @param {string} tooltipKey - ツールチップの LABELS キー
     * @param {boolean} checked - 初期状態
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addCheckbox(parent, labelKey, tooltipKey, checked) {
        var checkbox = parent.add("checkbox", undefined, getLabel(labelKey));
        checkbox.helpTip = getLabel(tooltipKey);
        checkbox.value = !!checked;
        return checkbox;
    }

    /**
     * ツールチップ付きのラジオボタンを追加する
     * @param {Group} parent - 追加先
     * @param {string} labelKey - ラベルの LABELS キー
     * @param {string} tooltipKey - ツールチップの LABELS キー
     * @returns {RadioButton} 追加したラジオボタン
     */
    function addRadio(parent, labelKey, tooltipKey) {
        var radio = parent.add("radiobutton", undefined, getLabel(labelKey));
        radio.helpTip = getLabel(tooltipKey);
        return radio;
    }

    /**
     * 長さの入力欄と単位ラベルを追加する
     * @param {Group} rowGroup - 追加先の行
     * @param {number} valuePt - 初期値（pt）
     * @param {string} prefKey - 単位の環境設定キー
     * @param {string} tooltipKey - ツールチップの LABELS キー
     * @returns {EditText} 追加した入力欄
     */
    function addLengthInput(rowGroup, valuePt, prefKey, tooltipKey) {
        var lengthInput = rowGroup.add("edittext", undefined, formatUnitValue(ptToUnit(valuePt, prefKey), prefKey));
        lengthInput.helpTip = getLabel(tooltipKey);
        lengthInput.characters = LENGTH_INPUT_CHARS;
        rowGroup.add("statictext", undefined, getUnitInfo(prefKey).label);
        return lengthInput;
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位テーブル（配列の添字が rulerType コードと一致：0=in, 1=mm, 2=pt …）/ Unit table; the array index equals the rulerType code */
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
     * 環境設定の単位の値を pt に換算する
     * @param {number} value - 単位の値
     * @param {string} prefKey - 環境設定キー
     * @returns {number} pt の値
     */
    function unitToPt(value, prefKey) {
        return value * getUnitInfo(prefKey).pointsPerUnit;
    }

    /**
     * pt の値を環境設定の単位に換算する
     * @param {number} valuePt - pt の値
     * @param {string} prefKey - 環境設定キー
     * @returns {number} 単位の値
     */
    function ptToUnit(valuePt, prefKey) {
        return valuePt / getUnitInfo(prefKey).pointsPerUnit;
    }

    /**
     * 単位の大きさに合わせて丸める（pt・Q は小数1桁、mm は2桁、in・cm・ft など1単位が10pt以上は3桁）
     * @param {number} value - 単位の値
     * @param {string} prefKey - 環境設定キー
     * @returns {number} 丸めた値
     */
    function roundUnitValue(value, prefKey) {
        var pointsPerUnit = getUnitInfo(prefKey).pointsPerUnit;
        var scale = (pointsPerUnit >= 10) ? 1000 : (pointsPerUnit >= 2) ? 100 : 10;
        return Math.round(value * scale) / scale;
    }

    /**
     * 単位の値を入力欄の文字列にする
     * @param {number} value - 単位の値
     * @param {string} prefKey - 環境設定キー
     * @returns {string} 表示用の文字列
     */
    function formatUnitValue(value, prefKey) {
        return String(roundUnitValue(value, prefKey));
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    var uiLang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "2つのオブジェクトの背景を作成", en: "Create Background for Two Objects" }
        },
        panel: {
            draw: { ja: "描画", en: "描画" },
            options: { ja: "オプション", en: "Options" },
            balance: { ja: "バランス", en: "Balance" }
        },
        fieldLabel: {
            height: { ja: "高さ", en: "Height" },
            width: { ja: "幅", en: "Width" },
            balanceWidth: { ja: "幅", en: "Width" },
            strokeWidth: { ja: "線幅", en: "Stroke width" },
            cornerRadius: { ja: "角丸", en: "Corner radius" }
        },
        checkbox: {
            fillLeft: { ja: "塗り（左）", en: "Fill (Left)" },
            fillRight: { ja: "塗り（右）", en: "Fill (Right)" },
            fillTop: { ja: "塗り（上）", en: "Fill (Top)" },
            fillBottom: { ja: "塗り（下）", en: "Fill (Bottom)" },
            overallFrame: { ja: "全体の枠", en: "Overall frame" },
            divider: { ja: "区切り線", en: "Divider" }
        },
        radio: {
            none: { ja: "なし", en: "None" },
            left: { ja: "左", en: "Left" },
            right: { ja: "右", en: "Right" },
            top: { ja: "上", en: "Top" },
            bottom: { ja: "下", en: "Bottom" }
        },
        tooltip: {
            sizePercent: { ja: "分割の割合（％）です。", en: "Split ratio, in percent." },
            fillSide: { ja: "この側に塗りを付けます。", en: "Fills this side." },
            overallFrame: {
                ja: "分割した全体を1つの枠線で囲みます。",
                en: "Draws a single frame around the whole split shape."
            },
            divider: { ja: "分割の境目にケイ線を引きます。", en: "Draws a rule along the split." },
            strokeWidth: { ja: "ケイ線の太さです。", en: "Weight of the rules." },
            cornerRadius: {
                ja: "角丸の半径です。0 で角のままになります。",
                en: "Corner radius. 0 leaves the corners square."
            },
            pinNone: {
                ja: "どちらの幅も固定せず、割合だけで分けます。",
                en: "Pins neither side; the ratio alone decides the split."
            },
            pinLeft: {
                ja: "左側の幅を固定し、残りを右側にします。",
                en: "Pins the width of the left side and gives the rest to the right."
            },
            pinRight: {
                ja: "右側の幅を固定し、残りを左側にします。",
                en: "Pins the width of the right side and gives the rest to the left."
            },
            pinTop: {
                ja: "上側の幅を固定し、残りを下側にします。",
                en: "Pins the width of the top side and gives the rest to the bottom."
            },
            pinBottom: {
                ja: "下側の幅を固定し、残りを上側にします。",
                en: "Pins the width of the bottom side and gives the rest to the top."
            },
            balanceWidth: { ja: "固定する側の幅です。", en: "Width of the pinned side." },
            balanceWidthSlider: {
                ja: "固定する側の幅をドラッグで決めます。",
                en: "Drag to set the width of the pinned side."
            }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            openDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            selectTwoItems: { ja: "2つのオブジェクトを選択してください。", en: "Please select two objects." },
            sizeInvalid: {
                ja: "サイズ（%）は0より大きい数値を入力してください。",
                en: "Enter a size (%) greater than 0."
            }
        }
    };

    /**
     * ドット区切りキーで現在の言語のラベルを取得する
     * @param {string} key - LABELS のキー（例: "panel.balance"）
     * @returns {string} ラベル文字列
     */
    function getLabel(key) {
        var keyParts = key.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < keyParts.length; i++) {
            labelNode = labelNode[keyParts[i]];
        }
        return labelNode[uiLang] || labelNode.en;
    }

    /**
     * コロン付きのラベルを取得する（日本語は全角、英語は半角）
     * @param {string} key - LABELS のキー
     * @returns {string} コロンを付けたラベル文字列
     */
    function labelText(key) {
        return getLabel(key) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // セッション設定 / Session settings
    // =========================================
    // Illustrator の起動中だけダイアログの値を保持する（長さは単位を変えても崩れないよう pt で持つ）
    // Dialog values are kept while Illustrator is running; lengths are stored in pt so unit changes do not skew them

    var SESSION_KEY = "SplitBackgroundForTwo_settings";

    /**
     * セッションに残したダイアログの値を返す（欠けている項目は初期値で補う）
     * @returns {Object} セッション設定
     */
    function getSessionSettings() {
        var session = $.global[SESSION_KEY] || {};
        var defaults = {
            percent: DEFAULT_SIZE_PERCENT,
            fillFirst: true,
            fillSecond: true,
            overallFrame: false,
            divider: false,
            strokeWidthPt: DEFAULT_STROKE_PT,
            cornerRadiusPt: DEFAULT_CORNER_PT,
            balanceMode: "none",
            balanceWidthPt: 0
        };
        for (var key in defaults) {
            if (session[key] === undefined) session[key] = defaults[key];
        }
        $.global[SESSION_KEY] = session;
        return session;
    }

    // =========================================
    // ダイアログ共通 / Dialog helpers
    // =========================================

    /**
     * ダイアログの表示位置をずらす（既存の onShow は先に呼ぶ）
     * @param {Window} dialog - 対象のダイアログ
     * @param {number} offsetX - 横のずらし量
     * @param {number} offsetY - 縦のずらし量
     * @returns {void}
     */
    function shiftDialogPosition(dialog, offsetX, offsetY) {
        var previousOnShow = dialog.onShow;
        dialog.onShow = function () {
            if (typeof previousOnShow === "function") previousOnShow();
            dialog.location = [dialog.location[0] + offsetX, dialog.location[1] + offsetY];
        };
    }

    /**
     * ダイアログの不透明度を設定する
     * @param {Window} dialog - 対象のダイアログ
     * @param {number} opacityValue - 不透明度（0〜1）
     * @returns {void}
     */
    function setDialogOpacity(dialog, opacityValue) {
        try {
            dialog.opacity = opacityValue;
        } catch (e) {
            /* 不透明度を持たない環境では既定のまま / keep the default where opacity is unsupported */
        }
    }

    /**
     * ↑↓キーで数値を増減する（shift で±10、option で±0.1）
     * @param {EditText} editText - 対象の入力欄
     * @param {boolean} allowNegative - 負の値を許すか
     * @param {Function} onChange - 値を変えたあとに呼ぶ関数
     * @returns {void}
     */
    function changeValueByArrowKey(editText, allowNegative, onChange) {
        if (!editText) return;

        editText.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;

            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                if (event.keyName === "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                } else if (event.keyName === "Down") {
                    value = Math.floor((value - 1) / delta) * delta;
                }
                event.preventDefault();
            } else if (keyboard.altKey) {
                delta = 0.1;
                if (event.keyName === "Up") value += delta;
                else if (event.keyName === "Down") value -= delta;
                event.preventDefault();
            } else {
                delta = 1;
                if (event.keyName === "Up") value += delta;
                else if (event.keyName === "Down") value -= delta;
                event.preventDefault();
            }

            if (keyboard.altKey) value = Math.round(value * 10) / 10;
            else value = Math.round(value);

            if (!allowNegative && value < 0) value = 0;

            editText.text = String(value);

            if (typeof onChange === "function") {
                try { onChange(); } catch (e) { }
            }
        });
    }

    // =========================================
    // 実行時の状態 / Run state
    // =========================================

    var doc = null;               /* 対象ドキュメント / target document */
    var targetItems = [];         /* 選択した2つのオブジェクト（左または上が先）/ the two selected objects, left or top first */
    var targetBounds = [];        /* targetItems の外接矩形 [左, 上, 右, 下] / bounds of targetItems [left, top, right, bottom] */
    var isVerticalSplit = false;  /* 上下に分けるか / whether the split runs top/bottom */

    /* プレビュー専用レイヤー（ダイアログを閉じると消す）/ Dedicated preview layer, removed when the dialog closes */
    var PREVIEW_LAYER_NAME = "__SplitBackgroundForTwo__PreviewLayer__";

    // =========================================
    // 計測 / Measuring
    // =========================================

    /**
     * テキストを複製してアウトライン化し、その外接矩形を返す（サイドベアリングを含めない）
     * @param {TextFrame} textFrame - 対象のテキスト
     * @returns {number[]|null} 外接矩形。測れなければ null
     */
    function getOutlineBounds(textFrame) {
        var duplicateText = null;
        var outlineGroup = null;
        var bounds = null;
        try {
            duplicateText = textFrame.duplicate(textFrame.layer, ElementPlacement.PLACEATBEGINNING);
            outlineGroup = duplicateText.createOutline();
            bounds = outlineGroup.geometricBounds;
        } catch (e) {
            /* 空白だけのテキストは中身の無いグループになり、geometricBounds が例外になる / whitespace-only text outlines to an empty group whose bounds throw */
            bounds = null;
        } finally {
            removeItem(outlineGroup);
            /* createOutline() が成功していれば複製は消費済み / the duplicate is already consumed when createOutline() succeeds */
            removeItem(duplicateText);
        }
        return bounds;
    }

    /**
     * オブジェクトの外接矩形を返す。テキストはアウトライン、クリップグループはマスクの範囲で測る
     * @param {PageItem} item - 対象のオブジェクト
     * @returns {number[]} 外接矩形 [左, 上, 右, 下]
     */
    function getItemBounds(item) {
        var bounds = null;
        if (item.typename === "TextFrame") {
            bounds = getOutlineBounds(item);
        } else if (item.typename === "GroupItem" && item.clipped && item.pageItems.length > 0) {
            /* クリップグループのマスクは pageItems[0] / the mask of a clip group is pageItems[0] */
            bounds = getItemBounds(item.pageItems[0]);
        }
        if (!bounds) bounds = item.geometricBounds;
        return [bounds[0], bounds[1], bounds[2], bounds[3]];
    }

    /**
     * 2つの外接矩形のすき間を返す（重なっていれば 0）
     * @param {Array} boundsPair - 2つの外接矩形
     * @param {boolean} vertical - 上下のすき間を測るか（false なら左右）
     * @returns {number} すき間（pt）
     */
    function getGapPt(boundsPair, vertical) {
        var bounds1 = boundsPair[0];
        var bounds2 = boundsPair[1];
        var gapPt;
        if (vertical) {
            gapPt = (bounds1[1] >= bounds2[1]) ? (bounds1[3] - bounds2[1]) : (bounds2[3] - bounds1[1]);
        } else {
            gapPt = (bounds1[0] <= bounds2[0]) ? (bounds2[0] - bounds1[2]) : (bounds1[0] - bounds2[2]);
        }
        return (gapPt > 0) ? gapPt : 0;
    }

    // =========================================
    // レイヤー / Layers
    // =========================================

    /**
     * レイヤーの重ね順を、最上位からの添字の並びで返す（添字が大きいほど背面）
     * @param {Layer} layer - 対象のレイヤー
     * @returns {number[]} 最上位レイヤーから順の添字
     */
    function getLayerStackPath(layer) {
        var stackPath = [];
        while (layer.typename === "Layer") {
            var siblingLayers = layer.parent.layers;
            for (var i = 0; i < siblingLayers.length; i++) {
                if (siblingLayers[i] === layer) break;
            }
            stackPath.unshift(i);
            layer = layer.parent;
        }
        return stackPath;
    }

    /**
     * 2つのレイヤーのうち背面側を返す（サブレイヤーも比べる）
     * @param {Layer} layerA - 1つ目のレイヤー
     * @param {Layer} layerB - 2つ目のレイヤー
     * @returns {Layer} 背面側のレイヤー
     */
    function getBackmostLayer(layerA, layerB) {
        var pathA = getLayerStackPath(layerA);
        var pathB = getLayerStackPath(layerB);
        for (var i = 0; i < pathA.length && i < pathB.length; i++) {
            if (pathA[i] !== pathB[i]) return (pathA[i] > pathB[i]) ? layerA : layerB;
        }
        /* 同じレイヤーか親子なら、浅いほう（親）に描く / same layer or parent and child: use the shallower one */
        return (pathA.length <= pathB.length) ? layerA : layerB;
    }

    /**
     * サブレイヤーをたどって最上位のレイヤーを返す
     * @param {Layer} layer - 対象のレイヤー
     * @returns {Layer} 最上位のレイヤー
     */
    function getTopLevelLayer(layer) {
        while (layer.parent.typename === "Layer") layer = layer.parent;
        return layer;
    }

    /**
     * プレビューレイヤーを探す
     * @returns {Layer|null} プレビューレイヤー。無ければ null
     */
    function findPreviewLayer() {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === PREVIEW_LAYER_NAME) return doc.layers[i];
        }
        return null;
    }

    /**
     * プレビューレイヤーを用意し、2つのオブジェクトのレイヤーより背面に置く
     * @returns {Layer} プレビューレイヤー
     */
    function ensurePreviewLayer() {
        var previewLayer = findPreviewLayer();
        if (previewLayer) return previewLayer;

        previewLayer = doc.layers.add();
        previewLayer.name = PREVIEW_LAYER_NAME;
        /* 背面側のレイヤーの、さらに背面へ（サブレイヤーなら最上位の親の背面へ）/ behind the backmost layer, or behind its top-level parent for a sublayer */
        var backLayer = getTopLevelLayer(getBackmostLayer(targetItems[0].layer, targetItems[1].layer));
        previewLayer.move(backLayer, ElementPlacement.PLACEAFTER);
        return previewLayer;
    }

    /**
     * プレビューレイヤーを中身ごと削除する
     * @returns {void}
     */
    function removePreviewLayer() {
        var previewLayer = findPreviewLayer();
        if (previewLayer) previewLayer.remove();
    }

    // =========================================
    // 描画 / Drawing
    // =========================================

    /**
     * RGB の配列から色を作る
     * @param {number[]} rgbValues - [R, G, B]
     * @returns {RGBColor} 色
     */
    function makeRGBColor(rgbValues) {
        var color = new RGBColor();
        color.red = rgbValues[0];
        color.green = rgbValues[1];
        color.blue = rgbValues[2];
        return color;
    }

    /**
     * CMYK の配列から色を作る
     * @param {number[]} cmykValues - [C, M, Y, K]
     * @returns {CMYKColor} 色
     */
    function makeCMYKColor(cmykValues) {
        var color = new CMYKColor();
        color.cyan = cmykValues[0];
        color.magenta = cmykValues[1];
        color.yellow = cmykValues[2];
        color.black = cmykValues[3];
        return color;
    }

    var FIRST_FILL_COLOR = makeRGBColor(FIRST_FILL_RGB);
    var SECOND_FILL_COLOR = makeRGBColor(SECOND_FILL_RGB);
    var LINE_COLOR = makeCMYKColor(LINE_CMYK);

    /**
     * 削除する（createOutline() に消費された複製など、すでに無いものは無視する）
     * @param {PageItem} item - 削除するオブジェクト
     * @returns {void}
     */
    function removeItem(item) {
        if (!item) return;
        try {
            item.remove();
        } catch (e) {
            /* すでに無い / already gone */
        }
    }

    /**
     * すき間を1つ目側と2つ目側の余白に振り分ける
     * @param {number} gapPt - すき間（pt）
     * @param {string} balanceMode - 'none' | 'first' | 'second'
     * @param {number} balanceWidthPt - 固定する側の幅（pt）
     * @returns {number[]} [1つ目側の余白, 2つ目側の余白]
     */
    function splitGap(gapPt, balanceMode, balanceWidthPt) {
        var pinnedPt = Math.min(Math.max(balanceWidthPt, 0), gapPt);
        if (balanceMode === "first") return [pinnedPt, gapPt - pinnedPt];
        if (balanceMode === "second") return [gapPt - pinnedPt, pinnedPt];
        return [gapPt / 2, gapPt / 2];
    }

    /**
     * 背景全体の範囲と分割位置を求める
     * @param {Object} drawOptions - 描画の設定（collectDrawOptions() の戻り値）
     * @returns {{left: number, top: number, right: number, bottom: number, split: number}} 範囲と分割位置（左右なら X、上下なら Y）
     */
    function computeSplitLayout(drawOptions) {
        var margins = splitGap(getGapPt(targetBounds, isVerticalSplit), drawOptions.balanceMode, drawOptions.balanceWidthPt);
        return isVerticalSplit
            ? computeTopBottomLayout(margins, drawOptions.sizeRatio)
            : computeLeftRightLayout(margins, drawOptions.sizeRatio);
    }

    /**
     * 左右分割の範囲：高さを倍率で広げ、横はすき間の余白ぶん外へ伸ばす
     * @param {number[]} margins - [左側の余白, 右側の余白]
     * @param {number} sizeRatio - 高さの倍率
     * @returns {{left: number, top: number, right: number, bottom: number, split: number}} 範囲と分割位置（X）
     */
    function computeLeftRightLayout(margins, sizeRatio) {
        var leftBounds = targetBounds[0];
        var rightBounds = targetBounds[1];
        var contentTop = Math.max(leftBounds[1], rightBounds[1]);
        var contentHeight = contentTop - Math.min(leftBounds[3], rightBounds[3]);
        var rectHeight = contentHeight * sizeRatio;
        var rectTop = contentTop - (contentHeight / 2) + (rectHeight / 2);
        return {
            left: leftBounds[0] - margins[0],
            top: rectTop,
            right: rightBounds[2] + margins[1],
            bottom: rectTop - rectHeight,
            split: leftBounds[2] + margins[0]
        };
    }

    /**
     * 上下分割の範囲：横幅を倍率で広げ、縦はすき間の余白ぶん外へ伸ばす
     * @param {number[]} margins - [上側の余白, 下側の余白]
     * @param {number} sizeRatio - 横幅の倍率
     * @returns {{left: number, top: number, right: number, bottom: number, split: number}} 範囲と分割位置（Y）
     */
    function computeTopBottomLayout(margins, sizeRatio) {
        var topBounds = targetBounds[0];
        var bottomBounds = targetBounds[1];
        var contentLeft = Math.min(topBounds[0], bottomBounds[0]);
        var contentWidth = Math.max(Math.max(topBounds[2], bottomBounds[2]) - contentLeft, 0);
        var rectWidth = contentWidth * sizeRatio;
        var rectLeft = contentLeft + (contentWidth / 2) - (rectWidth / 2);
        var rectTop = topBounds[1] + margins[0];
        return {
            left: rectLeft,
            top: rectTop,
            right: rectLeft + rectWidth,
            bottom: Math.min(bottomBounds[3] - margins[1], rectTop),
            split: topBounds[3] - margins[0]
        };
    }

    /**
     * パスのアンカーの選択をすべて外す
     * @param {PathItem} pathItem - 対象のパス
     * @returns {void}
     */
    function deselectAnchors(pathItem) {
        var points = pathItem.pathPoints;
        for (var i = 0; i < points.length; i++) {
            points[i].selected = PathPointSelection.NOSELECTION;
        }
    }

    /**
     * 長方形の、指定した辺の両端のアンカーだけを選択する
     * @param {PathItem} rectPath - 4点の長方形
     * @param {string} side - "left" | "right" | "top" | "bottom"
     * @returns {boolean} 選択できたか
     */
    function selectRectangleCorners(rectPath, side) {
        var points = rectPath.pathPoints;
        if (points.length !== 4) return false;

        var axis = (side === "left" || side === "right") ? 0 : 1;
        var order = [0, 1, 2, 3];
        order.sort(function (a, b) { return points[a].anchor[axis] - points[b].anchor[axis]; });
        /* 右と上は座標の大きい2点、左と下は小さい2点 / right and top take the two larger coordinates, left and bottom the two smaller */
        var picked = (side === "right" || side === "top") ? [order[2], order[3]] : [order[0], order[1]];

        deselectAnchors(rectPath);
        points[picked[0]].selected = PathPointSelection.ANCHORPOINT;
        points[picked[1]].selected = PathPointSelection.ANCHORPOINT;
        return true;
    }

    /**
     * 塗りの長方形の外側の2つの角を丸める
     * @param {PathItem} fillRect - 塗りの長方形
     * @param {string} side - 丸める辺（"left" | "right" | "top" | "bottom"）
     * @param {number} cornerRadiusPt - 角丸の半径（pt）
     * @returns {void}
     */
    function roundFillCorners(fillRect, side, cornerRadiusPt) {
        if (!(cornerRadiusPt > 0)) return;
        if (!selectRectangleCorners(fillRect, side)) return;
        try {
            roundAnyCorner([fillRect], { rr: cornerRadiusPt });
        } catch (e) {
            /* 幅0の長方形などで計算に失敗しても、角のまま描く / keep square corners if rounding fails, e.g. on a zero-width rectangle */
        }
        deselectAnchors(fillRect);
    }

    /**
     * 「角を丸くする」効果を適用する
     * @param {PathItem} targetItem - 対象のパス
     * @param {number} cornerRadiusPt - 角丸の半径（pt）
     * @returns {void}
     */
    function applyRoundCornersEffect(targetItem, cornerRadiusPt) {
        if (!(cornerRadiusPt > 0)) return;
        try {
            targetItem.applyEffect('<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + cornerRadiusPt + ' "/></LiveEffect>');
        } catch (e) {
            /* 効果を適用できない環境では角のまま描く / keep square corners where the effect cannot be applied */
        }
    }

    /**
     * 線を K100 のケイ線にする
     * @param {PathItem} pathItem - 対象のパス
     * @param {number} strokeWidthPt - 線幅（pt）
     * @returns {void}
     */
    function styleLine(pathItem, strokeWidthPt) {
        pathItem.filled = false;
        pathItem.stroked = true;
        pathItem.strokeWidth = strokeWidthPt;
        pathItem.strokeColor = LINE_COLOR;
    }

    /**
     * 塗りの長方形を追加し、外側の角を丸める
     * @param {PathItems} pathItems - 追加先
     * @param {number[]} rectBounds - [左, 上, 右, 下]
     * @param {RGBColor} fillColor - 塗りの色
     * @param {string} roundSide - 丸める辺
     * @param {number} cornerRadiusPt - 角丸の半径（pt）
     * @returns {PathItem} 追加した長方形
     */
    function addFillRect(pathItems, rectBounds, fillColor, roundSide, cornerRadiusPt) {
        var fillRect = pathItems.rectangle(rectBounds[1], rectBounds[0], rectBounds[2] - rectBounds[0], rectBounds[1] - rectBounds[3]);
        fillRect.fillColor = fillColor;
        fillRect.stroked = false;
        roundFillCorners(fillRect, roundSide, cornerRadiusPt);
        return fillRect;
    }

    /**
     * 背景の塗り・全体の枠・区切り線を描く
     * @param {Object} drawOptions - 描画の設定（collectDrawOptions() の戻り値）
     * @param {Layer} targetLayer - 描画先のレイヤー
     * @returns {PageItem[]} 描いたオブジェクト
     */
    function drawSplitBackground(drawOptions, targetLayer) {
        var layout = computeSplitLayout(drawOptions);
        var pathItems = targetLayer.pathItems;
        var fillRects = [];
        var drawnItems = [];

        if (drawOptions.fillFirst) {
            fillRects.push(addFillRect(pathItems,
                isVerticalSplit ? [layout.left, layout.top, layout.right, layout.split] : [layout.left, layout.top, layout.split, layout.bottom],
                FIRST_FILL_COLOR, isVerticalSplit ? "top" : "left", drawOptions.cornerRadiusPt));
        }
        if (drawOptions.fillSecond) {
            fillRects.push(addFillRect(pathItems,
                isVerticalSplit ? [layout.left, layout.split, layout.right, layout.bottom] : [layout.split, layout.top, layout.right, layout.bottom],
                SECOND_FILL_COLOR, isVerticalSplit ? "bottom" : "right", drawOptions.cornerRadiusPt));
        }

        if (drawOptions.overallFrame) {
            var frameRect = pathItems.rectangle(layout.top, layout.left, layout.right - layout.left, layout.top - layout.bottom);
            styleLine(frameRect, drawOptions.strokeWidthPt);
            applyRoundCornersEffect(frameRect, drawOptions.cornerRadiusPt);
            frameRect.zOrder(ZOrderMethod.BRINGTOFRONT);
            drawnItems.push(frameRect);
        }

        if (drawOptions.divider) {
            var dividerLine = pathItems.add();
            styleLine(dividerLine, drawOptions.strokeWidthPt);
            dividerLine.setEntirePath(isVerticalSplit
                ? [[layout.left, layout.split], [layout.right, layout.split]]
                : [[layout.split, layout.top], [layout.split, layout.bottom]]);
            dividerLine.zOrder(ZOrderMethod.BRINGTOFRONT);
            drawnItems.push(dividerLine);
        }

        /* 塗りは最背面へ（1つ目が2つ目の上に来る順）/ send the fills to the back, keeping the first above the second */
        for (var i = 0; i < fillRects.length; i++) {
            fillRects[i].zOrder(ZOrderMethod.SENDTOBACK);
            drawnItems.push(fillRects[i]);
        }
        return drawnItems;
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /**
     * プレビューレイヤーの中身を消す
     * @returns {void}
     */
    function clearPreview() {
        var previewLayer = findPreviewLayer();
        if (!previewLayer) return;
        for (var i = previewLayer.pageItems.length - 1; i >= 0; i--) {
            previewLayer.pageItems[i].remove();
        }
    }

    /**
     * プレビューを描き直す
     * @param {Object|null} drawOptions - 描画の設定。null ならプレビューを消すだけ
     * @returns {void}
     */
    function drawPreview(drawOptions) {
        clearPreview();
        if (drawOptions) {
            try {
                drawSplitBackground(drawOptions, ensurePreviewLayer());
            } catch (e) {
                /* 描けなくてもダイアログの入力は続けられるようにする / keep the dialog usable even if drawing fails */
                clearPreview();
            }
            /* 角丸の処理が塗りを選択するので外す / the corner rounding selects the fill; deselect it */
            doc.selection = null;
        }
        app.redraw();
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 設定ダイアログを組み立てる（イベントは bindDialogEvents() で付ける）
     * @param {Object} session - セッションに残した前回の値
     * @returns {Object} ダイアログとコントロールの参照
     */
    function buildSettingsDialog(session) {
        var controls = {};
        var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];
        dialog.margins = DIALOG_MARGINS;
        setDialogOpacity(dialog, DIALOG_OPACITY);
        controls.dialog = dialog;

        /* サイズ（%）：左右なら高さ、上下なら幅 / Size (%): height for a left/right split, width for top/bottom */
        var sizeRowGroup = addRow(dialog);
        sizeRowGroup.alignment = ["center", "top"];
        sizeRowGroup.add("statictext", undefined, labelText(isVerticalSplit ? "fieldLabel.width" : "fieldLabel.height"));
        controls.sizeInput = sizeRowGroup.add("edittext", undefined, String(session.percent));
        controls.sizeInput.helpTip = getLabel("tooltip.sizePercent");
        controls.sizeInput.characters = SIZE_INPUT_CHARS;
        controls.sizeInput.active = true;
        sizeRowGroup.add("statictext", undefined, "%");

        /* 2カラム：左＝描画、右＝オプション / Two columns: drawing on the left, options on the right */
        var columnsGroup = dialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.spacing = COLUMN_SPACING;

        var drawPanel = addPanel(columnsGroup, getLabel("panel.draw"));
        controls.fillFirstCheckbox = addCheckbox(drawPanel, isVerticalSplit ? "checkbox.fillTop" : "checkbox.fillLeft", "tooltip.fillSide", session.fillFirst);
        controls.fillSecondCheckbox = addCheckbox(drawPanel, isVerticalSplit ? "checkbox.fillBottom" : "checkbox.fillRight", "tooltip.fillSide", session.fillSecond);
        controls.overallFrameCheckbox = addCheckbox(drawPanel, "checkbox.overallFrame", "tooltip.overallFrame", session.overallFrame);
        controls.dividerCheckbox = addCheckbox(drawPanel, "checkbox.divider", "tooltip.divider", session.divider);

        var optionsPanel = addPanel(columnsGroup, getLabel("panel.options"));
        controls.strokeRowGroup = addRow(optionsPanel);
        controls.strokeRowGroup.add("statictext", undefined, getLabel("fieldLabel.strokeWidth"));
        controls.strokeInput = addLengthInput(controls.strokeRowGroup, session.strokeWidthPt, "strokeUnits", "tooltip.strokeWidth");
        controls.cornerRowGroup = addRow(optionsPanel);
        controls.cornerRowGroup.add("statictext", undefined, getLabel("fieldLabel.cornerRadius"));
        controls.cornerInput = addLengthInput(controls.cornerRowGroup, session.cornerRadiusPt, "rulerType", "tooltip.cornerRadius");

        buildBalancePanel(dialog, controls, session);

        var btnRowGroup = dialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = ["right", "center"];
        controls.btnCancel = btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        controls.btnOK = btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        dialog.defaultElement = controls.btnOK;
        dialog.cancelElement = controls.btnCancel;

        return controls;
    }

    /**
     * ［バランス］パネル（固定する側と、その幅）を組み立てる
     * @param {Window} dialog - 追加先のダイアログ
     * @param {Object} controls - コントロールの参照（ここで作ったものを足す）
     * @param {Object} session - セッションに残した前回の値
     * @returns {void}
     */
    function buildBalancePanel(dialog, controls, session) {
        var balancePanel = addPanel(dialog, getLabel("panel.balance"), "fill");
        balancePanel.alignment = ["fill", "top"];

        var balanceRadioGroup = addRow(balancePanel);
        controls.pinNoneRadio = addRadio(balanceRadioGroup, "radio.none", "tooltip.pinNone");
        controls.pinFirstRadio = addRadio(balanceRadioGroup, isVerticalSplit ? "radio.top" : "radio.left", isVerticalSplit ? "tooltip.pinTop" : "tooltip.pinLeft");
        controls.pinSecondRadio = addRadio(balanceRadioGroup, isVerticalSplit ? "radio.bottom" : "radio.right", isVerticalSplit ? "tooltip.pinBottom" : "tooltip.pinRight");
        controls.pinFirstRadio.value = (session.balanceMode === "first");
        controls.pinSecondRadio.value = (session.balanceMode === "second");
        controls.pinNoneRadio.value = !controls.pinFirstRadio.value && !controls.pinSecondRadio.value;

        controls.balanceWidthGroup = balancePanel.add("group");
        controls.balanceWidthGroup.orientation = "column";
        controls.balanceWidthGroup.alignChildren = ["fill", "top"];
        controls.balanceWidthGroup.spacing = BALANCE_SPACING;

        var balanceInputRow = addRow(controls.balanceWidthGroup);
        balanceInputRow.add("statictext", undefined, getLabel("fieldLabel.balanceWidth"));
        controls.balanceWidthInput = balanceInputRow.add("edittext", undefined, "");
        controls.balanceWidthInput.helpTip = getLabel("tooltip.balanceWidth");
        controls.balanceWidthInput.characters = SIZE_INPUT_CHARS;
        balanceInputRow.add("statictext", undefined, getUnitInfo("rulerType").label);

        /* 幅の上限は2つのオブジェクトのすき間 / the width cannot exceed the gap between the objects */
        controls.balanceWidthMax = roundUnitValue(ptToUnit(getGapPt(targetBounds, isVerticalSplit), "rulerType"), "rulerType");
        var balanceSliderRow = addRow(controls.balanceWidthGroup);
        balanceSliderRow.alignChildren = ["fill", "center"];
        controls.balanceWidthSlider = balanceSliderRow.add("slider", undefined, 0, 0, controls.balanceWidthMax);
        controls.balanceWidthSlider.helpTip = getLabel("tooltip.balanceWidthSlider");
        controls.balanceWidthSlider.preferredSize = BALANCE_SLIDER_SIZE;

        setBalanceWidth(controls, clampBalanceWidth(controls, ptToUnit(session.balanceWidthPt, "rulerType"), false));
    }

    /**
     * 線幅の欄を使うか（全体の枠か区切り線があるとき）
     * @param {Object} controls - コントロールの参照
     * @returns {boolean} 使うか
     */
    function needsStroke(controls) {
        return controls.overallFrameCheckbox.value || controls.dividerCheckbox.value;
    }

    /**
     * 角丸の欄を使うか（塗りか全体の枠があるとき）
     * @param {Object} controls - コントロールの参照
     * @returns {boolean} 使うか
     */
    function needsCorner(controls) {
        return controls.fillFirstCheckbox.value || controls.fillSecondCheckbox.value || controls.overallFrameCheckbox.value;
    }

    /**
     * 使わない欄を無効にする
     * @param {Object} controls - コントロールの参照
     * @returns {void}
     */
    function updateEnabledStates(controls) {
        controls.strokeRowGroup.enabled = needsStroke(controls);
        controls.cornerRowGroup.enabled = needsCorner(controls);
        controls.balanceWidthGroup.enabled = !controls.pinNoneRadio.value;
    }

    /**
     * サイズ（%）を読む
     * @param {Object} controls - コントロールの参照
     * @returns {number|null} サイズ（%）。0以下や数値でなければ null
     */
    function readSizePercent(controls) {
        var sizePercent = Number(controls.sizeInput.text);
        return (isNaN(sizePercent) || sizePercent <= 0) ? null : sizePercent;
    }

    /**
     * 長さの入力欄を pt に換算して読む
     * @param {EditText} lengthInput - 入力欄
     * @param {string} prefKey - 単位の環境設定キー
     * @param {boolean} allowZero - 0 を許すか
     * @returns {number|null} pt の値（小数1桁）。不正な値なら null
     */
    function readLengthPt(lengthInput, prefKey, allowZero) {
        var valueUnit = Number(lengthInput.text);
        if (isNaN(valueUnit) || valueUnit < 0) return null;
        var valuePt = Math.round(unitToPt(roundUnitValue(valueUnit, prefKey), prefKey) * 10) / 10;
        return (valuePt > 0 || allowZero) ? valuePt : null;
    }

    /**
     * バランスの幅を 0〜上限に収める
     * @param {Object} controls - コントロールの参照
     * @param {number|string} value - 幅（定規の単位）
     * @param {boolean} forceInteger - 整数に丸めるか
     * @returns {number} 収めた幅
     */
    function clampBalanceWidth(controls, value, forceInteger) {
        var widthUnit = Number(value);
        if (isNaN(widthUnit) || widthUnit < 0) widthUnit = 0;
        if (widthUnit > controls.balanceWidthMax) widthUnit = controls.balanceWidthMax;
        return forceInteger ? Math.round(widthUnit) : roundUnitValue(widthUnit, "rulerType");
    }

    /**
     * バランスの幅を入力欄とスライダーの両方に入れる
     * @param {Object} controls - コントロールの参照
     * @param {number} widthUnit - 幅（定規の単位）
     * @returns {void}
     */
    function setBalanceWidth(controls, widthUnit) {
        controls.balanceWidthSlider.value = widthUnit;
        controls.balanceWidthInput.text = String(widthUnit);
    }

    /**
     * バランスの幅を pt で読む
     * @param {Object} controls - コントロールの参照
     * @returns {number} 幅（pt、小数1桁）
     */
    function readBalanceWidthPt(controls) {
        var widthUnit = clampBalanceWidth(controls, controls.balanceWidthInput.text, false);
        return Math.round(unitToPt(widthUnit, "rulerType") * 10) / 10;
    }

    /**
     * 固定する側を返す
     * @param {Object} controls - コントロールの参照
     * @returns {string} 'none' | 'first' | 'second'
     */
    function getBalanceMode(controls) {
        if (controls.pinFirstRadio.value) return "first";
        if (controls.pinSecondRadio.value) return "second";
        return "none";
    }

    /**
     * ダイアログの値を描画の設定にまとめる
     * @param {Object} controls - コントロールの参照
     * @returns {Object|null} 描画の設定。使う欄に不正な値があれば null
     */
    function collectDrawOptions(controls) {
        var sizePercent = readSizePercent(controls);
        var strokeWidthPt = readLengthPt(controls.strokeInput, "strokeUnits", false);
        var cornerRadiusPt = readLengthPt(controls.cornerInput, "rulerType", true);
        if (sizePercent === null) return null;

        /* 使わない欄は、不正な値でも既定値で進める / a field that is not in use falls back to its default */
        if (strokeWidthPt === null) {
            if (needsStroke(controls)) return null;
            strokeWidthPt = DEFAULT_STROKE_PT;
        }
        if (cornerRadiusPt === null) {
            if (needsCorner(controls)) return null;
            cornerRadiusPt = DEFAULT_CORNER_PT;
        }

        return {
            sizeRatio: sizePercent / 100,
            fillFirst: controls.fillFirstCheckbox.value,
            fillSecond: controls.fillSecondCheckbox.value,
            overallFrame: controls.overallFrameCheckbox.value,
            divider: controls.dividerCheckbox.value,
            strokeWidthPt: strokeWidthPt,
            cornerRadiusPt: cornerRadiusPt,
            balanceMode: getBalanceMode(controls),
            balanceWidthPt: readBalanceWidthPt(controls)
        };
    }

    /**
     * ダイアログの値でプレビューを描き直す（使う欄に不正な値があれば消すだけ）
     * @param {Object} controls - コントロールの参照
     * @returns {void}
     */
    function refreshPreview(controls) {
        drawPreview(collectDrawOptions(controls));
    }

    /**
     * ダイアログのイベントを付ける
     * @param {Object} controls - コントロールの参照
     * @returns {void}
     */
    function bindDialogEvents(controls) {
        var refresh = function () { refreshPreview(controls); };
        var onOptionClick = function () {
            updateEnabledStates(controls);
            refresh();
        };
        var syncBalanceWidthFromInput = function () {
            setBalanceWidth(controls, clampBalanceWidth(controls, controls.balanceWidthInput.text, false));
            refresh();
        };

        changeValueByArrowKey(controls.sizeInput, false, refresh);
        changeValueByArrowKey(controls.strokeInput, false, refresh);
        changeValueByArrowKey(controls.cornerInput, false, refresh);
        changeValueByArrowKey(controls.balanceWidthInput, false, syncBalanceWidthFromInput);

        controls.sizeInput.addEventListener("changing", refresh);
        controls.strokeInput.addEventListener("changing", refresh);
        controls.cornerInput.addEventListener("changing", refresh);

        /* 入力中は欄を書き換えない（小数点を打てるように）。確定したら範囲内に収める
           Leave the field alone while typing so a decimal point can be entered; clamp it on commit */
        controls.balanceWidthInput.addEventListener("changing", function () {
            controls.balanceWidthSlider.value = clampBalanceWidth(controls, controls.balanceWidthInput.text, false);
            refresh();
        });
        controls.balanceWidthInput.onChange = syncBalanceWidthFromInput;
        controls.balanceWidthSlider.onChanging = function () {
            /* option を押しながらドラッグすると整数に丸める / hold Option while dragging to snap to whole units */
            var snapToInteger = ScriptUI.environment.keyboardState.altKey;
            setBalanceWidth(controls, clampBalanceWidth(controls, controls.balanceWidthSlider.value, snapToInteger));
            refresh();
        };

        controls.fillFirstCheckbox.onClick = onOptionClick;
        controls.fillSecondCheckbox.onClick = onOptionClick;
        controls.overallFrameCheckbox.onClick = onOptionClick;
        controls.dividerCheckbox.onClick = onOptionClick;
        controls.pinNoneRadio.onClick = onOptionClick;
        controls.pinFirstRadio.onClick = onOptionClick;
        controls.pinSecondRadio.onClick = onOptionClick;

        controls.btnOK.onClick = function () {
            if (readSizePercent(controls) === null) {
                alert(getLabel("alert.sizeInvalid"));
                return;
            }
            if (!collectDrawOptions(controls)) {
                /* 線幅か角丸が不正なときは、その欄に戻す / return to the invalid stroke or corner field */
                var strokeInvalid = needsStroke(controls) && readLengthPt(controls.strokeInput, "strokeUnits", false) === null;
                (strokeInvalid ? controls.strokeInput : controls.cornerInput).active = true;
                return;
            }
            controls.dialog.close(1);
        };
        controls.btnCancel.onClick = function () { controls.dialog.close(0); };

        controls.dialog.onShow = refresh;
        shiftDialogPosition(controls.dialog, DIALOG_OFFSET_X, DIALOG_OFFSET_Y);
        updateEnabledStates(controls);
    }

    /**
     * ダイアログの値をセッションに残す（不正な値の欄は前回の値のまま）
     * @param {Object} controls - コントロールの参照
     * @param {Object} session - セッション設定
     * @returns {void}
     */
    function saveDialogToSession(controls, session) {
        var sizePercent = readSizePercent(controls);
        var strokeWidthPt = readLengthPt(controls.strokeInput, "strokeUnits", false);
        var cornerRadiusPt = readLengthPt(controls.cornerInput, "rulerType", true);
        if (sizePercent !== null) session.percent = sizePercent;
        if (strokeWidthPt !== null) session.strokeWidthPt = strokeWidthPt;
        if (cornerRadiusPt !== null) session.cornerRadiusPt = cornerRadiusPt;

        session.fillFirst = controls.fillFirstCheckbox.value;
        session.fillSecond = controls.fillSecondCheckbox.value;
        session.overallFrame = controls.overallFrameCheckbox.value;
        session.divider = controls.dividerCheckbox.value;
        session.balanceMode = getBalanceMode(controls);
        session.balanceWidthPt = readBalanceWidthPt(controls);
    }

    /**
     * 設定ダイアログを表示する。閉じたらプレビューを片付ける
     * @returns {Object|null} 描画の設定。キャンセルなら null
     */
    function showSettingsDialog() {
        var session = getSessionSettings();
        var controls = buildSettingsDialog(session);
        bindDialogEvents(controls);

        /* 中断した実行で残ったプレビューレイヤーを消してから始める / start without a preview layer left over from an interrupted run */
        removePreviewLayer();
        var dialogResult = controls.dialog.show();
        removePreviewLayer();
        saveDialogToSession(controls, session);

        return (dialogResult === 1) ? collectDrawOptions(controls) : null;
    }

    // =========================================
    // 角丸アルゴリズム / Round Any Corner
    // 選択したアンカーだけを丸める / rounds only the selected anchors
    // Based on: Hiroyuki Sato (MIT) https://github.com/shspage
    // =========================================

    /**
     * 選択したアンカーの角を半径 rr で丸める
     * @param {PathItem[]} s - 対象のパス
     * @param {Object} conf - 設定（rr: 半径）
     * @returns {void}
     */
    function roundAnyCorner(s, conf) {
        var rr = conf.rr;

        var p, op, pnts;
        var skipList, adjRdirAtEnd, redrawFlg;
        var i, nxi, pvi, q, d, ds, r, g, t, qb;
        var anc1, ldir1, rdir1, anc2, ldir2, rdir2;

        var hanLen = 4 * (Math.sqrt(2) - 1) / 3;
        var ptyp = PointType.SMOOTH;

        for (var j = 0; j < s.length; j++) {
            p = s[j].pathPoints;
            if (readjustAnchors(p) < 2) continue;
            op = !s[j].closed;
            pnts = op ? [getDat(p[0])] : [];
            redrawFlg = false;
            adjRdirAtEnd = 0;

            skipList = [(op || !isSelected(p[0]) || !isCorner(p, 0))];
            for (i = 1; i < p.length; i++) {
                skipList.push((!isSelected(p[i])
                    || !isCorner(p, i)
                    || (op && i == p.length - 1)));
            }

            for (i = 0; i < p.length; i++) {
                nxi = parseIdx(p, i + 1);
                if (nxi < 0) break;

                pvi = parseIdx(p, i - 1);

                q = [p[i].anchor, p[i].rightDirection,
                p[nxi].leftDirection, p[nxi].anchor];

                ds = dist(q[0], q[3]) / 2;
                if (arrEq(q[0], q[1]) && arrEq(q[2], q[3])) {
                    r = Math.min(ds, rr);
                    g = getRad(q[0], q[3]);
                    anc1 = getPnt(q[0], g, r);
                    ldir1 = getPnt(anc1, g + Math.PI, r * hanLen);

                    if (skipList[nxi]) {
                        if (!skipList[i]) {
                            pnts.push([anc1, anc1, ldir1, ptyp]);
                            redrawFlg = true;
                        }
                        pnts.push(getDat(p[nxi]));
                    } else {
                        if (r < rr) {
                            pnts.push([anc1,
                                getPnt(anc1, getRad(ldir1, anc1), r * hanLen),
                                ldir1,
                                ptyp]);
                        } else {
                            if (!skipList[i]) pnts.push([anc1, anc1, ldir1, ptyp]);
                            anc2 = getPnt(q[3], g + Math.PI, r);
                            pnts.push([anc2,
                                getPnt(anc2, g, r * hanLen),
                                anc2,
                                ptyp]);
                        }
                        redrawFlg = true;
                    }
                } else {
                    d = getT4Len(q, 0) / 2;
                    r = Math.min(d, rr);
                    t = getT4Len(q, r);
                    anc1 = bezier(q, t);
                    rdir1 = defHan(t, q, 1);
                    ldir1 = getPnt(anc1, getRad(rdir1, anc1), r * hanLen);

                    if (skipList[nxi]) {
                        if (skipList[i]) {
                            pnts.push(getDat(p[nxi]));
                        } else {
                            pnts.push([anc1, rdir1, ldir1, ptyp]);
                            with (p[nxi]) pnts.push([anchor,
                                rightDirection,
                                adjHan(anchor, leftDirection, 1 - t),
                                ptyp]);
                            redrawFlg = true;
                        }
                    } else {
                        if (r < rr) {
                            if (skipList[i]) {
                                if (!op && i == 0) {
                                    adjRdirAtEnd = t;
                                } else {
                                    pnts[pnts.length - 1][1] = adjHan(q[0], q[1], t);
                                }
                                pnts.push([anc1,
                                    getPnt(anc1, getRad(ldir1, anc1), r * hanLen),
                                    defHan(t, q, 0),
                                    ptyp]);
                            } else {
                                pnts.push([anc1,
                                    getPnt(anc1, getRad(ldir1, anc1), r * hanLen),
                                    ldir1,
                                    ptyp]);
                            }
                        } else {
                            if (skipList[i]) {
                                t = getT4Len(q, -r);
                                anc2 = bezier(q, t);

                                if (!op && i == 0) {
                                    adjRdirAtEnd = t;
                                } else {
                                    pnts[pnts.length - 1][1] = adjHan(q[0], q[1], t);
                                }

                                ldir2 = defHan(t, q, 0);
                                rdir2 = getPnt(anc2, getRad(ldir2, anc2), r * hanLen);

                                pnts.push([anc2, rdir2, ldir2, ptyp]);
                            } else {
                                qb = [anc1, rdir1, adjHan(q[3], q[2], 1 - t), q[3]];
                                t = getT4Len(qb, -r);
                                anc2 = bezier(qb, t);
                                ldir2 = defHan(t, qb, 0);
                                rdir2 = getPnt(anc2, getRad(ldir2, anc2), r * hanLen);
                                rdir1 = adjHan(anc1, rdir1, t);

                                pnts.push([anc1, rdir1, ldir1, ptyp],
                                    [anc2, rdir2, ldir2, ptyp]);
                            }
                        }
                        redrawFlg = true;
                    }
                }
            }
            if (adjRdirAtEnd > 0) {
                pnts[pnts.length - 1][1] = adjHan(p[0].anchor, p[0].rightDirection, adjRdirAtEnd);
            }

            if (redrawFlg) {
                for (i = p.length - 1; i > 0; i--) p[i].remove();

                for (i = 0; i < pnts.length; i++) {
                    var pt = i > 0 ? p.add() : p[0];
                    with (pt) {
                        anchor = pnts[i][0];
                        rightDirection = pnts[i][1];
                        leftDirection = pnts[i][2];
                        pointType = pnts[i][3];
                    }
                }
            }
        }
        app.activeDocument.selection = s;
    }

    /**
     * 点から角度 rad の方向へ len 進んだ点を返す
     * @param {number[]} pt - 起点
     * @param {number} rad - 角度（ラジアン）
     * @param {number} len - 距離
     * @returns {number[]} 点
     */
    function getPnt(pt, rad, len) {
        return [pt[0] + Math.cos(rad) * len,
        pt[1] + Math.sin(rad) * len];
    }

    /**
     * ベジェ曲線の t における接線方向のハンドルを返す
     * @param {number} t - 曲線上の位置（0〜1）
     * @param {Array} q - 制御点4つ
     * @param {number} n - 0 なら前側、1 なら後側
     * @returns {number[]} ハンドルの点
     */
    function defHan(t, q, n) {
        return [t * (t * (q[n][0] - 2 * q[n + 1][0] + q[n + 2][0]) + 2 * (q[n + 1][0] - q[n][0])) + q[n][0],
        t * (t * (q[n][1] - 2 * q[n + 1][1] + q[n + 2][1]) + 2 * (q[n + 1][1] - q[n][1])) + q[n][1]];
    }

    /**
     * ベジェ曲線の t における点を返す
     * @param {Array} q - 制御点4つ
     * @param {number} t - 曲線上の位置（0〜1）
     * @returns {number[]} 点
     */
    function bezier(q, t) {
        var u = 1 - t;
        return [u * u * u * q[0][0] + 3 * u * t * (u * q[1][0] + t * q[2][0]) + t * t * t * q[3][0],
        u * u * u * q[0][1] + 3 * u * t * (u * q[1][1] + t * q[2][1]) + t * t * t * q[3][1]];
    }

    /**
     * ハンドルの長さを m 倍にする
     * @param {number[]} anc - アンカー
     * @param {number[]} dir - ハンドル
     * @param {number} m - 倍率
     * @returns {number[]} 新しいハンドルの点
     */
    function adjHan(anc, dir, m) {
        return [anc[0] + (dir[0] - anc[0]) * m,
        anc[1] + (dir[1] - anc[1]) * m];
    }

    /**
     * アンカーが角（折れ点）か
     * @param {PathPoints} p - パスポイント
     * @param {number} idx - 添字
     * @returns {boolean} 角か
     */
    function isCorner(p, idx) {
        var pnt0 = getAnglePnt(p, idx, -1);
        var pnt1 = getAnglePnt(p, idx, 1);
        if (!pnt0 || !pnt1) return false;
        if (pnt0.length < 1 || pnt1.length < 1) return false;
        var rad = getRad2(pnt0, p[idx].anchor, pnt1, true);
        if (rad > Math.PI - 0.1) return false;
        return true;
    }

    /**
     * 角度を測るための隣の点を返す
     * @param {PathPoints} p - パスポイント
     * @param {number} idx1 - 添字
     * @param {number} dir - -1 なら前、1 なら後
     * @returns {number[]|null} 点。無ければ null、重なっていれば空配列
     */
    function getAnglePnt(p, idx1, dir) {
        if (!dir) dir = -1;
        var idx2 = parseIdx(p, idx1 + dir);
        if (idx2 < 0) return null;
        var p2 = p[idx2];
        with (p[idx1]) {
            if (dir < 0) {
                if (arrEq(leftDirection, anchor)) {
                    if (arrEq(p2.anchor, anchor)) return [];
                    if (arrEq(p2.anchor, p2.rightDirection)
                        || arrEq(p2.rightDirection, anchor)) return p2.anchor;
                    else return p2.rightDirection;
                } else {
                    return leftDirection;
                }
            } else {
                if (arrEq(anchor, rightDirection)) {
                    if (arrEq(anchor, p2.anchor)) return [];
                    if (arrEq(p2.anchor, p2.leftDirection)
                        || arrEq(anchor, p2.leftDirection)) return p2.anchor;
                    else return p2.leftDirection;
                } else {
                    return rightDirection;
                }
            }
        }
    }

    /**
     * 2つの配列の要素が等しいか
     * @param {number[]} arr1 - 配列1
     * @param {number[]} arr2 - 配列2
     * @returns {boolean} 等しいか
     */
    function arrEq(arr1, arr2) {
        for (var i = 0; i < arr1.length; i++) {
            if (arr1[i] != arr2[i]) return false;
        }
        return true;
    }

    /**
     * 2点間の距離
     * @param {number[]} p1 - 点1
     * @param {number[]} p2 - 点2
     * @returns {number} 距離
     */
    function dist(p1, p2) {
        return Math.sqrt(Math.pow(p1[0] - p2[0], 2)
            + Math.pow(p1[1] - p2[1], 2));
    }

    /**
     * 2点間の距離の2乗
     * @param {number[]} p1 - 点1
     * @param {number[]} p2 - 点2
     * @returns {number} 距離の2乗
     */
    function dist2(p1, p2) {
        return Math.pow(p1[0] - p2[0], 2)
            + Math.pow(p1[1] - p2[1], 2);
    }

    /**
     * p1 から p2 への角度
     * @param {number[]} p1 - 起点
     * @param {number[]} p2 - 終点
     * @returns {number} 角度（ラジアン）
     */
    function getRad(p1, p2) {
        return Math.atan2(p2[1] - p1[1],
            p2[0] - p1[0]);
    }

    /**
     * o を頂点とする p1-o-p2 の角度
     * @param {number[]} p1 - 点1
     * @param {number[]} o - 頂点
     * @param {number[]} p2 - 点2
     * @returns {number} 角度（ラジアン）
     */
    function getRad2(p1, o, p2) {
        var v1 = normalize(p1, o);
        var v2 = normalize(p2, o);
        return Math.acos(v1[0] * v2[0] + v1[1] * v2[1]);
    }

    /**
     * o から p への単位ベクトル
     * @param {number[]} p - 点
     * @param {number[]} o - 原点
     * @returns {number[]} 単位ベクトル
     */
    function normalize(p, o) {
        var d = dist(p, o);
        return d == 0 ? [0, 0] : [(p[0] - o[0]) / d,
        (p[1] - o[1]) / d];
    }

    /**
     * ベジェ曲線上で長さ len の位置の t を返す（len が 0 なら全長）
     * @param {Array} q - 制御点4つ
     * @param {number} len - 長さ（負なら終点側から）
     * @returns {number} t、または全長
     */
    function getT4Len(q, len) {
        var m = [q[3][0] - q[0][0] + 3 * (q[1][0] - q[2][0]),
        q[0][0] - 2 * q[1][0] + q[2][0],
        q[1][0] - q[0][0]];
        var n = [q[3][1] - q[0][1] + 3 * (q[1][1] - q[2][1]),
        q[0][1] - 2 * q[1][1] + q[2][1],
        q[1][1] - q[0][1]];
        var k = [m[0] * m[0] + n[0] * n[0],
        4 * (m[0] * m[1] + n[0] * n[1]),
        2 * ((m[0] * m[2] + n[0] * n[2]) + 2 * (m[1] * m[1] + n[1] * n[1])),
        4 * (m[1] * m[2] + n[1] * n[2]),
        m[2] * m[2] + n[2] * n[2]];

        var fullLen = getLength(k, 1);

        if (len == 0) {
            return fullLen;
        } else if (len < 0) {
            len += fullLen;
            if (len < 0) return 0;
        } else if (len > fullLen) {
            return 1;
        }

        var t, d;
        var t0 = 0;
        var t1 = 1;
        var torelance = 0.001;

        for (var h = 1; h < 30; h++) {
            t = t0 + (t1 - t0) / 2;
            d = len - getLength(k, t);
            if (Math.abs(d) < torelance) break;
            else if (d < 0) t1 = t;
            else t0 = t;
        }
        return t;
    }

    /**
     * ベジェ曲線の 0〜t の長さ（シンプソン則）
     * @param {number[]} k - 係数
     * @param {number} t - 終わりの位置
     * @returns {number} 長さ
     */
    function getLength(k, t) {
        var h = t / 128;
        var hh = h * 2;
        var fc = function (t, k) {
            return Math.sqrt(t * (t * (t * (t * k[0] + k[1]) + k[2]) + k[3]) + k[4]) || 0;
        };
        var total = (fc(0, k) - fc(t, k)) / 2;
        for (var i = h; i < t; i += hh) total += 2 * fc(i, k) + fc(i + h, k);
        return total * hh;
    }

    /**
     * 重なったアンカーを1つにまとめる
     * @param {PathPoints} p - パスポイント
     * @returns {number} まとめたあとの点の数
     */
    function readjustAnchors(p) {
        var minDist = 0.0025;
        if (p.length < 2) return 1;
        var i;

        if (p.parent.closed) {
            for (i = p.length - 1; i >= 1; i--) {
                if (dist2(p[0].anchor, p[i].anchor) < minDist) {
                    p[0].leftDirection = p[i].leftDirection;
                    p[i].remove();
                } else {
                    break;
                }
            }
        }

        for (i = p.length - 1; i >= 1; i--) {
            if (dist2(p[i].anchor, p[i - 1].anchor) < minDist) {
                p[i - 1].rightDirection = p[i].rightDirection;
                p[i].remove();
            }
        }

        return p.length;
    }

    /**
     * 添字を点の数の範囲に収める（閉じたパスは巡回、開いたパスは範囲外で -1）
     * @param {PathPoints} p - パスポイント
     * @param {number} n - 添字
     * @returns {number} 収めた添字
     */
    function parseIdx(p, n) {
        var len = p.length;
        if (p.parent.closed) {
            return n >= 0 ? n % len : len - Math.abs(n % len);
        } else {
            return (n < 0 || n > len - 1) ? -1 : n;
        }
    }

    /**
     * パスポイントの座標と種類を配列で返す
     * @param {PathPoint} p - パスポイント
     * @returns {Array} [アンカー, 右ハンドル, 左ハンドル, 種類]
     */
    function getDat(p) {
        with (p) return [anchor, rightDirection, leftDirection, pointType];
    }

    /**
     * アンカーが選択されているか
     * @param {PathPoint} p - パスポイント
     * @returns {boolean} 選択されているか
     */
    function isSelected(p) {
        return p.selected == PathPointSelection.ANCHORPOINT;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 2つのオブジェクトを左から右（上下なら上から下）の順に並べ替える
     * @returns {void}
     */
    function orderTargetsByPosition() {
        var bounds1 = targetBounds[0];
        var bounds2 = targetBounds[1];
        var secondComesFirst = isVerticalSplit ? (bounds1[1] < bounds2[1]) : (bounds1[0] > bounds2[0]);
        if (secondComesFirst) {
            targetItems.reverse();
            targetBounds.reverse();
        }
    }

    /**
     * 選択した2つのオブジェクトの背面に、2分割の背景を作る
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.openDocument"));
            return;
        }

        doc = app.activeDocument;
        var selectedItems = doc.selection;
        /* 文字の編集中は TextRange が返り、length は文字数になる / while editing text, the selection is a TextRange whose length counts characters */
        if (!selectedItems || selectedItems.typename === "TextRange" || selectedItems.length !== 2) {
            alert(getLabel("alert.selectTwoItems"));
            return;
        }

        targetItems = [selectedItems[0], selectedItems[1]];
        targetBounds = [getItemBounds(targetItems[0]), getItemBounds(targetItems[1])];
        /* すき間の大きいほうの向きに分ける（同じなら左右）/ split along the larger gap; a tie goes left/right */
        isVerticalSplit = (getGapPt(targetBounds, true) > getGapPt(targetBounds, false));
        orderTargetsByPosition();

        var originalActiveLayer = doc.activeLayer;
        var drawOptions = showSettingsDialog();
        if (drawOptions) {
            drawSplitBackground(drawOptions, getBackmostLayer(targetItems[0].layer, targetItems[1].layer));
            /* 全体の枠と区切り線より前面に元のオブジェクトを出す / bring the objects in front of the frame and divider */
            targetItems[0].zOrder(ZOrderMethod.BRINGTOFRONT);
            targetItems[1].zOrder(ZOrderMethod.BRINGTOFRONT);
            doc.selection = null;
        } else {
            /* キャンセルしたら選択を戻す / restore the selection on cancel */
            doc.selection = targetItems;
        }

        try {
            doc.activeLayer = originalActiveLayer;
        } catch (e) {
            /* ロックされたレイヤーなどは戻せないことがある / a locked layer may refuse to become active again */
        }
    }

    main();

})();
