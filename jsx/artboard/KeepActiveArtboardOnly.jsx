#target illustrator
#targetengine "KeepActiveArtboardOnlyEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.KeepActiveArtboardOnly
Meta-Default-language=ja
Name=アクティブなアートボードだけを残す
Name[en]=Keep only the active artboard
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.2.13
Release-Date=2025-08-15
Tags=artboard,アートボード
X-com.swwwitch-Script-Name=KeepActiveArtboardOnly
X-com.swwwitch-Updated=2026-11-20
X-com.swwwitch-README-JA=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/KeepActiveArtboardOnly.md
X-com.swwwitch-README-EN=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/KeepActiveArtboardOnly.md
X-com.swwwitch-Source-URL=https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/KeepActiveArtboardOnly.jsx
Description-BEGIN
[ja]
アクティブなアートボードだけを残すか、空のアートボードをまとめて削除します。
アクティブなアートボードだけを残す場合は、その外側にあるオブジェクト（ガイドを含む）も削除できます。
[en]
Keeps only the active artboard, or removes every empty artboard in one pass.
When the active artboard is kept, the objects outside it — guides included — can be deleted as well.
Description-END
SCRIPTMETA-END
*/

// =========================================
// 基本情報 / Basic info
// =========================================
// スクリプト情報（再利用パーツ） / Script info (reusable)
/* 基本情報は先頭の SCRIPTMETA が元。実行時に読み、getScriptMetaValue("Version") のように直接読む（値は SCRIPTMETA を書き換え、helpers/sync-scriptmeta.py で同期）
   The basic info lives in the SCRIPTMETA block at the top; read its fields directly with getScriptMetaValue() */

/**
 * このスクリプトの先頭にある SCRIPTMETA の項目（Description の手前まで）を読む
 * @returns {Object} 項目名 → 値（同じ項目は最初の値）。ファイルが読めなければ空のオブジェクト
 */
function readScriptMeta() {
    var metaFields = {};
    var scriptFile = new File($.fileName);
    scriptFile.encoding = "UTF-8";
    if (!scriptFile.exists || !scriptFile.open("r")) return metaFields;
    /* 境界の文字列は分けて書く（コードの中に境界があると SCRIPTMETA を読むアプリが誤読するため）
       Marker strings are split so apps that scan for SCRIPTMETA do not find them in the code */
    var blockBegin = "SCRIPTMETA" + "-BEGIN";
    var blockEnd = "SCRIPTMETA" + "-END";
    var descriptionBegin = "Description" + "-BEGIN";
    var isInBlock = false;
    for (var lineCount = 0; lineCount < 200 && !scriptFile.eof; lineCount++) {
        var metaLine = scriptFile.readln().replace(/\r$/, "");
        if (metaLine === blockBegin) {
            isInBlock = true;
        } else if (isInBlock) {
            if (metaLine === blockEnd || metaLine === descriptionBegin) break;
            var separatorIndex = metaLine.indexOf("=");
            var fieldKey = metaLine.substring(0, separatorIndex);
            if (separatorIndex > 0 && !metaFields.hasOwnProperty(fieldKey)) metaFields[fieldKey] = metaLine.substring(separatorIndex + 1);
        }
    }
    scriptFile.close();
    return metaFields;
}

var SCRIPT_META = readScriptMeta();

/**
 * SCRIPTMETA の項目の値を返す
 * @param {string} fieldKey - 項目名（"Version"、"X-com.swwwitch-Updated" など）
 * @returns {string} 値。項目が無ければ空文字
 */
function getScriptMetaValue(fieldKey) {
    return SCRIPT_META.hasOwnProperty(fieldKey) ? SCRIPT_META[fieldKey] : "";
}

/* SCRIPTMETA が読めないとき（ファイルを経由しない実行）は、スクリプト名だけファイル名から補う / fall back to the file name */
var SCRIPT_NAME = getScriptMetaValue("X-com.swwwitch-Script-Name") || decodeURI(new File($.fileName).name).replace(/\.[^.]+$/, ""); /* スクリプト名（設定の保存名） / script name, the key for saved settings */
// スクリプト情報（再利用パーツ）ここまで / End of the reusable script info

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* アクティブなアートボード以外を削除するとき、外側のオブジェクトも削除するか / Delete the objects outside the kept artboard */
    var DEFAULT_DELETE_OUTSIDE_OBJECTS = true;

    /* 空のアートボードを探すとき、非表示のレイヤー・オブジェクトを無視するか / Ignore hidden layers and objects while looking for empty artboards */
    var DEFAULT_IGNORE_HIDDEN = true;

    // =========================================
    // レイアウト / Layout
    // =========================================

    // UIレイアウト（再利用パーツ） / UI layout (reusable)

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
    var TAB_MARGINS    = [15, 20, 5, 10];    /* タブ余白 [左,上,右,下] / tab margins */

    /**
     * ウィンドウの共通設定
     * @param {Window} targetWindow - 対象のウィンドウ
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(targetWindow, spacing) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルの共通設定（子は幅いっぱい。ボタンは alignment = "left" で広げない）
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * タブの共通設定
     * @param {Tab} targetTab - 対象のタブ
     * @param {number} [spacing] - 要素間隔（省略時は変えない）
     * @returns {void}
     */
    function setupTab(targetTab, spacing) {
        targetTab.orientation = "column";
        targetTab.alignChildren = "fill";
        targetTab.margins = TAB_MARGINS;
        if (typeof spacing === "number") targetTab.spacing = spacing;
    }

    /**
     * 横並びの行グループの共通設定（ボタン列など）。
     * alignment と alignChildren を対で指定し、中のボタンが横に伸びたり天地がずれたりしないようにする
     * @param {Group} rowGroup - 対象のグループ
     * @param {string|string[]} [rowAlignment] - 横方向の alignment（省略時は "left"）。配列ならそのまま使う
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(rowGroup, rowAlignment, spacing) {
        rowGroup.orientation = "row";
        rowGroup.alignment = (rowAlignment instanceof Array) ? rowAlignment : [rowAlignment || "left", "center"];
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ボタンの高さを指定した px だけ詰める（レイアウトが決まったあとに呼ぶ）
     * @param {Button} targetButton - 対象のボタン
     * @param {number} trimPixels - 詰める量（px）
     * @returns {void}
     */
    function trimButtonHeight(targetButton, trimPixels) {
        /* レイアウト前は size が無い / size is not set until the layout runs */
        if (!targetButton.size) return;
        targetButton.size = [targetButton.size.width, targetButton.size.height - trimPixels];
    }

    // UIレイアウト（再利用パーツ）ここまで / End of the reusable UI layout

    var OPTION_INDENT_MARGINS = [18, 0, 0, 2];  /* ラジオ直下のオプションのインデント */
    var COUNT_TEXT_WIDTH = 220;                 /* 件数表示の幅（後から伸びないため確保） */
    var BUTTON_WIDTH = 80;                      /* ボタン幅 */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ローカライズ（再利用パーツ） / Localization (reusable)

    /**
     * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /**
     * LABELS から今の UI 言語の文言を取り出す。
     * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
     * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
     * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
     */
    function getLabel(labelRef, placeholderValues) {
        var labelEntry = labelRef;
        if (typeof labelRef === "string") {
            var labelPathKeys = labelRef.split(".");
            labelEntry = LABELS;
            for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
                labelEntry = labelEntry[labelPathKeys[i]];
            }
        }
        var labelString;
        if (typeof labelEntry === "string") labelString = labelEntry;
        else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
        else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
        else return (typeof labelRef === "string") ? labelRef : "";
        return fillLabelPlaceholders(String(labelString), placeholderValues);
    }

    /**
     * 項目名の文言の末尾にコロンを付ける（日本語は半角スペース＋半角コロン「 :」、英語は「:」。Illustrator の線パネルなどの項目名に合わせる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? " :" : ":");
    }

    /**
     * 「項目名 : 値」の1行を返す（日本語は「件数 : 5」、英語は「Count: 5」。どちらもコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + " " + value;
    }

    /**
     * 文言の {name} や %1 に値を差し込む
     * @param {string} labelString - 文言
     * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
     * @returns {string} 差し込んだ文言
     */
    function fillLabelPlaceholders(labelString, placeholderValues) {
        if (placeholderValues == null) return labelString;
        if (placeholderValues instanceof Array) {
            /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
            for (var i = placeholderValues.length; i >= 1; i--) {
                labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
            }
            return labelString;
        }
        for (var placeholderKey in placeholderValues) {
            if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
            labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
        }
        return labelString;
    }

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "アートボードの削除", en: "Remove Artboards" }
        },
        panel: {
            mode: { ja: "削除するアートボード", en: "Artboards to remove" }
        },
        radio: {
            keepActive: { ja: "アクティブなアートボード以外", en: "All but the active artboard" },
            empty:      { ja: "空のアートボード", en: "Empty artboards" }
        },
        checkbox: {
            deleteOutsideObjects: { ja: "アートボード外のオブジェクトも削除", en: "Also delete the objects outside it" },
            ignoreHidden:         { ja: "非表示レイヤー、非表示オブジェクトは無視", en: "Ignore hidden layers and hidden objects" }
        },
        fieldLabel: {
            removalTarget: { ja: "削除対象アートボード数", en: "Artboards to remove" }
        },
        tooltip: {
            keepActive: {
                ja: "アクティブなアートボードだけを残し、ほかのアートボードを削除します。",
                en: "Keeps just the active artboard and removes the others."
            },
            empty: {
                ja: "オブジェクトが載っていないアートボードをまとめて削除します。",
                en: "Removes every artboard with nothing on it."
            },
            deleteOutsideObjects: {
                ja: "残したアートボードの外側にあるオブジェクトとオブジェクトガイドも削除します。",
                en: "Also deletes the objects and object guides that sit outside the artboard you keep."
            },
            ignoreHidden: {
                ja: "非表示のレイヤーやオブジェクトしか載っていないアートボードも「空」とみなします。オフにすると、見えていなくても中身があれば残します。",
                en: "Treats an artboard holding only hidden layers or objects as empty. Off keeps it as long as anything is on it."
            },
            singleArtboard: {
                ja: "アートボードが1枚しかないため削除できません。",
                en: "Only one artboard exists; cannot remove."
            }
        },
        button: {
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." }
        }
    };

    // =========================================
    // 重なり判定 / Hit testing
    // =========================================

    /* 共線とみなす外積の許容値 / Cross product below this counts as collinear */
    var COLLINEAR_EPSILON = 1e-6;

    /**
     * 点が矩形の内側（境界を含む）にあるか判定する
     * @param {number} x - X座標
     * @param {number} y - Y座標
     * @param {number[]} rect - 矩形 [left, top, right, bottom]
     * @returns {boolean} 内側にあれば true
     */
    function isPointInRect(x, y, rect) {
        return (x >= rect[0] && x <= rect[2] && y <= rect[1] && y >= rect[3]);
    }

    /**
     * 2Dベクトルの外積を返す
     * @param {number} ax - ベクトルAのX成分
     * @param {number} ay - ベクトルAのY成分
     * @param {number} bx - ベクトルBのX成分
     * @param {number} by - ベクトルBのY成分
     * @returns {number} 外積
     */
    function cross2D(ax, ay, bx, by) {
        return ax * by - ay * bx;
    }

    /**
     * 線分同士が交差するか判定する（共線で重なる場合も true）
     * @param {number} p1x - 線分Pの始点X
     * @param {number} p1y - 線分Pの始点Y
     * @param {number} p2x - 線分Pの終点X
     * @param {number} p2y - 線分Pの終点Y
     * @param {number} q1x - 線分Qの始点X
     * @param {number} q1y - 線分Qの始点Y
     * @param {number} q2x - 線分Qの終点X
     * @param {number} q2y - 線分Qの終点Y
     * @returns {boolean} 交差していれば true
     */
    function segmentsIntersect(p1x, p1y, p2x, p2y, q1x, q1y, q2x, q2y) {
        var pDx = p2x - p1x,
            pDy = p2y - p1y;
        var qDx = q2x - q1x,
            qDy = q2y - q1y;
        var denominator = cross2D(pDx, pDy, qDx, qDy);
        var p1ToQ1Dx = q1x - p1x,
            p1ToQ1Dy = q1y - p1y;

        if (denominator === 0) {
            /* 平行。共線でなければ交差しない / Parallel: no intersection unless collinear */
            if (Math.abs(cross2D(p1ToQ1Dx, p1ToQ1Dy, pDx, pDy)) > COLLINEAR_EPSILON) return false;
            /* 共線なら各軸への投影で重なりを見る / Collinear: check the overlap on each axis */
            var overlapX = !(Math.max(p1x, p2x) < Math.min(q1x, q2x) || Math.max(q1x, q2x) < Math.min(p1x, p2x));
            var overlapY = !(Math.max(p1y, p2y) < Math.min(q1y, q2y) || Math.max(q1y, q2y) < Math.min(p1y, p2y));
            return overlapX && overlapY;
        }

        var tParam = cross2D(p1ToQ1Dx, p1ToQ1Dy, qDx, qDy) / denominator;
        var uParam = cross2D(p1ToQ1Dx, p1ToQ1Dy, pDx, pDy) / denominator;
        return (tParam >= 0 && tParam <= 1 && uParam >= 0 && uParam <= 1);
    }

    /**
     * 線分が矩形と交差するか判定する
     * @param {number} x1 - 線分の始点X
     * @param {number} y1 - 線分の始点Y
     * @param {number} x2 - 線分の終点X
     * @param {number} y2 - 線分の終点Y
     * @param {number[]} rect - 矩形 [left, top, right, bottom]
     * @returns {boolean} 交差していれば true
     */
    function segmentIntersectsRect(x1, y1, x2, y2, rect) {
        /* 端点のどちらかが矩形内なら交差 / Either endpoint inside means a hit */
        if (isPointInRect(x1, y1, rect) || isPointInRect(x2, y2, rect)) return true;

        var left = rect[0],
            top = rect[1],
            right = rect[2],
            bottom = rect[3];

        return segmentsIntersect(x1, y1, x2, y2, left, top, right, top) ||
            segmentsIntersect(x1, y1, x2, y2, left, bottom, right, bottom) ||
            segmentsIntersect(x1, y1, x2, y2, left, top, left, bottom) ||
            segmentsIntersect(x1, y1, x2, y2, right, top, right, bottom);
    }

    /**
     * パスがアートボードと重なるか判定する（境界ボックスではなくパスの実体で判定）
     * @param {PathItem} pathItem - 判定するパス
     * @param {number[]} artboardRect - アートボードの矩形 [left, top, right, bottom]
     * @returns {boolean} 重なっていれば true
     */
    function pathIntersectsArtboard(pathItem, artboardRect) {
        try {
            var pathPoints = pathItem.pathPoints;
            if (!pathPoints || pathPoints.length === 0) return false;

            for (var i = 0; i < pathPoints.length; i++) {
                /* 開パスでは最後のアンカーから始点へは結ばない / Open paths do not wrap back to the start */
                if (!pathItem.closed && i === pathPoints.length - 1) break;
                var currentAnchor = pathPoints[i].anchor;
                var nextAnchor = pathPoints[(i + 1) % pathPoints.length].anchor;
                if (segmentIntersectsRect(currentAnchor[0], currentAnchor[1], nextAnchor[0], nextAnchor[1], artboardRect)) return true;
            }
            return false;
        } catch (e) {
            return false;
        }
    }

    /**
     * ページアイテムがアートボードと重なるか判定する
     * パス・複合パスはパスの実体で、それ以外は境界ボックスで判定する
     * @param {PageItem} pageItem - 判定するアイテム
     * @param {number[]} artboardRect - アートボードの矩形 [left, top, right, bottom]
     * @returns {boolean} 重なっていれば true
     */
    function itemIntersectsArtboard(pageItem, artboardRect) {
        if (pageItem.typename === 'CompoundPathItem') {
            var subPaths = pageItem.pathItems;
            for (var i = 0; i < subPaths.length; i++) {
                if (pathIntersectsArtboard(subPaths[i], artboardRect)) return true;
            }
            return false;
        }
        if (pageItem.typename === 'PathItem') {
            return pathIntersectsArtboard(pageItem, artboardRect);
        }
        /* その他は境界ボックスでフォールバック / Fall back to the bounding box */
        return rectsOverlap(artboardRect, pageItem.geometricBounds);
    }

    /**
     * アートボードと境界ボックスが重なるか判定する
     * @param {number[]} artboardRect - アートボードの矩形 [left, top, right, bottom]
     * @param {number[]} itemBounds - アイテムの境界 [left, top, right, bottom]
     * @returns {boolean} 重なっていれば true
     */
    function rectsOverlap(artboardRect, itemBounds) {
        return (
            itemBounds[0] < artboardRect[2] &&
            itemBounds[2] > artboardRect[0] &&
            itemBounds[1] > artboardRect[3] &&
            itemBounds[3] < artboardRect[1]
        );
    }

    // =========================================
    // ロック・表示状態の一時解除 / Relaxing locks and visibility
    // =========================================

    /**
     * レイヤーの状態を控えてロック・非表示・テンプレートを解除する（サブレイヤーとグループも再帰）
     * @param {Layer} layer - 対象レイヤー
     * @param {object[]} stateLog - 状態の控え
     * @returns {void}
     */
    function captureAndUnlockLayer(layer, stateLog) {
        stateLog.push({
            type: 'Layer',
            ref: layer,
            locked: layer.locked,
            visible: layer.visible,
            template: layer.template
        });

        layer.locked = false;
        layer.visible = true;
        if (layer.template === true) layer.template = false;

        for (var i = 0; i < layer.layers.length; i++) {
            captureAndUnlockLayer(layer.layers[i], stateLog);
        }
        for (var j = 0; j < layer.groupItems.length; j++) {
            captureAndUnlockGroup(layer.groupItems[j], stateLog);
        }
        captureAndUnlockItems(layer, stateLog);
    }

    /**
     * グループの状態を控えてロック・非表示を解除する（ネストしたグループも再帰）
     * @param {GroupItem} groupItem - 対象グループ
     * @param {object[]} stateLog - 状態の控え
     * @returns {void}
     */
    function captureAndUnlockGroup(groupItem, stateLog) {
        stateLog.push({
            type: 'Group',
            ref: groupItem,
            locked: groupItem.locked,
            visible: groupItem.visible
        });

        groupItem.locked = false;
        groupItem.visible = true;

        for (var i = 0; i < groupItem.groupItems.length; i++) {
            captureAndUnlockGroup(groupItem.groupItems[i], stateLog);
        }
        captureAndUnlockItems(groupItem, stateLog);
    }

    /**
     * コンテナ配下のページアイテムの状態を控えてロック・非表示を解除する
     * @param {Layer|GroupItem} container - レイヤーまたはグループ
     * @param {object[]} stateLog - 状態の控え
     * @returns {void}
     */
    function captureAndUnlockItems(container, stateLog) {
        try {
            var pageItems = container.pageItems;
            for (var i = 0; i < pageItems.length; i++) {
                var pageItem = pageItems[i];
                stateLog.push({
                    type: 'Item',
                    ref: pageItem,
                    locked: pageItem.locked,
                    hidden: pageItem.hidden
                });
                if (pageItem.locked) pageItem.locked = false;
                if (pageItem.hidden) pageItem.hidden = false;
            }
        } catch (e) { }
    }

    /**
     * ドキュメント全体のロック・表示状態を控えて一時解除する
     * @param {Document} doc - 対象ドキュメント
     * @returns {object[]} 復元用の状態の控え
     */
    function captureAndUnlockStructure(doc) {
        var stateLog = [];
        for (var i = 0; i < doc.layers.length; i++) {
            captureAndUnlockLayer(doc.layers[i], stateLog);
        }
        return stateLog;
    }

    /**
     * 控えた状態を復元する（下位から順に戻す）
     * @param {object[]} savedStates - captureAndUnlockStructure() が返した控え
     * @returns {void}
     */
    function restoreStructure(savedStates) {
        for (var i = savedStates.length - 1; i >= 0; i--) {
            var savedState = savedStates[i];
            /* 削除済みのオブジェクトは参照できない / deleted objects can no longer be restored */
            try {
                savedState.ref.locked = savedState.locked;
                if (savedState.type === 'Item') {
                    savedState.ref.hidden = savedState.hidden;
                } else {
                    savedState.ref.visible = savedState.visible;
                }
                if (savedState.type === 'Layer') savedState.ref.template = savedState.template;
            } catch (e) { }
        }
    }

    // =========================================
    // 空アートボードの判定 / Empty artboard detection
    // =========================================

    /**
     * オブジェクトと祖先がすべて表示状態か判定する
     * @param {PageItem} pageItem - 対象アイテム
     * @returns {boolean} 表示されていれば true
     */
    function isItemVisible(pageItem) {
        try {
            var ancestorNode = pageItem;
            while (ancestorNode && ancestorNode.typename !== "Document") {
                if (ancestorNode.typename === "Layer") {
                    if (ancestorNode.visible === false) return false;
                } else if (ancestorNode.hidden) {
                    return false;
                }
                ancestorNode = ancestorNode.parent;
            }
        } catch (err) {
            return false;
        }
        return true;
    }

    /**
     * アートボード上に占有オブジェクトがあるか判定する
     * @param {Document} doc - 対象ドキュメント
     * @param {number[]} artboardRect - アートボードの矩形 [left, top, right, bottom]
     * @param {boolean} ignoreHidden - 非表示のレイヤー・オブジェクトを無視するか
     * @returns {boolean} 占有オブジェクトがあれば true
     */
    function hasOccupantOnArtboard(doc, artboardRect, ignoreHidden) {
        var pageItems = doc.pageItems;
        for (var i = 0; i < pageItems.length; i++) {
            var pageItem = pageItems[i];
            var visible = isItemVisible(pageItem);
            if (ignoreHidden && !visible) continue;
            try {
                var bounds = visible ? pageItem.visibleBounds : pageItem.geometricBounds;
                if (rectsOverlap(artboardRect, bounds)) return true;
            } catch (err) {
                /* 境界を取得できないアイテムは無視 / Skip items without bounds */
            }
        }
        return false;
    }

    /**
     * 空アートボードのインデックス一覧を返す
     * @param {Document} doc - 対象ドキュメント
     * @param {boolean} ignoreHidden - 非表示のレイヤー・オブジェクトを無視するか
     * @returns {number[]} 空アートボードのインデックス
     */
    function findEmptyArtboardIndices(doc, ignoreHidden) {
        var indices = [];
        for (var i = 0; i < doc.artboards.length; i++) {
            if (!hasOccupantOnArtboard(doc, doc.artboards[i].artboardRect, ignoreHidden)) {
                indices.push(i);
            }
        }
        return indices;
    }

    // =========================================
    // 削除処理 / Removal
    // =========================================

    /**
     * 指定インデックスのアートボードを削除する（最後の1枚は残す）
     * @param {Document} doc - 対象ドキュメント
     * @param {number[]} indices - 削除するアートボードのインデックス
     * @returns {number} 削除した枚数
     */
    function removeArtboardsByIndices(doc, indices) {
        var artboards = doc.artboards;
        var removedCount = 0;
        /* 末尾から削除してインデックスのずれを避ける / Remove from the end so the indices stay valid */
        for (var i = indices.length - 1; i >= 0; i--) {
            if (artboards.length <= 1) break;
            try {
                artboards[indices[i]].remove();
                removedCount++;
            } catch (err) { }
        }
        return removedCount;
    }

    /**
     * アクティブなアートボードと重ならないオブジェクトガイドを削除する（ルーラーガイドは対象外）
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function removeGuidesOutsideArtboard(doc) {
        try {
            var artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;

            var guidesToRemove = [];
            for (var i = doc.pageItems.length - 1; i >= 0; i--) {
                var pageItem = doc.pageItems[i];
                if (pageItem.guides !== true) continue;
                if (!itemIntersectsArtboard(pageItem, artboardRect)) guidesToRemove.push(pageItem);
            }
            for (var j = 0; j < guidesToRemove.length; j++) guidesToRemove[j].remove();
        } catch (e) {
            /* 取得できないアイテムがあっても可能な範囲で続ける / Best effort */
        }
    }

    /**
     * アクティブなアートボードの外側にあるオブジェクトを削除する
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function deleteObjectsOutsideActiveArtboard(doc) {
        /* ロック・非表示・テンプレートのままでは選択できないため一時解除 / Relax locks and visibility first */
        var savedStructureStates = captureAndUnlockStructure(doc);

        removeGuidesOutsideArtboard(doc);

        /* アートボード上を選択→反転→削除 / Select on the artboard, invert, then delete */
        app.executeMenuCommand('selectallinartboard');
        app.executeMenuCommand('Inverse menu item');
        app.executeMenuCommand('clear');

        restoreStructure(savedStructureStates);
    }

    /**
     * アクティブなアートボードだけを残す
     * @param {Document} doc - 対象ドキュメント
     * @param {boolean} deleteOutsideObjects - 外側のオブジェクトも削除するか
     * @returns {void}
     */
    function keepActiveArtboardOnly(doc, deleteOutsideObjects) {
        if (doc.artboards.length > 1) {
            var activeArtboardIndex = doc.artboards.getActiveArtboardIndex();
            /* 末尾から削除してインデックスのずれを避ける / Remove from the end so the indices stay valid */
            for (var i = doc.artboards.length - 1; i >= 0; i--) {
                if (i !== activeArtboardIndex) doc.artboards.remove(i);
            }
        }

        if (deleteOutsideObjects) deleteObjectsOutsideActiveArtboard(doc);
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)

    var DIALOG_OPACITY = 0.98;       /* ダイアログの不透明度 / dialog opacity */
    var DIALOG_AVOID_MARGIN = 60;    /* 選択範囲の推定位置の両側に取る余裕（px）/ margin on each side of the estimated selection (px) */
    var DIALOG_AVOID_MAX_ITEMS = 100; /* 選択範囲を測るオブジェクトの上限 / max items measured for the selection bounds */

    /**
     * ダイアログの不透明度を設定し、前回閉じた位置で開いて、動かした位置を記録するようにする。
     * 開く位置が選択中のオブジェクトに重なりそうなときは、左右の反対側へずらす（Illustrator のみ）。
     * 既存の onShow / onMove / onClose は先に呼んでから、位置の復元・記録を行う。
     * @param {Window} dialog - 対象のダイアログ
     * @param {string} storageKey - 位置を覚えるキー（ふつうは SCRIPT_NAME）
     * @returns {void}
     */
    function prepareDialogWindow(dialog, storageKey) {
        /* 同じダイアログを開き直すときは、選択範囲を測り直すだけにする（ハンドラーを重ねない）
           When the same dialog is shown again, only re-measure the selection (don't stack handlers) */
        if (dialog.dialogWindowState) {
            dialog.dialogWindowState.selectionSpan = getSelectionViewSpan();
            dialog.dialogWindowState.avoidedLocation = null;
            return;
        }
        var locationKey = "__" + storageKey + "_DialogLocation";
        var previousOnShow = dialog.onShow;
        var previousOnMove = dialog.onMove;
        var previousOnClose = dialog.onClose;
        var windowState = {
            selectionSpan: getSelectionViewSpan(), /* 選択範囲は show() の前に測る / measured before show() */
            screenWidth: null,                     /* 最初に開いたときに推定する / estimated on the first show */
            avoidedLocation: null                  /* 避けるためにずらした位置（記録しない）/ location set to avoid the selection (not remembered) */
        };
        dialog.dialogWindowState = windowState;

        dialog.opacity = DIALOG_OPACITY;

        /* 今の位置を記録する / Remember the current location */
        function rememberDialogLocation() {
            var currentLocation = [dialog.location[0], dialog.location[1]];
            var avoidedLocation = windowState.avoidedLocation;
            if (avoidedLocation && currentLocation[0] === avoidedLocation[0] && currentLocation[1] === avoidedLocation[1]) return;
            $.global[locationKey] = currentLocation;
        }

        dialog.onShow = function () {
            /* 最初に開くときの既定の位置は画面の横中央なので、画面の幅を逆算できる。2回目からは前回の位置なので使い回す
               On the first show the default location is centered horizontally, which gives the screen width; reuse it afterwards */
            if (windowState.screenWidth === null) windowState.screenWidth = dialog.location[0] * 2 + dialog.bounds.width;
            if (previousOnShow) previousOnShow.apply(this, arguments);
            /* $.screens は実際の画面の大きさと合わない（Mac で 1280×524 など）ので、画面内かは判定しない
               $.screens does not match the real display (e.g. 1280x524 on a Mac), so no on-screen check */
            var savedLocation = $.global[locationKey];
            if (savedLocation) dialog.location = [savedLocation[0], savedLocation[1]];
            if (windowState.selectionSpan) {
                var avoidLeft = findDialogLeftAvoidingSelection(dialog.location[0], dialog.bounds.width, windowState.screenWidth, windowState.selectionSpan);
                if (avoidLeft !== null) {
                    dialog.location = [avoidLeft, dialog.location[1]];
                    /* 代入後の値で比べる（丸められることがある）/ Compare with the value after assignment, which may be rounded */
                    windowState.avoidedLocation = [dialog.location[0], dialog.location[1]];
                }
            }
        };
        dialog.onMove = function () {
            if (previousOnMove) previousOnMove.apply(this, arguments);
            rememberDialogLocation();
        };
        dialog.onClose = function () {
            rememberDialogLocation();
            /* false を返すと閉じるのを取りやめるので、戻り値は元の onClose のものを返す
               Returning false cancels the close, so pass the original onClose result through */
            if (previousOnClose) return previousOnClose.apply(this, arguments);
        };
    }

    /**
     * 選択中のオブジェクトが、ドキュメントの表示域の左端から画面上で何 px の範囲にあるかを返す。
     * @returns {{left: number, right: number, viewWidth: number}|null} 選択が無い・測れないときは null
     */
    function getSelectionViewSpan() {
        try {
            if (app.name !== "Adobe Illustrator" || !app.documents.length) return null;
            var targetDoc = app.activeDocument;
            var selectedItems = targetDoc.selection;
            /* 文字ツールで文字を選択しているときは TextRange が返り、[0] が無い / Selecting characters with the Type tool returns a TextRange, which has no [0] */
            if (!selectedItems || selectedItems.typename === "TextRange" || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
            var itemCount = Math.min(selectedItems.length, DIALOG_AVOID_MAX_ITEMS);
            var spanLeft = Infinity;
            var spanRight = -Infinity;
            for (var i = 0; i < itemCount; i++) {
                var itemBounds = selectedItems[i].visibleBounds;
                if (itemBounds[0] < spanLeft) spanLeft = itemBounds[0];
                if (itemBounds[2] > spanRight) spanRight = itemBounds[2];
            }
            var activeView = targetDoc.activeView; /* 複数ウィンドウで開いていても今のウィンドウ / the current window even with multiple windows */
            var viewBounds = activeView.bounds;
            var zoom = activeView.zoom;
            var viewWidth = (viewBounds[2] - viewBounds[0]) * zoom;
            /* 表示域の外にはみ出した部分は数えない / Ignore the part outside the view */
            var left = Math.max(0, (spanLeft - viewBounds[0]) * zoom);
            var right = Math.min(viewWidth, (spanRight - viewBounds[0]) * zoom);
            if (right <= left) return null;
            return { left: left, right: right, viewWidth: viewWidth };
        } catch (e) {
            /* テキスト編集中など測れないときは避けない / Do not avoid when it cannot be measured, e.g. while editing text */
            return null;
        }
    }

    /**
     * ダイアログが選択範囲に重なるなら、重ならない左端の位置を返す。
     * 表示域は画面の横中央にあるとみなし、ずれは DIALOG_AVOID_MARGIN で吸収する。
     * @param {number} dialogLeft - 今のダイアログの左端
     * @param {number} dialogWidth - ダイアログの幅
     * @param {number} screenWidth - 画面の幅
     * @param {{left: number, right: number, viewWidth: number}} selectionSpan - getSelectionViewSpan() の結果
     * @returns {number|null} ずらした左端。重ならない・どちらにも収まらないときは null
     */
    function findDialogLeftAvoidingSelection(dialogLeft, dialogWidth, screenWidth, selectionSpan) {
        var viewLeft = (screenWidth - selectionSpan.viewWidth) / 2;
        var avoidLeft = viewLeft + selectionSpan.left - DIALOG_AVOID_MARGIN;
        var avoidRight = viewLeft + selectionSpan.right + DIALOG_AVOID_MARGIN;
        if (dialogLeft + dialogWidth <= avoidLeft || dialogLeft >= avoidRight) return null;

        var leftSideLeft = avoidLeft - dialogWidth;   /* 選択範囲の左に置くとき / placed left of the selection */
        var rightSideLeft = avoidRight;               /* 選択範囲の右に置くとき / placed right of the selection */
        var fitsLeft = leftSideLeft >= 0;
        var fitsRight = rightSideLeft + dialogWidth <= screenWidth;
        /* 選択範囲が画面の右寄りなら左へ、左寄りなら右へ逃がす / Move away from the side the selection leans to */
        var preferLeft = (avoidLeft + avoidRight) / 2 > screenWidth / 2;
        if (preferLeft && fitsLeft) return leftSideLeft;
        if (fitsRight) return rightSideLeft;
        if (fitsLeft) return leftSideLeft;
        return null;
    }

    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_BOTTOM_MARGIN = 14; /* ボタン行の下の余白。ダイアログの下余白と合わせて約30px（Illustrator 標準のダイアログに合わせる） / bottom margin; with the dialog margin about 30px, like Illustrator's own dialogs */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
    var BUTTON_ROW_CENTER_MAX_WIDTH = 240; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

    /**
     * ダイアログ下部のボタン行を作る。
     * 通常は「左のグループ・伸びるスペーサー・右のグループ」、centered なら行そのものを左右中央に置く
     * @param {Window|Group|Panel} parent - 行を足す先（ふつうはダイアログ）
     * @param {Object} [rowOptions] - { centered: true } で左右中央に並べる
     * @returns {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} 行と左右のグループ（centered のときは左右が null）
     */
    function addButtonRow(parent, rowOptions) {
        var isCentered = !!(rowOptions && rowOptions.centered);
        var btnRowGroup = parent.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, BUTTON_ROW_BOTTOM_MARGIN];
        btnRowGroup.spacing = BUTTON_ROW_SPACING;

        if (isCentered) {
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            return { rowGroup: btnRowGroup, leftGroup: null, rightGroup: null };
        }

        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        btnLeftGroup.spacing = BUTTON_ROW_SPACING;

        /* 余りの幅を吸って、右のグループを右端に寄せる / Absorbs the extra width so the right group sits at the right edge */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.spacing = BUTTON_ROW_SPACING;

        return { rowGroup: btnRowGroup, leftGroup: btnLeftGroup, rightGroup: btnRightGroup };
    }

    /**
     * 左のグループにボタンが無い（右のボタンだけの）行を、ダイアログの幅に合わせて揃える。
     * 内側の幅（左右の余白を除く）が BUTTON_ROW_CENTER_MAX_WIDTH 以下なら左右中央、それより広ければ右揃えのまま。
     * 幅はレイアウトが決まるまで分からないので、ダイアログを表示した時点（show イベント）で判定する。
     * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
     * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
     * @returns {void}
     */
    function alignRightOnlyButtonRow(buttonRow) {
        if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
        var dialogWindow = buttonRow.rowGroup.window;
        dialogWindow.addEventListener("show", function () {
            if (!buttonRow.leftGroup) return;
            var btnRowGroup = buttonRow.rowGroup;
            /* 行の幅＝ダイアログの内側の幅（左右の余白を除く）/ The row spans the dialog's inner width (margins excluded) */
            if (!btnRowGroup.size || btnRowGroup.size.width > BUTTON_ROW_CENTER_MAX_WIDTH) return;
            /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
            btnRowGroup.remove(buttonRow.leftGroup);
            btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            buttonRow.leftGroup = null;
            dialogWindow.layout.layout(true);
        });
    }

    // ボタン行（再利用パーツ）ここまで / End of the reusable button row

    // UI の明暗（再利用パーツ） / UI theme (reusable)

    /**
     * UI がダークテーマかどうかを判定する（Illustrator は uiBrightness、InDesign は uiBrightnessPreference）
     * @returns {boolean} ダークなら true。取得できない環境では false（明るいUI扱い）
     */
    function isDarkUI() {
        try {
            if (app.preferences && app.preferences.getRealPreference) {
                return app.preferences.getRealPreference("uiBrightness") <= 0.5; /* Illustrator */
            }
            return app.generalPreferences.uiBrightnessPreference <= 0.5; /* InDesign */
        } catch (e) {
            return false;
        }
    }

    // UI の明暗（再利用パーツ）ここまで / End of the reusable UI theme

    // バージョン表示（再利用パーツ） / Version footer (reusable)

    var VERSION_FOOTER_TOP_MARGIN = 4;    /* バージョンの行の上に足す余白（px。ウィンドウの間隔に足す）/ extra space above the version row, added to the window spacing */
    var VERSION_FOOTER_BOTTOM_MARGIN = 8; /* バージョンの行の下の余白（px。ウィンドウの下余白をこの値にする）/ space below the version row (replaces the window's bottom margin) */
    var VERSION_FOOTER_PIXEL_SIZE = 1;    /* ドット1つの大きさ（px）/ size of one dot (px) */
    var VERSION_FOOTER_LETTER_GAP = 1;    /* 字の間（ドット）/ gap between glyphs (dots) */
    var VERSION_FOOTER_BRAND = "DTP Transit"; /* 右下に出す名前（字形にある文字だけ）/ name shown at the bottom right (glyph characters only) */
    var VERSION_FOOTER_COLOR = isDarkUI() ? [0.45, 0.45, 0.45, 1] : [0.62, 0.62, 0.62, 1]; /* 文字の色（ディム表示に近い灰色）/ text color, a gray close to the dimmed look */

    /* 字形：Sigi 7px Regular（Rasmus Andersson、CC0。https://github.com/rsms/sigi-pixel-font）の大文字・数字・「. - _ +」・空白を
       9 行のドット（# が点。上7行が大文字の高さ、下2行は小文字の下に出る部分）に写したもの。Sigi に小文字が無いので、a〜z は同じ大きさで描き足した。I は l と見分けるよう上下に横棒を足した
       / Glyphs: uppercase, digits, ". - _ +" and space from Sigi 7px Regular by Rasmus Andersson (CC0, https://github.com/rsms/sigi-pixel-font),
         nine rows each ("#" is a dot; seven rows of cap height plus two for descenders). Sigi has no lowercase, so a-z were drawn to match. I has top and bottom bars to tell it from l */
    var VERSION_FOOTER_GLYPH_HEIGHT = 9;  /* 字の高さ（ドット。下2行は g・j・p・q・y の下に出る部分）/ glyph height (dots), including two descender rows */
    var VERSION_FOOTER_GLYPHS = {
        "A": [".###.", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
        "B": ["####.", "#...#", "#...#", "####.", "#...#", "#...#", "####.", ".....", "....."],
        "C": [".###.", "#...#", "#....", "#....", "#....", "#...#", ".###.", ".....", "....."],
        "D": ["####.", "#...#", "#...#", "#...#", "#...#", "#...#", "####.", ".....", "....."],
        "E": ["#####", "#....", "#....", "####.", "#....", "#....", "#####", ".....", "....."],
        "F": ["#####", "#....", "#....", "####.", "#....", "#....", "#....", ".....", "....."],
        "G": [".###.", "#...#", "#....", "#..##", "#...#", "#...#", ".###.", ".....", "....."],
        "H": ["#...#", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
        "I": ["###", ".#.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
        "J": ["....#", "....#", "....#", "....#", "....#", "#...#", ".###.", ".....", "....."],
        "K": ["#...#", "#..#.", "#.#..", "##...", "#.#..", "#..#.", "#...#", ".....", "....."],
        "L": ["#...", "#...", "#...", "#...", "#...", "#...", "####", "....", "...."],
        "M": ["#...#", "##.##", "#.#.#", "#...#", "#...#", "#...#", "#...#", ".....", "....."],
        "N": ["#...#", "#...#", "##..#", "#.#.#", "#..##", "#...#", "#...#", ".....", "....."],
        "O": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "P": ["####.", "#...#", "#...#", "####.", "#....", "#....", "#....", ".....", "....."],
        "Q": [".###.", "#...#", "#...#", "#...#", "#...#", "#..##", ".####", ".....", "....."],
        "R": ["####.", "#...#", "#...#", "####.", "#.#..", "#..#.", "#...#", ".....", "....."],
        "S": [".###.", "#...#", "#....", ".###.", "....#", "#...#", ".###.", ".....", "....."],
        "T": ["#####", "..#..", "..#..", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
        "U": ["#...#", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "V": ["#...#", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", "..#..", ".....", "....."],
        "W": ["#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", ".##.##.", ".......", "......."],
        "X": ["#...#", "#...#", ".#.#.", "..#..", ".#.#.", "#...#", "#...#", ".....", "....."],
        "Y": ["#...#", "#...#", ".#.#.", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
        "Z": ["#####", "....#", "...#.", "..#..", ".#...", "#....", "#####", ".....", "....."],
        "a": ["....", "....", ".##.", "...#", ".###", "#..#", ".###", "....", "...."],
        "b": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "###.", "....", "...."],
        "c": ["....", "....", ".###", "#...", "#...", "#...", ".###", "....", "...."],
        "d": ["...#", "...#", ".###", "#..#", "#..#", "#..#", ".###", "....", "...."],
        "e": ["....", "....", ".##.", "#..#", "####", "#...", ".###", "....", "...."],
        "f": ["..#", ".#.", "###", ".#.", ".#.", ".#.", ".#.", "...", "..."],
        "g": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "###."],
        "h": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
        "i": ["#", ".", "#", "#", "#", "#", "#", ".", "."],
        "j": [".#", "..", ".#", ".#", ".#", ".#", ".#", ".#", "#."],
        "k": ["#...", "#...", "#..#", "#.#.", "##..", "#.#.", "#..#", "....", "...."],
        "l": ["#.", "#.", "#.", "#.", "#.", "#.", ".#", "..", ".."],
        "m": [".....", ".....", "####.", "#.#.#", "#.#.#", "#.#.#", "#.#.#", ".....", "....."],
        "n": ["....", "....", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
        "o": ["....", "....", ".##.", "#..#", "#..#", "#..#", ".##.", "....", "...."],
        "p": ["....", "....", "###.", "#..#", "#..#", "#..#", "###.", "#...", "#..."],
        "q": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "...#"],
        "r": ["...", "...", "#.#", "##.", "#..", "#..", "#..", "...", "..."],
        "s": ["....", "....", ".###", "#...", ".##.", "...#", "###.", "....", "...."],
        "t": ["...", ".#.", "###", ".#.", ".#.", ".#.", ".##", "...", "..."],
        "u": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "....", "...."],
        "v": [".....", ".....", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", ".....", "....."],
        "w": [".....", ".....", "#...#", "#...#", "#.#.#", "#.#.#", ".#.#.", ".....", "....."],
        "x": ["....", "....", "#..#", "#..#", ".##.", "#..#", "#..#", "....", "...."],
        "y": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "...#", "###."],
        "z": ["....", "....", "####", "...#", ".##.", "#...", "####", "....", "...."],
        "0": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "1": [".#.", "##.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
        "2": [".###.", "#...#", "....#", "...#.", "..#..", ".#...", "#####", ".....", "....."],
        "3": [".###.", "#...#", "....#", "..##.", "....#", "#...#", ".###.", ".....", "....."],
        "4": ["...#.", "..##.", ".#.#.", "#..#.", "#####", "...#.", "...#.", ".....", "....."],
        "5": ["#####", "#....", "####.", "....#", "....#", "#...#", ".###.", ".....", "....."],
        "6": [".###.", "#...#", "#....", "####.", "#...#", "#...#", ".###.", ".....", "....."],
        "7": ["#####", "....#", "....#", "...#.", "...#.", "..#..", "..#..", ".....", "....."],
        "8": [".###.", "#...#", "#...#", ".###.", "#...#", "#...#", ".###.", ".....", "....."],
        "9": [".###.", "#...#", "#...#", ".####", "....#", "#...#", ".###.", ".....", "....."],
        ".": [".", ".", ".", ".", ".", ".", "#", ".", "."],
        "-": ["...", "...", "...", "###", "...", "...", "...", "...", "..."],
        "_": ["....", "....", "....", "....", "....", "....", "####", "....", "...."],
        "+": [".....", "..#..", "..#..", "#####", "..#..", "..#..", ".....", ".....", "....."],
        " ": ["...", "...", "...", "...", "...", "...", "...", "...", "..."]
    };

    /**
     * ウィンドウの下端の左にスクリプト名とバージョン（SCRIPT_NAME・SCRIPTMETA の Version）、右に VERSION_FOOTER_BRAND を小さく薄く入れる。部品をすべて足したあと、show() の直前に呼ぶ。
     * 直前の行（ボタン行など）の下余白を 0 にし、ウィンドウの下余白を VERSION_FOOTER_BOTTOM_MARGIN にする。
     * InDesign の ScriptUI は文字サイズの指定を無視するので、onDraw でドットを並べて描く（UITheme の isDarkUI() を使う）
     * @param {Window} targetWindow - ダイアログまたはパレット（縦並び）
     * @returns {Group|null} バージョンを描くグループ。すでに足してあれば null
     */
    function addVersionFooter(targetWindow) {
        if (targetWindow.versionFooter) return null;

        /* ボタン行の下余白（BUTTON_ROW_BOTTOM_MARGIN）は、この行が引き受ける / this row takes over the button row's bottom margin */
        var lastChild = targetWindow.children[targetWindow.children.length - 1];
        if (lastChild && lastChild.type === "group") {
            var rowMargins = lastChild.margins;
            lastChild.margins = [rowMargins[0], rowMargins[1], rowMargins[2], 0];
        }
        var windowMargins = targetWindow.margins;
        targetWindow.margins = [windowMargins[0], windowMargins[1], windowMargins[2], VERSION_FOOTER_BOTTOM_MARGIN];

        var footerGroup = targetWindow.add("group");
        footerGroup.orientation = "row";
        footerGroup.alignment = ["fill", "bottom"];
        footerGroup.alignChildren = ["left", "center"];
        footerGroup.margins = [0, VERSION_FOOTER_TOP_MARGIN, 0, 0];

        /* 左は「スクリプト名 バージョン」。字形に無い文字（日本語など）を含む名前や、幅が足りないときはバージョンだけ。
           左の欄は余りの幅まで伸ばし、ダイアログの幅はバージョンの分しか求めない（名前でダイアログを広げない）
           / Left: "script name version", or just the version when the name has characters without glyphs (e.g. Japanese)
             or does not fit. The left area takes the extra width and only asks for the version's width, so the name never widens the dialog */
        var versionText = getVersionInfoValue("version");
        var leftCandidates = [versionText];
        if (typeof SCRIPT_NAME === "string" && hasVersionFooterGlyphs(SCRIPT_NAME)) leftCandidates.unshift(SCRIPT_NAME + " " + versionText);
        var versionCanvas = addVersionFooterCanvas(footerGroup, leftCandidates);
        versionCanvas.alignment = ["fill", "center"];
        bindVersionInfoClick(versionCanvas); /* クリックで「このスクリプトについて」を開く / click opens About This Script */
        var brandCanvas = addVersionFooterCanvas(footerGroup, [VERSION_FOOTER_BRAND]);
        brandCanvas.alignment = ["right", "center"];
        brandCanvas.maximumSize = brandCanvas.preferredSize;
        targetWindow.versionFooter = footerGroup;
        return versionCanvas;
    }

    /**
     * 文字列のすべての文字に字形があるかを返す
     * @param {string} dotText - 文字列
     * @returns {boolean} すべてあれば true
     */
    function hasVersionFooterGlyphs(dotText) {
        for (var i = 0; i < dotText.length; i++) {
            if (!VERSION_FOOTER_GLYPHS.hasOwnProperty(dotText.charAt(i))) return false;
        }
        return dotText.length > 0;
    }

    /**
     * ドットで文字列を描くグループを足す。候補のうち、欄の幅に収まる最初のものを描く
     * @param {Group} parentGroup - 足す先
     * @param {string[]} dotTexts - 描く文字列の候補（長い順。いちばん短いものの幅を欄の幅として求める）
     * @returns {Group} 描くグループ
     */
    function addVersionFooterCanvas(parentGroup, dotTexts) {
        var dotCanvas = parentGroup.add("group");
        var shortestText = dotTexts[dotTexts.length - 1];
        dotCanvas.preferredSize = [measureVersionFooterDots(shortestText) * VERSION_FOOTER_PIXEL_SIZE, VERSION_FOOTER_GLYPH_HEIGHT * VERSION_FOOTER_PIXEL_SIZE];
        dotCanvas.helpTip = dotTexts[0];
        dotCanvas.onDraw = function () {
            var availableWidth = this.size ? this.size.width : 0;
            var drawText = shortestText;
            for (var i = 0; i < dotTexts.length; i++) {
                if (measureVersionFooterDots(dotTexts[i]) * VERSION_FOOTER_PIXEL_SIZE <= availableWidth) {
                    drawText = dotTexts[i];
                    break;
                }
            }
            drawVersionFooterDots(this.graphics, drawText);
        };
        return dotCanvas;
    }

    /**
     * 文字列の幅をドットの数で返す（字形に無い文字は数えない）
     * @param {string} versionText - 文字列
     * @returns {number} 幅（ドット）
     */
    function measureVersionFooterDots(versionText) {
        var dotCount = 0;
        for (var i = 0; i < versionText.length; i++) {
            var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
            if (!glyphRows) continue;
            if (dotCount > 0) dotCount += VERSION_FOOTER_LETTER_GAP;
            dotCount += glyphRows[0].length;
        }
        return dotCount;
    }

    /**
     * 文字列をドットで描く。横に続くドットは1つの四角にまとめて塗る
     * @param {ScriptUIGraphics} canvasGraphics - 描く先の graphics
     * @param {string} versionText - 文字列
     * @returns {void}
     */
    function drawVersionFooterDots(canvasGraphics, versionText) {
        var footerBrush = canvasGraphics.newBrush(canvasGraphics.BrushType.SOLID_COLOR, VERSION_FOOTER_COLOR);
        var dotSize = VERSION_FOOTER_PIXEL_SIZE;
        var cursorDots = 0;
        canvasGraphics.newPath();
        for (var i = 0; i < versionText.length; i++) {
            var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
            if (!glyphRows) continue;
            if (cursorDots > 0) cursorDots += VERSION_FOOTER_LETTER_GAP;
            for (var row = 0; row < glyphRows.length; row++) {
                var rowText = glyphRows[row];
                var runStart = -1;
                for (var col = 0; col <= rowText.length; col++) {
                    var isDot = col < rowText.length && rowText.charAt(col) === "#";
                    if (isDot && runStart < 0) runStart = col;
                    if (!isDot && runStart >= 0) {
                        canvasGraphics.rectPath((cursorDots + runStart) * dotSize, row * dotSize, (col - runStart) * dotSize, dotSize);
                        runStart = -1;
                    }
                }
            }
            cursorDots += glyphRows[0].length;
        }
        canvasGraphics.fillPath(footerBrush);
    }

    // -----------------------------------------
    // このスクリプトについて / About this script
    // -----------------------------------------
    /* 左下のスクリプト名・バージョンをクリックすると開く。SCRIPTMETA の項目を getScriptMetaValue() で直接読み（スクリプト情報の部品）、
       この部品だけで動くよう、文言（VERSION_INFO_LABELS）とレイアウトの値も部品の中に持つ
       Opened by clicking the name and version at the bottom left. Reads the SCRIPTMETA fields directly with getScriptMetaValue(),
       and keeps its own labels and layout values so the part works on its own */

    var VERSION_INFO_LANG = (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en"; /* 文言の言語 / label language */
    var VERSION_INFO_IS_INDESIGN = (String(app.name).indexOf("InDesign") >= 0);         /* InDesign で動いているか / running in InDesign */
    /* 実行中のスクリプトのファイル。ボタンを押すころには $.fileName が別のものを指すことがあるので、読み込んだ時点で控える
       The running script file, captured at load time; $.fileName may point elsewhere by the time the button is clicked */
    var VERSION_INFO_SCRIPT_FILE = new File($.fileName);
    var VERSION_INFO_REVEAL_APP_PATH = "/Applications/RevealInFinder.app";   /* Finder で選択表示する補助アプリ / helper app that reveals a file in the Finder */
    var VERSION_INFO_REVEAL_PATH_FILE = "/tmp/illustrator_reveal_path.txt";  /* 補助アプリが読むパス（アプリの旧名のまま。変えると読めない）/ path the helper reads (keeps the app's old name) */
    var VERSION_INFO_FILE_VIEWER_APP_PATH = "/Applications/OpenInFileViewer.app";   /* Path Finder が起動中なら Path Finder、そうでなければ Finder で選択表示する補助アプリ / helper app that reveals a file in Path Finder when it is running, otherwise in the Finder */
    var VERSION_INFO_FILE_VIEWER_PATH_FILE = "/tmp/open_in_file_viewer_path.txt"; /* OpenInFileViewer.app が読むパス / path the helper reads */

    var VERSION_INFO_WINDOW_MARGINS = 16;               /* ウィンドウの余白 / window margin */
    var VERSION_INFO_WINDOW_SPACING = 12;               /* ウィンドウ内の間隔 / window spacing */
    var VERSION_INFO_PANEL_MARGINS = [16, 20, 16, 12];  /* パネルの余白 [左,上,右,下] / panel margins */
    var VERSION_INFO_ROW_SPACING = 6;                   /* 項目名と値の間隔 / gap between a label and its value */
    var VERSION_INFO_BUTTON_SPACING = 10;               /* ボタンどうしの間隔 / gap between buttons */
    var VERSION_INFO_BUTTON_ROW_MARGINS = [0, 5, 0, 14]; /* ボタン行の余白（Illustrator 標準のダイアログに合わせる）/ button row margins */
    var VERSION_INFO_COPY_BUTTON_SIZE = 22;             /* コピーボタンの一辺 / copy button size */
    /* コピーボタンの配色。∧∨（ステップボタン）と同じく、背景に重ねる半透明 / Copy button colors, translucent overlays like the stepper */
    var VERSION_INFO_COPY_FILL_COLOR    = isDarkUI() ? [0, 0, 0, 0.10] : [1, 1, 1, 0.50]; /* 地 / background */
    var VERSION_INFO_COPY_FRAME_COLOR   = isDarkUI() ? [1, 1, 1, 0.12] : [0, 0, 0, 0.15]; /* 枠線 / frame */
    var VERSION_INFO_COPY_PRESSED_COLOR = isDarkUI() ? [1, 1, 1, 0.12] : [0, 0, 0, 0.13]; /* 押下中 / pressed */
    var VERSION_INFO_COPY_ICON_COLOR    = isDarkUI() ? [1, 1, 1, 0.85] : [0, 0, 0, 0.70]; /* アイコン / icon */

    var VERSION_INFO_LABELS = {
        dialog: {
            title: { ja: "このスクリプトについて", en: "About This Script" }
        },
        panel: {
            generalInfo: { ja: "基本情報", en: "General" },
            links: { ja: "リンク", en: "Links" }
        },
        fieldLabel: {
            scriptName: { ja: "スクリプト名", en: "Script Name" },
            scriptFeature: { ja: "機能", en: "Feature" },
            scriptVersion: { ja: "バージョン", en: "Version" }
        },
        button: {
            openReadme: { ja: "README", en: "README" },
            openSourceCode: { ja: "ソースコード", en: "Source Code" },
            openArticle: { ja: "記事（note）", en: "Article (note)" },
            revealInFinder: { ja: "Finder に表示", en: "Show in Finder" },
            revealInExplorer: { ja: "エクスプローラーに表示", en: "Show in Explorer" },
            close: { ja: "閉じる", en: "Close" }
        },
        tooltip: {
            versionInfo: { ja: "クリックで、このスクリプトについての情報を表示", en: "Click for info about this script" },
            copyScriptName: { ja: "スクリプト名をコピー", en: "Copy Script Name" }
        },
        alert: {
            copyFailed: { ja: "スクリプト名をコピーできませんでした。", en: "Could not copy the script name." }
        }
    };

    /**
     * VERSION_INFO_LABELS の文言を今の言語で返す
     * @param {Object} labelEntry - { ja, en }
     * @returns {string} 文言
     */
    function getVersionInfoLabel(labelEntry) {
        return labelEntry[VERSION_INFO_LANG] || labelEntry.en;
    }

    /**
     * 「このスクリプトについて」とフッターに出す値を返す。SCRIPTMETA の項目を getScriptMetaValue() で直接読み、名前は SCRIPT_NAME。
     * getScriptMetaValue() が無いスクリプトは従来の SCRIPT_* を typeof で読む（未定義でも例外にならない）
     * @param {string} variableKey - "name" / "version" / "feature" / "readme" / "source" / "article"
     * @returns {string} 値（バージョンは「v1.2.0」の形）。無ければ空文字
     */
    function getVersionInfoValue(variableKey) {
        var infoValue = "";
        if (variableKey === "name") {
            infoValue = (typeof SCRIPT_NAME !== "undefined") ? SCRIPT_NAME : "";
        } else if (typeof getScriptMetaValue === "function") {
            var isJapanese = (VERSION_INFO_LANG === "ja");
            if (variableKey === "version") infoValue = getScriptMetaValue("Version") ? "v" + getScriptMetaValue("Version") : "";
            else if (variableKey === "feature") infoValue = getScriptMetaValue("Name");
            else if (variableKey === "readme") infoValue = getScriptMetaValue(isJapanese ? "X-com.swwwitch-README-JA" : "X-com.swwwitch-README-EN");
            else if (variableKey === "source") infoValue = getScriptMetaValue("X-com.swwwitch-Source-URL");
            else if (variableKey === "article") infoValue = getScriptMetaValue("X-com.swwwitch-Article-URL");
        } else {
            if (variableKey === "version") infoValue = (typeof SCRIPT_VERSION !== "undefined") ? SCRIPT_VERSION : "";
            else if (variableKey === "feature") infoValue = (typeof SCRIPT_DESCRIPTION_JA !== "undefined") ? SCRIPT_DESCRIPTION_JA : "";
            else if (variableKey === "readme") {
                if (VERSION_INFO_LANG === "ja") infoValue = (typeof SCRIPT_README_JA !== "undefined") ? SCRIPT_README_JA : "";
                else infoValue = (typeof SCRIPT_README_EN !== "undefined") ? SCRIPT_README_EN : "";
            }
            else if (variableKey === "source") infoValue = (typeof SCRIPT_SOURCE_URL !== "undefined") ? SCRIPT_SOURCE_URL : "";
            else if (variableKey === "article") infoValue = (typeof SCRIPT_ARTICLE_URL !== "undefined") ? SCRIPT_ARTICLE_URL : "";
        }
        return (infoValue === undefined || infoValue === null) ? "" : String(infoValue);
    }

    /**
     * 左下のスクリプト名・バージョンをクリックしたら、「このスクリプトについて」のダイアログを開くようにする
     * @param {Group} versionCanvas - addVersionFooterCanvas() で作った左の欄
     * @returns {void}
     */
    function bindVersionInfoClick(versionCanvas) {
        versionCanvas.helpTip = getVersionInfoLabel(VERSION_INFO_LABELS.tooltip.versionInfo);
        /* mousedown で開くと、離したときのイベントが閉じたあとの欄に届くので mouseup で開く
           Open on mouseup so the release does not land on the footer after the dialog closes */
        versionCanvas.addEventListener("mouseup", function () {
            showVersionInfoDialog(versionCanvas.window.type === "palette");
        });
    }

    /**
     * 「このスクリプトについて」のダイアログを開く。スクリプト名（コピーボタン付き）・機能・バージョンの［基本情報］パネル、
     * README・ソースコード・記事を開く［リンク］パネル、［Finder に表示］［閉じる］を並べる
     * @param {boolean} isFromPalette - 常駐パレットから開いたなら true（Illustrator のコピーを BridgeTalk 経由にする）
     * @returns {void}
     */
    function showVersionInfoDialog(isFromPalette) {
        var infoDialog = new Window("dialog", getVersionInfoLabel(VERSION_INFO_LABELS.dialog.title));
        infoDialog.orientation = "column";
        infoDialog.alignChildren = "fill";
        infoDialog.margins = VERSION_INFO_WINDOW_MARGINS;
        infoDialog.spacing = VERSION_INFO_WINDOW_SPACING;

        var infoPanel = addVersionInfoPanel(infoDialog, VERSION_INFO_LABELS.panel.generalInfo);
        /* コピーボタンはスクリプト名だけ / Only the script name gets a copy button */
        addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptName, getVersionInfoValue("name"), isFromPalette);
        /* 機能は README の見出しで日本語だけなので、日本語の画面だけに出す / The feature is the Japanese README title, shown only in the Japanese UI */
        if (VERSION_INFO_LANG === "ja" && getVersionInfoValue("feature") !== "") {
            addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptFeature, getVersionInfoValue("feature"), null);
        }
        addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptVersion, getVersionInfoValue("version"), null);
        alignVersionInfoLabels(infoPanel);

        /* URL が無いボタンはディム表示にして、並びはスクリプトごとに変えない / Dim buttons without a URL; keep the same buttons everywhere */
        var linksPanel = addVersionInfoPanel(infoDialog, VERSION_INFO_LABELS.panel.links);
        var linkButtonRow = linksPanel.add("group");
        linkButtonRow.orientation = "row";
        linkButtonRow.alignment = ["left", "center"];
        linkButtonRow.alignChildren = ["left", "center"];
        linkButtonRow.spacing = VERSION_INFO_BUTTON_SPACING;
        addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openReadme, getVersionInfoValue("readme"));
        addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openSourceCode, getVersionInfoValue("source"));
        addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openArticle, getVersionInfoValue("article"));

        /* ボタン行：左に［Finder に表示］、右に［閉じる］ / Button row: Show in Finder at the left, Close at the right */
        var buttonRow = infoDialog.add("group");
        buttonRow.orientation = "row";
        buttonRow.alignment = ["fill", "bottom"];
        buttonRow.alignChildren = ["left", "center"];
        buttonRow.margins = VERSION_INFO_BUTTON_ROW_MARGINS;
        buttonRow.spacing = VERSION_INFO_BUTTON_SPACING;
        var revealLabel = (File.fs === "Windows") ? VERSION_INFO_LABELS.button.revealInExplorer : VERSION_INFO_LABELS.button.revealInFinder;
        var revealButton = buttonRow.add("button", undefined, getVersionInfoLabel(revealLabel));
        if (VERSION_INFO_SCRIPT_FILE.exists) {
            revealButton.helpTip = VERSION_INFO_SCRIPT_FILE.fsName;
            revealButton.onClick = function () {
                revealVersionInfoScriptFile(VERSION_INFO_SCRIPT_FILE);
                infoDialog.close();
            };
        } else {
            revealButton.enabled = false;
        }
        /* 余りの幅を吸って、［閉じる］を右端に寄せる / Absorbs the extra width so Close sits at the right edge */
        var buttonSpacer = buttonRow.add("group");
        buttonSpacer.alignment = ["fill", "fill"];
        buttonSpacer.minimumSize.width = 0;
        var closeButton = buttonRow.add("button", undefined, getVersionInfoLabel(VERSION_INFO_LABELS.button.close), { name: "ok" });
        closeButton.alignment = ["right", "center"];
        /* return でも esc でも閉じる / Close with both Return and Esc */
        infoDialog.defaultElement = closeButton;
        infoDialog.cancelElement = closeButton;
        infoDialog.onShow = function () {
            pinVersionInfoCopyButton(infoPanel);
        };
        infoDialog.show();
    }

    /**
     * 「このスクリプトについて」のダイアログにパネルを足す
     * @param {Window} infoDialog - 足す先のダイアログ
     * @param {Object} titleEntry - パネル名の文言 { ja, en }
     * @returns {Panel} パネル
     */
    function addVersionInfoPanel(infoDialog, titleEntry) {
        var infoPanel = infoDialog.add("panel", undefined, getVersionInfoLabel(titleEntry));
        infoPanel.orientation = "column";
        infoPanel.alignChildren = ["fill", "top"];
        infoPanel.alignment = "fill";
        infoPanel.margins = VERSION_INFO_PANEL_MARGINS;
        infoPanel.spacing = VERSION_INFO_WINDOW_SPACING;
        return infoPanel;
    }

    /**
     * ［基本情報］パネルに「項目名 : 値」の1行を足す（値は見せるだけなので statictext）
     * @param {Panel} infoPanel - 足す先のパネル
     * @param {Object} labelEntry - 項目名の文言 { ja, en }
     * @param {string} valueText - 表示する値
     * @param {boolean|null} copyFromPalette - 値の右にコピーボタンを置くなら true / false（常駐パレットから開いたか）、置かないなら null
     * @returns {StaticText} 値の statictext
     */
    function addVersionInfoRow(infoPanel, labelEntry, valueText, copyFromPalette) {
        var infoRow = infoPanel.add("group");
        infoRow.orientation = "row";
        infoRow.alignment = ["fill", "center"];
        infoRow.alignChildren = ["left", "center"];
        infoRow.spacing = VERSION_INFO_ROW_SPACING;
        /* 項目名のコロンは日本語が「 :」、英語が「:」 / Colon: " :" in Japanese, ":" in English */
        var rowLabel = infoRow.add("statictext", undefined, getVersionInfoLabel(labelEntry) + (VERSION_INFO_LANG === "ja" ? " :" : ":"));
        rowLabel.justify = "right";
        var valueLabel = infoRow.add("statictext", undefined, valueText);
        if (copyFromPalette === null) return valueLabel;
        /* statictext の fill では伸びないことがあるので、伸びるスペーサーでコピーボタンをパネルの右端に寄せる
           A stretching spacer (statictext fill is unreliable) pushes the copy button to the panel's right edge */
        var copySpacer = infoRow.add("group");
        copySpacer.alignment = ["fill", "fill"];
        copySpacer.minimumSize.width = 0;
        infoRow.copyButton = addVersionInfoCopyButton(infoRow, valueLabel, copyFromPalette);
        return valueLabel;
    }

    /**
     * コピーボタンを［基本情報］パネルの右端に置き直す（表示時、レイアウトが決まったあとに呼ぶ）。
     * 行の fill・スペーサーだけでは、項目の値が短い行がパネルの幅まで伸びないことがあるので、行の幅とボタンの位置を直接入れる
     * @param {Panel} infoPanel - ［基本情報］パネル
     * @returns {void}
     */
    function pinVersionInfoCopyButton(infoPanel) {
        var innerWidth = infoPanel.size.width - VERSION_INFO_PANEL_MARGINS[0] - VERSION_INFO_PANEL_MARGINS[2];
        for (var i = 0; i < infoPanel.children.length; i++) {
            var infoRow = infoPanel.children[i];
            if (!infoRow.copyButton || !infoRow.size || infoRow.size.width >= innerWidth) continue;
            infoRow.size = [innerWidth, infoRow.size.height];
            infoRow.copyButton.location = [innerWidth - infoRow.copyButton.size.width, infoRow.copyButton.location[1]];
        }
    }

    /**
     * ［基本情報］パネルの項目名を、いちばん長いものの幅にそろえる
     * @param {Panel} infoPanel - 対象のパネル
     * @returns {void}
     */
    function alignVersionInfoLabels(infoPanel) {
        var rowLabels = [];
        var labelWidth = 0;
        for (var i = 0; i < infoPanel.children.length; i++) {
            var rowLabel = infoPanel.children[i].children[0]; /* 行の先頭は項目名 / each row starts with its label */
            rowLabels.push(rowLabel);
            labelWidth = Math.max(labelWidth, rowLabel.graphics.measureString(rowLabel.text)[0]);
        }
        for (var j = 0; j < rowLabels.length; j++) {
            rowLabels[j].preferredSize.width = Math.ceil(labelWidth) + 2;
        }
    }

    /**
     * URL を開くボタンを足す。URL が無ければディム表示にする。開いたらダイアログは閉じる
     * @param {Window} infoDialog - ボタンを置くダイアログ
     * @param {Group} buttonGroup - 足す先のグループ
     * @param {Object} labelEntry - ボタンの文言 { ja, en }
     * @param {string} targetUrl - 開く URL（空ならディム表示）
     * @returns {void}
     */
    function addVersionInfoLinkButton(infoDialog, buttonGroup, labelEntry, targetUrl) {
        var linkButton = buttonGroup.add("button", undefined, getVersionInfoLabel(labelEntry));
        if (!targetUrl) {
            linkButton.enabled = false;
            return;
        }
        linkButton.helpTip = targetUrl;
        linkButton.onClick = function () {
            openVersionInfoUrl(targetUrl);
            infoDialog.close();
        };
    }

    /**
     * 値をクリップボードへコピーするアイコンのボタンを足す。コピーしたらアイコンをチェックマークに変え、ポインタが外れたら戻す
     * @param {Group} infoRow - 足す先の行
     * @param {StaticText} valueLabel - コピーする値の statictext
     * @param {boolean} isFromPalette - 常駐パレットから開いたなら true
     * @returns {Group} ボタンとして使う group
     */
    function addVersionInfoCopyButton(infoRow, valueLabel, isFromPalette) {
        var buttonSize = VERSION_INFO_COPY_BUTTON_SIZE;
        var copyButton = infoRow.add("group");
        copyButton.preferredSize = [buttonSize, buttonSize];
        copyButton.minimumSize = [buttonSize, buttonSize];
        copyButton.maximumSize = [buttonSize, buttonSize];
        copyButton.helpTip = getVersionInfoLabel(VERSION_INFO_LABELS.tooltip.copyScriptName);
        copyButton.isPressed = false;
        copyButton.isCopied = false;

        copyButton.onDraw = function () {
            var buttonGraphics = copyButton.graphics;
            var fillColor = copyButton.isPressed ? VERSION_INFO_COPY_PRESSED_COLOR : VERSION_INFO_COPY_FILL_COLOR;
            buttonGraphics.newPath();
            buttonGraphics.rectPath(0, 0, buttonSize, buttonSize);
            buttonGraphics.fillPath(buttonGraphics.newBrush(buttonGraphics.BrushType.SOLID_COLOR, fillColor));
            buttonGraphics.newPath();
            buttonGraphics.rectPath(0.5, 0.5, buttonSize - 1, buttonSize - 1);
            buttonGraphics.strokePath(buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_FRAME_COLOR, 1));
            if (copyButton.isCopied) drawVersionInfoCheckIcon(buttonGraphics, buttonSize);
            else drawVersionInfoCopyIcon(buttonGraphics, buttonSize);
        };

        /**
         * 状態を変えて描き直す（group には notify() が無いので隠して再表示する）
         * @param {boolean} isPressed - 押下中なら true
         * @param {boolean} isCopied - コピー済みなら true
         * @returns {void}
         */
        function repaint(isPressed, isCopied) {
            if (copyButton.isPressed === isPressed && copyButton.isCopied === isCopied) return;
            copyButton.isPressed = isPressed;
            copyButton.isCopied = isCopied;
            copyButton.hide();
            copyButton.show();
        }
        copyButton.addEventListener("mousedown", function () { repaint(true, false); });
        copyButton.addEventListener("mouseup", function () {
            if (!copyButton.isPressed) return;
            var hasCopied = copyVersionInfoText(valueLabel.text, isFromPalette);
            repaint(false, hasCopied);
            if (!hasCopied) alert(getVersionInfoLabel(VERSION_INFO_LABELS.alert.copyFailed));
        });
        /* 押したまま外へ出たとき・コピーのあと離れたときは元の見た目に戻す / Reset when the pointer leaves */
        copyButton.addEventListener("mouseout", function () { repaint(false, false); });
        return copyButton;
    }

    /**
     * コピーのアイコン（後ろの四角に前の四角を重ねた形）を描く
     * @param {ScriptUIGraphics} buttonGraphics - 描画先
     * @param {number} buttonSize - ボタンの一辺
     * @returns {void}
     */
    function drawVersionInfoCopyIcon(buttonGraphics, buttonSize) {
        var iconPen = buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_ICON_COLOR, 1);
        var sheetSize = 9;
        var offset = 3;
        var frontLeft = Math.round((buttonSize - sheetSize - offset) / 2) + offset + 0.5;
        var frontTop = Math.round((buttonSize - sheetSize - offset) / 2) + offset + 0.5;
        var backLeft = frontLeft - offset;
        var backTop = frontTop - offset;
        /* 後ろの四角は、前の四角に隠れる部分を描かない（上と左の辺だけ）/ Draw only the visible top and left of the back sheet */
        buttonGraphics.newPath();
        buttonGraphics.moveTo(frontLeft, backTop + sheetSize);
        buttonGraphics.lineTo(backLeft, backTop + sheetSize);
        buttonGraphics.lineTo(backLeft, backTop);
        buttonGraphics.lineTo(backLeft + sheetSize, backTop);
        buttonGraphics.lineTo(backLeft + sheetSize, frontTop);
        buttonGraphics.strokePath(iconPen);
        buttonGraphics.newPath();
        buttonGraphics.rectPath(frontLeft, frontTop, sheetSize, sheetSize);
        buttonGraphics.strokePath(iconPen);
    }

    /**
     * コピーしたあとのチェックマークを描く
     * @param {ScriptUIGraphics} buttonGraphics - 描画先
     * @param {number} buttonSize - ボタンの一辺
     * @returns {void}
     */
    function drawVersionInfoCheckIcon(buttonGraphics, buttonSize) {
        var center = buttonSize / 2;
        buttonGraphics.newPath();
        buttonGraphics.moveTo(center - 4.5, center);
        buttonGraphics.lineTo(center - 1.5, center + 3);
        buttonGraphics.lineTo(center + 4.5, center - 3.5);
        buttonGraphics.strokePath(buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_ICON_COLOR, 1.5));
    }

    /**
     * 文字列をクリップボードにコピーする。
     * InDesign は AppleScript（Windows は VBScript）の doScript で書き込む。
     * Illustrator は文字列を直接送る API が無いので、一時レイヤーのテキストフレームを選んでコピーする（プレーンテキストも入る）。
     * Illustrator の常駐パレットからは DOM に触れないので、同じ処理を BridgeTalk でメインエンジンに渡す
     * @param {string} copyText - コピーする文字列
     * @param {boolean} isFromPalette - 常駐パレットから開いたなら true
     * @returns {boolean} コピーの操作まで進めたら true
     */
    function copyVersionInfoText(copyText, isFromPalette) {
        try {
            if (VERSION_INFO_IS_INDESIGN) {
                if (File.fs === "Windows") {
                    app.doScript('CreateObject("htmlfile").parentWindow.clipboardData.setData "text", "' + copyText.replace(/"/g, '""') + '"', ScriptLanguage.VISUAL_BASIC);
                } else {
                    app.doScript('set the clipboard to "' + copyText.replace(/\\/g, "\\\\").replace(/"/g, '\\"') + '"', ScriptLanguage.APPLESCRIPT_LANGUAGE);
                }
                return true;
            }
            var copyScript = buildVersionInfoCopyScript(copyText);
            if (!isFromPalette) return eval(copyScript) === true;
            var copyMessage = new BridgeTalk();
            copyMessage.target = BridgeTalk.appSpecifier; /* 複数バージョンを起動していても、今のバージョンに送る / the running version, even with several installed */
            copyMessage.body = copyScript;
            copyMessage.send();
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * Illustrator でコピーする処理を文字列で組み立てる（BridgeTalk でも送れるよう、関数の toString() は使わない）。
     * 一時レイヤーにテキストフレームを作り、フレームを選んでメニューコマンドでコピーして、レイヤーごと消す。
     * 文字として選択（textRange.select()）してコピーすると「オブジェクトをコピーできません」の警告が出るので使わない。
     * 選択と作業レイヤーは元に戻す。ドキュメントが無ければ一時ドキュメントを作って閉じる
     * @param {string} copyText - コピーする文字列
     * @returns {string} 実行すると true / false を返すスクリプト
     */
    function buildVersionInfoCopyScript(copyText) {
        return [
            "(function (copyText) {",
            "    var usingTempDoc = (app.documents.length === 0);",
            "    var targetDoc = usingTempDoc ? app.documents.add() : app.activeDocument;",
            "    var savedSelection = (!usingTempDoc && targetDoc.selection instanceof Array) ? targetDoc.selection : [];",
            "    var savedActiveLayer = usingTempDoc ? null : targetDoc.activeLayer;",
            "    var tempLayer = null;",
            "    var hasCopied = false;",
            "    try {",
            "        tempLayer = targetDoc.layers.add();",
            "        var tempFrame = tempLayer.textFrames.add();",
            "        tempFrame.contents = copyText;",
            "        app.redraw();",
            "        app.executeMenuCommand('deselectall');",
            "        tempFrame.selected = true;",
            "        app.redraw();",
            "        app.executeMenuCommand('copy');",
            "        app.redraw();",
            "        hasCopied = true;",
            "    } catch (e) { }",
            "    if (tempLayer) { try { tempLayer.remove(); } catch (e) { } }",
            "    if (usingTempDoc) { targetDoc.close(SaveOptions.DONOTSAVECHANGES); return hasCopied; }",
            "    try { targetDoc.activeLayer = savedActiveLayer; } catch (e) { }",
            "    targetDoc.selection = null;",
            "    for (var i = 0; i < savedSelection.length; i++) { try { savedSelection[i].selected = true; } catch (e) { } }",
            "    app.redraw();",
            "    return hasCopied;",
            "})(" + quoteVersionInfoString(copyText) + ");"
        ].join(" "); /* 行コメントを含まないので1行につなぐ（BridgeTalk の転送で改行が化けない）/ joined on one line; there are no line comments */
    }

    /**
     * 文字列を、スクリプトに埋め込める二重引用符のリテラルにする
     * @param {string} sourceText - 文字列
     * @returns {string} "…" の形のリテラル
     */
    function quoteVersionInfoString(sourceText) {
        return '"' + String(sourceText).replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/\r/g, "\\r").replace(/\n/g, "\\n") + '"';
    }

    /**
     * ファイルを選択した状態で表示する。OpenInFileViewer.app があれば Path Finder が起動中なら Path Finder、そうでなければ Finder で表示し、
     * 無ければ RevealInFinder.app で Finder に表示する。どちらも無いときは入っているフォルダーを開く
     * @param {File} scriptFile - 表示するファイル
     * @returns {void}
     */
    function revealVersionInfoScriptFile(scriptFile) {
        if (launchVersionInfoHelperApp(VERSION_INFO_FILE_VIEWER_APP_PATH, VERSION_INFO_FILE_VIEWER_PATH_FILE, scriptFile.fsName)) return;
        if (launchVersionInfoHelperApp(VERSION_INFO_REVEAL_APP_PATH, VERSION_INFO_REVEAL_PATH_FILE, scriptFile.fsName)) return;
        /* アプリが無い・起動できないときは、何も起きないままにせずフォルダーを開く / Otherwise open the enclosing folder */
        scriptFile.parent.execute();
    }

    /**
     * 補助アプリが読む一時ファイルにパスを書き出してから、アプリを起動する
     * @param {string} appPath - 補助アプリ（.app）のパス
     * @param {string} pathFilePath - アプリが読む一時ファイルのパス
     * @param {string} targetPath - 表示するファイルの絶対パス
     * @returns {boolean} 起動できたら true。Mac 以外・アプリが無い・起動できないときは false
     */
    function launchVersionInfoHelperApp(appPath, pathFilePath, targetPath) {
        /* .app は実体がフォルダーなので File.exists が false になる環境がある。Folder でも確かめる / An .app is a folder, so check with Folder too */
        if (File.fs !== "Macintosh" || !(new Folder(appPath).exists || new File(appPath).exists)) return false;
        return writeVersionInfoRevealPath(pathFilePath, targetPath) && new File(appPath).execute();
    }

    /**
     * 補助アプリに渡すパスを一時ファイルへ書き出す（UTF-8、改行なし）
     * @param {string} pathFilePath - 一時ファイルのパス
     * @param {string} targetPath - 絶対パス
     * @returns {boolean} 書き出せたら true
     */
    function writeVersionInfoRevealPath(pathFilePath, targetPath) {
        var pathFile = new File(pathFilePath);
        pathFile.encoding = "UTF-8";
        pathFile.lineFeed = "Unix";
        if (!pathFile.open("w")) return false;
        var hasWritten = pathFile.write(targetPath);
        pathFile.close();
        return hasWritten;
    }

    /**
     * URL を既定のブラウザーで開く（一時的なショートカットファイルを作って開く）
     * @param {string} targetUrl - 開く URL
     * @returns {void}
     */
    function openVersionInfoUrl(targetUrl) {
        var isMac = (File.fs === "Macintosh");
        var shortcutFile = new File(Folder.temp.fsName + "/VersionInfoLink" + (isMac ? ".webloc" : ".url"));
        shortcutFile.encoding = "UTF-8";
        shortcutFile.open("w");
        if (isMac) {
            var escapedUrl = targetUrl.replace(/&/g, "&amp;").replace(/</g, "&lt;");
            shortcutFile.write('<?xml version="1.0" encoding="UTF-8"?>\n'
                + '<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">\n'
                + '<plist version="1.0"><dict><key>URL</key><string>' + escapedUrl + '</string></dict></plist>\n');
        } else {
            shortcutFile.write("[InternetShortcut]\r\nURL=" + targetUrl + "\r\n");
        }
        shortcutFile.close();
        shortcutFile.execute();
    }

    // バージョン表示（再利用パーツ）ここまで / End of the reusable version footer

    /**
     * インデントしたオプション用のグループを作る
     * @param {Panel} parentPanel - 追加先のパネル
     * @returns {Group} オプションを並べるグループ
     */
    function addOptionGroup(parentPanel) {
        var optionGroup = parentPanel.add('group');
        optionGroup.orientation = 'column';
        optionGroup.alignChildren = 'left';
        optionGroup.margins = OPTION_INDENT_MARGINS;
        optionGroup.spacing = 6;
        return optionGroup;
    }

    /**
     * オプションダイアログを表示する
     * @param {object} artboardStats - { total: number, emptyIgnoreHidden: number, emptyAll: number }
     * @returns {object|null} 選択内容。キャンセル時は null
     */
    function showOptionsDialog(artboardStats) {
        var removeDialog = new Window('dialog', getLabel('dialog.title'));
        setupWindow(removeDialog);

        var modePanel = removeDialog.add('panel', undefined, getLabel('panel.mode'));
        setupPanel(modePanel, 6);

        /* ラジオは同じ親の中だけ排他になるため、2つともパネル直下に置く
           Radios are only exclusive within one parent, so both sit directly in the panel */
        var keepActiveRadio = modePanel.add('radiobutton', undefined, getLabel('radio.keepActive'));
        keepActiveRadio.helpTip = getLabel('tooltip.keepActive');
        keepActiveRadio.value = true;

        var keepActiveOptionGroup = addOptionGroup(modePanel);
        var deleteOutsideCheckbox = keepActiveOptionGroup.add('checkbox', undefined, getLabel('checkbox.deleteOutsideObjects'));
        deleteOutsideCheckbox.helpTip = getLabel('tooltip.deleteOutsideObjects');
        deleteOutsideCheckbox.value = DEFAULT_DELETE_OUTSIDE_OBJECTS;

        var emptyRadio = modePanel.add('radiobutton', undefined, getLabel('radio.empty'));
        emptyRadio.helpTip = getLabel('tooltip.empty');

        var emptyOptionGroup = addOptionGroup(modePanel);
        var emptyCountText = emptyOptionGroup.add('statictext', undefined, '');
        emptyCountText.preferredSize.width = COUNT_TEXT_WIDTH;
        var ignoreHiddenCheckbox = emptyOptionGroup.add('checkbox', undefined, getLabel('checkbox.ignoreHidden'));
        ignoreHiddenCheckbox.helpTip = getLabel('tooltip.ignoreHidden');
        ignoreHiddenCheckbox.value = DEFAULT_IGNORE_HIDDEN;

        /* アートボードが1枚のときは「空のアートボード」を選べない / Only one artboard leaves nothing to remove */
        var canRemoveEmpty = (artboardStats.total > 1);
        if (!canRemoveEmpty) {
            emptyRadio.enabled = false;
            emptyRadio.helpTip = getLabel('tooltip.singleArtboard');
        }

        /* ボタン行（右寄せの キャンセル／OK）/ Button row (Cancel / OK on the right) */
        var buttonRow = addButtonRow(removeDialog);
        var btnCancel = buttonRow.rightGroup.add('button', undefined, getLabel('button.cancel'), { name: 'cancel' });
        var btnOK = buttonRow.rightGroup.add('button', undefined, getLabel('button.ok'), { name: 'ok' });
        btnCancel.preferredSize.width = BUTTON_WIDTH;
        btnOK.preferredSize.width = BUTTON_WIDTH;

        /**
         * モードに応じてオプションの活性・件数表示・OKの活性を更新する
         * @returns {void}
         */
        function syncDialogState() {
            deleteOutsideCheckbox.enabled = keepActiveRadio.value;

            var emptyModeSelected = emptyRadio.value;
            emptyCountText.enabled = emptyModeSelected;
            ignoreHiddenCheckbox.enabled = emptyModeSelected;

            var emptyCount = ignoreHiddenCheckbox.value ? artboardStats.emptyIgnoreHidden : artboardStats.emptyAll;
            /* 英語はコロンのあとに空白を入れる / add a space after the colon in English */
            emptyCountText.text = labelText('fieldLabel.removalTarget') + (uiLang === 'ja' ? '' : ' ') + emptyCount;

            /* 空アートボードが0枚のときは実行できない / Nothing to do when no artboard is empty */
            btnOK.enabled = emptyModeSelected ? (emptyCount > 0) : true;
        }

        /* クリックだけでなくキーボード操作でも更新 / Refresh on click and on keyboard change */
        keepActiveRadio.onClick = keepActiveRadio.onChanging = syncDialogState;
        emptyRadio.onClick = emptyRadio.onChanging = syncDialogState;
        ignoreHiddenCheckbox.onClick = syncDialogState;
        syncDialogState();

        alignRightOnlyButtonRow(buttonRow);
        addVersionFooter(removeDialog);
        prepareDialogWindow(removeDialog, SCRIPT_NAME);
        if (removeDialog.show() !== 1) return null;

        return {
            mode: keepActiveRadio.value ? 'keepActive' : 'empty',
            deleteOutsideObjects: deleteOutsideCheckbox.value,
            ignoreHidden: ignoreHiddenCheckbox.value
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * エントリポイント：ダイアログ→削除
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel('alert.noDocument'));
            return;
        }

        var doc = app.activeDocument;

        /* 非表示の扱いを切り替えても即座に件数を出せるよう、先に両方を数えておく
           Count both ways up front so the dialog can switch instantly */
        var emptyIndicesIgnoreHidden = findEmptyArtboardIndices(doc, true);
        var emptyIndicesAll = findEmptyArtboardIndices(doc, false);

        var removalOptions = showOptionsDialog({
            total: doc.artboards.length,
            emptyIgnoreHidden: emptyIndicesIgnoreHidden.length,
            emptyAll: emptyIndicesAll.length
        });
        if (!removalOptions) return;

        if (removalOptions.mode === 'keepActive') {
            keepActiveArtboardOnly(doc, removalOptions.deleteOutsideObjects);
        } else {
            removeArtboardsByIndices(doc, removalOptions.ignoreHidden ? emptyIndicesIgnoreHidden : emptyIndicesAll);
        }
    }

    main();

})();
