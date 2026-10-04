#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトに合わせて、アクティブビューをズーム＆センタリングします。
複数選択時は全体の外接矩形にフィットし、選択がないときは100%表示に戻します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ZoomToSelection.md

### Overview

Zooms and centers the active view on the selected objects.
With several objects selected it fits their overall bounding box, and with nothing selected it returns to 100%.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ZoomToSelection.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ZoomToSelection";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.1.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ZoomToSelection.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ZoomToSelection.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/**
 * @author John Wundes (www.wundes.com)
 * @discussion http://www.wundes.com/js4ai/copyright.txt
 */

(function () {

// =========================================
// ユーザー設定 / User Settings
// =========================================

/* フィット時に残す余白の係数（1.0 でぴったり、0.9 で 10% 余白） / Fit margin ratio (1.0 = exact fit, 0.9 = 10% margin) */
var ZOOM_FIT_RATIO = 0.9;

/* 補間ステップ数の上限（大きな移動・ズーム時） / Maximum interpolation steps (for large moves or zooms) */
var MAX_ANIMATION_STEP_COUNT = 32;
/* 各フレームの待機ミリ秒（redraw 自体も時間を食う） / Delay per frame in ms (redraw itself also takes time) */
var FRAME_DELAY_MS = 3;

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
    alert: {
        noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." }
    }
};

// 選択の収集と境界（再利用パーツ） / Selection items and bounds (reusable)

/* 座標を同じと見なす許容値（pt） / Tolerance for treating coordinates as equal, in points */
var SELECTION_ITEMS_TOLERANCE = 0.001;

/**
 * 選択やコレクションを、オブジェクトの配列にそろえる
 * TextRange・PathItem は length を持つので、typename で1個か集まりかを見分ける
 * @param {*} source - doc.selection、配列、DOM のコレクション、または単独のオブジェクト
 * @returns {Array} オブジェクトの配列（空なら []）
 */
function normalizeSelectionItems(source) {
    var items = [];
    if (!source) return items;
    var typeName = "";
    try { typeName = source.typename || ""; } catch (e) { /* 読めない種類 / unreadable kind */ }
    /* 単数形の typename は1個（PageItems などのコレクションは s で終わる）
       A singular typename is one object (collections such as PageItems end in s) */
    if (typeName && !/s$/.test(typeName)) return [source];
    if (typeof source.length !== "number") return items;
    for (var i = 0; i < source.length; i++) items.push(source[i]);
    return items;
}

/**
 * 文字カーソルの選択（TextRange）を、それを含むテキストフレームに読み替える
 * @param {TextRange} textRange - 文字の範囲
 * @returns {TextFrame|null} テキストフレーム（たどれなければ null）
 */
function resolveTextRangeFrame(textRange) {
    var current = textRange;
    /* parent をたどる（深さは念のため制限） / Walk up the parents, with a safety limit */
    for (var depth = 0; depth < 10 && current; depth++) {
        try {
            if (current.typename === "TextFrame") return current;
            current = current.parent;
        } catch (e) {
            break;
        }
    }
    /* ストーリーの先頭フレームで代用する / Fall back to the first frame of the story */
    try {
        var storyFrames = textRange.story.textFrames;
        if (storyFrames.length > 0) return storyFrames[0];
    } catch (e2) { /* ストーリーを持たない / no story */ }
    return null;
}

/**
 * 選択から条件に合うオブジェクトを集める（グループ・レイヤーを再帰でたどり、重複は除く）
 * 条件に合ったオブジェクトの中へは進まない
 * @param {*} source - doc.selection、配列、コレクション、または単独のオブジェクト
 * @param {Object} [options] - 収集の設定
 * @param {function(PageItem): boolean} [options.accept] - 集める条件（既定はグループ・レイヤー以外すべて）
 * @param {boolean} [options.enterGroups] - グループの中をたどる（既定 true）
 * @param {boolean} [options.enterClipGroups] - クリップグループの中をたどる（既定は enterGroups と同じ）
 * @param {boolean} [options.enterCompoundPaths] - 複合パスの中のパスをたどる（既定 false）
 * @param {boolean} [options.textRangeToFrame] - 文字の選択をテキストフレームに読み替える（既定 true）
 * @param {boolean} [options.skipLocked] - ロックされたものを中ごと外す（既定 false）
 * @param {boolean} [options.skipHidden] - 非表示のものを中ごと外す（既定 false）
 * @param {boolean} [options.skipClipMasks] - クリッピングマスクを外す（既定 false）
 * @param {boolean} [options.skipGuides] - ガイドを外す（既定 false）
 * @param {boolean} [options.unique] - 同じ参照を1回だけにする（既定 true。数千件で遅ければ false）
 * @returns {Array} 集めたオブジェクト（前面→背面の順）
 */
function collectSelectionItems(source, options) {
    var opts = options || {};
    var enterGroups = (opts.enterGroups !== false);
    var enterClipGroups = (opts.enterClipGroups === undefined) ? enterGroups : (opts.enterClipGroups === true);
    var accept = opts.accept || function (item) {
        return item.typename !== "GroupItem" && item.typename !== "Layer";
    };
    var collected = [];

    /**
     * 集めた配列に加える（unique のときは同じ参照を足さない）
     * @param {PageItem} item - 加えるオブジェクト
     * @returns {void}
     */
    function pushItem(item) {
        if (opts.unique !== false) {
            for (var k = 0; k < collected.length; k++) {
                if (collected[k] === item) return;
            }
        }
        collected.push(item);
    }

    /**
     * 設定に従って外すオブジェクトか判定する
     * @param {PageItem} item - 判定するオブジェクト
     * @returns {boolean} 外すなら true
     */
    function isSkipped(item) {
        try {
            if (item.typename === "Layer") {
                if (opts.skipLocked && item.locked) return true;
                if (opts.skipHidden && !item.visible) return true;
                return false;
            }
            if (opts.skipLocked && item.locked) return true;
            if (opts.skipHidden && item.hidden) return true;
            if (opts.skipGuides && item.guides === true) return true;
            if (opts.skipClipMasks && isClipMaskItem(item)) return true;
        } catch (e) {
            /* 読めないプロパティは「外さない」に倒す / Unreadable properties do not exclude */
        }
        return false;
    }

    /**
     * 1件をたどって集める
     * @param {PageItem} item - 対象のオブジェクト
     * @returns {void}
     */
    function visit(item) {
        if (!item) return;
        var typeName = "";
        try { typeName = item.typename; } catch (e) { return; }

        if (typeName === "TextRange" || typeName === "InsertionPoint") {
            if (opts.textRangeToFrame === false) {
                if (accept(item)) pushItem(item);
                return;
            }
            visit(resolveTextRangeFrame(item));
            return;
        }
        if (isSkipped(item)) return;
        if (accept(item)) {
            pushItem(item);
            return;
        }

        var children = null;
        if (typeName === "GroupItem") {
            var isClipped = false;
            try { isClipped = (item.clipped === true); } catch (e2) { }
            if (isClipped ? enterClipGroups : enterGroups) children = item.pageItems;
        } else if (typeName === "CompoundPathItem") {
            if (opts.enterCompoundPaths) children = item.pathItems;
        } else if (typeName === "Layer") {
            /* 重なり順はサブレイヤーとページアイテムで別々なので、ページアイテム→サブレイヤーの順にする
               Page items and sublayers stack separately; visit page items first, then sublayers */
            walk(item.pageItems);
            walk(item.layers);
            return;
        }
        if (children) walk(children);
    }

    /**
     * 集まりの各要素をたどる
     * @param {*} list - 配列またはコレクション
     * @returns {void}
     */
    function walk(list) {
        var listItems = normalizeSelectionItems(list);
        for (var i = 0; i < listItems.length; i++) visit(listItems[i]);
    }

    walk(source);
    return collected;
}

/**
 * テキストフレームの種類を "point" / "area" / "path" で返す
 * @param {TextFrame} textFrame - テキストフレーム
 * @returns {string} 種類のキー（判定できなければ ""）
 */
function getTextFrameKindKey(textFrame) {
    try {
        if (textFrame.kind === TextType.POINTTEXT) return "point";
        if (textFrame.kind === TextType.AREATEXT) return "area";
        if (textFrame.kind === TextType.PATHTEXT) return "path";
    } catch (e) { /* kind を読めない / kind is unreadable */ }
    return "";
}

/**
 * 選択からテキストフレームを集める（グループの中・文字カーソルの選択を含む）
 * @param {*} source - doc.selection など
 * @param {Object} [options] - collectSelectionItems と同じ設定に加えて次を受ける
 * @param {string[]} [options.kinds] - 集める種類（"point" / "area" / "path"。既定はすべて）
 * @returns {TextFrame[]} テキストフレーム（前面→背面の順）
 */
function collectSelectionTextFrames(source, options) {
    var opts = {};
    var sourceOptions = options || {};
    for (var key in sourceOptions) {
        if (sourceOptions.hasOwnProperty(key)) opts[key] = sourceOptions[key];
    }
    var kindFilter = null;
    if (opts.kinds && opts.kinds.length) {
        kindFilter = {};
        for (var i = 0; i < opts.kinds.length; i++) kindFilter[opts.kinds[i]] = true;
    }
    opts.accept = function (item) {
        if (item.typename !== "TextFrame") return false;
        return !kindFilter || kindFilter[getTextFrameKindKey(item)] === true;
    };
    /* 種類で外したテキストは中をたどらない（accept が false でも子は無い） / Text frames have no children to walk */
    return collectSelectionItems(source, opts);
}

/**
 * 選択からパスを集める（グループの中を含む）
 * @param {*} source - doc.selection など
 * @param {Object} [options] - collectSelectionItems と同じ設定に加えて次を受ける
 * @param {string} [options.compoundPaths] - 複合パスの扱い。"children"（中のパス、既定）/ "whole"（複合パスごと）/ "skip"（外す）
 * @returns {Array} PathItem（"whole" のときは CompoundPathItem も）の配列
 */
function collectSelectionPathItems(source, options) {
    var opts = {};
    var sourceOptions = options || {};
    for (var key in sourceOptions) {
        if (sourceOptions.hasOwnProperty(key)) opts[key] = sourceOptions[key];
    }
    var compoundMode = opts.compoundPaths || "children";
    opts.enterCompoundPaths = (compoundMode === "children");
    opts.accept = function (item) {
        if (item.typename === "PathItem") return true;
        return compoundMode === "whole" && item.typename === "CompoundPathItem";
    };
    return collectSelectionItems(source, opts);
}

/**
 * クリッピングマスク（クリップグループの型）か判定する
 * パスは clipping、複合パスは中の先頭パスの clipping、テキストは clipping が無いので「クリップグループの先頭」で見る
 * @param {PageItem} item - 判定するオブジェクト
 * @returns {boolean} マスクなら true
 */
function isClipMaskItem(item) {
    try {
        if (item.typename === "PathItem") return item.clipping === true;
        if (item.typename === "CompoundPathItem") {
            return item.pathItems.length > 0 && item.pathItems[0].clipping === true;
        }
        if (item.typename === "TextFrame") {
            var parentGroup = item.parent;
            return parentGroup.typename === "GroupItem" && parentGroup.clipped === true &&
                parentGroup.pageItems.length > 0 && parentGroup.pageItems[0] === item;
        }
    } catch (e) { /* 読めない種類はマスクではない / unreadable kinds are not masks */ }
    return false;
}

/**
 * クリップグループの型（マスク）を返す
 * フラグで探し、見つからなければ先頭（pageItems[0]）を返す（型は常に最前面。テキストの型はフラグを持たない）
 * @param {GroupItem} groupItem - 対象のグループ
 * @returns {PageItem|null} マスク（クリップグループでなければ null）
 */
function getClipMaskItem(groupItem) {
    try {
        if (!groupItem || groupItem.typename !== "GroupItem" || groupItem.clipped !== true) return null;
        var groupChildren = groupItem.pageItems;
        if (groupChildren.length === 0) return null;
        for (var i = 0; i < groupChildren.length; i++) {
            var childType = groupChildren[i].typename;
            if ((childType === "PathItem" || childType === "CompoundPathItem") && isClipMaskItem(groupChildren[i])) {
                return groupChildren[i];
            }
        }
        return groupChildren[0];
    } catch (e) {
        return null;
    }
}

/**
 * グループの中（入れ子を含む）にクリップグループがあるか判定する
 * @param {GroupItem} groupItem - 対象のグループ
 * @returns {boolean} あれば true
 */
function hasClippedDescendant(groupItem) {
    try {
        var groupChildren = groupItem.pageItems;
        for (var i = 0; i < groupChildren.length; i++) {
            if (groupChildren[i].typename !== "GroupItem") continue;
            if (groupChildren[i].clipped === true || hasClippedDescendant(groupChildren[i])) return true;
        }
    } catch (e) { /* 中を読めない / cannot read the children */ }
    return false;
}

/**
 * 環境設定の［プレビュー境界を使用］を読む
 * @returns {boolean} オンなら true（読めなければ false）
 */
function readUsePreviewBoundsPreference() {
    try {
        return app.preferences.getBooleanPreference("includeStrokeInBounds");
    } catch (e) {
        return false;
    }
}

/**
 * 見た目どおりの境界を返す。クリップグループはマスクの境界、
 * 中にクリップグループを含むグループは子の境界を合わせたもの（隠れた部分を含めない）
 * @param {PageItem} item - 対象のオブジェクト
 * @param {boolean} [usePreviewBounds] - true で visibleBounds、false で geometricBounds（省略時は環境設定に従う）
 * @returns {number[]|null} [左, 上, 右, 下] の新しい配列（測れなければ null）
 */
function getClipAwareBounds(item, usePreviewBounds) {
    var usePreview = (usePreviewBounds === undefined || usePreviewBounds === null) ?
        readUsePreviewBoundsPreference() : (usePreviewBounds === true);
    try {
        var measuredItem = item;
        if (item.typename === "GroupItem") {
            var maskItem = getClipMaskItem(item);
            if (maskItem) {
                measuredItem = maskItem;
            } else if (hasClippedDescendant(item)) {
                /* グループ自体の効果（影など）の広がりは含まれなくなる
                   This leaves out the reach of effects applied to the group itself (drop shadows etc.) */
                var childBounds = getClipAwareUnionBounds(filterMeasurableChildren(item.pageItems), usePreview);
                if (childBounds) return childBounds;
            }
        }
        var bounds = usePreview ? measuredItem.visibleBounds : measuredItem.geometricBounds;
        return [bounds[0], bounds[1], bounds[2], bounds[3]];
    } catch (e) {
        return null;
    }
}

/**
 * 境界の計算に入れる子だけを残す（非表示とガイドを外す）
 * @param {*} childList - 子のコレクション
 * @returns {Array} 残した子
 */
function filterMeasurableChildren(childList) {
    var childItems = normalizeSelectionItems(childList);
    var measurable = [];
    for (var i = 0; i < childItems.length; i++) {
        try {
            if (childItems[i].hidden === true || childItems[i].guides === true) continue;
        } catch (e) { /* 読めなければ残す / keep when unreadable */ }
        measurable.push(childItems[i]);
    }
    return measurable;
}

/**
 * 複数のオブジェクトを囲む外接範囲を返す（クリップグループはマスクで測る）
 * @param {*} items - オブジェクトの配列・コレクション・選択
 * @param {boolean} [usePreviewBounds] - true で visibleBounds、false で geometricBounds（省略時は環境設定に従う）
 * @returns {number[]|null} [左, 上, 右, 下]（測れるものが無ければ null）
 */
function getClipAwareUnionBounds(items, usePreviewBounds) {
    var usePreview = (usePreviewBounds === undefined || usePreviewBounds === null) ?
        readUsePreviewBoundsPreference() : (usePreviewBounds === true);
    var itemList = normalizeSelectionItems(items);
    var unionBounds = null;
    for (var i = 0; i < itemList.length; i++) {
        var itemBounds = getClipAwareBounds(itemList[i], usePreview);
        if (!itemBounds) continue;
        if (!unionBounds) {
            unionBounds = itemBounds;
            continue;
        }
        if (itemBounds[0] < unionBounds[0]) unionBounds[0] = itemBounds[0];
        if (itemBounds[1] > unionBounds[1]) unionBounds[1] = itemBounds[1];
        if (itemBounds[2] > unionBounds[2]) unionBounds[2] = itemBounds[2];
        if (itemBounds[3] < unionBounds[3]) unionBounds[3] = itemBounds[3];
    }
    return unionBounds;
}

/**
 * 2つの座標を許容値つきで比べる
 * @param {number} valueA - 座標A（pt）
 * @param {number} valueB - 座標B（pt）
 * @param {number} [tolerance] - 許容値（pt、既定は SELECTION_ITEMS_TOLERANCE）
 * @returns {boolean} 差が許容値以下なら true
 */
function isNearlySameCoordinate(valueA, valueB, tolerance) {
    var limit = (typeof tolerance === "number") ? tolerance : SELECTION_ITEMS_TOLERANCE;
    return Math.abs(valueA - valueB) <= limit;
}

/**
 * 2つの境界を許容値つきで比べる
 * @param {number[]} boundsA - [左, 上, 右, 下]
 * @param {number[]} boundsB - [左, 上, 右, 下]
 * @param {number} [tolerance] - 許容値（pt、既定は SELECTION_ITEMS_TOLERANCE）
 * @returns {boolean} 4辺とも許容値以内なら true
 */
function areBoundsNearlyEqual(boundsA, boundsB, tolerance) {
    if (!boundsA || !boundsB) return false;
    for (var i = 0; i < 4; i++) {
        if (!isNearlySameCoordinate(boundsA[i], boundsB[i], tolerance)) return false;
    }
    return true;
}

// 選択の収集と境界（再利用パーツ）ここまで / End of the reusable selection items and bounds

// =========================================
// ビューの計算 / View calculation
// =========================================

/**
 * 対象を画面に収めるズーム倍率を、現在のズームと表示領域から求める
 * （現在の表示を基準にするので、開始位置からそのまま動かせる）
 * @param {number} startZoom - 現在のズーム倍率
 * @param {number} visibleWidth - 現在の表示領域の幅（ドキュメント座標）
 * @param {number} visibleHeight - 現在の表示領域の高さ（ドキュメント座標）
 * @param {number} targetWidth - 収めたい範囲の幅
 * @param {number} targetHeight - 収めたい範囲の高さ
 * @returns {number} 幅・高さの両方が収まるズーム倍率（求められないときは startZoom）
 */
function computeFitZoom(startZoom, visibleWidth, visibleHeight, targetWidth, targetHeight) {
    var fitZoom = null;

    if (targetWidth > 0 && visibleWidth > 0) {
        fitZoom = startZoom * visibleWidth / targetWidth;
    }
    if (targetHeight > 0 && visibleHeight > 0) {
        var fitZoomByHeight = startZoom * visibleHeight / targetHeight;
        if (fitZoom === null || fitZoomByHeight < fitZoom) {
            fitZoom = fitZoomByHeight;
        }
    }

    return (fitZoom === null) ? startZoom : fitZoom;
}

// =========================================
// アニメーション / Animation
// =========================================

/**
 * 指定ミリ秒だけ待機する（アニメーションのフレーム間隔）
 * @param {number} durationMs - 待機するミリ秒
 * @returns {void}
 */
function sleep(durationMs) {
    var startTime = new Date().getTime();
    while (new Date().getTime() - startTime < durationMs) { }
}

/**
 * 移動量に応じて補間ステップ数を決める（小さな移動はステップを減らしてもたつきを防ぐ）
 * 移動距離（現在の表示領域に対する割合）とズームの変化率の大きい方を「移動量」とする
 * @param {View} activeView - 対象のビュー
 * @param {number} startCenterX - 開始時の中心 X
 * @param {number} startCenterY - 開始時の中心 Y
 * @param {number} startZoom - 開始時のズーム倍率
 * @param {number} targetCenterX - 目標の中心 X
 * @param {number} targetCenterY - 目標の中心 Y
 * @param {number} targetZoom - 目標のズーム倍率
 * @returns {number} 補間ステップ数
 */
function resolveStepCount(activeView, startCenterX, startCenterY, startZoom, targetCenterX, targetCenterY, targetZoom) {
    var viewBounds = activeView.bounds; /* [left, top, right, bottom] */
    var visibleExtent = Math.max(
        Math.abs(viewBounds[2] - viewBounds[0]),
        Math.abs(viewBounds[1] - viewBounds[3])
    );

    var dx = targetCenterX - startCenterX;
    var dy = targetCenterY - startCenterY;
    var moveDistance = Math.sqrt(dx * dx + dy * dy);
    var moveFraction = (visibleExtent > 0) ? (moveDistance / visibleExtent) : 1;

    var maxZoom = Math.max(startZoom, targetZoom);
    var zoomFraction = (maxZoom > 0) ? (Math.abs(targetZoom - startZoom) / maxZoom) : 0;

    var magnitude = Math.min(Math.max(moveFraction, zoomFraction), 1);

    var minSteps = Math.max(2, Math.round(MAX_ANIMATION_STEP_COUNT * 0.2));
    var resolvedSteps = Math.round(MAX_ANIMATION_STEP_COUNT * magnitude);
    if (resolvedSteps < minSteps) resolvedSteps = minSteps;
    if (resolvedSteps > MAX_ANIMATION_STEP_COUNT) resolvedSteps = MAX_ANIMATION_STEP_COUNT;
    return resolvedSteps;
}

/**
 * 現在のビューから目標の中心・ズームへ少しずつ補間して動かす
 * @param {View} activeView - 対象のビュー
 * @param {number} targetCenterX - 目標の中心 X
 * @param {number} targetCenterY - 目標の中心 Y
 * @param {number} targetZoom - 目標のズーム倍率
 * @returns {void}
 */
function animateView(activeView, targetCenterX, targetCenterY, targetZoom) {
    var startCenter = activeView.centerPoint;
    var startCenterX = startCenter[0];
    var startCenterY = startCenter[1];
    var startZoom = activeView.zoom;

    /* 移動量に応じてステップ数を決める / Decide the step count from the amount of movement */
    var stepCount = resolveStepCount(
        activeView, startCenterX, startCenterY, startZoom,
        targetCenterX, targetCenterY, targetZoom
    );

    /* ズームは倍率を一定割合ずつ動かすと滑らかに見えるので、線形ではなく幾何補間にする
       Zoom is interpolated geometrically, which looks smoother than linear */
    var zoomRatio = (startZoom > 0) ? (targetZoom / startZoom) : 1;

    for (var i = 1; i <= stepCount; i++) {
        var progress = i / stepCount;
        /* easeOutQuad：終わりに向かって減速 / easeOutQuad: slow down toward the end */
        var easedProgress = 1 - (1 - progress) * (1 - progress);

        activeView.centerPoint = [
            startCenterX + (targetCenterX - startCenterX) * easedProgress,
            startCenterY + (targetCenterY - startCenterY) * easedProgress
        ];
        /* 幾何補間：startZoom × (targetZoom / startZoom)^t / Geometric interpolation */
        activeView.zoom = startZoom * Math.pow(zoomRatio, easedProgress);

        app.redraw();
        sleep(FRAME_DELAY_MS);
    }

    /* 最後に正確な値へ合わせる / Snap to the exact target at the end */
    activeView.centerPoint = [targetCenterX, targetCenterY];
    activeView.zoom = targetZoom;
    app.redraw();
}

// =========================================
// メイン処理 / Main
// =========================================

/**
 * 選択範囲にズームする。選択が無いときは現在位置のまま 100% に戻す
 * @returns {void}
 */
function main() {
    if (app.documents.length < 1) {
        alert(getLabel("alert.noDocument"));
        return;
    }
    var doc = app.activeDocument;
    var selectedItems = doc.selection;
    var activeView = doc.views[0];

    if (!(selectedItems && selectedItems.length > 0)) {
        /* 選択が無いときは現在位置のまま 100% へ（こちらも滑らかに） / No selection: animate back to 100% in place */
        var currentCenter = activeView.centerPoint;
        animateView(activeView, currentCenter[0], currentCenter[1], 1);
        return;
    }

    /* visibleBounds 基準の外接矩形 [左, 上, 右, 下]。クリップグループはマスクの範囲
       Union of visible bounds [L, T, R, B]; a clip group is measured by its mask */
    var selectionBounds = getClipAwareUnionBounds(selectedItems, true);
    if (!selectionBounds) return;
    var selectionWidth = selectionBounds[2] - selectionBounds[0];
    var selectionHeight = selectionBounds[1] - selectionBounds[3];

    /* 現在の表示領域（ドキュメント座標系）を基準にフィット倍率を求める / Fit zoom from the current visible area */
    var viewBounds = activeView.bounds; /* [left, top, right, bottom] */
    var targetZoom = computeFitZoom(
        activeView.zoom,
        viewBounds[2] - viewBounds[0],
        viewBounds[1] - viewBounds[3],
        selectionWidth,
        selectionHeight
    ) * ZOOM_FIT_RATIO;

    animateView(
        activeView,
        selectionBounds[0] + (selectionWidth / 2),
        selectionBounds[1] - (selectionHeight / 2),
        targetZoom
    );
}

main();

})();
