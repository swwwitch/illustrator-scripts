#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択されているオブジェクト群の端または中心を取得し、下に揃えます。
揃え先は、アクティブアートボードの端、または条件に合うガイドです。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GroupEdgeAlignBOTTOM.md

### Overview

Takes the edges or the center of the selected objects and aligns them to the bottom edge.
The target is the edge of the active artboard, or a matching guide.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GroupEdgeAlignBOTTOM.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "GroupEdgeAlignBOTTOM";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-04-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GroupEdgeAlignBOTTOM.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GroupEdgeAlignBOTTOM.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    /* アートボード端の手前にガイドがあれば、そのガイドへ揃える / snap to a guide when one is in the way */
    var USE_GUIDES = true;

    /* 境界の取り方 "preference"（環境設定に従う）| "preview"（線を含む）| "geometric"（線を含まない）*/
    var BOUNDS_MODE = "preference";

    /* ガイドの探し方 "inside"（揃える向きで最も近いもの）| "nearest"（距離が最も近いもの）*/
    var GUIDE_SEARCH_MODE = "inside";

    /* ガイドの水平・垂直判定に使う許容値（pt）/ tolerance for classifying a guide as H or V */
    var GUIDE_ORIENTATION_TOLERANCE = 0.01;

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
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "オブジェクトが選択されていません。", en: "No objects are selected." },
            invalidBoundsMode: { ja: "BOUNDS_MODE の指定が不正です", en: "Invalid BOUNDS_MODE" },
            invalidGuideSearchMode: { ja: "GUIDE_SEARCH_MODE の指定が不正です", en: "Invalid GUIDE_SEARCH_MODE" },
            invalidAlignmentSide: { ja: "揃える方向を判定できませんでした", en: "Could not determine the alignment direction" }
        }
    };

    /**
     * ラベルに言語別のコロンと値を続けた文字列を返す（日本語は「 : 」、英語は「: 」）
     * @param {Object} labelSet - LABELS のリーフ（{ ja, en }）
     * @param {string} value - コロンの後ろに続ける値
     * @returns {string} 「ラベル：値」の文字列
     */
    function labelWithValue(labelSet, value) {
        return getLabel(labelSet) + (uiLang === "ja" ? " : " : ": ") + value;
    }

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
    // メイン処理 / Main
    // =========================================

    /**
     * 選択オブジェクト群の端または中心を、ファイル名から判定した方向へ揃える
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var documentRef = app.activeDocument;
        var selectedItems = documentRef.selection;
        if (selectedItems.length === 0) {
            alert(getLabel(LABELS.alert.noSelection));
            return;
        }

        var alignmentSide = resolveAlignmentSideFromFileName();

        var includeStrokeInBounds = resolveIncludeStrokeInBounds(BOUNDS_MODE);
        if (includeStrokeInBounds === null) {
            alert(labelWithValue(LABELS.alert.invalidBoundsMode, BOUNDS_MODE));
            return;
        }

        var artboards = documentRef.artboards;
        var activeArtboardRect = artboards[artboards.getActiveArtboardIndex()].artboardRect;
        var selectionBounds = getSelectionBounds(selectedItems, includeStrokeInBounds);

        var moveOffset = getAlignmentOffset(documentRef, selectionBounds, activeArtboardRect, alignmentSide);
        if (moveOffset === null) return; /* 方向の指定が不正。getAlignmentOffset が通知済み */

        for (var i = 0; i < selectedItems.length; i++) {
            selectedItems[i].translate(moveOffset.x, moveOffset.y);
        }
    }

    /**
     * スクリプトのファイル名から揃える方向を判定する（例: GroupEdgeAlignRIGHT.jsx → "right"）
     * CENTERX / CENTERY は CENTER より先に判定する必要がある。
     * @returns {string} 揃える方向。判定できなければ "left"
     */
    function resolveAlignmentSideFromFileName() {
        var fileNameUpper = File($.fileName).name.toUpperCase();

        if (fileNameUpper.indexOf("CENTERX") !== -1) return "CENTER_X";
        if (fileNameUpper.indexOf("CENTERY") !== -1) return "CENTER_Y";
        if (fileNameUpper.indexOf("CENTER") !== -1) return "CENTER";
        if (fileNameUpper.indexOf("LEFT") !== -1) return "left";
        if (fileNameUpper.indexOf("RIGHT") !== -1) return "right";
        if (fileNameUpper.indexOf("TOP") !== -1) return "top";
        if (fileNameUpper.indexOf("BOTTOM") !== -1) return "bottom";
        return "left";
    }

    /**
     * BOUNDS_MODE から「線を境界に含めるか」を決める
     * @param {string} boundsMode - "preference" | "preview" | "geometric"
     * @returns {boolean|null} 線を含めるなら true。指定が不正なら null
     */
    function resolveIncludeStrokeInBounds(boundsMode) {
        if (boundsMode === "preference") return app.preferences.getBooleanPreference("includeStrokeInBounds");
        if (boundsMode === "preview") return true;
        if (boundsMode === "geometric") return false;
        return null;
    }

    /**
     * 選択オブジェクト全体を囲む境界を返す
     * @param {PageItem[]} selectedItems - 対象のオブジェクト
     * @param {boolean} includeStrokeInBounds - 線を境界に含めるか
     * @returns {number[]} [left, top, right, bottom]
     */
    function getSelectionBounds(selectedItems, includeStrokeInBounds) {
        var selectionBounds = getItemBounds(selectedItems[0], includeStrokeInBounds).slice(0);

        for (var i = 1; i < selectedItems.length; i++) {
            var itemBounds = getItemBounds(selectedItems[i], includeStrokeInBounds);
            if (itemBounds[0] < selectionBounds[0]) selectionBounds[0] = itemBounds[0];
            if (itemBounds[1] > selectionBounds[1]) selectionBounds[1] = itemBounds[1];
            if (itemBounds[2] > selectionBounds[2]) selectionBounds[2] = itemBounds[2];
            if (itemBounds[3] < selectionBounds[3]) selectionBounds[3] = itemBounds[3];
        }

        return selectionBounds;
    }

    /**
     * 指定オブジェクトの境界を返す
     * クリップグループは線を含めるかどうかによらず、マスクの geometricBounds を使う。
     * 中にクリップグループを含むグループは、子の境界を合わせて測る（隠れた部分を含めない）。
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @param {boolean} includeStrokeInBounds - 線を境界に含めるか
     * @returns {number[]} 境界 [L, T, R, B]
     */
    function getItemBounds(pageItem, includeStrokeInBounds) {
        /* クリップグループのマスクは常に幾何境界で測る / Clip-group masks are always measured by geometric bounds */
        var isClipGroup = (getClipMaskItem(pageItem) !== null);
        return getClipAwareBounds(pageItem, isClipGroup ? false : includeStrokeInBounds);
    }

    /**
     * 揃える方向から移動量を求める
     * @param {Document} documentRef - 対象のドキュメント
     * @param {number[]} selectionBounds - 選択範囲の境界 [L, T, R, B]
     * @param {number[]} artboardRect - アクティブアートボードの矩形 [L, T, R, B]
     * @param {string} alignmentSide - 揃える方向
     * @returns {{x: number, y: number}|null} 移動量。方向の指定が不正なら null
     */
    function getAlignmentOffset(documentRef, selectionBounds, artboardRect, alignmentSide) {
        var centerOffsetX = getHorizontalCenterValue(artboardRect) - getHorizontalCenterValue(selectionBounds);
        var centerOffsetY = getVerticalCenterValue(artboardRect) - getVerticalCenterValue(selectionBounds);

        if (alignmentSide === "CENTER_X") return { x: centerOffsetX, y: 0 };
        if (alignmentSide === "CENTER_Y") return { x: 0, y: centerOffsetY };
        if (alignmentSide === "CENTER") return { x: centerOffsetX, y: centerOffsetY };

        var selectionEdge = getEdgeValueForAlignmentSide(selectionBounds, alignmentSide);
        var artboardEdge = getEdgeValueForAlignmentSide(artboardRect, alignmentSide);
        if (selectionEdge === null || artboardEdge === null) {
            alert(labelWithValue(LABELS.alert.invalidAlignmentSide, alignmentSide));
            return null;
        }

        var targetEdge = artboardEdge;
        if (USE_GUIDES) {
            var snappedGuideValue = findGuideSnapValue(documentRef, artboardRect, selectionEdge, alignmentSide);
            if (snappedGuideValue !== null) targetEdge = snappedGuideValue;
        }

        var edgeOffset = targetEdge - selectionEdge;
        return isHorizontalSide(alignmentSide) ? { x: edgeOffset, y: 0 } : { x: 0, y: edgeOffset };
    }

    /**
     * 揃える方向が左右（X方向の移動）かを返す
     * @param {string} alignmentSide - 揃える方向
     * @returns {boolean} "left" / "right" なら true
     */
    function isHorizontalSide(alignmentSide) {
        return alignmentSide === "left" || alignmentSide === "right";
    }

    /**
     * 指定した方向に対応する境界値を返す
     * @param {number[]} boundsRect - 境界 [L, T, R, B]
     * @param {string} alignmentSide - 揃える方向
     * @returns {number|null} 境界値。方向が不正なら null
     */
    function getEdgeValueForAlignmentSide(boundsRect, alignmentSide) {
        if (alignmentSide === "left") return boundsRect[0];
        if (alignmentSide === "top") return boundsRect[1];
        if (alignmentSide === "right") return boundsRect[2];
        if (alignmentSide === "bottom") return boundsRect[3];
        return null;
    }

    /**
     * 左右中央の座標を返す
     * @param {number[]} boundsRect - 境界 [L, T, R, B]
     * @returns {number} 左右中央のX
     */
    function getHorizontalCenterValue(boundsRect) {
        return (boundsRect[0] + boundsRect[2]) / 2;
    }

    /**
     * 上下中央の座標を返す
     * @param {number[]} boundsRect - 境界 [L, T, R, B]
     * @returns {number} 上下中央のY
     */
    function getVerticalCenterValue(boundsRect) {
        return (boundsRect[1] + boundsRect[3]) / 2;
    }

    /**
     * アートボード内側にあるガイドのうち、揃える方向と GUIDE_SEARCH_MODE に合う吸着先座標を返す
     * @param {Document} documentRef - 対象のドキュメント
     * @param {number[]} artboardRect - アクティブアートボードの矩形 [L, T, R, B]
     * @param {number} selectionEdge - 選択範囲の該当する端の座標
     * @param {string} alignmentSide - 揃える方向（"left" | "right" | "top" | "bottom"）
     * @returns {number|null} 吸着先の座標。該当するガイドがなければ null
     */
    function findGuideSnapValue(documentRef, artboardRect, selectionEdge, alignmentSide) {
        if (GUIDE_SEARCH_MODE !== "inside" && GUIDE_SEARCH_MODE !== "nearest") {
            alert(labelWithValue(LABELS.alert.invalidGuideSearchMode, GUIDE_SEARCH_MODE));
            return null;
        }

        /* アートボード端へ向かう向きの符号。left / bottom は座標が減る向き / sign toward the artboard edge */
        var outwardSign = (alignmentSide === "left" || alignmentSide === "bottom") ? -1 : 1;
        var nearestGuideValue = null;
        var nearestGuideDistance = null;
        var documentPathItems = documentRef.pathItems;

        for (var i = 0; i < documentPathItems.length; i++) {
            var guidePathItem = documentPathItems[i];
            if (guidePathItem.guides !== true) continue;

            var guideValue = getGuideValueForAlignmentSide(guidePathItem.geometricBounds, alignmentSide);
            if (guideValue === null) continue;
            if (!isGuideValueInsideArtboard(guideValue, artboardRect, alignmentSide)) continue;

            if (GUIDE_SEARCH_MODE === "inside") {
                /* 選択範囲より揃える向きの側にあり、その中で選択範囲に最も近いもの / nearest guide on the outward side */
                if ((guideValue - selectionEdge) * outwardSign <= 0) continue;
                if (nearestGuideValue === null || (guideValue - nearestGuideValue) * outwardSign < 0) {
                    nearestGuideValue = guideValue;
                }
            } else {
                var guideDistance = Math.abs(guideValue - selectionEdge);
                if (nearestGuideDistance === null || guideDistance < nearestGuideDistance) {
                    nearestGuideValue = guideValue;
                    nearestGuideDistance = guideDistance;
                }
            }
        }

        return nearestGuideValue;
    }

    /**
     * 揃える方向に対応するガイド座標を返す（左右なら垂直ガイドのX、上下なら水平ガイドのY）
     * @param {number[]} guideBounds - ガイドの境界 [L, T, R, B]
     * @param {string} alignmentSide - 揃える方向（"left" | "right" | "top" | "bottom"）
     * @returns {number|null} ガイドの座標。向きが対応しなければ null
     */
    function getGuideValueForAlignmentSide(guideBounds, alignmentSide) {
        if (isHorizontalSide(alignmentSide)) {
            var isVerticalGuide = Math.abs(guideBounds[2] - guideBounds[0]) <= GUIDE_ORIENTATION_TOLERANCE;
            return isVerticalGuide ? guideBounds[0] : null;
        }
        var isHorizontalGuide = Math.abs(guideBounds[1] - guideBounds[3]) <= GUIDE_ORIENTATION_TOLERANCE;
        return isHorizontalGuide ? guideBounds[1] : null;
    }

    /**
     * ガイド座標がアクティブアートボード内にあるかを返す
     * @param {number} guideValue - ガイドの座標
     * @param {number[]} artboardRect - アクティブアートボードの矩形 [L, T, R, B]
     * @param {string} alignmentSide - 揃える方向（"left" | "right" | "top" | "bottom"）
     * @returns {boolean} 内側にあれば true
     */
    function isGuideValueInsideArtboard(guideValue, artboardRect, alignmentSide) {
        if (isHorizontalSide(alignmentSide)) {
            return guideValue >= artboardRect[0] && guideValue <= artboardRect[2];
        }
        return guideValue <= artboardRect[1] && guideValue >= artboardRect[3];
    }

    main();

})();
