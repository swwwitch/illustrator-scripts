#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

バラバラに分割されたテキストオブジェクトを行単位にまとめ、タブ区切りの1つのエリア内文字に再構成します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextMergeToAreaBox-tab.md

note記事も参照してください。
https://note.com/dtp_tranist/n/ne8d31278c266

### Overview

Gathers scattered text objects line by line and rebuilds them as a single tab-separated area text.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextMergeToAreaBox-tab.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TextMergeToAreaBox-tab";       /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-18";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-21";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextMergeToAreaBox-tab.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextMergeToAreaBox-tab.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/ne8d31278c266"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 同じ行と見なすY座標差の閾値（pt） / Y threshold for grouping items into the same line (pt) */
    var LINE_Y_THRESHOLD = 5;

    /* 行送りの最小倍率（フォントサイズ基準） / Minimum leading ratio against the font size */
    var MIN_LEADING_RATIO = 1.2;

    /* 同じ行の断片どうしの区切り / Separator between the items in a line */
    var CELL_SEPARATOR = "\t";

    /* 断片の中の改行を置き換える文字（行の区切りと混ざらないように）
       Replaces line breaks inside an item so they do not mix with the row breaks */
    var LINE_BREAK_PLACEHOLDER = "〓";

    /* エリア内文字に適用する禁則 / Kinsoku applied to the area text */
    var DEFAULT_KINSOKU  = "Soft_v2"; /* 弱い禁則 v2 / Loose v2 */
    var FALLBACK_KINSOKU = "Soft";    /* v2がないバージョン向け / For versions without v2 */

    /* エリア内文字に適用する行揃え / Justification applied to the area text */
    var AREA_TEXT_JUSTIFICATION = Justification.FULLJUSTIFYLASTLINELEFT; /* 両端揃え（最終行左揃え） / Justify, last line left */

    /* あふれ解消で枠を下に伸ばす最大回数 / Maximum number of times the frame is extended downward to clear overflow */
    var MAX_HEIGHT_GROW_STEPS = 20;

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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
     * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    var LABELS = {
        alert: {
            noText: { ja: "変換できるテキストはありません。", en: "No convertible text found." }
        }
    };

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // 選択の収集と境界（再利用パーツ） / Selection items and bounds (reusable)
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // 選択の収集と境界（再利用パーツ）ここまで / End of the reusable selection items and bounds
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // 行の分類 / Grouping into lines
    // =========================================

    /**
     * テキストフレームを左から右へ並び替える
     * @param {Array<TextFrame>} framesToSort - 対象テキストフレーム
     * @returns {Array<TextFrame>} 並び替えた新しい配列
     */
    function sortFramesLeftToRight(framesToSort) {
        return framesToSort.slice(0).sort(function (frameA, frameB) {
            return frameA.left - frameB.left;
        });
    }

    /**
     * テキストフレームを上から下へ並び替える
     * @param {Array<TextFrame>} framesToSort - 対象テキストフレーム
     * @returns {Array<TextFrame>} 並び替えた新しい配列
     */
    function sortFramesTopToBottom(framesToSort) {
        return framesToSort.slice(0).sort(function (frameA, frameB) {
            return frameB.position[1] - frameA.position[1];
        });
    }

    /**
     * 上から下に並べたフレームを、Y位置が近いものごとに同じ行としてまとめる
     * @param {Array<TextFrame>} framesTopToBottom - 上から下へ並び替え済みのテキストフレーム
     * @param {number} yThreshold - 同じ行と見なすY座標差の閾値（pt）
     * @returns {Array<Array<TextFrame>>} 行ごとのテキストフレーム配列
     */
    function groupFramesIntoLines(framesTopToBottom, yThreshold) {
        var lineGroups = [];
        for (var i = 0; i < framesTopToBottom.length; i++) {
            var frame = framesTopToBottom[i];
            /* 上から順に並んでいるので、直前の行とだけ比較すればよい
               The frames are sorted top-down, so comparing with the last line is enough */
            var lastLine = lineGroups[lineGroups.length - 1];
            if (lastLine && Math.abs(lastLine[0].position[1] - frame.position[1]) <= yThreshold) {
                lastLine.push(frame);
            } else {
                lineGroups.push([frame]);
            }
        }
        return lineGroups;
    }

    /**
     * 選択からテキストフレームだけを集め、Y位置で行単位に分類する
     * @param {Array<PageItem>} selectedItems - 選択中のオブジェクト
     * @returns {Array<Array<TextFrame>>} 上から順に並べた行ごとのテキストフレーム配列
     */
    function groupTextFramesByLine(selectedItems) {
        var textFrames = [];
        for (var i = 0; i < selectedItems.length; i++) {
            if (selectedItems[i].typename === "TextFrame") {
                textFrames.push(selectedItems[i]);
            }
        }
        return groupFramesIntoLines(sortFramesTopToBottom(textFrames), LINE_Y_THRESHOLD);
    }

    // =========================================
    // 連結 / Merging
    // =========================================

    /**
     * テキストフレームをまとめて削除する
     * @param {Array<TextFrame>} framesToRemove - 削除するテキストフレーム
     * @returns {void}
     */
    function removeFrames(framesToRemove) {
        for (var i = 0; i < framesToRemove.length; i++) {
            framesToRemove[i].remove();
        }
    }

    /**
     * 1行分のテキストフレームをX順にタブ区切りで連結し、1つのテキストフレームにまとめる
     * @param {Array<TextFrame>} lineFrames - 同じ行に属するテキストフレーム
     * @returns {TextFrame} 連結後のテキストフレーム（元のフレームは削除される）
     */
    function mergeLineFrames(lineFrames) {
        var framesInLine = sortFramesLeftToRight(lineFrames);

        var cellTexts = [];
        for (var i = 0; i < framesInLine.length; i++) {
            cellTexts.push(framesInLine[i].contents.replace(/\r/g, LINE_BREAK_PLACEHOLDER));
        }

        var leftmostFrame = framesInLine[0];
        var originalPosition = leftmostFrame.position;

        var mergedFrame = leftmostFrame.duplicate();
        mergedFrame.orientation = TextOrientation.HORIZONTAL;
        mergedFrame.move(leftmostFrame, ElementPlacement.PLACEBEFORE);
        mergedFrame.contents = cellTexts.join(CELL_SEPARATOR);
        mergedFrame.position = originalPosition;

        removeFrames(framesInLine);
        return mergedFrame;
    }

    /**
     * 各行のテキストを、行ごとに改行して1つの文字列にする（表の行をそのまま保つ）
     * @param {Array<TextFrame>} mergedLineFrames - 行ごとに連結済みのテキストフレーム
     * @returns {string} エリア内文字に流し込む文字列
     */
    function joinLineContents(mergedLineFrames) {
        var lineTexts = [];
        for (var i = 0; i < mergedLineFrames.length; i++) {
            lineTexts.push(mergedLineFrames[i].contents);
        }
        return lineTexts.join("\r");
    }

    // =========================================
    // エリア内文字の作成 / Creating the area text
    // =========================================

    /**
     * 隣り合う2行のY差から行送りを求める
     * @param {Array<TextFrame>} mergedLineFrames - 行ごとに連結済みのテキストフレーム
     * @param {number} fontSize - 基準になるフォントサイズ（pt）
     * @returns {number|null} 行送り（pt）。2行未満のときは null
     */
    function computeLeading(mergedLineFrames, fontSize) {
        if (mergedLineFrames.length < 2) {
            return null;
        }
        /* 行送りとしてY差を使用 / Use the Y gap as the leading */
        var leading = Math.abs(mergedLineFrames[0].position[1] - mergedLineFrames[1].position[1]);
        if (leading < fontSize) {
            /* 最小でも MIN_LEADING_RATIO 倍にする / Keep at least MIN_LEADING_RATIO times the font size */
            leading = fontSize * MIN_LEADING_RATIO;
        }
        return leading;
    }

    /**
     * 禁則を適用する（指定の禁則名が使えないバージョンではフォールバックに切り替える）
     * @param {TextRange} targetRange - 適用先のテキスト範囲
     * @param {string} kinsokuName - 適用したい禁則名
     * @param {string} fallbackKinsokuName - 使えなかったときに適用する禁則名
     * @returns {void}
     */
    function applyKinsoku(targetRange, kinsokuName, fallbackKinsokuName) {
        /* 未対応の禁則名は例外になる / An unsupported kinsoku name throws */
        try {
            targetRange.paragraphAttributes.kinsoku = kinsokuName;
            /* 例外にならず無視される場合もあるため読み戻して確認
               Read the value back, since an unknown name can be ignored instead of throwing */
            if (targetRange.paragraphAttributes.kinsoku === kinsokuName) {
                return;
            }
        } catch (e) { }
        targetRange.paragraphAttributes.kinsoku = fallbackKinsokuName;
    }

    /**
     * エリア内文字からテキストがあふれているかを判定する
     * @param {TextFrame} areaTextFrame - 判定するエリア内文字
     * @returns {boolean} あふれていれば true
     */
    function isTextOverflowing(areaTextFrame) {
        var composedCharacters = 0;
        for (var i = 0; i < areaTextFrame.lines.length; i++) {
            composedCharacters += areaTextFrame.lines[i].characters.length;
        }
        /* 改行コードは各行の文字数に含まれないため、比較対象から除く
           Line breaks are not counted in each line, so exclude them from the total */
        return composedCharacters < areaTextFrame.contents.replace(/[\r\n]/g, "").length;
    }

    /**
     * テキストがあふれなくなるまで枠を下方向に伸ばす（上端は動かさない）
     * @param {TextFrame} areaTextFrame - 対象のエリア内文字
     * @param {number} heightStep - 1回あたりに伸ばす高さ（pt）
     * @returns {void}
     */
    function growFrameUntilTextFits(areaTextFrame, heightStep) {
        var frameTop = areaTextFrame.top;
        for (var i = 0; i < MAX_HEIGHT_GROW_STEPS && isTextOverflowing(areaTextFrame); i++) {
            areaTextFrame.height = areaTextFrame.height + heightStep;
            /* 上端は元の位置に固定 / Keep the top edge where it was */
            areaTextFrame.top = frameTop;
        }
    }

    /**
     * エリア内文字にフォント・サイズ・行揃え・禁則・行送りを適用する
     * @param {TextFrame} areaTextFrame - 適用先のエリア内文字
     * @param {TextFrame} sourceFrame - 書式の引き継ぎ元（一番上の行）
     * @param {number} fontSize - 適用するフォントサイズ（pt）
     * @param {number|null} leading - 適用する行送り（pt）。null なら自動行送りのまま
     * @returns {void}
     */
    function applyAreaTextFormatting(areaTextFrame, sourceFrame, fontSize, leading) {
        var targetRange = areaTextFrame.textRange;
        targetRange.characterAttributes.textFont = sourceFrame.textRange.characterAttributes.textFont;
        targetRange.characterAttributes.size = fontSize;

        /* 禁則・行揃えはテキスト全体に適用する（paragraphs[0] だけでは［段落］パネルに反映されない）
           Apply kinsoku and justification to the whole text (paragraphs[0] alone is not reflected in the Paragraph panel) */
        applyKinsoku(targetRange, DEFAULT_KINSOKU, FALLBACK_KINSOKU);
        targetRange.paragraphAttributes.justification = AREA_TEXT_JUSTIFICATION;

        if (leading !== null) {
            targetRange.characterAttributes.autoLeading = false;
            targetRange.characterAttributes.leading = leading;
        }
    }

    /**
     * 行ごとのフレーム全体の外接矩形からエリア内文字を作成し、体裁を引き継ぐ
     * @param {Array<TextFrame>} mergedLineFrames - 行ごとに連結済みのテキストフレーム（2行以上）
     * @param {string} joinedText - 流し込む文字列
     * @returns {TextFrame} 作成したエリア内文字
     */
    function createAreaTextFrame(mergedLineFrames, joinedText) {
        var doc = app.activeDocument;
        var fontSize = mergedLineFrames[0].textRange.characterAttributes.size;

        /* 選択状態に依存しないよう mergedLineFrames から外接矩形を取得
           Read the bounding box from mergedLineFrames so it does not depend on the current selection */
        var lineBounds = getClipAwareUnionBounds(mergedLineFrames, true);
        var boundsLeft = lineBounds[0];
        var boundsTop = lineBounds[1];
        var boundsWidth = lineBounds[2] - boundsLeft;
        var areaHeight = boundsTop - lineBounds[3];

        /* 作成幅は1文字分縮める。縮めると狭くなりすぎる選択では最低1文字分を確保
           Shrink the created width by one character, but keep at least one character for narrow selections */
        var areaWidth = boundsWidth - fontSize;
        if (areaWidth < fontSize) {
            areaWidth = Math.max(boundsWidth, fontSize);
        }

        /* 長方形を作成してエリア内文字に変換 / Create a rectangle and convert it to area text */
        var areaRect = doc.pathItems.rectangle(boundsTop, boundsLeft, areaWidth, areaHeight);
        areaRect.stroked = false;
        areaRect.filled = false;

        var leading = computeLeading(mergedLineFrames, fontSize);
        var areaTextFrame = doc.textFrames.areaText(areaRect);
        areaTextFrame.contents = joinedText;
        applyAreaTextFormatting(areaTextFrame, mergedLineFrames[0], fontSize, leading);

        /* 枠の高さは元の外接矩形どおりなので、最終行があふれる分だけ下に伸ばす
           The frame keeps the original bounding box height, so extend it downward until the last line fits */
        growFrameUntilTextFits(areaTextFrame, (leading !== null) ? leading : fontSize * MIN_LEADING_RATIO);
        return areaTextFrame;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択テキストを行ごとにタブ区切りで連結し、エリア内文字を生成する
     * @returns {void}
     */
    function main() {
        /* ドキュメント未オープン時は終了 / Exit when no document is open */
        if (app.documents.length === 0) {
            return;
        }
        var doc = app.activeDocument;

        var lineGroups = groupTextFramesByLine(doc.selection);
        if (lineGroups.length === 0) {
            alert(getLabel(LABELS.alert.noText));
            return;
        }

        /* 1行だけの場合は別処理：エリア内文字にせず左揃えで出力
           Single line: merge and left-align only, without converting to area text */
        if (lineGroups.length === 1) {
            var singleLineFrame = mergeLineFrames(lineGroups[0]);
            singleLineFrame.textRange.paragraphAttributes.justification = Justification.LEFT;
            doc.selection = [singleLineFrame];
            return;
        }

        /* 各行を1つのテキストフレームに連結 / Merge each line into a single text frame */
        var mergedLineFrames = [];
        for (var i = 0; i < lineGroups.length; i++) {
            mergedLineFrames.push(mergeLineFrames(lineGroups[i]));
        }

        var areaTextFrame = createAreaTextFrame(mergedLineFrames, joinLineContents(mergedLineFrames));

        /* 連結に使った元のフレームを削除 / Delete the merged source frames */
        removeFrames(mergedLineFrames);

        /* 生成されたエリア内文字を選択状態にする / Select the generated area text */
        doc.selection = null;
        doc.selection = [areaTextFrame];
        app.redraw();
    }

    main();

})();
