#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストフレームから、Illustrator標準の「箇条書きと番号付きリスト」を解除します。
テキストの内容・文字属性・段落設定・タブストップは控えて戻すため、リスト書式だけが外れます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ClearBulletsAndNumbering.md

### Overview

Removes Illustrator's built-in Bullets and Numbering from the selected text frames.
The text content, character attributes, paragraph settings and tab stops are captured and restored, so only the list formatting comes off.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ClearBulletsAndNumbering.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ClearBulletsAndNumbering";     /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-08-18";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ClearBulletsAndNumbering.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ClearBulletsAndNumbering.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

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

    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document open." },
            noSelection: { ja: "テキストオブジェクトを選択してください。", en: "Please select text objects." },
            noTextFrame: { ja: "テキストフレームを選択してください。", en: "Please select text frames." }
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
    // メイン処理 / Main
    // =========================================

    main();

    /**
     * 前提チェックののち、選択したテキストフレームのリスト書式を解除する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) { alert(getLabel("alert.noDocument")); return; }
        var selectedItems = app.activeDocument.selection;
        if (selectedItems.length === 0) { alert(getLabel("alert.noSelection")); return; }

        /* 文字カーソルの選択はフレームに読み替えない（従来どおり） / Do not resolve a text-cursor selection to its frame (as before) */
        var targetFrames = collectSelectionTextFrames(selectedItems, { textRangeToFrame: false });
        if (targetFrames.length === 0) {
            alert(getLabel("alert.noTextFrame"));
            return;
        }

        for (var i = 0; i < targetFrames.length; i++) {
            clearListFormatting(targetFrames[i]);
        }
        app.redraw();
    }

    /**
     * 1フレームの「箇条書きと番号付きリスト」を解除する
     * contents を入れ直すとリスト書式が外れる（同時に文字書式も初期化されるため、控えてから戻す）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {void}
     */
    function clearListFormatting(textFrame) {
        var frameState = captureFrameState(textFrame);
        if (frameState.contents == null) return;

        try {
            textFrame.contents = frameState.contents;
        } catch (e) {
            return; /* 入れ直せなければ書式も戻さない / leave the frame untouched when the text cannot be reassigned */
        }

        restoreFrameState(textFrame, frameState);
    }

    // =========================================
    // 書式の退避・復元 / Format snapshot & restore
    // =========================================
    // contents の再設定でフレーム全体の書式が初期化されるため、文字属性・段落属性を控えて復元する
    // Setting .contents resets the frame's formatting, so character and paragraph attributes are snapshotted and restored.

    /**
     * テキストと書式の現状を控える
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {{contents: (string|null), charSnapshots: Object[], paraSnapshots: Object[]}} 控えた状態
     */
    function captureFrameState(textFrame) {
        var frameState = { contents: null, charSnapshots: [], paraSnapshots: [] };
        try { frameState.contents = textFrame.contents; } catch (e) { return frameState; }
        frameState.charSnapshots = captureCharAttributes(textFrame);
        frameState.paraSnapshots = captureParagraphFormats(textFrame);
        return frameState;
    }

    /**
     * 控えておいた書式をフレームへ復元する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {{charSnapshots: Object[], paraSnapshots: Object[]}} frameState - captureFrameState() が返した控え
     * @returns {void}
     */
    function restoreFrameState(textFrame, frameState) {
        restoreAllCharAttributes(textFrame, frameState.charSnapshots);
        restoreParagraphFormats(textFrame, frameState.paraSnapshots);
    }

    /**
     * フレーム内の全文字の文字属性を控える
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {Object[]} 文字ごとの属性（文字順）
     */
    function captureCharAttributes(textFrame) {
        var charSnapshots = [];
        try {
            var characters = textFrame.textRange.characters;
            for (var i = 0; i < characters.length; i++) {
                charSnapshots.push(snapshotCharAttributes(characters[i].characterAttributes));
            }
        } catch (e) { }
        return charSnapshots;
    }

    /**
     * 控えた文字属性を全文字へ復元する（文字数は不変なので先頭から順に対応づける）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {Object[]} charSnapshots - captureCharAttributes() が返した控え
     * @returns {void}
     */
    function restoreAllCharAttributes(textFrame, charSnapshots) {
        if (!charSnapshots || charSnapshots.length === 0) return;
        try {
            var characters = textFrame.textRange.characters;
            for (var i = 0; i < characters.length && i < charSnapshots.length; i++) {
                restoreCharAttributes(characters[i].characterAttributes, charSnapshots[i]);
            }
        } catch (e) { }
    }

    /**
     * 1文字分の主要な文字属性を控える
     * @param {CharacterAttributes} characterAttr - 対象の文字属性
     * @returns {Object} 控えた属性
     */
    function snapshotCharAttributes(characterAttr) {
        var charSnapshot = {};
        /* 途中で失敗しても、それまでに読めた属性は残る / attributes read before a failure are kept */
        try {
            charSnapshot.textFont = characterAttr.textFont;
            charSnapshot.size = characterAttr.size;
            charSnapshot.horizontalScale = characterAttr.horizontalScale;
            charSnapshot.verticalScale = characterAttr.verticalScale;
            charSnapshot.baselineShift = characterAttr.baselineShift;
            charSnapshot.tracking = characterAttr.tracking;
            charSnapshot.leading = characterAttr.leading;
            charSnapshot.autoLeading = characterAttr.autoLeading;
            charSnapshot.fillColor = characterAttr.fillColor;
        } catch (e) { }
        return charSnapshot;
    }

    /**
     * 控えた文字属性を1文字へ復元する
     * @param {CharacterAttributes} characterAttr - 復元先の文字属性
     * @param {Object} charSnapshot - snapshotCharAttributes() が返した控え
     * @returns {void}
     */
    function restoreCharAttributes(characterAttr, charSnapshot) {
        if (!charSnapshot) return;
        /* フォントは失敗しやすいので分けて囲み、他の属性の復元を巻き込まない
           Guard the font separately so a failure there does not skip the remaining attributes */
        if (charSnapshot.textFont) { try { characterAttr.textFont = charSnapshot.textFont; } catch (eFont) { } }
        try {
            if (charSnapshot.size != null) characterAttr.size = charSnapshot.size;
            if (charSnapshot.horizontalScale != null) characterAttr.horizontalScale = charSnapshot.horizontalScale;
            if (charSnapshot.verticalScale != null) characterAttr.verticalScale = charSnapshot.verticalScale;
            if (charSnapshot.baselineShift != null) characterAttr.baselineShift = charSnapshot.baselineShift;
            if (charSnapshot.tracking != null) characterAttr.tracking = charSnapshot.tracking;
            /* 行送りは自動行送りより先に戻す（先に autoLeading を立てると固定値が入らない）
               Restore leading before auto-leading (setting auto-leading first would drop the fixed value) */
            if (charSnapshot.leading != null) characterAttr.leading = charSnapshot.leading;
            if (charSnapshot.autoLeading != null) characterAttr.autoLeading = charSnapshot.autoLeading;
            if (charSnapshot.fillColor) characterAttr.fillColor = charSnapshot.fillColor;
        } catch (e) { }
    }

    /**
     * 各段落の段落属性を控える
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {Object[]} 段落ごとの属性（段落順）
     */
    function captureParagraphFormats(textFrame) {
        var paraSnapshots = [];
        try {
            var paragraphs = textFrame.paragraphs;
            for (var i = 0; i < paragraphs.length; i++) {
                paraSnapshots.push(snapshotParagraphAttributes(paragraphs[i].paragraphAttributes));
            }
        } catch (e) { }
        return paraSnapshots;
    }

    /**
     * 控えた段落属性を各段落へ復元する（段落数は不変なので先頭から順に対応づける）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {Object[]} paraSnapshots - captureParagraphFormats() が返した控え
     * @returns {void}
     */
    function restoreParagraphFormats(textFrame, paraSnapshots) {
        if (!paraSnapshots || paraSnapshots.length === 0) return;
        try {
            var paragraphs = textFrame.paragraphs;
            for (var i = 0; i < paragraphs.length && i < paraSnapshots.length; i++) {
                restoreParagraphAttributes(paragraphs[i].paragraphAttributes, paraSnapshots[i]);
            }
        } catch (e) { }
    }

    /**
     * 1段落分の主要な段落属性を控える
     * @param {ParagraphAttributes} paragraphAttr - 対象の段落属性
     * @returns {Object} 控えた属性
     */
    function snapshotParagraphAttributes(paragraphAttr) {
        var paraSnapshot = {};
        try {
            paraSnapshot.justification = paragraphAttr.justification;
            paraSnapshot.spaceBefore = paragraphAttr.spaceBefore;
            paraSnapshot.spaceAfter = paragraphAttr.spaceAfter;
            paraSnapshot.leftIndent = paragraphAttr.leftIndent;
            paraSnapshot.rightIndent = paragraphAttr.rightIndent;
            paraSnapshot.firstLineIndent = paragraphAttr.firstLineIndent;
            paraSnapshot.tabStops = copyTabStops(paragraphAttr.tabStops);
        } catch (e) { }
        return paraSnapshot;
    }

    /**
     * 控えた段落属性を1段落へ復元する
     * @param {ParagraphAttributes} paragraphAttr - 復元先の段落属性
     * @param {Object} paraSnapshot - snapshotParagraphAttributes() が返した控え
     * @returns {void}
     */
    function restoreParagraphAttributes(paragraphAttr, paraSnapshot) {
        if (!paraSnapshot) return;
        try {
            if (paraSnapshot.justification != null) paragraphAttr.justification = paraSnapshot.justification;
            if (paraSnapshot.spaceBefore != null) paragraphAttr.spaceBefore = paraSnapshot.spaceBefore;
            if (paraSnapshot.spaceAfter != null) paragraphAttr.spaceAfter = paraSnapshot.spaceAfter;
            if (paraSnapshot.leftIndent != null) paragraphAttr.leftIndent = paraSnapshot.leftIndent;
            if (paraSnapshot.rightIndent != null) paragraphAttr.rightIndent = paraSnapshot.rightIndent;
            if (paraSnapshot.firstLineIndent != null) paragraphAttr.firstLineIndent = paraSnapshot.firstLineIndent;
        } catch (e) { }
        /* タブストップは TabStopInfo を作り直して差し替える / rebuild TabStopInfo objects for the tab stops */
        if (paraSnapshot.tabStops) {
            try { paragraphAttr.tabStops = makeTabStops(paraSnapshot.tabStops); } catch (eTab) { }
        }
    }

    /**
     * タブストップを位置と揃えだけの配列として控える
     * @param {TabStopInfo[]} tabStops - 対象のタブストップ
     * @returns {Array<{position: number, alignment: TabStopAlignment}>|null} 控えた内容（取得できなければ null）
     */
    function copyTabStops(tabStops) {
        var tabSpecs = [];
        try {
            for (var i = 0; i < tabStops.length; i++) {
                tabSpecs.push({ position: tabStops[i].position, alignment: tabStops[i].alignment });
            }
        } catch (e) {
            return null;
        }
        return tabSpecs;
    }

    /**
     * 控えた内容から TabStopInfo の配列を作る
     * @param {Array<{position: number, alignment: TabStopAlignment}>} tabSpecs - 控えたタブストップ
     * @returns {TabStopInfo[]} 生成したタブストップ
     */
    function makeTabStops(tabSpecs) {
        var tabStopInfos = [];
        for (var i = 0; i < tabSpecs.length; i++) {
            var tabStopInfo = new TabStopInfo();
            tabStopInfo.alignment = tabSpecs[i].alignment;
            tabStopInfo.position = tabSpecs[i].position;
            tabStopInfos.push(tabStopInfo);
        }
        return tabStopInfos;
    }
})();
