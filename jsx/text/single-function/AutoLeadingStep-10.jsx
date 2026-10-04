#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストの行送り（表示値）が 整数で10ステップ小さく するように、自動行送り量（％）を逆算して設定します。
ダイアログは表示せず、実行するとその場で反映されます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AutoLeadingStep-10.md

### Overview

Works out the auto-leading percentage that decreases by ten whole steps the displayed leading of the selected text, and applies it.
There is no dialog; running it applies the change straight away.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AutoLeadingStep-10.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AutoLeadingStep-10";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AutoLeadingStep-10.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AutoLeadingStep-10.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // ステップ数はファイル名末尾の符号付き整数から取得する（例: AutoLeadingStep+10 → 10 / AutoLeadingStep-1 → -1）
    // 下記はファイル名から数値を読めなかったときの既定値
    // The step is read from the trailing signed integer of the file name (e.g. AutoLeadingStep+10 → 10);
    // this is the fallback used only when no number can be read.
    var DEFAULT_LEADING_STEP = 1;

    (function () {

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
                noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
                noSelection: { ja: "テキストが選択されていません。", en: "No text is selected." }
            }
        };

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
            return { code: unitCode, label: unit.label, pointsPerUnit: unit.pointsPerUnit };
        }

        /* 実行中スクリプトのファイル名末尾の符号付き整数をステップ数として取得（読めなければ既定値）
           Read the trailing signed integer of the running script's file name as the step (fallback to default)
           例: AutoLeadingStep+10.jsx → 10 / AutoLeadingStep-1.jsx → -1 / AutoLeadingStep+1.jsx → 1
           @param {number} defaultStep ファイル名から数値を読めないときの既定値
           @returns {number} ステップ数（整数） */
        function getScriptStep(defaultStep) {
            try {
                var baseName = new File($.fileName).name.replace(/\.[^.]*$/, "");
                var matched = baseName.match(/([+\-]?\d+)\s*$/);
                if (matched) {
                    var parsed = parseInt(matched[1], 10);
                    if (!isNaN(parsed)) return parsed;
                }
            } catch (e) { }
            return defaultStep;
        }

        /* 型名を安全に取得 / Safely resolve a type name */
        function getTypeName(obj) {
            if (obj === null || obj === undefined) return "";
            if (obj.typename) return obj.typename;
            try { return obj.constructor ? obj.constructor.name : ""; } catch (e) { return ""; }
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

        /* 範囲が触れている段落(段落全体)を対象配列へ追加 / Add the full paragraphs the range touches */
        function collectParagraphs(range, paragraphTargets) {
            try {
                var paragraphs = range.paragraphs;
                for (var i = 0; i < paragraphs.length; i++) paragraphTargets.push(paragraphs[i]);
            } catch (e) { }
        }

        /* 1段落の行送り(表示値)を整数で step だけ動かす自動行送り量(%)を逆算して適用
           Apply the auto-leading amount (%) that steps the paragraph's displayed leading by `step` integers
           @param {object} paragraph 段落範囲(Paragraph)
           @param {number} unitFactor 表示単位の pt 換算係数
           @param {number} step 行送り(表示単位・整数)のステップ数 */
        function applyLeadingStep(paragraph, unitFactor, step) {
            if (!paragraph.characters || paragraph.characters.length === 0) return;
            try {
                var charAttr = paragraph.characters[0].characterAttributes;
                var sizePt = charAttr.size;
                var leadingPt = charAttr.leading;
                if (isNaN(sizePt) || sizePt <= 0 || isNaN(leadingPt)) return;

                // 現在の行送りを表示単位に換算し、次の整数を目標にする(小さな誤差は吸収)
                // Convert the current leading to display units and target the next integer (absorb tiny float error)
                var currentInUnit = leadingPt / unitFactor;
                var base = (step >= 0) ? Math.floor(currentInUnit + 1e-4) : Math.ceil(currentInUnit - 1e-4);
                var targetInUnit = base + step;
                if (targetInUnit <= 0) return;

                // 目標の整数行送りになる自動行送り量(%)を逆算 / Back-calculate the auto-leading amount (%) for the target integer leading
                var targetLeadingPt = targetInUnit * unitFactor;
                paragraph.paragraphAttributes.autoLeadingAmount = (targetLeadingPt / sizePt) * 100;
                paragraph.characterAttributes.autoLeading = true;
            } catch (e) { }
        }

        function main() {
            if (app.documents.length === 0) {
                alert(getLabel("alert.noDocument"));
                return;
            }

            var selection = app.activeDocument.selection;
            var paragraphTargets = []; // 対象の段落範囲 / Target paragraph ranges
            var typeFrames = [];       // leadingType を設定するフレーム / Frames to set leadingType on

            if (getTypeName(selection) === "TextRange") {
                // テキスト編集モード：選択が触れている段落だけを対象(一部の文字選択でも段落全体に適用)
                // Text-edit mode: target only the paragraphs the selection touches (partial char selection → whole paragraph)
                collectParagraphs(selection, paragraphTargets);
                var editFrame = resolveTextRangeFrame(selection);
                if (editFrame) typeFrames.push(editFrame);
            } else {
                // 選択ツール：選択したフレーム(グループ内含む)の全段落を対象
                // Selection tool: target every paragraph of the selected frames (including those inside groups)
                /* 中身があり行を持つテキストフレームだけ / Only text frames that have contents and lines */
                var frames = collectSelectionItems(selection, {
                    accept: function (item) {
                        return item.typename === "TextFrame" && !!item.contents && !!item.lines && item.lines.length > 0;
                    },
                    unique: false
                });
                for (var f = 0; f < frames.length; f++) {
                    collectParagraphs(frames[f].textRange, paragraphTargets);
                    typeFrames.push(frames[f]);
                }
            }

            if (paragraphTargets.length === 0) {
                alert(getLabel("alert.noSelection"));
                return;
            }

            // ステップ数はファイル名から取得（例: AutoLeadingStep+10 → 10）/ The step comes from the file name (e.g. +10)
            var step = getScriptStep(DEFAULT_LEADING_STEP);
            // 段落ごとに行送りを整数で step 動かす / Step each paragraph's leading by `step` integers
            var unitFactor = getUnitInfo("text/units").pointsPerUnit;
            for (var p = 0; p < paragraphTargets.length; p++) applyLeadingStep(paragraphTargets[p], unitFactor, step);
            // 基準は仮想ボディの上に固定(フレーム単位) / Fix the leading basis to the top of the virtual body (per frame)
            for (var g = 0; g < typeFrames.length; g++) {
                try { typeFrames[g].textRange.leadingType = AutoLeadingType.TOPTOTOP; } catch (e) { }
            }
            app.redraw();
        }

        main();

    })();

})();
