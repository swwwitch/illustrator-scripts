#target illustrator
#targetengine "ReplaceWithPasteEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

クリップボードのテキストで、選択中のテキストフレームの内容をまとめて置き換えます。
クリップボードがテキスト以外なら、選択したオブジェクトをその内容で置き換えます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReplaceWithPaste.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nf14ce08eb618

### Overview

Replaces the contents of the selected text frames with the text on the clipboard.
When the clipboard holds something other than text, it replaces the selected objects with it.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ReplaceWithPaste.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ReplaceWithPaste";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.0.6";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2024-10-28";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReplaceWithPaste.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ReplaceWithPaste.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nf14ce08eb618"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* テキスト以外で置き換えるとき、大きさの扱いを選ぶダイアログを開くか
       Whether to open a dialog to choose the sizing when replacing with non-text contents */
    var SHOW_SIZE_DIALOG = true;

    /* ダイアログを開かないときの大きさの扱い（開くときは初期選択）
       "keep"：大きさ保持、"long"：長辺に合わせる、"short"：短辺に合わせる
       Sizing used without the dialog (its initial choice otherwise)
       "keep": keep the pasted size, "long": fit the long side, "short": fit the short side */
    var DEFAULT_SIZE_MODE = "long";

    // =========================================
    // レイアウト / Layout
    // =========================================

    var WINDOW_MARGINS        = 16;               /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING        = 12;               /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS         = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING         = 6;                /* パネル内の要素間隔 / panel spacing */

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

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

    var LABELS = {
        dialog: {
            title: { ja: "クリップボードで置き換え", en: "Replace with Clipboard" }
        },
        panel: {
            size: { ja: "大きさ", en: "Size" }
        },
        radio: {
            keepSize: { ja: "大きさ保持", en: "Keep Size" },
            fitLongSide: { ja: "長辺に合わせる", en: "Fit Long Side" },
            fitShortSide: { ja: "短辺に合わせる", en: "Fit Short Side" }
        },
        tooltip: {
            keepSize: {
                ja: "拡大・縮小せず、中心だけを元のオブジェクトにそろえます",
                en: "Centers on the original object without scaling"
            },
            fitLongSide: {
                ja: "長辺どうしが同じ長さになるよう、縦横比を保ったまま拡大・縮小します",
                en: "Scales proportionally so the long sides match"
            },
            fitShortSide: {
                ja: "短辺どうしが同じ長さになるよう、縦横比を保ったまま拡大・縮小します",
                en: "Scales proportionally so the short sides match"
            }
        },
        checkbox: {
            preview: { ja: "プレビューを表示", en: "Show Preview" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            emptyClipboard: {
                ja: "クリップボードが空か、Illustrator に貼り付けられない内容です。",
                en: "The clipboard is empty, or Illustrator cannot paste its contents."
            },
            noTextInClipboard: {
                ja: "クリップボードにテキストが見つかりませんでした。",
                en: "No text was found on the clipboard."
            },
            clipboardError: {
                ja: "クリップボードからの取得に失敗しました：\n",
                en: "Failed to get text from clipboard:\n"
            },
            textCreateError: {
                ja: "新規テキスト作成時にエラーが発生しました：\n",
                en: "An error occurred while creating new text:\n"
            },
            replaceError: {
                ja: "テキスト置換中にエラーが発生しました：\n",
                en: "An error occurred while replacing text:\n"
            },
            objectReplaceError: {
                ja: "オブジェクトの置き換え中にエラーが発生しました：\n",
                en: "An error occurred while replacing objects:\n"
            }
        }
    };

    // =========================================
    // 選択の操作 / Selection helpers
    // =========================================

    /**
     * 選択オブジェクトを配列に写し取る（selection は操作で変化するため）。
     * テキスト編集中は selection が配列ではなく TextRange そのものになることがあり、
     * その length は選択した文字数を指す。添字で取り出すと undefined が並び、
     * あとで選択に戻すときに Illustrator が落ちるため、形を判定してから写す。
     * @param {Document} doc - 対象ドキュメント
     * @returns {Object[]} 選択オブジェクトの配列。選択がなければ空配列
     */
    function captureSelection(doc) {
        var capturedItems = [];
        var currentSelection = doc.selection;
        if (!currentSelection) return capturedItems;

        if (!(currentSelection instanceof Array)) {
            if (currentSelection.typename === "TextRange") capturedItems.push(currentSelection);
            return capturedItems;
        }

        for (var i = 0; i < currentSelection.length; i++) {
            if (currentSelection[i]) capturedItems.push(currentSelection[i]);
        }
        return capturedItems;
    }

    /**
     * 選択状態を差し替える（失敗しても処理を止めない）
     * @param {Document} doc - 対象ドキュメント
     * @param {Object[]} itemsToSelect - 選択するオブジェクトの配列。空または null で選択解除
     * @returns {void}
     */
    function setSelection(doc, itemsToSelect) {
        try {
            doc.selection = (itemsToSelect && itemsToSelect.length > 0) ? itemsToSelect : null;
        } catch (e) {
            /* Illustrator が選択を拒む場合は現在の選択のままにする / Keep whatever stays selected */
        }
    }

    /**
     * テキスト編集中の選択（TextRange）を、あとから使える数値として控える。
     * オブジェクト参照は編集モードの解除やペーストで無効になり得るため、
     * 親テキストフレームと文字位置だけを持たせる。
     * @param {Object[]} capturedItems - 退避した選択
     * @returns {{frame: TextFrame, start: number, end: number}|null} 文字範囲の情報。編集中でなければ null
     */
    function captureEditingRange(capturedItems) {
        if (!capturedItems || capturedItems.length !== 1) return null;

        var textRange = capturedItems[0];
        if (!textRange || textRange.typename !== "TextRange") return null;

        try {
            var parentFrame = resolveTextRangeFrame(textRange);
            if (!parentFrame) return null;

            return { frame: parentFrame, start: textRange.start, end: textRange.end };
        } catch (e) {
            return null;
        }
    }

    /**
     * テキスト編集モードを抜ける。
     * 編集中のままペーストするとクリップボードの内容が文字として流し込まれ、
     * さらに再描画を挟むと Illustrator が不安定になるため、読み取りの前に必ず抜ける。
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function leaveTextEditing(doc) {
        try {
            /* 選択ツールへ切り替えると編集が確定して抜けられる / Switching tools commits the edit and leaves it */
            app.selectTool("Adobe Select Tool");
        } catch (e) {
            /* 切り替えられない場合は次の選択解除に任せる / Leave it to the deselect below */
        }
        setSelection(doc, null);
    }

    /**
     * 指定のオブジェクトをまとめて削除する
     * @param {Object[]} itemsToRemove - 削除対象の配列
     * @returns {void}
     */
    function removeItems(itemsToRemove) {
        if (!itemsToRemove || !itemsToRemove.length) return;
        for (var i = itemsToRemove.length - 1; i >= 0; i--) {
            try {
                /* TextRange の remove() は文字そのものを消すため対象外にする / TextRange.remove() would delete characters, so skip it */
                if (itemsToRemove[i].typename === "TextRange") continue;
                itemsToRemove[i].remove();
            } catch (e) {
                /* 既に消えているものは無視する / Ignore items that are already gone */
            }
        }
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

    /**
     * 配列から最初のテキストフレームを探す。
     * 他アプリからのペーストはグループやクリップグループにまとめられることがあるため、中も再帰的にたどる。
     * @param {Object[]} searchItems - 探索対象の配列またはコレクション
     * @returns {TextFrame|null} 見つかったテキストフレーム。なければ null
     */
    function findFirstTextFrame(searchItems) {
        var foundFrames = collectSelectionTextFrames(searchItems, { unique: false });
        return (foundFrames.length > 0) ? foundFrames[0] : null;
    }

    // =========================================
    // クリップボードの取得 / Clipboard access
    // =========================================

    /**
     * 選択を解除してからペーストし、貼り付いたオブジェクトを返す。
     * 先に解除するのは、ペーストが実行されなかったときに元の選択を
     * 「貼り付いたもの」と取り違えないため。
     * @param {Document} doc - 対象ドキュメント
     * @returns {Object[]} 貼り付いたオブジェクトの配列。貼り付かなかった場合は空配列
     */
    function pasteAndCapture(doc) {
        setSelection(doc, null);
        app.paste();
        /* 貼り付け直後は selection に反映されないことがあるため、描画を確定させてから読む / Flush the paste before reading the selection */
        app.redraw();
        return captureSelection(doc);
    }

    /**
     * クリップボードの内容をドキュメントへ貼り付け、貼り付いたオブジェクトを返す。
     * Illustrator は自分がコピーした内容を内部に保持していて、他アプリがクリップボードを
     * 書き換えたあとの1回目のペーストでは古い内容が貼り付く。その1回目が内部の更新を促すため、
     * 1回目は捨てて2回目の結果を使う。
     * @param {Document} doc - 対象ドキュメント
     * @returns {Object[]} 貼り付いたオブジェクトの配列。貼り付かなかった場合は空配列
     */
    function pasteClipboardItems(doc) {
        try {
            /* 1回目は内部クリップボードを最新にするためだけのペースト / The first paste only refreshes Illustrator's cached clipboard */
            removeItems(pasteAndCapture(doc));
        } catch (e) {
            /* 更新目的なので、失敗しても2回目の結果で判断する / Judge by the second paste even if this one fails */
        }
        return pasteAndCapture(doc);
    }

    /**
     * 一度ペーストして、クリップボードの中身がテキストかそれ以外かを調べる。
     * テキストなら貼り付いたテキストフレームから内容と座標を読み取る。
     * 読み取り後は貼り付けたオブジェクトを削除し、元の選択へ戻す。
     * @param {Document} doc - 対象ドキュメント
     * @param {Object[]} originalSelection - 復元する元の選択
     * @returns {{kind: string, bounds: number[], contents: string}|null} kind は "text" または "objects"（objects のときは bounds と contents を持たない）。貼り付けに失敗した場合は null
     */
    function readClipboard(doc, originalSelection) {
        var clipboardInfo = null;
        var pastedItems = null;
        var pasteError = null;

        try {
            pastedItems = pasteClipboardItems(doc);

            var pastedTextFrame = findFirstTextFrame(pastedItems);
            if (pastedTextFrame) {
                clipboardInfo = {
                    kind: "text",
                    bounds: pastedTextFrame.geometricBounds,
                    contents: pastedTextFrame.contents
                };
            } else if (pastedItems.length > 0) {
                clipboardInfo = { kind: "objects" };
            }
        } catch (e) {
            pasteError = String(e);
        }

        /* 成否にかかわらず、貼り付けた分を消して元の選択へ戻す / Clean up and restore regardless of the outcome */
        removeItems(pastedItems);
        setSelection(doc, originalSelection);

        /* 画面を元に戻してから知らせる / Report only after the canvas is back to its original state */
        if (pasteError) {
            alert(getLabel("alert.clipboardError") + pasteError);
            return null;
        }
        if (!clipboardInfo) alert(getLabel("alert.emptyClipboard"));
        return clipboardInfo;
    }

    // =========================================
    // テキストの適用 / Text application
    // =========================================

    /**
     * 同じ内容のエラーを重複させずに追加する
     * @param {string[]} errorMessages - 収集先
     * @param {string} errorMessage - 追加するエラーメッセージ
     * @returns {void}
     */
    function addUniqueError(errorMessages, errorMessage) {
        for (var i = 0; i < errorMessages.length; i++) {
            if (errorMessages[i] === errorMessage) return;
        }
        errorMessages.push(errorMessage);
    }

    /**
     * 置き換えで起きたエラーを1回の alert でまとめて知らせる（選択数だけダイアログが出ないように）
     * @param {string[]} errorMessages - 収集したエラーメッセージ
     * @param {string} [labelPath] - 見出しに使うラベルのキー。省略時は "alert.replaceError"
     * @returns {void}
     */
    function alertReplaceErrors(errorMessages, labelPath) {
        if (errorMessages.length === 0) return;
        alert(getLabel(labelPath || "alert.replaceError") + errorMessages.join("\n"));
    }

    /**
     * クリップボードのテキストを、貼り付いた位置に新規テキストフレームとして作成する
     * @param {Layer} targetLayer - 作成先のレイヤー
     * @param {number[]} pastedBounds - 貼り付いたテキストフレームの geometricBounds
     * @param {string} textContent - 作成するテキスト
     * @returns {void}
     */
    function createNewTextFrame(targetLayer, pastedBounds, textContent) {
        try {
            var newFrame = targetLayer.textFrames.add();
            newFrame.contents = textContent;

            var newFrameBounds = newFrame.geometricBounds;
            newFrame.translate(pastedBounds[0] - newFrameBounds[0], pastedBounds[1] - newFrameBounds[1]);
        } catch (e) {
            alert(getLabel("alert.textCreateError") + e);
        }
    }

    /**
     * 編集中に選択していた文字範囲だけを、クリップボードのテキストで置き換える。
     * 編集モードを抜けてから読み取り、終わったら対象のテキストフレームを選択しておく。
     * @param {Document} doc - 対象ドキュメント
     * @param {{frame: TextFrame, start: number, end: number}} rangeInfo - 置き換える範囲
     * @returns {void}
     */
    function replaceSelectedCharacters(doc, rangeInfo) {
        leaveTextEditing(doc);

        var clipboardInfo = readClipboard(doc, []);
        if (!clipboardInfo) return;

        var alertMessage = null;
        if (clipboardInfo.kind !== "text") {
            /* 文字の中にはテキスト以外を流し込めない / Only text can go into a character range */
            alertMessage = getLabel("alert.noTextInClipboard");
        } else {
            try {
                var targetRange = rangeInfo.frame.textRange;
                targetRange.start = rangeInfo.start;
                targetRange.end = rangeInfo.end;
                targetRange.contents = clipboardInfo.contents;
            } catch (e) {
                alertMessage = getLabel("alert.replaceError") + e;
            }
        }

        /* 編集モードは抜けているので、対象のテキストフレームを選択して終える / Editing is over, so leave the frame itself selected */
        setSelection(doc, [rangeInfo.frame]);
        app.redraw();
        if (alertMessage) alert(alertMessage);
    }

    /**
     * 集めたテキストフレームの内容をまとめて置き換える
     * @param {TextFrame[]} targetFrames - 対象のテキストフレーム
     * @param {string} textContent - 適用するテキスト
     * @param {string[]} errorMessages - 発生したエラーの収集先（呼び出し元でまとめて通知する）
     * @returns {void}
     */
    function applyTextToFrames(targetFrames, textContent, errorMessages) {
        for (var i = 0; i < targetFrames.length; i++) {
            try {
                targetFrames[i].contents = textContent;
            } catch (e) {
                addUniqueError(errorMessages, String(e));
            }
        }
    }

    // =========================================
    // オブジェクトの置き換え / Object replacement
    // =========================================

    /**
     * 境界（visibleBounds。クリップグループはマスクの範囲）から中心と長辺・短辺を求める
     * @param {number[]} bounds - [左, 上, 右, 下]
     * @returns {{centerX: number, centerY: number, longSide: number, shortSide: number}} 中心座標と長辺・短辺の長さ
     */
    function measureBounds(bounds) {
        var width = bounds[2] - bounds[0];
        var height = bounds[1] - bounds[3];
        return {
            centerX: (bounds[0] + bounds[2]) / 2,
            centerY: (bounds[1] + bounds[3]) / 2,
            longSide: Math.max(width, height),
            shortSide: Math.min(width, height)
        };
    }

    /**
     * 大きさの扱いに応じた拡大・縮小率を求める
     * @param {{longSide: number, shortSide: number}} pastedMetrics - 貼り付けた内容の長辺・短辺
     * @param {{longSide: number, shortSide: number}} targetMetrics - 置き換え先の長辺・短辺
     * @param {string} sizeMode - "keep"・"long"・"short" のいずれか
     * @returns {number} 拡大・縮小率（1 で等倍）
     */
    function getFitScale(pastedMetrics, targetMetrics, sizeMode) {
        if (sizeMode === "long" && pastedMetrics.longSide > 0) return targetMetrics.longSide / pastedMetrics.longSide;
        if (sizeMode === "short" && pastedMetrics.shortSide > 0) return targetMetrics.shortSide / pastedMetrics.shortSide;
        return 1;
    }

    /**
     * 貼り付けたオブジェクト全体を指定の倍率で拡大・縮小し、中心を置き換え先にそろえる。
     * 複数ある場合は互いの配置を保ったまま、ひとまとまりとして扱う
     * @param {PageItem[]} pastedItems - 貼り付けたオブジェクト
     * @param {{centerX: number, centerY: number}} targetMetrics - 置き換え先の中心
     * @param {number} scale - 今の大きさに対する倍率（1 で等倍）
     * @returns {void}
     */
    function scalePastedItems(pastedItems, targetMetrics, scale) {
        var pastedMetrics = measureBounds(getClipAwareUnionBounds(pastedItems, true));

        for (var i = 0; i < pastedItems.length; i++) {
            var pastedItem = pastedItems[i];
            var itemMetrics = measureBounds(getClipAwareBounds(pastedItem, true));

            /* 各オブジェクトの中心は、全体の中心からの距離を同じ比率で伸縮させた位置へ / Keep each item's offset from the group center, scaled */
            var destX = targetMetrics.centerX + (itemMetrics.centerX - pastedMetrics.centerX) * scale;
            var destY = targetMetrics.centerY + (itemMetrics.centerY - pastedMetrics.centerY) * scale;

            if (scale !== 1) pastedItem.resize(scale * 100, scale * 100);

            var scaledMetrics = measureBounds(getClipAwareBounds(pastedItem, true));
            pastedItem.translate(destX - scaledMetrics.centerX, destY - scaledMetrics.centerY);
        }
    }

    /**
     * 選択した各オブジェクトの直前（前面）に、クリップボードの内容を貼り付ける。
     * 元のオブジェクトはまだ消さず、確定するかキャンセルするかをあとで決められるようにする
     * @param {Document} doc - 対象ドキュメント
     * @param {Object[]} targetItems - 置き換えるオブジェクト
     * @param {string[]} errorMessages - 発生したエラーの収集先（呼び出し元でまとめて通知する）
     * @returns {Object[]} 置き換えの組（targetItem・pastedItems・targetMetrics・baseMetrics・currentScale）の配列
     */
    function pasteOverTargets(doc, targetItems, errorMessages) {
        var replacements = [];
        for (var i = targetItems.length - 1; i >= 0; i--) {
            try {
                var targetItem = targetItems[i];
                var targetMetrics = measureBounds(getClipAwareBounds(targetItem, true));

                var pastedItems = pasteAndCapture(doc);
                if (pastedItems.length === 0) continue;

                for (var j = 0; j < pastedItems.length; j++) {
                    try {
                        pastedItems[j].move(targetItem, ElementPlacement.PLACEBEFORE);
                    } catch (e) {
                        /* 移せない場合はペースト先のレイヤーに残す / Leave it on the paste layer if it cannot be moved */
                    }
                }

                replacements.push({
                    targetItem: targetItem,
                    pastedItems: pastedItems,
                    targetMetrics: targetMetrics,
                    baseMetrics: measureBounds(getClipAwareUnionBounds(pastedItems, true)),
                    currentScale: 1
                });
            } catch (e) {
                /* ロック中のオブジェクトなど、DOM が操作を拒む場合 / The DOM may refuse, e.g. for locked objects */
                addUniqueError(errorMessages, String(e));
            }
        }
        setSelection(doc, null);
        return replacements;
    }

    /**
     * 貼り付けた内容を、指定の扱いの大きさにする。
     * 倍率は貼り付けた時点の大きさから求め、今の倍率との差だけ拡大・縮小する（切り替えを繰り返しても誤差をためない）
     * @param {Object[]} replacements - pasteOverTargets() が返した置き換えの組
     * @param {string} sizeMode - "keep"（大きさ保持）・"long"（長辺に合わせる）・"short"（短辺に合わせる）
     * @returns {void}
     */
    function applySizeMode(replacements, sizeMode) {
        for (var i = 0; i < replacements.length; i++) {
            var replacement = replacements[i];
            var desiredScale = getFitScale(replacement.baseMetrics, replacement.targetMetrics, sizeMode);
            scalePastedItems(replacement.pastedItems, replacement.targetMetrics, desiredScale / replacement.currentScale);
            replacement.currentScale = desiredScale;
        }
    }

    /**
     * プレビューの表示を切り替える。表示中は元のオブジェクトを隠し、貼り付けた内容を見せる
     * @param {Object[]} replacements - 置き換えの組
     * @param {boolean} showResult - true で置き換え後、false で置き換え前を表示
     * @returns {void}
     */
    function setPreviewVisible(replacements, showResult) {
        for (var i = 0; i < replacements.length; i++) {
            replacements[i].targetItem.hidden = showResult;
            for (var j = 0; j < replacements[i].pastedItems.length; j++) {
                replacements[i].pastedItems[j].hidden = !showResult;
            }
        }
        app.redraw();
    }

    /**
     * 置き換えを確定する。貼り付けた内容を表示し、元のオブジェクトを削除する
     * @param {Object[]} replacements - 置き換えの組
     * @returns {PageItem[]} 貼り付けたオブジェクトすべて
     */
    function commitReplacements(replacements) {
        var replacedItems = [];
        for (var i = 0; i < replacements.length; i++) {
            for (var j = 0; j < replacements[i].pastedItems.length; j++) {
                replacements[i].pastedItems[j].hidden = false;
                replacedItems.push(replacements[i].pastedItems[j]);
            }
            replacements[i].targetItem.remove();
        }
        return replacedItems;
    }

    /**
     * 置き換えを取りやめる。貼り付けた内容を削除し、元のオブジェクトを表示に戻す
     * @param {Object[]} replacements - 置き換えの組
     * @returns {void}
     */
    function discardReplacements(replacements) {
        for (var i = 0; i < replacements.length; i++) {
            removeItems(replacements[i].pastedItems);
            replacements[i].targetItem.hidden = false;
        }
    }

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

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */

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
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
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
     * 左のグループにボタンが無い（右のボタンだけの）とき、行を左右中央に並べ直す。
     * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
     * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
     * @returns {void}
     */
    function centerButtonRowIfRightOnly(buttonRow) {
        if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
        var btnRowGroup = buttonRow.rowGroup;
        /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
        btnRowGroup.remove(buttonRow.leftGroup);
        btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
        btnRowGroup.alignment = ["center", "bottom"];
        btnRowGroup.alignChildren = ["center", "center"];
        buttonRow.leftGroup = null;
    }

    // ボタン行（再利用パーツ）ここまで / End of the reusable button row

    /**
     * 大きさの扱いを選ぶダイアログを開く。選択を切り替えるたびにプレビューへ反映する
     * @param {string} initialMode - 初期選択。"keep"・"long"・"short" のいずれか
     * @param {Object[]} replacements - プレビューに使う置き換えの組
     * @returns {string|null} 選んだ扱い（"keep"・"long"・"short"）。キャンセルなら null
     */
    function showSizeDialog(initialMode, replacements) {
        var sizeDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        sizeDialog.orientation = "column";
        sizeDialog.alignChildren = ["fill", "top"];
        sizeDialog.margins = WINDOW_MARGINS;
        sizeDialog.spacing = WINDOW_SPACING;

        var sizePanel = sizeDialog.add("panel", undefined, getLabel("panel.size"));
        sizePanel.orientation = "column";
        sizePanel.alignChildren = ["left", "top"];
        sizePanel.margins = PANEL_MARGINS;
        sizePanel.spacing = PANEL_SPACING;

        var sizeRadios = {
            keep: sizePanel.add("radiobutton", undefined, getLabel("radio.keepSize")),
            "long": sizePanel.add("radiobutton", undefined, getLabel("radio.fitLongSide")),
            "short": sizePanel.add("radiobutton", undefined, getLabel("radio.fitShortSide"))
        };
        sizeRadios.keep.helpTip = getLabel("tooltip.keepSize");
        sizeRadios["long"].helpTip = getLabel("tooltip.fitLongSide");
        sizeRadios["short"].helpTip = getLabel("tooltip.fitShortSide");

        var selectedMode = sizeRadios[initialMode] ? initialMode : "long";
        sizeRadios[selectedMode].value = true;

        var previewCheckbox = sizeDialog.add("checkbox", undefined, getLabel("checkbox.preview"));
        previewCheckbox.alignment = "center";
        previewCheckbox.value = true;

        var buttonRow = addButtonRow(sizeDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        /**
         * ラジオボタンのクリックで大きさを切り替える関数を作る
         * @param {string} sizeMode - 切り替え先の扱い
         * @returns {Function} onClick に渡す関数
         */
        function createSizeClickHandler(sizeMode) {
            return function () {
                selectedMode = sizeMode;
                applySizeMode(replacements, sizeMode);
                app.redraw();
            };
        }
        for (var radioMode in sizeRadios) {
            sizeRadios[radioMode].onClick = createSizeClickHandler(radioMode);
        }
        previewCheckbox.onClick = function () {
            setPreviewVisible(replacements, previewCheckbox.value);
        };

        /* 表示前のチェックボックスは値を読み戻せないため、初期状態は直接渡す / A checkbox cannot be read back before show(), so pass the initial state directly */
        applySizeMode(replacements, selectedMode);
        setPreviewVisible(replacements, true);

        centerButtonRowIfRightOnly(buttonRow);
        prepareDialogWindow(sizeDialog, SCRIPT_NAME);
        return (sizeDialog.show() === 1) ? selectedMode : null;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択した各オブジェクトを、クリップボードの内容（テキスト以外）で置き換える。
     * 選択がなければ通常のペーストと同じく画面の中央へ貼り付ける。
     * @param {Document} doc - 対象ドキュメント
     * @param {Object[]} originalSelection - 置き換えるオブジェクト
     * @returns {void}
     */
    function replaceObjects(doc, originalSelection) {
        if (originalSelection.length === 0) {
            pasteAndCapture(doc);
            return;
        }

        var errorMessages = [];
        var replacements = pasteOverTargets(doc, originalSelection, errorMessages);

        var sizeMode = DEFAULT_SIZE_MODE;
        if (SHOW_SIZE_DIALOG && replacements.length > 0) {
            sizeMode = showSizeDialog(DEFAULT_SIZE_MODE, replacements);
            if (!sizeMode) {
                discardReplacements(replacements);
                setSelection(doc, originalSelection);
                app.redraw();
                return;
            }
        }
        applySizeMode(replacements, sizeMode);

        var replacedItems = commitReplacements(replacements);
        alertReplaceErrors(errorMessages, "alert.objectReplaceError");
        setSelection(doc, replacedItems);
        app.redraw();
    }

    /**
     * クリップボードのテキストで、集めたテキストフレームの内容を置き換える。
     * 選択がなければ、貼り付いた位置に新規テキストフレームを作成する。
     * @param {Document} doc - 対象ドキュメント
     * @param {Object[]} originalSelection - 元の選択（終わったら選択し直す）
     * @param {TextFrame[]} targetFrames - 置き換えるテキストフレーム
     * @param {{bounds: number[], contents: string}} clipboardInfo - readClipboard() が返したテキストの情報
     * @returns {void}
     */
    function replaceTexts(doc, originalSelection, targetFrames, clipboardInfo) {
        if (originalSelection.length === 0) {
            createNewTextFrame(doc.activeLayer, clipboardInfo.bounds, clipboardInfo.contents);
        } else {
            var errorMessages = [];
            applyTextToFrames(targetFrames, clipboardInfo.contents, errorMessages);
            /* 選択数だけダイアログが出ないよう、まとめて1回だけ知らせる / Report every failure in a single alert */
            alertReplaceErrors(errorMessages);
        }

        /* 置換後の表示を確実に更新するため、選択を解除してから戻す / Clear and reset the selection so the redraw reflects the change */
        setSelection(doc, null);
        app.redraw();
        setSelection(doc, originalSelection);
    }

    /**
     * 選択の状態とクリップボードの中身に応じて、挿入・置き換えの処理を振り分ける
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var originalSelection = captureSelection(doc);

        var editingRange = captureEditingRange(originalSelection);
        if (editingRange) {
            if (editingRange.start === editingRange.end) {
                /* カーソルを立てただけ（文字の選択なし）なら、書式なしでカーソル位置へ挿入する / With only a caret, insert at it without formatting */
                app.executeMenuCommand("pasteWithoutFormatting");
            } else {
                /* 文字を選択して編集中なら、その範囲だけを置き換える / When characters are selected, replace just that range */
                replaceSelectedCharacters(doc, editingRange);
            }
            return;
        }

        /* グループやクリップグループの中は、ペーストを挟んで参照が古くなる前にたどっておく / Walk into groups before the paste cycle can stale the references */
        var targetFrames = collectSelectionTextFrames(originalSelection, { unique: false });

        var clipboardInfo = readClipboard(doc, originalSelection);
        if (!clipboardInfo) return;

        /* テキスト以外なら、選択したオブジェクトそのものを置き換える / For non-text contents, replace the selected objects themselves */
        if (clipboardInfo.kind === "objects") {
            replaceObjects(doc, originalSelection);
        } else {
            replaceTexts(doc, originalSelection, targetFrames, clipboardInfo);
        }
    }

    main();

})();
