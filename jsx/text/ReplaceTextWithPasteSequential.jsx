#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

クリップボードの複数行テキストから1行目を取り出して、選択中のテキストフレームに適用します（空行は読み飛ばします）。
適用した行はクリップボードから取り除かれるため、繰り返し実行すると次の行へ順に進みます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReplaceTextWithPasteSequential.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nf4b285b87940

### Overview

Takes the first line of the multi-line text on the clipboard and applies it to the selected text frame, skipping blank lines.
The line used is removed from the clipboard, so running it again moves on to the next one.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ReplaceTextWithPasteSequential.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ReplaceTextWithPasteSequential";  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-08-14";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReplaceTextWithPasteSequential.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ReplaceTextWithPasteSequential.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nf4b285b87940"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内のローカライズ節（LABELS の直前）に貼る。
    //    uiLang を使うコード（StepperButtons・LinkToggle の部品など）より前に置く
    // 2. 識別子は uiLang / getCurrentLang / getLabel / labelText / labelValueText / fillLabelPlaceholders。
    //    同じ役割の既存の関数・変数（getCurrentLanguage、currentLanguage、formatLabel など）は消して、これに寄せる
    // 3. 呼び出しはどちらの形でもよい（混ぜてもよい）
    //      getLabel("dialog.title")        … パス
    //      getLabel(LABELS.dialog.title)   … { ja, en } を直接
    //      getLabel("alert.count", { count: 3 })  … "{count} 個" の {count} を差し込む
    //      getLabel("alert.range", [1, 10])       … "%1〜%2" の %1・%2 を差し込む
    //      labelText("fieldLabel.width")   … 末尾にコロン（日本語は全角「：」、英語は半角「:」）
    //      labelValueText("message.count", 5) … 「件数：5」／「Count: 5」（値が続く1行。英語はコロンのあとに空白）
    // 4. 見つからないパスはパスの文字列をそのまま返す（表示で気づけるように）。{ ja, en } が無いときは空文字
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
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            emptyClipboard: {
                ja: "クリップボードが空か、Illustrator に貼り付けられない内容です。",
                en: "The clipboard is empty, or Illustrator cannot paste its contents."
            },
            noTextInClipboard: {
                ja: "クリップボードにテキストが見つかりませんでした。",
                en: "No text was found on the clipboard."
            },
            lastLineApplied: {
                ja: "最後の行を適用しました。\nクリップボードの内容はそのまま残っています。",
                en: "Applied the last line.\nThe clipboard has been left as it is."
            },
            clipboardError: {
                ja: "クリップボードからの取得に失敗しました：\n",
                en: "Failed to get text from clipboard:\n"
            },
            clipboardWriteError: {
                ja: "クリップボードの更新に失敗しました：\n",
                en: "Failed to update the clipboard:\n"
            },
            replaceError: {
                ja: "テキスト置換中にエラーが発生しました：\n",
                en: "An error occurred while replacing text:\n"
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
     * テキスト編集中の選択（TextRange）から、親のテキストフレームを取り出す。
     * オブジェクト参照は編集モードの解除やペーストで無効になり得るため、
     * フレームだけを取り出して以降の対象にする。
     * @param {Object[]} capturedItems - 退避した選択
     * @returns {TextFrame|null} 親のテキストフレーム。編集中でなければ null
     */
    function captureEditingFrame(capturedItems) {
        if (!capturedItems || capturedItems.length !== 1) return null;

        var textRange = capturedItems[0];
        if (!textRange || textRange.typename !== "TextRange") return null;

        /* 取り出せない場合は編集中として扱わない（null） / Treat it as a normal selection when we cannot tell (null) */
        return resolveTextRangeFrame(textRange);
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

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // 選択の収集と境界（再利用パーツ） / Selection items and bounds (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。使わない関数も消さずに残してよい（互いに呼び合う）。
    //    識別子は SELECTION_ITEMS_TOLERANCE / normalizeSelectionItems / resolveTextRangeFrame /
    //    collectSelectionItems / getTextFrameKindKey / collectSelectionTextFrames / collectSelectionPathItems /
    //    isClipMaskItem / getClipMaskItem / hasClippedDescendant / readUsePreviewBoundsPreference /
    //    getClipAwareBounds / filterMeasurableChildren / getClipAwareUnionBounds / isNearlySameCoordinate / areBoundsNearlyEqual
    // 2. 選択は normalizeSelectionItems(doc.selection) で配列にする。文字カーソルの選択（TextRange）は
    //    配列ではなく1個で返り、しかも .length（文字数）を持つので、length だけで配列と見なさない
    // 3. テキストフレーム:
    //      var frames = collectSelectionTextFrames(doc.selection);                           // 全種類
    //      var frames = collectSelectionTextFrames(doc.selection, { kinds: ["point", "path"] });
    //    パス:
    //      var paths = collectSelectionPathItems(doc.selection);                             // 複合パスは中のパスへ
    //      var paths = collectSelectionPathItems(doc.selection, { compoundPaths: "whole", skipClipMasks: true });
    //    それ以外は collectSelectionItems(source, { accept: function (item) { … } }) で条件を書く
    // 4. 並びは選択と同じ前面→背面（グループの中も pageItems の順）。重なり順を使う処理はこの順を前提にしてよい
    // 5. doc.selection に代入し直す配列は skipLocked / skipHidden を true にする。
    //    ロック・非表示を選択に代入すると例外になり、中の子が選択に残る
    // 6. 境界は getClipAwareBounds(item, usePreviewBounds) / getClipAwareUnionBounds(items, usePreviewBounds)。
    //    usePreviewBounds を省くと環境設定の［プレビュー境界を使用］に従う。返り値は [左, 上, 右, 下] の新しい配列
    //    （書き換えても元のオブジェクトに影響しない）。測れないときは null
    // 7. 座標の一致・前後の判定は isNearlySameCoordinate / areBoundsNearlyEqual で許容値を挟む
    //    （吸着させた辺とガイドは 1e-12 ほどずれる）
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
    // クリップボードの読み書き / Clipboard access
    // =========================================

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
            setSelection(doc, null);
            app.paste();
            app.redraw();
        } catch (e) {
            /* 更新目的なので、失敗しても2回目の結果で判断する / Judge by the second paste even if this one fails */
        }
        removeItems(captureSelection(doc));

        setSelection(doc, null);
        app.paste();
        /* 貼り付け直後は selection に反映されないことがあるため、描画を確定させてから読む / Flush the paste before reading the selection */
        app.redraw();
        return captureSelection(doc);
    }

    /**
     * ペーストの失敗（例外、または何も貼り付かなかった）を知らせる
     * @param {string|null} pasteError - ペースト中の例外の文字列。無ければ null
     * @param {Object[]|null} pastedItems - 貼り付いたオブジェクト
     * @returns {boolean} 失敗を知らせた場合は true
     */
    function alertPasteFailure(pasteError, pastedItems) {
        if (pasteError) {
            alert(getLabel("alert.clipboardError") + pasteError);
            return true;
        }
        if (!pastedItems || pastedItems.length === 0) {
            /* ペースト自体が起きなかった場合 / Nothing was pasted at all */
            alert(getLabel("alert.emptyClipboard"));
            return true;
        }
        return false;
    }

    /**
     * 空白だけの行かどうかを判定する（半角スペース、タブ、全角スペースを空白として扱う）
     * @param {string} lineText - 判定する行
     * @returns {boolean} 空、または空白だけなら true
     */
    function isBlankLine(lineText) {
        return /^[\s　]*$/.test(lineText);
    }

    /**
     * 空行（空白だけの行を含む）を取り除く。
     * 残った行のあいだの改行は、元の文字（段落改行の CR、強制改行の LF）をそのまま使う。
     * @param {string} textContent - 対象の文字列
     * @returns {string} 空行を取り除いた文字列
     */
    function removeEmptyLines(textContent) {
        if (!textContent) return "";

        var keptText = "";
        var lineBuffer = "";
        var breakBeforeLine = "";

        /* 末尾は改行が無くても1行として確定させるため、長さの位置まで回す / Close the last line even without a trailing break */
        for (var i = 0; i <= textContent.length; i++) {
            var isEndOfText = (i === textContent.length);
            var currentChar = isEndOfText ? "" : textContent.charAt(i);

            if (!isEndOfText && currentChar !== "\r" && currentChar !== "\n") {
                lineBuffer += currentChar;
                continue;
            }

            /* 空行は改行ごと捨てるので、残す行の直前の改行だけが書き出される / Dropping a blank line drops its break too */
            if (!isBlankLine(lineBuffer)) {
                keptText += (keptText === "" ? "" : breakBeforeLine) + lineBuffer;
            }
            lineBuffer = "";

            if (isEndOfText) break;

            breakBeforeLine = currentChar;
            /* CR+LF は1つの改行として扱う / Treat CR+LF as a single break */
            if (currentChar === "\r" && textContent.charAt(i + 1) === "\n") {
                breakBeforeLine = "\r\n";
                i++;
            }
        }
        return keptText;
    }

    /**
     * テキストフレームの内容から空行を落とした文字列を返す。
     * ここで空行を落としておけば、1行目の取り出しも書き戻しも空行を意識しなくてよい。
     * @param {TextFrame} textFrame - 読み取るテキストフレーム
     * @returns {string|null} 空行を除いた文字列。何も残らなければ null
     */
    function readNonBlankText(textFrame) {
        var nonBlankText = removeEmptyLines(textFrame.contents);
        return (nonBlankText.length > 0) ? nonBlankText : null;
    }

    /**
     * 一度ペーストして、貼り付けられたテキストフレームから文字列を読み取る。
     * 読み取り後は貼り付けたオブジェクトを削除し、元の選択へ戻す。
     * 貼り付け前に選択を解除するのは、ペーストが実行されなかったときに
     * 元の選択を「貼り付いたもの」と誤認して削除しないため。
     * @param {Document} doc - 対象ドキュメント
     * @param {Object[]} originalSelection - 復元する元の選択
     * @returns {string|null} クリップボードの文字列。テキストが無い、または失敗した場合は null
     */
    function readClipboardText(doc, originalSelection) {
        var clipboardText = null;
        var pastedItems = null;
        var pasteError = null;

        try {
            /* 貼り付いたものだけを確実に拾うため、先に選択を空にする / Clear the selection first so only the pasted items are captured */
            setSelection(doc, null);
            pastedItems = pasteClipboardItems(doc);

            var pastedTextFrame = findFirstTextFrame(pastedItems);
            if (pastedTextFrame) {
                /* 中身が空なら「残りなし」として null のまま / Nothing usable means there is nothing left to apply */
                clipboardText = readNonBlankText(pastedTextFrame);
            }
        } catch (e) {
            pasteError = String(e);
        }

        /* 成否にかかわらず、貼り付けた分を消して元の選択へ戻す / Clean up and restore regardless of the outcome */
        removeItems(pastedItems);
        setSelection(doc, originalSelection);

        /* 画面を元に戻してから知らせる。ペースト自体が起きなかった場合と、貼り付いたがテキストが無い場合を区別する
           Report only after the canvas is restored; tell an unusable clipboard apart from a paste without text */
        if (!alertPasteFailure(pasteError, pastedItems) && clipboardText === null) {
            alert(getLabel("alert.noTextInClipboard"));
        }
        return clipboardText;
    }

    /**
     * 一時テキストフレーム経由でクリップボードを書き換える。
     * Illustrator には文字列を直接クリップボードへ送る API が無いため、
     * 内容を持つフレームを作ってコピーし、すぐに削除する。
     * 追加直後のフレームは再描画しないとコピー対象として扱われないことがあり、
     * また app.copy() は黙って無視される場合があるためメニューコマンドを使う。
     * @param {Document} doc - 対象ドキュメント
     * @param {string} textContent - クリップボードに残す文字列
     * @returns {void}
     */
    function writeTextToClipboard(doc, textContent) {
        var tempFrame = null;
        var writeError = null;

        try {
            tempFrame = doc.activeLayer.textFrames.add();
            tempFrame.contents = textContent;

            /* 追加したフレームを画面に反映してから選択する / Flush the new frame to the canvas before selecting it */
            app.redraw();
            app.executeMenuCommand("deselectall");
            tempFrame.selected = true;
            app.redraw();

            app.executeMenuCommand("copy");
            /* コピーが確定してから元のフレームを消す / Let the copy settle before deleting the source */
            app.redraw();
        } catch (e) {
            writeError = String(e);
        }

        if (tempFrame) removeItems([tempFrame]);
        if (writeError) {
            alert(getLabel("alert.clipboardWriteError") + writeError);
        }
    }

    // =========================================
    // ウィンドウ中央への配置 / Placement at the center of the window
    // =========================================

    /**
     * 表示中のウィンドウの中心座標を返す
     * @param {Document} doc - 対象ドキュメント
     * @returns {{x: number, y: number}} ウィンドウ中心の座標
     */
    function getActiveViewCenter(doc) {
        var viewCenter = doc.activeView.centerPoint;
        return { x: viewCenter[0], y: viewCenter[1] };
    }

    /**
     * オブジェクト群を囲む矩形の中心座標を返す（クリップグループはマスクの範囲で測る）
     * @param {Object[]} targetItems - 対象のページアイテム
     * @returns {{x: number, y: number}|null} 中心座標。座標を取れるものが無ければ null
     */
    function getItemsCenter(targetItems) {
        /* TextRange など座標を持たないものは測れずに外れる / Things without geometry, such as a TextRange, drop out */
        var unionBounds = getClipAwareUnionBounds(targetItems, false);
        if (!unionBounds) return null;
        return { x: (unionBounds[0] + unionBounds[2]) / 2, y: (unionBounds[1] + unionBounds[3]) / 2 };
    }

    /**
     * オブジェクト群をまとめてウィンドウの中央へ移動する。
     * 複数まとめて貼り付いた場合も並びを崩さないよう、全体を同じ量だけ動かす。
     * @param {Document} doc - 対象ドキュメント
     * @param {Object[]} targetItems - 移動するページアイテム
     * @returns {void}
     */
    function moveItemsToViewCenter(doc, targetItems) {
        var itemsCenter = getItemsCenter(targetItems);
        if (!itemsCenter) return;

        var viewCenter = getActiveViewCenter(doc);
        var offsetX = viewCenter.x - itemsCenter.x;
        var offsetY = viewCenter.y - itemsCenter.y;

        for (var i = 0; i < targetItems.length; i++) {
            try {
                targetItems[i].translate(offsetX, offsetY);
            } catch (e) {
                /* 動かせないものはその位置に残す / Leave behind whatever cannot be moved */
            }
        }
    }

    /**
     * 貼り付いたテキストフレームの内容を1行目だけに切り詰め、2行目以降を返す。
     * @param {TextFrame} pastedTextFrame - 貼り付いたテキストフレーム
     * @returns {string|null} 2行目以降の文字列。1行しかなければ空文字、使えるテキストが無ければ null
     */
    function trimPastedTextToFirstLine(pastedTextFrame) {
        var pastedText = readNonBlankText(pastedTextFrame);
        if (pastedText === null) return null;

        var splitResult = splitFirstLine(pastedText);
        pastedTextFrame.contents = splitResult.firstLine;
        /* 切り詰めた分を座標へ反映させてから中央を測る / Flush the trim before the bounds are measured */
        app.redraw();
        return splitResult.remainder;
    }

    /**
     * クリップボードの1行目だけをウィンドウの中央へ貼り付ける。
     * 適用先のテキストフレームが無いときの動作で、貼り付いたフレームを1行目だけに切り詰め、
     * 2行目以降はクリップボードへ書き戻して次の実行に引き継ぐ。
     * ペースト位置もウィンドウの中央だが、1行目に切り詰めるとフレームの大きさが変わるため、
     * 切り詰めたあとに測り直して中央へ置き直す。
     * テキストを含まない内容は切り詰めも書き戻しも行わず、そのまま中央へ置く。
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function pasteFirstLineAtViewCenter(doc) {
        var pastedItems = null;
        var pasteError = null;

        try {
            pastedItems = pasteClipboardItems(doc);
        } catch (e) {
            pasteError = String(e);
        }

        if (alertPasteFailure(pasteError, pastedItems)) return;

        /* テキスト以外は取り出す行が無いので、貼り付いたまま中央へ置く / Non-text has no line to take, so leave the paste as it is */
        var pastedTextFrame = findFirstTextFrame(pastedItems);
        var remainder = pastedTextFrame ? trimPastedTextToFirstLine(pastedTextFrame) : "";

        if (remainder === null) {
            /* 空行だけの内容は、貼り付けた分を片付けてから知らせる / Clean up before reporting a paste of blank lines only */
            removeItems(pastedItems);
            setSelection(doc, null);
            alert(getLabel("alert.noTextInClipboard"));
            return;
        }

        moveItemsToViewCenter(doc, pastedItems);

        /* 貼り付けた行を取り除いた残りをクリップボードへ戻し、次の実行で続きから進めるようにする / Put the remaining lines back so the next run continues where this one stopped */
        var hasRemainder = (remainder.length > 0);
        if (hasRemainder) {
            writeTextToClipboard(doc, remainder);
        }

        /* 貼り付けた直後に手を加えられるよう選択したままにする / Keep the paste selected so it can be edited right away */
        setSelection(doc, pastedItems);
        app.redraw();

        if (pastedTextFrame && !hasRemainder) {
            alert(getLabel("alert.lastLineApplied"));
        }
    }

    // =========================================
    // テキストの適用 / Text application
    // =========================================

    /**
     * 文字列を最初の改行で1行目と残りに分ける。
     * Illustrator のテキストは段落改行が CR、強制改行が LF になるため、どちらも区切りとして扱う。
     * @param {string} textContent - 分割する文字列
     * @returns {{firstLine: string, remainder: string}} 1行目と、改行を取り除いた残り
     */
    function splitFirstLine(textContent) {
        var lineBreak = /\r\n|\r|\n/.exec(textContent);
        if (!lineBreak) {
            return { firstLine: textContent, remainder: "" };
        }
        return {
            firstLine: textContent.substring(0, lineBreak.index),
            remainder: textContent.substring(lineBreak.index + lineBreak[0].length)
        };
    }

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
    // メイン処理 / Main
    // =========================================

    /**
     * クリップボードの1行目を選択中のテキストフレームへ適用し、残りをクリップボードへ書き戻す。
     * 適用先が無い場合は、1行目をウィンドウ中央へ貼り付ける。
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var originalSelection = captureSelection(doc);

        /* 文字を選択して編集中なら、編集を抜けて親テキストフレームを対象にする / Leave text editing and target the parent frame */
        var editingFrame = captureEditingFrame(originalSelection);
        if (editingFrame) {
            leaveTextEditing(doc);
            originalSelection = [editingFrame];
            setSelection(doc, originalSelection);
        }

        /* グループやクリップグループの中は、ペーストを挟んで参照が古くなる前にたどっておく / Walk into groups before the paste cycle can stale the references */
        var targetFrames = collectSelectionTextFrames(originalSelection, { unique: false });
        /* 適用先が無いときは、1行目をウィンドウ中央へ貼り付けて残りをクリップボードへ戻す / Without a target, paste the first line at the center of the window and keep the rest */
        if (targetFrames.length === 0) {
            pasteFirstLineAtViewCenter(doc);
            return;
        }

        var clipboardText = readClipboardText(doc, originalSelection);
        if (clipboardText === null) return;

        var splitResult = splitFirstLine(clipboardText);

        var errorMessages = [];
        applyTextToFrames(targetFrames, splitResult.firstLine, errorMessages);
        /* 選択数だけダイアログが出ないよう、まとめて1回だけ知らせる / Report every failure in a single alert */
        if (errorMessages.length > 0) {
            alert(getLabel("alert.replaceError") + errorMessages.join("\n"));
        }

        /* 適用した行を取り除いた残りをクリップボードへ戻し、次の実行で続きから進めるようにする / Put the remaining lines back so the next run continues where this one stopped */
        var hasRemainder = (splitResult.remainder.length > 0);
        if (hasRemainder) {
            writeTextToClipboard(doc, splitResult.remainder);
        }

        /* 置換後の表示を確実に更新するため、選択を解除してから戻す / Clear and reset the selection so the redraw reflects the change */
        setSelection(doc, null);
        app.redraw();
        setSelection(doc, originalSelection);

        if (!hasRemainder) {
            alert(getLabel("alert.lastLineApplied"));
        }
    }

    main();

})();
