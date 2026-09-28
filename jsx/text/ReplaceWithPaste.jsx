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
var SCRIPT_VERSION  = "v2.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2024-10-28";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

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
    var BUTTON_ROW_TOP_MARGIN = 10;               /* ボタンエリアの上余白 / top margin of the button row */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * Illustrator の UI 言語から表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = detectUILanguage();

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

    /**
     * ラベルをドット区切りのキーで引く
     * @param {string} labelPath - "alert.noDocument" のようなドット区切りキー
     * @returns {string} 現在の表示言語のラベル。見つからない場合はキーをそのまま返す
     */
    function getLabel(labelPath) {
        var pathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            if (!labelNode) return labelPath;
            labelNode = labelNode[pathKeys[i]];
        }
        return (labelNode && labelNode[uiLang]) ? labelNode[uiLang] : labelPath;
    }

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
            var parentFrame = null;
            if (textRange.parent && textRange.parent.typename === "TextFrame") {
                parentFrame = textRange.parent;
            } else if (textRange.story && textRange.story.textFrames.length > 0) {
                parentFrame = textRange.story.textFrames[0];
            }
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

    /**
     * テキストフレームを再帰的に集める。
     * クリップグループも typename は GroupItem なので、同じ経路で中までたどれる。
     * 子は生のコレクションのまま持たず配列へ写し取る。内容を書き換えると
     * コレクションの中身が変わり、たどっている途中で取りこぼすため。
     * TextRange が渡された場合は親のテキストフレームに読み替える。
     * @param {Object} searchItem - 探索対象のページアイテムまたは TextRange
     * @param {TextFrame[]} collectedFrames - 収集先の配列
     * @returns {void}
     */
    function collectTextFrames(searchItem, collectedFrames) {
        if (!searchItem) return;

        try {
            if (searchItem.typename === "TextRange" && searchItem.parent && searchItem.parent.typename === "TextFrame") {
                searchItem = searchItem.parent;
            }

            if (searchItem.typename === "TextFrame") {
                collectedFrames.push(searchItem);
                return;
            }

            if (searchItem.typename !== "GroupItem") return;

            var childItems = [];
            for (var i = 0; i < searchItem.pageItems.length; i++) {
                childItems.push(searchItem.pageItems[i]);
            }
            for (var j = 0; j < childItems.length; j++) {
                collectTextFrames(childItems[j], collectedFrames);
            }
        } catch (e) {
            /* シンボルやエンベロープなど、中をたどれないものは対象外にする / Skip what we cannot walk into, such as symbols and envelopes */
        }
    }

    /**
     * 配列やコレクションから、テキストフレームをまとめて集める
     * @param {Object[]} searchItems - 探索対象の配列またはコレクション
     * @returns {TextFrame[]} 見つかったテキストフレームの配列。順序は探索順
     */
    function collectTextFramesFrom(searchItems) {
        var collectedFrames = [];
        if (!searchItems) return collectedFrames;
        for (var i = 0; i < searchItems.length; i++) {
            collectTextFrames(searchItems[i], collectedFrames);
        }
        return collectedFrames;
    }

    /**
     * 配列から最初のテキストフレームを探す。
     * 他アプリからのペーストはグループやクリップグループにまとめられることがあるため、中も再帰的にたどる。
     * @param {Object[]} searchItems - 探索対象の配列またはコレクション
     * @returns {TextFrame|null} 見つかったテキストフレーム。なければ null
     */
    function findFirstTextFrame(searchItems) {
        var foundFrames = collectTextFramesFrom(searchItems);
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
     * visibleBounds から中心と長辺・短辺を求める
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
     * 複数オブジェクトの visibleBounds を合わせた外接矩形を返す
     * @param {PageItem[]} pageItems - 対象オブジェクト
     * @returns {number[]} [左, 上, 右, 下]
     */
    function getUnionBounds(pageItems) {
        var unionBounds = pageItems[0].visibleBounds;
        for (var i = 1; i < pageItems.length; i++) {
            var itemBounds = pageItems[i].visibleBounds;
            unionBounds = [
                Math.min(unionBounds[0], itemBounds[0]),
                Math.max(unionBounds[1], itemBounds[1]),
                Math.max(unionBounds[2], itemBounds[2]),
                Math.min(unionBounds[3], itemBounds[3])
            ];
        }
        return unionBounds;
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
        var pastedMetrics = measureBounds(getUnionBounds(pastedItems));

        for (var i = 0; i < pastedItems.length; i++) {
            var pastedItem = pastedItems[i];
            var itemMetrics = measureBounds(pastedItem.visibleBounds);

            /* 各オブジェクトの中心は、全体の中心からの距離を同じ比率で伸縮させた位置へ / Keep each item's offset from the group center, scaled */
            var destX = targetMetrics.centerX + (itemMetrics.centerX - pastedMetrics.centerX) * scale;
            var destY = targetMetrics.centerY + (itemMetrics.centerY - pastedMetrics.centerY) * scale;

            if (scale !== 1) pastedItem.resize(scale * 100, scale * 100);

            var scaledMetrics = measureBounds(pastedItem.visibleBounds);
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
                var targetMetrics = measureBounds(targetItem.visibleBounds);

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
                    baseMetrics: measureBounds(getUnionBounds(pastedItems)),
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

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子は DIALOG_* / prepareDialogWindow / *DialogLeft* / getSelectionViewSpan の名前
    // 2. スクリプトの先頭（#target の次の行）に #targetengine "<SCRIPT_NAME>Engine" を置く。
    //    #targetengine が無いと $.global が実行ごとに消え、位置を覚えられない。すでにあればそのまま使う
    // 3. ダイアログの show() の直前で prepareDialogWindow(dialog, SCRIPT_NAME) を呼ぶ。
    //    それまでに入れた onShow / onMove / onClose はそのまま生かし、あとに位置の復元・記録をつなぐ
    //      prepareDialogWindow(mainDialog, SCRIPT_NAME);
    //      var dialogResult = mainDialog.show();
    //    同じスクリプトで複数のダイアログを開くときは、2つ目以降のキーを変える（SCRIPT_NAME + "_colorPicker" など）
    //    同じダイアログを何度も開くときも、毎回 show() の直前で呼んでよい（2回目からは選択範囲を測り直すだけ）
    // 4. 初めて開くとき（記録が無いとき）は、スクリプト側の配置（中央・オフセットなど）がそのまま効く
    // 5. 開く位置が選択中のオブジェクトに重なりそうなら左右の反対側へずらす（Illustrator のみ）。
    //    ずらした位置は記録せず、ユーザーが動かしたときだけ記録する
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var DIALOG_OPACITY = 0.97;       /* ダイアログの不透明度 / dialog opacity */
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
            if (!selectedItems || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // ダイアログ / Dialog
    // =========================================

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

        var btnRowGroup = sizeDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.alignment = ["fill", "bottom"];

        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnRightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

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
        var targetFrames = collectTextFramesFrom(originalSelection);

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
