#targetengine "TextMemoEngine"
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

メモ入力用のフローティングパレットです。
選択オブジェクトやクリップボードからテキストを読み込み、空行・改行を整理して、テキストファイルへの保存やクリップボードへのコピーができます。

詳細は README を参照してください。

### Overview

A floating palette for taking notes.
It can pull text in from the selection or the clipboard, tidy up blank lines and returns, and save the result to a text file or copy it back to the clipboard.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiMemoPallete";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-06-15";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-16";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiMemoPallete.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiMemoPallete.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n41e91e4b1a09"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    /**
     * 原案 / Original idea
     * @author こじらせたクマー (@nice_lotus120)
     * @discussion https://note.com/nice_lotus120/n/n6291a432b30d
     */

    // Released under the MIT license
    // http://opensource.org/licenses/mit-license.php

    // =========================================
    // ユーザー設定 / User settings
    // =========================================
    /* パレットの透明度 / Palette opacity */
    var PALETTE_OPACITY = 0.97;

    /* 保存先モード / Save destination mode
       'dialog'  = 保存ダイアログで場所と名前を選ぶ（A）/ Choose location & name via the save dialog (A)
       'desktop' = 常にデスクトップへ memo-<ドキュメント名>-<yyyymmdd>.txt（B）/ Always save to the Desktop (B) */
    var SAVE_LOCATION_MODE = 'desktop'; // デフォルトは B / Default is B

    /* パレットを閉じたときに内容をクリアするか / Whether to clear the content when the palette is closed
       true  = 閉じると内容をクリア / Clear on close
       false = 内容を保持して次回復元（デフォルト）/ Keep content and restore next time (default) */
    var CLEAR_ON_CLOSE = false;

    /* 同じ行とみなす上端の差（pt）/ Difference in top edge treated as the same row (pt) */
    var SAME_ROW_TOLERANCE_PT = 10;

    // =========================================
    // レイアウト / Layout
    // =========================================
    var PALETTE_MARGINS       = [15, 15, 15, 15]; /* パレット外周の余白 [左,上,右,下] */
    var PALETTE_MIN_SIZE      = [300, 200];       /* パレットの最小サイズ [幅,高さ] */
    var PALETTE_INITIAL_WIDTH = 400;              /* 保存位置が無いときの初期幅 */
    var MEMO_AREA_MIN_SIZE    = [300, 100];       /* テキスト欄の最小サイズ [幅,高さ] */
    var BUTTON_SIZE           = [76, 24];         /* 標準ボタンのサイズ [幅,高さ] */
    var WIDE_BUTTON_SIZE      = [120, 24];        /* 長いラベルのボタンのサイズ [幅,高さ] */
    var BUTTON_HEIGHT         = 24;               /* 幅をラベルに合わせるボタンの高さ */

    // =========================================
    // ローカライズ / Localization
    // =========================================
    /* 現在の言語を判定（ロケールが ja 始まりなら日本語）/ Detect UI language (Japanese if locale starts with "ja") */
    var uiLang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        dialog: {
            title: { ja: "テキスト一時保管", en: "Text Stash" }
        },
        fieldLabel: {
            mode: { ja: "モード:", en: "Mode:" }
        },
        radio: {
            replace: { ja: "置き換え", en: "Replace" },
            append: { ja: "追加", en: "Append" }
        },
        button: {
            load: { ja: "読み込む", en: "Load" },
            save: { ja: "保存", en: "Save" },
            clear: { ja: "クリア", en: "Clear" },
            copyAll: { ja: "すべてをコピー", en: "Copy All" },
            removeBlanks: { ja: "空行削除", en: "Remove Blanks" },
            removeBreaks: { ja: "改行削除", en: "Remove Breaks" }
        },
        tooltip: {
            load: {
                ja: "選択オブジェクトのテキストを読み込みます（モードに従って追加／置き換え）。「追加」モードでは既存テキストとの間に空行を入れます",
                en: "Load text from the selected objects (append or replace per mode). In Append mode, a blank line is inserted between the existing and new text"
            },
            removeBlanks: {
                ja: "テキスト欄の空行（空白のみの行を含む）を削除します",
                en: "Remove blank lines (including whitespace-only lines) from the memo"
            },
            removeBreaks: {
                ja: "テキスト欄の改行をすべて削除して1行にまとめます",
                en: "Remove all line breaks in the memo, joining everything into one line"
            },
            clear: {
                ja: "テキスト欄を空にして再起動します",
                en: "Clear the memo and restart"
            },
            save: {
                ja: "メモをデスクトップにテキストファイルとして保存します",
                en: "Save the memo to the Desktop as a text file"
            },
            copyAll: {
                ja: "テキスト欄の内容をクリップボードへコピーします",
                en: "Copy the memo to the clipboard"
            }
        },
        alert: {
            savePrompt: {
                ja: "テキストファイル（.txt）の保存先を指定してください",
                en: "Choose where to save the text file (.txt)"
            },
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noTextFrame: {
                ja: "選択されたオブジェクトにテキストフレームがありません。",
                en: "No text frame is found in the selection."
            },
            loadFailed: { ja: "テキストを追加できませんでした: ", en: "Could not load text: " },
            writeFailed: { ja: "ファイルを書き込めませんでした。", en: "Could not write the file." },
            clipboardEmpty: {
                ja: "選択オブジェクトもクリップボードのテキストもありません。",
                en: "No selected object and no clipboard text."
            },
            copyFailed: { ja: "クリップボードへコピーできませんでした。", en: "Could not copy to the clipboard." }
        },
        status: {
            unsaved: { ja: "(ドキュメント未保存)", en: "(Unsaved document)" }
        },
        footer: {
            source: { ja: "保存元", en: "Source" },
            savedAt: { ja: "保存日時", en: "Saved at" }
        }
    };

    /**
     * ラベルノードから現在の言語の文言を返す（{slash} は / に置換）
     * @param {object} labelNode - { ja: string, en: string } 形式のラベル
     * @returns {string} 現在の言語の文言
     */
    function getLabel(labelNode) {
        if (!labelNode) return "";
        var labelText = labelNode[uiLang] || labelNode.en || "";
        return labelText.replace(/\{slash\}/g, "/");
    }

    // =========================================
    // メインエンジン用ワーカー / Main-engine workers
    // =========================================
    /* 以下の WORKER_* はメインエンジン側で評価される文字列。常駐エンジンの app はパレット表示中に
       DOM 接続を失い "there is no document" を投げるため、DOM に触る処理は BridgeTalk で
       メインエンジンへ渡す（結果は非同期）。結果は "マーカー:ペイロード" 形式で返し、
       マルチバイト欠落を防ぐためテキストは encodeURIComponent で受け渡す。
       These WORKER_* strings are evaluated in the main engine, which keeps a live DOM connection
       while the palette is up. Results come back asynchronously as "MARKER:payload". */

    /* アクティブドキュメントの取得 / Resolve the active document */
    var WORKER_TARGET_DOC =
        '    if (app.documents.length === 0) return "NODOC:";' +
        '    var targetDoc = app.activeDocument;';

    /* 選択の待避と復元（breakLink・貼り付けで選択が変わるため）/ Save and restore the selection */
    var WORKER_SELECTION =
        '    var savedSelection = [];' +
        '    function saveSelection() {' +
        '        savedSelection = [];' +
        '        for (var i = 0; i < targetDoc.selection.length; i++) savedSelection.push(targetDoc.selection[i]);' +
        '    }' +
        '    function restoreSelection() {' +
        '        try {' +
        '            targetDoc.selection = null;' +
        // 文字ツールで編集中の TextRange には selected が無いため飛ばす
        '            for (var i = 0; i < savedSelection.length; i++) {' +
        '                if (savedSelection[i].typename !== "TextRange") savedSelection[i].selected = true;' +
        '            }' +
        '        } catch (err) {}' +
        '    }';

    /* テキストの収集・整列・連結 / Collect, sort and join text */
    var WORKER_COLLECTOR =
        '    var collected = [];' +
        // シンボルの読み取り方は呼び出し側で差し替える（未設定ならシンボルは無視）
        '    var readSymbolItem = null;' +
        // 実質的に空のフレームは無視し、テキストと位置（左端・上端）を控える
        '    function pushFrame(frame) {' +
        '        var contents = frame.contents;' +
        '        if (!contents || contents.replace(/[\\r\\n\\x03]/g, "").replace(/\\s+/g, "") === "") return;' +
        '        var frameBounds = frame.geometricBounds;' +
        '        collected.push({ text: contents, left: frameBounds[0], top: frameBounds[1] });' +
        '    }' +
        // グループ・クリップグループの中もたどる
        '    function collectFrom(item) {' +
        '        if (item.typename === "TextFrame") { pushFrame(item); return; }' +
        '        if (item.typename === "SymbolItem") { if (readSymbolItem) readSymbolItem(item); return; }' +
        '        if (item.pageItems) { for (var i = 0; i < item.pageItems.length; i++) collectFrom(item.pageItems[i]); }' +
        '    }' +
        // カンバス上の位置で上から順（同じ高さは左から右）に並べて連結する
        '    function joinCollected(dedupe) {' +
        '        collected.sort(function (a, b) {' +
        '            if (Math.abs(b.top - a.top) <= ' + SAME_ROW_TOLERANCE_PT + ') return a.left - b.left;' +
        '            return b.top - a.top;' +
        '        });' +
        '        var seenTexts = {};' +
        '        var texts = [];' +
        '        for (var i = 0; i < collected.length; i++) {' +
        '            var seenKey = "k_" + collected[i].text;' +
        '            if (dedupe && seenTexts[seenKey]) continue;' +
        '            seenTexts[seenKey] = true;' +
        '            texts.push(collected[i].text);' +
        '        }' +
        '        return texts.join(String.fromCharCode(10));' +
        '    }';

    /* 選択オブジェクトからテキストを収集する（シンボル内も対象）/ Collect text from the selection (symbols included) */
    var WORKER_READ_SELECTION =
        '(function () {' +
        WORKER_TARGET_DOC +
        '    if (!targetDoc.selection || targetDoc.selection.length === 0) return "NOSEL:";' +
        WORKER_SELECTION +
        WORKER_COLLECTOR +
        // シンボルは一時レイヤーへ複製してから展開する（元のシンボルには触れない）
        '    var symbolReadLayer = null;' +
        '    function getSymbolReadLayer() {' +
        '        if (!symbolReadLayer) {' +
        '            symbolReadLayer = targetDoc.layers.add();' +
        '            symbolReadLayer.name = "__TextMemo_symbol_read__";' +
        '        }' +
        '        return symbolReadLayer;' +
        '    }' +
        // 複製なのでその場で breakLink してよい（入れ子シンボルも多段展開）
        '    function harvestDuplicate(item) {' +
        '        if (!item) return;' +
        '        if (item.typename === "TextFrame") { pushFrame(item); return; }' +
        '        if (item.typename === "SymbolItem") {' +
        '            try {' +
        '                targetDoc.selection = null;' +
        '                item.selected = true;' +
        '                item.breakLink();' +
        '                var expandedItems = [];' +
        '                for (var i = 0; i < targetDoc.selection.length; i++) expandedItems.push(targetDoc.selection[i]);' +
        '                for (var j = 0; j < expandedItems.length; j++) harvestDuplicate(expandedItems[j]);' +
        '            } catch (err) {}' +
        '            return;' +
        '        }' +
        '        if (item.pageItems) { for (var k = 0; k < item.pageItems.length; k++) harvestDuplicate(item.pageItems[k]); }' +
        '    }' +
        '    readSymbolItem = function (symbolItem) {' +
        '        harvestDuplicate(symbolItem.duplicate(getSymbolReadLayer(), ElementPlacement.PLACEATBEGINNING));' +
        '    };' +
        '    saveSelection();' +
        '    for (var s = 0; s < savedSelection.length; s++) collectFrom(savedSelection[s]);' +
        // 一時レイヤーは中身ごと削除する
        '    if (symbolReadLayer) { try { symbolReadLayer.remove(); } catch (err) {} }' +
        '    restoreSelection();' +
        '    if (collected.length === 0) return "NOTF:";' +
        '    return "OK:" + encodeURIComponent(joinCollected(false));' +
        '})();';

    /* クリップボードを貼り付けてテキストを読み取る（貼り付けたものは削除）/ Paste the clipboard, read the text, remove the pasted items */
    var WORKER_PASTE_CLIPBOARD =
        '(function () {' +
        WORKER_TARGET_DOC +
        WORKER_SELECTION +
        WORKER_COLLECTOR +
        // 貼り付け前に必ず選択解除する（解除しないと、貼り付かなかったときに元の選択を削除してしまう）
        '    saveSelection();' +
        '    targetDoc.selection = null;' +
        '    var pastedItems = [];' +
        '    try {' +
        '        app.executeMenuCommand("pasteInAllArtboard");' +
        // 文字ツールで編集中は TextRange が返る。削除すると選択中の文字自体が消えるため除外する
        '        for (var i = 0; i < targetDoc.selection.length; i++) {' +
        '            if (targetDoc.selection[i].typename !== "TextRange") pastedItems.push(targetDoc.selection[i]);' +
        '        }' +
        '    } catch (err) {}' +
        '    for (var j = 0; j < pastedItems.length; j++) collectFrom(pastedItems[j]);' +
        '    for (var k = 0; k < pastedItems.length; k++) { try { pastedItems[k].remove(); } catch (err) {} }' +
        '    restoreSelection();' +
        '    if (collected.length === 0) return "NOTX:";' +
        // pasteInAllArtboard はアートボードごとに複製を作るため重複を除く
        '    return "OK:" + encodeURIComponent(joinCollected(true));' +
        '})();';

    /**
     * テキストをクリップボードへ書き戻すワーカーのソースを組み立てる
     * Illustrator には app.system が無いため、一時テキストフレームを作ってコピーし、すぐ削除する
     * @param {string} memoText - コピーするテキスト
     * @returns {string} メインエンジンで評価するソース
     */
    function buildCopyWorkerSource(memoText) {
        return '(function () {' +
            '    var copyText = decodeURIComponent("' + encodeURIComponent(memoText) + '");' +
            // ドキュメントが無ければ一時ドキュメントを作り、コピー後に閉じる
            '    var usingTempDoc = (app.documents.length === 0);' +
            '    var targetDoc = usingTempDoc ? app.documents.add() : app.activeDocument;' +
            WORKER_SELECTION +
            '    if (!usingTempDoc) saveSelection();' +
            // ロック／非表示レイヤーでは textFrames.add() が 8705 で失敗するため一時的に解除する
            '    var editLayer = targetDoc.activeLayer;' +
            '    var layerWasLocked = editLayer.locked;' +
            '    var layerWasHidden = !editLayer.visible;' +
            '    var tempFrame = null;' +
            '    var copied = false;' +
            '    try {' +
            '        editLayer.locked = false;' +
            '        editLayer.visible = true;' +
            '        tempFrame = editLayer.textFrames.add();' +
            '        tempFrame.contents = copyText;' +
            // app.copy() は黙って無視されることがあるため、再描画を挟んでメニューコマンドでコピーする
            '        app.redraw();' +
            '        app.executeMenuCommand("deselectall");' +
            '        tempFrame.selected = true;' +
            '        app.redraw();' +
            '        app.executeMenuCommand("copy");' +
            '        app.redraw();' + // コピー確定前に削除すると空になる
            '        copied = true;' +
            '    } catch (err) {}' +
            '    if (tempFrame) { try { tempFrame.remove(); } catch (err) {} }' +
            '    editLayer.locked = layerWasLocked;' +
            '    editLayer.visible = !layerWasHidden;' +
            '    if (usingTempDoc) { try { targetDoc.close(SaveOptions.DONOTSAVECHANGES); } catch (err) {} }' +
            '    else restoreSelection();' +
            '    return copied ? "OK" : "ERR";' +
            '})();';
    }

    /**
     * ワーカーをメインエンジンで評価し、"マーカー:ペイロード" の結果を状態文字列にして返す
     * @param {string} workerSource - メインエンジンで評価するソース
     * @param {object} statusByMarker - マーカーと状態文字列の対応表
     * @param {function} onComplete - 完了コールバック onComplete(status, text)
     * @returns {void}
     */
    function callMainEngine(workerSource, statusByMarker, onComplete) {
        var bridgeMessage = new BridgeTalk();
        bridgeMessage.target = 'illustrator';
        bridgeMessage.body = workerSource;
        bridgeMessage.onResult = function (response) {
            var payload = response.body || '';
            var colonIndex = payload.indexOf(':');
            var marker = (colonIndex >= 0) ? payload.substring(0, colonIndex) : payload;
            var status = statusByMarker[marker];
            if (!status) {
                onComplete('error', payload);
                return;
            }
            onComplete(status, (status === 'ok' && colonIndex >= 0) ? decodeURIComponent(payload.substring(colonIndex + 1)) : '');
        };
        bridgeMessage.onError = function (response) {
            onComplete('error', response ? response.body : '');
        };
        bridgeMessage.send();
    }

    /**
     * 選択オブジェクトのテキストを取得する
     * @param {function} onComplete - 完了コールバック onComplete(status, text)
     *   status = 'ok' | 'nodoc' | 'nosel' | 'notextframe' | 'error'
     * @returns {void}
     */
    function fetchSelectedText(onComplete) {
        callMainEngine(WORKER_READ_SELECTION, {
            OK: 'ok',
            NODOC: 'nodoc',
            NOSEL: 'nosel',
            NOTF: 'notextframe'
        }, onComplete);
    }

    /**
     * クリップボードのテキストを取得する
     * @param {function} onComplete - 完了コールバック onComplete(status, text)
     *   status = 'ok' | 'nodoc' | 'notext' | 'error'
     * @returns {void}
     */
    function fetchClipboardText(onComplete) {
        callMainEngine(WORKER_PASTE_CLIPBOARD, {
            OK: 'ok',
            NODOC: 'nodoc',
            NOTX: 'notext'
        }, onComplete);
    }

    /**
     * テキストをクリップボードへコピーする
     * @param {string} memoText - コピーするテキスト
     * @param {function} onComplete - 完了コールバック onComplete(status) status = 'ok' | 'error'
     * @returns {void}
     */
    function copyTextToClipboard(memoText, onComplete) {
        callMainEngine(buildCopyWorkerSource(memoText), { OK: 'ok' }, onComplete);
    }

    // =========================================
    // 二重起動防止 / Prevent duplicate launch
    // =========================================
    /* 既にパレットが開いていれば前面に出して終了する（常駐エンジンなので二重に生成しない）/
       If the palette is already open, bring it to the front and stop (avoid a second instance on this persistent engine) */
    if ($.global.__TextMemoWindow) {
        try {
            if ($.global.__TextMemoWindow.visible) {
                $.global.__TextMemoWindow.active = true; // 前面へ / Bring to front
                return;
            }
        } catch (err) {
            // 参照が失効していれば通常どおり新規作成へ進む / If the reference is stale, fall through and create a new window
        }
    }

    // =========================================
    // セッション保持用変数 / Session state
    // =========================================
    var restoredMemoText = $.global.__TextMemoContent || '';
    var restoredPaletteBounds = $.global.__TextMemoBounds || null;

    /* アクティブドキュメント名（保存フッター・ファイル名用）/ Active document name (used in the footer and file name) */
    var sourceDocumentName = getLabel(LABELS.status.unsaved);
    var initialDocument = getActiveDocument();
    if (initialDocument) {
        sourceDocumentName = initialDocument.fullName ? decodeURI(initialDocument.fullName.name) : initialDocument.name;
    }

    /* 再起動（保存・クリア）由来の close では内容を消さないためのフラグ / Flag so a restart-triggered close does not clear the content */
    var isRestartingScript = false;

    // =========================================
    // UI 構築 / Build UI
    // =========================================
    var initialPaletteBounds = (restoredPaletteBounds && restoredPaletteBounds.length === 4) ? restoredPaletteBounds : undefined;
    var memoPalette = new Window('palette', getLabel(LABELS.dialog.title) + ' ' + SCRIPT_VERSION, initialPaletteBounds, {
        resizeable: true
    });
    memoPalette.opacity = PALETTE_OPACITY;
    memoPalette.orientation = 'column';
    memoPalette.alignChildren = ['fill', 'top'];
    memoPalette.margins = PALETTE_MARGINS;
    memoPalette.minimumSize = PALETTE_MIN_SIZE;
    memoPalette.preferredSize.width = PALETTE_INITIAL_WIDTH; // 初回（保存位置がないとき）の幅 / Initial width when there is no saved position
    $.global.__TextMemoWindow = memoPalette; // 二重起動防止用に参照を保持 / Keep a reference so a second launch can detect this window

    /* モード選択（置き換え / 追加）/ Mode selection (replace / append) */
    var modeSelectGroup = memoPalette.add('group');
    modeSelectGroup.orientation = 'row';
    modeSelectGroup.alignment = ['center', 'top'];      // 左右中央 / horizontally centered
    modeSelectGroup.alignChildren = ['left', 'center']; // 天地中央 / vertically centered
    modeSelectGroup.add('statictext', undefined, getLabel(LABELS.fieldLabel.mode));
    var appendModeRadio = modeSelectGroup.add('radiobutton', undefined, getLabel(LABELS.radio.append));
    var replaceModeRadio = modeSelectGroup.add('radiobutton', undefined, getLabel(LABELS.radio.replace));
    appendModeRadio.value = true; // デフォルトは追加 / Append by default

    /* 読み込み行（テキスト欄の上）/ Load row (above the text area) */
    var loadButtonRow = memoPalette.add('group');
    loadButtonRow.orientation = 'row';
    loadButtonRow.alignment = ['center', 'top'];
    loadButtonRow.alignChildren = ['center', 'center'];

    var loadButton = loadButtonRow.add('button', undefined, getLabel(LABELS.button.load));
    var removeBlanksButton = loadButtonRow.add('button', undefined, getLabel(LABELS.button.removeBlanks));
    var removeBreaksButton = loadButtonRow.add('button', undefined, getLabel(LABELS.button.removeBreaks));
    loadButton.preferredSize = BUTTON_SIZE;
    removeBlanksButton.preferredSize.height = BUTTON_HEIGHT; // 幅はラベルに合わせて自動 / Auto width to fit the label
    removeBreaksButton.preferredSize.height = BUTTON_HEIGHT; // 幅はラベルに合わせて自動 / Auto width to fit the label
    loadButton.helpTip = getLabel(LABELS.tooltip.load);
    removeBlanksButton.helpTip = getLabel(LABELS.tooltip.removeBlanks);
    removeBreaksButton.helpTip = getLabel(LABELS.tooltip.removeBreaks);

    /* メモ入力テキストエリア / Memo text area */
    var memoTextArea = memoPalette.add('edittext', undefined, restoredMemoText, {
        multiline: true,
        scrolling: true
    });
    memoTextArea.minimumSize.width = MEMO_AREA_MIN_SIZE[0];
    memoTextArea.minimumSize.height = MEMO_AREA_MIN_SIZE[1];
    memoTextArea.alignment = ['fill', 'fill']; // 上下リサイズ対応 / Resize vertically

    /* 保存・コピー・クリア行（テキスト欄の下、3カラム）/ Save / Copy / Clear row (below the text area, 3 columns) */
    var bottomButtonRow = memoPalette.add('group');
    bottomButtonRow.orientation = 'row';
    bottomButtonRow.alignment = ['fill', 'top']; // 行を幅いっぱいに広げて左右に振り分ける / Span full width to push columns left & right
    bottomButtonRow.alignChildren = ['fill', 'center'];

    /* 左カラム：保存・すべてをコピー / Left column: Save & Copy All */
    var saveCopyGroup = bottomButtonRow.add('group');
    saveCopyGroup.orientation = 'row';
    saveCopyGroup.alignment = ['left', 'center'];
    var saveButton = saveCopyGroup.add('button', undefined, getLabel(LABELS.button.save));
    saveButton.preferredSize = BUTTON_SIZE;
    saveButton.helpTip = getLabel(LABELS.tooltip.save);
    var copyAllButton = saveCopyGroup.add('button', undefined, getLabel(LABELS.button.copyAll));
    copyAllButton.preferredSize = WIDE_BUTTON_SIZE;
    copyAllButton.helpTip = getLabel(LABELS.tooltip.copyAll);

    /* 中央カラム：スペーサー（余白を吸収して左右を振り分ける）/ Center column: spacer that absorbs slack */
    var bottomSpacerGroup = bottomButtonRow.add('group');
    bottomSpacerGroup.alignment = ['fill', 'center'];

    /* 右カラム：クリア / Right column: Clear */
    var clearButtonGroup = bottomButtonRow.add('group');
    clearButtonGroup.orientation = 'row';
    clearButtonGroup.alignment = ['right', 'center'];
    var clearButton = clearButtonGroup.add('button', undefined, getLabel(LABELS.button.clear));
    clearButton.preferredSize = BUTTON_SIZE;
    clearButton.helpTip = getLabel(LABELS.tooltip.clear);

    /* レイアウトとリサイズ / Layout and resize */
    memoPalette.layout.layout(true);
    memoPalette.onResize = function () {
        memoPalette.layout.resize();
        memoPalette.layout.layout(true);
    };

    /* 表示時に入力欄をアクティブ化 / Focus the field on show */
    memoPalette.onShow = function () {
        memoTextArea.active = true;
    };

    /* ウィンドウ移動時に位置を自動保存 / Auto-save position on move */
    memoPalette.onMove = function () {
        storePaletteBounds();
    };

    /**
     * パレットの位置・サイズをセッションに保存する
     * @returns {void}
     */
    function storePaletteBounds() {
        $.global.__TextMemoBounds = [
            memoPalette.bounds.left, memoPalette.bounds.top,
            memoPalette.bounds.right, memoPalette.bounds.bottom
        ];
    }

    // =========================================
    // ボタンアクション / Button actions
    // =========================================
    /* 選択オブジェクト（無ければクリップボード）のテキストを読み込む / Load text from the selection, falling back to the clipboard */
    loadButton.onClick = function () {
        fetchSelectedText(function (status, loadedText) {
            if (status === 'nosel' || status === 'nodoc') {
                fetchClipboardText(function (clipboardStatus, clipboardText) {
                    if (clipboardStatus === 'nodoc') {
                        alert(getLabel(LABELS.alert.noDocument));
                        return;
                    }
                    if (clipboardStatus !== 'ok' || !clipboardText) {
                        alert(getLabel(LABELS.alert.clipboardEmpty));
                        return;
                    }
                    applyLoadedText(clipboardText);
                });
                return;
            }
            if (status === 'notextframe') {
                alert(getLabel(LABELS.alert.noTextFrame));
                return;
            }
            if (status !== 'ok') {
                alert(getLabel(LABELS.alert.loadFailed) + loadedText);
                return;
            }
            applyLoadedText(loadedText);
        });
    };

    /* メモをテキストファイルへ保存（パレットは開いたまま）/ Save the memo to a text file (the palette stays open) */
    saveButton.onClick = function () {
        var saveFile = resolveSaveFile();
        if (!saveFile) return;
        if (!writeTextFile(saveFile, buildSaveContent(memoTextArea.text))) {
            alert(getLabel(LABELS.alert.writeFailed));
            return;
        }
        // 保存後もパレットは開いたまま：内容・ウィンドウ位置をそのまま保持する（再起動しない）
        // Keep the palette open after saving: preserve the content and window position as-is (no restart)
        $.global.__TextMemoContent = memoTextArea.text; // フッターは含めず入力内容のみ保持 / Keep the typed memo only (without footer)
        storePaletteBounds();
    };

    /* メモをクリアして再起動 / Clear the memo, then restart */
    clearButton.onClick = function () {
        $.global.__TextMemoContent = '';
        storePaletteBounds();
        restartScript();
    };

    /* すべてをコピー：テキスト欄の内容をクリップボードへ / Copy All: copy the memo to the clipboard */
    copyAllButton.onClick = function () {
        if (!memoTextArea.text) return;
        copyTextToClipboard(memoTextArea.text, function (status) {
            if (status !== 'ok') alert(getLabel(LABELS.alert.copyFailed));
        });
    };

    /* 空行削除：テキスト欄の空行を除去 / Remove Blanks: strip blank lines from the memo */
    removeBlanksButton.onClick = function () {
        setMemoText(removeBlankLines(memoTextArea.text));
    };

    /* 改行削除：テキスト欄の改行をすべて除去して1行にまとめる / Remove Breaks: strip all line breaks from the memo */
    removeBreaksButton.onClick = function () {
        setMemoText(removeLineBreaks(memoTextArea.text));
    };

    memoTextArea.onChanging = updateButtonState; // 入力中もリアルタイムに反映 / Update live while typing
    updateButtonState();

    /**
     * テキスト欄の内容を差し替え、セッションとボタンの状態を更新する
     * @param {string} newText - テキスト欄に入れる文字列
     * @returns {void}
     */
    function setMemoText(newText) {
        memoTextArea.text = newText;
        $.global.__TextMemoContent = newText;
        updateButtonState(); // プログラム的な変更では onChanging が発火しないため明示 / onChanging does not fire on programmatic changes
    }

    /**
     * 読み込んだテキストをモードに応じてテキスト欄へ反映する（濁点を NFC へ補正）
     * @param {string} loadedText - 読み込んだテキスト
     * @returns {void}
     */
    function applyLoadedText(loadedText) {
        var normalizedText = normalizeDakuten(loadedText);
        if (appendModeRadio.value) {
            // 既存テキストとの間に空行を1つ挟む / Insert one blank line between existing and new text
            setMemoText(memoTextArea.text + (memoTextArea.text ? '\n\n' : '') + normalizedText);
        } else {
            setMemoText(normalizedText);
        }
    }

    /**
     * テキストの有無・空行の有無でボタンの活性を切り替える
     * @returns {void}
     */
    function updateButtonState() {
        var memoText = memoTextArea.text;
        var hasText = (memoText.length > 0);
        clearButton.enabled = hasText;
        copyAllButton.enabled = hasText;
        removeBlanksButton.enabled = hasText && hasBlankLines(memoText); // 空行が無ければディム / Dim when there are no blank lines
        removeBreaksButton.enabled = hasText && hasLineBreaks(memoText); // 改行が無ければディム / Dim when there are no line breaks
    }

    // =========================================
    // キー操作 & 閉じる / Keys & close
    // =========================================
    /* Escape でパレットを閉じる / Close the palette with Escape */
    memoPalette.addEventListener('keydown', function (e) {
        if (e.keyName === 'Escape') memoPalette.close();
    });

    /* 閉じたときの内容を CLEAR_ON_CLOSE で制御（再起動時は触れない）。ウィンドウ位置は常に保持 /
       On close, control the content via CLEAR_ON_CLOSE (untouched on restart); always keep the window position */
    memoPalette.onClose = function () {
        if (!isRestartingScript) {
            $.global.__TextMemoContent = CLEAR_ON_CLOSE ? '' : memoTextArea.text;
        }
        storePaletteBounds();
        $.global.__TextMemoWindow = null; // 参照を解放（次回起動で新規作成できるように）/ Release the reference so the next launch creates a fresh window
        return true;
    };

    /* 表示（既存位置がない場合のみ中央）/ Show the palette (center only when no saved position) */
    if (!initialPaletteBounds) memoPalette.center();
    memoPalette.show();

    // =========================================
    // テキスト整形 / Text utilities
    // =========================================
    /**
     * 改行を LF に正規化して行の配列へ分割する
     * @param {string} sourceText - 対象のテキスト
     * @returns {string[]} 行の配列
     */
    function splitLines(sourceText) {
        return sourceText.replace(/\r\n|\r/g, '\n').split('\n');
    }

    /**
     * 空行（空白のみの行を含む）かどうかを判定する
     * @param {string} line - 1行分のテキスト
     * @returns {boolean} 空行なら true
     */
    function isBlankLine(line) {
        return line.replace(/^\s+|\s+$/g, '') === '';
    }

    /**
     * 空行（空白のみの行を含む）を除去する
     * @param {string} sourceText - 対象のテキスト
     * @returns {string} 空行を除いたテキスト
     */
    function removeBlankLines(sourceText) {
        var lines = splitLines(sourceText);
        var keptLines = [];
        for (var i = 0; i < lines.length; i++) {
            if (!isBlankLine(lines[i])) keptLines.push(lines[i]);
        }
        return keptLines.join('\n');
    }

    /**
     * 空行が1つでもあるか判定する
     * @param {string} sourceText - 対象のテキスト
     * @returns {boolean} 空行があれば true
     */
    function hasBlankLines(sourceText) {
        var lines = splitLines(sourceText);
        for (var i = 0; i < lines.length; i++) {
            if (isBlankLine(lines[i])) return true;
        }
        return false;
    }

    /**
     * 改行をすべて除去して1行にまとめる
     * @param {string} sourceText - 対象のテキスト
     * @returns {string} 1行にまとめたテキスト
     */
    function removeLineBreaks(sourceText) {
        return sourceText.replace(/\r\n|\r|\n/g, '');
    }

    /**
     * 改行が1つでもあるか判定する
     * @param {string} sourceText - 対象のテキスト
     * @returns {boolean} 改行があれば true
     */
    function hasLineBreaks(sourceText) {
        return /\r|\n/.test(sourceText);
    }

    /**
     * 分解済みの濁点・半濁点（NFD）を合成形（NFC）へ補正する
     * macOS のファイル名やクリップボード由来の「か゛」を「が」に直す
     * @param {string} rawText - 補正前のテキスト
     * @returns {string} 補正後のテキスト
     */
    function normalizeDakuten(rawText) {
        return decodeURIComponent(composeDakutenOnEncoded(encodeURIComponent(rawText)));
    }

    /**
     * パーセントエンコード列上で分解済み濁点を合成済みへ置換する
     * @param {string} encodedText - encodeURIComponent 済みのテキスト
     * @returns {string} 置換後のエンコード列
     */
    function composeDakutenOnEncoded(encodedText) {
        var decomposedForms = ['%E3%81%86%E3%82%99', '%E3%81%8B%E3%82%99', '%E3%81%8D%E3%82%99', '%E3%81%8F%E3%82%99', '%E3%81%91%E3%82%99', '%E3%81%93%E3%82%99',
            '%E3%81%95%E3%82%99', '%E3%81%97%E3%82%99', '%E3%81%99%E3%82%99', '%E3%81%9B%E3%82%99', '%E3%81%9D%E3%82%99', '%E3%81%9F%E3%82%99',
            '%E3%81%A1%E3%82%99', '%E3%81%A4%E3%82%99', '%E3%81%A6%E3%82%99', '%E3%81%A8%E3%82%99', '%E3%81%AF%E3%82%99', '%E3%81%AF%E3%82%9A',
            '%E3%81%B2%E3%82%99', '%E3%81%B2%E3%82%9A', '%E3%81%B5%E3%82%99', '%E3%81%B5%E3%82%9A', '%E3%81%B8%E3%82%99', '%E3%81%B8%E3%82%9A',
            '%E3%81%BB%E3%82%99', '%E3%81%BB%E3%82%9A', '%E3%82%A6%E3%82%99', '%E3%82%AB%E3%82%99', '%E3%82%AD%E3%82%99', '%E3%82%AF%E3%82%99',
            '%E3%82%B1%E3%82%99', '%E3%82%B3%E3%82%99', '%E3%82%B5%E3%82%99', '%E3%82%B7%E3%82%99', '%E3%82%B9%E3%82%99', '%E3%82%BB%E3%82%99',
            '%E3%82%BD%E3%82%99', '%E3%82%BF%E3%82%99', '%E3%83%81%E3%82%99', '%E3%83%84%E3%82%99', '%E3%83%86%E3%82%99', '%E3%83%88%E3%82%99',
            '%E3%83%8F%E3%82%99', '%E3%83%8F%E3%82%9A', '%E3%83%92%E3%82%99', '%E3%83%92%E3%82%9A', '%E3%83%95%E3%82%99', '%E3%83%95%E3%82%9A',
            '%E3%83%98%E3%82%99', '%E3%83%98%E3%82%9A', '%E3%83%9B%E3%82%99', '%E3%83%9B%E3%82%9A'
        ];
        var composedForms = ['%E3%82%94', '%E3%81%8C', '%E3%81%8E', '%E3%81%90', '%E3%81%92', '%E3%81%94', '%E3%81%96', '%E3%81%98', '%E3%81%9A', '%E3%81%9C', '%E3%81%9E',
            '%E3%81%A0', '%E3%81%A2', '%E3%81%A5', '%E3%81%A7', '%E3%81%A9', '%E3%81%B0', '%E3%81%B1', '%E3%81%B3', '%E3%81%B4', '%E3%81%B6', '%E3%81%B7',
            '%E3%81%B9', '%E3%81%BA', '%E3%81%BC', '%E3%81%BD', '%E3%83%B4', '%E3%82%AC', '%E3%82%AE', '%E3%82%B0', '%E3%82%B2', '%E3%82%B4', '%E3%82%B6',
            '%E3%82%B8', '%E3%82%BA', '%E3%82%BC', '%E3%82%BE', '%E3%83%80', '%E3%83%82', '%E3%83%85', '%E3%83%87', '%E3%83%89', '%E3%83%90', '%E3%83%91',
            '%E3%83%93', '%E3%83%94', '%E3%83%96', '%E3%83%97'
        ];
        for (var i = 0; i < decomposedForms.length; i++) {
            encodedText = encodedText.replace(new RegExp(decomposedForms[i], 'ig'), composedForms[i]);
        }
        return encodedText;
    }

    // =========================================
    // 保存 / Save
    // =========================================
    /**
     * SAVE_LOCATION_MODE に従って保存先ファイルを決める
     * @returns {File|null} 保存先ファイル（保存ダイアログでキャンセルされたときは null）
     */
    function resolveSaveFile() {
        if (SAVE_LOCATION_MODE === 'desktop') {
            // B: 常にデスクトップへ memo-<ドキュメント名>-<yyyymmdd>.txt（日本語名対応で encodeURI）
            var fileName = 'memo-' + getDocumentBaseName() + '-' + getTimeStamp().substring(0, 8) + '.txt';
            return new File(Folder.desktop + '/' + encodeURI(fileName));
        }
        // A: 保存ダイアログで場所と名前を選ぶ
        var suggestedName = getDocumentBaseName() + '_' + getTimeStamp() + '.txt';
        var chosenFile = File.saveDialog(getLabel(LABELS.alert.savePrompt), suggestedName);
        if (!chosenFile) return null;
        return /\.txt$/i.test(chosenFile.name) ? chosenFile : new File(chosenFile.fsName + '.txt');
    }

    /**
     * 保存本文（フッター付き、改行を LF に正規化）を組み立てる
     * @param {string} memoText - テキスト欄の内容
     * @returns {string} ファイルへ書き込む文字列
     */
    function buildSaveContent(memoText) {
        var bodyText = memoText.replace(/\r\n|\r/g, '\n');
        var footerTimeStamp = getTimeStamp().replace(/(\d{4})(\d{2})(\d{2})_(\d{2})(\d{2})(\d{2})/, '$1-$2-$3 $4:$5:$6');
        return bodyText + '\n\n---\n[' + getLabel(LABELS.footer.source) + ': ' + sourceDocumentName + ']\n[' +
            getLabel(LABELS.footer.savedAt) + ': ' + footerTimeStamp + ']\n';
    }

    /**
     * UTF-8 でテキストファイルへ書き込む
     * @param {File} targetFile - 書き込み先ファイル
     * @param {string} fileContent - 書き込む文字列
     * @returns {boolean} 書き込めたら true
     */
    function writeTextFile(targetFile, fileContent) {
        targetFile.encoding = 'UTF-8';
        if (!targetFile.open('w')) return false;
        var writeSucceeded = targetFile.write(fileContent);
        targetFile.close();
        return writeSucceeded;
    }

    /**
     * タイムスタンプ YYYYMMDD_HHMMSS を返す
     * @returns {string} タイムスタンプ文字列
     */
    function getTimeStamp() {
        var now = new Date();

        /**
         * 2桁になるよう 0 を補う
         * @param {number} value - 対象の数値
         * @returns {string} 2桁の文字列
         */
        function padZero(value) {
            return ('0' + value).slice(-2);
        }
        return now.getFullYear() + padZero(now.getMonth() + 1) + padZero(now.getDate()) + '_' +
            padZero(now.getHours()) + padZero(now.getMinutes()) + padZero(now.getSeconds());
    }

    // =========================================
    // ドキュメント / Document helpers
    // =========================================
    /**
     * アクティブドキュメントを安全に取得する（起動直後・パレット表示前向け）
     * パレット表示中はこのエンジンの app が DOM 接続を失うため、DOM に触る処理は
     * fetchSelectedText などメインエンジン側へ委譲すること
     * @returns {Document|null} アクティブドキュメント（取得できなければ null）
     */
    function getActiveDocument() {
        try {
            return (app.documents.length > 0) ? app.activeDocument : null;
        } catch (err) {
            return null;
        }
    }

    /**
     * 保存ファイルの基底名（拡張子なし）を返す
     * パレット表示中は DOM 取得が失敗しうるため、起動時に控えた名前を使う
     * @returns {string} ファイルの基底名
     */
    function getDocumentBaseName() {
        if (sourceDocumentName === getLabel(LABELS.status.unsaved)) return 'Untitled';
        return sourceDocumentName.replace(/\.[^\.]+$/, '');
    }

    /**
     * スクリプト自身をメインエンジンで再評価して再起動する
     * @returns {void}
     */
    function restartScript() {
        var scriptFile = new File($.fileName);
        var bridgeMessage = new BridgeTalk();
        bridgeMessage.target = 'illustrator';
        bridgeMessage.body = '$.evalFile("' + scriptFile.fsName.replace(/\\/g, '\\\\') + '");';
        bridgeMessage.send(100);
        isRestartingScript = true; // この close では内容を消さない / This close must not clear the content
        memoPalette.close();
    }

})();
