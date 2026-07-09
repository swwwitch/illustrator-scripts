#targetengine "TextMemoEngine"
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

Illustrator 用のメモ入力フローティングパレット。

#### 読み込み

- 選択オブジェクトからテキストを収集（**追加**＝既定 / **置き換え**）。追加時は既存テキストとの間に空行を1つ挟む
- グループ・クリップグループ・シンボル内（入れ子も多段展開）のテキストも対象
- 選択が無いときはクリップボードを一時的に貼り付け（`pasteInAllArtboard`）→テキスト収集→貼り付けたオブジェクトを削除、の手順で読み込む（アートボードごとの複製は重複除去）
- 複数取得時はカンバス上の位置で上から順（同じ高さは左から右）に整列。実質的に空のフレームは無視
- 取り込み時に濁点・半濁点を NFC へ補正
- 各ボタンにはツールチップ（helpTip）を表示

#### 編集・コピー

- 「空行削除」でテキスト欄の空行（空白のみの行を含む）を除去
- 「改行削除」でテキスト欄の改行をすべて除去し1行にまとめる
- 「すべてをコピー」でテキスト欄の内容をシステムクリップボードへコピー（一時テキストフレーム + `app.copy()` を BridgeTalk 実行）
- テキストが空／対象が無いときは該当ボタンを自動でディム（空行が無ければ「空行削除」、改行が無ければ「改行削除」もディム）

#### 保存・その他

- 入力したメモを UTF-8 のテキストファイル（.txt）へ保存（保存元・保存日時のフッター付き）
- 保存先は `SAVE_LOCATION_MODE` で切替：保存ダイアログで選択（A）/ 常にデスクトップへ `memo-<ドキュメント名>-<yyyymmdd>.txt`（B＝既定）
- 保存後はパレットを閉じず、入力内容とウィンドウ位置をそのまま保持（再起動しない）
- クリア後はスクリプトを再起動して空の状態で起動し直す
- ボタンエリアはテキスト欄の上＝読み込み系（読み込む／空行削除／改行削除）、下＝3カラム（左:保存・すべてをコピー／中央:スペーサー／右:クリア）
- 閉じたときの挙動は `CLEAR_ON_CLOSE` で切替：内容を保持（既定）/ クリア。ウィンドウ位置は常に保持
- UI・メッセージ・保存フッターまで日本語／英語にローカライズ（ロケール自動判定）

#### 補足

選択テキストの取得は、生きた DOM を持つメインエンジンへ BridgeTalk で問い合わせて行う
（常駐エンジンの app はパレット表示中に DOM 接続を失い `there is no document` を投げるため）。
シンボル内テキストは一時レイヤーへ複製→breakLink で展開して読み取り、複製は削除して元のシンボルには触れない。
Illustrator には `app.system` が無いため、クリップボードへのコピーは一時テキストフレーム + `app.copy()` で行う。
逆方向（クリップボードの取り込み）も `app.system` で直接読めないため、ドキュメントへ一時的に貼り付けてからテキストを読み取り、貼り付けたオブジェクトは削除する（元の選択は復元）。

### 解説

https://note.com/dtp_tranist/n/n41e91e4b1a09

### オリジナル

こじらせたクマーさんの以下の記事をもとに、機能追加やリファクタリングを行いました。
https://note.com/nice_lotus120/n/n6291a432b30d

*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.1.1";

// =========================================
// ユーザー設定 / User settings
// =========================================
/* パレットの透明度 / Palette opacity */
var DIALOG_OPACITY = 0.97;

/* 保存先モード / Save destination mode
   'dialog'  = 保存ダイアログで場所と名前を選ぶ（A）/ Choose location & name via the save dialog (A)
   'desktop' = 常にデスクトップへ memo-<ドキュメント名><yyyymmdd>.txt（B）/ Always save to the Desktop (B) */
var SAVE_LOCATION_MODE = 'desktop'; // デフォルトは B / Default is B

/* パレットを閉じたときに内容をクリアするか / Whether to clear the content when the palette is closed
   true  = 閉じると内容をクリア / Clear on close
   false = 内容を保持して次回復元（デフォルト）/ Keep content and restore next time (default) */
var CLEAR_ON_CLOSE = false;

// =========================================
// ローカライズ / Localization
// =========================================
/* 現在の言語を判定（ロケールが ja 始まりなら日本語）/ Detect UI language (Japanese if locale starts with "ja") */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var currentLanguage = getCurrentLang();

var LABELS = {
    dialog: {
        title: { ja: "テキスト一時保管", en: "Text Stash" }
    },
    field: {
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
    message: {
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

/* ラベルノードから現在言語の文言を返す（{slash}→/）/ Resolve a label node to the current language ({slash}→/) */
function L(labelNode) {
    if (!labelNode) return "";
    var text = labelNode[currentLanguage] || labelNode.en || "";
    return text.replace(/\{slash\}/g, "/");
}

(function () {

    // =========================================
    // 二重起動防止 / Prevent duplicate launch
    // =========================================
    /* 既にパレットが開いていれば前面に出して終了する（常駐エンジンなので二重に生成しない）/
       If the palette is already open, bring it to the front and stop (avoid a second instance on this persistent engine) */
    if ($.global.__TextMemoWindow) {
        try {
            var existingWindow = $.global.__TextMemoWindow;
            if (existingWindow && existingWindow.visible) {
                existingWindow.active = true; // 前面へ / Bring to front
                return;
            }
        } catch (duplicateLaunchCheckError) {
            // 参照が失効していれば通常どおり新規作成へ進む / If the reference is stale, fall through and create a new window
        }
    }

    // =========================================
    // セッション保持用変数 / Session state
    // =========================================
    var savedText = $.global.__TextMemoContent || '';
    var savedBounds = $.global.__TextMemoBounds || null;

    /* アクティブドキュメント名（保存フッター用）/ Active document name (used in the save footer) */
    var activeDocName = L(LABELS.status.unsaved);
    var initialDoc = getActiveDocument();
    if (initialDoc) {
        activeDocName = initialDoc.fullName ? decodeURI(initialDoc.fullName.name) : initialDoc.name;
    }

    // =========================================
    // UI 構築 / Build UI
    // =========================================
    var winBounds = (savedBounds && savedBounds.length === 4) ? savedBounds : undefined;
    var win = new Window('palette', L(LABELS.dialog.title) + ' ' + SCRIPT_VERSION, winBounds, {
        resizeable: true
    });
    win.opacity = DIALOG_OPACITY;
    win.orientation = 'column';
    win.alignChildren = ['fill', 'top'];
    win.margins = [15, 15, 15, 15]; // 左右上下を対称に / Symmetric margins on all sides
    win.minimumSize = [300, 200];
    win.preferredSize.width = 400; // 初回（保存位置がないとき）の幅 / Initial width when there is no saved position
    $.global.__TextMemoWindow = win; // 二重起動防止用に参照を保持 / Keep a reference so a second launch can detect this window

    /* モード選択（置き換え / 追加）/ Mode selection (replace / append) */
    var modeSelectGroup = win.add('group');
    modeSelectGroup.orientation = 'row';
    modeSelectGroup.alignment = ['center', 'top'];      // 左右中央 / horizontally centered
    modeSelectGroup.alignChildren = ['left', 'center']; // 天地中央 / vertically centered
    modeSelectGroup.add('statictext', undefined, L(LABELS.field.mode));
    var modeAppend = modeSelectGroup.add('radiobutton', undefined, L(LABELS.radio.append));
    var modeReplace = modeSelectGroup.add('radiobutton', undefined, L(LABELS.radio.replace));
    modeAppend.value = true; // デフォルトは追加 / Append by default

    /* 読み込み行（テキスト欄の上）/ Load row (above the text area) */
    var buttonRow = win.add('group');
    buttonRow.orientation = 'row';
    buttonRow.alignment = ['center', 'top'];
    buttonRow.alignChildren = ['center', 'center'];

    var loadButton = buttonRow.add('button', undefined, L(LABELS.button.load));
    var removeBlanksButton = buttonRow.add('button', undefined, L(LABELS.button.removeBlanks));
    var removeBreaksButton = buttonRow.add('button', undefined, L(LABELS.button.removeBreaks));
    /* ボタンは少し小さめに統一 / Make the buttons a bit smaller and uniform */
    loadButton.preferredSize = [76, 24];
    removeBlanksButton.preferredSize.height = 24; // 幅はラベルに合わせて自動 / Auto width to fit the label
    removeBreaksButton.preferredSize.height = 24; // 幅はラベルに合わせて自動 / Auto width to fit the label
    /* ツールチップ / Tooltips */
    loadButton.helpTip = L(LABELS.tooltip.load);
    removeBlanksButton.helpTip = L(LABELS.tooltip.removeBlanks);
    removeBreaksButton.helpTip = L(LABELS.tooltip.removeBreaks);

    /* メモ入力テキストエリア / Memo text area */
    var memoTextArea = win.add('edittext', undefined, savedText, {
        multiline: true,
        scrolling: true
    });
    memoTextArea.minimumSize.height = 100;
    memoTextArea.minimumSize.width = 300;
    memoTextArea.alignment = ['fill', 'fill']; // 上下リサイズ対応 / Resize vertically

    /* 保存・コピー・クリア行（テキスト欄の下、3カラム）/ Save / Copy / Clear row (below the text area, 3 columns) */
    var saveButtonRow = win.add('group');
    saveButtonRow.orientation = 'row';
    saveButtonRow.alignment = ['fill', 'top']; // 行を幅いっぱいに広げて左右に振り分ける / Span full width to push columns left & right
    saveButtonRow.alignChildren = ['fill', 'center'];

    /* 左カラム：保存・すべてをコピー / Left column: Save & Copy All */
    var leftButtonGroup = saveButtonRow.add('group');
    leftButtonGroup.orientation = 'row';
    leftButtonGroup.alignment = ['left', 'center'];
    var saveButton = leftButtonGroup.add('button', undefined, L(LABELS.button.save));
    saveButton.preferredSize = [76, 24];
    saveButton.helpTip = L(LABELS.tooltip.save);
    var copyAllButton = leftButtonGroup.add('button', undefined, L(LABELS.button.copyAll));
    copyAllButton.preferredSize = [120, 24];
    copyAllButton.helpTip = L(LABELS.tooltip.copyAll);

    /* 中央カラム：スペーサー（余白を吸収して左右を振り分ける）/ Center column: spacer that absorbs slack */
    var centerSpacer = saveButtonRow.add('group');
    centerSpacer.alignment = ['fill', 'center'];

    /* 右カラム：クリア / Right column: Clear */
    var rightButtonGroup = saveButtonRow.add('group');
    rightButtonGroup.orientation = 'row';
    rightButtonGroup.alignment = ['right', 'center'];
    var clearButton = rightButtonGroup.add('button', undefined, L(LABELS.button.clear));
    clearButton.preferredSize = [76, 24];
    clearButton.helpTip = L(LABELS.tooltip.clear);

    /* レイアウトとリサイズ / Layout and resize */
    win.layout.layout(true);
    win.onResize = function () {
        win.layout.resize();
        win.layout.layout(true);
    };

    /* 表示時に入力欄をアクティブ化 / Focus the field on show */
    win.onShow = function () {
        memoTextArea.active = true;
    };

    /* ウィンドウ移動時に位置を自動保存 / Auto-save position on move */
    win.onMove = function () {
        storeWinBounds();
    };

    /* ウィンドウ位置をセッションに保存 / Persist the window bounds into the session */
    function storeWinBounds() {
        $.global.__TextMemoBounds = [
            win.bounds.left, win.bounds.top,
            win.bounds.right, win.bounds.bottom
        ];
    }

    // =========================================
    // ボタンアクション / Button actions
    // =========================================
    /* 選択テキストフレームを読み込んでテキストエリアへ反映 / Load selected text frames into the memo area */
    loadButton.onClick = function () {
        // 選択テキストの取得は DOM 接続を保つメインエンジンに委譲（非同期）
        fetchSelectedTextFrames(function (status, loadedText) {
            // 選択オブジェクトが無ければクリップボードを貼り付けて読み込む
            if (status === 'nosel' || status === 'nodoc') {
                pasteAndReadClipboardText(function (pasteStatus, clipboardText) {
                    if (pasteStatus === 'nodoc') {
                        alert(L(LABELS.message.noDocument));
                        return;
                    }
                    if (pasteStatus !== 'ok' || !clipboardText) {
                        alert(L(LABELS.message.clipboardEmpty));
                        return;
                    }
                    applyLoadedText(clipboardText);
                });
                return;
            }
            if (status === 'notextframe') {
                alert(L(LABELS.message.noTextFrame));
                return;
            }
            if (status === 'error') {
                alert(L(LABELS.message.loadFailed) + loadedText);
                return;
            }
            applyLoadedText(loadedText);
        });
    };

    /* 読み込んだテキストをモードに応じてテキストエリアへ反映（濁点を NFC へ補正）/ Apply loaded text to the memo area per mode (normalize dakuten to NFC) */
    function applyLoadedText(text) {
        var normalized = fixDakuten(text);
        if (modeAppend.value) {
            // 既存テキストとの間に空行を1つ挟む / Insert one blank line between existing and new text
            memoTextArea.text += (memoTextArea.text ? "\n\n" : "") + normalized;
        } else {
            memoTextArea.text = normalized;
        }
        $.global.__TextMemoContent = memoTextArea.text;
        updateButtonState(); // プログラム的な変更では onChanging が発火しないため明示 / onChanging does not fire on programmatic changes
    }

    /* 空行（空白のみの行を含む）を除去 / Remove blank lines (including whitespace-only lines) */
    function removeBlankLines(text) {
        var lines = text.replace(/\r\n|\r/g, '\n').split('\n');
        var kept = [];
        for (var i = 0; i < lines.length; i++) {
            if (lines[i].replace(/^\s+|\s+$/g, '') !== '') kept.push(lines[i]);
        }
        return kept.join('\n');
    }

    /* 空行が1つでもあるか判定 / Check whether there is at least one blank line */
    function hasBlankLines(text) {
        var lines = text.replace(/\r\n|\r/g, '\n').split('\n');
        for (var i = 0; i < lines.length; i++) {
            if (lines[i].replace(/^\s+|\s+$/g, '') === '') return true;
        }
        return false;
    }

    /* 改行をすべて除去して1行にまとめる / Remove all line breaks, joining everything into one line */
    function removeLineBreaks(text) {
        return text.replace(/\r\n|\r|\n/g, '');
    }

    /* 改行が1つでもあるか判定 / Check whether there is at least one line break */
    function hasLineBreaks(text) {
        return /\r|\n/.test(text);
    }

    /* クリップボードを貼り付けてテキストを読み取る / Paste the clipboard and read its text */
    /* 選択オブジェクトが無いときのフォールバック。Illustrator には app.system が無く
       システムクリップボードを直接読めないため、いったんドキュメントへ貼り付け
       （pasteInAllArtboard）→ 貼り付いたオブジェクトのテキストを収集 → そのオブジェクトを削除、
       という手順で取り出す。パレット常駐エンジンは表示中に DOM が切れるため、
       生きた DOM を持つメインエンジンへ BridgeTalk で委譲（非同期）。
       pasteInAllArtboard はアートボードごとに複製を作るため、同一テキストは重複除去する。
       テキストは encodeURIComponent で安全に受け渡す。
       コールバック: onDone(status, text)  status = 'ok' | 'nodoc' | 'notext' | 'error' */
    function pasteAndReadClipboardText(onDone) {
        var pasteSource =
            '(function () {' +
            '    if (app.documents.length === 0) return "NODOC:";' +
            '    var targetDoc;' +
            '    try { targetDoc = app.activeDocument; } catch (e0) { targetDoc = app.documents[0]; }' +
            // 貼り付け前の選択を控えて解除
            '    var prevSel = [];' +
            '    try { for (var i = 0; i < targetDoc.selection.length; i++) prevSel.push(targetDoc.selection[i]); } catch (e1) {}' +
            '    try { targetDoc.selection = null; } catch (e2) {}' +
            // クリップボードを貼り付け。直後の選択が貼り付いたオブジェクト
            '    var pasted = [];' +
            '    try {' +
            '        app.executeMenuCommand("pasteInAllArtboard");' +
            '        for (var p = 0; p < targetDoc.selection.length; p++) pasted.push(targetDoc.selection[p]);' +
            '    } catch (e3) {}' +
            // テキストと位置（geometricBounds の上端・左端）を収集。実質的に空のフレームは無視
            '    var collected = [];' +
            '    function pushFrame(frame) {' +
            '        var contents = frame.contents;' +
            '        if (!contents || contents.replace(/[\\r\\n\\x03]/g, "").replace(/\\s+/g, "") === "") return;' +
            '        var topY = 0, leftX = 0;' +
            '        try { var bounds = frame.geometricBounds; leftX = bounds[0]; topY = bounds[1]; } catch (eg) {}' +
            '        collected.push({ text: contents, top: topY, left: leftX });' +
            '    }' +
            '    function collectFrom(item) {' +
            '        if (item.typename === "TextFrame") pushFrame(item);' +
            '        else if (item.typename === "GroupItem") { for (var k = 0; k < item.pageItems.length; k++) collectFrom(item.pageItems[k]); }' +
            '    }' +
            '    for (var j = 0; j < pasted.length; j++) collectFrom(pasted[j]);' +
            // 貼り付けたオブジェクトを削除（元のドキュメントは汚さない）
            '    for (var d = 0; d < pasted.length; d++) { try { pasted[d].remove(); } catch (e4) {} }' +
            // 元の選択を復元
            '    try { targetDoc.selection = null; } catch (e5) {}' +
            '    for (var q = 0; q < prevSel.length; q++) { try { prevSel[q].selected = true; } catch (e6) {} }' +
            '    if (collected.length === 0) return "NOTX:";' +
            // カンバス上の位置で上から順（同じ高さは左から右）に並べる
            '    collected.sort(function (a, b) {' +
            '        if (Math.abs(b.top - a.top) <= 10) return a.left - b.left;' +
            '        return b.top - a.top;' +
            '    });' +
            // 重複テキストを除去（アートボードごとの複製を1つにまとめる）
            '    var seen = {};' +
            '    var texts = [];' +
            '    for (var x = 0; x < collected.length; x++) {' +
            '        var key = "k_" + collected[x].text;' +
            '        if (seen[key]) continue;' +
            '        seen[key] = true;' +
            '        texts.push(collected[x].text);' +
            '    }' +
            '    return "OK:" + encodeURIComponent(texts.join(String.fromCharCode(10)));' +
            '})();';

        var bridge = new BridgeTalk();
        bridge.target = 'illustrator';
        bridge.body = pasteSource;
        bridge.onResult = function (response) {
            var payload = response.body || '';
            var colonIndex = payload.indexOf(':');
            var marker = colonIndex >= 0 ? payload.substring(0, colonIndex) : payload;
            var rest = colonIndex >= 0 ? payload.substring(colonIndex + 1) : '';
            if (marker === 'OK') {
                onDone('ok', decodeURIComponent(rest));
            } else if (marker === 'NODOC') {
                onDone('nodoc');
            } else if (marker === 'NOTX') {
                onDone('notext');
            } else {
                onDone('error', payload);
            }
        };
        bridge.onError = function (response) {
            onDone('error', response.body);
        };
        bridge.send();
    }

    /* テキストをクリップボードへコピー / Copy text to the clipboard */
    /* Illustrator には app.system が無いため、一時テキストフレーム + app.copy() でコピーする。
       パレット常駐エンジンは表示中に DOM が切れるため、メインエンジンへ BridgeTalk で委譲（非同期）。
       ドキュメントが開いていればその上で（画面外に一瞬作って削除）、無ければ一時ドキュメントで実行する。
       アクティブレイヤーがロック/非表示だと textFrames.add() が 8705（Target layer cannot be modified）で
       失敗するため、一時的にロック解除＋表示にしてからフレームを作り、処理後に元の状態へ戻す。
       テキストは encodeURIComponent で安全に受け渡す。onDone(status): 'ok' | 'error' */
    function copyTextToClipboard(text, onDone) {
        var copyScript =
            '(function () {' +
            '    var memoText = decodeURIComponent("' + encodeURIComponent(text) + '");' +
            '    var usingTempDoc = (app.documents.length === 0);' +
            '    var targetDoc;' +
            '    if (usingTempDoc) {' +
            '        targetDoc = app.documents.add();' +
            '    } else {' +
            '        try { targetDoc = app.activeDocument; } catch (e0) { targetDoc = app.documents[0]; }' +
            '    }' +
            '    var previousSelection = [];' +
            '    if (!usingTempDoc) { try { for (var i = 0; i < targetDoc.selection.length; i++) previousSelection.push(targetDoc.selection[i]); } catch (e1) {} }' +
            '    var tempFrame = null;' +
            '    var editLayer = null, savedLocked = false, savedVisible = true;' +
            '    try { editLayer = targetDoc.activeLayer; savedLocked = editLayer.locked; savedVisible = editLayer.visible; } catch (eL) { editLayer = null; }' +
            '    try {' +
            '        if (editLayer) { if (editLayer.locked) editLayer.locked = false; if (!editLayer.visible) editLayer.visible = true; }' +
            '        tempFrame = (editLayer ? editLayer.textFrames.add() : targetDoc.textFrames.add());' +
            '        tempFrame.contents = memoText;' +
            '        try { targetDoc.selection = null; } catch (e2) {}' +
            '        tempFrame.selected = true;' +
            '        app.copy();' +
            '    } catch (e3) {' +
            '        if (tempFrame) { try { tempFrame.remove(); } catch (e4) {} }' +
            '        if (editLayer) { try { editLayer.locked = savedLocked; } catch (eR1) {} try { editLayer.visible = savedVisible; } catch (eR2) {} }' +
            '        if (usingTempDoc) { try { targetDoc.close(SaveOptions.DONOTSAVECHANGES); } catch (e5) {} }' +
            '        return "ERR";' +
            '    }' +
            '    if (usingTempDoc) {' +
            '        try { targetDoc.close(SaveOptions.DONOTSAVECHANGES); } catch (e6) {}' +
            '    } else {' +
            '        try { tempFrame.remove(); } catch (e7) {}' +
            '        if (editLayer) { try { editLayer.locked = savedLocked; } catch (eR3) {} try { editLayer.visible = savedVisible; } catch (eR4) {} }' +
            '        try { targetDoc.selection = null; } catch (e8) {}' +
            '        for (var j = 0; j < previousSelection.length; j++) { try { previousSelection[j].selected = true; } catch (e9) {} }' +
            '    }' +
            '    return "OK";' +
            '})();';

        var bridge = new BridgeTalk();
        bridge.target = 'illustrator';
        bridge.body = copyScript;
        bridge.onResult = function (response) {
            onDone((response.body === 'OK') ? 'ok' : 'error');
        };
        bridge.onError = function () {
            onDone('error');
        };
        bridge.send();
    }

    /* 濁点・半濁点を結合（NFD → NFC）/ Combine dakuten / handakuten (NFD → NFC) */
    /* macOS のファイル名やクリップボード由来の分解形（か゛）を合成形（が）へ補正する */
    function fixDakuten(rawText) {
        return decodeURIComponent(combineDakuten(encodeURIComponent(rawText)));
    }

    /* パーセントエンコード列上で分解済み濁点を合成済みへ置換 / Replace decomposed dakuten with composed ones on the percent-encoded string */
    function combineDakuten(encodedText) {
        var nfd = ['%E3%81%86%E3%82%99', '%E3%81%8B%E3%82%99', '%E3%81%8D%E3%82%99', '%E3%81%8F%E3%82%99', '%E3%81%91%E3%82%99', '%E3%81%93%E3%82%99',
            '%E3%81%95%E3%82%99', '%E3%81%97%E3%82%99', '%E3%81%99%E3%82%99', '%E3%81%9B%E3%82%99', '%E3%81%9D%E3%82%99', '%E3%81%9F%E3%82%99',
            '%E3%81%A1%E3%82%99', '%E3%81%A4%E3%82%99', '%E3%81%A6%E3%82%99', '%E3%81%A8%E3%82%99', '%E3%81%AF%E3%82%99', '%E3%81%AF%E3%82%9A',
            '%E3%81%B2%E3%82%99', '%E3%81%B2%E3%82%9A', '%E3%81%B5%E3%82%99', '%E3%81%B5%E3%82%9A', '%E3%81%B8%E3%82%99', '%E3%81%B8%E3%82%9A',
            '%E3%81%BB%E3%82%99', '%E3%81%BB%E3%82%9A', '%E3%82%A6%E3%82%99', '%E3%82%AB%E3%82%99', '%E3%82%AD%E3%82%99', '%E3%82%AF%E3%82%99',
            '%E3%82%B1%E3%82%99', '%E3%82%B3%E3%82%99', '%E3%82%B5%E3%82%99', '%E3%82%B7%E3%82%99', '%E3%82%B9%E3%82%99', '%E3%82%BB%E3%82%99',
            '%E3%82%BD%E3%82%99', '%E3%82%BF%E3%82%99', '%E3%83%81%E3%82%99', '%E3%83%84%E3%82%99', '%E3%83%86%E3%82%99', '%E3%83%88%E3%82%99',
            '%E3%83%8F%E3%82%99', '%E3%83%8F%E3%82%9A', '%E3%83%92%E3%82%99', '%E3%83%92%E3%82%9A', '%E3%83%95%E3%82%99', '%E3%83%95%E3%82%9A',
            '%E3%83%98%E3%82%99', '%E3%83%98%E3%82%9A', '%E3%83%9B%E3%82%99', '%E3%83%9B%E3%82%9A'
        ];
        var nfc = ['%E3%82%94', '%E3%81%8C', '%E3%81%8E', '%E3%81%90', '%E3%81%92', '%E3%81%94', '%E3%81%96', '%E3%81%98', '%E3%81%9A', '%E3%81%9C', '%E3%81%9E',
            '%E3%81%A0', '%E3%81%A2', '%E3%81%A5', '%E3%81%A7', '%E3%81%A9', '%E3%81%B0', '%E3%81%B1', '%E3%81%B3', '%E3%81%B4', '%E3%81%B6', '%E3%81%B7',
            '%E3%81%B9', '%E3%81%BA', '%E3%81%BC', '%E3%81%BD', '%E3%83%B4', '%E3%82%AC', '%E3%82%AE', '%E3%82%B0', '%E3%82%B2', '%E3%82%B4', '%E3%82%B6',
            '%E3%82%B8', '%E3%82%BA', '%E3%82%BC', '%E3%82%BE', '%E3%83%80', '%E3%83%82', '%E3%83%85', '%E3%83%87', '%E3%83%89', '%E3%83%90', '%E3%83%91',
            '%E3%83%93', '%E3%83%94', '%E3%83%96', '%E3%83%97'
        ];
        for (var i = 0; i < nfd.length; i++) {
            var pattern = new RegExp(nfd[i], 'ig');
            encodedText = encodedText.replace(pattern, nfc[i]);
        }
        return encodedText;
    }

    /* メモをテキストファイルへ保存し、再起動 / Save the memo to a text file, then restart */
    saveButton.onClick = function () {
        var file;
        if (SAVE_LOCATION_MODE === 'desktop') {
            // B: 常にデスクトップへ memo-<ドキュメント名><yyyymmdd>.txt（日本語名対応で encodeURI）
            var dateStamp = getTimeStamp().substring(0, 8); // yyyymmdd
            var fileName = 'memo-' + getDocBaseName() + '-' + dateStamp + '.txt';
            file = new File(Folder.desktop + '/' + encodeURI(fileName));
        } else {
            // A: 保存ダイアログで場所と名前を選ぶ
            var defaultFileName = getDocBaseName() + '_' + getTimeStamp() + '.txt';
            file = File.saveDialog(L(LABELS.message.savePrompt), defaultFileName);
            if (!file) return;
            if (!/\.txt$/i.test(file.name)) file = new File(file.fsName + '.txt');
        }

        var content = buildSaveContent(memoTextArea.text);
        if (!writeTextFile(file, content)) {
            alert(L(LABELS.message.writeFailed));
            return;
        }

        // 保存後もパレットは開いたまま：内容・ウィンドウ位置をそのまま保持する（再起動しない）
        // Keep the palette open after saving: preserve the content and window position as-is (no restart)
        $.global.__TextMemoContent = memoTextArea.text; // フッターは含めず入力内容のみ保持 / Keep the typed memo only (without footer)
        storeWinBounds();
    };

    /* メモをクリアして再起動 / Clear the memo, then restart */
    clearButton.onClick = function () {
        $.global.__TextMemoContent = '';
        storeWinBounds();
        restartScript();
    };

    /* すべてをコピー：テキスト欄の内容をクリップボードへ / Copy All: copy the memo to the clipboard */
    copyAllButton.onClick = function () {
        if (!memoTextArea.text) return;
        copyTextToClipboard(memoTextArea.text, function (status) {
            if (status !== 'ok') alert(L(LABELS.message.copyFailed));
        });
    };

    /* 空行削除：テキスト欄の空行を除去 / Remove Blanks: strip blank lines from the memo */
    removeBlanksButton.onClick = function () {
        memoTextArea.text = removeBlankLines(memoTextArea.text);
        $.global.__TextMemoContent = memoTextArea.text;
        updateButtonState();
    };

    /* 改行削除：テキスト欄の改行をすべて除去して1行にまとめる / Remove Breaks: strip all line breaks from the memo */
    removeBreaksButton.onClick = function () {
        memoTextArea.text = removeLineBreaks(memoTextArea.text);
        $.global.__TextMemoContent = memoTextArea.text;
        updateButtonState();
    };

    /* テキストの有無・空行の有無でボタンの活性を切り替え / Toggle buttons by whether there is text / blank lines */
    function updateButtonState() {
        var text = memoTextArea.text;
        var hasText = (text.length > 0);
        clearButton.enabled = hasText;
        copyAllButton.enabled = hasText;
        removeBlanksButton.enabled = hasText && hasBlankLines(text); // 空行が無ければディム / Dim when there are no blank lines
        removeBreaksButton.enabled = hasText && hasLineBreaks(text); // 改行が無ければディム / Dim when there are no line breaks
    }
    memoTextArea.onChanging = updateButtonState; // 入力中もリアルタイムに反映 / Update live while typing
    updateButtonState();

    // =========================================
    // キー操作 & 閉じる / Keys & close
    // =========================================
    /* 再起動（保存・クリア）由来の close では内容を消さないためのフラグ / Flag so a restart-triggered close does not clear the content */
    var isRestarting = false;

    /* Escape でパレットを閉じる / Close the palette with Escape */
    win.addEventListener('keydown', function (e) {
        if (e.keyName === 'Escape') win.close();
    });

    /* 閉じたときの内容を CLEAR_ON_CLOSE で制御（再起動時は触れない）。ウィンドウ位置は常に保持 /
       On close, control the content via CLEAR_ON_CLOSE (untouched on restart); always keep the window position */
    win.onClose = function () {
        if (!isRestarting) {
            $.global.__TextMemoContent = CLEAR_ON_CLOSE ? '' : memoTextArea.text;
        }
        storeWinBounds();
        $.global.__TextMemoWindow = null; // 参照を解放（次回起動で新規作成できるように）/ Release the reference so the next launch creates a fresh window
        return true;
    };

    /* 表示（既存位置がない場合のみ中央）/ Show the palette (center only when no saved position) */
    if (!winBounds) win.center();
    win.show();

    // =========================================
    // 保存ユーティリティ / Save utilities
    // =========================================
    /* 保存本文（フッター付き、改行を LF に正規化）を組み立て / Build the file content (with footer, newlines normalized to LF) */
    function buildSaveContent(text) {
        var bodyText = text.replace(/\r\n|\r/g, '\n');
        var logTimeStamp = getTimeStamp().replace(/(\d{4})(\d{2})(\d{2})_(\d{2})(\d{2})(\d{2})/,
            '$1-$2-$3 $4:$5:$6');
        return bodyText + '\n\n---\n[' + L(LABELS.footer.source) + ': ' + activeDocName + ']\n[' +
            L(LABELS.footer.savedAt) + ': ' + logTimeStamp + ']\n';
    }

    /* UTF-8 でテキストファイルへ書き込み（成功可否を返す）/ Write a UTF-8 text file (returns success) */
    function writeTextFile(file, content) {
        file.encoding = 'UTF-8';
        if (!file.open('w')) return false;
        var writeOk = file.write(content);
        file.close();
        return writeOk;
    }

    // =========================================
    // ドキュメント / Document helpers
    // =========================================
    /* タイムスタンプ YYYYMMDD_HHMMSS を返す / Return a timestamp formatted as YYYYMMDD_HHMMSS */
    function getTimeStamp() {
        var now = new Date();

        function padZero(num) {
            return ('0' + num).slice(-2);
        }
        return now.getFullYear() + padZero(now.getMonth() + 1) + padZero(now.getDate()) + '_' +
            padZero(now.getHours()) + padZero(now.getMinutes()) + padZero(now.getSeconds());
    }

    /* アクティブドキュメントを安全に取得（起動直後・パレット表示前向け）/ Safely get the active document (for the moment before the palette is shown) */
    /* パレット表示中はこのエンジンの app が DOM 接続を失うため、
       選択取得は fetchSelectedTextFrames（メインエンジン）を使うこと */
    function getActiveDocument() {
        if (app.documents.length === 0) return null;
        try {
            return app.activeDocument;
        } catch (docAccessError) { }
        try {
            return app.documents[0]; // パレット表示中はこれも throw しうるので保険 / may also throw while the palette is focused
        } catch (fallbackError) { }
        return null;
    }

    /* 保存ファイルの基底名（拡張子なし）を取得 / Get the base name (without extension) for the save file */
    /* パレット表示中は DOM 取得が失敗しうるため、起動時に控えた activeDocName を使う */
    function getDocBaseName() {
        if (activeDocName === L(LABELS.status.unsaved)) return 'Untitled';
        return activeDocName.replace(/\.[^\.]+$/, '');
    }

    // =========================================
    // 再起動 & 選択取得 / Restart & fetch selection
    // =========================================
    /* スクリプト自身をメインエンジンで再評価して再起動 / Restart by re-evaluating this script in the main engine */
    function restartScript() {
        var scriptFile = new File($.fileName);
        var bridge = new BridgeTalk();
        bridge.target = 'illustrator';
        bridge.body = '$.evalFile("' + scriptFile.fsName.replace(/\\/g, '\\\\') + '");';
        bridge.send(100);
        isRestarting = true; // この close では内容を消さない / This close must not clear the content
        win.close();
    }

    /* 選択テキストフレームをメインエンジン経由で取得 / Fetch selected text frames via the main engine */
    /* 常駐エンジンの app はパレット表示中に DOM 接続を失い "there is no document" を
       投げるため、生きた DOM を持つメインエンジンに BridgeTalk で問い合わせる。
       テキストフレームに加え、グループ・クリップグループ・シンボル内のテキストも対象。
       シンボルは一時レイヤーへ複製→breakLink で展開して読み取り（入れ子シンボルも多段展開）、
       複製は削除して元のシンボルには触れない。
       結果は onResult で非同期に返る。マルチバイト欠落防止に encodeURIComponent で受け渡す。
       改行はエスケープ非依存の String.fromCharCode(10)（LF）で連結する。
       コールバック: onComplete(status, text)
         status = 'ok' | 'nodoc' | 'nosel' | 'notextframe' | 'error' */
    function fetchSelectedTextFrames(onComplete) {
        // ↓ この本文はメインエンジン側で評価される / This body is evaluated in the main engine
        var collectorSource =
            '(function () {' +
            '    if (app.documents.length === 0) return "NODOC:";' +
            '    var targetDoc;' +
            '    try { targetDoc = app.activeDocument; } catch (e) { targetDoc = app.documents[0]; }' +
            '    var picked = targetDoc.selection;' +
            '    if (!picked || picked.length === 0) return "NOSEL:";' +
            '    var collected = [];' +
            // テキストと位置（geometricBounds の上端・左端）をまとめて控える。実質的に空のフレームは無視
            '    function pushFrame(frame) {' +
            '        var contents = frame.contents;' +
            '        if (!contents || contents.replace(/[\\r\\n\\x03]/g, "").replace(/\\s+/g, "") === "") return;' +
            '        var topY = 0, leftX = 0;' +
            '        try { var bounds = frame.geometricBounds; leftX = bounds[0]; topY = bounds[1]; } catch (eg) {}' +
            '        collected.push({ text: contents, top: topY, left: leftX });' +
            '    }' +
            // シンボル展開用の一時レイヤー（シンボルがある場合のみ作成）。
            // 名前と note の両方で判定し、同名のユーザーレイヤーを誤って触らないようにする
            '    var TEMP_LAYER_NAME = "__TextMemo_symbol_read__";' +
            '    var TEMP_LAYER_NOTE = "__TextMemo_symbol_read__";' +
            '    var tempLayer = null;' +
            '    var tempLayerCreated = false;' +
            '    function getTempLayer() {' +
            '        if (tempLayer) return tempLayer;' +
            '        for (var n = 0; n < targetDoc.layers.length; n++) {' +
            '            if (targetDoc.layers[n].name === TEMP_LAYER_NAME && targetDoc.layers[n].note === TEMP_LAYER_NOTE) { tempLayer = targetDoc.layers[n]; return tempLayer; }' +
            '        }' +
            '        tempLayer = targetDoc.layers.add();' +
            '        tempLayer.name = TEMP_LAYER_NAME;' +
            '        tempLayer.note = TEMP_LAYER_NOTE;' +
            '        tempLayerCreated = true;' +
            '        return tempLayer;' +
            '    }' +
            // 一時レイヤー上の使い捨て複製を、入れ子シンボルも含めて再帰展開しテキストを収集する。
            // 複製なので SymbolItem はその場で breakLink してよい。残骸は最後にまとめて削除する。
            '    function harvestCopy(item) {' +
            '        if (!item) return;' +
            '        if (item.typename === "TextFrame") { pushFrame(item); return; }' +
            '        if (item.typename === "SymbolItem") {' +
            '            try {' +
            '                targetDoc.selection = null;' +
            '                item.selected = true;' +
            '                item.breakLink();' +
            '                var subItems = [];' +
            '                for (var s = 0; s < targetDoc.selection.length; s++) subItems.push(targetDoc.selection[s]);' +
            '                for (var b = 0; b < subItems.length; b++) harvestCopy(subItems[b]);' +
            '            } catch (eh) {}' +
            '            return;' +
            '        }' +
            '        if (item.pageItems) {' +
            '            for (var p = 0; p < item.pageItems.length; p++) harvestCopy(item.pageItems[p]);' +
            '        }' +
            '    }' +
            // カンバス上のシンボルは複製してから（元には触れず）入れ子も含めて展開する
            '    function readSymbolItem(symbolItem) {' +
            '        var layer = getTempLayer();' +
            '        var duplicatedItem = symbolItem.duplicate(layer, ElementPlacement.PLACEATBEGINNING);' +
            '        harvestCopy(duplicatedItem);' +
            '    }' +
            // グループ・クリップグループ・シンボル内のテキストを収集
            '    function collectFrom(item) {' +
            '        if (item.typename === "TextFrame") {' +
            '            pushFrame(item);' +
            '        } else if (item.typename === "GroupItem") {' +
            '            for (var k = 0; k < item.pageItems.length; k++) collectFrom(item.pageItems[k]);' +
            '        } else if (item.typename === "SymbolItem") {' +
            '            readSymbolItem(item);' +
            '        }' +
            '    }' +
            // breakLink で選択が変わるため、元の選択を控えてから処理
            '    var savedSelection = [];' +
            '    for (var i = 0; i < picked.length; i++) savedSelection.push(picked[i]);' +
            '    for (var j = 0; j < savedSelection.length; j++) collectFrom(savedSelection[j]);' +
            // 一時レイヤーの後始末（元のシンボルには触れない）
            '    if (tempLayer) {' +
            '        try { while (tempLayer.pageItems.length > 0) tempLayer.pageItems[0].remove(); } catch (e4) {}' +
            '        if (tempLayerCreated) { try { tempLayer.remove(); } catch (e5) {} }' +
            '    }' +
            // 元の選択を復元
            '    try { targetDoc.selection = null; } catch (e6) {}' +
            '    for (var q = 0; q < savedSelection.length; q++) { try { savedSelection[q].selected = true; } catch (e7) {} }' +
            '    if (collected.length === 0) return "NOTF:";' +
            // カンバス上の位置で上から順（同じ高さは左から右）に並べる
            '    collected.sort(function (a, b) {' +
            '        if (Math.abs(b.top - a.top) <= 10) return a.left - b.left;' +
            '        return b.top - a.top;' +
            '    });' +
            '    var texts = [];' +
            '    for (var x = 0; x < collected.length; x++) texts.push(collected[x].text);' +
            '    return "OK:" + encodeURIComponent(texts.join(String.fromCharCode(10)));' +
            '})();';

        var bridge = new BridgeTalk();
        bridge.target = 'illustrator';
        bridge.body = collectorSource;

        bridge.onResult = function (response) {
            var payload = response.body || '';
            var colonIndex = payload.indexOf(':');
            var marker = colonIndex >= 0 ? payload.substring(0, colonIndex) : payload;
            var rest = colonIndex >= 0 ? payload.substring(colonIndex + 1) : '';

            if (marker === 'OK') {
                onComplete('ok', decodeURIComponent(rest));
            } else if (marker === 'NODOC') {
                onComplete('nodoc');
            } else if (marker === 'NOSEL') {
                onComplete('nosel');
            } else if (marker === 'NOTF') {
                onComplete('notextframe');
            } else {
                onComplete('error', payload);
            }
        };

        bridge.onError = function (response) {
            onComplete('error', response.body);
        };

        bridge.send();
    }

})();
