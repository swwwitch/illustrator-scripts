#target illustrator
#targetengine "LockHistoryPalette"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

パレットの［ロック］でオブジェクトをロックすると、そのまとまりを履歴に1件として記録し、
あとからその単位だけ解除できます。記録は各オブジェクトのタグに書き込まれ、ドキュメントに保存されます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/LockHistoryPalette.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n577d8a654ec1

### 注意

Illustrator にはロック操作を通知する仕組みが無いため、ロックそのものをパレットで行います。
command＋2 でロックしたぶんは記録されません。

### Overview

Locking objects with the palette's Lock button records that batch as one history entry, so you can
release just that batch later. The record is written into each object's tag and stored in the document.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/LockHistoryPalette.md

### Notes

Illustrator never notifies a script that something was locked, so the locking happens in the
palette. Objects locked with command+2 are not recorded.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "LockHistoryPalette";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-23";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/LockHistoryPalette.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/LockHistoryPalette.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n577d8a654ec1"; /* 紹介記事 / article URL */

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 常駐エンジンにパレットを保持するキー（GC とパレットの二重起動を防ぐ）
       / Key that keeps the palette in the resident engine (avoids GC and duplicate palettes) */
    var PALETTE_GLOBAL_KEY = "__LockHistoryPaletteWindow";

    /* 履歴の番号を書き込むタグの名前。ドキュメントに保存されるので、ファイルを開き直しても復元できる
       値を変えると既存のドキュメントに残った記録が読めなくなるので、むやみに変えない
       / Name of the tag holding the entry number; stored in the document, so it survives reopening.
         Changing the value orphans the records already written into documents */
    var HISTORY_TAG_NAME = "LockHistoryPalette";

    /* 一覧で選んだときに描く枠（赤・塗りなし・線 10 pt・不透明度 50%）
       / Frame drawn when an entry is selected: red, no fill, 10 pt stroke, 50% opacity */
    var FRAME_RGB          = [255, 0, 0];      /* RGB ドキュメントの線の色 / Stroke color in RGB documents */
    var FRAME_CMYK         = [0, 100, 100, 0]; /* CMYK ドキュメントの線の色 / Stroke color in CMYK documents */
    var FRAME_STROKE_WIDTH = 10;               /* 線幅（pt） / Stroke width (pt) */
    var FRAME_OPACITY      = 50;               /* 不透明度（%） / Opacity (%) */
    var FRAME_LAYER_NAME   = "LockHistoryPalette Preview"; /* 枠を描く一時レイヤー名 / Temporary frame layer name */

    /* BridgeTalk の返信を区切る文字。改行やタブは転送時にリテラルの "\n" / "\t" に化けるため、
       印字できる文字を使う / Replies use printable separators: newlines and tabs arrive as literal
       two-character escapes after BridgeTalk's transport */
    var REPLY_RECORD_SEPARATOR = ";";
    var REPLY_FIELD_SEPARATOR  = "|";

    /* ［ロック］のショートカット（option＋L）。macOS では option を押すと keyName に
       合成文字が入ることがあるため、"L" と option＋L が生む "¬" の両方を受ける
       / Shortcut for Lock (option+L); macOS may report the composed character, so accept both */
    var LOCK_SHORTCUT_KEYS = { "L": true, "¬": true };

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* パレット外周の余白と行間 / Palette margins and spacing */
    var PALETTE_MARGINS = 12;
    var PALETTE_SPACING = 6;

    /* 履歴一覧の表示サイズ。パレットの幅はここで決まる
       / Preferred size of the entry list; it sets the palette width */
    var ENTRY_LIST_SIZE = [280, 200];

    /* 一覧の上下に足す余白 / Extra space above and below the list */
    var ENTRY_LIST_VERTICAL_MARGIN = 5;

    /* 最下部のステータス行の上に足す余白 / Extra space above the status line at the bottom */
    var STATUS_TOP_MARGIN = 5;

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

    var uiLanguage = detectUILanguage();

    var LABELS = {
        dialog: {
            title: { ja: "ロックの履歴と管理", en: "Lock History and Management" }
        },
        button: {
            lockSelection: { ja: "ロック", en: "Lock" },
            unlockEntry: { ja: "ロック解除", en: "Unlock" },
            unlockAll: { ja: "すべてロック解除", en: "Unlock All" },
            forgetEntry: { ja: "リストから外す", en: "Remove from List" },
            forgetAll: { ja: "すべてリストから外す", en: "Remove All from List" }
        },
        tooltip: {
            lockSelection: {
                ja: "選択中のオブジェクトを履歴に1件として記録してからロックします（option＋L）。記録はドキュメントに保存されます",
                en: "Records the selected objects as one history entry, then locks them (option+L); the record is stored in the document"
            },
            entryList: {
                ja: "選ぶと、その履歴のオブジェクト全体を囲む赤い枠がドキュメント上に表示されます",
                en: "Selecting an entry draws a red frame around the whole extent of its objects"
            },
            unlockEntry: {
                ja: "選んだ履歴のオブジェクトのロックを解除し、記録を削除します",
                en: "Unlocks the objects of the selected entry and deletes its record"
            },
            unlockAll: {
                ja: "履歴にあるすべてのオブジェクトのロックを解除し、記録を削除します",
                en: "Unlocks the objects of every entry and deletes all records"
            },
            forgetEntry: {
                ja: "ロックはそのままに、選んだ履歴の記録だけを削除します",
                en: "Deletes the selected entry's record but leaves the objects locked"
            },
            forgetAll: {
                ja: "ロックはそのままに、履歴にあるすべての記録を削除します",
                en: "Deletes every record in the history but leaves the objects locked"
            }
        },
        listItem: {
            entry: { ja: "{id}回目（{count}個 / {type}）", en: "#{id} ({count} items / {type})" }
        },
        status: {
            summary: { ja: "{units}件 ／ {items}個", en: "{units} entries / {items} items" },
            working: { ja: "処理中…", en: "Working..." }
        },
        /* PageItem.typename の表示名。ここに無い型名はそのまま出す
           / Display names for PageItem.typename; anything missing is shown as-is */
        itemType: {
            PathItem: { ja: "パス", en: "Path" },
            CompoundPathItem: { ja: "複合パス", en: "Compound Path" },
            TextFrame: { ja: "テキスト", en: "Text" },
            GroupItem: { ja: "グループ", en: "Group" },
            PlacedItem: { ja: "配置画像", en: "Placed Image" },
            RasterItem: { ja: "画像", en: "Image" },
            SymbolItem: { ja: "シンボル", en: "Symbol" },
            MeshItem: { ja: "メッシュ", en: "Mesh" },
            PluginItem: { ja: "プラグインオブジェクト", en: "Plug-in Object" },
            GraphItem: { ja: "グラフ", en: "Graph" },
            LegacyTextItem: { ja: "旧バージョンのテキスト", en: "Legacy Text" },
            NonNativeItem: { ja: "非ネイティブアート", en: "Non-native Art" }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてから実行してください。", en: "Open a document first." },
            noSelection: { ja: "ロックするオブジェクトを選択してください。", en: "Select the objects you want to lock." },
            entryGone: { ja: "この記録のオブジェクトが見つかりません。", en: "The objects of this record could not be found." },
            remoteFailed: {
                ja: "Illustrator 側の処理に失敗しました：{message}",
                en: "The operation failed inside Illustrator: {message}"
            }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "button.unlockEntry" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
            if (labelNode === undefined || labelNode === null) return labelPath;
        }
        var localizedText = labelNode[uiLanguage];
        if (typeof localizedText !== "string") localizedText = labelNode.en;
        return (typeof localizedText === "string") ? localizedText : labelPath;
    }

    /**
     * ラベル内の {name} を実際の値に差し替える
     * @param {string} labelPath - LABELS のドット区切りキー
     * @param {Object} replacements - {name: 値} の組
     * @returns {string} 差し替え後のテキスト
     */
    function formatLabel(labelPath, replacements) {
        var formattedText = getLabel(labelPath);
        for (var replacementName in replacements) {
            if (!replacements.hasOwnProperty(replacementName)) continue;
            formattedText = formattedText.replace("{" + replacementName + "}", String(replacements[replacementName]));
        }
        return formattedText;
    }

    /**
     * DOM の型名を表示用の名前にする
     * @param {string} typeName - PageItem.typename の値
     * @returns {string} 表示用の名前（対応が無ければ型名をそのまま返す）
     */
    function localizedTypeName(typeName) {
        return LABELS.itemType.hasOwnProperty(typeName) ? getLabel("itemType." + typeName) : typeName;
    }

    // =========================================
    // メインエンジンへ渡す処理 / Code delegated to the main engine
    // =========================================

    /*
       常駐パレットのエンジンからは app.activeDocument を取得できないため、DOM を触る処理は
       すべて文字列として組み立て、BridgeTalk でメインエンジンへ渡して実行する。
       ここは向こう側で eval されるコードなので JSDoc は付けない（toString の出力が壊れるのと同じ理由）。

       The palette's engine cannot reach app.activeDocument, so every DOM operation is built as
       source text and executed in the main engine through BridgeTalk. No JSDoc in here: this is
       code that gets evaluated on the other side.

       戻り値の書式 / Reply format:
         "OK;番号|個数|型名;番号|個数|型名..." または "ERROR|コード"
         改行とタブは BridgeTalk の転送で化けるので使わない
         / Newlines and tabs do not survive BridgeTalk's transport, so they are never used
    */

    /* パレット側の設定を向こう側へ渡す / Hand the palette's settings to the other side */
    var REMOTE_CONSTANTS = [
        'var LUP_TAG_NAME = "' + HISTORY_TAG_NAME + '";',
        'var LUP_FRAME_LAYER = "' + FRAME_LAYER_NAME + '";',
        'var LUP_FRAME_RGB = [' + FRAME_RGB.join(", ") + '];',
        'var LUP_FRAME_CMYK = [' + FRAME_CMYK.join(", ") + '];',
        'var LUP_FRAME_STROKE = ' + FRAME_STROKE_WIDTH + ';',
        'var LUP_FRAME_OPACITY = ' + FRAME_OPACITY + ';',
        'var LUP_RECORD_SEP = "' + REPLY_RECORD_SEPARATOR + '";',
        'var LUP_FIELD_SEP = "' + REPLY_FIELD_SEPARATOR + '";'
    ];

    /* タグの読み書きと、履歴のまとまりを1回の走査で集める処理
       / Reading and writing tags, and collecting the whole history in a single pass */
    var REMOTE_TAG_HELPERS = [
        '// アイテムに付いた履歴のタグを名前で探す（tags["名前"] は無いと例外になる）',
        'function lupFindTag(item) {',
        '    var itemTags = item.tags;',
        '    for (var i = 0; i < itemTags.length; i++) {',
        '        if (itemTags[i].name === LUP_TAG_NAME) return itemTags[i];',
        '    }',
        '    return null;',
        '}',
        '',
        '// タグはロック中のアイテムには付けられないので、呼ぶ側で必ず解除しておく',
        'function lupRemoveTags(items) {',
        '    for (var i = 0; i < items.length; i++) {',
        '        var historyTag = lupFindTag(items[i]);',
        '        if (historyTag) historyTag.remove();',
        '    }',
        '}',
        '',
        '// タグの付いたアイテムを1回の走査で集める。番号順の一覧・番号ごとのアイテム・先頭の型名を返す',
        '// pageItems はグループの中身も平坦に返るので、入れ子をたどる必要はない',
        'function lupCollectEntries(doc) {',
        '    var itemsById = {}, typeById = {}, entryIds = [];',
        '    var allItems = doc.pageItems;',
        '    for (var i = 0; i < allItems.length; i++) {',
        '        var historyTag = lupFindTag(allItems[i]);',
        '        if (!historyTag) continue;',
        '        var entryKey = "entry" + historyTag.value;',
        '        if (!itemsById[entryKey]) {',
        '            itemsById[entryKey] = [];',
        '            typeById[entryKey] = allItems[i].typename;',
        '            entryIds.push(Number(historyTag.value));',
        '        }',
        '        itemsById[entryKey].push(allItems[i]);',
        '    }',
        '    entryIds.sort(function (a, b) { return a - b; });',
        '    return { ids: entryIds, itemsById: itemsById, typeById: typeById };',
        '}'
    ];

    /* ロック状態の読み書き / Reading and writing the lock state */
    var REMOTE_LOCK_HELPERS = [
        '// 戻すために、いまのロック状態を控える',
        'function lupLockStates(items) {',
        '    var lockStates = [];',
        '    for (var i = 0; i < items.length; i++) lockStates.push(items[i].locked);',
        '    return lockStates;',
        '}',
        '',
        '// lockState は真偽値か、lupLockStates() が返した配列',
        'function lupSetLocked(items, lockState) {',
        '    var perItem = (lockState instanceof Array);',
        '    for (var i = 0; i < items.length; i++) {',
        '        // レイヤーごとロックされていると代入が通らない',
        '        try { items[i].locked = perItem ? lockState[i] : lockState; } catch (e) {}',
        '    }',
        '}'
    ];

    /* 位置を示す赤い枠 / The red frame that marks a location */
    var REMOTE_FRAME_HELPERS = [
        '// 枠のレイヤーは名前で掃く。前回のパレットが残したものも拾える',
        'function lupClearFrame(doc) {',
        '    var documentLayers = doc.layers;',
        '    for (var i = documentLayers.length - 1; i >= 0; i--) {',
        '        if (documentLayers[i].name === LUP_FRAME_LAYER) documentLayers[i].remove();',
        '    }',
        '}',
        '',
        'function lupFrameColor(doc) {',
        '    if (doc.documentColorSpace === DocumentColorSpace.CMYK) {',
        '        var cmykColor = new CMYKColor();',
        '        cmykColor.cyan = LUP_FRAME_CMYK[0];',
        '        cmykColor.magenta = LUP_FRAME_CMYK[1];',
        '        cmykColor.yellow = LUP_FRAME_CMYK[2];',
        '        cmykColor.black = LUP_FRAME_CMYK[3];',
        '        return cmykColor;',
        '    }',
        '    var rgbColor = new RGBColor();',
        '    rgbColor.red = LUP_FRAME_RGB[0];',
        '    rgbColor.green = LUP_FRAME_RGB[1];',
        '    rgbColor.blue = LUP_FRAME_RGB[2];',
        '    return rgbColor;',
        '}',
        '',
        '// 線や効果を含めた外接範囲 [左, 上, 右, 下]',
        'function lupUnionBounds(items) {',
        '    var unionRect = null;',
        '    for (var i = 0; i < items.length; i++) {',
        '        var bounds = items[i].visibleBounds;',
        '        if (!unionRect) { unionRect = [bounds[0], bounds[1], bounds[2], bounds[3]]; continue; }',
        '        if (bounds[0] < unionRect[0]) unionRect[0] = bounds[0];',
        '        if (bounds[1] > unionRect[1]) unionRect[1] = bounds[1];',
        '        if (bounds[2] > unionRect[2]) unionRect[2] = bounds[2];',
        '        if (bounds[3] < unionRect[3]) unionRect[3] = bounds[3];',
        '    }',
        '    return unionRect;',
        '}',
        '',
        '// 最上位の一時レイヤーに枠を描き、アクティブレイヤーはその場で戻す',
        'function lupDrawFrame(doc, frameBounds) {',
        '    var layerBefore = doc.activeLayer;',
        '    var frameLayer = doc.layers.add();',
        '    frameLayer.name = LUP_FRAME_LAYER;',
        '    frameLayer.zOrder(ZOrderMethod.BRINGTOFRONT);',
        '    var entryFrame = frameLayer.pathItems.rectangle(frameBounds[1], frameBounds[0], frameBounds[2] - frameBounds[0], frameBounds[1] - frameBounds[3]);',
        '    entryFrame.filled = false;',
        '    entryFrame.stroked = true;',
        '    entryFrame.strokeColor = lupFrameColor(doc);',
        '    entryFrame.strokeWidth = LUP_FRAME_STROKE;',
        '    entryFrame.opacity = LUP_FRAME_OPACITY;',
        '    doc.activeLayer = layerBefore;',
        '}'
    ];

    /* パレットから呼ばれる命令。どれも最後に最新の一覧を返す
       / The commands the palette calls; each one ends by returning the current list */
    var REMOTE_COMMANDS = [
        'function lupReadEntries(doc) {',
        '    var collected = lupCollectEntries(doc);',
        '    var replyRecords = ["OK"];',
        '    for (var i = 0; i < collected.ids.length; i++) {',
        '        var entryKey = "entry" + collected.ids[i];',
        '        replyRecords.push(collected.ids[i] + LUP_FIELD_SEP + collected.itemsById[entryKey].length + LUP_FIELD_SEP + collected.typeById[entryKey]);',
        '    }',
        '    return replyRecords.join(LUP_RECORD_SEP);',
        '}',
        '',
        'function lupLockSelection(doc) {',
        '    var selected = doc.selection || [];',
        '    if (!selected.length) return "ERROR" + LUP_FIELD_SEP + "NOSEL";',
        '    // ロックすると選択から外れるので、先に配列へ写す',
        '    var items = [];',
        '    for (var i = 0; i < selected.length; i++) items.push(selected[i]);',
        '    // 番号は昇順に並んでいるので、末尾が最大値',
        '    var entryIds = lupCollectEntries(doc).ids;',
        '    var newEntryId = String(entryIds.length ? entryIds[entryIds.length - 1] + 1 : 1);',
        '    for (var j = 0; j < items.length; j++) {',
        '        var historyTag = lupFindTag(items[j]);',
        '        if (!historyTag) { historyTag = items[j].tags.add(); historyTag.name = LUP_TAG_NAME; }',
        '        historyTag.value = newEntryId;',
        '    }',
        '    lupSetLocked(items, true);',
        '    doc.selection = null;',
        '    return lupReadEntries(doc);',
        '}',
        '',
        '// entryIdList が null なら履歴のすべてが対象',
        'function lupDiscard(doc, entryIdList, keepLocked) {',
        '    var collected = lupCollectEntries(doc);',
        '    var targetIds = entryIdList || collected.ids;',
        '    for (var i = 0; i < targetIds.length; i++) {',
        '        var items = collected.itemsById["entry" + targetIds[i]];',
        '        if (!items) continue;',
        '        var lockStates = keepLocked ? lupLockStates(items) : null;',
        '        lupSetLocked(items, false);',
        '        lupRemoveTags(items);',
        '        if (keepLocked) lupSetLocked(items, lockStates);',
        '    }',
        '    return lupReadEntries(doc);',
        '}',
        '',
        'function lupShowLocation(doc, entryId) {',
        '    var items = lupCollectEntries(doc).itemsById["entry" + entryId];',
        '    if (!items) return "ERROR" + LUP_FIELD_SEP + "NOITEMS";',
        '    // ロックされたままでは範囲を測れないので、いったん外す',
        '    var lockStates = lupLockStates(items);',
        '    lupSetLocked(items, false);',
        '    try {',
        '        var frameBounds = lupUnionBounds(items);',
        '        if (frameBounds) lupDrawFrame(doc, frameBounds);',
        '    } finally {',
        '        // 途中で失敗してもロック状態は必ず戻す',
        '        lupSetLocked(items, lockStates);',
        '    }',
        '    return lupReadEntries(doc);',
        '}',
        '',
        'function lupRun(command, entryId, keepLocked) {',
        '    if (app.documents.length === 0) return "ERROR" + LUP_FIELD_SEP + "NODOC";',
        '    var doc = app.activeDocument;',
        '    // どの操作でも、まず前回の枠を片づける',
        '    lupClearFrame(doc);',
        '    var reply;',
        '    if (command === "lock") reply = lupLockSelection(doc);',
        '    else if (command === "show") reply = lupShowLocation(doc, entryId);',
        '    else if (command === "discard") reply = lupDiscard(doc, [entryId], keepLocked);',
        '    else if (command === "discardAll") reply = lupDiscard(doc, null, keepLocked);',
        '    else reply = lupReadEntries(doc);',
        '    app.redraw();',
        '    return reply;',
        '}'
    ];

    var REMOTE_PRELUDE = [].concat(
        REMOTE_CONSTANTS,
        REMOTE_TAG_HELPERS,
        REMOTE_LOCK_HELPERS,
        REMOTE_FRAME_HELPERS,
        REMOTE_COMMANDS
    ).join("\n");

    // =========================================
    // パレットの状態 / Palette state
    // =========================================

    /* メインエンジンから受け取った表示用の一覧 / Display rows received from the main engine */
    var historyEntries = [];

    var historyPalette = null;
    var entryListBox = null;
    var statusLabel = null;

    /* パレットが開いているか。閉じたあとに返信が届いても UI に触らないための目印。
       これが true の間は UI の参照がそろっているので、個別の null チェックは要らない
       / Whether the palette is open; replies arriving after it closed must not touch the UI.
         While it is true the UI references are all in place, so no per-control null checks */
    var paletteIsOpen = false;

    /* 一覧を作り直している間は true。プログラムから選択を戻すと onChange が発火するため、
       そのまま位置の表示に流すと送信と再描画が無限に往復する
       / True while the list is being rebuilt; restoring the selection fires onChange, and passing
         that straight to the location request would bounce between request and redraw forever */
    var isRebuildingList = false;

    /* 応答待ちの BridgeTalk メッセージ。ローカル変数のままだと、返信が届く前に回収されて落ちる
       / Pending BridgeTalk messages; held as locals they can be collected before the reply arrives */
    var pendingMessages = [];

    // =========================================
    // メインエンジンとのやりとり / Talking to the main engine
    // =========================================

    /**
     * 返信が届いたメッセージを応答待ちから外す
     * @param {BridgeTalk} bridgeMessage - 外すメッセージ
     * @returns {void}
     */
    function forgetPendingMessage(bridgeMessage) {
        for (var i = 0; i < pendingMessages.length; i++) {
            if (pendingMessages[i] !== bridgeMessage) continue;
            pendingMessages.splice(i, 1);
            return;
        }
    }

    /**
     * メインエンジンへ命令を送る（結果は非同期で戻る）
     * @param {string} command - "read" / "lock" / "show" / "discard" / "discardAll"
     * @param {string} entryId - 対象の履歴番号（不要なときは空文字）
     * @param {boolean} keepLocked - 記録を消すときにロックを残すか
     * @returns {void}
     */
    function sendToMainEngine(command, entryId, keepLocked) {
        if (!paletteIsOpen) return;

        var bridgeMessage = new BridgeTalk();
        bridgeMessage.target = "illustrator";
        bridgeMessage.body = REMOTE_PRELUDE + "\n" +
            'lupRun("' + command + '", "' + (entryId || "") + '", ' + (keepLocked ? "true" : "false") + ');';
        bridgeMessage.onResult = function (replyMessage) {
            forgetPendingMessage(bridgeMessage);
            applyRemoteReply(replyMessage.body);
        };
        bridgeMessage.onError = function (replyMessage) {
            forgetPendingMessage(bridgeMessage);
            applyRemoteReply("ERROR" + REPLY_FIELD_SEPARATOR + replyMessage.body);
        };

        pendingMessages.push(bridgeMessage);
        statusLabel.text = getLabel("status.working");
        bridgeMessage.send();
    }

    /**
     * エラーコードに対応するメッセージを出す
     * @param {string} errorCode - "NODOC" / "NOSEL" / "NOITEMS" など
     * @returns {void}
     */
    function reportRemoteError(errorCode) {
        if (errorCode === "NODOC") { alert(getLabel("alert.noDocument")); return; }
        if (errorCode === "NOSEL") { alert(getLabel("alert.noSelection")); return; }
        if (errorCode === "NOITEMS") { alert(getLabel("alert.entryGone")); return; }
        alert(formatLabel("alert.remoteFailed", { message: errorCode }));
    }

    /**
     * メインエンジンからの返信を一覧に反映する
     * @param {string} replyText - lupRun() が返した文字列
     * @returns {void}
     */
    function applyRemoteReply(replyText) {
        if (!paletteIsOpen) return;

        var replyRecords = String(replyText).split(REPLY_RECORD_SEPARATOR);
        if (replyRecords[0] !== "OK") {
            /* 一覧はそのまま。作り直すのは「処理中…」を消すため
               / The list is unchanged; rebuilding only clears the "Working..." line */
            rebuildEntryList();
            var errorFields = replyRecords[0].split(REPLY_FIELD_SEPARATOR);
            reportRemoteError(errorFields[1] || errorFields[0]);
            return;
        }

        historyEntries = [];
        for (var i = 1; i < replyRecords.length; i++) {
            if (!replyRecords[i]) continue;
            var fields = replyRecords[i].split(REPLY_FIELD_SEPARATOR);
            historyEntries.push({ id: Number(fields[0]), count: Number(fields[1]), typeName: fields[2] });
        }
        rebuildEntryList();
    }

    // =========================================
    // 表示更新 / Display refresh
    // =========================================

    /**
     * 一覧で選ばれている履歴を返す
     * @returns {Object} 選択中の履歴（無ければ null）
     */
    function getSelectedEntry() {
        var listSelection = entryListBox.selection;
        if (!listSelection) return null;
        var selectedIndex = listSelection.index;
        return (selectedIndex >= 0 && selectedIndex < historyEntries.length) ? historyEntries[selectedIndex] : null;
    }

    /**
     * 記録中のアイテムの総数を数える
     * @returns {number} アイテム数
     */
    function countRecordedItems() {
        var totalCount = 0;
        for (var i = 0; i < historyEntries.length; i++) totalCount += historyEntries[i].count;
        return totalCount;
    }

    /**
     * 履歴一覧とステータス行を作り直す（選択位置は保つ）
     * @returns {void}
     */
    function rebuildEntryList() {
        var selectedIndex = entryListBox.selection ? entryListBox.selection.index : -1;

        isRebuildingList = true;
        entryListBox.removeAll();
        for (var i = 0; i < historyEntries.length; i++) {
            entryListBox.add("item", formatLabel("listItem.entry", {
                id: historyEntries[i].id,
                count: historyEntries[i].count,
                type: localizedTypeName(historyEntries[i].typeName)
            }));
        }
        if (selectedIndex >= 0 && selectedIndex < entryListBox.items.length) entryListBox.selection = selectedIndex;
        isRebuildingList = false;

        statusLabel.text = formatLabel("status.summary", {
            units: historyEntries.length,
            items: countRecordedItems()
        });
    }

    // =========================================
    // 操作 / Actions
    // =========================================

    /**
     * アクティブなドキュメントから記録を読み直す
     * @returns {void}
     */
    function refreshHistory() {
        sendToMainEngine("read", "", false);
    }

    /**
     * 選択中のオブジェクトを履歴に1件として記録してからロックする
     * @returns {void}
     */
    function lockSelectionAsEntry() {
        sendToMainEngine("lock", "", false);
    }

    /**
     * 選んだ履歴の範囲を赤い枠で示す（一覧で選んだときに呼ばれる）
     * @returns {void}
     */
    function showEntryLocation() {
        var targetEntry = getSelectedEntry();
        if (targetEntry) sendToMainEngine("show", targetEntry.id, false);
    }

    /**
     * 選んだ履歴の記録を削除する
     * @param {boolean} keepLocked - true ならロックを残し、false ならロックも解除する
     * @returns {void}
     */
    function discardSelectedEntry(keepLocked) {
        var targetEntry = getSelectedEntry();
        if (targetEntry) sendToMainEngine("discard", targetEntry.id, keepLocked);
    }

    /**
     * 選んだ履歴のロックを解除し、記録を削除する
     * @returns {void}
     */
    function unlockSelectedEntry() {
        discardSelectedEntry(false);
    }

    /**
     * ロックはそのままに、選んだ履歴の記録だけを削除する
     * @returns {void}
     */
    function forgetSelectedEntry() {
        discardSelectedEntry(true);
    }

    /**
     * 履歴にあるすべてのロックを解除し、記録を削除する
     * @returns {void}
     */
    function unlockAllEntries() {
        sendToMainEngine("discardAll", "", false);
    }

    /**
     * ロックはそのままに、履歴にあるすべての記録を削除する
     * @returns {void}
     */
    function forgetAllEntries() {
        sendToMainEngine("discardAll", "", true);
    }

    // =========================================
    // パレットの組み立て / Palette construction
    // =========================================

    /**
     * 上下に余白を持つ囲みグループを足す（ScriptUI は子ごとの余白を持てない）
     * @param {Window} parentPalette - 追加先のパレット
     * @param {number} topMargin - 上の余白
     * @param {number} bottomMargin - 下の余白
     * @returns {Group} 作成したグループ
     */
    function addMarginBox(parentPalette, topMargin, bottomMargin) {
        var marginBox = parentPalette.add("group");
        marginBox.orientation = "column";
        marginBox.alignment = ["fill", "top"];
        marginBox.alignChildren = ["fill", "top"];
        marginBox.margins = [0, topMargin, 0, bottomMargin];
        return marginBox;
    }

    /**
     * ボタンを横一列に並べる
     * @param {Window} parentPalette - 追加先のパレット
     * @param {Object[]} buttonSpecs - {label, tooltip, onClick} の配列
     * @returns {Group} 作成したボタン行
     */
    function addButtonRow(parentPalette, buttonSpecs) {
        var buttonRow = parentPalette.add("group");
        buttonRow.orientation = "row";
        /* alignChildren を "left" にしないと親の fill を継承してボタンが横いっぱいに伸びる
           / Without "left" the row inherits the parent's fill and stretches the buttons */
        buttonRow.alignment = ["fill", "top"];
        buttonRow.alignChildren = ["left", "center"];
        buttonRow.spacing = PALETTE_SPACING;
        for (var i = 0; i < buttonSpecs.length; i++) {
            var rowButton = buttonRow.add("button", undefined, getLabel(buttonSpecs[i].label));
            rowButton.helpTip = getLabel(buttonSpecs[i].tooltip);
            rowButton.onClick = buttonSpecs[i].onClick;
        }
        return buttonRow;
    }

    /**
     * 履歴一覧を足す（選ぶとその位置を枠で示す）
     * @param {Window} parentPalette - 追加先のパレット
     * @returns {void}
     */
    function addEntryList(parentPalette) {
        var listArea = addMarginBox(parentPalette, ENTRY_LIST_VERTICAL_MARGIN, ENTRY_LIST_VERTICAL_MARGIN);

        entryListBox = listArea.add("listbox", undefined, undefined, { numberOfColumns: 1, showHeaders: false });
        entryListBox.preferredSize = ENTRY_LIST_SIZE;
        entryListBox.helpTip = getLabel("tooltip.entryList");

        /* 閉じる最中にも届きうるので、開いているかどうかも見る
           / This can still fire while the palette is closing, so check that it is open */
        entryListBox.onChange = function () {
            if (!paletteIsOpen || isRebuildingList) return;
            showEntryLocation();
        };
    }

    /**
     * option＋L で［ロック］を実行できるようにする
     * @param {Window} targetPalette - 対象のパレット
     * @returns {void}
     */
    function bindLockShortcut(targetPalette) {
        /* キーはパレットにフォーカスがあるときだけ届く / Keys only reach the palette while it has focus */
        targetPalette.addEventListener("keydown", function (keyEvent) {
            if (!keyEvent.altKey) return;
            var pressedKey = String(keyEvent.keyName || "").toUpperCase();
            if (!LOCK_SHORTCUT_KEYS[pressedKey]) return;
            keyEvent.preventDefault();
            lockSelectionAsEntry();
        });
    }

    /**
     * パレットの中身を組み立てる
     * @returns {Window} 組み立てたパレット
     */
    function buildPalette() {
        var builtPalette = new Window("palette", getLabel("dialog.title"), undefined, { resizeable: true });
        builtPalette.orientation = "column";
        builtPalette.alignChildren = ["fill", "top"];
        builtPalette.spacing = PALETTE_SPACING;
        builtPalette.margins = PALETTE_MARGINS;

        /* 記録する / Record a new entry */
        addButtonRow(builtPalette, [
            { label: "button.lockSelection", tooltip: "tooltip.lockSelection", onClick: lockSelectionAsEntry }
        ]);

        addEntryList(builtPalette);

        /* ロックも解除する（左が選んだ1件、右が履歴のすべて）
           / Unlocking: the selected entry on the left, the whole history on the right */
        addButtonRow(builtPalette, [
            { label: "button.unlockEntry", tooltip: "tooltip.unlockEntry", onClick: unlockSelectedEntry },
            { label: "button.unlockAll", tooltip: "tooltip.unlockAll", onClick: unlockAllEntries }
        ]);

        /* 記録だけ消す（ロックは残す）/ Discarding records only, leaving the locks in place */
        addButtonRow(builtPalette, [
            { label: "button.forgetEntry", tooltip: "tooltip.forgetEntry", onClick: forgetSelectedEntry },
            { label: "button.forgetAll", tooltip: "tooltip.forgetAll", onClick: forgetAllEntries }
        ]);

        statusLabel = addMarginBox(builtPalette, STATUS_TOP_MARGIN, 0).add("statictext", undefined, "");

        bindLockShortcut(builtPalette);

        return builtPalette;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 前回のパレットが残っていれば閉じる（二重起動の防止）
     * @returns {void}
     */
    function closeKeptPalette() {
        var keptPalette = $.global[PALETTE_GLOBAL_KEY];
        if (!keptPalette) return;
        /* すでに閉じられたウィンドウは close() が例外になる
           / close() throws when the window is already gone */
        try { keptPalette.close(); } catch (e) {}
        $.global[PALETTE_GLOBAL_KEY] = null;
    }

    /**
     * パレットを表示して、ドキュメントの記録を読み込む
     * @returns {void}
     */
    function showLockHistoryPalette() {
        closeKeptPalette();

        historyPalette = buildPalette();

        /* パレットは選択やロックの変更を通知されないため、アクティブになったタイミングで読み直す
           / The palette is never notified of selection or lock changes, so re-read on activation */
        historyPalette.onActivate = refreshHistory;

        historyPalette.onClose = function () {
            /* 閉じる最中に DOM を触る・$.global の参照を外すと Illustrator が落ちる。
               ここでは「閉じた」印を立てて UI の参照を捨てるだけにし、
               $.global の参照と描いたままの枠は、次回起動時の closeKeptPalette() と読み直しで片づける
               / Touching the DOM or releasing the $.global reference while the window is being
                 destroyed crashes Illustrator. Only mark it closed and drop the UI references;
                 the global reference and any leftover frame are cleaned up on the next launch */
            paletteIsOpen = false;
            historyPalette = null;
            entryListBox = null;
            statusLabel = null;
            return true;
        };

        /* 常駐エンジンに保持して GC を避ける / Keep it in the resident engine so it is not garbage-collected */
        $.global[PALETTE_GLOBAL_KEY] = historyPalette;

        paletteIsOpen = true;
        historyPalette.show();
        refreshHistory();
    }

    showLockHistoryPalette();

})();
