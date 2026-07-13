#targetengine "DocumentFontListEngine"
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ドキュメント内で使用しているテキストの組み合わせ（フォント・サイズ・行送り・
文字ツメ・トラッキング・自動カーニング・プロポーショナルメトリクス）を収集し、
重複を除いた一覧をパレットに表示する常駐スクリプトです。同じフォントでも
これらの設定が異なれば、別の候補としてリストアップします。

一覧は「フォント／サイズ／行送り／自動カーニング／文字ツメ／トラッキング／
プロポーショナル」の7列で表示します（プロポーショナルメトリクスは ON/OFF）。

例）
  フォント               サイズ   行送り  自動カーニング  文字ツメ  トラッキング  プロポーショナル
  ヒラギノ角ゴシック W3   12pt     16pt    メトリクス       0         0             ON
  ヒラギノ角ゴシック W3   10.5pt   14pt    メトリクス       10        25            OFF

- リスト上部に3つのフィルタチェックボックス（横並び・左右中央）：「非表示のテキスト
  を含める」「ロックされたテキストを含める」「現在のアートボードのみ」。前2つは初期
  ON、最後は初期 OFF。OFF/ON 変更で一覧を即再走査して対象を絞り込む
- パレットは列付きリストと、その下のボタンエリア（左：「クリックで適用」チェック
  ボックス／中央：スペーサー／右：「選択から探す」「再読み込み」「選択した条件の
  フォントを選択」ボタン）で構成。リスト左端の「使用数」列はその組み合わせを含む
  テキストフレーム数。トラッキング・プロポーショナル列は幅が狭いので列名は
  リストの tooltip で補足
- 「クリックで適用」が ON のとき、行をクリックすると選択中のテキストへその組み合わせ
  （フォント・サイズ・行送り・自動カーニング・文字ツメ・トラッキング・
  プロポーショナルメトリクス）をまとめて適用。OFF のあいだはクリックしても適用しない
- 「選択した条件のフォントを選択」ボタンで、その条件に一致する文字を含むテキスト
  フレームをドキュメント上で選択する（押すと「クリックで適用」は自動 OFF）
- 「選択から探す」ボタンで、現在ドキュメントで選択中のテキスト（先頭文字）の組み合わせ
  に一致する行を一覧でハイライトする。Illustrator に選択変更イベントもタイマーAPI
  （app.scheduleTask 等）も無いため、自動同期ではなく手動ボタンとしている
- 走査はドキュメント全体（全テキストフレーム）。document.textFrames は
  グループ内・入れ子グループ内のテキストフレームも再帰的に含む（複合パスは
  仕様上テキストを含まない）。適用先は現在の選択
- 一覧はパレット表示時に走査。ドックは開きっぱなしで onShow が再発火しないため、
  テキスト編集後は「再読み込み」ボタンで手動再走査する
- 常駐エンジン（#targetengine）でパレット表示。常駐エンジンの app は
  パレット表示中に DOM 接続を失うため、DOM 処理はメインエンジンへ
  BridgeTalk で都度委譲する（コードは encodeURIComponent で包んで送信）

### Overview

A docked palette that lists the text-composition combinations used in the
document (font, size, leading, Tsume, tracking, auto-kerning and proportional
metrics), deduped. Even for the same font, a different combination of these is
listed as a separate candidate.

The list is shown as seven columns: Font / Size / Leading / Auto Kerning / Tsume /
Tracking / Prop. Metrics (proportional metrics shown as ON/OFF).

- Above the list are three filter checkboxes (in a centered row): "Include hidden
  text", "Include locked text" and "Current artboard only". The first two default
  ON, the last OFF. Toggling any one immediately rescans the list to narrow targets
- The palette is the column list plus a button area below it (left: an
  "Apply on click" checkbox / center: a spacer / right: "Find from selection",
  "Reload" and "Select matching text" buttons). The leftmost "Count" column is the
  number of text frames containing that combination. The Tracking & Prop. Metrics
  columns are narrow, so the full column names are shown in the list's tooltip
- While "Apply on click" is ON, clicking a row applies that combination (font,
  size, leading, auto kerning, Tsume, tracking, proportional metrics) to the
  current selection; while OFF, clicking does not apply
- The "Select matching text" button selects the text frames containing
  characters that match the selected condition (this turns "Apply on click" OFF)
- The "Find from selection" button highlights the row matching the combo of the
  text currently selected in the document (its first character). Illustrator has no
  selection-changed event and no timer API (app.scheduleTask etc.), so this is a
  manual button rather than automatic sync
- The scan covers the whole document; document.textFrames recursively includes
  text frames nested in groups (compound paths can't contain text). The target
  is the current selection
- The list is scanned when the palette is shown; since onShow won't re-fire
  while docked, use the "Reload" button to rescan after editing text
- Runs as a persistent palette (#targetengine). The persistent engine's app
  loses its DOM connection while the palette is shown, so all DOM work is
  delegated to the main engine via BridgeTalk (code wrapped in encodeURIComponent)

*/

(function () {

    // =========================================
    // バージョン / Version
    // =========================================
    var SCRIPT_VERSION = "v1.1.0";

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 言語判定 / Detect UI language */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    /* ラベル定義 / Label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "ドキュメントフォントリスト", en: "Document Font List" }
        },
        column: {
            count: { ja: "使用数", en: "Count" },
            font: { ja: "フォント", en: "Font" },
            size: { ja: "サイズ", en: "Size" },
            leading: { ja: "行送り", en: "Leading" },
            kern: { ja: "自動カーニング", en: "Auto Kerning" },
            tsume: { ja: "文字ツメ", en: "Tsume" },
            tracking: { ja: "トラッキング", en: "Tracking" },
            propMetrics: { ja: "プロポーショナル", en: "Prop. Metrics" }
        },
        control: {
            applyOnClick: { ja: "クリックで適用", en: "Apply on click" },
            includeHidden: { ja: "非表示のテキストを含める", en: "Include hidden text" },
            includeLocked: { ja: "ロックされたテキストを含める", en: "Include locked text" },
            currentArtboardOnly: { ja: "現在のアートボードのみ", en: "Current artboard only" }
        },
        button: {
            fromSelection: { ja: "選択から探す", en: "Find from selection" },
            reload: { ja: "再読み込み", en: "Reload" },
            selectMatching: { ja: "一致するテキストを選択", en: "Select matching text" }
        },
        listHelpTip: {
            ja: "列: 使用数 / フォント / サイズ / 行送り / 自動カーニング / 文字ツメ / トラッキング / プロポーショナルメトリクス",
            en: "Columns: Count / Font / Size / Leading / Auto Kerning / Tsume / Tracking / Proportional Metrics"
        },
        autoLeading: { ja: "自動", en: "Auto" },
        autoKern: {
            mono: { ja: "和文等幅", en: "Metrics - Roman Only" },
            zero: { ja: "0", en: "0" },
            metrics: { ja: "メトリクス", en: "Metrics" },
            optical: { ja: "オプティカル", en: "Optical" }
        },
        status: {
            needSelection: { ja: "適用先のテキストを選択してください。", en: "Select the text to apply to." },
            needCondition: { ja: "一覧から条件を選択してください。", en: "Select a condition from the list." },
            noMatch: { ja: "一致するテキストが見つかりませんでした。", en: "No matching text was found." }
        }
    };

    /* 言語に応じたラベル文字列を取得 / Resolve a label string for the current language */
    function getLocalizedText(entry) {
        if (!entry) return "";
        return entry[currentLanguage] || entry.ja || entry.en || "";
    }

    // =========================================
    // メインエンジンで実行する DOM 処理 / DOM helpers run on the main engine
    //
    // 以下の関数群は toString() で連結し、BridgeTalk 本文に同梱して
    // メインエンジン（生きた DOM を持つ）で eval される。常駐エンジン側
    // からは直接呼ばず、本文への埋め込み用途のみ。
    // These are concatenated via toString() and embedded in the BridgeTalk
    // body, then eval'd on the main engine (which has a live DOM).
    //
    // ★重要 / IMPORTANT: この Illustrator の Function.toString() は改行を全削除する。
    // そのため WORKER_FUNCTIONS と dispatch の「本体内」コメントに `//` 行コメントを
    // 使うと、改行が消えて以降のコード（閉じ括弧含む）が丸ごとコメントアウトされ、
    // eval が「コメントが終了していません」で壊れる。本体内コメントは必ず /* */ を使うこと。
    // toString() here STRIPS all newlines, so a `//` line comment inside any shipped
    // function (or dispatch) would comment out the rest of the body. Use /* */ only.
    // =========================================

    /* 型名を安全に取得 / Safely resolve a type name */
    function getTypeName(targetObject) {
        if (targetObject === null || targetObject === undefined) return "";
        if (targetObject.typename) return targetObject.typename;
        try {
            return targetObject.constructor ? targetObject.constructor.name : "";
        } catch (e) {
            return "";
        }
    }

    /* 選択アイテムからテキスト範囲を集める（グループは再帰）/ Collect text ranges from an item (recurse into groups) */
    function collectRangesFromItem(selectedItem, rangesOut) {
        var itemTypeName = getTypeName(selectedItem);
        if (itemTypeName === "TextFrame") {
            rangesOut.push(selectedItem.textRange);
        } else if (itemTypeName === "TextRange") {
            rangesOut.push(selectedItem);
        } else if (itemTypeName === "GroupItem") {
            for (var i = 0; i < selectedItem.pageItems.length; i++) collectRangesFromItem(selectedItem.pageItems[i], rangesOut);
        }
    }

    /* 選択中のテキスト範囲を取得 / Get selected text ranges from the current document */
    function getSelectedTextRanges() {
        var activeDocument = app.activeDocument;
        var currentSelection = activeDocument.selection;
        var selectedRanges = [];
        if (!currentSelection) return selectedRanges;
        /* テキスト編集モードでは selection が TextRange になる / In text-edit mode the selection is a TextRange */
        if (getTypeName(currentSelection) === "TextRange") {
            selectedRanges.push(currentSelection);
            return selectedRanges;
        }
        if (currentSelection.length === 0) return selectedRanges;
        for (var i = 0; i < currentSelection.length; i++) {
            collectRangesFromItem(currentSelection[i], selectedRanges);
        }
        return selectedRanges;
    }

    /* AutoKernType を ID 文字列へ / Convert an AutoKernType value to an id string */
    function kernMethodToId(kernMethodValue) {
        var kernMethodText = String(kernMethodValue);
        if (kernMethodText === String(AutoKernType.AUTO)) return "metrics";
        if (kernMethodText === String(AutoKernType.OPTICAL)) return "optical";
        if (kernMethodText === String(AutoKernType.METRICSROMANONLY)) return "mono";
        if (kernMethodText === String(AutoKernType.NOAUTOKERN)) return "zero";
        return "zero";
    }

    /* 方式 ID を AutoKernType の列挙値へ変換 / Resolve a method id to an AutoKernType enum value
       和文等幅は欧文のみメトリクス（METRICSROMANONLY）/ Japanese equal width = METRICSROMANONLY */
    function resolveAutoKernType(kernMethodId) {
        if (kernMethodId === "mono") return AutoKernType.METRICSROMANONLY;
        if (kernMethodId === "metrics") return AutoKernType.AUTO;
        if (kernMethodId === "optical") return AutoKernType.OPTICAL;
        return AutoKernType.NOAUTOKERN;
    }

    /* 小数2桁までで末尾ゼロを落とす / Round to 2 decimals and trim trailing zeros */
    function formatNumber(numericValue) {
        if (numericValue === null || numericValue === undefined || isNaN(numericValue)) return "?";
        var roundedValue = Math.round(numericValue * 100) / 100;
        return String(roundedValue);
    }

    /* フォントの表示名（ファミリー＋スタイル）/ Display name of a font (family + style) */
    function getFontDisplayName(textFont) {
        if (!textFont) return "(no font)";
        var fontFamily = textFont.family || textFont.name || "";
        var fontStyle = textFont.style || "";
        return fontFamily + (fontStyle ? " " + fontStyle : "");
    }

    /* 1文字の属性から組み合わせオブジェクトを作る（読めなければ null）/ Build a combination object from one character's attributes (null if unreadable)
       属性は同じ characterAttributes 上にあり、読めるときは一括で読めるので try は1つ
       The attributes live on the same characterAttributes object, so a single try suffices */
    function readComboFromCharacter(character) {
        try {
            var charAttributes = character.characterAttributes;
            var textFont = charAttributes.textFont;
            if (!textFont) return null;
            return {
                psName: textFont.name,
                displayName: getFontDisplayName(textFont),
                size: charAttributes.size,
                leading: charAttributes.leading,
                autoLeading: charAttributes.autoLeading ? true : false,
                tsume: charAttributes.Tsume,
                tracking: charAttributes.tracking,
                kernId: kernMethodToId(charAttributes.kerningMethod),
                proportionalMetrics: charAttributes.proportionalMetrics ? true : false
            };
        } catch (e) {
            return null;
        }
    }

    /* 組み合わせの重複排除キー / Dedupe key for a combination */
    function comboKey(combo) {
        return [
            combo.psName,
            formatNumber(combo.size),
            combo.autoLeading ? "auto" : formatNumber(combo.leading),
            String(Math.round(combo.tsume)),
            String(Math.round(combo.tracking)),
            combo.kernId,
            combo.proportionalMetrics ? "p1" : "p0"
        ].join("|");
    }

    /* アイテム自身または親（グループ・レイヤー）が非表示か / Whether the item or an ancestor (group/layer) is hidden */
    function isItemHidden(item) {
        var node = item;
        while (node) {
            var nodeType = getTypeName(node);
            if (nodeType === "Document" || nodeType === "") break;
            if (nodeType === "Layer") { if (node.visible === false) return true; }
            else if (node.hidden === true) return true;
            node = node.parent;
        }
        return false;
    }

    /* アイテム自身または親（グループ・レイヤー）がロックされているか / Whether the item or an ancestor is locked */
    function isItemLocked(item) {
        var node = item;
        while (node) {
            var nodeType = getTypeName(node);
            if (nodeType === "Document" || nodeType === "") break;
            if (node.locked === true) return true;
            node = node.parent;
        }
        return false;
    }

    /* 2つの矩形 [左,上,右,下]（Yは上が大）が交差するか / Whether two rects [L,T,R,B] (Y up) intersect */
    function boundsIntersect(rectA, rectB) {
        return !(rectA[2] < rectB[0] || rectA[0] > rectB[2] || rectA[3] > rectB[1] || rectA[1] < rectB[3]);
    }

    /* ドキュメント全体を走査し、組み合わせを TAB 区切り行（LF 連結）で返す / Scan the whole document; return combos as TAB-joined rows (LF-joined)
       1行 = psName \t displayName \t size \t leading \t autoLeading(0/1) \t tsume \t tracking \t kernId \t proportionalMetrics(0/1) \t frameCount
       frameCount = その組み合わせを含むテキストフレーム数（同一フレーム内の重複は1と数える）/ number of text frames containing the combo (counted once per frame)
       scanOptions = { includeHidden, includeLocked, currentArtboardOnly } で対象を絞り込む / scanOptions filters the targets */
    function collectDocumentCombosRaw(activeDocument, scanOptions) {
        var comboByKey = {}, comboOrder = [];
        var TAB = String.fromCharCode(9), LF = String.fromCharCode(10);
        var includeHidden = !scanOptions || scanOptions.includeHidden !== false;
        var includeLocked = !scanOptions || scanOptions.includeLocked !== false;
        var currentArtboardOnly = scanOptions && scanOptions.currentArtboardOnly === true;
        /* 「現在のアートボードのみ」用に現在のアートボード矩形を取得 / Active artboard rect for "current artboard only" */
        var artboardRect = null;
        if (currentArtboardOnly) {
            try {
                var activeArtboardIndex = activeDocument.artboards.getActiveArtboardIndex();
                artboardRect = activeDocument.artboards[activeArtboardIndex].artboardRect;
            } catch (eArtboard) { artboardRect = null; }
        }
        /* document.textFrames はグループ内・入れ子グループ内のテキストフレームも再帰的に含む */
        /* document.textFrames recursively includes text frames nested inside groups */
        var textFrames = activeDocument.textFrames;
        for (var f = 0; f < textFrames.length; f++) {
            var frame = textFrames[f];
            /* フィルタ：非表示／ロック／アートボード外を除外 / Filter out hidden, locked, off-artboard frames */
            if (!includeHidden && isItemHidden(frame)) continue;
            if (!includeLocked && isItemLocked(frame)) continue;
            if (currentArtboardOnly && artboardRect) {
                var frameBounds;
                try { frameBounds = frame.geometricBounds; } catch (eBounds) { frameBounds = null; }
                if (!frameBounds || !boundsIntersect(frameBounds, artboardRect)) continue;
            }
            var characters;
            try { characters = frame.characters; } catch (eFrame) { continue; }
            var seenInThisFrame = {};
            for (var c = 0; c < characters.length; c++) {
                var combo = readComboFromCharacter(characters[c]);
                if (!combo) continue;
                var dedupeKey = comboKey(combo);
                var entry = comboByKey[dedupeKey];
                if (!entry) {
                    entry = { combo: combo, frameCount: 0 };
                    comboByKey[dedupeKey] = entry;
                    comboOrder.push(dedupeKey);
                }
                /* フレーム単位で1回だけカウント / Count once per frame */
                if (!seenInThisFrame[dedupeKey]) {
                    seenInThisFrame[dedupeKey] = true;
                    entry.frameCount++;
                }
            }
        }
        var comboRows = [];
        for (var k = 0; k < comboOrder.length; k++) {
            var rowEntry = comboByKey[comboOrder[k]];
            var rowCombo = rowEntry.combo;
            comboRows.push([
                rowCombo.psName, rowCombo.displayName, String(rowCombo.size), String(rowCombo.leading),
                rowCombo.autoLeading ? "1" : "0", String(rowCombo.tsume), String(rowCombo.tracking), rowCombo.kernId,
                rowCombo.proportionalMetrics ? "1" : "0", String(rowEntry.frameCount)
            ].join(TAB));
        }
        return comboRows.join(LF);
    }

    /* 選択範囲に組み合わせを適用 / Apply a combination to the given ranges
       戻り値: 適用できた範囲数 / Returns the number of ranges applied */
    function applyComboToRanges(selectedRanges, combo) {
        var textFont = null;
        try { textFont = app.textFonts.getByName(combo.psName); } catch (eGet) { textFont = null; }
        var kernType = resolveAutoKernType(combo.kernId);
        var appliedCount = 0;
        for (var i = 0; i < selectedRanges.length; i++) {
            /* 属性は同じ characterAttributes に連続代入するので try は1範囲につき1つ */
            /* All attributes go on the same characterAttributes, so one try per range */
            try {
                var charAttributes = selectedRanges[i].characterAttributes;
                if (textFont) charAttributes.textFont = textFont;
                if (!isNaN(combo.size)) charAttributes.size = combo.size;
                if (combo.autoLeading) {
                    charAttributes.autoLeading = true;
                } else if (!isNaN(combo.leading)) {
                    charAttributes.autoLeading = false;
                    charAttributes.leading = combo.leading;
                }
                charAttributes.Tsume = combo.tsume;
                if (!isNaN(combo.tracking)) charAttributes.tracking = combo.tracking;
                charAttributes.kerningMethod = kernType;
                charAttributes.proportionalMetrics = combo.proportionalMetrics ? true : false;
                appliedCount++;
            } catch (eApplyRange) {
                /* ロック・表示外などで代入できない範囲はスキップし、他の範囲は続行 */
                /* Skip ranges that can't be set (locked/off, etc.) and keep going with the rest */
            }
        }
        return appliedCount;
    }

    /* 組み合わせに一致する文字を含むテキストフレームを選択 / Select text frames containing characters that match the combo
       戻り値: 選択したフレーム数 / Returns the number of selected frames */
    function selectFramesMatchingCombo(activeDocument, combo) {
        var targetKey = comboKey(combo);
        var textFrames = activeDocument.textFrames;
        var matchedFrames = [];
        for (var f = 0; f < textFrames.length; f++) {
            var characters;
            try { characters = textFrames[f].characters; } catch (eFrame) { continue; }
            for (var c = 0; c < characters.length; c++) {
                var charCombo = readComboFromCharacter(characters[c]);
                if (charCombo && comboKey(charCombo) === targetKey) { matchedFrames.push(textFrames[f]); break; }
            }
        }
        /* 一致0件のときは現在の選択を保持（消してから「なし」と出さない）/ Keep the current selection when nothing matched */
        if (matchedFrames.length === 0) return 0;
        /* 配列を直接代入すれば既存選択を一括置換できる / Assigning the array replaces the current selection wholesale */
        activeDocument.selection = matchedFrames;
        return matchedFrames.length;
    }

    // =========================================
    // メインエンジン委譲 / Main-engine delegation (BridgeTalk)
    // =========================================

    /* メインエンジンへ送る処理関数 / Functions shipped to the main engine */
    var WORKER_FUNCTIONS = [
        getTypeName, collectRangesFromItem, getSelectedTextRanges,
        kernMethodToId, resolveAutoKernType, formatNumber,
        getFontDisplayName, readComboFromCharacter, comboKey,
        isItemHidden, isItemLocked, boundsIntersect,
        collectDocumentCombosRaw, applyComboToRanges, selectFramesMatchingCombo
    ];

    /* 関数配列を toString() で連結してソース文字列にする / Join the function array into a source string */
    function buildWorkerSource() {
        var workerSource = "";
        for (var i = 0; i < WORKER_FUNCTIONS.length; i++) workerSource += WORKER_FUNCTIONS[i].toString() + "\n";
        return workerSource;
    }
    var workerSourceCache = null;
    function getWorkerSource() {
        if (workerSourceCache === null) workerSourceCache = buildWorkerSource();
        return workerSourceCache;
    }

    /* メインエンジンで実行されるディスパッチャ / Dispatcher executed on the main engine
       getCombos: 組み合わせ一覧 / applyCombo: 選択へ適用 / selectMatching: 条件一致フレームを選択
       getSelectionCombo: 現在選択の先頭文字の組み合わせキー（選択と一覧の同期用）
       結果は "OK:<payload>" または "ERR:<msg>" の文字列で返す */
    function dispatch(actionId, params) {
        if (app.documents.length === 0) return "ERR:nodoc";
        try { app.activeDocument; } catch (e) { return "ERR:nodoc"; }

        if (actionId === "getCombos") {
            /* テキストの表示単位はライブな app から "text/units" を読む（rulerType ではない） */
            /* Read the text display unit from the live app via "text/units" (not rulerType) */
            var textUnitIndex = 2;
            try { textUnitIndex = app.preferences.getIntegerPreference("text/units"); } catch (eUnit) { textUnitIndex = 2; }
            /* 1行目に単位インデックス、2行目以降に組み合わせ行 / First line = unit index, rest = combo rows */
            var comboPayload = String(textUnitIndex) + String.fromCharCode(10) + collectDocumentCombosRaw(app.activeDocument, params);
            return "OK:" + encodeURIComponent(comboPayload);
        }
        if (actionId === "applyCombo") {
            var selectedRanges = getSelectedTextRanges();
            if (selectedRanges.length === 0) return "OK:nosel";
            var appliedCount = 0;
            try {
                appliedCount = applyComboToRanges(selectedRanges, params);
            } catch (errApply) {
                return "ERR:" + (errApply && errApply.message ? errApply.message : String(errApply));
            }
            app.redraw();
            return "OK:" + appliedCount;
        }
        if (actionId === "selectMatching") {
            var matchedCount = selectFramesMatchingCombo(app.activeDocument, params);
            app.redraw();
            return "OK:" + matchedCount;
        }
        if (actionId === "getSelectionCombo") {
            /* 現在選択の先頭テキスト範囲・先頭文字の組み合わせキーを返す（混在時は先頭の1つ） */
            /* Return the combo key of the first character of the first selected range (first one when mixed) */
            var syncRanges = getSelectedTextRanges();
            if (syncRanges.length === 0) return "OK:nosel";
            var syncChars;
            try { syncChars = syncRanges[0].characters; } catch (eSyncChars) { return "OK:nosel"; }
            if (!syncChars || syncChars.length === 0) return "OK:nosel";
            var syncCombo = readComboFromCharacter(syncChars[0]);
            if (!syncCombo) return "OK:nosel";
            return "OK:" + encodeURIComponent(comboKey(syncCombo));
        }
        return "OK:0";
    }
    var DISPATCH_SOURCE = "(" + dispatch.toString() + ")";

    /* パラメータを安全な JS リテラル文字列へ / Serialize params to a safe JS literal
       params は comboToParams（全フィールドあり）か null のどちらか / params is either a full combo-params object or null */
    function paramsToSource(params) {
        if (!params) return "{}";
        function numberLiteral(value) { return isNaN(value) ? "NaN" : value; }
        function stringLiteral(value) { return 'decodeURIComponent("' + encodeURIComponent(value) + '")'; }
        function boolLiteral(value) { return value ? "true" : "false"; }
        return "{" + [
            "psName:" + stringLiteral(params.psName),
            "kernId:" + stringLiteral(params.kernId),
            "size:" + numberLiteral(params.size),
            "leading:" + numberLiteral(params.leading),
            "tsume:" + (isNaN(params.tsume) ? 0 : params.tsume),
            "tracking:" + numberLiteral(params.tracking),
            "autoLeading:" + boolLiteral(params.autoLeading),
            "proportionalMetrics:" + boolLiteral(params.proportionalMetrics)
        ].join(",") + "}";
    }

    /*
     * 本文をメインエンジンへ送り、結果マーカーを解析して onDone(status, payload) を呼ぶ。
     * BridgeTalk は本文送信時にバックスラッシュをエスケープするため、コード全体を
     * encodeURIComponent で包んで送り、ターゲットで decodeURIComponent + eval して復元する。
     * Send the body to the main engine and parse the result marker.
     */
    function sendWorker(workerCode, onDone) {
        var bridge = new BridgeTalk();
        bridge.target = "illustrator";
        bridge.body = "eval(decodeURIComponent(\"" + encodeURIComponent(workerCode) + "\"));";
        bridge.onResult = function (response) {
            var responseBody = response.body || "";
            var colonIndex = responseBody.indexOf(":");
            var statusMarker = colonIndex >= 0 ? responseBody.substring(0, colonIndex) : responseBody;
            var markerPayload = colonIndex >= 0 ? responseBody.substring(colonIndex + 1) : "";
            if (statusMarker === "OK") onDone("ok", markerPayload);
            else onDone("error", markerPayload);
        };
        bridge.onError = function (response) { onDone("error", response && response.body ? response.body : "BridgeTalk error"); };
        bridge.send();
    }

    /* getCombos 用のフィルタオプションを JS リテラルへ / Serialize getCombos filter options to a JS literal */
    function scanOptionsToSource(scanOptions) {
        if (!scanOptions) return "{}";
        function boolLiteral(value) { return value ? "true" : "false"; }
        return "{includeHidden:" + boolLiteral(scanOptions.includeHidden) +
            ",includeLocked:" + boolLiteral(scanOptions.includeLocked) +
            ",currentArtboardOnly:" + boolLiteral(scanOptions.currentArtboardOnly) + "}";
    }

    /* メインエンジンへアクションを委譲（非同期）/ Delegate an action to the main engine (async)
       getCombos はフィルタオプション、それ以外は組み合わせパラメータを渡す / getCombos passes filter options; others pass combo params */
    function runWorker(actionId, params, onDone) {
        var paramSource = (actionId === "getCombos") ? scanOptionsToSource(params) : paramsToSource(params);
        var workerCode = getWorkerSource() + "\nvar __result=" + DISPATCH_SOURCE + "(\"" + actionId + "\"," + paramSource + ");__result;";
        sendWorker(workerCode, onDone);
    }

    // =========================================
    // ラベル・解析（パレット側）/ Labels & parsing (palette side)
    // =========================================

    /* カーニング method id を表示ラベルへ / Kerning method id to a display label */
    function kernLabel(kernMethodId) {
        var labelEntry = LABELS.autoKern[kernMethodId];
        return labelEntry ? getLocalizedText(labelEntry) : "—";
    }

    /* テキスト単位インデックスを {ラベル, pt換算係数} へ / Resolve a "text/units" index to {label, pointsPerUnit}
       0=inch 1=mm 2=pt 3=pica 4=cm 5=Q 6=px。pointsPerUnit は「1単位=何pt」/ pointsPerUnit = points per unit */
    function resolveTextUnitInfo(textUnitIndex) {
        switch (textUnitIndex) {
            case 0: return { label: "inch", pointsPerUnit: 72.0 };
            case 1: return { label: "mm", pointsPerUnit: 72.0 / 25.4 };
            case 2: return { label: "pt", pointsPerUnit: 1.0 };
            case 3: return { label: "pica", pointsPerUnit: 12.0 };
            case 4: return { label: "cm", pointsPerUnit: 72.0 / 2.54 };
            case 5: return { label: "Q", pointsPerUnit: 72.0 / 25.4 * 0.25 };
            case 6: return { label: "px", pointsPerUnit: 1.0 };
            default: return { label: "pt", pointsPerUnit: 1.0 };
        }
    }

    /* 各列のセル文字列を組み立て / Build the per-column cell strings
       [使用数, フォント, サイズ, 行送り, 自動カーニング, 文字ツメ, トラッキング, プロポーショナル]
       サイズ・行送りは pt 値を unitInfo で表示単位へ換算 / Size & leading converted from pt to the display unit */
    function comboCells(combo, unitInfo) {
        var countText = isNaN(combo.frameCount) ? "?" : String(combo.frameCount);
        var sizeText = formatNumber(combo.size / unitInfo.pointsPerUnit) + unitInfo.label;
        var leadingText = combo.autoLeading
            ? getLocalizedText(LABELS.autoLeading)
            : formatNumber(combo.leading / unitInfo.pointsPerUnit) + unitInfo.label;
        var tsumeText = String(Math.round(combo.tsume));
        var trackingText = isNaN(combo.tracking) ? "?" : String(Math.round(combo.tracking));
        var propMetricsText = combo.proportionalMetrics ? "ON" : "OFF";
        return [countText, combo.displayName, sizeText, leadingText, kernLabel(combo.kernId), tsumeText, trackingText, propMetricsText];
    }

    /* 委譲結果（デコード済み）を組み合わせ配列に復元し、表示順に並べる / Parse the worker payload into combos and sort them */
    function parseCombosPayload(rawPayload) {
        var combos = [];
        if (!rawPayload) return combos;
        var TAB = String.fromCharCode(9), LF = String.fromCharCode(10);
        var comboRows = rawPayload.split(LF);
        // 1行目はワーカーが付けた単位インデックス / First line is the unit index added by the worker
        var textUnitIndex = parseInt(comboRows.shift(), 10);
        var unitInfo = resolveTextUnitInfo(isNaN(textUnitIndex) ? 2 : textUnitIndex);
        for (var i = 0; i < comboRows.length; i++) {
            var fields = comboRows[i].split(TAB);
            if (fields.length < 10) continue;
            var combo = {
                psName: fields[0],
                displayName: fields[1],
                size: parseFloat(fields[2]),
                leading: parseFloat(fields[3]),
                autoLeading: fields[4] === "1",
                tsume: parseFloat(fields[5]),
                tracking: parseFloat(fields[6]),
                kernId: fields[7],
                proportionalMetrics: fields[8] === "1",
                frameCount: parseInt(fields[9], 10)
            };
            combo.cells = comboCells(combo, unitInfo);
            combos.push(combo);
        }
        // フォント名→サイズ→行送り→自動カーニング→ツメ→トラッキング→プロポーショナル の順で並べる / Sort by font, size, leading, kerning, Tsume, tracking, prop-metrics
        combos.sort(function (a, b) {
            if (a.displayName !== b.displayName) return a.displayName < b.displayName ? -1 : 1;
            if ((a.size || 0) !== (b.size || 0)) return (a.size || 0) - (b.size || 0);
            if ((a.leading || 0) !== (b.leading || 0)) return (a.leading || 0) - (b.leading || 0);
            if (a.kernId !== b.kernId) return a.kernId < b.kernId ? -1 : 1;
            if ((a.tsume || 0) !== (b.tsume || 0)) return (a.tsume || 0) - (b.tsume || 0);
            if ((a.tracking || 0) !== (b.tracking || 0)) return (a.tracking || 0) - (b.tracking || 0);
            if (a.proportionalMetrics !== b.proportionalMetrics) return a.proportionalMetrics ? 1 : -1;
            return 0;
        });
        return combos;
    }

    /* 組み合わせを適用パラメータへ / Convert a combo to apply params */
    function comboToParams(combo) {
        return {
            psName: combo.psName,
            size: combo.size,
            leading: combo.leading,
            autoLeading: combo.autoLeading,
            tsume: combo.tsume,
            tracking: combo.tracking,
            kernId: combo.kernId,
            proportionalMetrics: combo.proportionalMetrics
        };
    }

    // =========================================
    // UI構築 / Build UI
    // =========================================

    /* パレットを組み立てて参照を返す（イベント未接続）/ Build the palette and return references (events not wired yet) */
    function createPaletteUI() {
        var palette = new Window("palette", getLocalizedText(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        palette.alignChildren = ["fill", "top"];
        palette.margins = 16;
        palette.spacing = 10;

        // 上部フィルタ：3つのチェックボックスを横並び・左右中央に / Top filters: three checkboxes in a horizontally-centered row
        var filterRow = palette.add("group");
        filterRow.orientation = "row";
        filterRow.alignment = ["center", "top"];
        filterRow.alignChildren = ["left", "center"];
        filterRow.spacing = 16;

        // 非表示／ロックは初期 ON（含める）、アートボード絞り込みは初期 OFF / Hidden & locked default ON (include); artboard filter default OFF
        var includeHiddenCheckbox = filterRow.add("checkbox", undefined, getLocalizedText(LABELS.control.includeHidden));
        includeHiddenCheckbox.value = true;
        var includeLockedCheckbox = filterRow.add("checkbox", undefined, getLocalizedText(LABELS.control.includeLocked));
        includeLockedCheckbox.value = true;
        var currentArtboardOnlyCheckbox = filterRow.add("checkbox", undefined, getLocalizedText(LABELS.control.currentArtboardOnly));
        currentArtboardOnlyCheckbox.value = false;

        var columnTitles = [
            getLocalizedText(LABELS.column.count),
            getLocalizedText(LABELS.column.font),
            getLocalizedText(LABELS.column.size),
            getLocalizedText(LABELS.column.leading),
            getLocalizedText(LABELS.column.kern),
            getLocalizedText(LABELS.column.tsume),
            getLocalizedText(LABELS.column.tracking),
            getLocalizedText(LABELS.column.propMetrics)
        ];
        // トラッキング・プロポーショナルは値が短い（数値 / ON・OFF）ので幅を約4文字に / Tracking & prop-metrics carry short values, so ~4-char width
        var columnWidths = [50, 200, 60, 70, 90, 70, 48, 48];

        var comboListBox = palette.add("listbox", undefined, [], {
            multiselect: false,
            numberOfColumns: columnTitles.length,
            showHeaders: true,
            columnTitles: columnTitles,
            columnWidths: columnWidths
        });
        comboListBox.preferredSize = [650, 300];
        // 列幅が狭く見出しが切れるため、tooltip で全列名を補足 / Headers may clip in narrow columns, so tooltip supplements full column names
        comboListBox.helpTip = getLocalizedText(LABELS.listHelpTip);

        // ボタンエリア：左＝クリックで適用 / 中央＝スペーサー / 右＝再読み込み・条件一致テキストを選択
        // Button area: left = apply-on-click / center = spacer / right = reload & select matching text
        var buttonRow = palette.add("group");
        buttonRow.orientation = "row";
        buttonRow.alignment = ["fill", "center"];

        // 左：クリックで適用するか（OFF のあいだはクリックしても適用しない）/ Left: whether a click applies (no apply while OFF)
        var applyOnClickCheckbox = buttonRow.add("checkbox", undefined, getLocalizedText(LABELS.control.applyOnClick));
        applyOnClickCheckbox.alignment = ["left", "center"];
        applyOnClickCheckbox.value = true;

        // 中央：スペーサー（伸びて左右を端へ押し出す。空グループが幅0に潰れないよう最小幅を持たせる）
        // Center: spacer that stretches to push the sides apart (minimum width so an empty group won't collapse to 0)
        var spacer = buttonRow.add("group");
        spacer.alignment = ["fill", "center"];
        spacer.minimumSize = [10, 1];

        // 右：選択中テキストの条件を一覧でハイライト（このIllustratorに選択イベントが無いため手動）/ Right: highlight the current selection's combo (manual, since there is no selection event)
        var fromSelectionButton = buttonRow.add("button", undefined, getLocalizedText(LABELS.button.fromSelection));
        fromSelectionButton.alignment = ["right", "center"];

        // 右：一覧を再走査（ドックは開きっぱなしで onShow が再発火しないため手動で）/ Right: rescan the list (onShow won't re-fire while docked)
        var reloadButton = buttonRow.add("button", undefined, getLocalizedText(LABELS.button.reload));
        reloadButton.alignment = ["right", "center"];

        // 右：選択中の条件に一致するテキストを選択 / Right: select text matching the selected condition
        var selectMatchingButton = buttonRow.add("button", undefined, getLocalizedText(LABELS.button.selectMatching));
        selectMatchingButton.alignment = ["right", "center"];

        return {
            palette: palette,
            comboListBox: comboListBox,
            includeHiddenCheckbox: includeHiddenCheckbox,
            includeLockedCheckbox: includeLockedCheckbox,
            currentArtboardOnlyCheckbox: currentArtboardOnlyCheckbox,
            applyOnClickCheckbox: applyOnClickCheckbox,
            fromSelectionButton: fromSelectionButton,
            reloadButton: reloadButton,
            selectMatchingButton: selectMatchingButton
        };
    }

    /* パレットにイベントを接続 / Wire palette events */
    function bindPaletteEvents(paletteUI) {
        // BridgeTalk は非同期なので、連打による委譲の重複を抑止 / Guard against overlapping async delegations
        var applyInProgress = false;
        var currentCombos = [];
        // プログラムから listbox 選択を変えるとき onChange（クリック適用）を抑止 / Suppress onChange when setting the listbox selection programmatically
        var suppressOnChange = false;

        /* 現在のフィルタ設定を取得 / Read the current filter settings */
        function currentScanOptions() {
            return {
                includeHidden: paletteUI.includeHiddenCheckbox.value,
                includeLocked: paletteUI.includeLockedCheckbox.value,
                currentArtboardOnly: paletteUI.currentArtboardOnlyCheckbox.value
            };
        }

        /* 組み合わせ一覧を読み直して listbox に反映（フィルタ適用）/ Reload the combos into the listbox (with filters) */
        function refreshCombos() {
            runWorker("getCombos", currentScanOptions(), function (status, payload) {
                paletteUI.comboListBox.removeAll();
                currentCombos = (status === "ok") ? parseCombosPayload(decodeURIComponent(payload)) : [];
                for (var i = 0; i < currentCombos.length; i++) {
                    var cells = currentCombos[i].cells;
                    var listItem = paletteUI.comboListBox.add("item", cells[0]);
                    for (var col = 1; col < cells.length; col++) listItem.subItems[col - 1].text = cells[col];
                    listItem.comboIndex = i;
                }
            });
        }

        /* listbox の選択をプログラム的に設定（onChange を抑止）/ Set the listbox selection programmatically (onChange suppressed) */
        function setListboxSelection(itemIndex) {
            suppressOnChange = true;
            try {
                paletteUI.comboListBox.selection = (itemIndex >= 0) ? itemIndex : null;
            } catch (eSel) {
                // 一覧が空／インデックス範囲外でも無視（ハイライトしないだけ）
                // Ignore if the list is empty or the index is out of range (just no highlight)
            }
            suppressOnChange = false;
        }

        // ---- シングルクリックで選択テキストへ適用 / A single click applies to the selection ----
        // 適用先は委譲時にメインエンジンが現在の選択を読む。選択が無ければワーカーが "OK:nosel" を返すのでそのときだけ警告。
        // The main engine reads the live selection at apply time; it returns "OK:nosel" when nothing is selected.
        paletteUI.comboListBox.onChange = function () {
            // プログラム的に行を選んだときは適用しない / Skip when the selection was set programmatically
            if (suppressOnChange) return;
            if (!paletteUI.comboListBox.selection) return;
            // 「クリックで適用」が OFF なら何もしない / Do nothing while "Apply on click" is OFF
            if (!paletteUI.applyOnClickCheckbox.value) return;
            if (applyInProgress) return;
            var combo = currentCombos[paletteUI.comboListBox.selection.comboIndex];
            if (!combo) return;
            applyInProgress = true;
            runWorker("applyCombo", comboToParams(combo), function (status, payload) {
                applyInProgress = false;
                // 選択が無いときだけ警告（適用0件＝ロック範囲などとは区別）/ Warn only when nothing was selected
                if (status === "ok" && payload === "nosel") {
                    alert(getLocalizedText(LABELS.status.needSelection));
                }
            });
        };

        // ---- フィルタ変更で一覧を再走査 / Rescan when a filter changes ----
        paletteUI.includeHiddenCheckbox.onClick = function () { refreshCombos(); };
        paletteUI.includeLockedCheckbox.onClick = function () { refreshCombos(); };
        paletteUI.currentArtboardOnlyCheckbox.onClick = function () { refreshCombos(); };

        // ---- 一覧を再走査（テキスト編集後の使用数・組み合わせを反映）/ Rescan the list (reflect edits to counts/combos) ----
        paletteUI.reloadButton.onClick = function () { refreshCombos(); };

        // ---- 選択中テキストの条件を一覧でハイライト / Highlight the current selection's combo in the list ----
        // 自動同期は不可：このIllustratorに選択イベントもタイマーAPI（app.scheduleTask等）も無いため手動ボタン
        // Auto-sync isn't possible: no selection event and no timer API (app.scheduleTask etc.), so this is a manual button
        paletteUI.fromSelectionButton.onClick = function () {
            runWorker("getSelectionCombo", null, function (status, payload) {
                if (status !== "ok") return;
                if (payload === "nosel") { alert(getLocalizedText(LABELS.status.needSelection)); return; }
                var selectionKey = decodeURIComponent(payload);
                var foundIndex = -1;
                for (var i = 0; i < currentCombos.length; i++) {
                    if (comboKey(currentCombos[i]) === selectionKey) { foundIndex = i; break; }
                }
                if (foundIndex < 0) { alert(getLocalizedText(LABELS.status.noMatch)); return; }
                setListboxSelection(foundIndex);
            });
        };

        // ---- 選択中の条件に一致するテキストを選択 / Select text matching the selected condition ----
        paletteUI.selectMatchingButton.onClick = function () {
            if (!paletteUI.comboListBox.selection) {
                alert(getLocalizedText(LABELS.status.needCondition));
                return;
            }
            var combo = currentCombos[paletteUI.comboListBox.selection.comboIndex];
            if (!combo) return;
            // 一致選択後は選択が変わるので、クリック適用を自動 OFF にする / Selection changes after matching, so turn apply-on-click OFF
            paletteUI.applyOnClickCheckbox.value = false;
            runWorker("selectMatching", comboToParams(combo), function (status, payload) {
                if (status === "ok" && payload === "0") {
                    alert(getLocalizedText(LABELS.status.noMatch));
                }
            });
        };

        // 表示のたびに一覧を再走査 / Rescan the combos whenever the palette is shown
        paletteUI.palette.onShow = function () { refreshCombos(); };

        // Esc でパレットを閉じる / Close the palette on Esc
        paletteUI.palette.addEventListener("keydown", function (keyEvent) {
            if (keyEvent.keyName === "Escape") paletteUI.palette.close();
        });

        // 初回ロード：常駐ドックの再実行では onShow が再発火しないことがあるため、ここでも明示的に走査
        // Initial load: onShow may not re-fire when re-running in a persistent dock, so scan explicitly here too
        refreshCombos();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================
    function main() {
        var paletteUI = createPaletteUI();
        bindPaletteEvents(paletteUI);
        paletteUI.palette.show();
    }

    main();

})();
