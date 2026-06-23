#targetengine "DocumentFontListEngine"
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ドキュメント内で使用しているテキストの組み合わせ（フォント・サイズ・行送り・
文字ツメ・トラッキング・自動カーニング・プロポーショナルメトリクス）を収集し、
重複を除いた一覧をパレットに表示する常駐スクリプトです。同じフォントでも
これらの設定が異なれば、別の候補としてリストアップします。

一覧は「使用数／フォント／サイズ／行送り／自動カーニング／文字ツメ／トラッキング／
プロポーショナル」の8列で表示します（プロポーショナルメトリクスは ON/OFF）。
使用数は、その組み合わせを含むテキストフレームの数（1フレーム内で複数回使っても1）。

例）
  使用数  フォント               サイズ   行送り  自動カーニング  文字ツメ  トラッキング  プロポーショナル
  3       ヒラギノ角ゴシック W3   12pt     16pt    メトリクス       0         0             ON
  1       ヒラギノ角ゴシック W3   10.5pt   14pt    メトリクス       10        25            OFF

- パレットは上部に横並び・中央寄せのチェックボックス（「ロックされたテキストを含む」
  「非表示のテキストを含む」「現在のアートボードに限定」）、列付きリスト、
  その下のボタンエリア（左：「クリックで適用」チェックボックス／中央：スペーサー／
  右：「リストを更新」「選択した条件のフォントを選択」ボタン）で構成
- 「クリックで適用」が ON のとき、行をクリックすると選択中のテキストへその組み合わせ
  （フォント・サイズ・行送り・自動カーニング・文字ツメ・トラッキング・
  プロポーショナルメトリクス）をまとめて適用。OFF のあいだはクリックしても適用しない
- 「選択した条件のフォントを選択」ボタンで、その条件に一致する文字を含むテキスト
  フレームをドキュメント上で選択する（押すと「クリックで適用」は自動 OFF）
- 「リストを更新」ボタンで一覧をその場で再走査する
- 「現在のアートボードに限定」を ON にすると、走査・条件一致選択を現在のアートボードと
  重なるテキストフレームだけに絞る（OFF はドキュメント全体）。適用先は常に現在の選択
- 「非表示のテキストを含む」が OFF のあいだは、非表示のテキスト（hidden なオブジェクト・
  グループや、非表示レイヤー上のもの）を走査・条件一致選択から除外する（ON で含める）
- 「ロックされたテキストを含む」が OFF のあいだは、ロックされたテキスト（locked な
  オブジェクト・グループや、ロックレイヤー上のもの）を走査・条件一致選択から除外する（ON で含める）
- 一覧の並び順は フォント名→サイズ→行送り→自動カーニング→文字ツメ→トラッキング→
  プロポーショナル（使用数では並べ替えない）
- 一覧はパレット表示時・「リストを更新」・各チェックボックス切り替え時に再走査する
- 常駐エンジン（#targetengine）でパレット表示。常駐エンジンの app は
  パレット表示中に DOM 接続を失うため、DOM 処理はメインエンジンへ
  BridgeTalk で都度委譲する（コードは encodeURIComponent で包んで送信）

### Overview

A docked palette that lists the text-composition combinations used in the
document (font, size, leading, Tsume, tracking, auto-kerning and proportional
metrics), deduped. Even for the same font, a different combination of these is
listed as a separate candidate.

The list is shown as eight columns: Count / Font / Size / Leading / Auto Kerning /
Tsume / Tracking / Prop. Metrics (proportional metrics shown as ON/OFF). Count is
the number of text frames that use the combination (a frame counts once).

- The palette is a centered row of top checkboxes ("Include locked text" /
  "Include hidden text" / "Current artboard only"), the column list, and
  a button area below it (left: an "Apply on click" checkbox / center: a spacer /
  right: "Refresh list" and "Select matching text" buttons)
- While "Apply on click" is ON, clicking a row applies that combination (font,
  size, leading, auto kerning, Tsume, tracking, proportional metrics) to the
  current selection; while OFF, clicking does not apply
- The "Select matching text" button selects the text frames containing
  characters that match the selected condition (this turns "Apply on click" OFF)
- The "Refresh list" button rescans the list on demand
- Turning on "Current artboard only" limits the scan and matching selection to
  text frames overlapping the active artboard (off = whole doc); the apply target
  is always the current selection
- While "Include hidden text" is OFF, hidden text (hidden items/groups, or items
  on hidden layers) is excluded from the scan and matching selection (ON includes it)
- While "Include locked text" is OFF, locked text (locked items/groups, or items
  on locked layers) is excluded from the scan and matching selection (ON includes it)
- The list is sorted by font, then size, leading, auto kerning, Tsume, tracking,
  and prop. metrics (not by count)
- The list is rescanned when shown, on "Refresh list", and when a checkbox is toggled
- Runs as a persistent palette (#targetengine). The persistent engine's app
  loses its DOM connection while the palette is shown, so all DOM work is
  delegated to the main engine via BridgeTalk (code wrapped in encodeURIComponent)

*/

(function () {

    // =========================================
    // バージョン / Version
    // =========================================
    var SCRIPT_VERSION = "v1.1.3";

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
            currentArtboardOnly: { ja: "現在のアートボードに限定", en: "Current artboard only" },
            includeHidden: { ja: "非表示のテキストを含む", en: "Include hidden text" },
            includeLocked: { ja: "ロックされたテキストを含む", en: "Include locked text" }
        },
        button: {
            refresh: { ja: "リストを更新", en: "Refresh list" },
            selectMatching: { ja: "選択した条件のフォントを選択", en: "Select matching text" }
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

    /* 現在のアートボードの矩形を取得（取れなければ null）/ Get the active artboard rect ([L,T,R,B]); null if unavailable */
    function getActiveArtboardRect(activeDocument) {
        try {
            var activeIndex = activeDocument.artboards.getActiveArtboardIndex();
            return activeDocument.artboards[activeIndex].artboardRect;
        } catch (e) {
            return null;
        }
    }

    /* テキストフレームがアートボード矩形と重なるか / Whether a text frame overlaps the artboard rect
       Illustrator の y 軸は上が大きい（top > bottom）。矩形が無ければ常に true */
    function frameOverlapsArtboard(textFrame, artboardRect) {
        if (!artboardRect) return true;
        var bounds;
        try { bounds = textFrame.geometricBounds; } catch (e) { return false; }
        var frameLeft = bounds[0], frameTop = bounds[1], frameRight = bounds[2], frameBottom = bounds[3];
        var artLeft = artboardRect[0], artTop = artboardRect[1], artRight = artboardRect[2], artBottom = artboardRect[3];
        if (frameRight < artLeft || frameLeft > artRight) return false;
        if (frameBottom > artTop || frameTop < artBottom) return false;
        return true;
    }

    /* アイテムが表示されているか（自身と親グループの hidden、所属レイヤーの visible を遡って確認）/ Whether an item is visible
       自身・親グループの hidden、所属レイヤーの visible を親方向にたどって判定 / Walks up parents checking item.hidden and layer.visible */
    function isItemVisible(pageItem) {
        var node = pageItem;
        try {
            while (node) {
                var nodeType = node.typename;
                if (nodeType === "Document") break;
                if (nodeType === "Layer") {
                    if (!node.visible) return false;
                } else {
                    if (node.hidden) return false;
                }
                node = node.parent;
            }
        } catch (e) { }
        return true;
    }

    /* アイテムがロックされていないか（自身と親グループの locked、所属レイヤーの locked を遡って確認）/ Whether an item is unlocked
       自身・親グループの locked、所属レイヤーの locked を親方向にたどって判定 / Walks up parents checking item.locked and layer.locked */
    function isItemUnlocked(pageItem) {
        var node = pageItem;
        try {
            while (node) {
                var nodeType = node.typename;
                if (nodeType === "Document") break;
                if (node.locked) return false;
                node = node.parent;
            }
        } catch (e) { }
        return true;
    }

    /* ドキュメントを走査し、組み合わせを TAB 区切り行（LF 連結）で返す / Scan the document; return combos as TAB-joined rows (LF-joined)
       currentArtboardOnly が true なら現在のアートボードと重なるフレームだけ対象 / If currentArtboardOnly, only frames overlapping the active artboard
       includeHidden が false なら非表示のテキストフレームは除外 / If includeHidden is false, skip hidden text frames
       includeLocked が false ならロックされたテキストフレームは除外 / If includeLocked is false, skip locked text frames
       使用数 = その組み合わせを含むテキストフレームの数（1フレーム内で複数回使っても1）/ count = number of text frames that use the combo (a frame counts once)
       1行 = psName \t displayName \t size \t leading \t autoLeading(0/1) \t tsume \t tracking \t kernId \t proportionalMetrics(0/1) \t count */
    function collectDocumentCombosRaw(activeDocument, currentArtboardOnly, includeHidden, includeLocked) {
        var comboOrder = [], comboData = {}, comboRows = [];
        var TAB = String.fromCharCode(9), LF = String.fromCharCode(10);
        var artboardRect = currentArtboardOnly ? getActiveArtboardRect(activeDocument) : null;
        var textFrames = activeDocument.textFrames;
        for (var f = 0; f < textFrames.length; f++) {
            if (!includeHidden && !isItemVisible(textFrames[f])) continue;
            if (!includeLocked && !isItemUnlocked(textFrames[f])) continue;
            if (artboardRect && !frameOverlapsArtboard(textFrames[f], artboardRect)) continue;
            var characters;
            try { characters = textFrames[f].characters; } catch (eFrame) { continue; }
            var seenInFrame = {};
            for (var c = 0; c < characters.length; c++) {
                var combo = readComboFromCharacter(characters[c]);
                if (!combo) continue;
                var dedupeKey = comboKey(combo);
                if (!comboData[dedupeKey]) {
                    comboData[dedupeKey] = { combo: combo, count: 0 };
                    comboOrder.push(dedupeKey);
                }
                // 同じフレーム内で同じ組み合わせを複数回使っても使用数は1だけ増やす / Count each frame once per combo
                if (!seenInFrame[dedupeKey]) {
                    seenInFrame[dedupeKey] = true;
                    comboData[dedupeKey].count++;
                }
            }
        }
        for (var i = 0; i < comboOrder.length; i++) {
            var entry = comboData[comboOrder[i]];
            var entryCombo = entry.combo;
            comboRows.push([
                entryCombo.psName, entryCombo.displayName, String(entryCombo.size), String(entryCombo.leading),
                entryCombo.autoLeading ? "1" : "0", String(entryCombo.tsume), String(entryCombo.tracking), entryCombo.kernId,
                entryCombo.proportionalMetrics ? "1" : "0", String(entry.count)
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
            // 属性は同じ characterAttributes に連続代入するので try は1範囲につき1つ
            // All attributes go on the same characterAttributes, so one try per range
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
            } catch (eApplyRange) { }
        }
        return appliedCount;
    }

    /* 組み合わせに一致する文字を含むテキストフレームを選択 / Select text frames containing characters that match the combo
       戻り値: 選択したフレーム数 / Returns the number of selected frames */
    function selectFramesMatchingCombo(activeDocument, combo, currentArtboardOnly, includeHidden, includeLocked) {
        var targetKey = comboKey(combo);
        var artboardRect = currentArtboardOnly ? getActiveArtboardRect(activeDocument) : null;
        var textFrames = activeDocument.textFrames;
        var matchedFrames = [];
        for (var f = 0; f < textFrames.length; f++) {
            if (!includeHidden && !isItemVisible(textFrames[f])) continue;
            if (!includeLocked && !isItemUnlocked(textFrames[f])) continue;
            if (artboardRect && !frameOverlapsArtboard(textFrames[f], artboardRect)) continue;
            var characters;
            try { characters = textFrames[f].characters; } catch (eFrame) { continue; }
            for (var c = 0; c < characters.length; c++) {
                var charCombo = readComboFromCharacter(characters[c]);
                if (charCombo && comboKey(charCombo) === targetKey) { matchedFrames.push(textFrames[f]); break; }
            }
        }
        // 一致0件のときは現在の選択を保持（消してから「なし」と出さない）/ Keep the current selection when nothing matched
        if (matchedFrames.length === 0) return 0;
        // 配列を直接代入すれば既存選択を一括置換できる / Assigning the array replaces the current selection wholesale
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
        getActiveArtboardRect, frameOverlapsArtboard, isItemVisible, isItemUnlocked,
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
       結果は "OK:<payload>" または "ERR:<msg>" の文字列で返す */
    function dispatch(actionId, params) {
        if (app.documents.length === 0) return "ERR:nodoc";
        try { app.activeDocument; } catch (e) { return "ERR:nodoc"; }

        if (actionId === "getCombos") {
            return "OK:" + encodeURIComponent(collectDocumentCombosRaw(app.activeDocument, params && params.currentArtboardOnly, params && params.includeHidden, params && params.includeLocked));
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
            var matchedCount = selectFramesMatchingCombo(app.activeDocument, params, params && params.currentArtboardOnly, params && params.includeHidden, params && params.includeLocked);
            app.redraw();
            return "OK:" + matchedCount;
        }
        return "OK:0";
    }
    var DISPATCH_SOURCE = "(" + dispatch.toString() + ")";

    /* パラメータを安全な JS リテラル文字列へ / Serialize params to a safe JS literal
       params はフラグだけ（getCombos）か comboToParams（全フィールド＋フラグ）か null / flag-only, full combo+flag, or null
       currentArtboardOnly / includeHidden は常に出力する / the flags are always emitted */
    function paramsToSource(params) {
        if (!params) return "{}";
        function numberLiteral(value) { return isNaN(value) ? "NaN" : value; }
        function stringLiteral(value) { return 'decodeURIComponent("' + encodeURIComponent(value) + '")'; }
        function boolLiteral(value) { return value ? "true" : "false"; }
        var fields = [];
        if (params.psName !== undefined) {
            fields.push("psName:" + stringLiteral(params.psName));
            fields.push("kernId:" + stringLiteral(params.kernId));
            fields.push("size:" + numberLiteral(params.size));
            fields.push("leading:" + numberLiteral(params.leading));
            fields.push("tsume:" + (isNaN(params.tsume) ? 0 : params.tsume));
            fields.push("tracking:" + numberLiteral(params.tracking));
            fields.push("autoLeading:" + boolLiteral(params.autoLeading));
            fields.push("proportionalMetrics:" + boolLiteral(params.proportionalMetrics));
        }
        fields.push("currentArtboardOnly:" + boolLiteral(params.currentArtboardOnly));
        fields.push("includeHidden:" + boolLiteral(params.includeHidden));
        fields.push("includeLocked:" + boolLiteral(params.includeLocked));
        return "{" + fields.join(",") + "}";
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

    /* メインエンジンへアクションを委譲（非同期）/ Delegate an action to the main engine (async) */
    function runWorker(actionId, params, onDone) {
        var workerCode = getWorkerSource() + "\nvar __result=" + DISPATCH_SOURCE + "(\"" + actionId + "\"," + paramsToSource(params) + ");__result;";
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

    /* 各列のセル文字列を組み立て / Build the per-column cell strings
       [使用数, フォント, サイズ, 行送り, 自動カーニング, 文字ツメ, トラッキング, プロポーショナル] */
    function comboCells(combo) {
        var countText = isNaN(combo.count) ? "?" : String(combo.count);
        var sizeText = formatNumber(combo.size) + "pt";
        var leadingText = combo.autoLeading
            ? getLocalizedText(LABELS.autoLeading)
            : formatNumber(combo.leading) + "pt";
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
                count: parseInt(fields[9], 10)
            };
            combo.cells = comboCells(combo);
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
        var columnWidths = [50, 160, 60, 70, 90, 70, 70, 70];

        // 最上部：走査オプションのチェックボックスを横並び・左右中央に配置 / Topmost: scan-option checkboxes in one row, centered horizontally
        //   ロックされたテキストを含む（OFF ならロックされたフレームを除外）/ Include locked text (off = skip locked frames)
        //   非表示のテキストを含む（OFF なら非表示フレームを除外）/ Include hidden text (off = skip hidden frames)
        //   現在のアートボードに限定 / Limit the scan to the active artboard
        var topRow = palette.add("group");
        topRow.orientation = "row";
        topRow.alignment = ["center", "top"];
        topRow.spacing = 16;
        // 下マージンを +5（チェックボックス行とリストの間隔を少し広げる）/ +5 bottom margin (a bit more gap below the checkboxes)
        topRow.margins = [0, 0, 0, 5];

        var includeLockedCheckbox = topRow.add("checkbox", undefined, getLocalizedText(LABELS.control.includeLocked));
        includeLockedCheckbox.value = true;

        var includeHiddenCheckbox = topRow.add("checkbox", undefined, getLocalizedText(LABELS.control.includeHidden));
        includeHiddenCheckbox.value = true;

        var currentArtboardCheckbox = topRow.add("checkbox", undefined, getLocalizedText(LABELS.control.currentArtboardOnly));
        currentArtboardCheckbox.value = false;

        var comboListBox = palette.add("listbox", undefined, [], {
            multiselect: false,
            numberOfColumns: columnTitles.length,
            showHeaders: true,
            columnTitles: columnTitles,
            columnWidths: columnWidths
        });
        comboListBox.preferredSize = [640, 300];

        // ボタンエリア：左＝クリックで適用 / 中央＝スペーサー / 右＝条件一致テキストを選択
        // Button area: left = apply-on-click / center = spacer / right = select matching text
        var buttonRow = palette.add("group");
        buttonRow.orientation = "row";
        buttonRow.alignment = ["fill", "center"];
        // 上マージンを +10（リストとボタンの間隔を少し広げる）/ +10 top margin (a bit more gap above the buttons)
        buttonRow.margins = [0, 10, 0, 0];

        // 左：クリックで適用するか（OFF のあいだはクリックしても適用しない）/ Left: whether a click applies (no apply while OFF)
        var applyOnClickCheckbox = buttonRow.add("checkbox", undefined, getLocalizedText(LABELS.control.applyOnClick));
        applyOnClickCheckbox.alignment = ["left", "center"];
        applyOnClickCheckbox.value = true;

        // 中央：スペーサー（伸びて左右を端へ押し出す。空グループが幅0に潰れないよう最小幅を持たせる）
        // Center: spacer that stretches to push the sides apart (minimum width so an empty group won't collapse to 0)
        var spacer = buttonRow.add("group");
        spacer.alignment = ["fill", "center"];
        spacer.minimumSize = [10, 1];

        // 右：一覧を再走査 / Right: rescan the list
        var refreshButton = buttonRow.add("button", undefined, getLocalizedText(LABELS.button.refresh));
        refreshButton.alignment = ["right", "center"];

        // 右：選択中の条件に一致するテキストを選択 / Right: select text matching the selected condition
        var selectMatchingButton = buttonRow.add("button", undefined, getLocalizedText(LABELS.button.selectMatching));
        selectMatchingButton.alignment = ["right", "center"];

        return {
            palette: palette,
            comboListBox: comboListBox,
            includeLockedCheckbox: includeLockedCheckbox,
            includeHiddenCheckbox: includeHiddenCheckbox,
            currentArtboardCheckbox: currentArtboardCheckbox,
            applyOnClickCheckbox: applyOnClickCheckbox,
            refreshButton: refreshButton,
            selectMatchingButton: selectMatchingButton
        };
    }

    /* パレットにイベントを接続 / Wire palette events */
    function bindPaletteEvents(paletteUI) {
        // BridgeTalk は非同期なので、連打による委譲の重複を抑止 / Guard against overlapping async delegations
        var applyInProgress = false;
        var currentCombos = [];

        /* 組み合わせ一覧を読み直して listbox に反映 / Reload the combos into the listbox */
        function refreshCombos() {
            runWorker("getCombos", { currentArtboardOnly: paletteUI.currentArtboardCheckbox.value, includeHidden: paletteUI.includeHiddenCheckbox.value, includeLocked: paletteUI.includeLockedCheckbox.value }, function (status, payload) {
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

        // ---- シングルクリックで選択テキストへ適用 / A single click applies to the selection ----
        // 適用先は委譲時にメインエンジンが現在の選択を読む。選択が無ければワーカーが
        // "OK:nosel" を返すので、その場合だけ警告する。
        // The main engine reads the live selection at apply time; if nothing is
        // selected it returns "OK:nosel", and only then do we warn.
        paletteUI.comboListBox.onChange = function () {
            if (!paletteUI.comboListBox.selection) return;
            // 「クリックで適用」が OFF なら何もしない / Do nothing while "Apply on click" is OFF
            if (!paletteUI.applyOnClickCheckbox.value) return;
            if (applyInProgress) return;
            var combo = currentCombos[paletteUI.comboListBox.selection.comboIndex];
            if (!combo) return;
            applyInProgress = true;
            runWorker("applyCombo", comboToParams(combo), function (status, payload) {
                applyInProgress = false;
                // 選択が無いときだけ警告（適用0件＝ロック範囲などとは区別）/ Warn only when nothing was selected (distinct from 0 applied)
                if (status === "ok" && payload === "nosel") {
                    alert(getLocalizedText(LABELS.status.needSelection));
                }
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
            var matchingParams = comboToParams(combo);
            matchingParams.currentArtboardOnly = paletteUI.currentArtboardCheckbox.value;
            matchingParams.includeHidden = paletteUI.includeHiddenCheckbox.value;
            matchingParams.includeLocked = paletteUI.includeLockedCheckbox.value;
            runWorker("selectMatching", matchingParams, function (status, payload) {
                if (status === "ok" && payload === "0") {
                    alert(getLocalizedText(LABELS.status.noMatch));
                }
            });
        };

        // 「リストを更新」で一覧を再走査 / Rescan the list on demand
        paletteUI.refreshButton.onClick = function () { refreshCombos(); };

        // 「現在のアートボードに限定」を切り替えたら一覧を再走査 / Rescan when the artboard-only toggle changes
        paletteUI.currentArtboardCheckbox.onClick = function () { refreshCombos(); };

        // 「非表示のテキストを含む」を切り替えたら一覧を再走査 / Rescan when the include-hidden toggle changes
        paletteUI.includeHiddenCheckbox.onClick = function () { refreshCombos(); };

        // 「ロックされたテキストを含む」を切り替えたら一覧を再走査 / Rescan when the include-locked toggle changes
        paletteUI.includeLockedCheckbox.onClick = function () { refreshCombos(); };

        // 表示のたびに一覧を再走査 / Rescan the combos whenever the palette is shown
        paletteUI.palette.onShow = function () { refreshCombos(); };

        // Esc でパレットを閉じる / Close the palette on Esc
        paletteUI.palette.addEventListener("keydown", function (keyEvent) {
            if (keyEvent.keyName === "Escape") paletteUI.palette.close();
        });
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
