#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

(function () {

    // =========================================
    // バージョンとローカライズ
    // =========================================

    /* スクリプトバージョン / Script version */
    var SCRIPT_VERSION = "v1.0";

    /* ロケール判定 / Locale detection */
    function getCurrentLang() {
        return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    /* ラベル定義 / Label definitions */
    var LABELS = {
        dialogTitle: { ja: "シンボル一覧を配置", en: "List All Symbols" },
        placePanel: { ja: "アートボードの位置", en: "Artboard position" },
        placeRight: { ja: "最終アートボードの右側", en: "Right of last artboard" },
        placeBelow: { ja: "最終アートボードの下側", en: "Below last artboard" },
        artboardPanel: { ja: "アートボード", en: "Artboard" },
        symbolPanel: { ja: "シンボルの配置", en: "Symbol placement" },
        gap: { ja: "間隔", en: "Gap" },
        margin: { ja: "マージン", en: "Margin" },
        update: { ja: "更新", en: "Update" },
        tipUpdate: { ja: "既存の「シンボル一覧」アートボードがあれば置き換え", en: "Replace existing Symbol List artboard if present" },
        showCaption: { ja: "シンボル名を表示", en: "Show symbol names" },
        tipShowCaption: { ja: "各シンボルの下にシンボル名を表示", en: "Show symbol name below each symbol" },
        filterAll: { ja: "すべて", en: "All" },
        filterUsed: { ja: "使用中のみ", en: "Used only" },
        tipFilterAll: { ja: "登録されているすべてのシンボルを並べる", en: "List every registered symbol" },
        tipFilterUsed: { ja: "ドキュメント内で実際に使用されているシンボルのみ", en: "Only symbols actually placed in the document" },
        cancel: { ja: "キャンセル", en: "Cancel" },
        tipPlaceRight: { ja: "末尾アートボードの右に新規アートボードを作成", en: "Place the new artboard to the right of the last one" },
        tipPlaceBelow: { ja: "末尾アートボードの下に新規アートボードを作成", en: "Place the new artboard below the last one" },
        tipArtboardGap: { ja: "既存アートボードと新規アートボードの間隔", en: "Distance between the last and new artboards" },
        tipMargin: { ja: "新規アートボード内側に設ける余白", en: "Padding inside the new artboard around the symbols" },
        tipSymbolGap: { ja: "シンボル同士の間隔", en: "Spacing between adjacent symbols" },
        maxRowWidth: { ja: "最大幅", en: "Max row width" },
        tipMaxRowWidth: { ja: "1行に並べる最大幅。超えたら折り返し", en: "Max width per row before wrapping" },
        noDoc: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        noSymbols: { ja: "登録されているシンボルがありません。", en: "No symbols are registered." },
        noUsedSymbols: { ja: "ドキュメント内で使用中のシンボルがありません。", en: "No symbols are currently used in the document." }
    };
    /* 現在の言語の文字列を取得（不足時は英語→日本語→空文字にフォールバック） / Get localized text with fallback */
    function L(labelObj) {
        if (!labelObj) return "";
        return labelObj[currentLanguage] || labelObj.en || labelObj.ja || "";
    }

    /* コロン付きラベル（日本語は全角、英語は半角） / Label with colon (full-width JA, half-width EN) */
    function labelText(labelObj) {
        return L(labelObj) + (currentLanguage === "ja" ? "：" : ":");
    }

    // =========================================
    // 設定値
    // =========================================

    /* レイアウト初期値（ポイント単位） / Layout defaults (points) */
    var DEFAULT_MAX_ROW_WIDTH_PT = 800;
    var DEFAULT_SYMBOL_GAP_PT = 20;
    var DEFAULT_ARTBOARD_GAP_PT = 100;
    var DEFAULT_ARTBOARD_MARGIN_PT = 50;
    var CAPTION_FONT_SIZE_PT = 9;
    var CAPTION_GAP_PT = 6;
    var LAYER_NAME = "シンボル一覧";
    var ARTBOARD_NAME = "シンボル一覧";

    /* 設定保存用キー / Preference key for saved settings */
    var PREF_KEY = "swwwitch.listupallsymbol.settings";

    /* rulerType から表示単位とポイント変換係数を取得
     * Resolve display unit & point conversion from rulerType */
    function getRulerUnitInfo() {
        var rulerUnit = app.preferences.getIntegerPreference("rulerType");
        switch (rulerUnit) {
            case 0: return { label: "inch", factor: 72.0 };
            case 1: return { label: "mm", factor: 72.0 / 25.4 };
            case 3: return { label: "pica", factor: 12.0 };
            case 4: return { label: "cm", factor: 72.0 / 2.54 };
            case 5: return { label: "Q", factor: 72.0 / 25.4 * 0.25 };
            case 6: return { label: "px", factor: 1.0 };
            default: return { label: "pt", factor: 1.0 };
        }
    }
    var UNIT = getRulerUnitInfo();

    function ptToUnit(pt) { return pt / UNIT.factor; }
    function unitToPt(value) { return value * UNIT.factor; }

    // =========================================
    // 設定の保存・復元
    // =========================================

    /* 設定オブジェクトを文字列化（手動 JSON, ES3 互換）
     * Serialize settings object (manual JSON, ES3 compatible) */
    function serializeSettings(settings) {
        return [
            '{',
            '"position":"' + settings.position + '",',
            '"artboardGap":' + settings.artboardGap + ',',
            '"margin":' + settings.margin + ',',
            '"update":' + (settings.update ? "true" : "false") + ',',
            '"showCaption":' + (settings.showCaption ? "true" : "false") + ',',
            '"filter":"' + settings.filter + '",',
            '"symbolGap":' + settings.symbolGap + ',',
            '"maxRowWidth":' + settings.maxRowWidth,
            '}'
        ].join('');
    }

    /* 文字列から設定オブジェクトに復元（eval 不使用） / Parse settings string without eval */
    function parseSettings(str) {
        if (!str) return null;

        function readString(key) {
            var match = str.match(new RegExp('"' + key + '"\\s*:\\s*"([^"\\\\]*)"'));
            return match ? match[1] : null;
        }

        function readNumber(key) {
            var match = str.match(new RegExp('"' + key + '"\\s*:\\s*(-?\\d+(?:\\.\\d+)?)'));
            return match ? Number(match[1]) : null;
        }

        function readBoolean(key) {
            var match = str.match(new RegExp('"' + key + '"\\s*:\\s*(true|false)'));
            return match ? (match[1] === "true") : null;
        }

        var position = readString("position");
        var filter = readString("filter");
        var artboardGap = readNumber("artboardGap");
        var margin = readNumber("margin");
        var symbolGap = readNumber("symbolGap");
        var maxRowWidth = readNumber("maxRowWidth");
        var update = readBoolean("update");
        var showCaption = readBoolean("showCaption");

        if ((position !== "right" && position !== "below") ||
            (filter !== "all" && filter !== "used") ||
            artboardGap === null || margin === null ||
            symbolGap === null || maxRowWidth === null ||
            update === null || showCaption === null) {
            return null;
        }

        return {
            position: position,
            artboardGap: artboardGap,
            margin: margin,
            update: update,
            showCaption: showCaption,
            filter: filter,
            symbolGap: symbolGap,
            maxRowWidth: maxRowWidth
        };
    }

    function saveSettings(settings) {
        try { app.preferences.setStringPreference(PREF_KEY, serializeSettings(settings)); } catch (e) { }
    }

    function loadSettings() {
        try { return parseSettings(app.preferences.getStringPreference(PREF_KEY)); } catch (e) { return null; }
    }

    /* 設定値の取得（保存値 → 既定値の順） / Resolve setting (saved → default) */
    function getSavedSetting(savedSettings, key, defaultValue) {
        if (!savedSettings) return defaultValue;
        var savedValue = savedSettings[key];
        return (typeof savedValue === "undefined" || savedValue === null) ? defaultValue : savedValue;
    }

    // =========================================
    // UI ヘルパー
    // =========================================

    /* パネル共通設定 / Panel layout helpers */
    var PANEL_MARGINS = [15, 20, 15, 10];
    var VALUE_ROW_LABEL_WIDTH = 70;

    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = "left";
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        if (typeof spacing === "number") panel.spacing = spacing;
    }

    /* ↑↓キーで値を増減
     * - ↑↓: ±1
     * - Shift+↑↓: ±10（10の倍数にスナップ） */
    function changeValueByArrowKey(editText, allowNegative, onChange) {
        editText.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;
            var current = Number(editText.text);
            if (isNaN(current)) return;

            var sign = (event.keyName === "Up") ? 1 : -1;
            var step = ScriptUI.environment.keyboardState.shiftKey ? 10 : 1;
            var next = (step === 10)
                ? (sign > 0 ? Math.ceil((current + 1) / step) * step
                    : Math.floor((current - 1) / step) * step)
                : current + sign * step;

            next = Math.round(next);
            if (!allowNegative && next < 0) next = 0;

            editText.text = next;
            event.preventDefault();
            if (typeof onChange === "function") onChange();
        });
    }

    // =========================================
    // レイアウト処理
    // =========================================

    /* ドキュメント内で使用中のシンボル名セットを返す（管理レイヤーは除外）
     * Return a set of symbol names currently placed in the document (excluding our layer) */
    function getUsedSymbolNames(doc) {
        var nameSet = {};
        var items = doc.symbolItems;
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            try {
                if (item.layer && item.layer.name === LAYER_NAME) continue;
            } catch (e) { }
            if (item.symbol && item.symbol.name) nameSet[item.symbol.name] = true;
        }
        return nameSet;
    }

    /* フィルタ条件で対象シンボル配列を返す / Resolve target symbols by filter */
    function getTargetSymbols(doc, settings) {
        var symbols = doc.symbols;
        if (settings.filter !== "used") {
            var all = [];
            for (var i = 0; i < symbols.length; i++) all.push(symbols[i]);
            return all;
        }
        var usedNames = getUsedSymbolNames(doc);
        var filtered = [];
        for (var j = 0; j < symbols.length; j++) {
            if (usedNames[symbols[j].name]) filtered.push(symbols[j]);
        }
        return filtered;
    }

    /* 確定後に他の「シンボル一覧」アートボード／レイヤーを削除（state は保持）
     * Remove other symbol-list artboards/layers, keeping the one we just built */
    function removeOtherSymbolLists(doc, state) {
        if (!state || !state.artboardInfo) return;

        var keepIndex = findArtboardIndexByState(doc, state);
        if (keepIndex < 0) keepIndex = state.artboardInfo.index;

        for (var i = doc.artboards.length - 1; i >= 0; i--) {
            if (i === keepIndex) continue;
            if (doc.artboards[i].name === ARTBOARD_NAME && doc.artboards.length > 1) {
                try {
                    doc.artboards.remove(i);
                    if (i < keepIndex) keepIndex--;
                } catch (artboardRemoveError) { }
            }
        }
        state.artboardInfo.index = keepIndex;

        for (var j = doc.layers.length - 1; j >= 0; j--) {
            var layer = doc.layers[j];
            if (layer === state.layer) continue;
            if (layer.name === LAYER_NAME && doc.layers.length > 1) {
                try { layer.remove(); } catch (layerRemoveError) { }
            }
        }
    }

    /* シンボル配置用のレイヤーを作成（失敗時は中断用に null を返す）
     * Create the layer for the symbol list; return null on failure */
    function createSymbolLayer(doc) {
        try {
            var layer = doc.layers.add();
            layer.name = LAYER_NAME;
            return layer;
        } catch (layerCreateError) {
            return null;
        }
    }

    /* 1 シンボル分のキャプション TextFrame を生成 / Create a caption text frame */
    function createCaption(layer, symbolName) {
        var caption = layer.textFrames.add();
        caption.contents = symbolName;
        try {
            caption.textRange.characterAttributes.size = CAPTION_FONT_SIZE_PT;
        } catch (e) { }
        return caption;
    }

    /* シンボルをインスタンス化し、必要ならキャプションも作成、エントリ配列を返す
     * Instantiate symbols (and optional captions), measure, return entries */
    function instantiateSymbols(doc, layer, settings, targetSymbols) {
        var entries = [];
        for (var i = 0; i < targetSymbols.length; i++) {
            var symbol = targetSymbols[i];
            var symbolItem = doc.symbolItems.add(symbol);
            symbolItem.moveToBeginning(layer);
            var bounds = symbolItem.visibleBounds; // [left, top, right, bottom]
            var symbolWidth = bounds[2] - bounds[0];
            var symbolHeight = bounds[1] - bounds[3];

            var entry = {
                item: symbolItem,
                symbolWidth: symbolWidth,
                symbolHeight: symbolHeight
            };

            if (settings.showCaption) {
                var caption = createCaption(layer, symbol.name);
                var capBounds = caption.visibleBounds;
                entry.caption = caption;
                entry.captionWidth = capBounds[2] - capBounds[0];
                entry.captionHeight = capBounds[1] - capBounds[3];
                entry.width = Math.max(symbolWidth, entry.captionWidth);
                entry.height = symbolHeight + CAPTION_GAP_PT + entry.captionHeight;
            } else {
                entry.width = symbolWidth;
                entry.height = symbolHeight;
            }
            entries.push(entry);
        }
        return entries;
    }

    /* シェルフパッキングで各エントリに x,y を割り付け、合計サイズを返す / Shelf packing */
    function packShelf(entries, maxWidth, gap) {
        var rowX = 0, rowY = 0, rowHeight = 0, totalWidth = 0;
        for (var i = 0; i < entries.length; i++) {
            var entry = entries[i];
            if (rowX > 0 && rowX + entry.width > maxWidth) {
                rowY -= rowHeight + gap;
                rowX = 0;
                rowHeight = 0;
            }
            entry.x = rowX;
            entry.y = rowY;
            rowX += entry.width + gap;
            if (rowX - gap > totalWidth) totalWidth = rowX - gap;
            if (entry.height > rowHeight) rowHeight = entry.height;
        }
        return { width: totalWidth, height: -rowY + rowHeight };
    }

    /* 新アートボードの基準座標（左上）を決定。
     * - 更新 ON: 既存「シンボル一覧」は OK 時に消えるので基準から除外（最後の通常アートボードを基準）
     * - 更新 OFF: 既存「シンボル一覧」を含むドキュメント末尾のアートボードを基準（その右／下に追加）
     * Compute new artboard origin:
     *   - update=ON: skip シンボル一覧 artboards (they will be removed at OK)
     *   - update=OFF: use the very last artboard (so new is placed beyond existing) */
    function computeArtboardOrigin(doc, settings) {
        var lastRect = null;
        if (!settings.update) {
            lastRect = doc.artboards[doc.artboards.length - 1].artboardRect;
        } else {
            for (var i = doc.artboards.length - 1; i >= 0; i--) {
                if (doc.artboards[i].name !== ARTBOARD_NAME) {
                    lastRect = doc.artboards[i].artboardRect;
                    break;
                }
            }
            if (!lastRect) lastRect = doc.artboards[doc.artboards.length - 1].artboardRect;
        }
        if (settings.position === "right") {
            return { left: lastRect[2] + settings.artboardGap, top: lastRect[1] };
        }
        return { left: lastRect[0], top: lastRect[3] - settings.artboardGap };
    }

    /* アートボードを追加し、識別用情報を返す / Add artboard and return identifying info */
    function addArtboard(doc, left, top, width, height, name) {
        var index = doc.artboards.length;
        var rect = [left, top, left + width, top - height];
        var artboard = doc.artboards.add(rect);
        artboard.name = name;
        return { index: index, name: name, rect: rect };
    }

    /* パッキング結果に従って各エントリ（シンボル＋任意のキャプション）を配置
     * Place entries (symbol + optional caption) based on packing result */
    function placeEntries(entries, originLeft, originTop, margin) {
        for (var i = 0; i < entries.length; i++) {
            var entry = entries[i];
            var slotLeft = originLeft + margin + entry.x;
            var slotTop = originTop - margin + entry.y;

            /* スロット内でシンボルを水平中央寄せ / Center symbol within slot */
            var symbolLeft = slotLeft + (entry.width - entry.symbolWidth) / 2;
            entry.item.position = [symbolLeft, slotTop];

            if (entry.caption) {
                var captionLeft = slotLeft + (entry.width - entry.captionWidth) / 2;
                var captionTop = slotTop - entry.symbolHeight - CAPTION_GAP_PT;
                entry.caption.position = [captionLeft, captionTop];
            }
        }
    }

    /* ひとまとまりのレイアウトを構築（既存はプレビュー中常に保持、削除は OK 時のみ）
     * Build layout (existing symbol-list artboards are kept during preview; removal only at OK) */
    function buildLayout(doc, settings) {
        var targetSymbols = getTargetSymbols(doc, settings);
        if (targetSymbols.length === 0) return null;
        var layer = createSymbolLayer(doc);
        if (!layer) return null;
        var entries = instantiateSymbols(doc, layer, settings, targetSymbols);
        var size = packShelf(entries, settings.maxRowWidth, settings.symbolGap);

        var origin = computeArtboardOrigin(doc, settings);
        var artboardInfo = addArtboard(
            doc, origin.left, origin.top,
            size.width + settings.margin * 2,
            size.height + settings.margin * 2,
            ARTBOARD_NAME
        );
        placeEntries(entries, origin.left, origin.top, settings.margin);

        return { layer: layer, artboardInfo: artboardInfo, count: entries.length };
    }

    /* 数値を許容誤差つきで比較 / Compare numbers with tolerance */
    function nearlyEqual(a, b) {
        return Math.abs(a - b) < 0.01;
    }

    /* アートボード矩形を比較 / Compare artboard rectangles */
    function isSameArtboardRect(a, b) {
        if (!a || !b || a.length !== 4 || b.length !== 4) return false;
        for (var i = 0; i < 4; i++) {
            if (!nearlyEqual(a[i], b[i])) return false;
        }
        return true;
    }

    /* state から削除対象アートボードの現在 index を探す / Find current artboard index from state */
    function findArtboardIndexByState(doc, state) {
        if (!state || !state.artboardInfo) return -1;
        var fallbackIndex = state.artboardInfo.index;
        if (fallbackIndex >= 0 && fallbackIndex < doc.artboards.length) {
            try {
                var fallbackArtboard = doc.artboards[fallbackIndex];
                if (fallbackArtboard.name === state.artboardInfo.name &&
                    isSameArtboardRect(fallbackArtboard.artboardRect, state.artboardInfo.rect)) {
                    return fallbackIndex;
                }
            } catch (fallbackError) { }
        }

        for (var i = doc.artboards.length - 1; i >= 0; i--) {
            try {
                var artboard = doc.artboards[i];
                if (artboard.name === state.artboardInfo.name &&
                    isSameArtboardRect(artboard.artboardRect, state.artboardInfo.rect)) {
                    return i;
                }
            } catch (searchError) { }
        }
        return -1;
    }

    /* レイアウト結果を取り消す（プレビュー用、API 経由削除） / Remove layout via API */
    function clearLayout(doc, state) {
        if (!state) return;
        var artboardIndex = findArtboardIndexByState(doc, state);
        if (artboardIndex >= 0 && doc.artboards.length > 1) {
            try { doc.artboards.remove(artboardIndex); } catch (artboardRemoveError) { }
        }
        try { state.layer.remove(); } catch (layerRemoveError) { }
    }

    // =========================================
    // ダイアログ
    // =========================================

    /* ラベル付き数値入力行を作成 / Build a labeled numeric input row */
    function addValueRow(parent, label, defaultPt, tooltip) {
        var row = parent.add("group");
        row.orientation = "row";
        row.alignChildren = "center";
        row.spacing = 6;

        /* 固定幅・右寄せのラベル列（statictext の justify は不安定なため group で寄せる）
         * Fixed-width right-aligned label column (statictext justify is unreliable; align via group) */
        var labelColumn = row.add("group");
        labelColumn.orientation = "row";
        labelColumn.alignChildren = ["right", "center"];
        labelColumn.margins = 0;
        labelColumn.spacing = 0;
        labelColumn.preferredSize.width = VALUE_ROW_LABEL_WIDTH;
        var labelControl = labelColumn.add("statictext", undefined, label);

        var input = row.add("edittext", undefined, String(Math.round(ptToUnit(defaultPt))));
        input.characters = 4;
        row.add("statictext", undefined, UNIT.label);
        if (tooltip) {
            labelControl.helpTip = tooltip;
            input.helpTip = tooltip;
        }
        return input;
    }

    function buildPlacementPanel(parent, savedSettings) {
        var panel = parent.add("panel", undefined, L(LABELS.placePanel));
        setupPanel(panel, 6);
        var savedPosition = getSavedSetting(savedSettings, "position", "right");
        var rightRadio = panel.add("radiobutton", undefined, L(LABELS.placeRight));
        rightRadio.helpTip = L(LABELS.tipPlaceRight);
        var belowRadio = panel.add("radiobutton", undefined, L(LABELS.placeBelow));
        belowRadio.helpTip = L(LABELS.tipPlaceBelow);
        rightRadio.value = (savedPosition === "right");
        belowRadio.value = (savedPosition === "below");
        var artboardGapInput = addValueRow(panel, labelText(LABELS.gap),
            getSavedSetting(savedSettings, "artboardGap", DEFAULT_ARTBOARD_GAP_PT), L(LABELS.tipArtboardGap));
        return { rightRadio: rightRadio, belowRadio: belowRadio, artboardGapInput: artboardGapInput };
    }

    function buildArtboardPanel(parent, savedSettings) {
        var panel = parent.add("panel", undefined, L(LABELS.artboardPanel));
        setupPanel(panel, 6);
        var marginInput = addValueRow(panel, labelText(LABELS.margin),
            getSavedSetting(savedSettings, "margin", DEFAULT_ARTBOARD_MARGIN_PT), L(LABELS.tipMargin));
        var updateCheckbox = panel.add("checkbox", undefined, L(LABELS.update));
        updateCheckbox.value = getSavedSetting(savedSettings, "update", true);
        updateCheckbox.helpTip = L(LABELS.tipUpdate);
        return { marginInput: marginInput, updateCheckbox: updateCheckbox };
    }

    function buildSymbolPanel(parent, savedSettings) {
        var panel = parent.add("panel", undefined, L(LABELS.symbolPanel));
        setupPanel(panel, 6);

        /* 対象フィルタ / Filter radios */
        var filterGroup = panel.add("group");
        filterGroup.orientation = "row";
        filterGroup.spacing = 12;
        var filterAllRadio = filterGroup.add("radiobutton", undefined, L(LABELS.filterAll));
        filterAllRadio.helpTip = L(LABELS.tipFilterAll);
        var filterUsedRadio = filterGroup.add("radiobutton", undefined, L(LABELS.filterUsed));
        filterUsedRadio.helpTip = L(LABELS.tipFilterUsed);
        var savedFilter = getSavedSetting(savedSettings, "filter", "all");
        filterAllRadio.value = (savedFilter === "all");
        filterUsedRadio.value = (savedFilter === "used");

        /* キャプション表示 / Caption checkbox */
        var showCaptionCheckbox = panel.add("checkbox", undefined, L(LABELS.showCaption));
        showCaptionCheckbox.value = getSavedSetting(savedSettings, "showCaption", false);
        showCaptionCheckbox.helpTip = L(LABELS.tipShowCaption);

        var symbolGapInput = addValueRow(panel, labelText(LABELS.gap),
            getSavedSetting(savedSettings, "symbolGap", DEFAULT_SYMBOL_GAP_PT), L(LABELS.tipSymbolGap));
        var maxRowWidthInput = addValueRow(panel, labelText(LABELS.maxRowWidth),
            getSavedSetting(savedSettings, "maxRowWidth", DEFAULT_MAX_ROW_WIDTH_PT), L(LABELS.tipMaxRowWidth));

        return {
            filterAllRadio: filterAllRadio,
            filterUsedRadio: filterUsedRadio,
            showCaptionCheckbox: showCaptionCheckbox,
            symbolGapInput: symbolGapInput,
            maxRowWidthInput: maxRowWidthInput
        };
    }

    function buildButtonRow(parent) {
        var group = parent.add("group");
        group.alignment = "center";
        group.spacing = 8;
        var cancelButton = group.add("button", undefined, L(LABELS.cancel), { name: "cancel" });
        var okButton = group.add("button", undefined, "OK", { name: "OK" });
        return { okButton: okButton, cancelButton: cancelButton };
    }

    /* ダイアログを構築し、コントロール参照を返す / Build dialog and return refs */
    function buildDialog(savedSettings) {
        var dialog = new Window("dialog", L(LABELS.dialogTitle) + " " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = "fill";
        dialog.margins = 16;
        dialog.spacing = 12;

        var placementRefs = buildPlacementPanel(dialog, savedSettings);
        var artboardRefs = buildArtboardPanel(dialog, savedSettings);
        var symbolRefs = buildSymbolPanel(dialog, savedSettings);

        var buttonRefs = buildButtonRow(dialog);

        return {
            dialog: dialog,
            rightRadio: placementRefs.rightRadio,
            belowRadio: placementRefs.belowRadio,
            artboardGapInput: placementRefs.artboardGapInput,
            marginInput: artboardRefs.marginInput,
            updateCheckbox: artboardRefs.updateCheckbox,
            filterAllRadio: symbolRefs.filterAllRadio,
            filterUsedRadio: symbolRefs.filterUsedRadio,
            showCaptionCheckbox: symbolRefs.showCaptionCheckbox,
            symbolGapInput: symbolRefs.symbolGapInput,
            maxRowWidthInput: symbolRefs.maxRowWidthInput,
            okButton: buttonRefs.okButton,
            cancelButton: buttonRefs.cancelButton
        };
    }

    /* 入力テキストを pt に変換 / Read input text as points */
    function readInputPt(inputText, defaultPt) {
        var parsedValue = parseFloat(inputText);
        return isNaN(parsedValue) ? defaultPt : unitToPt(parsedValue);
    }

    function readDialogSettings(controls) {
        return {
            position: controls.rightRadio.value ? "right" : "below",
            artboardGap: readInputPt(controls.artboardGapInput.text, DEFAULT_ARTBOARD_GAP_PT),
            margin: readInputPt(controls.marginInput.text, DEFAULT_ARTBOARD_MARGIN_PT),
            update: controls.updateCheckbox.value,
            showCaption: controls.showCaptionCheckbox.value,
            filter: controls.filterUsedRadio.value ? "used" : "all",
            symbolGap: readInputPt(controls.symbolGapInput.text, DEFAULT_SYMBOL_GAP_PT),
            maxRowWidth: readInputPt(controls.maxRowWidthInput.text, DEFAULT_MAX_ROW_WIDTH_PT)
        };
    }

    /* プレビュー再構築用クロージャ（常時 ON）。前回プレビューを削除してから再構築。
     * Build a refresher (always-on preview) that removes the previous preview before rebuilding. */
    function makePreviewRefresher(doc, controls) {
        var state = { current: null };
        function refresh() {
            if (state.current) {
                clearLayout(doc, state.current);
                state.current = null;
            }
            var settings = readDialogSettings(controls);
            /* 使用中のみ指定で対象 0 個の場合はプレビューせず終了 / Skip preview if no target */
            if (settings.filter === "used" && getTargetSymbols(doc, settings).length === 0) {
                app.redraw();
                return;
            }
            state.current = buildLayout(doc, settings);
            app.redraw();
        }
        return { refresh: refresh, state: state };
    }

    /* 入力イベントを refresh に結線 / Wire input events to refresh */
    function wireDialogEvents(controls, refresh) {
        var triggers = [
            controls.rightRadio, controls.belowRadio,
            controls.updateCheckbox,
            controls.filterAllRadio, controls.filterUsedRadio,
            controls.showCaptionCheckbox
        ];
        for (var i = 0; i < triggers.length; i++) triggers[i].onClick = refresh;

        var inputs = [
            controls.artboardGapInput,
            controls.marginInput,
            controls.symbolGapInput,
            controls.maxRowWidthInput
        ];
        for (var j = 0; j < inputs.length; j++) {
            inputs[j].onChange = refresh;
            changeValueByArrowKey(inputs[j], false, refresh);
        }
    }

    /* OK/キャンセルを結線して show / Wire buttons and show */
    function runDialog(controls) {
        var result = "cancel";
        controls.okButton.onClick = function () { result = "ok"; controls.dialog.close(1); };
        controls.cancelButton.onClick = function () { result = "cancel"; controls.dialog.close(2); };
        controls.dialog.show();
        return result;
    }

    /* ダイアログ全体のオーケストレーション / Orchestrate the dialog session */
    function showDialog(doc) {
        var savedSettings = loadSettings();
        var controls = buildDialog(savedSettings);
        var refresher = makePreviewRefresher(doc, controls);
        wireDialogEvents(controls, refresher.refresh);

        /* 起動直後に初回プレビューを表示 / Initial preview before showing dialog */
        refresher.refresh();

        var result = runDialog(controls);

        if (result === "ok") {
            var finalSettings = readDialogSettings(controls);
            if (refresher.state.current) {
                clearLayout(doc, refresher.state.current);
                refresher.state.current = null;
            }
            var finalState = buildLayout(doc, finalSettings);
            if (!finalState) {
                alert(L(LABELS.noUsedSymbols));
                app.redraw();
                return null;
            }
            if (finalSettings.update) {
                removeOtherSymbolLists(doc, finalState);
            }
            saveSettings(finalSettings);
            app.redraw();
            return finalState;
        }

        if (refresher.state.current) {
            clearLayout(doc, refresher.state.current);
            app.redraw();
        }
        return null;
    }

    // =========================================
    // メイン
    // =========================================

    function main() {
        if (app.documents.length === 0) {
            alert(L(LABELS.noDoc));
            return;
        }
        var doc = app.activeDocument;
        if (doc.symbols.length === 0) {
            alert(L(LABELS.noSymbols));
            return;
        }
        showDialog(doc);
    }

    main();

})();