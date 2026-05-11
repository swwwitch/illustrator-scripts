#target illustrator
/*
 * SymbolListBuilder.jsx
 * ----------------------------------------------------------------------------
 * 概要 / Overview
 *   Illustrator ドキュメントに登録されたシンボルを一覧表示する専用アートボード
 *   「シンボル一覧」を自動生成するスクリプト。ダイアログでパラメータを操作しな
 *   がらライブプレビューでき、OK で確定（プレビュー削除 → 最終ビルド → 旧版掃除）。
 *
 *   Generates a dedicated "Symbol List" artboard that lays out every (or only
 *   used) symbol registered in the active Illustrator document. Live preview
 *   updates while the dialog is open; clicking OK rebuilds the final layout
 *   and sweeps any previous Symbol List artboards/backgrounds.
 *
 * 主な機能 / Key features
 *   - 配置位置：最終アートボードの右／下を選択
 *   - 大きさとマージン：幅・高さ・内側マージン（幅変更時は最大幅も自動追従）
 *   - 背景色：なし／黒／白／グレー（K50）。背景黒のときキャプションを白に
 *   - シンボルの絞り込み：すべて／使用中のみ
 *   - シンボル名キャプション：しない／上／下、フォントサイズ（pt 単位）
 *   - 既定キャプションフォントはロケール別（ja → HiraginoSans-W3 / en → MyriadPro-Regular）
 *   - 「更新」ON で既存「シンボル一覧」アートボード／背景塗りを置換
 *
 * 単位系 / Units
 *   ルーラー単位 (rulerType) … 寸法・マージン・間隔
 *   テキスト単位 (text/units) … フォントサイズ
 *
 * 互換性 / Compatibility
 *   Illustrator ExtendScript（ES3 相当）。設定は app.preferences の string
 *   プリファレンスに JSON 風文字列で保存／復元（eval 不使用）。
 *
 * 更新履歴 / Changelog
 *   v1.1.0  パネル再構成（2 カラム）、背景色、幅・高さ override、
 *           キャプション 3 ラジオ（しない／上／下）、ロケール別既定フォント、
 *           黒背景時のキャプション白文字、更新時の旧背景塗り掃除。
 *   v1.0    初版 / Initial release.
 */
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

(function () {

    // =========================================
    // バージョンとローカライズ
    // =========================================

    /* スクリプトバージョン / Script version */
    var SCRIPT_VERSION = "v1.1.1";

    /* ロケール判定 / Locale detection */
    function getCurrentLang() {
        return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    /* ラベル定義 / Label definitions */
    var LABELS = {
        dialogTitle: { ja: "シンボル一覧を配置", en: "List All Symbols" },
        placePanel: { ja: "作成する位置", en: "Create at" },
        placeRight: { ja: "最終アートボードの右側", en: "Right of last artboard" },
        placeBelow: { ja: "最終アートボードの下側", en: "Below last artboard" },
        artboardPanel: { ja: "アートボード", en: "Artboard" },
        marginPanel: { ja: "大きさとマージン", en: "Size & margin" },
        bgColor: { ja: "背景色", en: "Background" },
        bgNone: { ja: "なし", en: "None" },
        bgBlack: { ja: "黒", en: "Black" },
        bgWhite: { ja: "白", en: "White" },
        bgGray: { ja: "グレー", en: "Gray" },
        tipBgNone: { ja: "背景の塗りを作成しない", en: "Do not create a background fill" },
        tipBgBlack: { ja: "アートボード背面に黒（K100）の塗りを敷く", en: "Place a solid black (K100) fill behind the artboard" },
        tipBgWhite: { ja: "アートボード背面に白の塗りを敷く", en: "Place a solid white fill behind the artboard" },
        tipBgGray: { ja: "アートボード背面にグレー（K50）の塗りを敷く", en: "Place a 50% gray (K50) fill behind the artboard" },
        symbolGroupPanel: { ja: "シンボル", en: "Symbol" },
        symbolPanel: { ja: "シンボルの配置", en: "Symbol placement" },
        captionPanel: { ja: "シンボル名", en: "Symbol name" },
        gap: { ja: "間隔", en: "Gap" },
        margin: { ja: "マージン", en: "Margin" },
        width: { ja: "幅", en: "Width" },
        height: { ja: "高さ", en: "Height" },
        tipWidth: { ja: "アートボードの幅。空欄または 0 で自動計算、入力すると指定値で固定", en: "Artboard width. Empty/0 = auto; a number forces that exact size" },
        tipHeight: { ja: "アートボードの高さ。空欄または 0 で自動計算、入力すると指定値で固定", en: "Artboard height. Empty/0 = auto; a number forces that exact size" },
        update: { ja: "更新", en: "Update" },
        tipUpdate: { ja: "既存の「シンボル一覧」アートボードがあれば置き換え", en: "Replace existing Symbol List artboard if present" },
        tipShowCaption: { ja: "各シンボルの下にシンボル名を表示", en: "Show symbol name below each symbol" },
        captionPosition: { ja: "表示位置", en: "Position" },
        captionNone: { ja: "しない", en: "Off" },
        captionAbove: { ja: "上", en: "Top" },
        captionBelow: { ja: "下", en: "Bottom" },
        tipCaptionNone: { ja: "シンボル名を表示しない", en: "Do not show the symbol name" },
        tipCaptionAbove: { ja: "シンボルの上にシンボル名を表示", en: "Place the name above the symbol" },
        tipCaptionBelow: { ja: "シンボルの下にシンボル名を表示", en: "Place the name below the symbol" },
        fontSize: { ja: "フォントサイズ", en: "Font size" },
        tipFontSize: { ja: "シンボル名のフォントサイズ（pt）", en: "Symbol name font size (pt)" },
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
    var DEFAULT_CAPTION_FONT_SIZE_PT = 9;
    /* ロケール別の既定フォント。UI からは指定しないが createCaption 内で適用する
     * Locale-based default caption font (no UI; applied silently in createCaption) */
    var DEFAULT_CAPTION_FONT_NAME = (currentLanguage === "ja") ? "HiraginoSans-W3" : "MyriadPro-Regular";
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

    /* text/units（文字単位）から表示単位とポイント変換係数を取得
     * Resolve type-unit display label & point conversion from `text/units` preference */
    function getTypeUnitInfo() {
        var u;
        try { u = app.preferences.getIntegerPreference("text/units"); }
        catch (e) { u = -1; }
        switch (u) {
            case 0: return { label: "inch", factor: 72.0 };
            case 1: return { label: "mm", factor: 72.0 / 25.4 };
            case 3: return { label: "pica", factor: 12.0 };
            case 4: return { label: "cm", factor: 72.0 / 2.54 };
            case 5: return { label: "Q", factor: 72.0 / 25.4 * 0.25 };
            case 6: return { label: "px", factor: 1.0 };
            default: return { label: "pt", factor: 1.0 };
        }
    }
    var TYPE_UNIT = getTypeUnitInfo();

    function ptToTypeUnit(pt) { return pt / TYPE_UNIT.factor; }
    function typeUnitToPt(value) { return value * TYPE_UNIT.factor; }

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
            '"maxRowWidth":' + settings.maxRowWidth + ',',
            '"bgColor":"' + settings.bgColor + '",',
            '"captionPosition":"' + settings.captionPosition + '",',
            '"fontSize":' + settings.fontSize,
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
        var bgColor = readString("bgColor");
        var captionPosition = readString("captionPosition");
        var artboardGap = readNumber("artboardGap");
        var margin = readNumber("margin");
        var symbolGap = readNumber("symbolGap");
        var maxRowWidth = readNumber("maxRowWidth");
        var fontSize = readNumber("fontSize");
        var update = readBoolean("update");
        var showCaption = readBoolean("showCaption");

        if ((position !== "right" && position !== "below") ||
            (filter !== "all" && filter !== "used") ||
            artboardGap === null || margin === null ||
            symbolGap === null || maxRowWidth === null ||
            update === null || showCaption === null) {
            return null;
        }

        if (bgColor !== "none" && bgColor !== "black" && bgColor !== "white" && bgColor !== "gray") {
            bgColor = "none";
        }

        if (captionPosition !== "above" && captionPosition !== "below") {
            captionPosition = "above";
        }

        if (fontSize === null || fontSize <= 0) {
            fontSize = DEFAULT_CAPTION_FONT_SIZE_PT;
        }

        return {
            position: position,
            artboardGap: artboardGap,
            margin: margin,
            update: update,
            showCaption: showCaption,
            filter: filter,
            symbolGap: symbolGap,
            maxRowWidth: maxRowWidth,
            bgColor: bgColor,
            captionPosition: captionPosition,
            fontSize: fontSize
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
    var PANEL_SPACING = 8;
    var VALUE_ROW_LABEL_WIDTH = 70;

    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = "left";
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
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

    /* 指定アートボード矩形と (ほぼ) 一致する塗り長方形を削除（過去版が別レイヤーへ移されていても掴めるよう全レイヤー対象）
     * keepLayer を指定すると、そのレイヤー上のパスは保護（同位置・同寸の新アートボードに作った
     * 新しい背景塗りを巻き添えで消さないため）。
     * Remove fill rectangles whose bounds match the artboard rect; preserves any path on keepLayer
     * so the freshly-created background fill is not collateral-deleted when old/new artboards
     * happen to share bounds. */
    function removeBackgroundFillsForArtboard(doc, artboardRect, keepLayer) {
        var tolerance = 1.0;
        var allPaths = doc.pathItems;
        var toRemove = [];
        for (var i = 0; i < allPaths.length; i++) {
            var p = allPaths[i];
            try {
                if (keepLayer && p.layer === keepLayer) continue;
                var b = p.geometricBounds;
                if (Math.abs(b[0] - artboardRect[0]) < tolerance &&
                    Math.abs(b[1] - artboardRect[1]) < tolerance &&
                    Math.abs(b[2] - artboardRect[2]) < tolerance &&
                    Math.abs(b[3] - artboardRect[3]) < tolerance &&
                    p.filled && !p.stroked) {
                    toRemove.push(p);
                }
            } catch (boundsError) { }
        }
        for (var j = 0; j < toRemove.length; j++) {
            try { toRemove[j].remove(); } catch (removeError) { }
        }
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
                /* 旧アートボード矩形と一致する塗り（背景）を先に掃除。新レイヤー上の塗りは保護
                 * Sweep old bg fills first; never touch the just-created (kept) layer */
                removeBackgroundFillsForArtboard(doc, doc.artboards[i].artboardRect, state.layer);
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
    function createCaption(layer, symbolName, fontSizePt, fillColor) {
        var caption = layer.textFrames.add();
        caption.contents = symbolName;
        try {
            caption.textRange.characterAttributes.size = fontSizePt;
        } catch (sizeError) { }
        try {
            caption.textRange.characterAttributes.textFont = app.textFonts.getByName(DEFAULT_CAPTION_FONT_NAME);
        } catch (fontError) { }
        try {
            caption.textRange.paragraphAttributes.justification = Justification.CENTER;
        } catch (justifyError) { }
        if (fillColor) {
            try {
                caption.textRange.characterAttributes.fillColor = fillColor;
            } catch (colorError) { }
        }
        return caption;
    }

    /* シンボルをインスタンス化し、必要ならキャプションも作成、エントリ配列を返す
     * Instantiate symbols (and optional captions), measure, return entries */
    function instantiateSymbols(doc, layer, settings, targetSymbols) {
        /* 背景が黒のときだけキャプションを白に。それ以外は既定（黒）/ White caption only when bg is black */
        var captionFillColor = (settings.bgColor === "black") ? createBgFillColor(doc, "white") : null;
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
                var caption = createCaption(layer, symbol.name, settings.fontSize, captionFillColor);
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
    function placeEntries(entries, originLeft, originTop, margin, captionPosition) {
        var captionAbove = (captionPosition === "above");
        for (var i = 0; i < entries.length; i++) {
            var entry = entries[i];
            var slotLeft = originLeft + margin + entry.x;
            var slotTop = originTop - margin + entry.y;

            /* スロット内でシンボルを水平中央寄せ / Center symbol within slot */
            var symbolLeft = slotLeft + (entry.width - entry.symbolWidth) / 2;
            var symbolTop = (entry.caption && captionAbove)
                ? slotTop - entry.captionHeight - CAPTION_GAP_PT
                : slotTop;
            entry.item.position = [symbolLeft, symbolTop];

            if (entry.caption) {
                var captionLeft = slotLeft + (entry.width - entry.captionWidth) / 2;
                var captionTop = captionAbove
                    ? slotTop
                    : slotTop - entry.symbolHeight - CAPTION_GAP_PT;
                entry.caption.position = [captionLeft, captionTop];
            }
        }
    }

    /* 背景色選択からドキュメントカラースペースに合った塗り色を生成
     * Resolve fill color matching the document color space */
    function createBgFillColor(doc, choice) {
        if (!choice || choice === "none") return null;
        var isCMYK = (doc.documentColorSpace === DocumentColorSpace.CMYK);
        var k;
        if (choice === "black") k = 100;
        else if (choice === "white") k = 0;
        else if (choice === "gray") k = 50;
        else return null;

        if (isCMYK) {
            var cmyk = new CMYKColor();
            cmyk.cyan = 0; cmyk.magenta = 0; cmyk.yellow = 0; cmyk.black = k;
            return cmyk;
        }
        var v = Math.round(255 * (1 - k / 100));
        var rgb = new RGBColor();
        rgb.red = v; rgb.green = v; rgb.blue = v;
        return rgb;
    }

    /* アートボード矩形いっぱいに背景塗りを敷く（最背面に送る）
     * Place a background fill across the artboard rect and send it to back */
    function createBackgroundFill(layer, rect, color) {
        if (!color) return null;
        var left = rect[0];
        var top = rect[1];
        var width = rect[2] - rect[0];
        var height = rect[1] - rect[3];
        var bg = layer.pathItems.rectangle(top, left, width, height);
        bg.filled = true;
        bg.stroked = false;
        bg.fillColor = color;
        try { bg.zOrder(ZOrderMethod.SENDTOBACK); } catch (e) { }
        return bg;
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
        var artboardWidth = (settings.widthOverridePt && settings.widthOverridePt > 0)
            ? settings.widthOverridePt
            : size.width + settings.margin * 2;
        var artboardHeight = (settings.heightOverridePt && settings.heightOverridePt > 0)
            ? settings.heightOverridePt
            : size.height + settings.margin * 2;
        var artboardInfo = addArtboard(
            doc, origin.left, origin.top,
            artboardWidth, artboardHeight,
            ARTBOARD_NAME
        );
        placeEntries(entries, origin.left, origin.top, settings.margin, settings.captionPosition);

        createBackgroundFill(layer, artboardInfo.rect, createBgFillColor(doc, settings.bgColor));

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
        setupPanel(panel);
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

    function buildMarginPanel(parent, savedSettings) {
        var panel = parent.add("panel", undefined, L(LABELS.marginPanel));
        setupPanel(panel);
        var widthInput = addValueRow(panel, labelText(LABELS.width), 0, L(LABELS.tipWidth));
        var heightInput = addValueRow(panel, labelText(LABELS.height), 0, L(LABELS.tipHeight));
        var marginInput = addValueRow(panel, labelText(LABELS.margin),
            getSavedSetting(savedSettings, "margin", DEFAULT_ARTBOARD_MARGIN_PT), L(LABELS.tipMargin));
        return {
            widthInput: widthInput,
            heightInput: heightInput,
            marginInput: marginInput
        };
    }

    function buildBgColorPanel(parent, savedSettings) {
        var panel = parent.add("panel", undefined, L(LABELS.bgColor));
        setupPanel(panel);

        /* 2 行 2 列。各ラジオの幅を揃えて列をきれいに整える / 2x2 grid; equal widths align columns */
        var BG_RADIO_WIDTH = 60;

        var row1 = panel.add("group");
        row1.orientation = "row";
        row1.spacing = 10;
        var bgNoneRadio = row1.add("radiobutton", undefined, L(LABELS.bgNone));
        bgNoneRadio.helpTip = L(LABELS.tipBgNone);
        bgNoneRadio.preferredSize.width = BG_RADIO_WIDTH;
        var bgBlackRadio = row1.add("radiobutton", undefined, L(LABELS.bgBlack));
        bgBlackRadio.helpTip = L(LABELS.tipBgBlack);
        bgBlackRadio.preferredSize.width = BG_RADIO_WIDTH;

        var row2 = panel.add("group");
        row2.orientation = "row";
        row2.spacing = 10;
        var bgWhiteRadio = row2.add("radiobutton", undefined, L(LABELS.bgWhite));
        bgWhiteRadio.helpTip = L(LABELS.tipBgWhite);
        bgWhiteRadio.preferredSize.width = BG_RADIO_WIDTH;
        var bgGrayRadio = row2.add("radiobutton", undefined, L(LABELS.bgGray));
        bgGrayRadio.helpTip = L(LABELS.tipBgGray);
        bgGrayRadio.preferredSize.width = BG_RADIO_WIDTH;

        var savedBgColor = getSavedSetting(savedSettings, "bgColor", "none");
        bgNoneRadio.value = (savedBgColor === "none");
        bgBlackRadio.value = (savedBgColor === "black");
        bgWhiteRadio.value = (savedBgColor === "white");
        bgGrayRadio.value = (savedBgColor === "gray");
        return {
            bgNoneRadio: bgNoneRadio,
            bgBlackRadio: bgBlackRadio,
            bgWhiteRadio: bgWhiteRadio,
            bgGrayRadio: bgGrayRadio
        };
    }

    function buildCaptionPanel(parent, savedSettings) {
        var panel = parent.add("panel", undefined, L(LABELS.captionPanel));
        setupPanel(panel);

        /* 表示位置（横並び 3 ラジオ）。「しない」が「シンボル名を表示」OFF を兼ねる
         * Caption position as a 3-way horizontal radio; "Off" replaces the show-caption checkbox */
        var posRow = panel.add("group");
        posRow.orientation = "row";
        posRow.alignChildren = ["left", "center"];
        posRow.spacing = 6;
        posRow.add("statictext", undefined, labelText(LABELS.captionPosition));
        var posRadios = posRow.add("group");
        posRadios.orientation = "row";
        posRadios.alignChildren = "left";
        posRadios.spacing = 10;
        var captionNoneRadio = posRadios.add("radiobutton", undefined, L(LABELS.captionNone));
        captionNoneRadio.helpTip = L(LABELS.tipCaptionNone);
        var captionAboveRadio = posRadios.add("radiobutton", undefined, L(LABELS.captionAbove));
        captionAboveRadio.helpTip = L(LABELS.tipCaptionAbove);
        var captionBelowRadio = posRadios.add("radiobutton", undefined, L(LABELS.captionBelow));
        captionBelowRadio.helpTip = L(LABELS.tipCaptionBelow);
        var savedShowCaption = getSavedSetting(savedSettings, "showCaption", false);
        var savedCaptionPosition = getSavedSetting(savedSettings, "captionPosition", "above");
        captionNoneRadio.value = (savedShowCaption === false);
        captionAboveRadio.value = (savedShowCaption !== false && savedCaptionPosition === "above");
        captionBelowRadio.value = (savedShowCaption !== false && savedCaptionPosition === "below");

        /* フォントサイズ（単位は環境設定 text/units に従う）/ Font size in the user's text-unit pref */
        var fontSizeRow = panel.add("group");
        fontSizeRow.orientation = "row";
        fontSizeRow.alignChildren = "center";
        fontSizeRow.spacing = 6;
        var fsLabelColumn = fontSizeRow.add("group");
        fsLabelColumn.orientation = "row";
        fsLabelColumn.alignChildren = ["right", "center"];
        fsLabelColumn.margins = 0;
        fsLabelColumn.spacing = 0;
        fsLabelColumn.preferredSize.width = VALUE_ROW_LABEL_WIDTH;
        var fsLabel = fsLabelColumn.add("statictext", undefined, labelText(LABELS.fontSize));
        fsLabel.helpTip = L(LABELS.tipFontSize);
        var savedFontSize = getSavedSetting(savedSettings, "fontSize", DEFAULT_CAPTION_FONT_SIZE_PT);
        var fontSizeInput = fontSizeRow.add("edittext", undefined, String(Math.round(ptToTypeUnit(savedFontSize))));
        fontSizeInput.characters = 4;
        fontSizeInput.helpTip = L(LABELS.tipFontSize);
        fontSizeRow.add("statictext", undefined, TYPE_UNIT.label);

        return {
            captionNoneRadio: captionNoneRadio,
            captionAboveRadio: captionAboveRadio,
            captionBelowRadio: captionBelowRadio,
            fontSizeRow: fontSizeRow,
            fontSizeInput: fontSizeInput
        };
    }

    function buildSymbolPanel(parent, savedSettings) {
        var panel = parent.add("panel", undefined, L(LABELS.symbolPanel));
        setupPanel(panel);

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

        var symbolGapInput = addValueRow(panel, labelText(LABELS.gap),
            getSavedSetting(savedSettings, "symbolGap", DEFAULT_SYMBOL_GAP_PT), L(LABELS.tipSymbolGap));
        var maxRowWidthInput = addValueRow(panel, labelText(LABELS.maxRowWidth),
            getSavedSetting(savedSettings, "maxRowWidth", DEFAULT_MAX_ROW_WIDTH_PT), L(LABELS.tipMaxRowWidth));

        return {
            filterAllRadio: filterAllRadio,
            filterUsedRadio: filterUsedRadio,
            symbolGapInput: symbolGapInput,
            maxRowWidthInput: maxRowWidthInput
        };
    }

    function buildButtonRow(parent, savedSettings) {
        var row = parent.add("group");
        row.alignment = "fill";
        row.orientation = "row";
        row.spacing = 8;

        /* 左：更新チェックボックス / Left: update checkbox */
        var leftCol = row.add("group");
        leftCol.alignment = ["left", "center"];
        var updateCheckbox = leftCol.add("checkbox", undefined, L(LABELS.update));
        updateCheckbox.value = true;
        updateCheckbox.helpTip = L(LABELS.tipUpdate);

        /* 中央：スペーサー / Center: spacer */
        var spacer = row.add("group");
        spacer.alignment = ["fill", "center"];

        /* 右：キャンセル / OK / Right: Cancel / OK */
        var rightCol = row.add("group");
        rightCol.alignment = ["right", "center"];
        rightCol.spacing = 8;
        var cancelButton = rightCol.add("button", undefined, L(LABELS.cancel), { name: "cancel" });
        var okButton = rightCol.add("button", undefined, "OK", { name: "OK" });

        return {
            updateCheckbox: updateCheckbox,
            okButton: okButton,
            cancelButton: cancelButton
        };
    }

    /* ダイアログを構築し、コントロール参照を返す / Build dialog and return refs */
    function buildDialog(savedSettings) {
        var dialog = new Window("dialog", L(LABELS.dialogTitle) + " " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = "fill";
        dialog.margins = 16;
        dialog.spacing = 12;

        /* 2 カラム構成 / Two-column layout */
        var columns = dialog.add("group");
        columns.orientation = "row";
        columns.alignChildren = ["fill", "top"];
        columns.spacing = 12;

        var leftColumn = columns.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = "fill";
        leftColumn.spacing = 12;

        var rightColumn = columns.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = "fill";
        rightColumn.spacing = 12;

        /* アートボード関連パネルを内包するラッパー / Wrapper for artboard-related panels */
        var artboardGroupPanel = leftColumn.add("panel", undefined, L(LABELS.artboardPanel));
        setupPanel(artboardGroupPanel);
        artboardGroupPanel.margins = [15, 20, 15, 15];

        var placementRefs = buildPlacementPanel(artboardGroupPanel, savedSettings);
        var marginRefs = buildMarginPanel(artboardGroupPanel, savedSettings);
        var bgColorRefs = buildBgColorPanel(artboardGroupPanel, savedSettings);

        /* シンボル関連パネルを内包するラッパー / Wrapper for symbol-related panels */
        var symbolGroupPanel = rightColumn.add("panel", undefined, L(LABELS.symbolGroupPanel));
        setupPanel(symbolGroupPanel);
        symbolGroupPanel.margins = [15, 20, 15, 15];
        var symbolRefs = buildSymbolPanel(symbolGroupPanel, savedSettings);
        var captionRefs = buildCaptionPanel(symbolGroupPanel, savedSettings);

        var buttonRefs = buildButtonRow(dialog, savedSettings);

        return {
            dialog: dialog,
            rightRadio: placementRefs.rightRadio,
            belowRadio: placementRefs.belowRadio,
            artboardGapInput: placementRefs.artboardGapInput,
            widthInput: marginRefs.widthInput,
            heightInput: marginRefs.heightInput,
            marginInput: marginRefs.marginInput,
            bgNoneRadio: bgColorRefs.bgNoneRadio,
            bgBlackRadio: bgColorRefs.bgBlackRadio,
            bgWhiteRadio: bgColorRefs.bgWhiteRadio,
            bgGrayRadio: bgColorRefs.bgGrayRadio,
            updateCheckbox: buttonRefs.updateCheckbox,
            fontSizeInput: captionRefs.fontSizeInput,
            filterAllRadio: symbolRefs.filterAllRadio,
            filterUsedRadio: symbolRefs.filterUsedRadio,
            captionNoneRadio: captionRefs.captionNoneRadio,
            captionAboveRadio: captionRefs.captionAboveRadio,
            captionBelowRadio: captionRefs.captionBelowRadio,
            fontSizeRow: captionRefs.fontSizeRow,
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

    function readBgColorChoice(controls) {
        if (controls.bgBlackRadio.value) return "black";
        if (controls.bgWhiteRadio.value) return "white";
        if (controls.bgGrayRadio.value) return "gray";
        return "none";
    }

    /* フォントサイズ入力。表示は TYPE_UNIT、内部は常に pt
     * Font size input is shown in the user's text-unit preference; stored in pt internally */
    function readFontSizePt(controls) {
        var v = parseFloat(controls.fontSizeInput.text);
        if (isNaN(v) || v <= 0) return DEFAULT_CAPTION_FONT_SIZE_PT;
        return typeUnitToPt(v);
    }

    function readDialogSettings(controls) {
        return {
            position: controls.rightRadio.value ? "right" : "below",
            artboardGap: readInputPt(controls.artboardGapInput.text, DEFAULT_ARTBOARD_GAP_PT),
            margin: readInputPt(controls.marginInput.text, DEFAULT_ARTBOARD_MARGIN_PT),
            update: controls.updateCheckbox.value,
            showCaption: !controls.captionNoneRadio.value,
            filter: controls.filterUsedRadio.value ? "used" : "all",
            symbolGap: readInputPt(controls.symbolGapInput.text, DEFAULT_SYMBOL_GAP_PT),
            maxRowWidth: readInputPt(controls.maxRowWidthInput.text, DEFAULT_MAX_ROW_WIDTH_PT),
            bgColor: readBgColorChoice(controls),
            captionPosition: controls.captionBelowRadio.value ? "below" : "above",
            fontSize: readFontSizePt(controls)
        };
    }

    /* プレビュー再構築用クロージャ（常時 ON）。前回プレビューを削除してから再構築。
     * - widthOverridePt / heightOverridePt が非 null の場合、その固定サイズでアートボード生成
     * - ビルド後、有効サイズを width/height 入力欄に書き戻す（プログラム的設定なので onChange は発火しない）
     * Build a refresher (always-on preview); supports user-typed width/height overrides */
    function makePreviewRefresher(doc, controls) {
        var state = {
            current: null,
            widthOverridePt: null,
            heightOverridePt: null
        };
        function clearOverrides() {
            state.widthOverridePt = null;
            state.heightOverridePt = null;
        }
        function refresh() {
            if (state.current) {
                clearLayout(doc, state.current);
                state.current = null;
            }
            var settings = readDialogSettings(controls);
            settings.widthOverridePt = state.widthOverridePt;
            settings.heightOverridePt = state.heightOverridePt;

            if (settings.filter === "used" && getTargetSymbols(doc, settings).length === 0) {
                app.redraw();
                return;
            }
            state.current = buildLayout(doc, settings);
            if (state.current) {
                var rect = state.current.artboardInfo.rect;
                var widthPt = rect[2] - rect[0];
                var heightPt = rect[1] - rect[3];
                if (state.widthOverridePt === null) {
                    controls.widthInput.text = String(Math.round(ptToUnit(widthPt)));
                }
                if (state.heightOverridePt === null) {
                    controls.heightInput.text = String(Math.round(ptToUnit(heightPt)));
                }
            }
            app.redraw();
        }
        return { refresh: refresh, state: state, clearOverrides: clearOverrides };
    }

    /* 入力イベントを refresh に結線 / Wire input events to refresh */
    function wireDialogEvents(controls, refresher) {
        function refresh() { refresher.refresh(); }
        function clearAndRefresh() {
            refresher.clearOverrides();
            refresher.refresh();
        }

        /* 背景色は親グループが分かれているため ScriptUI 標準の排他選択が効かない。手動で他を解除。
         * Background radios live in separate parent groups, so we enforce exclusivity manually. */
        var bgRadios = [
            controls.bgNoneRadio, controls.bgBlackRadio, controls.bgWhiteRadio, controls.bgGrayRadio
        ];
        function applyBgExclusive(target) {
            for (var b = 0; b < bgRadios.length; b++) {
                if (bgRadios[b] !== target) bgRadios[b].value = false;
            }
        }
        for (var k = 0; k < bgRadios.length; k++) {
            (function (radio) {
                radio.onClick = function () {
                    applyBgExclusive(radio);
                    clearAndRefresh();
                };
            })(bgRadios[k]);
        }

        /* キャプション「しない」のときフォントサイズ行をディム
         * Dim font-size row when caption is set to "Off" */
        function updateCaptionEnabled() {
            controls.fontSizeRow.enabled = !controls.captionNoneRadio.value;
        }

        /* 寸法以外の操作は幅・高さオーバーライドを解除して再描画
         * Non-size triggers clear the width/height override before refreshing */
        var triggers = [
            controls.rightRadio, controls.belowRadio,
            controls.updateCheckbox,
            controls.filterAllRadio, controls.filterUsedRadio
        ];
        for (var i = 0; i < triggers.length; i++) triggers[i].onClick = clearAndRefresh;

        var captionRadios = [
            controls.captionNoneRadio, controls.captionAboveRadio, controls.captionBelowRadio
        ];
        for (var c = 0; c < captionRadios.length; c++) {
            captionRadios[c].onClick = function () {
                updateCaptionEnabled();
                clearAndRefresh();
            };
        }
        updateCaptionEnabled();

        var inputs = [
            controls.artboardGapInput,
            controls.marginInput,
            controls.symbolGapInput,
            controls.maxRowWidthInput,
            controls.fontSizeInput
        ];
        for (var j = 0; j < inputs.length; j++) {
            inputs[j].onChange = clearAndRefresh;
            changeValueByArrowKey(inputs[j], false, clearAndRefresh);
        }

        /* 幅・高さ：入力されたら override として固定。0/空欄は自動 / Width-height override
         * 幅を変えたときは「最大幅」も内側マージンを引いた値に追従させる
         * Adjusting width also retunes maxRowWidth (= width − margin × 2) so packing fits naturally. */
        function applyWidthFromInput() {
            var pt = readInputPt(controls.widthInput.text, 0);
            refresher.state.widthOverridePt = (pt > 0) ? pt : null;
            if (pt > 0) {
                var marginPt = readInputPt(controls.marginInput.text, DEFAULT_ARTBOARD_MARGIN_PT);
                var newMaxPt = pt - 2 * marginPt;
                if (newMaxPt > 0) {
                    controls.maxRowWidthInput.text = String(Math.round(ptToUnit(newMaxPt)));
                }
            }
            refresh();
        }
        function applyHeightFromInput() {
            var pt = readInputPt(controls.heightInput.text, 0);
            refresher.state.heightOverridePt = (pt > 0) ? pt : null;
            refresh();
        }
        controls.widthInput.onChange = applyWidthFromInput;
        controls.heightInput.onChange = applyHeightFromInput;
        changeValueByArrowKey(controls.widthInput, false, applyWidthFromInput);
        changeValueByArrowKey(controls.heightInput, false, applyHeightFromInput);
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
        wireDialogEvents(controls, refresher);

        /* 起動直後に初回プレビューを表示 / Initial preview before showing dialog */
        refresher.refresh();

        var result = runDialog(controls);

        if (result === "ok") {
            var finalSettings = readDialogSettings(controls);
            finalSettings.widthOverridePt = refresher.state.widthOverridePt;
            finalSettings.heightOverridePt = refresher.state.heightOverridePt;
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