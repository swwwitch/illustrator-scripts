#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

## SymbolListBuilder.jsx

### 概要 / Overview

Illustrator ドキュメントに登録されたシンボルを一覧表示する専用アートボード「シンボル一覧」を自動生成するスクリプト。ダイアログでパラメータを操作しながらライブプレビューでき、OK で確定（プレビュー削除 → 最終ビルド → 旧版掃除）。

### 主な機能 / Key features

- 作成位置の基準：最終アートボード／番号指定（既定はキャンバス最右下のアートボードを自動採用）
- 作成方向：基準アートボードの右側／下側を選択
- サイズと余白：幅・高さ・内側余白（幅変更時は最大幅も自動追従）
- 背景色：なし／黒／白／グレー（K50）。背景黒のときキャプションを白に
- シンボルの絞り込み：すべて／使用中のみ
- キャプション：しない／上／下、フォントサイズ（Illustrator の文字設定単位）
- 既定キャプションフォントはロケール別（ja → HiraginoSans-W3 / en → MyriadPro-Regular）
- レイヤー／アートボード名はロケール別（ja「シンボル一覧」／ en「Symbol List」）。既存判定は両言語に対応
- 「更新」ON で既存「シンボル一覧」アートボードと、その上に乗っているオブジェクトをすべて削除して置換

### 単位系 / Units

- ルーラー単位 (rulerType) … 寸法・マージン・間隔
- テキスト単位 (text/units) … フォントサイズ

### 紹介記事（note）

https://note.com/dtp_tranist/n/ncac687d0a3a0

### 更新履歴 / Changelog

- v1.0.0（2026-05-09）：初版 / Initial release.
- v1.2.1（2026-06-03）：作成位置を「基準（最終アートボード／番号指定）＋方向（右側／下側）」の 2 グループに再編し、指定番号の既定値にキャンバス最右下のアートボードを自動採用。更新 ON 時は旧「シンボル一覧」アートボード上のオブジェクトをすべて削除。レイヤー／アートボード名をロケール別にし、既存判定は両言語対応。
- v1.2.2（2026-06-03）：パネル名・ラベル・ツールチップの文言を整理。画面ズームスライダーを廃止し、ビュー合わせボタン（シンボル一覧＝作成アートボードにフィット＋90%／全体表示＝全アートボードにフィット＋90%）を最下段ボタン行の左に配置。更新チェックを「作成するアートボード」パネル末尾へ移動。収集対象ラジオを横並びに。

*/

/* スクリプトバージョン / Script version */
var SCRIPT_VERSION = "v1.2.2";

(function () {

    // =========================================
    // ローカライズ
    // =========================================

    /* ロケール判定 / Locale detection */
    function getCurrentLang() {
        return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    /* ラベル定義 / Label definitions */
    var LABELS = {
        /* ダイアログ / Dialog */
        dialog: {
            title: { ja: "シンボル一覧を作成", en: "Create Symbol List" }
        },
        /* パネル見出し / Panel titles */
        panel: {
            place: { ja: "作成位置", en: "Location" },
            artboard: { ja: "作成するアートボード", en: "New artboard" },
            size: { ja: "サイズと余白", en: "Size & padding" },
            bgColor: { ja: "背景", en: "Background" },
            symbolGroup: { ja: "収集するシンボル", en: "Symbols" },
            placement: { ja: "並べ方", en: "Layout" },
            caption: { ja: "キャプション", en: "Caption" },
            target: { ja: "収集対象", en: "Collect" }
        },
        /* ラジオボタン / Radio buttons */
        radio: {
            baseLast: { ja: "最終アートボード", en: "Last artboard" },
            baseSpecified: { ja: "指定", en: "Specified" },
            directionRight: { ja: "右側", en: "Right" },
            directionBelow: { ja: "下側", en: "Below" },
            bgNone: { ja: "なし", en: "None" },
            bgBlack: { ja: "黒", en: "Black" },
            bgWhite: { ja: "白", en: "White" },
            bgGray: { ja: "グレー", en: "Gray" },
            captionAbove: { ja: "上", en: "Top" },
            captionBelow: { ja: "下", en: "Bottom" },
            filterAll: { ja: "すべて", en: "All" },
            filterUsed: { ja: "使用中のみ", en: "Used only" }
        },
        /* チェックボックス / Checkboxes */
        checkbox: {
            update: { ja: "更新", en: "Update" },
            showCaption: { ja: "シンボル名を表示", en: "Show symbol name" }
        },
        /* ボタン / Buttons */
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            fitCreated: { ja: "シンボル一覧", en: "Symbol list" },
            fitAll: { ja: "全体表示", en: "Fit all" }
        },
        /* 行ラベル（statictext）/ Inline labels (statictext) */
        label: {
            direction: { ja: "方向", en: "Direction" },
            captionPosition: { ja: "位置", en: "Position" }
        },
        /* 数値入力欄ラベル / Numeric input labels */
        field: {
            gap: { ja: "間隔", en: "Gap" },
            margin: { ja: "余白", en: "Padding" },
            width: { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" },
            fontSize: { ja: "フォントサイズ", en: "Font size" },
            maxRowWidth: { ja: "最大幅", en: "Max row width" }
        },
        /* ツールチップ / Tooltips */
        tip: {
            baseLast: { ja: "アートボード一覧の末尾を基準にする。更新 ON の場合は既存のシンボル一覧を除外", en: "Use the last artboard in the list as the base; with Update on, the existing symbol list is excluded" },
            baseSpecified: { ja: "指定した番号のアートボードを基準にする（1 始まり）", en: "Use the artboard with the given number as the base (1-based)" },
            baseArtboardNumber: { ja: "基準にするアートボード番号（1 始まり）", en: "Base artboard number (1-based)" },
            directionRight: { ja: "基準アートボードの右に新規アートボードを作成", en: "Place the new artboard to the right of the base" },
            directionBelow: { ja: "基準アートボードの下に新規アートボードを作成", en: "Place the new artboard below the base" },
            artboardGap: { ja: "基準アートボードと新規アートボードの間隔", en: "Distance between the base and new artboards" },
            margin: { ja: "新規アートボードの内側に設ける余白", en: "Padding inside the new artboard around the symbols" },
            symbolGap: { ja: "シンボル同士の間隔", en: "Spacing between adjacent symbols" },
            maxRowWidth: { ja: "1行に並べる最大幅。超えると次の行に折り返し", en: "Max width per row; symbols wrap to the next row when exceeded" },
            width: { ja: "アートボードの幅。空欄または 0 で自動計算、入力すると指定値で固定", en: "Artboard width. Empty/0 = auto; a number forces that exact size" },
            height: { ja: "アートボードの高さ。空欄または 0 で自動計算、入力すると指定値で固定", en: "Artboard height. Empty/0 = auto; a number forces that exact size" },
            update: { ja: "既存の「シンボル一覧」アートボードと対象オブジェクトを削除して作り直す", en: "Delete the existing Symbol List artboard and its objects, then rebuild" },
            showCaption: { ja: "各シンボルの近くにシンボル名をキャプションとして表示", en: "Show the symbol name as a caption near each symbol" },
            captionAbove: { ja: "シンボルの上にシンボル名を表示", en: "Place the name above the symbol" },
            captionBelow: { ja: "シンボルの下にシンボル名を表示", en: "Place the name below the symbol" },
            fontSize: { ja: "キャプションのフォントサイズ。単位は Illustrator の文字設定に従う", en: "Caption font size; unit follows Illustrator's type preferences" },
            filterAll: { ja: "ドキュメントに登録されているすべてのシンボルを並べる", en: "List every symbol registered in the document" },
            filterUsed: { ja: "ドキュメント内に配置されているシンボルだけを並べる", en: "List only symbols placed in the document" },
            bgNone: { ja: "背景の塗りを作成しない", en: "Do not create a background fill" },
            bgBlack: { ja: "アートボード背面に黒（K100）の塗りを敷く", en: "Place a solid black (K100) fill behind the artboard" },
            bgWhite: { ja: "アートボード背面に白の塗りを敷く", en: "Place a solid white fill behind the artboard" },
            bgGray: { ja: "アートボード背面にグレー（K50）の塗りを敷く", en: "Place a 50% gray (K50) fill behind the artboard" },
            fitCreated: { ja: "作成したアートボードのみをウィンドウに合わせて表示", en: "Fit the created artboard to the window" },
            fitAll: { ja: "すべてのアートボードをウィンドウに合わせて表示", en: "Fit all artboards to the window" }
        },
        /* メッセージ / Messages */
        message: {
            noDoc: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSymbols: { ja: "登録されているシンボルがありません。", en: "No symbols are registered." },
            noUsedSymbols: { ja: "ドキュメント内で使用中のシンボルがありません。", en: "No symbols are currently used in the document." }
        }
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

    /* レイヤー／アートボード名（ロケール別）。生成にはこの名前を使う。
     * Locale-based layer/artboard name; used when creating items. */
    var SYMBOL_LIST_LABEL = { ja: "シンボル一覧", en: "Symbol List" };
    var LAYER_NAME = L(SYMBOL_LIST_LABEL);
    var ARTBOARD_NAME = L(SYMBOL_LIST_LABEL);
    /* 既存判定用の全ロケール名称セット。別言語で作成済みでも更新できるよう全変種を対象にする。
     * All locale variants for matching existing items, so a list built in another language is still updatable. */
    var SYMBOL_LIST_NAMES = [SYMBOL_LIST_LABEL.ja, SYMBOL_LIST_LABEL.en];
    function isSymbolListName(name) {
        for (var i = 0; i < SYMBOL_LIST_NAMES.length; i++) {
            if (name === SYMBOL_LIST_NAMES[i]) return true;
        }
        return false;
    }

    /* 設定保存用キー / Preference key for saved settings */
    var PREF_KEY = "swwwitch.listupallsymbol.settings";

    // =========================================
    // 単位
    // =========================================

    /* 単位コード（0=inch / 1=mm / 3=pica / 4=cm / 5=Q / 6=px / その他=pt）から表示単位とポイント変換係数を取得
     * Map a unit code to its display label and point conversion factor */
    function getUnitInfoByCode(code) {
        switch (code) {
            case 0: return { label: "inch", factor: 72.0 };
            case 1: return { label: "mm", factor: 72.0 / 25.4 };
            case 3: return { label: "pica", factor: 12.0 };
            case 4: return { label: "cm", factor: 72.0 / 2.54 };
            case 5: return { label: "Q", factor: 72.0 / 25.4 * 0.25 };
            case 6: return { label: "px", factor: 1.0 };
            default: return { label: "pt", factor: 1.0 };
        }
    }

    /* 整数プリファレンス（rulerType / text/units）から単位情報を取得（取得失敗時は pt 相当）
     * Resolve unit info from an integer preference key (defaults to pt on failure) */
    function getUnitInfoFromPreference(prefKey) {
        var code;
        try { code = app.preferences.getIntegerPreference(prefKey); }
        catch (e) { code = -1; }
        return getUnitInfoByCode(code);
    }

    /* ルーラー単位（寸法・マージン・間隔）/ Ruler unit for sizes, margins, gaps */
    var UNIT = getUnitInfoFromPreference("rulerType");
    function ptToUnit(pt) { return pt / UNIT.factor; }
    function unitToPt(value) { return value * UNIT.factor; }

    /* テキスト単位（フォントサイズ）/ Type unit for font size */
    var TYPE_UNIT = getUnitInfoFromPreference("text/units");

    function ptToTypeUnit(pt) { return pt / TYPE_UNIT.factor; }
    function typeUnitToPt(value) { return value * TYPE_UNIT.factor; }

    // =========================================
    // 設定の保存・復元
    // =========================================

    /* 設定オブジェクトを JSON 風文字列へ（手動 JSON, ES3 互換）。
     * キー一覧を 1 か所にまとめ、文字列値は引用、数値・真偽はそのまま出力する。
     * Serialize settings to a JSON-like string: single key list; quote strings, emit numbers/booleans raw. */
    var SETTINGS_KEYS = [
        "position", "baseMode", "baseArtboardNumber", "artboardGap", "margin",
        "update", "showCaption", "filter", "symbolGap", "maxRowWidth",
        "bgColor", "captionPosition", "fontSize"
    ];
    function serializeSettings(settings) {
        var parts = [];
        for (var i = 0; i < SETTINGS_KEYS.length; i++) {
            var key = SETTINGS_KEYS[i];
            var value = settings[key];
            var encoded = (typeof value === "string") ? ('"' + value + '"') : String(value);
            parts.push('"' + key + '":' + encoded);
        }
        return "{" + parts.join(",") + "}";
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
        var baseMode = readString("baseMode");
        var baseArtboardNumber = readNumber("baseArtboardNumber");
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

        /* 基準モード（保存値が無ければ最終アートボード基準）/ Base mode (default: last artboard) */
        if (baseMode !== "last" && baseMode !== "specified") {
            baseMode = "last";
        }
        if (baseArtboardNumber === null || baseArtboardNumber < 1) {
            baseArtboardNumber = 1;
        }

        return {
            position: position,
            baseMode: baseMode,
            baseArtboardNumber: baseArtboardNumber,
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

    /* 設定を文字列化して app.preferences に保存 / Serialize settings and store in app.preferences */
    function saveSettings(settings) {
        try { app.preferences.setStringPreference(PREF_KEY, serializeSettings(settings)); } catch (e) { }
    }

    /* app.preferences から設定文字列を読み出して復元（無ければ null）/ Load and parse saved settings (null if none) */
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

    /* パネルの余白と間隔 / Panel margins and spacing */
    var PANEL_MARGINS = [16, 20, 16, 12];
    var PANEL_SPACING = 8;
    var VALUE_ROW_LABEL_WIDTH = 70;

    /* パネルの共通設定 / Apply shared panel layout */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* 固定幅の右寄せラベル列を追加 / Add a fixed-width right-aligned label column */
    function addFixedLabelColumn(parent, label) {
        var labelColumn = parent.add("group");
        labelColumn.orientation = "row";
        labelColumn.alignChildren = ["right", "center"];
        labelColumn.margins = 0;
        labelColumn.spacing = 0;
        labelColumn.preferredSize.width = VALUE_ROW_LABEL_WIDTH;
        return labelColumn.add("statictext", undefined, label);
    }

    /* 同一グループ外のラジオボタンを手動で排他化 / Enforce exclusivity for radios in different parent groups */
    function selectExclusiveRadio(radios, selectedRadio) {
        for (var i = 0; i < radios.length; i++) {
            radios[i].value = (radios[i] === selectedRadio);
        }
    }

    /* ↑↓キーで値を増減
     * - ↑↓: ±1
     * - Shift+↑↓: ±10（10の倍数にスナップ）
     * minValue を指定すると、それ未満には下がらない（指定省略時は allowNegative=false で 0 が下限）
     * Pass minValue to clamp the lower bound (otherwise 0 when allowNegative is false) */
    function changeValueByArrowKey(editText, allowNegative, onChange, minValue) {
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
            if (typeof minValue === "number" && next < minValue) next = minValue;

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
                if (item.layer && isSymbolListName(item.layer.name)) continue;
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

    /* オブジェクトの中心が指定アートボード矩形の内側にあるか / Is the item's center inside the artboard rect
     * Illustrator 座標系：x は左<右、y は上>下（上方向が正）/ x grows right, y grows up */
    function isItemCenterOnArtboard(itemBounds, artboardRect, tolerance) {
        var centerX = (itemBounds[0] + itemBounds[2]) / 2;
        var centerY = (itemBounds[1] + itemBounds[3]) / 2;
        return centerX >= artboardRect[0] - tolerance &&
            centerX <= artboardRect[2] + tolerance &&
            centerY <= artboardRect[1] + tolerance &&
            centerY >= artboardRect[3] - tolerance;
    }

    /* 指定アートボード上に乗っているオブジェクトをすべて削除（過去版が別レイヤーへ移されていても掴めるよう全レイヤー対象）
     * keepLayer を指定すると、そのレイヤー上のアイテムは保護（同位置・同寸の新アートボードに作った
     * 新しいシンボル・キャプション・背景塗りを巻き添えで消さないため）。
     * Remove every page item whose center sits on the artboard rect; preserves any item on keepLayer
     * so freshly-created symbols/captions/background are not collateral-deleted when old/new artboards
     * happen to share bounds. */
    function removeItemsOnArtboard(doc, artboardRect, keepLayer) {
        var tolerance = 1.0;
        var allItems = doc.pageItems;
        var toRemove = [];
        for (var i = 0; i < allItems.length; i++) {
            var item = allItems[i];
            try {
                if (keepLayer && item.layer === keepLayer) continue;
                if (isItemCenterOnArtboard(item.geometricBounds, artboardRect, tolerance)) {
                    toRemove.push(item);
                }
            } catch (boundsError) { }
        }
        /* 親グループが先に消えると子の remove が投げるが、try/catch で握りつぶす
         * If a parent group is removed first, removing its child throws; swallow via try/catch */
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
            if (isSymbolListName(doc.artboards[i].name) && doc.artboards.length > 1) {
                /* 旧アートボード上のオブジェクトをすべて先に掃除。新レイヤー上のものは保護
                 * Sweep every item on the old artboard first; never touch the just-created (kept) layer */
                removeItemsOnArtboard(doc, doc.artboards[i].artboardRect, state.layer);
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
            if (isSymbolListName(layer.name) && doc.layers.length > 1) {
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

    /* キャンバス上で最も右下に位置するアートボードの番号（1 始まり）を返す。
     * 右端（rect[2] 最大）と下端（rect[3] 最小）を同時に満たす隅を、score = right − bottom で評価。
     * 「シンボル一覧」アートボードは基準候補から除外する。
     * Return the 1-based number of the bottom-right-most artboard on the canvas
     * (max right edge + lowest bottom edge), excluding シンボル一覧 artboards. */
    function findBottomRightArtboardNumber(doc) {
        var bestNumber = 1;
        var bestScore = null;
        for (var i = 0; i < doc.artboards.length; i++) {
            if (isSymbolListName(doc.artboards[i].name)) continue;
            var rect = doc.artboards[i].artboardRect; // [left, top, right, bottom]
            var score = rect[2] - rect[3]; // 右に大きく・下に大きいほど大 / larger toward bottom-right
            if (bestScore === null || score > bestScore) {
                bestScore = score;
                bestNumber = i + 1;
            }
        }
        return bestNumber;
    }

    /* 基準アートボードの矩形を取得（基準モード別）。
     * - specified: 指定番号（1 始まり、範囲外はクランプ）のアートボード
     * - last + 更新 ON: 既存「シンボル一覧」は OK 時に消えるので除外し、最後の通常アートボード
     * - last + 更新 OFF: ドキュメント末尾のアートボード
     * Resolve the base artboard rect by base mode (specified number / last, update-aware). */
    function resolveBaseArtboardRect(doc, settings) {
        if (settings.baseMode === "specified") {
            var idx = settings.baseArtboardNumber - 1;
            if (idx < 0) idx = 0;
            if (idx > doc.artboards.length - 1) idx = doc.artboards.length - 1;
            return doc.artboards[idx].artboardRect;
        }
        if (!settings.update) {
            return doc.artboards[doc.artboards.length - 1].artboardRect;
        }
        for (var i = doc.artboards.length - 1; i >= 0; i--) {
            if (!isSymbolListName(doc.artboards[i].name)) {
                return doc.artboards[i].artboardRect;
            }
        }
        return doc.artboards[doc.artboards.length - 1].artboardRect;
    }

    /* 基準アートボードの右／下に新アートボードを置く左上座標を決定
     * Compute the new artboard origin (top-left) to the right/below the base artboard */
    function computeArtboardOrigin(doc, settings) {
        var baseRect = resolveBaseArtboardRect(doc, settings);
        if (settings.position === "right") {
            return { left: baseRect[2] + settings.artboardGap, top: baseRect[1] };
        }
        return { left: baseRect[0], top: baseRect[3] - settings.artboardGap };
    }

    /* アートボードを追加し、識別用情報を返す / Add artboard and return identifying info */
    function addArtboard(doc, left, top, width, height, name) {
        var index = doc.artboards.length;
        var rect = [left, top, left + width, top - height];
        var artboard = doc.artboards.add(rect);
        artboard.name = name;
        return { index: index, name: name, rect: rect };
    }

    /* 指定アートボードをアクティブにし、その中心に指定倍率でズーム
     * Activate the artboard and center the view on it at the given zoom factor */
    function zoomToArtboard(doc, index, zoomFactor) {
        if (index === null || index < 0 || index >= doc.artboards.length) return;
        try { doc.artboards.setActiveArtboardIndex(index); } catch (e) { }
        try {
            var rect = doc.artboards[index].artboardRect; // [left, top, right, bottom]
            var center = [(rect[0] + rect[2]) / 2, (rect[1] + rect[3]) / 2];
            var view = doc.views[0];
            view.zoom = zoomFactor;
            view.centerPoint = center;
        } catch (e) { }
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
    // ビュー状態の退避・復元 / View state capture & restore
    // =========================================

    /* 現在のビュー状態（view/zoom/centerPoint）を退避 / Capture current view state (view/zoom/center) */
    function captureViewState(doc) {
        var st = { view: null, zoom: null, center: null };
        try {
            st.view = doc.activeView;
            st.zoom = st.view.zoom;
            st.center = st.view.centerPoint;
        } catch (_) { }
        return st;
    }

    /* 退避したビュー状態（zoom/centerPoint）を復元 / Restore a previously captured view state */
    function restoreViewState(doc, state) {
        if (!state) return;
        try {
            var v = state.view || doc.activeView;
            if (v && state.zoom != null) v.zoom = state.zoom;
            if (v && state.center != null) v.centerPoint = state.center;
        } catch (_) { }
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

        var labelControl = addFixedLabelColumn(row, label);

        var input = row.add("edittext", undefined, String(Math.round(ptToUnit(defaultPt))));
        input.characters = 4;
        row.add("statictext", undefined, UNIT.label);
        if (tooltip) {
            labelControl.helpTip = tooltip;
            input.helpTip = tooltip;
        }
        return input;
    }

    /* 「作成位置」パネル。基準（最終アートボード／指定番号）と方向（右側／下側）の
     * 2 つの独立ラジオグループ＋間隔入力で構成。指定番号の既定値はキャンバス最右下のアートボード番号。
     * Build the "Location" panel: base (last / specified number) and direction (right / below)
     * as two independent radio groups, plus the gap input. The specified-number default is the
     * bottom-right-most artboard number on the canvas. */
    function buildPlacementPanel(parent, savedSettings, doc) {
        var panel = parent.add("panel", undefined, L(LABELS.panel.place));
        setupPanel(panel);

        var savedPosition = getSavedSetting(savedSettings, "position", "below");
        var savedBaseMode = getSavedSetting(savedSettings, "baseMode", "last");

        /* 基準：最終アートボード／指定［番号］/ Base: last artboard or specified [number] */
        var baseGroup = panel.add("group");
        baseGroup.orientation = "column";
        baseGroup.alignChildren = "left";
        baseGroup.spacing = 6;

        var baseLastRadio = baseGroup.add("radiobutton", undefined, L(LABELS.radio.baseLast));
        baseLastRadio.helpTip = L(LABELS.tip.baseLast);

        var baseSpecRow = baseGroup.add("group");
        baseSpecRow.orientation = "row";
        baseSpecRow.alignChildren = ["left", "center"];
        baseSpecRow.spacing = 6;
        var baseSpecifiedRadio = baseSpecRow.add("radiobutton", undefined, L(LABELS.radio.baseSpecified));
        baseSpecifiedRadio.helpTip = L(LABELS.tip.baseSpecified);
        /* 既定値はキャンバス上で最も右下のアートボード番号を自動採用 / Default to bottom-right-most artboard */
        var baseNumberInput = baseSpecRow.add("edittext", undefined, String(findBottomRightArtboardNumber(doc)));
        baseNumberInput.characters = 4;
        baseNumberInput.helpTip = L(LABELS.tip.baseArtboardNumber);

        baseLastRadio.value = (savedBaseMode !== "specified");
        baseSpecifiedRadio.value = (savedBaseMode === "specified");

        var artboardGapInput = addValueRow(panel, labelText(LABELS.field.gap),
            getSavedSetting(savedSettings, "artboardGap", DEFAULT_ARTBOARD_GAP_PT), L(LABELS.tip.artboardGap));

        /* 方向：右側／下側 / Direction: right or below */
        var directionGroup = panel.add("group");
        directionGroup.orientation = "row";
        directionGroup.alignChildren = ["left", "center"];
        directionGroup.spacing = 6;

        /* 間隔行と同じ右寄せ固定幅列にして桁を揃える / Align with other value-row labels */
        addFixedLabelColumn(directionGroup, labelText(LABELS.label.direction));

        var directionRadios = directionGroup.add("group");
        directionRadios.orientation = "row";
        directionRadios.alignChildren = ["left", "center"];
        directionRadios.spacing = 10;
        var rightRadio = directionRadios.add("radiobutton", undefined, L(LABELS.radio.directionRight));
        rightRadio.helpTip = L(LABELS.tip.directionRight);
        var belowRadio = directionRadios.add("radiobutton", undefined, L(LABELS.radio.directionBelow));
        belowRadio.helpTip = L(LABELS.tip.directionBelow);
        rightRadio.value = (savedPosition === "right");
        belowRadio.value = (savedPosition !== "right");

        return {
            baseLastRadio: baseLastRadio,
            baseSpecifiedRadio: baseSpecifiedRadio,
            baseNumberInput: baseNumberInput,
            rightRadio: rightRadio,
            belowRadio: belowRadio,
            artboardGapInput: artboardGapInput
        };
    }

    /* 「サイズと余白」パネル（幅・高さ・内側余白）を構築 / Build the "Size & padding" panel (width/height/inner padding) */
    function buildMarginPanel(parent, savedSettings) {
        var panel = parent.add("panel", undefined, L(LABELS.panel.size));
        setupPanel(panel);
        var widthInput = addValueRow(panel, labelText(LABELS.field.width), 0, L(LABELS.tip.width));
        var heightInput = addValueRow(panel, labelText(LABELS.field.height), 0, L(LABELS.tip.height));
        var marginInput = addValueRow(panel, labelText(LABELS.field.margin),
            getSavedSetting(savedSettings, "margin", DEFAULT_ARTBOARD_MARGIN_PT), L(LABELS.tip.margin));
        return {
            widthInput: widthInput,
            heightInput: heightInput,
            marginInput: marginInput
        };
    }

    /* 「背景」パネル（なし／黒／白／グレーの 2×2 ラジオ）を構築 / Build the "Background" panel (none/black/white/gray 2x2 radios) */
    function buildBgColorPanel(parent, savedSettings) {
        var panel = parent.add("panel", undefined, L(LABELS.panel.bgColor));
        setupPanel(panel);

        /* 2 行 2 列。各ラジオの幅を揃えて列をきれいに整える / 2x2 grid; equal widths align columns */
        var BG_RADIO_WIDTH = 60;

        var row1 = panel.add("group");
        row1.orientation = "row";
        row1.spacing = 10;
        var bgNoneRadio = row1.add("radiobutton", undefined, L(LABELS.radio.bgNone));
        bgNoneRadio.helpTip = L(LABELS.tip.bgNone);
        bgNoneRadio.preferredSize.width = BG_RADIO_WIDTH;
        var bgBlackRadio = row1.add("radiobutton", undefined, L(LABELS.radio.bgBlack));
        bgBlackRadio.helpTip = L(LABELS.tip.bgBlack);
        bgBlackRadio.preferredSize.width = BG_RADIO_WIDTH;

        var row2 = panel.add("group");
        row2.orientation = "row";
        row2.spacing = 10;
        var bgWhiteRadio = row2.add("radiobutton", undefined, L(LABELS.radio.bgWhite));
        bgWhiteRadio.helpTip = L(LABELS.tip.bgWhite);
        bgWhiteRadio.preferredSize.width = BG_RADIO_WIDTH;
        var bgGrayRadio = row2.add("radiobutton", undefined, L(LABELS.radio.bgGray));
        bgGrayRadio.helpTip = L(LABELS.tip.bgGray);
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

    /* 「キャプション」パネル（表示 ON/OFF・上下位置・フォントサイズ）を構築 / Build the "Caption" panel (toggle/position/font size) */
    function buildCaptionPanel(parent, savedSettings) {
        var panel = parent.add("panel", undefined, L(LABELS.panel.caption));
        setupPanel(panel);

        var savedShowCaption = getSavedSetting(savedSettings, "showCaption", false);
        var savedCaptionPosition = getSavedSetting(savedSettings, "captionPosition", "below");

        /* シンボル名表示の ON/OFF / Toggle caption visibility */
        var showCaptionCheckbox = panel.add("checkbox", undefined, L(LABELS.checkbox.showCaption));
        showCaptionCheckbox.helpTip = L(LABELS.tip.showCaption);
        showCaptionCheckbox.value = (savedShowCaption !== false);

        /* 表示位置（横並び 2 ラジオ）/ Caption position as 2-way horizontal radio */
        var posRow = panel.add("group");
        posRow.orientation = "row";
        posRow.alignChildren = ["left", "center"];
        posRow.spacing = 6;
        posRow.add("statictext", undefined, labelText(LABELS.label.captionPosition));
        var posRadios = posRow.add("group");
        posRadios.orientation = "row";
        posRadios.alignChildren = "left";
        posRadios.spacing = 10;
        var captionAboveRadio = posRadios.add("radiobutton", undefined, L(LABELS.radio.captionAbove));
        captionAboveRadio.helpTip = L(LABELS.tip.captionAbove);
        var captionBelowRadio = posRadios.add("radiobutton", undefined, L(LABELS.radio.captionBelow));
        captionBelowRadio.helpTip = L(LABELS.tip.captionBelow);
        captionAboveRadio.value = (savedCaptionPosition === "above");
        captionBelowRadio.value = (savedCaptionPosition !== "above");

        /* フォントサイズ（単位は環境設定 text/units に従う）/ Font size in the user's text-unit pref */
        var fontSizeRow = panel.add("group");
        fontSizeRow.orientation = "row";
        fontSizeRow.alignChildren = "center";
        fontSizeRow.spacing = 6;
        var fsLabel = addFixedLabelColumn(fontSizeRow, labelText(LABELS.field.fontSize));
        fsLabel.helpTip = L(LABELS.tip.fontSize);
        var savedFontSize = getSavedSetting(savedSettings, "fontSize", DEFAULT_CAPTION_FONT_SIZE_PT);
        var fontSizeInput = fontSizeRow.add("edittext", undefined, String(Math.round(ptToTypeUnit(savedFontSize))));
        fontSizeInput.characters = 4;
        fontSizeInput.helpTip = L(LABELS.tip.fontSize);
        fontSizeRow.add("statictext", undefined, TYPE_UNIT.label);

        return {
            showCaptionCheckbox: showCaptionCheckbox,
            posRow: posRow,
            captionAboveRadio: captionAboveRadio,
            captionBelowRadio: captionBelowRadio,
            fontSizeRow: fontSizeRow,
            fontSizeInput: fontSizeInput
        };
    }

    /* 「収集対象」パネル（すべて／使用中のみ）を構築。起動時は常に「すべて」 / Build the "Collect" panel (all / used only; always "all" at launch) */
    function buildTargetPanel(parent) {
        var panel = parent.add("panel", undefined, L(LABELS.panel.target));
        setupPanel(panel);
        /* ラジオを横並びに / Lay out radios in a row */
        var row = panel.add("group");
        row.orientation = "row";
        row.alignChildren = ["left", "center"];
        row.spacing = 12;
        var filterAllRadio = row.add("radiobutton", undefined, L(LABELS.radio.filterAll));
        filterAllRadio.helpTip = L(LABELS.tip.filterAll);
        var filterUsedRadio = row.add("radiobutton", undefined, L(LABELS.radio.filterUsed));
        filterUsedRadio.helpTip = L(LABELS.tip.filterUsed);
        /* 起動時は常に「すべて」を選択（保存値は無視）/ Always start with "all" on launch */
        filterAllRadio.value = true;
        filterUsedRadio.value = false;
        return {
            filterAllRadio: filterAllRadio,
            filterUsedRadio: filterUsedRadio
        };
    }

    /* 「並べ方」パネル（シンボル間隔・最大幅）を構築 / Build the "Layout" panel (symbol gap / max row width) */
    function buildSymbolPanel(parent, savedSettings) {
        var panel = parent.add("panel", undefined, L(LABELS.panel.placement));
        setupPanel(panel);

        var symbolGapInput = addValueRow(panel, labelText(LABELS.field.gap),
            getSavedSetting(savedSettings, "symbolGap", DEFAULT_SYMBOL_GAP_PT), L(LABELS.tip.symbolGap));
        var maxRowWidthInput = addValueRow(panel, labelText(LABELS.field.maxRowWidth),
            getSavedSetting(savedSettings, "maxRowWidth", DEFAULT_MAX_ROW_WIDTH_PT), L(LABELS.tip.maxRowWidth));

        return {
            symbolGapInput: symbolGapInput,
            maxRowWidthInput: maxRowWidthInput
        };
    }

    /* 「作成するアートボード」パネル末尾の更新チェックボックスを構築 / Build the update checkbox at the bottom of the artboard panel */
    function buildUpdateCheckbox(parent) {
        var row = parent.add("group");
        row.orientation = "row";
        row.alignChildren = ["left", "center"];
        row.margins = 0;
        var updateCheckbox = row.add("checkbox", undefined, L(LABELS.checkbox.update));
        updateCheckbox.value = true;
        updateCheckbox.helpTip = L(LABELS.tip.update);
        return { updateCheckbox: updateCheckbox };
    }

    /* 最下段のボタン行（左：表示ボタン／右：キャンセル・OK）を構築 / Build the bottom button row (left: view-fit buttons / right: Cancel, OK) */
    function buildButtonRow(parent) {
        var row = parent.add("group");
        row.alignment = "fill";
        row.orientation = "row";
        row.spacing = 8;

        /* 左：表示ボタン（作成したアートボードのみ／全体）/ Left: view-fit buttons */
        var leftCol = row.add("group");
        leftCol.alignment = ["left", "center"];
        leftCol.spacing = 8;
        var fitCreatedButton = leftCol.add("button", undefined, L(LABELS.button.fitCreated));
        fitCreatedButton.helpTip = L(LABELS.tip.fitCreated);
        var fitAllButton = leftCol.add("button", undefined, L(LABELS.button.fitAll));
        fitAllButton.helpTip = L(LABELS.tip.fitAll);

        /* 中央：スペーサー / Center: spacer */
        var spacer = row.add("group");
        spacer.alignment = ["fill", "center"];

        /* 右：キャンセル / OK / Right: Cancel / OK */
        var rightCol = row.add("group");
        rightCol.alignment = ["right", "center"];
        rightCol.spacing = 8;
        var cancelButton = rightCol.add("button", undefined, L(LABELS.button.cancel), { name: "cancel" });
        var okButton = rightCol.add("button", undefined, "OK", { name: "OK" });

        return {
            fitCreatedButton: fitCreatedButton,
            fitAllButton: fitAllButton,
            okButton: okButton,
            cancelButton: cancelButton
        };
    }

    /* ダイアログを構築し、コントロール参照を返す / Build dialog and return refs */
    function buildDialog(savedSettings, doc) {
        var dialog = new Window("dialog", L(LABELS.dialog.title) + " " + SCRIPT_VERSION);
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
        var artboardGroupPanel = leftColumn.add("panel", undefined, L(LABELS.panel.artboard));
        setupPanel(artboardGroupPanel);

        var placementRefs = buildPlacementPanel(artboardGroupPanel, savedSettings, doc);
        var marginRefs = buildMarginPanel(artboardGroupPanel, savedSettings);
        var bgColorRefs = buildBgColorPanel(artboardGroupPanel, savedSettings);
        /* 「作成するアートボード」パネル末尾の更新チェック / Update checkbox at the bottom of the artboard panel */
        var updateRefs = buildUpdateCheckbox(artboardGroupPanel);

        /* シンボル関連パネルを内包するラッパー / Wrapper for symbol-related panels */
        var symbolGroupPanel = rightColumn.add("panel", undefined, L(LABELS.panel.symbolGroup));
        setupPanel(symbolGroupPanel);
        var targetRefs = buildTargetPanel(symbolGroupPanel);
        var symbolRefs = buildSymbolPanel(symbolGroupPanel, savedSettings);
        var captionRefs = buildCaptionPanel(symbolGroupPanel, savedSettings);

        var buttonRefs = buildButtonRow(dialog);

        return {
            dialog: dialog,
            baseLastRadio: placementRefs.baseLastRadio,
            baseSpecifiedRadio: placementRefs.baseSpecifiedRadio,
            baseNumberInput: placementRefs.baseNumberInput,
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
            updateCheckbox: updateRefs.updateCheckbox,
            fontSizeInput: captionRefs.fontSizeInput,
            filterAllRadio: targetRefs.filterAllRadio,
            filterUsedRadio: targetRefs.filterUsedRadio,
            showCaptionCheckbox: captionRefs.showCaptionCheckbox,
            captionPosRow: captionRefs.posRow,
            captionAboveRadio: captionRefs.captionAboveRadio,
            captionBelowRadio: captionRefs.captionBelowRadio,
            fontSizeRow: captionRefs.fontSizeRow,
            symbolGapInput: symbolRefs.symbolGapInput,
            maxRowWidthInput: symbolRefs.maxRowWidthInput,
            fitCreatedButton: buttonRefs.fitCreatedButton,
            fitAllButton: buttonRefs.fitAllButton,
            okButton: buttonRefs.okButton,
            cancelButton: buttonRefs.cancelButton
        };
    }

    /* 入力テキストを pt に変換 / Read input text as points */
    function readInputPt(inputText, defaultPt) {
        var parsedValue = parseFloat(inputText);
        return isNaN(parsedValue) ? defaultPt : unitToPt(parsedValue);
    }

    /* 背景色ラジオの選択を文字列で返す / Read the selected background color as a string */
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

    /* 指定基準のアートボード番号を読む（1 始まり、不正値は 1）/ Read specified base artboard number (1-based) */
    function readBaseArtboardNumber(controls) {
        var n = parseInt(controls.baseNumberInput.text, 10);
        if (isNaN(n) || n < 1) n = 1;
        return n;
    }

    function readDialogSettings(controls) {
        return {
            position: controls.rightRadio.value ? "right" : "below",
            baseMode: controls.baseSpecifiedRadio.value ? "specified" : "last",
            baseArtboardNumber: readBaseArtboardNumber(controls),
            artboardGap: readInputPt(controls.artboardGapInput.text, DEFAULT_ARTBOARD_GAP_PT),
            margin: readInputPt(controls.marginInput.text, DEFAULT_ARTBOARD_MARGIN_PT),
            update: controls.updateCheckbox.value,
            showCaption: controls.showCaptionCheckbox.value,
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
        return { refresh: refresh, state: state, clearOverrides: clearOverrides, doc: doc };
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
            selectExclusiveRadio(bgRadios, target);
        }
        for (var k = 0; k < bgRadios.length; k++) {
            (function (radio) {
                radio.onClick = function () {
                    applyBgExclusive(radio);
                    clearAndRefresh();
                };
            })(bgRadios[k]);
        }

        /* 「シンボル名を表示」OFF のとき表示位置とフォントサイズ行をディム
         * Dim position row and font-size row when caption is off */
        function updateCaptionEnabled() {
            var on = controls.showCaptionCheckbox.value;
            controls.captionPosRow.enabled = on;
            controls.fontSizeRow.enabled = on;
        }

        /* 基準が「指定」のときだけ番号入力欄を有効化 / Enable the number field only when base is "specified" */
        function updateBaseEnabled() {
            controls.baseNumberInput.enabled = controls.baseSpecifiedRadio.value;
        }

        /* 作成位置（基準・方向・基準番号）の変更時は再構築後に明示的に再描画
         * On any position change (base/direction/number), rebuild then force a redraw */
        function refreshPlacement() {
            clearAndRefresh();
            app.redraw();
        }

        /* 基準ラジオ（最終アートボード／指定）は親グループが異なり ScriptUI の排他選択が効かないため手動で排他化。
         * The two base radios live in separate parent groups, so enforce exclusivity manually. */
        var baseRadios = [controls.baseLastRadio, controls.baseSpecifiedRadio];
        controls.baseLastRadio.onClick = function () {
            selectExclusiveRadio(baseRadios, controls.baseLastRadio);
            updateBaseEnabled();
            refreshPlacement();
        };
        controls.baseSpecifiedRadio.onClick = function () {
            selectExclusiveRadio(baseRadios, controls.baseSpecifiedRadio);
            updateBaseEnabled();
            refreshPlacement();
        };

        /* 作成方向（右側／下側）も位置変更なので再描画付き / Direction radios are placement changes too */
        controls.rightRadio.onClick = refreshPlacement;
        controls.belowRadio.onClick = refreshPlacement;

        /* 寸法以外の操作は幅・高さオーバーライドを解除して再描画
         * Non-size triggers clear the width/height override before refreshing */
        var triggers = [
            controls.updateCheckbox,
            controls.filterAllRadio, controls.filterUsedRadio
        ];
        for (var i = 0; i < triggers.length; i++) triggers[i].onClick = clearAndRefresh;
        /* 基準番号は 1 未満を 1 に補正してから再描画（位置変更なので redraw 付き）/ Clamp base number to >= 1, then refresh+redraw */
        function clampBaseNumberAndRefresh() {
            var n = parseInt(controls.baseNumberInput.text, 10);
            if (isNaN(n) || n < 1) n = 1;
            controls.baseNumberInput.text = String(n);
            refreshPlacement();
        }
        controls.baseNumberInput.onChange = clampBaseNumberAndRefresh;
        changeValueByArrowKey(controls.baseNumberInput, false, clampBaseNumberAndRefresh, 1);
        updateBaseEnabled();

        controls.showCaptionCheckbox.onClick = function () {
            updateCaptionEnabled();
            clearAndRefresh();
        };

        var captionRadios = [controls.captionAboveRadio, controls.captionBelowRadio];
        for (var c = 0; c < captionRadios.length; c++) {
            captionRadios[c].onClick = clearAndRefresh;
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
         * 幅を変えたときは「最大幅」も内側の余白を引いた値に追従させる
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

        /* 「表示」パネル：作成したアートボードのみ／全体をウィンドウに合わせる
         * "View" panel: fit the created artboard / all artboards to the window */
        controls.fitCreatedButton.onClick = function () {
            var st = refresher.state.current;
            if (!st || !st.artboardInfo) return;
            try { refresher.doc.artboards.setActiveArtboardIndex(st.artboardInfo.index); } catch (e) { }
            app.executeMenuCommand("fitin"); /* Fit Artboard in Window（アクティブアートボード）*/
            /* フィット後に少し引いて 90% に / Back off slightly to 90% after fitting */
            try {
                var v = refresher.doc.activeView;
                if (v) v.zoom = v.zoom * 0.9;
            } catch (e2) { }
            app.redraw();
        };
        controls.fitAllButton.onClick = function () {
            app.executeMenuCommand("fitall");
            /* フィット後に少し引いて 90% に / Back off slightly to 90% after fitting */
            try {
                var v = refresher.doc.activeView;
                if (v) v.zoom = v.zoom * 0.9;
            } catch (e3) { }
            app.redraw();
        };
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
        var zoomState = captureViewState(doc);
        var controls = buildDialog(savedSettings, doc);
        var refresher = makePreviewRefresher(doc, controls);
        wireDialogEvents(controls, refresher);

        /* 起動直後に初回プレビューを表示し、新アートボードの中心に 60% ズーム
         * Initial preview, then center on the new artboard at 60% zoom */
        refresher.refresh();
        if (refresher.state.current) {
            zoomToArtboard(doc, refresher.state.current.artboardInfo.index, 0.6);
            app.redraw();
        }

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
                alert(L(LABELS.message.noUsedSymbols));
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
        }
        restoreViewState(doc, zoomState);
        app.redraw();
        return null;
    }

    // =========================================
    // メイン
    // =========================================

    /* エントリポイント：ドキュメント／シンボルの有無を確認してダイアログを起動 / Entry point: validate doc/symbols, then open the dialog */
    function main() {
        if (app.documents.length === 0) {
            alert(L(LABELS.message.noDoc));
            return;
        }
        var doc = app.activeDocument;
        if (doc.symbols.length === 0) {
            alert(L(LABELS.message.noSymbols));
            return;
        }
        showDialog(doc);
    }

    main();

})();