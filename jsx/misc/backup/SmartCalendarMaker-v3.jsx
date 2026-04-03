#target illustrator
#targetengine "SmartCalendarMaker"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

var SCRIPT_VERSION = "v1.1";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: {
        ja: "カレンダー作成 " + SCRIPT_VERSION,
        en: "Calendar Maker " + SCRIPT_VERSION
    },

    alertNoDoc: { ja: "ドキュメントを開いてから実行してください。", en: "Please open a document before running." },

    // Panels
    panelBase: { ja: "基本設定", en: "Basics" },
    panelYear: { ja: "年", en: "Year" },
    panelMonth: { ja: "月", en: "Month" },
    panelWeekday: { ja: "曜日", en: "Weekday" },
    panelLayout: { ja: "レイアウト", en: "Layout" },
    panelCell: { ja: "セル", en: "Cell" },
    panelFormat: { ja: "書式", en: "Style" },
    panelPreset: { ja: "プリセット", en: "Preset" },

    // Basics
    year: { ja: "年:", en: "Year:" },
    month: { ja: "月:", en: "Month:" },
    day: { ja: "日:", en: "Day:" },

    // Presets / Base
    monthCountLabel: { ja: "月数：", en: "Months:" },
    preset1: { ja: "1ヶ月", en: "1 mo" },
    preset3: { ja: "3ヶ月", en: "3 mo" },
    preset12: { ja: "12ヶ月", en: "12 mo" },

    baseLabel: { ja: "基準：", en: "Base:" },
    baseCurrent: { ja: "当月基準", en: "Current month" },
    baseJan: { ja: "1月から", en: "From January" },

    // Common UI
    bottomMargin: { ja: "下マージン", en: "Bottom margin" },
    alignLabel: { ja: "揃え：", en: "Align:" },
    left: { ja: "左", en: "Left" },
    center: { ja: "中央", en: "Center" },
    right: { ja: "右", en: "Right" },

    // Month
    chkIncludeYear: { ja: "年を併記", en: "Include year" },
    panelNotation: { ja: "表記", en: "Notation" },
    monthNum: { ja: "数字", en: "Number" },
    monthPad: { ja: "数字（0埋め）", en: "Zero-padded" },
    monthEn: { ja: "英語", en: "English" },
    monthEnS: { ja: "英語（短縮）", en: "English (short)" },
    chkBottomBorder: { ja: "下ボーダー", en: "Bottom border" },

    // Year
    chkShowYear: { ja: "年を表示", en: "Show year" },

    // Weekday
    weekStart: { ja: "週の始まり", en: "Week starts" },
    monday: { ja: "月曜日", en: "Monday" },
    sunday: { ja: "日曜日", en: "Sunday" },
    labelNotation: { ja: "表記", en: "Notation" },
    labelMargin: { ja: "下マージン", en: "Bottom margin" },
    labelFontSize: { ja: "フォントサイズ", en: "Font size" },

    // Layout
    months: { ja: "月数:", en: "Months:" },
    cols: { ja: "列数:", en: "Columns:" },
    width: { ja: "幅", en: "W" },
    height: { ja: "高さ", en: "H" },
    lr: { ja: "左右", en: "L/R" },
    ud: { ja: "上下", en: "U/D" },

    // Format
    font: { ja: "フォント", en: "Font" },
    favorites: { ja: "お気に入り", en: "Favorites" },
    sundayLabel: { ja: "日曜日", en: "Sunday" },
    holidayLabel: { ja: "祝日", en: "Holidays" },

    // Buttons
    preview: { ja: "プレビュー", en: "Preview" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    create: { ja: "作成", en: "OK" },

    // Errors
    errBadDate: { ja: "日付形式が正しくありません。例: 2026-02-14", en: "Invalid date. e.g. 2026-02-14" },
    errBadFontSize: { ja: "フォントサイズが正しくありません。", en: "Invalid font size." },
    errBadCellW: { ja: "セル幅が正しくありません。", en: "Invalid cell width." },
    errBadCellH: { ja: "セル高さが正しくありません。", en: "Invalid cell height." },
    errBuild: { ja: "作成中にエラーが発生しました。", en: "An error occurred while creating." },

    // Loading
    loadingTitle: { ja: "準備中", en: "Loading" },
    loadingText: { ja: "処理中…", en: "Working…" },
    loadingFontsRead: { ja: "フォント一覧を読み込み中…", en: "Loading fonts…" },
    loadingFontsSet: { ja: "フォント一覧をセット中…", en: "Applying fonts…" },
    ready: { ja: "準備完了", en: "Ready" },

    presetSave: { ja: "保存", en: "Save" },
    presetLoad: { ja: "読込", en: "Load" },
    presetSaveDialog: { ja: "プリセットを保存", en: "Save preset" },
    presetLoadDialog: { ja: "プリセットを読み込み", en: "Load preset" },
    presetFileFilter: { ja: "プリセットファイル", en: "Preset file" },
    presetSaveHint: { ja: "（.jsxpreset 推奨）", en: "(recommended .jsxpreset)" },
    presetSaved: { ja: "プリセットを保存しました。", en: "Preset saved." },
    presetLoaded: { ja: "プリセットを読み込みました。", en: "Preset loaded." },
    presetErr: { ja: "プリセットの読み込みに失敗しました。", en: "Failed to load preset." }
};

function L(key) {
    try {
        var o = LABELS[key];
        if (!o) return key;
        return o[lang] || o.en || o.ja || key;
    } catch (_) {
        return key;
    }
}

    /* =====================================================
     * SmartCalendarMaker.jsx
     * バージョン: v1.1
     * 概要: 基準日を基準に、指定月数ぶんのカレンダー（月曜はじまり）をアートボード中心に作成します。
     *      ダイアログ上の変更はプレビューで即時反映し、OKで確定レイヤーとして残します。
     * 更新日: 2026-02-14
     * ===================================================== */


(function () {
    if (app.documents.length === 0) {
        alert(L('alertNoDoc'));
        return;
    }

    var doc = app.activeDocument;

    // ===== 単位ユーティリティ（Preferences 参照 / Q・H 切替対応）=====
    // ▼ 各設定キーの意味：
    // - "rulerType"       ：一般（定規の単位）
    // - "strokeUnits"     ：線
    // - "text/units"      ：文字
    // - "text/asianunits" ：東アジア言語のオプション（日本語・中国語など）

    // 単位コード → ラベル（Q/H は getUnitLabel() で分岐）
    var __SCM_unitMap = {
        0: "in",
        1: "mm",
        2: "pt",
        3: "pica",
        4: "cm",
        6: "px",
        7: "ft/in",
        8: "m",
        9: "yd",
        10: "ft"
    };

    function __SCM_getUnitLabel(code, prefKey) {
        // code===5 は Q/H
        if (code === 5) {
            var hKeys = {
                "text/asianunits": true,
                "rulerType": true,
                "strokeUnits": true
            };
            return hKeys[prefKey] ? "H" : "Q";
        }
        return __SCM_unitMap[code] || "pt";
    }

    function __SCM_getPrefUnitLabel(prefKey, fallback) {
        var fb = fallback || "pt";
        try {
            var code = app.preferences.getIntegerPreference(prefKey);
            return __SCM_getUnitLabel(code, prefKey) || fb;
        } catch (_) {
            return fb;
        }
    }

    // 画面上の寸法系はドキュメントの rulerUnits ではなく、Illustrator の一般単位（rulerType）に合わせる
    var unitLabel = __SCM_getPrefUnitLabel("rulerType", "pt");
    // 文字系（フォントサイズ等）の単位は text/units を参照
    var textUnitLabel = __SCM_getPrefUnitLabel("text/units", "pt");

    // --- Unit conversion helpers (rulerType-based) ---
    // 寸法系の換算は Illustrator の一般単位（rulerType）に統一する。
    function __SCM_getPtFactorFromUnitCode(code) {
        switch (code) {
            case 0: return 72.0;                        // in
            case 1: return 72.0 / 25.4;                 // mm
            case 2: return 1.0;                         // pt
            case 3: return 12.0;                        // pica
            case 4: return 72.0 / 2.54;                 // cm
            case 5: return 72.0 / 25.4 * 0.25;          // Q or H（1Q=0.25mm）
            case 6: return 1.0;                         // px（環境依存のため 1px=1pt 扱い）
            case 7: return 72.0 * 12.0;                 // ft/in
            case 8: return 72.0 / 25.4 * 1000.0;        // m
            case 9: return 72.0 * 36.0;                 // yd
            case 10: return 72.0 * 12.0;                // ft
            default: return 1.0;
        }
    }

    function unitValueToPt(v) {
        v = Number(v);
        if (isNaN(v)) return NaN;
        try {
            var code = app.preferences.getIntegerPreference("rulerType");
            var f = __SCM_getPtFactorFromUnitCode(code);
            return v * f;
        } catch (_) { }
        return v;
    }

    function toPtFromUI(editText) {
        return unitValueToPt(editText && editText.text);
    }

    function ptToUnitValue(pt) {
        pt = Number(pt);
        if (isNaN(pt)) return NaN;
        try {
            var code = app.preferences.getIntegerPreference("rulerType");
            var f = __SCM_getPtFactorFromUnitCode(code);
            if (!f) f = 1.0;
            return pt / f;
        } catch (_) { }
        return pt;
    }

    /* ================================
     * Japan Holidays (2020-2035) built-in
     * - Includes substitute holidays (振替休日) and citizen's holidays (国民の休日)
     * - Handles 2020/2021 Olympic special holiday moves
     * ================================ */
    function __pad2(n) { return (n < 10) ? ("0" + n) : String(n); }
    function __dateKey(y, m1, d) { return y + "-" + __pad2(m1) + "-" + __pad2(d); }
    function __isLeap(y) { return (y % 4 === 0 && y % 100 !== 0) || (y % 400 === 0); }
    function __daysInMonth(y, m1) {
        var md = [31, (__isLeap(y) ? 29 : 28), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
        return md[m1 - 1];
    }
    function __dow(y, m1, d) { return (new Date(y, m1 - 1, d)).getDay(); } // 0=Sun..6
    function __nthWeekday(y, m1, weekday0Sun, nth) {
        // weekday0Sun: 0=Sun..6
        var firstDow = __dow(y, m1, 1);
        var offset = (weekday0Sun - firstDow + 7) % 7;
        var day = 1 + offset + (nth - 1) * 7;
        return day;
    }
    function __vernalEquinoxDay(y) {
        // Approximation valid for 1980-2099
        return Math.floor(20.8431 + 0.242194 * (y - 1980) - Math.floor((y - 1980) / 4));
    }
    function __autumnEquinoxDay(y) {
        // Approximation valid for 1980-2099
        return Math.floor(23.2488 + 0.242194 * (y - 1980) - Math.floor((y - 1980) / 4));
    }
    function __buildHolidaysForYear(y) {
        var map = {};
        function add(m1, d, name) { map[__dateKey(y, m1, d)] = name; }

        // Fixed / rule-based holidays
        add(1, 1, "元日");
        add(1, __nthWeekday(y, 1, 1, 2), "成人の日"); // 2nd Mon
        add(2, 11, "建国記念の日");
        add(2, 23, "天皇誕生日"); // since 2020
        add(3, __vernalEquinoxDay(y), "春分の日");
        add(4, 29, "昭和の日");
        add(5, 3, "憲法記念日");
        add(5, 4, "みどりの日");
        add(5, 5, "こどもの日");

        // Marine / Mountain / Sports (with Olympic special cases)
        if (y === 2020) {
            add(7, 23, "海の日");
            add(7, 24, "スポーツの日");
            add(8, 10, "山の日");
        } else if (y === 2021) {
            add(7, 22, "海の日");
            add(7, 23, "スポーツの日");
            add(8, 8, "山の日"); // substitute handled later
        } else {
            add(7, __nthWeekday(y, 7, 1, 3), "海の日");      // 3rd Mon
            add(8, 11, "山の日");
            add(10, __nthWeekday(y, 10, 1, 2), "スポーツの日"); // 2nd Mon
        }

        add(9, __nthWeekday(y, 9, 1, 3), "敬老の日"); // 3rd Mon
        add(9, __autumnEquinoxDay(y), "秋分の日");
        add(11, 3, "文化の日");
        add(11, 23, "勤労感謝の日");

        // --- Substitute holiday (振替休日): holiday on Sunday -> next non-holiday weekday
        var keys = [];
        for (var k in map) if (map.hasOwnProperty(k)) keys.push(k);
        keys.sort();
        for (var i = 0; i < keys.length; i++) {
            var key = keys[i];
            var parts = key.split("-");
            var m1 = Number(parts[1]);
            var d = Number(parts[2]);
            if (__dow(y, m1, d) === 0) {
                var dd = d;
                var mm = m1;
                while (true) {
                    dd++;
                    var dim = __daysInMonth(y, mm);
                    if (dd > dim) { dd = 1; mm++; }
                    var k2 = __dateKey(y, mm, dd);
                    if (!map[k2]) {
                        map[k2] = "振替休日";
                        break;
                    }
                }
            }
        }

        // --- Citizen's holiday (国民の休日): a working day between two holidays
        // Iterate through all days in year
        var m;
        for (m = 1; m <= 12; m++) {
            var dim2 = __daysInMonth(y, m);
            for (var day = 1; day <= dim2; day++) {
                var k0 = __dateKey(y, m, day);
                if (map[k0]) continue;
                // prev
                var py = y, pm = m, pd = day - 1;
                if (pd < 1) { pm--; if (pm < 1) continue; pd = __daysInMonth(y, pm); }
                var kyPrev = __dateKey(y, pm, pd);
                // next
                var ny = y, nm = m, nd = day + 1;
                var dim3 = __daysInMonth(y, nm);
                if (nd > dim3) { nm++; if (nm > 12) continue; nd = 1; }
                var kyNext = __dateKey(y, nm, nd);
                if (map[kyPrev] && map[kyNext]) {
                    map[k0] = "国民の休日";
                }
            }
        }

        return map;
    }
    function __buildHolidayMap(fromY, toY) {
        var all = {};
        for (var yy = fromY; yy <= toY; yy++) {
            var m = __buildHolidaysForYear(yy);
            for (var k in m) if (m.hasOwnProperty(k)) all[k] = m[k];
        }
        return all;
    }
    var JP_HOLIDAYS_2020_2035 = __buildHolidayMap(2020, 2035);
    function getJPHolidayName(yyyy, m1, d) {
        return JP_HOLIDAYS_2020_2035[__dateKey(yyyy, m1, d)] || "";
    }

    var PREVIEW_LAYER_NAME = "__CAL_PREVIEW__";

    // 既存プレビューが残っていたら消す
    removeLayerIfExists(doc, PREVIEW_LAYER_NAME);

    // ===== ダイアログ =====
    var today = new Date();

    var dlg = new Window("dialog", L("dialogTitle"));
    // ===== ダイアログ位置・透明度調整 =====
    var offsetX = 300;
    var offsetY = 0;
    var dialogOpacity = 0.98;

    function shiftDialogPosition(dlg, offsetX, offsetY) {
        dlg.onShow = function () {
            try {
                var currentX = dlg.location[0];
                var currentY = dlg.location[1];
                dlg.location = [currentX + offsetX, currentY + offsetY];
            } catch (_) { }
            try { refreshPreview(); } catch (_) { }
        };
    }

    function setDialogOpacity(dlg, opacityValue) {
        try { dlg.opacity = opacityValue; } catch (_) { }
    }

    setDialogOpacity(dlg, dialogOpacity);
    shiftDialogPosition(dlg, offsetX, offsetY);
    dlg.orientation = "column";
    dlg.alignChildren = "fill";

    // ===== プリセット（上部・全幅） =====
    var pnlPresetTop = dlg.add("panel", undefined, L("panelPreset"));
    pnlPresetTop.orientation = "row";
    pnlPresetTop.alignChildren = ["left", "center"];
    pnlPresetTop.alignment = "fill";
    pnlPresetTop.margins = [15, 20, 15, 10];

    var gPresetBtns = pnlPresetTop.add("group");
    gPresetBtns.orientation = "row";
    gPresetBtns.alignChildren = ["left", "center"];

    var presetLoadBtn = gPresetBtns.add("button", undefined, L("presetLoad"));
    var presetSaveBtn = gPresetBtns.add("button", undefined, L("presetSave"));

    // プリセット一覧（読み込んだプリセットを追加）
    var stPresetList = gPresetBtns.add("statictext", undefined, " ");
    var ddPresetList = gPresetBtns.add("dropdownlist", undefined, ["-"]);
    ddPresetList.preferredSize = [220, 22];
    try { ddPresetList.selection = 0; } catch (_) { }

    // 読み込んだプリセットを保持（セッション中のみ）
    var __SCM_PRESET_STORE = ($.global.__SCM_CACHE = $.global.__SCM_CACHE || {});
    __SCM_PRESET_STORE.presets = __SCM_PRESET_STORE.presets || {}; // label -> preset object
    __SCM_PRESET_STORE.presetOrder = __SCM_PRESET_STORE.presetOrder || []; // labels

    // ===== 組み込みプリセット（初期登録） =====
    (function () {
        function registerBuiltin(label, obj) {
            try {
                if (!__SCM_PRESET_STORE.presets[label]) {
                    __SCM_PRESET_STORE.presets[label] = obj;
                    __SCM_PRESET_STORE.presetOrder.push(label);
                }
            } catch (_) { }
        }

        registerBuiltin("1)EN Month Title", {
            y: 2026, m: 2, d: 14,
            monthCount: 1, colCount: 1,
            startFrom: "current",
            showTopYear: false,
            topYearBottomMargin: 3,
            includeYearInMonthTitle: false,
            monthTitleAlign: "center",
            monthTitleMode: "en",
            monthBottomMargin: 2,
            monthTitleBottomBorder: true,
            weekStart: "mon",
            weekdayLabelMode: "jp",
            weekdayBottomMargin: 2,
            weekdayFontSize: 9,
            cellW: 7,
            cellH: 6.6,
            outerMarginX: 10,
            outerMarginY: 3,
            fontName: "DNPShueiMGoStd-B",
            favFont: "",
            fontSize: 12,
            align: "center",
            sundayRed: true,
            holidayRed: true
        });

        registerBuiltin("2)ENS + Mon Labels", {
            y: 2026, m: 2, d: 14,
            monthCount: 1, colCount: 1,
            startFrom: "current",
            showTopYear: false,
            topYearBottomMargin: 3,
            includeYearInMonthTitle: false,
            monthTitleAlign: "center",
            monthTitleMode: "ens",
            monthBottomMargin: 2,
            monthTitleBottomBorder: true,
            weekStart: "mon",
            weekdayLabelMode: "mon",
            weekdayBottomMargin: 0,
            weekdayFontSize: 8,
            cellW: 8.7,
            cellH: 6.6,
            outerMarginX: 10,
            outerMarginY: 3,
            fontName: "HeganteDisplay-Regular",
            favFont: "Hegante Display Regular",
            fontSize: 12,
            align: "center",
            sundayRed: true,
            holidayRed: true
        });

        registerBuiltin("3)12M Script Layout", {
            y: 2026, m: 2, d: 14,
            monthCount: 12, colCount: 3,
            startFrom: "jan",
            showTopYear: true,
            topYearBottomMargin: 10,
            includeYearInMonthTitle: false,
            monthTitleAlign: "center",
            monthTitleMode: "en",
            monthBottomMargin: 2,
            monthTitleBottomBorder: false,
            weekStart: "mon",
            weekdayLabelMode: "mon",
            weekdayBottomMargin: 0,
            weekdayFontSize: 7,
            cellW: 8.3,
            cellH: 5.1,
            outerMarginX: 10,
            outerMarginY: 6,
            fontName: "CaflischScriptPro-Light",
            favFont: "Caflisch Script Pro Light",
            fontSize: 12,
            align: "center",
            sundayRed: true,
            holidayRed: true
        });

        registerBuiltin("4)3M Bodoni Layout", {
            y: 2026, m: 2, d: 14,
            monthCount: 3, colCount: 3,
            startFrom: "current",
            showTopYear: false,
            topYearBottomMargin: 3,
            includeYearInMonthTitle: false,
            monthTitleAlign: "center",
            monthTitleMode: "pad",
            monthBottomMargin: 2,
            monthTitleBottomBorder: true,
            weekStart: "mon",
            weekdayLabelMode: "mon",
            weekdayBottomMargin: 0,
            weekdayFontSize: 8,
            cellW: 7.3,
            cellH: 6,
            outerMarginX: 10,
            outerMarginY: 3,
            fontName: "BodoniModaRoman-SemiBold",
            favFont: "Bodoni Moda SemiBold",
            fontSize: 12,
            align: "center",
            sundayRed: true,
            holidayRed: true
        });

        registerBuiltin("5)3M DIN Layout", {
            y: 2026, m: 2, d: 14,
            monthCount: 3, colCount: 1,
            startFrom: "current",
            showTopYear: false,
            topYearBottomMargin: 3,
            includeYearInMonthTitle: false,
            monthTitleAlign: "center",
            monthTitleMode: "pad",
            monthBottomMargin: 2,
            monthTitleBottomBorder: true,
            weekStart: "mon",
            weekdayLabelMode: "mtw",
            weekdayBottomMargin: 0,
            weekdayFontSize: 8,
            cellW: 8,
            cellH: 7,
            outerMarginX: 10,
            outerMarginY: 9,
            fontName: "DINCondensedVF-Regular",
            favFont: "DIN Condensed VF Regular",
            fontSize: 20,
            align: "center",
            sundayRed: true,
            holidayRed: true
        });

        // dropdown 初期化
        try {
            ddPresetList.removeAll();
        } catch (_) {
            try { while (ddPresetList.items.length) ddPresetList.items[0].remove(); } catch (__e) { }
        }
        ddPresetList.add("item", "-");
        for (var i = 0; i < __SCM_PRESET_STORE.presetOrder.length; i++) {
            ddPresetList.add("item", __SCM_PRESET_STORE.presetOrder[i]);
        }
        try { ddPresetList.selection = 0; } catch (_) { }
    })();

    function __SCM_addPresetToDropdown(label, obj) {
        try {
            var lb = String(label || "");
            if (!lb || lb === "-") return;
            __SCM_PRESET_STORE.presets[lb] = obj;
            // 既に登録済みなら順序は維持
            var exists = false;
            for (var i = 0; i < __SCM_PRESET_STORE.presetOrder.length; i++) {
                if (__SCM_PRESET_STORE.presetOrder[i] === lb) { exists = true; break; }
            }
            if (!exists) __SCM_PRESET_STORE.presetOrder.push(lb);

            // dd の items を作り直す（重くない件数想定）
            try {
                ddPresetList.removeAll();
            } catch (_) {
                // removeAll が無い環境向け
                try { while (ddPresetList.items.length) ddPresetList.items[0].remove(); } catch (__e) { }
            }
            ddPresetList.add("item", "-");
            for (var j = 0; j < __SCM_PRESET_STORE.presetOrder.length; j++) {
                ddPresetList.add("item", __SCM_PRESET_STORE.presetOrder[j]);
            }
            // 追加したものを選択
            for (var k = 0; k < ddPresetList.items.length; k++) {
                if (ddPresetList.items[k].text === lb) { ddPresetList.selection = k; break; }
            }
        } catch (_) { }
    }

    ddPresetList.onChange = function () {
        try {
            if (!ddPresetList.selection) return;
            var lb = ddPresetList.selection.text;
            if (!lb || lb === "-") return;
            var obj = __SCM_PRESET_STORE.presets[lb];
            if (!obj) return;
            __SCM_applyPreset(obj);
        } catch (_) { }
    };

    // ===== 2カラム =====
    var gCols = dlg.add("group");
    gCols.orientation = "row";
    gCols.alignChildren = ["fill", "top"];

    var gColL = gCols.add("group");
    gColL.orientation = "column";
    gColL.alignChildren = "fill";

    var gColR = gCols.add("group");
    gColR.orientation = "column";
    gColR.alignChildren = "fill";

    // ===== 基本設定パネル =====
    var pnlBaseDate = gColL.add("panel", undefined, L("panelBase"));
    pnlBaseDate.orientation = "column";
    pnlBaseDate.alignChildren = "left";
    pnlBaseDate.margins = [15, 20, 15, 10];

    var g1 = pnlBaseDate.add("group");
    g1.add("statictext", undefined, L("year"));
    var inputY = g1.add("edittext", undefined, String(today.getFullYear()));
    inputY.characters = 4;

    g1.add("statictext", undefined, L("month"));
    var inputM = g1.add("edittext", undefined, String(today.getMonth() + 1));
    inputM.characters = 2;

    g1.add("statictext", undefined, L("day"));
    var inputD = g1.add("edittext", undefined, String(today.getDate()));
    inputD.characters = 2;

    // 月数プリセット（ラジオ）
    var gMonthPreset = pnlBaseDate.add("group");
    gMonthPreset.orientation = "row";
    gMonthPreset.alignChildren = ["left", "center"];

    gMonthPreset.add("statictext", undefined, L("monthCountLabel"));
    var rbPreset1 = gMonthPreset.add("radiobutton", undefined, L("preset1"));
    var rbPreset3 = gMonthPreset.add("radiobutton", undefined, L("preset3"));
    var rbPreset12 = gMonthPreset.add("radiobutton", undefined, L("preset12"));

    function applyMonthPreset(months, cols) {
        try { inputMonths.text = String(months); } catch (_) { }
        try {
            if (typeof cols === "number" && cols > 0) inputCols.text = String(cols);
        } catch (_) { }
        // 「年」panel: 年表示は 12ヶ月 のときだけON
        try {
            chkTopYear.value = (months === 12);
            chkTopYear.enabled = (months === 12);
        } catch (_) { }

        try { gYearMargin.enabled = chkTopYear.value; } catch (_) { }

        // 12ヶ月のときは月タイトルの「年を併記」をOFF
        try {
            if (months === 12) chkMonthYear.value = false;
        } catch (_) { }

        // 12ヶ月 は「1月から」を自動選択
        try {
            if (months === 12) {
                rbStartJan.value = true;
                rbStartCurrent.value = false;
            }
        } catch (_) { }
        try { schedulePreviewRefresh(true); } catch (_) { }
    }

    rbPreset1.onClick = function () { applyMonthPreset(1, 1); };
    rbPreset3.onClick = function () { applyMonthPreset(3, 3); };
    rbPreset12.onClick = function () { applyMonthPreset(12, 3); };

    // 月の起点
    var gMonthStart = pnlBaseDate.add("group");
    gMonthStart.orientation = "row";
    gMonthStart.alignChildren = ["left", "center"];

    gMonthStart.add("statictext", undefined, L("baseLabel"));
    var rbStartCurrent = gMonthStart.add("radiobutton", undefined, L("baseCurrent"));
    var rbStartJan = gMonthStart.add("radiobutton", undefined, L("baseJan"));

    rbStartCurrent.value = true; // デフォルト
    rbStartCurrent.onClick = schedulePreviewRefresh;
    rbStartJan.onClick = schedulePreviewRefresh;




    // ===== 年パネル =====
    var pnlYear = gColR.add("panel", undefined, L("panelYear"));
    pnlYear.orientation = "column";
    pnlYear.alignChildren = "left";
    pnlYear.margins = [15, 20, 15, 10];

    var gYear = pnlYear.add("group");
    gYear.orientation = "row";
    gYear.alignChildren = ["left", "center"];


    var chkTopYear = gYear.add("checkbox", undefined, L("chkShowYear"));
    chkTopYear.value = false;      // 12ヶ月のときだけONにする
    chkTopYear.enabled = false;    // 12ヶ月以外は触れない
    chkTopYear.onClick = function () {
        try { gYearMargin.enabled = chkTopYear.value; } catch (_) { }
        schedulePreviewRefresh(true);
    };

    // 年タイトル下のマージン（pt）
    var gYearMargin = pnlYear.add("group");
    gYearMargin.orientation = "row";
    gYearMargin.alignChildren = ["left", "center"];
    gYearMargin.add("statictext", undefined, L("bottomMargin"));
    var inputTopYearBottomMargin = gYearMargin.add("edittext", undefined, "3");
    inputTopYearBottomMargin.characters = 4;
    gYearMargin.add("statictext", undefined, unitLabel);
    inputTopYearBottomMargin.onChanging = schedulePreviewRefresh;

    try { gYearMargin.enabled = chkTopYear.value; } catch (_) { }

    // ===== 月パネル =====
    var pnlMonth = gColR.add("panel", undefined, L("panelMonth"));
    pnlMonth.orientation = "column";
    pnlMonth.alignChildren = "left";
    pnlMonth.margins = [15, 20, 15, 10];
    pnlMonth.alignment = "fill";
    pnlMonth.minimumSize.height = 44;
    pnlMonth.preferredSize.height = 44;
    pnlMonth.visible = true;
    pnlMonth.enabled = true;


    var gMonth = pnlMonth.add("group");
    gMonth.orientation = "row";
    gMonth.alignChildren = ["left", "center"];

    var chkMonthYear = gMonth.add("checkbox", undefined, L("chkIncludeYear"));
    chkMonthYear.value = true;
    chkMonthYear.onClick = schedulePreviewRefresh;

    // 月タイトルの揃え
    var gMonthAlign = pnlMonth.add("group");
    gMonthAlign.orientation = "row";
    gMonthAlign.alignChildren = ["left", "center"];
    gMonthAlign.add("statictext", undefined, L("alignLabel"));
    var rbMonthAlignL = gMonthAlign.add("radiobutton", undefined, L("left"));
    var rbMonthAlignC = gMonthAlign.add("radiobutton", undefined, L("center"));
    var rbMonthAlignR = gMonthAlign.add("radiobutton", undefined, L("right"));
    rbMonthAlignC.value = true; // default
    rbMonthAlignL.onClick = schedulePreviewRefresh;
    rbMonthAlignC.onClick = schedulePreviewRefresh;
    rbMonthAlignR.onClick = schedulePreviewRefresh;

    // 月タイトルの表記
    var pnlMonthFmt = pnlMonth.add("panel", undefined, L("panelNotation"));
    pnlMonthFmt.orientation = "column";
    pnlMonthFmt.alignChildren = "left";
    pnlMonthFmt.margins = [15, 20, 15, 10];

    var gMonthFmt = pnlMonthFmt.add("group");
    gMonthFmt.orientation = "column";
    gMonthFmt.alignChildren = ["left", "top"];

    var rbMonthNum = gMonthFmt.add("radiobutton", undefined, L("monthNum"));
    var rbMonthPad = gMonthFmt.add("radiobutton", undefined, L("monthPad"));
    var rbMonthEn = gMonthFmt.add("radiobutton", undefined, L("monthEn"));
    var rbMonthEnS = gMonthFmt.add("radiobutton", undefined, L("monthEnS"));

    // デフォルトは 0埋め
    rbMonthPad.value = true;

    rbMonthNum.onClick = schedulePreviewRefresh;
    rbMonthPad.onClick = schedulePreviewRefresh;
    rbMonthEn.onClick = schedulePreviewRefresh;
    rbMonthEnS.onClick = schedulePreviewRefresh;

    // 月タイトル下のマージン（pt）
    var gMonthMargin = pnlMonth.add("group");
    gMonthMargin.orientation = "row";
    gMonthMargin.alignChildren = ["left", "center"];
    gMonthMargin.add("statictext", undefined, L("bottomMargin"));
    var inputMonthBottomMargin = gMonthMargin.add("edittext", undefined, "3");
    inputMonthBottomMargin.characters = 4;
    gMonthMargin.add("statictext", undefined, unitLabel);
    inputMonthBottomMargin.onChanging = schedulePreviewRefresh;

    // 月タイトル行の下ボーダー
    var chkMonthBottomBorder = pnlMonth.add("checkbox", undefined, L("chkBottomBorder"));
    chkMonthBottomBorder.value = true; // デフォルトON
    chkMonthBottomBorder.onClick = schedulePreviewRefresh;
    function getWeekdayLabelMode() {
        // "jp" | "mtw" | "mon"
        try {
            if (rbWdMon && rbWdMon.value) return "mon";
            if (rbWdMTW && rbWdMTW.value) return "mtw";
        } catch (_) { }
        return "jp";
    }

    function getMonthTitleMode() {
        // "num" | "pad" | "en" | "ens"
        try {
            if (rbMonthEnS && rbMonthEnS.value) return "ens";
            if (rbMonthEn && rbMonthEn.value) return "en";
            if (rbMonthNum && rbMonthNum.value) return "num";
        } catch (_) { }
        return "pad"; // default
    }

    function formatMonthTitle(year, month0, includeYear, mode) {
        var m = month0 + 1;

        if (mode === "en" || mode === "ens") {
            var enFull = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
            var enShort = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
            var name = (mode === "ens") ? enShort[month0] : enFull[month0];
            return includeYear ? (String(year) + " " + name) : name;
        }

        var mm = (mode === "num") ? String(m) : pad2(m);
        return includeYear ? (String(year) + "/" + mm) : mm;
    }

    // 曜日表記（panel）
    var pnlWeekdayLabel = gColR.add("panel", undefined, L("panelWeekday"));
    pnlWeekdayLabel.orientation = "column";
    pnlWeekdayLabel.alignChildren = "left";
    pnlWeekdayLabel.margins = [15, 20, 15, 10];

    // 週の始まり（ラベル + ラジオ）
    var gWeekStart = pnlWeekdayLabel.add("group");
    gWeekStart.orientation = "row";
    gWeekStart.alignChildren = ["left", "center"];

    gWeekStart.add("statictext", undefined, L("weekStart"));
    var rbWeekMon = gWeekStart.add("radiobutton", undefined, L("monday"));
    var rbWeekSun = gWeekStart.add("radiobutton", undefined, L("sunday"));

    rbWeekMon.value = true; // デフォルト
    rbWeekMon.onClick = schedulePreviewRefresh;
    rbWeekSun.onClick = schedulePreviewRefresh;

    // 曜日ヘッダの表記（ラベル + ラジオ）
    var gWeekdayLabel = pnlWeekdayLabel.add("group");
    gWeekdayLabel.orientation = "row";
    gWeekdayLabel.alignChildren = ["left", "center"];

    gWeekdayLabel.add("statictext", undefined, L("labelNotation"));
    var rbWdJP = gWeekdayLabel.add("radiobutton", undefined, "月");
    var rbWdMTW = gWeekdayLabel.add("radiobutton", undefined, "M");
    var rbWdMon = gWeekdayLabel.add("radiobutton", undefined, "Mon");

    // 曜日ヘッダ下のマージン（pt）
    var gWeekdayBottomMargin = pnlWeekdayLabel.add("group");
    gWeekdayBottomMargin.orientation = "row";
    gWeekdayBottomMargin.alignChildren = ["left", "center"];
    gWeekdayBottomMargin.add("statictext", undefined, L("labelMargin"));
    var inputWeekdayBottomMargin = gWeekdayBottomMargin.add("edittext", undefined, "2");
    inputWeekdayBottomMargin.characters = 4;
    gWeekdayBottomMargin.add("statictext", undefined, unitLabel);
    inputWeekdayBottomMargin.onChanging = schedulePreviewRefresh;

    // 曜日フォントサイズ（未入力時は全体フォントサイズを使用）
    var gWeekdayFontSize = pnlWeekdayLabel.add("group");
    gWeekdayFontSize.orientation = "row";
    gWeekdayFontSize.alignChildren = ["left", "center"];
    gWeekdayFontSize.add("statictext", undefined, L("labelFontSize"));
    var inputWeekdayFontSize = gWeekdayFontSize.add("edittext", undefined, "12");
    inputWeekdayFontSize.characters = 4;
    gWeekdayFontSize.add("statictext", undefined, textUnitLabel);
    inputWeekdayFontSize.onChanging = schedulePreviewRefresh;
    changeValueByArrowKey(inputWeekdayFontSize, { integer: false, min: 0.1, max: 9999 }, schedulePreviewRefresh);

    rbWdJP.value = true; // default
    rbWdJP.onClick = schedulePreviewRefresh;
    rbWdMTW.onClick = schedulePreviewRefresh;
    rbWdMon.onClick = schedulePreviewRefresh;

    // ===== レイアウトパネル =====
    var pnlLayout = gColL.add("panel", undefined, L("panelLayout"));
    pnlLayout.orientation = "column";
    pnlLayout.alignChildren = "left";
    pnlLayout.margins = [15, 20, 15, 10];

    // 月数/列数
    var gCount = pnlLayout.add("group");
    gCount.orientation = "row";
    gCount.alignChildren = ["left", "center"];
    gCount.add("statictext", undefined, L("months"));
    var inputMonths = gCount.add("edittext", undefined, "1");
    inputMonths.characters = 3;
    try { rbPreset1.value = true; } catch (_) { }
    gCount.add("statictext", undefined, L("cols"));
    var inputCols = gCount.add("edittext", undefined, "1");
    inputCols.characters = 3;

    // セル（panel）
    var pnlCell = pnlLayout.add("panel", undefined, L("panelCell"));
    pnlCell.orientation = "column";
    pnlCell.alignChildren = "left";
    pnlCell.margins = [15, 20, 15, 10];

    var gCellW = pnlCell.add("group");
    var stCellW = gCellW.add("statictext", undefined, L("width"));
    stCellW.justification = "right";
    stCellW.preferredSize.width = 30;
    var __fs0 = 12;
    try {
        if (typeof inputFontSize !== "undefined" && inputFontSize) {
            __fs0 = Number(inputFontSize.text);
        }
    } catch (_) { }
    if (!__fs0 || __fs0 <= 0) __fs0 = 12;
    var __defaultCellW_pt = Math.round(__fs0 * 1.5);
    var __defaultCellW = Math.round(ptToUnitValue(__defaultCellW_pt));
    var inputCellW = gCellW.add("edittext", undefined, String(__defaultCellW));
    inputCellW.characters = 5;
    gCellW.add("statictext", undefined, unitLabel);

    var gCellH = pnlCell.add("group");
    var stCellH = gCellH.add("statictext", undefined, L("height"));
    stCellH.justification = "right";
    stCellH.preferredSize.width = 30;
    var __defaultCellH_pt = Math.round(__fs0 * 1.3);
    var __defaultCellH = Math.round(ptToUnitValue(__defaultCellH_pt));
    var inputCellH = gCellH.add("edittext", undefined, String(__defaultCellH));
    inputCellH.characters = 5;
    gCellH.add("statictext", undefined, unitLabel);

    inputCellW.onChanging = schedulePreviewRefresh;
    inputCellH.onChanging = schedulePreviewRefresh;

    changeValueByArrowKey(inputCellW, { integer: false, min: 1, max: 2000 }, schedulePreviewRefresh);
    changeValueByArrowKey(inputCellH, { integer: false, min: 1, max: 2000 }, schedulePreviewRefresh);

    // 月（ユニット間マージン）panel
    var pnlMonthOuter = pnlLayout.add("panel", undefined, L("panelMonth"));
    pnlMonthOuter.orientation = "column";
    pnlMonthOuter.alignChildren = "left";
    pnlMonthOuter.margins = [15, 20, 15, 10];
    pnlMonthOuter.enabled = false; // 月数=1 のときはディム（refreshPreviewで同期）

    var gOuterH = pnlMonthOuter.add("group");
    gOuterH.orientation = "row";
    gOuterH.alignChildren = ["left", "center"];
    gOuterH.add("statictext", undefined, L("lr"));
    var inputOuterMarginX = gOuterH.add("edittext", undefined, "10");
    inputOuterMarginX.characters = 4;
    gOuterH.add("statictext", undefined, unitLabel);
    inputOuterMarginX.onChanging = schedulePreviewRefresh;

    var gOuterV = pnlMonthOuter.add("group");
    gOuterV.orientation = "row";
    gOuterV.alignChildren = ["left", "center"];
    gOuterV.add("statictext", undefined, L("ud"));
    var inputOuterMarginY = gOuterV.add("edittext", undefined, "3");
    inputOuterMarginY.characters = 4;
    gOuterV.add("statictext", undefined, unitLabel);
    inputOuterMarginY.onChanging = schedulePreviewRefresh;

    // ===== 書式パネル =====
    var pnlFormat = gColL.add("panel", undefined, L("panelFormat"));
    pnlFormat.orientation = "column";
    pnlFormat.alignChildren = "left";
    pnlFormat.margins = [15, 20, 15, 10];



    // フォント（インストール済み）選択
    var gFontName = pnlFormat.add("group");
    gFontName.add("statictext", undefined, L("font"));
    var ddFont = gFontName.add("dropdownlist", undefined, []);
    gFontName.alignChildren = ["left", "center"];
    ddFont.minimumSize = [80, 22];
    ddFont.preferredSize = [200, 22];
    ddFont.maximumSize = [200, 22];

    // お気に入り（よく使うフォント）
    var gFontFav = pnlFormat.add("group");
    gFontFav.orientation = "row";
    gFontFav.alignChildren = ["left", "center"];
    gFontFav.add("statictext", undefined, L("favorites"));
    var ddFavFont = gFontFav.add("dropdownlist", undefined, [
        "-",
        "Automate OT Light",
        "Hegante Display Regular",
        "Bodoni Moda SemiBold",
        "DIN Condensed VF Regular",
        "Caflisch Script Pro Light",
        "AgencyFB Regular"
    ]);
    ddFavFont.minimumSize = [80, 22];
    ddFavFont.preferredSize = [200, 22];
    ddFavFont.maximumSize = [200, 22];
    try { ddFavFont.selection = 0; } catch (_) { }

    // お気に入りの解決は重いので、初回にだけ辞書化してキャッシュする
    var __SCM_FAV_LABELS = [
        "Automate OT Light",
        "Hegante Display Regular",
        "Bodoni Moda SemiBold",
        "DIN Condensed VF Regular",
        "Caflisch Script Pro Light",
        "AgencyFB Regular"
    ];

    function __SCM_buildFavFontMapOnce() {
        try {
            var cache = ($.global.__SCM_CACHE = $.global.__SCM_CACHE || {});
            if (cache.favFontMap && cache.favFontMapBuilt) return cache.favFontMap;

            // lookup set
            var wantSet = {};
            for (var wi = 0; wi < __SCM_FAV_LABELS.length; wi++) {
                var k = String(__SCM_FAV_LABELS[wi] || "").replace(/\s+/g, " ").replace(/^\s+|\s+$/g, "");
                if (k) wantSet[k] = true;
            }

            var map = {};
            var len = app.textFonts.length;
            for (var i = 0; i < len; i++) {
                var f = app.textFonts[i];
                var disp = "";
                try {
                    var fam = f.family || "";
                    var sty = f.style || "";
                    disp = (fam + " " + sty).replace(/\s+/g, " ").replace(/^\s+|\s+$/g, "");
                } catch (_) { }

                if (disp && wantSet[disp]) {
                    try { map[disp] = f.name; } catch (_) { }
                }
            }

            cache.favFontMap = map;
            cache.favFontMapBuilt = true;
            return map;
        } catch (_) { }
        return {};
    }

    function __SCM_resolveFontInternalNameFromLabel(label) {
        try {
            var want = String(label || "").replace(/\s+/g, " ").replace(/^\s+|\s+$/g, "");
            if (!want) return null;

            var map = __SCM_buildFavFontMapOnce();
            if (map && map[want]) return map[want];

            // フォールバック：一致が無い場合は label をそのまま name として扱う
            return want;
        } catch (_) { }
        return null;
    }

    ddFavFont.onChange = function () {
        try {
            if (!ddFont || !ddFont.items || ddFont.items.length === 0) return;
            if (!ddFavFont.selection) return;

            var label = ddFavFont.selection.text;
            if (label === "-") return; // 無指定

            var internalName = __SCM_resolveFontInternalNameFromLabel(label);
            if (!internalName) return;

            // ddFont の items は TextFont.name で構成されている
            var __hit = false;
            for (var i = 0; i < ddFont.items.length; i++) {
                if (ddFont.items[i].text === internalName) {
                    ddFont.selection = i;
                    __hit = true;
                    break;
                }
            }

            if (__hit) {
                try {
                    if (rbWdMTW) rbWdMTW.value = true;
                    if (rbWdJP) rbWdJP.value = false;
                    if (rbWdMon) rbWdMon.value = false;
                } catch (_) { }
                schedulePreviewRefresh(true);
            }
        } catch (_) { }
    };

    // ===== フォント一覧（キャッシュ対応）=====
    // 1) targetengine 上のメモリキャッシュ（Illustrator起動中だけ）
    // 2) うまく効かない環境向けに app.putCustomOptions の簡易キャッシュも併用（必要なら自動再構築）
    var __SCM_CACHE = ($.global.__SCM_CACHE = $.global.__SCM_CACHE || {});
    var fontNames = __SCM_CACHE.fontNames;

    // app.putCustomOptions 用キー（将来変更するときはバージョンを変える）
    var __FONT_CACHE_KEY = "SmartCalendarMaker_fontNames_v1";

    function loadFontNamesFromCustomOptions() {
        try {
            var desc = app.getCustomOptions(__FONT_CACHE_KEY);
            if (!desc) return null;
            var s = desc.getString(0);
            if (!s) return null;
            var arr = s.split("\n");
            return (arr && arr.length) ? arr : null;
        } catch (_) { }
        return null;
    }

    function saveFontNamesToCustomOptions(arr) {
        try {
            var desc = new ActionDescriptor();
            // key 0 の string として保存
            desc.putString(0, arr.join("\n"));
            app.putCustomOptions(__FONT_CACHE_KEY, desc);
        } catch (_) { }
    }

    // engine キャッシュが無い場合は custom options も見る
    if (!fontNames || !fontNames.length) {
        fontNames = loadFontNamesFromCustomOptions();
        if (fontNames && fontNames.length) {
            __SCM_CACHE.fontNames = fontNames;
        }
    }

    // どちらも無ければ列挙して作成
    if (!fontNames || !fontNames.length) {
        // ===== 準備中プログレス（重い処理向け）=====
        function __createLoadingPalette(title, initialText, maxValue) {
            var w = new Window("palette", title || L("loadingTitle"));
            w.orientation = "column";
            w.alignChildren = "fill";
            w.margins = 12;

            var msg = w.add("statictext", undefined, initialText || L("loadingText"));
            var bar = w.add("progressbar", undefined, 0, (typeof maxValue === "number" && maxValue > 0) ? maxValue : 100);
            bar.preferredSize = [260, 14];

            try { w.show(); } catch (_) { }

            return {
                win: w,
                msg: msg,
                bar: bar,
                setText: function (t) { try { msg.text = t; w.update(); } catch (_) { } },
                setValue: function (v) { try { bar.value = v; w.update(); } catch (_) { } },
                setMax: function (m) { try { bar.maxvalue = m; w.update(); } catch (_) { } },
                close: function () { try { w.close(); } catch (_) { } }
            };
        }

        var __loading = __createLoadingPalette(L("loadingTitle"), L("loadingFontsRead"), 100);

        // フォント一覧を作成（名前でソート）
        fontNames = [];
        try {
            var __len = app.textFonts.length;
            for (var fi = 0; fi < __len; fi++) {
                try { fontNames.push(app.textFonts[fi].name); } catch (_) { }
                // 進捗更新（描画負荷を下げるため間引き）
                if ((fi % 50) === 0) {
                    try {
                        __loading.setMax(__len);
                        __loading.setValue(fi);
                        __loading.setText(L("loadingFontsRead") + " (" + fi + "/" + __len + ")");
                    } catch (_) { }
                }
            }
        } catch (_) { }

        try { fontNames.sort(); } catch (_) { }

        // キャッシュに保存（両方）
        __SCM_CACHE.fontNames = fontNames;
        saveFontNamesToCustomOptions(fontNames);

        // 読み込み完了
        try {
            __loading.setMax((fontNames && fontNames.length) ? fontNames.length : 100);
            __loading.setValue((fontNames && fontNames.length) ? fontNames.length : 100);
            __loading.setText(L("ready"));
        } catch (_) { }
        __loading.close();
    }

    // dropdown に投入（毎回 UI は作り直す）
    // キャッシュされていても items 追加が重い環境があるため、必要に応じてプログレスを表示
    var __ddLoading = null;
    try {
        if (fontNames && fontNames.length >= 200) {
            // 既存の列挙用ヘルパーが未定義の可能性があるため、ここでも最小版を用意
            if (typeof __createLoadingPalette !== "function") {
                var __createLoadingPalette = function (title, initialText, maxValue) {
                    var w = new Window("palette", title || L("loadingTitle"));
                    w.orientation = "column";
                    w.alignChildren = "fill";
                    w.margins = 12;
                    var msg = w.add("statictext", undefined, initialText || L("loadingText"));
                    var bar = w.add("progressbar", undefined, 0, (typeof maxValue === "number" && maxValue > 0) ? maxValue : 100);
                    bar.preferredSize = [260, 14];
                    try { w.show(); } catch (_) { }
                    return {
                        win: w,
                        msg: msg,
                        bar: bar,
                        setText: function (t) { try { msg.text = t; w.update(); } catch (_) { } },
                        setValue: function (v) { try { bar.value = v; w.update(); } catch (_) { } },
                        setMax: function (m) { try { bar.maxvalue = m; w.update(); } catch (_) { } },
                        close: function () { try { w.close(); } catch (_) { } }
                    };
                };
            }
            __ddLoading = __createLoadingPalette(L("loadingTitle"), L("loadingFontsSet"), fontNames.length);
        }
    } catch (_) { }

    for (var iFont = 0; iFont < fontNames.length; iFont++) {
        ddFont.add("item", fontNames[iFont]);
        if (__ddLoading && (iFont % 50) === 0) {
            try {
                __ddLoading.setValue(iFont);
                __ddLoading.setText(L("loadingFontsSet") + " (" + iFont + "/" + fontNames.length + ")");
            } catch (_) { }
        }
    }

    if (__ddLoading) {
        try {
            __ddLoading.setValue(fontNames.length);
            __ddLoading.setText(L("ready"));
        } catch (_) { }
        __ddLoading.close();
    }


    var gFont = pnlFormat.add("group");
    gFont.add("statictext", undefined, L("labelFontSize"));
    var inputFontSize = gFont.add("edittext", undefined, "12");
    inputFontSize.characters = 5;
    gFont.add("statictext", undefined, textUnitLabel);

    // デフォルト: DNPShueiMGoStd-B があれば選択、なければ先頭
    var defaultFontName = "DNPShueiMGoStd-B";
    var foundIndex = -1;
    for (var di = 0; di < ddFont.items.length; di++) {
        if (ddFont.items[di].text === defaultFontName) { foundIndex = di; break; }
    }
    if (ddFont.items.length > 0) {
        ddFont.selection = (foundIndex >= 0) ? foundIndex : 0;
    }

    // ===== 揃え =====
    var gAlign = pnlFormat.add("group");
    gAlign.orientation = "row";
    gAlign.alignChildren = ["left", "center"];

    gAlign.add("statictext", undefined, L("alignLabel"));

    var rbLeft = gAlign.add("radiobutton", undefined, L("left"));
    var rbCenter = gAlign.add("radiobutton", undefined, L("center"));
    var rbRight = gAlign.add("radiobutton", undefined, L("right"));
    rbCenter.value = true; // デフォルト中央
    rbLeft.onClick = schedulePreviewRefresh;
    rbCenter.onClick = schedulePreviewRefresh;
    rbRight.onClick = schedulePreviewRefresh;


    // 日曜日・祝日（横並び）
    var gHoliday = pnlFormat.add("group");
    gHoliday.orientation = "row";
    gHoliday.alignChildren = ["left", "center"];

    var chkSundayRed = gHoliday.add("checkbox", undefined, L("sundayLabel"));
    chkSundayRed.value = true;
    chkSundayRed.onClick = schedulePreviewRefresh;

    var chkHolidayRed = gHoliday.add("checkbox", undefined, L("holidayLabel"));
    chkHolidayRed.value = true;
    chkHolidayRed.onClick = schedulePreviewRefresh;


    // ===== 下部コントロール行（左：プレビュー／右：キャンセル・OK）=====
    var gBottom = dlg.add("group");
    gBottom.orientation = "row";
    gBottom.alignChildren = ["fill", "center"];
    gBottom.alignment = "fill";

    // 左側：プレビュー
    var gBottomLeft = gBottom.add("group");
    gBottomLeft.alignment = ["left", "center"];
    var previewChk = gBottomLeft.add("checkbox", undefined, L("preview"));
    previewChk.value = true;

    // スペーサー（左右を分離）
    var gSpacer = gBottom.add("group");
    gSpacer.alignment = ["fill", "fill"];

    // 右側：ボタン
    var gBottomRight = gBottom.add("group");
    gBottomRight.alignment = ["right", "center"];

    var cancelBtn = gBottomRight.add("button", undefined, L("cancel"), { name: "cancel" });
    var okBtn = gBottomRight.add("button", undefined, L("create"), { name: "ok" });

    // ===== キー操作（↑↓で数値増減）=====
    changeValueByArrowKey(inputY, { integer: true, min: 1, max: 9999 }, schedulePreviewRefresh);
    changeValueByArrowKey(inputM, { integer: true, min: 1, max: 12 }, schedulePreviewRefresh);
    changeValueByArrowKey(inputD, { integer: true, min: 1, max: 31 }, schedulePreviewRefresh);
    changeValueByArrowKey(inputMonths, { integer: true, min: 1, max: 24 }, schedulePreviewRefresh);
    changeValueByArrowKey(inputCols, { integer: true, min: 1, max: 12 }, schedulePreviewRefresh);
    changeValueByArrowKey(inputMonthBottomMargin, { integer: false, min: 0, max: 2000 }, schedulePreviewRefresh);
    changeValueByArrowKey(inputTopYearBottomMargin, { integer: false, min: 0, max: 2000 }, schedulePreviewRefresh);
    changeValueByArrowKey(inputWeekdayBottomMargin, { integer: false, min: 0, max: 2000 }, schedulePreviewRefresh);
    changeValueByArrowKey(inputFontSize, { integer: false, min: 0.1, max: 9999 }, schedulePreviewRefresh);
    changeValueByArrowKey(inputOuterMarginX, { integer: false, min: 0, max: 5000 }, schedulePreviewRefresh);
    changeValueByArrowKey(inputOuterMarginY, { integer: false, min: 0, max: 5000 }, schedulePreviewRefresh);

    // ===== プレビュー更新のデバウンス（入力中の連打を抑制）=====
    // onChanging が連続発火すると「全消去→大量生成」を連打してしまい重くなるため、
    // 最後の入力から少し待って1回だけ描画する。
    var __PREVIEW_TASK_ID = null;
    var __PREVIEW_DELAY_MS = 220; // 150〜300ms程度が体感的に良い

    // scheduleTask から呼ぶため global に置く（#targetengine で Illustrator 起動中だけ保持）
    $.global.__SCM_doRefreshPreview = function () {
        try {
            __PREVIEW_TASK_ID = null;
            refreshPreview();
        } catch (_) { }
    };

    function schedulePreviewRefresh(immediate) {
        // immediate===true のときは遅延なしで1回だけ実行
        try {
            if (__PREVIEW_TASK_ID) {
                try { app.cancelTask(__PREVIEW_TASK_ID); } catch (_) { }
                __PREVIEW_TASK_ID = null;
            }
            var delay = (immediate === true) ? 0 : __PREVIEW_DELAY_MS;
            __PREVIEW_TASK_ID = app.scheduleTask('$.global.__SCM_doRefreshPreview()', delay, false);
        } catch (_) {
            // scheduleTask が使えない/失敗した環境では従来通り即時更新
            try { refreshPreview(); } catch (__e) { }
        }
    }

    // ===== プリセット（書き出し / 読み込み） =====
    function __SCM_getSelectedRadioValue(map) {
        try {
            for (var k in map) {
                if (map.hasOwnProperty(k) && map[k] && map[k].value) return k;
            }
        } catch (_) { }
        return null;
    }

    function __SCM_setRadioByKey(map, key) {
        try {
            for (var k in map) {
                if (!map.hasOwnProperty(k) || !map[k]) continue;
                map[k].value = (k === key);
            }
        } catch (_) { }
    }

    function __SCM_getDropdownText(dd) {
        try { return (dd && dd.selection) ? (dd.selection.text || "") : ""; } catch (_) { }
        return "";
    }

    function __SCM_setDropdownByText(dd, text) {
        try {
            if (!dd || !dd.items || dd.items.length === 0) return false;
            var t = String(text || "");
            for (var i = 0; i < dd.items.length; i++) {
                if (dd.items[i].text === t) { dd.selection = i; return true; }
            }
        } catch (_) { }
        return false;
    }

    function __SCM_serializePreset() {
        return {
            // date
            y: Number(inputY.text),
            m: Number(inputM.text),
            d: Number(inputD.text),

            // layout counts
            monthCount: Number(inputMonths.text),
            colCount: Number(inputCols.text),

            // base
            startFrom: __SCM_getSelectedRadioValue({ current: rbStartCurrent, jan: rbStartJan }),

            // year
            showTopYear: !!(chkTopYear && chkTopYear.value),
            topYearBottomMargin: Number(inputTopYearBottomMargin.text),

            // month
            includeYearInMonthTitle: !!(chkMonthYear && chkMonthYear.value),
            monthTitleAlign: __SCM_getSelectedRadioValue({ left: rbMonthAlignL, center: rbMonthAlignC, right: rbMonthAlignR }),
            monthTitleMode: __SCM_getSelectedRadioValue({ num: rbMonthNum, pad: rbMonthPad, en: rbMonthEn, ens: rbMonthEnS }),
            monthBottomMargin: Number(inputMonthBottomMargin.text),
            monthTitleBottomBorder: !!(chkMonthBottomBorder && chkMonthBottomBorder.value),

            // weekday
            weekStart: __SCM_getSelectedRadioValue({ mon: rbWeekMon, sun: rbWeekSun }),
            weekdayLabelMode: __SCM_getSelectedRadioValue({ jp: rbWdJP, mtw: rbWdMTW, mon: rbWdMon }),
            weekdayBottomMargin: Number(inputWeekdayBottomMargin.text),
            weekdayFontSize: Number(inputWeekdayFontSize.text),

            // layout
            cellW: Number(inputCellW.text),
            cellH: Number(inputCellH.text),
            outerMarginX: Number(inputOuterMarginX.text),
            outerMarginY: Number(inputOuterMarginY.text),

            // format
            fontName: __SCM_getDropdownText(ddFont),
            favFont: __SCM_getDropdownText(ddFavFont),
            fontSize: Number(inputFontSize.text),

            align: __SCM_getSelectedRadioValue({ left: rbLeft, center: rbCenter, right: rbRight }),
            sundayRed: !!(chkSundayRed && chkSundayRed.value),
            holidayRed: !!(chkHolidayRed && chkHolidayRed.value)
        };
    }

    function __SCM_applyPreset(obj) {
        if (!obj) return;

        function setTextSafe(et, v) { try { if (et && v != null) et.text = String(v); } catch (_) { } }
        function setCheckSafe(chk, v) { try { if (chk && v != null) chk.value = !!v; } catch (_) { } }

        // date
        setTextSafe(inputY, obj.y);
        setTextSafe(inputM, obj.m);
        setTextSafe(inputD, obj.d);

        setTextSafe(inputMonths, obj.monthCount);
        setTextSafe(inputCols, obj.colCount);

        if (obj.startFrom) __SCM_setRadioByKey({ current: rbStartCurrent, jan: rbStartJan }, obj.startFrom);

        // year
        setCheckSafe(chkTopYear, obj.showTopYear);
        setTextSafe(inputTopYearBottomMargin, obj.topYearBottomMargin);
        try { gYearMargin.enabled = !!(chkTopYear && chkTopYear.value); } catch (_) { }

        // month
        setCheckSafe(chkMonthYear, obj.includeYearInMonthTitle);
        if (obj.monthTitleAlign) __SCM_setRadioByKey({ left: rbMonthAlignL, center: rbMonthAlignC, right: rbMonthAlignR }, obj.monthTitleAlign);
        if (obj.monthTitleMode) __SCM_setRadioByKey({ num: rbMonthNum, pad: rbMonthPad, en: rbMonthEn, ens: rbMonthEnS }, obj.monthTitleMode);
        setTextSafe(inputMonthBottomMargin, obj.monthBottomMargin);
        setCheckSafe(chkMonthBottomBorder, obj.monthTitleBottomBorder);

        // weekday
        if (obj.weekStart) __SCM_setRadioByKey({ mon: rbWeekMon, sun: rbWeekSun }, obj.weekStart);
        if (obj.weekdayLabelMode) __SCM_setRadioByKey({ jp: rbWdJP, mtw: rbWdMTW, mon: rbWdMon }, obj.weekdayLabelMode);
        setTextSafe(inputWeekdayBottomMargin, obj.weekdayBottomMargin);
        setTextSafe(inputWeekdayFontSize, obj.weekdayFontSize);

        // layout
        setTextSafe(inputCellW, obj.cellW);
        setTextSafe(inputCellH, obj.cellH);
        setTextSafe(inputOuterMarginX, obj.outerMarginX);
        setTextSafe(inputOuterMarginY, obj.outerMarginY);

        // format
        setTextSafe(inputFontSize, obj.fontSize);
        if (obj.align) __SCM_setRadioByKey({ left: rbLeft, center: rbCenter, right: rbRight }, obj.align);
        setCheckSafe(chkSundayRed, obj.sundayRed);
        setCheckSafe(chkHolidayRed, obj.holidayRed);

        // dropdowns
        try { if (ddFavFont && obj.favFont != null) __SCM_setDropdownByText(ddFavFont, obj.favFont); } catch (_) { }
        try { if (ddFont && obj.fontName != null) __SCM_setDropdownByText(ddFont, obj.fontName); } catch (_) { }

        // enable sync
        try {
            var mc = Math.round(Number(inputMonths.text));
            pnlMonthOuter.enabled = (mc !== 1);
        } catch (_) { }

        try { schedulePreviewRefresh(true); } catch (_) { }
    }

    function __SCM_toJsLiteral(obj) {
        // Human-editable preset format
        // Prefer JSON.stringify when available; otherwise use ExtendScript toSource().
        try {
            if (typeof JSON !== "undefined" && JSON && typeof JSON.stringify === "function") {
                var json = JSON.stringify(obj, null, 2);
                return "// SmartCalendarMaker preset\n// " + SCRIPT_VERSION + "\n(" + json + ")\n";
            }
        } catch (_) { }

        // JSON が無い/使えない環境向け（ExtendScript）
        try {
            if (obj && typeof obj.toSource === "function") {
                // toSource() は既に ({...}) 形式を返すことが多いので、そのまま包む
                var src = String(obj.toSource());
                // 念のため式として評価できる形に
                if (!/^\s*\(.*\)\s*$/.test(src)) src = "(" + src + ")";
                return "// SmartCalendarMaker preset\n// " + SCRIPT_VERSION + "\n" + src + "\n";
            }
        } catch (_) { }

        // 最終フォールバック：空オブジェクト
        return "// SmartCalendarMaker preset\n// " + SCRIPT_VERSION + "\n({})\n";
    }

    function __SCM_parsePresetText(text) {
        // Try JSON first, then JS literal (eval)
        var s = String(text || "");
        // strip BOM
        s = s.replace(/^\uFEFF/, "");
        try {
            if (typeof JSON !== "undefined" && JSON && typeof JSON.parse === "function") {
                return JSON.parse(s);
            }
        } catch (_) { }
        try {
            // allow files saved as: ( { ... } ) or ({...}) or just {...}
            var body = s;
            // Remove leading comments
            body = body.replace(/^\s*\/\/.*$/mg, "");
            body = body.replace(/^\s*\/\*[\s\S]*?\*\//, "");
            // Ensure object literal is evaluated correctly
            var wrapped = body;
            if (!/^\s*\(.*\)\s*$/.test(body)) {
                wrapped = "(" + body + ")";
            }
            var obj = eval(wrapped);
            return obj;
        } catch (e) {
            return null;
        }
    }

    function __SCM_savePresetToFile() {
        try {
            var obj = __SCM_serializePreset();
            var txt = __SCM_toJsLiteral(obj);

            var f = File.saveDialog(L("presetSaveDialog") + " " + L("presetSaveHint"), "*.jsxpreset;*.json");
            if (!f) return;

            var nameLower = String(f.name || "").toLowerCase();
            if (!/\.(jsxpreset|json)$/i.test(nameLower)) {
                // default extension
                try { f = new File(f.fsName + ".jsxpreset"); } catch (_) { }
                nameLower = String(f.name || "").toLowerCase();
            }

            // If user explicitly chose .json, save JSON (only if available)
            if (/\.json$/i.test(nameLower)) {
                if (typeof JSON === "undefined" || !JSON || typeof JSON.stringify !== "function") {
                    alert(L("presetErr") + "\n\n" + "JSON.stringify が利用できないため .json では保存できません。\n拡張子 .jsxpreset を選んでください。\n");
                    return;
                }
                txt = JSON.stringify(obj, null, 2);
            }

            f.encoding = "UTF-8";
            if (f.open("w")) {
                f.write(txt);
                f.close();
                alert(L("presetSaved"));
            }
        } catch (e) {
            alert(L("presetErr") + "\n\n" + e);
        }
    }

    function __SCM_loadPresetFromFile() {
        try {
            var f = File.openDialog(L("presetLoadDialog"), "*.json;*.jsxpreset");
            if (!f) return;
            f.encoding = "UTF-8";
            if (!f.open("r")) return;
            var s = f.read();
            f.close();
            var obj = __SCM_parsePresetText(s);
            if (!obj) throw new Error("Preset parse error");

            // 読み込んだプリセットを一覧に追加（ラベルはファイル名）
            try {
                var label = "";
                try {
                    label = f.displayName || f.name || "";
                } catch (_) { }
                label = String(label || "");
                label = label.replace(/\.(json|jsxpreset)$/i, "");
                if (label) __SCM_addPresetToDropdown(label, obj);
            } catch (_) { }

            __SCM_applyPreset(obj);
            alert(L("presetLoaded"));
        } catch (e) {
            alert(L("presetErr") + "\n\n" + e);
        }
    }

    // button handlers
    presetSaveBtn.onClick = function () { __SCM_savePresetToFile(); };
    presetLoadBtn.onClick = function () { __SCM_loadPresetFromFile(); };


    function refreshPreview() {
        // UI同期（プレビューOFFでも反映）: 月数=1のとき外側マージンをディム
        try {
            var __mc = Math.round(Number(inputMonths.text));
            if (!__mc || __mc < 1) __mc = 1;
            if (__mc > 24) __mc = 24;
            pnlMonthOuter.enabled = (__mc !== 1);
        } catch (_) { }

        if (!previewChk.value) return;

        // 既存プレビューを再利用（レイヤーは残し、中身だけ消す）
        clearLayerContents(doc, PREVIEW_LAYER_NAME);

        var base = parseYMDFields(inputY.text, inputM.text, inputD.text);
        if (!base) return; // 不正入力はプレビューしない
        var monthCount = Math.round(Number(inputMonths.text));
        if (!monthCount || monthCount < 1) monthCount = 1;
        if (monthCount > 24) monthCount = 24;

        var colCount = Math.round(Number(inputCols.text));
        if (!colCount || colCount < 1) colCount = 1;
        if (colCount > 12) colCount = 12;
        if (colCount > monthCount) colCount = monthCount;

        // 月数プリセット（1/3/12）と同期（列数も含めて判定）
        try {
            var isP1 = (monthCount === 1 && colCount === 1);
            var isP3 = (monthCount === 3 && colCount === 3);
            var isP12 = (monthCount === 12 && colCount === 3);

            rbPreset1.value = isP1;
            rbPreset3.value = isP3;
            rbPreset12.value = isP12;

            if (!isP1 && !isP3 && !isP12) {
                rbPreset1.value = false;
                rbPreset3.value = false;
                rbPreset12.value = false;
            }

            // 「年」panel: 年表示は 12ヶ月 のときだけON（手入力で一致した場合も含む）
            chkTopYear.value = isP12;
            chkTopYear.enabled = isP12;

            try { gYearMargin.enabled = chkTopYear.value; } catch (_) { }

            // 12ヶ月 プリセット時は「1月から」を強制
            if (isP12) {
                rbStartJan.value = true;
                rbStartCurrent.value = false;
                // 12ヶ月のときは月タイトルの「年を併記」をOFF
                try { chkMonthYear.value = false; } catch (_) { }
            }
        } catch (_) { }

        var fs = Number(inputFontSize.text);
        if (!fs || fs <= 0) return;

        var align = rbLeft.value ? "left" : (rbRight.value ? "right" : "center");
        var monthTitleAlign = rbMonthAlignL.value ? "left" : (rbMonthAlignR.value ? "right" : "center");
        var fontName = (ddFont && ddFont.selection) ? ddFont.selection.text : null;
        var sundayRed = (chkSundayRed && chkSundayRed.value) ? true : false;
        var holidayRed = (chkHolidayRed && chkHolidayRed.value) ? true : false;

        var cellW = toPtFromUI(inputCellW);
        var cellH = toPtFromUI(inputCellH);
        if (!cellW || isNaN(cellW) || cellW <= 0) return;
        if (!cellH || isNaN(cellH) || cellH <= 0) return;

        var showTopYear = (chkTopYear && chkTopYear.value) ? true : false;
        var includeYearInMonthTitle = (chkMonthYear && chkMonthYear.value) ? true : false;
        var monthTitleBottomBorder = (chkMonthBottomBorder && chkMonthBottomBorder.value) ? true : false;
        var monthTitleMode = getMonthTitleMode();
        var topYearBottomMargin = toPtFromUI(inputTopYearBottomMargin);
        if (isNaN(topYearBottomMargin) || topYearBottomMargin < 0) topYearBottomMargin = 0;
        if (topYearBottomMargin > 2000) topYearBottomMargin = 2000;

        var monthBottomMargin = toPtFromUI(inputMonthBottomMargin);
        if (isNaN(monthBottomMargin) || monthBottomMargin < 0) monthBottomMargin = 0;
        if (monthBottomMargin > 2000) monthBottomMargin = 2000;

        var weekdayBottomMargin = toPtFromUI(inputWeekdayBottomMargin);
        if (isNaN(weekdayBottomMargin) || weekdayBottomMargin < 0) weekdayBottomMargin = 0;
        if (weekdayBottomMargin > 2000) weekdayBottomMargin = 2000;

        var weekdayFontSize = Number(inputWeekdayFontSize.text);
        if (!weekdayFontSize || weekdayFontSize <= 0) weekdayFontSize = fs;

        var outerMarginX = toPtFromUI(inputOuterMarginX);
        if (isNaN(outerMarginX) || outerMarginX < 0) outerMarginX = 0;
        if (outerMarginX > 5000) outerMarginX = 5000;

        var outerMarginY = toPtFromUI(inputOuterMarginY);
        if (isNaN(outerMarginY) || outerMarginY < 0) outerMarginY = 0;
        if (outerMarginY > 5000) outerMarginY = 5000;

        var startFromJanuary = false;
        try { startFromJanuary = (rbStartJan && rbStartJan.value) ? true : false; } catch (_) { }

        var weekStartMonday = true;
        try { weekStartMonday = (rbWeekSun && rbWeekSun.value) ? false : true; } catch (_) { }

        var weekdayLabelMode = getWeekdayLabelMode();
        buildCalendar(doc, base, PREVIEW_LAYER_NAME, fs, align, monthTitleAlign, fontName, sundayRed, holidayRed, cellW, cellH, showTopYear, includeYearInMonthTitle, monthTitleBottomBorder, topYearBottomMargin, monthCount, colCount, monthTitleMode, monthBottomMargin, weekdayBottomMargin, outerMarginX, outerMarginY, startFromJanuary, weekStartMonday, weekdayLabelMode, weekdayFontSize);
        try { app.redraw(); } catch (_) { }
    }

    inputY.onChanging = schedulePreviewRefresh;
    inputM.onChanging = schedulePreviewRefresh;
    inputD.onChanging = schedulePreviewRefresh;
    inputMonths.onChanging = schedulePreviewRefresh;
    inputCols.onChanging = schedulePreviewRefresh;
    previewChk.onClick = function () { schedulePreviewRefresh(true); };
    if (ddFont) ddFont.onChange = function () { schedulePreviewRefresh(true); };

    // キーボードショートカット: L=左 / C=中央 / R=右
    dlg.addEventListener("keydown", function (event) {
        if (event.keyName === "L") {
            rbLeft.value = true;
            schedulePreviewRefresh(true);
            event.preventDefault();
        } else if (event.keyName === "C") {
            rbCenter.value = true;
            schedulePreviewRefresh(true);
            event.preventDefault();
        } else if (event.keyName === "R") {
            rbRight.value = true;
            schedulePreviewRefresh(true);
            event.preventDefault();
        }
    });

    okBtn.onClick = function () {
        var base = parseYMDFields(inputY.text, inputM.text, inputD.text);
        if (!base) {
            alert(L("errBadDate"));
            return;
        }
        var monthCount = Math.round(Number(inputMonths.text));
        if (!monthCount || monthCount < 1) monthCount = 1;
        if (monthCount > 24) monthCount = 24;

        var colCount = Math.round(Number(inputCols.text));
        if (!colCount || colCount < 1) colCount = 1;
        if (colCount > 12) colCount = 12;
        if (colCount > monthCount) colCount = monthCount;

        var fs = Number(inputFontSize.text);
        if (!fs || fs <= 0) {
            alert(L("errBadFontSize"));
            return;
        }

        // プレビューONで既に描画済みなら、それをそのまま確定（再描画しない）
        var __existingPreview = null;
        try { __existingPreview = getLayerByName(doc, PREVIEW_LAYER_NAME); } catch (_) { }
        var __finalName = "Calendar_" + base.getFullYear() + "_" + pad2(base.getMonth() + 1);
        try {
            if (previewChk && previewChk.value && __existingPreview && __existingPreview.pageItems && __existingPreview.pageItems.length > 0) {
                __existingPreview.name = __finalName;
                try { app.redraw(); } catch (_) { }
                dlg.close(1);
                return;
            }
        } catch (_) { }

        // プレビューOFFや未描画の場合は、ここで確定生成
        clearLayerContents(doc, PREVIEW_LAYER_NAME);
        var align = rbLeft.value ? "left" : (rbRight.value ? "right" : "center");
        var monthTitleAlign = rbMonthAlignL.value ? "left" : (rbMonthAlignR.value ? "right" : "center");
        var fontName = (ddFont && ddFont.selection) ? ddFont.selection.text : null;
        var sundayRed = (chkSundayRed && chkSundayRed.value) ? true : false;
        var holidayRed = (chkHolidayRed && chkHolidayRed.value) ? true : false;

        var cellW = toPtFromUI(inputCellW);
        var cellH = toPtFromUI(inputCellH);
        if (!cellW || isNaN(cellW) || cellW <= 0) return alert(L("errBadCellW"));
        if (!cellH || isNaN(cellH) || cellH <= 0) return alert(L("errBadCellH"));

        var showTopYear = (chkTopYear && chkTopYear.value) ? true : false;
        var includeYearInMonthTitle = (chkMonthYear && chkMonthYear.value) ? true : false;
        var monthTitleBottomBorder = (chkMonthBottomBorder && chkMonthBottomBorder.value) ? true : false;
        var monthTitleMode = getMonthTitleMode();
        var topYearBottomMargin = toPtFromUI(inputTopYearBottomMargin);
        if (isNaN(topYearBottomMargin) || topYearBottomMargin < 0) topYearBottomMargin = 0;
        if (topYearBottomMargin > 2000) topYearBottomMargin = 2000;

        var monthBottomMargin = toPtFromUI(inputMonthBottomMargin);
        if (isNaN(monthBottomMargin) || monthBottomMargin < 0) monthBottomMargin = 0;
        if (monthBottomMargin > 2000) monthBottomMargin = 2000;

        var weekdayBottomMargin = toPtFromUI(inputWeekdayBottomMargin);
        if (isNaN(weekdayBottomMargin) || weekdayBottomMargin < 0) weekdayBottomMargin = 0;
        if (weekdayBottomMargin > 2000) weekdayBottomMargin = 2000;

        var outerMarginX = toPtFromUI(inputOuterMarginX);
        if (isNaN(outerMarginX) || outerMarginX < 0) outerMarginX = 0;
        if (outerMarginX > 5000) outerMarginX = 5000;

        var outerMarginY = toPtFromUI(inputOuterMarginY);
        if (isNaN(outerMarginY) || outerMarginY < 0) outerMarginY = 0;
        if (outerMarginY > 5000) outerMarginY = 5000;

        var startFromJanuary = false;
        try { startFromJanuary = (rbStartJan && rbStartJan.value) ? true : false; } catch (_) { }

        var weekStartMonday = true;
        try { weekStartMonday = (rbWeekSun && rbWeekSun.value) ? false : true; } catch (_) { }

        var weekdayLabelMode = getWeekdayLabelMode();
        try {
            buildCalendar(doc, base, PREVIEW_LAYER_NAME, fs, align, monthTitleAlign, fontName, sundayRed, holidayRed, cellW, cellH, showTopYear, includeYearInMonthTitle, monthTitleBottomBorder, topYearBottomMargin, monthCount, colCount, monthTitleMode, monthBottomMargin, weekdayBottomMargin, outerMarginX, outerMarginY, startFromJanuary, weekStartMonday, weekdayLabelMode, weekdayFontSize);
        } catch (e) {
            // 生成に失敗した場合、空レイヤーだけ残るのを防ぐ
            try { removeLayerIfExists(doc, PREVIEW_LAYER_NAME); } catch (_) { }
            alert(L("errBuild") + "\n\n" + e);
            return;
        }

        // 確定：プレビューレイヤー名を変更して残す
        var lyr = getLayerByName(doc, PREVIEW_LAYER_NAME);
        if (lyr) lyr.name = __finalName;

        app.redraw();
        dlg.close(1);
    };

    cancelBtn.onClick = function () {
        removeLayerIfExists(doc, PREVIEW_LAYER_NAME);
        app.redraw();
        dlg.close(0);
    };


    dlg.show();

    // ====== functions ======
    function clampNumber(v, minV, maxV) {
        if (typeof minV === "number" && v < minV) v = minV;
        if (typeof maxV === "number" && v > maxV) v = maxV;
        return v;
    }

    /**
     * ↑↓キーで数値を増減（Shift: ±10スナップ / Option: ±0.1）
     * @param {EditText} editText
     * @param {{integer?:boolean, min?:number, max?:number, allowNegative?:boolean}} opt
     * @param {Function=} onChanged 変更後に呼ぶ（プレビュー更新など）
     */
    function changeValueByArrowKey(editText, opt, onChanged) {
        opt = opt || {};
        editText.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;

            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;

                // Shiftキー押下時は10の倍数にスナップ
                if (event.keyName === "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                } else {
                    value = Math.floor((value - 1) / delta) * delta;
                }
            } else if (keyboard.altKey) {
                delta = 0.1;

                if (event.keyName === "Up") value += delta;
                else value -= delta;
            } else {
                delta = 1;

                if (event.keyName === "Up") value += delta;
                else value -= delta;
            }

            // 丸め
            if (keyboard.altKey && opt.integer !== true) {
                value = Math.round(value * 10) / 10; // 小数第1位
            } else {
                value = Math.round(value); // 整数
            }

            // 下限
            if (opt.allowNegative !== true && value < 0) value = 0;

            // min/max
            value = clampNumber(value, opt.min, opt.max);

            event.preventDefault();
            editText.text = String(value);

            if (typeof onChanged === "function") onChanged();
        });
    }

    function buildCalendar(doc, baseDate, layerName, fontSize, alignMode, monthTitleAlign, fontName, sundayRed, holidayRed, cellW, cellH, showTopYear, includeYearInMonthTitle, monthTitleBottomBorder, topYearBottomMargin, monthCount, colCount, monthTitleMode, monthBottomMargin, weekdayBottomMargin, outerMarginX, outerMarginY, startFromJanuary, weekStartMonday, weekdayLabelMode, weekdayFontSize) {
        if (!monthTitleMode) monthTitleMode = "pad";
        monthBottomMargin = Number(monthBottomMargin);
        if (isNaN(monthBottomMargin) || monthBottomMargin < 0) monthBottomMargin = 0;
        if (monthBottomMargin > 2000) monthBottomMargin = 2000;
        weekdayBottomMargin = Number(weekdayBottomMargin);
        if (isNaN(weekdayBottomMargin) || weekdayBottomMargin < 0) weekdayBottomMargin = 0;
        if (weekdayBottomMargin > 2000) weekdayBottomMargin = 2000;
        topYearBottomMargin = Number(topYearBottomMargin);
        if (isNaN(topYearBottomMargin) || topYearBottomMargin < 0) topYearBottomMargin = 0;
        if (topYearBottomMargin > 2000) topYearBottomMargin = 2000;
        outerMarginX = Number(outerMarginX);
        if (isNaN(outerMarginX) || outerMarginX < 0) outerMarginX = 0;
        if (outerMarginX > 5000) outerMarginX = 5000;
        outerMarginY = Number(outerMarginY);
        if (isNaN(outerMarginY) || outerMarginY < 0) outerMarginY = 0;
        if (outerMarginY > 5000) outerMarginY = 5000;
        var layer = getOrCreateLayer(doc, layerName);

        layer.locked = false;
        layer.visible = true;

        var abIndex = doc.artboards.getActiveArtboardIndex();
        var abRect = doc.artboards[abIndex].artboardRect; // [L,T,R,B]
        var abCX = (abRect[0] + abRect[2]) / 2;
        var abCY = (abRect[1] + abRect[3]) / 2;

        weekStartMonday = (weekStartMonday !== false); // default true
        if (!weekdayLabelMode) weekdayLabelMode = "jp";

        var headersMon;
        if (weekdayLabelMode === "mtw") {
            headersMon = ["M", "T", "W", "T", "F", "S", "S"]; // Mon..Sun
        } else if (weekdayLabelMode === "mon") {
            headersMon = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]; // Mon..Sun
        } else {
            headersMon = ["月", "火", "水", "木", "金", "土", "日"]; // Mon..Sun
        }

        // order by week start
        var headers = weekStartMonday
            ? headersMon
            : [headersMon[6], headersMon[0], headersMon[1], headersMon[2], headersMon[3], headersMon[4], headersMon[5]];

        var sundayCol = weekStartMonday ? 6 : 0;
        // 年タイトル（最上部）
        var yearTitleH = showTopYear ? (cellH + topYearBottomMargin) : 0; // 1行 + 下マージン

        monthCount = Math.round(Number(monthCount));
        if (!monthCount || monthCount < 1) monthCount = 1;
        if (monthCount > 24) monthCount = 24;

        colCount = Math.round(Number(colCount));
        if (!colCount || colCount < 1) colCount = 1;
        if (colCount > 12) colCount = 12;
        if (colCount > monthCount) colCount = monthCount;

        // 月ごとに必要な高さ（週数）を先に計算し、全体をアートボード中心にレイアウト
        var baseYear = baseDate.getFullYear();
        var baseMonth0 = baseDate.getMonth();

        startFromJanuary = (startFromJanuary === true);

        var monthInfos = [];
        var totalW = 7 * cellW;
        var gapX = outerMarginX; // 月ユニット間の左右マージン
        var gapY = outerMarginY; // 月ユニット間の上下マージン

        var maxCellH = 0; // 1つの月ブロックの最大高さ

        var startOffset = -Math.floor((monthCount - 1) / 2);
        for (var mi = 0; mi < monthCount; mi++) {
            var d0;
            if (startFromJanuary) {
                // 1月から順に
                d0 = new Date(baseYear, 0 + mi, 1);
            } else {
                // 当月を中心に前後
                var mOff = startOffset + mi;
                d0 = new Date(baseYear, baseMonth0 + mOff, 1);
            }
            var y = d0.getFullYear();
            var m0 = d0.getMonth();
            var first = new Date(y, m0, 1);
            var last = new Date(y, m0 + 1, 0);
            var daysInMonth = last.getDate();

            // 週はじまり列（0=開始曜日..6）
            var startIndex = weekStartMonday ? ((first.getDay() + 6) % 7) : first.getDay();

            // 行数計算
            var dayCells = startIndex + daysInMonth;
            var weeks = Math.ceil(dayCells / 7);
            var rows = 1 + weeks;      // 1=曜日ヘッダ
            var totalH = (rows + 1) * cellH + monthBottomMargin + weekdayBottomMargin; // +1=タイトル行 + タイトル下マージン + 曜日下マージン

            if (totalH > maxCellH) maxCellH = totalH;

            monthInfos.push({
                year: y,
                month0: m0,
                daysInMonth: daysInMonth,
                startIndex: startIndex,
                weeks: weeks,
                rows: rows,
                totalH: totalH
            });
        }

        // グリッド配置（列数指定）
        var rowsCount = Math.ceil(monthCount / colCount);

        // 全体の幅・高さ（アートボード中心に配置）
        var blockW = colCount * totalW + (colCount - 1) * gapX;
        var blockH = rowsCount * maxCellH + (rowsCount - 1) * gapY;
        var blockHWithYear = blockH + yearTitleH;

        var startX0 = abCX - blockW / 2;
        var startY0 = abCY + blockHWithYear / 2;

        // 年タイトル（最上部）
        if (showTopYear) {
            addText(layer, String(baseYear), startX0, startY0, fontSize + 6, "center", blockW, fontName, false, null);
        }

        for (var bi = 0; bi < monthInfos.length; bi++) {
            var info = monthInfos[bi];

            // 月ごとのグループ（後から選択/移動しやすくする）
            var monthGroup = null;
            try {
                monthGroup = layer.groupItems.add();
                monthGroup.name = "Month_" + info.year + "_" + pad2(info.month0 + 1);
            } catch (_) { monthGroup = null; }

            var colI = bi % colCount;
            var rowI = Math.floor(bi / colCount);

            // 各月の左上（各セルは maxCellH の上揃え）
            var startX = startX0 + colI * (totalW + gapX);
            var startY = (startY0 - yearTitleH) - rowI * (maxCellH + gapY);

            // タイトル
            var titleText = formatMonthTitle(info.year, info.month0, includeYearInMonthTitle, monthTitleMode);
            addText(layer, titleText, startX, startY, fontSize + 2, monthTitleAlign || alignMode, totalW, fontName, false, monthGroup);

            // タイトル行の下ボーダー（タイトルと曜日の中間）
            if (monthTitleBottomBorder) {
                var by = startY - cellH - (monthBottomMargin * 0.5);
                addHLine(layer, startX, startX + totalW, by, 0.3, monthGroup);
            }

            // 曜日ヘッダ（タイトル下にマージンを入れる）
            var headerY = startY - cellH - monthBottomMargin;
            for (var c = 0; c < 7; c++) {
                var isSunH = (sundayRed && c === sundayCol);
                addText(layer, headers[c], startX + c * cellW, headerY, weekdayFontSize || fontSize, alignMode, cellW, fontName, isSunH, monthGroup);
            }

            // 日付
            var row = 1;
            var col = info.startIndex;
            for (var d = 1; d <= info.daysInMonth; d++) {
                var x = startX + col * cellW;
                var yy = headerY - weekdayBottomMargin - row * cellH;
                var isSunD = (sundayRed && col === sundayCol);
                var hname = getJPHolidayName(info.year, info.month0 + 1, d);
                var isHol = (holidayRed && hname);

                var colorFlag = isSunD;
                var tf = addText(layer, String(d), x, yy, fontSize, alignMode, cellW, fontName, colorFlag, monthGroup);

                if (isHol) {
                    try {
                        var oc = new RGBColor();
                        oc.red = 255;
                        oc.green = 140;
                        oc.blue = 0; // orange
                        tf.textRange.characterAttributes.fillColor = oc;
                    } catch (_) { }
                }

                col++;
                if (col >= 7) {
                    col = 0;
                    row++;
                }
            }
        }
    }

    function addHLine(layer, x1, x2, y, strokeW, parentGroup) {
        var p = layer.pathItems.add();
        p.setEntirePath([[x1, y], [x2, y]]);
        p.stroked = true;
        p.filled = false;
        p.strokeWidth = (typeof strokeW === "number" && strokeW > 0) ? strokeW : 0.3;

        try {
            var c = new RGBColor();
            c.red = 0; c.green = 0; c.blue = 0;
            p.strokeColor = c;
        } catch (_) { }

        // 生成後にグループへ移動
        if (parentGroup) {
            try { p.move(parentGroup, ElementPlacement.PLACEATEND); } catch (_) { }
        }
        return p;
    }

    function addText(layer, str, xLeft, y, size, alignMode, boxW, fontName, isRed, parentGroup) {
        var tf = layer.textFrames.add();
        tf.contents = str;

        var x = xLeft;
        if (alignMode === "right") x = xLeft + boxW;
        else if (alignMode === "left") x = xLeft;
        else x = xLeft + boxW / 2;

        tf.position = [x, y];
        tf.textRange.size = size;

        // フォント適用（選択されている場合）
        if (fontName) {
            try {
                tf.textRange.characterAttributes.textFont = app.textFonts.getByName(fontName);
            } catch (_) { }
        }

        // 日曜日カラー（赤）
        if (isRed) {
            try {
                var rc = new RGBColor();
                rc.red = 255; rc.green = 0; rc.blue = 0;
                tf.textRange.characterAttributes.fillColor = rc;
            } catch (_) { }
        }

        if (alignMode === "right") tf.textRange.justification = Justification.RIGHT;
        else if (alignMode === "left") tf.textRange.justification = Justification.LEFT;
        else tf.textRange.justification = Justification.CENTER;

        // 生成後にグループへ移動
        if (parentGroup) {
            try { tf.move(parentGroup, ElementPlacement.PLACEATEND); } catch (_) { }
        }
        return tf;
    }

    function toYMD(d) {
        return d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate());
    }

    function pad2(n) {
        return (n < 10) ? ("0" + n) : String(n);
    }

    function parseYMDFields(yStr, mStr, dStr) {
        var yy = Number(yStr);
        var mm = Number(mStr);
        var dd = Number(dStr);

        if (!yy || !mm || !dd) return null;
        if (mm < 1 || mm > 12) return null;
        if (dd < 1 || dd > 31) return null;

        var d = new Date(yy, mm - 1, dd);
        if (d.getFullYear() !== yy || d.getMonth() !== (mm - 1) || d.getDate() !== dd) return null;

        return d;
    }

    function getLayerByName(doc, name) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === name) return doc.layers[i];
        }
        return null;
    }

    function getOrCreateLayer(doc, name) {
        var lyr = getLayerByName(doc, name);
        if (lyr) return lyr;
        lyr = doc.layers.add();
        lyr.name = name;
        return lyr;
    }

    function clearLayerContents(doc, name) {
        var lyr = getLayerByName(doc, name);
        if (!lyr) return;

        try {
            // ロック解除
            lyr.locked = false;
            lyr.visible = true;

            // レイヤーごと削除（中身を1つずつ消すより高速な場合が多い）
            lyr.remove();
        } catch (_) { }

        try {
            // 同名レイヤーを再作成
            var newLyr = doc.layers.add();
            newLyr.name = name;
        } catch (_) { }
    }

    function removeLayerIfExists(doc, name) {
        var lyr = getLayerByName(doc, name);
        if (!lyr) return;

        // ロック解除して削除
        lyr.locked = false;
        lyr.visible = true;
        lyr.remove();
    }
})();