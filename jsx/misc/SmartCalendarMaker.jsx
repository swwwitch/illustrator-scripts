#target illustrator
#targetengine "SmartCalendarMaker"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

# SmartCalendarMaker.jsx

バージョン / Version: v1.3.1
更新日 / Updated: 2026-02-15

### 概要

- 基準日をもとに、指定した月数ぶんのカレンダー（月曜はじまり）をアートボード中心に作成します。
- セル幅／高さ・月間隔（左右／上下）などの設定はダイアログで調整でき、プレビューで即時反映されます。
- 日本の祝日（2020–2035、振替休日・国民の休日を含む）に対応しています。

### Overview

- Creates calendars (Monday start) for the given number of months, centered on the artboard, based on a reference date.
- Cell size, month spacing (L/R, U/D) and more are set in the dialog with live preview.
- Japanese national holidays (2020–2035, incl. substitute & citizen's holidays) are supported.

### 紹介記事 / Article (note)

- https://note.com/dtp_tranist/n/nc54c315c5dc3

*/


// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartCalendarMaker";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-02-15";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-07-15";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ユーザー設定 / User Settings
// =========================================
// よく調整する値をまとめています / Frequently adjusted values are collected here
var offsetX = 300;                          /* ダイアログ表示位置の横オフセット / Dialog X offset */
var offsetY = 0;                            /* ダイアログ表示位置の縦オフセット / Dialog Y offset */
var dialogOpacity = 0.985;                   /* ダイアログの不透明度 / Dialog opacity */
var __PREVIEW_DELAY_MS = 450;               /* プレビュー更新の遅延(ms) / Preview debounce delay (ms) */
var PREVIEW_LAYER_NAME = "__CAL_PREVIEW__"; /* プレビュー用レイヤー名 / Preview layer name */

// =========================================
// ローカライズ / Localization
// =========================================
/* UIロケールから言語(ja/en)を判定 / Detect UI language (ja/en) from locale */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義（カテゴリ別に構造化）/ Japanese-English label definitions (grouped by category) */
var LABELS = {
    /* ダイアログ / Dialog */
    dialog: {
        title: { ja: "カレンダー作成 " + SCRIPT_VERSION, en: "Calendar Maker " + SCRIPT_VERSION }
    },

    /* 警告メッセージ / Alerts */
    alert: {
        noDoc: {
            ja: "ドキュメントを開いてから実行してください。",
            en: "Please open a document before running."
        }
    },

    /* パネル・タブ見出し / Panel & tab titles */
    panel: {
        base: { ja: "全般", en: "General" },
        view: { ja: "画面表示", en: "View" },
        unit: { ja: "セル", en: "Cells" },
        year: { ja: "年", en: "Year" },
        month: { ja: "月", en: "Month" },
        weekday: { ja: "曜日", en: "Weekday" },
        date: { ja: "日付", en: "Date" },
        layout: { ja: "レイアウト", en: "Layout" },
        cell: { ja: "セル", en: "Cell" },
        formatTab: { ja: "文字", en: "Text" },
        font: { ja: "フォント", en: "Font" },
        option: { ja: "表記", en: "Labels" },
        fontSize: { ja: "フォントサイズ", en: "Font Size" },
        preset: { ja: "プリセット", en: "Preset" },
        notation: { ja: "表記", en: "Notation" }
    },

    /* 基準日フィールド / Base date fields */
    field: {
        year: { ja: "年:", en: "Year:" },
        month: { ja: "月:", en: "Month:" },
        day: { ja: "日:", en: "Day:" }
    },

    /* 月数プリセット・起点 / Month-count presets & base */
    base: {
        monthCount: { ja: "月数：", en: "Months:" },
        preset1: { ja: "1ヶ月", en: "1 mo" },
        preset3: { ja: "3ヶ月", en: "3 mo" },
        preset12: { ja: "12ヶ月", en: "12 mo" },
        label: { ja: "基準：", en: "Base:" },
        current: { ja: "当月基準", en: "Current month" },
        jan: { ja: "1月から", en: "From January" },
        ghost: { ja: "ゴースト", en: "Ghost" }
    },

    /* 共通UI / Common UI */
    common: {
        bottomMargin: { ja: "下マージン：", en: "Bottom margin: " },
        align: { ja: "揃え：", en: "Align:" },
        left: { ja: "左", en: "Left" },
        center: { ja: "中央", en: "Center" },
        right: { ja: "右", en: "Right" },
        zoom: { ja: "ズーム：", en: "Zoom:" },
        lr: { ja: "左右：", en: "L/R:" },
        ud: { ja: "上下：", en: "U/D:" }
    },

    /* 月タイトル / Month title */
    month: {
        includeYear: { ja: "年を併記", en: "Include year" },
        num: { ja: "数字", en: "Number" },
        pad: { ja: "数字（0埋め）", en: "Zero-padded" },
        en: { ja: "英語", en: "English" },
        ens: { ja: "英語（短縮）", en: "English (short)" },
        bottomBorder: { ja: "下ボーダー", en: "Bottom border" }
    },

    /* 年タイトル / Year title */
    year: {
        show: { ja: "年を表示", en: "Show year" }
    },

    /* 曜日ヘッダ / Weekday header */
    weekday: {
        start: { ja: "週の始まり：", en: "Week starts:" },
        monday: { ja: "月曜日", en: "Monday" },
        sunday: { ja: "日曜日", en: "Sunday" },
        notation: { ja: "表記：", en: "Notation:" },
        margin: { ja: "下マージン：", en: "Bottom margin: " },
        fontSize: { ja: "基本：", en: "Basic: " }
    },

    /* レイアウト / Layout */
    layout: {
        months: { ja: "月数:", en: "Months: " },
        cols: { ja: "列数:", en: "Columns: " },
        width: { ja: "幅：", en: "W:" },
        height: { ja: "高さ：", en: "H:" },
        cellFill: { ja: "塗り", en: "Fill" }
    },

    /* 書式（フォント・色）/ Format (font & color) */
    format: {
        font: { ja: "フォント：", en: "Font: " },
        favorites: { ja: "お気に入り：", en: "Favorites: " },
        sunday: { ja: "日曜日", en: "Sunday" },
        holiday: { ja: "祝日", en: "Holidays" }
    },

    /* ボタン / Buttons */
    button: {
        preview: { ja: "プレビュー", en: "Preview" },
        cancel: { ja: "キャンセル", en: "Cancel" },
        create: { ja: "作成", en: "OK" }
    },

    /* エラーメッセージ / Error messages */
    error: {
        badDate: {
            ja: "日付形式が正しくありません。例: 2026-02-14",
            en: "Invalid date. e.g. 2026-02-14"
        },
        badFontSize: { ja: "フォントサイズが正しくありません。", en: "Invalid font size." },
        badCellW: { ja: "セル幅が正しくありません。", en: "Invalid cell width." },
        badCellH: { ja: "セル高さが正しくありません。", en: "Invalid cell height." },
        build: {
            ja: "作成中にエラーが発生しました。",
            en: "An error occurred while creating."
        }
    },

    /* 進捗表示 / Loading progress */
    loading: {
        title: { ja: "準備中", en: "Loading" },
        text: { ja: "処理中…", en: "Working…" },
        fontsRead: { ja: "フォント一覧を読み込み中…", en: "Loading fonts…" },
        fontsSet: { ja: "フォント一覧をセット中…", en: "Applying fonts…" },
        ready: { ja: "準備完了", en: "Ready" }
    },

    /* プリセット入出力 / Preset save & load */
    preset: {
        save: { ja: "保存", en: "Save" },
        load: { ja: "読込", en: "Load" },
        saveDialog: { ja: "プリセットを保存", en: "Save preset" },
        loadDialog: { ja: "プリセットを読み込み", en: "Load preset" },
        fileFilter: { ja: "プリセットファイル", en: "Preset file" },
        saveHint: { ja: "（.jsxpreset 推奨）", en: "(recommended .jsxpreset)" },
        saved: { ja: "プリセットを保存しました。", en: "Preset saved." },
        loaded: { ja: "プリセットを読み込みました。", en: "Preset loaded." },
        err: {
            ja: "プリセットの読み込みに失敗しました。",
            en: "Failed to load preset."
        }
    }
};

/* ドット区切りキー（例 "panel.base"）でラベルを引く / Look up a label by dot-separated key (e.g. "panel.base") */
function L(key) {
    try {
        var parts = String(key).split(".");
        var o = LABELS;
        for (var i = 0; i < parts.length; i++) {
            if (!o) return key;
            o = o[parts[i]];
        }
        if (!o) return key;
        return o[lang] || o.en || o.ja || key;
    } catch (_) {
        return key;
    }
}

// =========================================
// 単位 / Units
// =========================================
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

/* 単位コードを表示ラベルに変換（Q/H対応）/ Convert a unit code to a display label (Q/H aware) */
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

/* 環境設定の単位コードからラベルを取得 / Get a unit label from a preference's unit code */
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
/* 単位コード→ポイント換算係数 / Unit code to points conversion factor */
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

/* 現在の定規単位の値をポイントへ換算 / Convert a value in the current ruler unit to points */
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

/* 入力欄の値をポイントへ換算 / Convert an edit field's value to points */
function toPtFromUI(editText) {
    return unitValueToPt(editText && editText.text);
}

/* ポイント値を現在の定規単位へ換算 / Convert a point value to the current ruler unit */
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


// =========================================
// UIレイアウトの共通設定 / Shared UI layout
// =========================================

/* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
var PANEL_SPACING  = 8;                  /* パネル内の要素間隔 / panel spacing */
var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
var TAB_MARGINS    = [15, 20, 5, 10];    /* タブ余白 [左,上,右,下] / tab margins */

/* ウィンドウの共通設定 / Apply shared window layout */
function setupWindow(win, spacing) {
    win.orientation = "column";
    win.alignChildren = "fill";
    win.margins = WINDOW_MARGINS;
    win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
}

/* パネルの共通設定 / Apply shared panel layout */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = ["fill", "top"];
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/* タブの共通設定 / Apply shared tab layout */
function setupTab(tab, spacing) {
    tab.orientation = "column";
    tab.alignChildren = "fill";
    tab.margins = TAB_MARGINS;
    if (typeof spacing === "number") tab.spacing = spacing;
}

/* 行グループの共通設定（ボタン列など）/ Apply a horizontal row group */
function setupRow(group, alignment, spacing) {
    group.orientation = "row";
    group.alignment = alignment || "left";
    group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/* ボタンの高さを指定 px 詰める（レイアウト確定後に呼ぶ）/ Trim a button's height by the given px (call after layout) */
function trimButtonHeight(button, px) {
    try {
        button.size = [button.size.width, button.size.height - px];
    } catch (e) {}
}

(function () {
    if (app.documents.length === 0) {
        alert(L("alert.noDoc"));
        return;
    }

    var doc = app.activeDocument;

    // ===== Zoom (view) helpers =====
    var __SCM_VIEW = null;
    var __SCM_ORG_ZOOM = null;
    var __SCM_ORG_CENTER = null;

    try {
        __SCM_VIEW = doc.views[0];
        __SCM_ORG_ZOOM = __SCM_VIEW.zoom;
        __SCM_ORG_CENTER = __SCM_VIEW.centerPoint;
    } catch (_) { }

    /* アクティブアートボードの中心座標を返す / Return the center point of the active artboard */
    function __SCM_getActiveArtboardCenterPoint() {
        try {
            var idx = doc.artboards.getActiveArtboardIndex();
            var r = doc.artboards[idx].artboardRect; // [L, T, R, B]
            return [r[0] + (r[2] - r[0]) / 2, r[1] + (r[3] - r[1]) / 2];
        } catch (_) { }
        try { return (__SCM_VIEW && __SCM_VIEW.centerPoint) ? __SCM_VIEW.centerPoint : [0, 0]; } catch (_) { }
        return [0, 0];
    }

    /* ズーム倍率をIllustratorの許容範囲に丸める / Clamp zoom factor to Illustrator's allowed range */
    function __SCM_clampZoomFactor(z) {
        if (z < 0.0313) z = 0.0313;
        if (z > 640.0) z = 640.0;
        return z;
    }

    /* パーセント指定でビューをズーム（パン維持）/ Zoom the view by percent (keeping pan) */
    function __SCM_applyZoomPct(pct) {
        try {
            if (!__SCM_VIEW) return;
            var p = Number(pct);
            if (isNaN(p)) return;
            if (p < 10) p = 10;
            if (p > 1600) p = 1600;
            var z = __SCM_clampZoomFactor(p / 100.0);
            __SCM_VIEW.zoom = z;
            // Keep current pan offsets when zoom changes
            __SCM_applyViewCenterWithPan();
            app.redraw();
        } catch (_) { }
    }

    /* ズームUI（スライダー/入力）を同期し必要なら適用 / Sync zoom UI and optionally apply */
    function __SCM_syncZoomUI(inputZoomPct, sldZoom, valPct, apply) {
        try {
            var v = Number(valPct);
            if (isNaN(v)) v = Number(sldZoom.value);
            v = Math.round(v);
            if (v < 10) v = 10;
            if (v > 1600) v = 1600;
            try { sldZoom.value = v; } catch (_) { }
            try { inputZoomPct.text = String(v); } catch (_) { }
            if (apply !== false) __SCM_applyZoomPct(v);
        } catch (_) { }
    }

    // ===== Pan (view center) helpers =====
    var __SCM_PAN_X = 0; // pt
    var __SCM_PAN_Y = 0; // pt (positive => down)

    /* アートボードサイズからパンの可動範囲(pt)を求める / Derive pan range (pt) from artboard size */
    function __SCM_getPanRangePt() {
        // Range is based on active artboard size.
        // Use half of width/height (pt) as max, with safety clamps.
        try {
            var idx = doc.artboards.getActiveArtboardIndex();
            var r = doc.artboards[idx].artboardRect; // [L, T, R, B]
            var w = Math.abs(r[2] - r[0]);
            var h = Math.abs(r[1] - r[3]);
            var xMax = Math.round(w / 2);
            var yMax = Math.round(h / 2);
            if (!xMax || xMax < 100) xMax = 100;
            if (!yMax || yMax < 100) yMax = 100;
            // avoid extreme ranges
            if (xMax > 50000) xMax = 50000;
            if (yMax > 50000) yMax = 50000;
            return { xMax: xMax, yMax: yMax };
        } catch (_) { }
        return { xMax: 2000, yMax: 2000 };
    }

    /* パン量を反映してビュー中心を設定 / Set the view center reflecting the current pan offsets */
    function __SCM_applyViewCenterWithPan() {
        try {
            if (!__SCM_VIEW) return;
            var c = __SCM_getActiveArtboardCenterPoint();
            var x = c[0] + Number(__SCM_PAN_X || 0);
            // Illustrator's Y axis increases upward; treat positive panY as "down" for UI.
            var y = c[1] - Number(__SCM_PAN_Y || 0);
            __SCM_VIEW.centerPoint = [x, y];
            app.redraw();
        } catch (_) { }
    }

    /* 左右パンを範囲内に丸めて適用 / Clamp and apply horizontal pan */
    function __SCM_applyPanX(v) {
        try {
            var n = Number(v);
            if (isNaN(n)) n = 0;
            var rg = __SCM_getPanRangePt();
            var mx = rg.xMax;
            if (n < -mx) n = -mx;
            if (n > mx) n = mx;
            __SCM_PAN_X = n;
            __SCM_applyViewCenterWithPan();
        } catch (_) { }
    }

    /* 上下パンを範囲内に丸めて適用 / Clamp and apply vertical pan */
    function __SCM_applyPanY(v) {
        try {
            var n = Number(v);
            if (isNaN(n)) n = 0;
            var rg = __SCM_getPanRangePt();
            var my = rg.yMax;
            if (n < -my) n = -my;
            if (n > my) n = my;
            __SCM_PAN_Y = n;
            __SCM_applyViewCenterWithPan();
        } catch (_) { }
    }

    /* ================================
     * Japan Holidays (2020-2035) built-in
     * - Includes substitute holidays (振替休日) and citizen's holidays (国民の休日)
     * - Handles 2020/2021 Olympic special holiday moves
     * ================================ */
    function __pad2(n) { return (n < 10) ? ("0" + n) : String(n); }
    function __dateKey(y, m1, d) { return y + "-" + __pad2(m1) + "-" + __pad2(d); }
    function __isLeap(y) { return (y % 4 === 0 && y % 100 !== 0) || (y % 400 === 0); }
    /* 指定年月の日数を返す / Return the number of days in a given year/month */
    function __daysInMonth(y, m1) {
        var md = [31, (__isLeap(y) ? 29 : 28), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
        return md[m1 - 1];
    }
    function __dow(y, m1, d) { return (new Date(y, m1 - 1, d)).getDay(); } // 0=Sun..6
    /* 第n・指定曜日の日を返す / Return the date of the nth given weekday */
    function __nthWeekday(y, m1, weekday0Sun, nth) {
        // weekday0Sun: 0=Sun..6
        var firstDow = __dow(y, m1, 1);
        var offset = (weekday0Sun - firstDow + 7) % 7;
        var day = 1 + offset + (nth - 1) * 7;
        return day;
    }
    /* 春分の日を近似計算(1980-2099) / Approximate the vernal equinox day (1980-2099) */
    function __vernalEquinoxDay(y) {
        // Approximation valid for 1980-2099
        return Math.floor(20.8431 + 0.242194 * (y - 1980) - Math.floor((y - 1980) / 4));
    }
    /* 秋分の日を近似計算(1980-2099) / Approximate the autumn equinox day (1980-2099) */
    function __autumnEquinoxDay(y) {
        // Approximation valid for 1980-2099
        return Math.floor(23.2488 + 0.242194 * (y - 1980) - Math.floor((y - 1980) / 4));
    }
    /* 指定年の祝日マップを生成（振替・国民の休日含む）/ Build the year's holiday map (incl. substitute & citizen's holidays) */
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
    /* 複数年ぶんの祝日マップを結合 / Merge holiday maps across multiple years */
    function __buildHolidayMap(fromY, toY) {
        var all = {};
        for (var yy = fromY; yy <= toY; yy++) {
            var m = __buildHolidaysForYear(yy);
            for (var k in m) if (m.hasOwnProperty(k)) all[k] = m[k];
        }
        return all;
    }
    var JP_HOLIDAYS_2020_2035 = __buildHolidayMap(2020, 2035);
    /* 指定日の祝日名を返す（無ければ""）/ Return the holiday name for a date ("" if none) */
    function getJPHolidayName(yyyy, m1, d) {
        return JP_HOLIDAYS_2020_2035[__dateKey(yyyy, m1, d)] || "";
    }

    // 既存プレビューが残っていたら消す
    removeLayerIfExists(doc, PREVIEW_LAYER_NAME);

    // ===== ダイアログ =====
    var today = new Date();

    var dlg = new Window("dialog", L("dialog.title"));
    /* ダイアログ表示位置をずらし初回プレビューを起動 / Offset the dialog and trigger the first preview */
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

    /* ダイアログの不透明度を設定 / Set the dialog opacity */
    function setDialogOpacity(dlg, opacityValue) {
        try { dlg.opacity = opacityValue; } catch (_) { }
    }

    setDialogOpacity(dlg, dialogOpacity);
    shiftDialogPosition(dlg, offsetX, offsetY);
    dlg.orientation = "column";
    dlg.alignChildren = "fill";

    dlg.preferredSize.width = 320;
    dlg.minimumSize.width = 320;

    // ===== タブ（プリセット / 設定） =====
    var tabs = dlg.add("tabbedpanel");
    tabs.alignment = "fill";
    tabs.alignChildren = "fill";
    // tabs.margins = [15, 20, 15, 10]; // 上に余白を追加

    var tabGeneral = tabs.add("tab", undefined, L("panel.base"));
    setupTab(tabGeneral);

    var tabCell = tabs.add("tab", undefined, L("panel.unit"));
    setupTab(tabCell);

    var tabText = tabs.add("tab", undefined, L("panel.formatTab"));
    setupTab(tabText, 10);

    var tabLabel = tabs.add("tab", undefined, L("panel.option"));
    setupTab(tabLabel);

    // 既定は「基本設定」タブ
    try { tabs.selection = tabGeneral; } catch (_) { }

    // ===== 2カラム =====
    var gCols = tabGeneral.add("group");
    gCols.orientation = "column";
    gCols.alignChildren = "fill";

    var gColL = gCols.add("group");
    gColL.orientation = "column";
    gColL.alignChildren = "fill";

    var gColR = gCols.add("group");
    gColR.orientation = "column";
    gColR.alignChildren = "fill";

    // ===== プリセット（「基本設定」タブ最下部・全幅）/ Preset (bottom of Basics tab, full width) =====
    var pnlPresetTop = tabGeneral.add("panel", undefined, L("panel.preset"));
    pnlPresetTop.orientation = "column";
    pnlPresetTop.alignChildren = ["fill", "top"];
    pnlPresetTop.alignment = "fill";
    pnlPresetTop.margins = PANEL_MARGINS;

    var gPresetBtns = pnlPresetTop.add("group");
    gPresetBtns.orientation = "row";
    gPresetBtns.alignChildren = ["left", "center"];

    var presetLoadBtn = gPresetBtns.add("button", undefined, L("preset.load"));
    var presetSaveBtn = gPresetBtns.add("button", undefined, L("preset.save"));
    // ボタンはパネル幅いっぱいに広げず左寄せ / Keep buttons at natural width, left-aligned
    presetLoadBtn.alignment = "left";
    presetSaveBtn.alignment = "left";

    // 2行目：プリセット一覧
    var gPresetListRow = pnlPresetTop.add("group");
    gPresetListRow.orientation = "row";
    gPresetListRow.alignChildren = ["left", "center"];
    gPresetListRow.alignment = "fill";

    // プリセット一覧（読み込んだプリセットを追加）
    var stPresetList = gPresetListRow.add("statictext", undefined, " ");
    var ddPresetList = gPresetListRow.add("dropdownlist", undefined, ["-"]);
    ddPresetList.preferredSize = [220, 22];
    ddPresetList.alignment = "fill";
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

    /* プリセットを保持しドロップダウンへ追加 / Store a preset and add it to the dropdown */
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

    // ===== オプションタブ内コンテナ =====
    var gOption = tabLabel.add("group");
    gOption.orientation = "column";
    gOption.alignChildren = "fill";

    // ===== 基本設定パネル =====
    var pnlBaseDate = gColL.add("panel", undefined, L("panel.base"));
    pnlBaseDate.orientation = "column";
    pnlBaseDate.alignChildren = "left";
    pnlBaseDate.margins = PANEL_MARGINS;

    var gBaseDateRow = pnlBaseDate.add("group");
    gBaseDateRow.add("statictext", undefined, L("field.year"));
    var inputY = gBaseDateRow.add("edittext", undefined, String(today.getFullYear()));
    inputY.characters = 4;

    gBaseDateRow.add("statictext", undefined, L("field.month"));
    var inputM = gBaseDateRow.add("edittext", undefined, String(today.getMonth() + 1));
    inputM.characters = 2;

    gBaseDateRow.add("statictext", undefined, L("field.day"));
    var inputD = gBaseDateRow.add("edittext", undefined, String(today.getDate()));
    inputD.characters = 2;

    // 月数プリセット（ラジオ）
    var gMonthPreset = pnlBaseDate.add("group");
    gMonthPreset.orientation = "row";
    gMonthPreset.alignChildren = ["left", "center"];

    gMonthPreset.add("statictext", undefined, L("base.monthCount"));
    var rbPreset1 = gMonthPreset.add("radiobutton", undefined, L("base.preset1"));
    var rbPreset3 = gMonthPreset.add("radiobutton", undefined, L("base.preset3"));
    var rbPreset12 = gMonthPreset.add("radiobutton", undefined, L("base.preset12"));

    /* 月数プリセット(1/3/12)に合わせUIを一括設定 / Apply a month-count preset (1/3/12) across the UI */
    function applyMonthPreset(months, cols) {
        try { inputMonths.text = String(months); } catch (_) { }
        try {
            if (typeof cols === "number" && cols > 0) inputCols.text = String(cols);
        } catch (_) { }
        // 「年」panel: 年表示は 12ヶ月 のときだけON
        try {
            chkTopYear.value = (months === 12);
            chkTopYear.enabled = (months === 12);
            __SCM_syncYearPanelDimToFontSize();
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
        // レイアウトの月数/列数は 1ヶ月 のときディム
        try { gCount.enabled = (months !== 1); } catch (_) { }
        try {
            chkGhost.enabled = (months === 1);
            if (months !== 1) chkGhost.value = false;
        } catch (_) { }
        try { schedulePreviewRefresh(true); } catch (_) { }
    }

    /* 「年」パネルの有効/無効を年表示チェックに同期 / Sync the Year panel enabled-state with the show-year checkbox */
    function __SCM_syncYearPanelDimToFontSize() {
        var en = false;
        try { en = !!(chkTopYear && chkTopYear.enabled); } catch (_) { }
        try { if (inputYearFontSize) inputYearFontSize.enabled = en; } catch (_) { }
        try { if (stFSYear) stFSYear.enabled = en; } catch (_) { }
    }

    rbPreset1.onClick = function () { applyMonthPreset(1, 1); };
    rbPreset3.onClick = function () { applyMonthPreset(3, 3); };
    rbPreset12.onClick = function () { applyMonthPreset(12, 3); };

    // 月の起点
    var gMonthStart = pnlBaseDate.add("group");
    gMonthStart.orientation = "row";
    gMonthStart.alignChildren = ["left", "center"];

    gMonthStart.add("statictext", undefined, L("base.label"));
    var rbStartCurrent = gMonthStart.add("radiobutton", undefined, L("base.current"));
    var rbStartJan = gMonthStart.add("radiobutton", undefined, L("base.jan"));

    rbStartCurrent.value = true; // デフォルト
    rbStartCurrent.onClick = schedulePreviewRefresh;
    rbStartJan.onClick = schedulePreviewRefresh;

    // ゴースト（ロジックは後で追加）
    var chkGhost = pnlBaseDate.add("checkbox", undefined, L("base.ghost"));
    chkGhost.value = false;
    chkGhost.onClick = schedulePreviewRefresh;

    chkGhost.enabled = true; // 初期値（refreshPreviewで正しく同期される）

    // ===== 日付（panel） =====
    var pnlDate = tabCell.add("panel", undefined, L("panel.date"));
    pnlDate.orientation = "column";
    pnlDate.alignChildren = "left";
    pnlDate.margins = PANEL_MARGINS;

    // ===== 揃え =====
    var gAlign = pnlDate.add("group");
    gAlign.orientation = "row";
    gAlign.alignChildren = ["left", "center"];

    gAlign.add("statictext", undefined, L("common.align"));

    var rbLeft = gAlign.add("radiobutton", undefined, L("common.left"));
    var rbCenter = gAlign.add("radiobutton", undefined, L("common.center"));
    var rbRight = gAlign.add("radiobutton", undefined, L("common.right"));
    rbCenter.value = true; // デフォルト中央
    rbLeft.onClick = schedulePreviewRefresh;
    rbCenter.onClick = schedulePreviewRefresh;
    rbRight.onClick = schedulePreviewRefresh;

    // 日曜日・祝日（横並び）
    var gHoliday = pnlDate.add("group");
    gHoliday.orientation = "row";
    gHoliday.alignChildren = ["left", "center"];

    var chkSundayRed = gHoliday.add("checkbox", undefined, L("format.sunday"));
    chkSundayRed.value = true;
    chkSundayRed.onClick = schedulePreviewRefresh;

    var chkHolidayRed = gHoliday.add("checkbox", undefined, L("format.holiday"));
    chkHolidayRed.value = true;
    chkHolidayRed.onClick = schedulePreviewRefresh;

    // ===== 曜日表記（panel） =====
    var pnlWeekdayLabel = gOption.add("panel", undefined, L("panel.weekday"));
    pnlWeekdayLabel.orientation = "column";
    pnlWeekdayLabel.alignChildren = "left";
    pnlWeekdayLabel.margins = PANEL_MARGINS;

    // 週の始まり（ラベル + ラジオ）
    var gWeekStart = pnlWeekdayLabel.add("group");
    gWeekStart.orientation = "row";
    gWeekStart.alignChildren = ["left", "center"];

    gWeekStart.add("statictext", undefined, L("weekday.start"));
    var rbWeekMon = gWeekStart.add("radiobutton", undefined, L("weekday.monday"));
    var rbWeekSun = gWeekStart.add("radiobutton", undefined, L("weekday.sunday"));

    rbWeekMon.value = true; // デフォルト
    rbWeekMon.onClick = schedulePreviewRefresh;
    rbWeekSun.onClick = schedulePreviewRefresh;

    // 曜日ヘッダの表記（ラベル + ラジオ）
    var gWeekdayLabel = pnlWeekdayLabel.add("group");
    gWeekdayLabel.orientation = "row";
    gWeekdayLabel.alignChildren = ["left", "center"];

    gWeekdayLabel.add("statictext", undefined, L("weekday.notation"));
    var rbWdJP = gWeekdayLabel.add("radiobutton", undefined, "月");
    var rbWdMTW = gWeekdayLabel.add("radiobutton", undefined, "M");
    var rbWdMon = gWeekdayLabel.add("radiobutton", undefined, "Mon");

    // 曜日ヘッダ下のマージン（pt）
    var gWeekdayBottomMargin = pnlWeekdayLabel.add("group");
    gWeekdayBottomMargin.orientation = "row";
    gWeekdayBottomMargin.alignChildren = ["left", "center"];
    gWeekdayBottomMargin.add("statictext", undefined, L("weekday.margin"));
    var inputWeekdayBottomMargin = gWeekdayBottomMargin.add("edittext", undefined, "2");
    inputWeekdayBottomMargin.characters = 4;
    gWeekdayBottomMargin.add("statictext", undefined, unitLabel);
    inputWeekdayBottomMargin.onChanging = schedulePreviewRefresh;

    // 曜日フォントサイズ（未入力時は全体フォントサイズを使用）
    var inputWeekdayFontSize;

    rbWdJP.value = true; // default
    rbWdJP.onClick = schedulePreviewRefresh;
    rbWdMTW.onClick = schedulePreviewRefresh;
    rbWdMon.onClick = schedulePreviewRefresh;


    // ===== 月パネル =====
    var pnlMonth = gOption.add("panel", undefined, L("panel.month"));
    pnlMonth.orientation = "column";
    pnlMonth.alignChildren = "left";
    pnlMonth.margins = PANEL_MARGINS;
    pnlMonth.alignment = "fill";
    pnlMonth.minimumSize.height = 44;
    pnlMonth.preferredSize.height = 44;
    pnlMonth.visible = true;
    pnlMonth.enabled = true;


    var gMonth = pnlMonth.add("group");
    gMonth.orientation = "row";
    gMonth.alignChildren = ["left", "center"];

    var chkMonthYear = gMonth.add("checkbox", undefined, L("month.includeYear"));
    chkMonthYear.value = true;
    chkMonthYear.onClick = schedulePreviewRefresh;

    // 月タイトルの揃え
    var gMonthAlign = pnlMonth.add("group");
    gMonthAlign.orientation = "row";
    gMonthAlign.alignChildren = ["left", "center"];
    gMonthAlign.add("statictext", undefined, L("common.align"));
    var rbMonthAlignL = gMonthAlign.add("radiobutton", undefined, L("common.left"));
    var rbMonthAlignC = gMonthAlign.add("radiobutton", undefined, L("common.center"));
    var rbMonthAlignR = gMonthAlign.add("radiobutton", undefined, L("common.right"));
    rbMonthAlignC.value = true; // default
    rbMonthAlignL.onClick = schedulePreviewRefresh;
    rbMonthAlignC.onClick = schedulePreviewRefresh;
    rbMonthAlignR.onClick = schedulePreviewRefresh;

    // 月タイトルの表記
    var pnlMonthFmt = pnlMonth.add("panel", undefined, L("panel.notation"));
    pnlMonthFmt.orientation = "column";
    pnlMonthFmt.alignChildren = "left";
    pnlMonthFmt.margins = PANEL_MARGINS;

    var gMonthFmt = pnlMonthFmt.add("group");
    gMonthFmt.orientation = "column";
    gMonthFmt.alignChildren = ["left", "top"];

    var rbMonthNum = gMonthFmt.add("radiobutton", undefined, L("month.num"));
    var rbMonthPad = gMonthFmt.add("radiobutton", undefined, L("month.pad"));
    var rbMonthEn = gMonthFmt.add("radiobutton", undefined, L("month.en"));
    var rbMonthEnS = gMonthFmt.add("radiobutton", undefined, L("month.ens"));

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
    gMonthMargin.add("statictext", undefined, L("common.bottomMargin"));
    var inputMonthBottomMargin = gMonthMargin.add("edittext", undefined, "3");
    inputMonthBottomMargin.characters = 4;
    gMonthMargin.add("statictext", undefined, unitLabel);
    inputMonthBottomMargin.onChanging = schedulePreviewRefresh;

    // 月タイトル行の下ボーダー
    var chkMonthBottomBorder = pnlMonth.add("checkbox", undefined, L("month.bottomBorder"));
    chkMonthBottomBorder.value = true; // デフォルトON
    chkMonthBottomBorder.onClick = schedulePreviewRefresh;

    // ===== 年パネル =====
    var pnlYear = gOption.add("panel", undefined, L("panel.year"));
    pnlYear.orientation = "column";
    pnlYear.alignChildren = "left";
    pnlYear.margins = PANEL_MARGINS;

    var gYear = pnlYear.add("group");
    gYear.orientation = "row";
    gYear.alignChildren = ["left", "center"];


    var chkTopYear = gYear.add("checkbox", undefined, L("year.show"));
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
    gYearMargin.add("statictext", undefined, L("common.bottomMargin"));
    var inputTopYearBottomMargin = gYearMargin.add("edittext", undefined, "3");
    inputTopYearBottomMargin.characters = 4;
    gYearMargin.add("statictext", undefined, unitLabel);
    inputTopYearBottomMargin.onChanging = schedulePreviewRefresh;

    try { gYearMargin.enabled = chkTopYear.value; } catch (_) { }
    /* 選択中の曜日表記モードを返す / Return the selected weekday-label mode */
    function getWeekdayLabelMode() {
        // "jp" | "mtw" | "mon"
        try {
            if (rbWdMon && rbWdMon.value) return "mon";
            if (rbWdMTW && rbWdMTW.value) return "mtw";
        } catch (_) { }
        return "jp";
    }

    /* 選択中の月タイトル表記モードを返す / Return the selected month-title mode */
    function getMonthTitleMode() {
        // "num" | "pad" | "en" | "ens"
        try {
            if (rbMonthEnS && rbMonthEnS.value) return "ens";
            if (rbMonthEn && rbMonthEn.value) return "en";
            if (rbMonthNum && rbMonthNum.value) return "num";
        } catch (_) { }
        return "pad"; // default
    }

    /* 年・月・モードから月タイトル文字列を組み立てる / Build the month-title string from year/month/mode */
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


    // ===== レイアウトパネル =====
    var pnlLayout = gColL.add("panel", undefined, L("panel.layout"));
    pnlLayout.orientation = "column";
    pnlLayout.alignChildren = "left";
    pnlLayout.margins = PANEL_MARGINS;

    // 月数/列数
    var gCount = pnlLayout.add("group");
    gCount.orientation = "row";
    gCount.alignChildren = ["left", "center"];
    gCount.add("statictext", undefined, L("layout.months"));
    var inputMonths = gCount.add("edittext", undefined, "1");
    inputMonths.characters = 3;
    try { rbPreset1.value = true; } catch (_) { }
    gCount.add("statictext", undefined, L("layout.cols"));
    var inputCols = gCount.add("edittext", undefined, "1");
    inputCols.characters = 3;

    // セル＋月（縦並びコンテナ）
    var gCellMonth = pnlLayout.add("group");
    gCellMonth.orientation = "column";
    gCellMonth.alignChildren = ["fill", "top"];
    gCellMonth.alignment = "fill";

    // 月（ユニット間マージン）panel
    var pnlMonthOuter = gCellMonth.add("panel", undefined, L("panel.month") + "（" + unitLabel + "）");
    pnlMonthOuter.orientation = "column";
    pnlMonthOuter.alignChildren = "left";
    pnlMonthOuter.margins = PANEL_MARGINS;
    pnlMonthOuter.enabled = false; // 月数=1 のときはディム（refreshPreviewで同期）

    // 横並びコンテナ
    var gOuterRow = pnlMonthOuter.add("group");
    gOuterRow.orientation = "row";
    gOuterRow.alignChildren = ["left", "center"];

    var gOuterH = gOuterRow.add("group");
    gOuterH.orientation = "row";
    gOuterH.alignChildren = ["left", "center"];
    gOuterH.add("statictext", undefined, L("common.lr"));
    var inputOuterMarginX = gOuterH.add("edittext", undefined, "10");
    inputOuterMarginX.characters = 3;
    inputOuterMarginX.onChanging = schedulePreviewRefresh;

    var gOuterV = gOuterRow.add("group");
    gOuterV.orientation = "row";
    gOuterV.alignChildren = ["left", "center"];
    gOuterV.add("statictext", undefined, L("common.ud"));
    var inputOuterMarginY = gOuterV.add("edittext", undefined, "3");
    inputOuterMarginY.characters = 3;
    inputOuterMarginY.onChanging = schedulePreviewRefresh;

    // Slider fine control: hold Option(Alt) to move at 1/10 speed
    /* Option押下時にスライダーを1/10速度で微調整 / Move a slider at 1/10 speed while Option is held */
    function __SCM_applySliderWithAltFine(slider, st, applyFn) {
        try {
            if (!slider || !st || typeof applyFn !== "function") return;

            var ks = null;
            try { ks = ScriptUI.environment.keyboardState; } catch (_) { ks = null; }
            var alt = false;
            try { alt = !!(ks && ks.altKey); } catch (_) { alt = false; }

            var raw = Number(slider.value);
            if (isNaN(raw)) raw = 0;

            // init
            if (st.raw == null || st.eff == null) {
                st.raw = raw;
                st.eff = raw;
            }

            if (alt) {
                var d = raw - Number(st.raw);
                var eff = Number(st.eff) + d * 0.1;
                st.raw = raw;
                st.eff = eff;
                try { slider.value = eff; } catch (_) { }
                applyFn(eff);
            } else {
                st.raw = raw;
                st.eff = raw;
                applyFn(raw);
            }
        } catch (_) { }
    }

    // ===== 画面表示（ズーム） =====
    var pnlView = gColL.add("panel", undefined, L("panel.view"));
    pnlView.orientation = "column";
    pnlView.alignChildren = "left";
    pnlView.margins = PANEL_MARGINS;

    var gZoom = pnlView.add("group");
    gZoom.orientation = "row";
    gZoom.alignChildren = ["left", "center"];

    var stZoom = gZoom.add("statictext", undefined, L("common.zoom"));
    try { stZoom.preferredSize.width = 58; } catch (_) { }

    var __initZoomPct = 100;
    try { if (__SCM_ORG_ZOOM != null) __initZoomPct = Math.round(Number(__SCM_ORG_ZOOM) * 100); } catch (_) { }
    if (!__initZoomPct || __initZoomPct < 10) __initZoomPct = 100;

    var sldZoom = gZoom.add("slider", undefined, __initZoomPct, 10, 1600);
    try { sldZoom.preferredSize.width = 180; } catch (_) { }

    var __stZoom = { raw: null, eff: null };
    sldZoom.onChanging = function () {
        __SCM_applySliderWithAltFine(this, __stZoom, function (v) {
            __SCM_applyZoomPct(Math.round(v));
        });
    };

    // 左右（パン）
    var gPanX = pnlView.add("group");
    gPanX.orientation = "row";
    gPanX.alignChildren = ["left", "center"];

    var stPanX = gPanX.add("statictext", undefined, L("common.lr"));
    try { stPanX.preferredSize.width = 58; } catch (_) { }

    var __panRange = __SCM_getPanRangePt();
    var sldPanX = gPanX.add("slider", undefined, 0, -__panRange.xMax, __panRange.xMax);
    try { sldPanX.preferredSize.width = 180; } catch (_) { }

    var __stPanX = { raw: null, eff: null };
    sldPanX.onChanging = function () {
        __SCM_applySliderWithAltFine(this, __stPanX, function (v) {
            __SCM_applyPanX(Math.round(v));
        });
    };

    // 上下（パン）
    var gPanY = pnlView.add("group");
    gPanY.orientation = "row";
    gPanY.alignChildren = ["left", "center"];

    var stPanY = gPanY.add("statictext", undefined, L("common.ud"));
    try { stPanY.preferredSize.width = 58; } catch (_) { }

    var sldPanY = gPanY.add("slider", undefined, 0, -__panRange.yMax, __panRange.yMax);
    try { sldPanY.preferredSize.width = 180; } catch (_) { }

    var __stPanY = { raw: null, eff: null };
    sldPanY.onChanging = function () {
        __SCM_applySliderWithAltFine(this, __stPanY, function (v) {
            __SCM_applyPanY(Math.round(v));
        });
    };

    // ===== セル（panel）をレイアウトから移動してフォントパネルの直前に配置 =====
    // セル（panel）
    var pnlCell = tabCell.add("panel", undefined, L("panel.cell"));
    pnlCell.orientation = "column";
    pnlCell.alignChildren = "left";
    pnlCell.margins = PANEL_MARGINS;

    // セル上段コンテナ（大きさ／セル間隔 横並び）
    var gCellTopRow = pnlCell.add("group");
    gCellTopRow.orientation = "row";
    gCellTopRow.alignChildren = ["fill", "top"];
    gCellTopRow.alignment = "fill";

    // セル：大きさ（sub panel）
    var pnlCellSize = gCellTopRow.add("panel", undefined, "大きさ");
    pnlCellSize.orientation = "column";
    pnlCellSize.alignChildren = "left";
    pnlCellSize.margins = PANEL_MARGINS;

    var gCellW = pnlCellSize.add("group");
    var stCellW = gCellW.add("statictext", undefined, L("layout.width"));
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
    inputCellW.characters = 3;
    gCellW.add("statictext", undefined, unitLabel);

    var gCellH = pnlCellSize.add("group");
    var stCellH = gCellH.add("statictext", undefined, L("layout.height"));
    stCellH.justification = "right";
    stCellH.preferredSize.width = 30;
    var __defaultCellH_pt = Math.round(__fs0 * 1.3);
    var __defaultCellH = Math.round(ptToUnitValue(__defaultCellH_pt));
    var inputCellH = gCellH.add("edittext", undefined, String(__defaultCellH));
    inputCellH.characters = 3;
    gCellH.add("statictext", undefined, unitLabel);

    // セル：間隔（sub panel）
    var pnlCellGap = gCellTopRow.add("panel", undefined, "セル間隔（" + unitLabel + "）");
    pnlCellGap.orientation = "column";
    pnlCellGap.alignChildren = "left";
    pnlCellGap.margins = PANEL_MARGINS;

    // 縦並び（左右 / 上下）
    var gCellGapRow = pnlCellGap.add("group");
    gCellGapRow.orientation = "column";
    gCellGapRow.alignChildren = ["left", "center"];
    gCellGapRow.spacing = 6;

    var gCellGapX = gCellGapRow.add("group");
    gCellGapX.orientation = "row";
    gCellGapX.alignChildren = ["left", "center"];
    gCellGapX.add("statictext", undefined, L("common.lr"));
    var inputCellGapX = gCellGapX.add("edittext", undefined, "0");
    inputCellGapX.characters = 3;

    var gCellGapY = gCellGapRow.add("group");
    gCellGapY.orientation = "row";
    gCellGapY.alignChildren = ["left", "center"];
    gCellGapY.add("statictext", undefined, L("common.ud"));
    var inputCellGapY = gCellGapY.add("edittext", undefined, "0");
    inputCellGapY.characters = 3;

    inputCellGapX.onChanging = schedulePreviewRefresh;
    inputCellGapY.onChanging = schedulePreviewRefresh;

    changeValueByArrowKey(inputCellGapX, { integer: false, min: 0, max: 2000 }, schedulePreviewRefresh);
    changeValueByArrowKey(inputCellGapY, { integer: false, min: 0, max: 2000 }, schedulePreviewRefresh);

    // セル位置調整（上下）
    var gCellPosAdj = pnlCell.add("group");
    gCellPosAdj.orientation = "row";
    gCellPosAdj.alignChildren = ["left", "center"];
    gCellPosAdj.add("statictext", undefined, "セル位置調整：");
    var inputCellPosAdjY = gCellPosAdj.add("edittext", undefined, "0");
    inputCellPosAdjY.characters = 4;
    gCellPosAdj.add("statictext", undefined, unitLabel);

    inputCellPosAdjY.onChanging = schedulePreviewRefresh;
    changeValueByArrowKey(inputCellPosAdjY, { integer: false, min: -2000, max: 2000, allowNegative: true }, schedulePreviewRefresh);

    // セル背景の塗り
    var chkCellFill = pnlCell.add("checkbox", undefined, L("layout.cellFill"));
    chkCellFill.value = false; // 既定：OFF
    chkCellFill.onClick = schedulePreviewRefresh;

    // 罫線（panel）
    var pnlCellStroke = pnlCell.add("panel", undefined, "罫線");
    pnlCellStroke.orientation = "column";
    pnlCellStroke.alignChildren = "left";
    pnlCellStroke.margins = PANEL_MARGINS;

    // 罫線モード
    var gStrokeMode = pnlCellStroke.add("group");
    gStrokeMode.orientation = "row";
    gStrokeMode.alignChildren = ["left", "center"];

    var rbStrokeNone = gStrokeMode.add("radiobutton", undefined, "なし");
    var rbStrokeAll = gStrokeMode.add("radiobutton", undefined, "すべて");
    var rbStrokeCustom = gStrokeMode.add("radiobutton", undefined, "個別");

    rbStrokeNone.value = true; // default

    var gStrokeAll = pnlCellStroke.add("group");
    gStrokeAll.orientation = "row";
    gStrokeAll.alignChildren = ["left", "center"];

    var chkStrokeTop = gStrokeAll.add("checkbox", undefined, "上");
    var chkStrokeBottom = gStrokeAll.add("checkbox", undefined, "下");
    var chkStrokeLeft = gStrokeAll.add("checkbox", undefined, "左");
    var chkStrokeRight = gStrokeAll.add("checkbox", undefined, "右");

    // デフォルトはすべてON
    chkStrokeTop.value = true;
    chkStrokeBottom.value = true;
    chkStrokeLeft.value = true;
    chkStrokeRight.value = true;

    chkStrokeTop.onClick = schedulePreviewRefresh;
    chkStrokeBottom.onClick = schedulePreviewRefresh;
    chkStrokeLeft.onClick = schedulePreviewRefresh;
    chkStrokeRight.onClick = schedulePreviewRefresh;

    // 罫線モード切替ロジック
    /* 罫線モード（なし/すべて/個別）でチェック状態を切替 / Toggle border checkboxes by stroke mode */
    function updateStrokeMode() {
        if (rbStrokeNone.value) {
            chkStrokeTop.value = false;
            chkStrokeBottom.value = false;
            chkStrokeLeft.value = false;
            chkStrokeRight.value = false;

            chkStrokeTop.enabled = false;
            chkStrokeBottom.enabled = false;
            chkStrokeLeft.enabled = false;
            chkStrokeRight.enabled = false;
        }
        else if (rbStrokeAll.value) {
            chkStrokeTop.value = true;
            chkStrokeBottom.value = true;
            chkStrokeLeft.value = true;
            chkStrokeRight.value = true;

            chkStrokeTop.enabled = false;
            chkStrokeBottom.enabled = false;
            chkStrokeLeft.enabled = false;
            chkStrokeRight.enabled = false;
        }
        else {
            chkStrokeTop.enabled = true;
            chkStrokeBottom.enabled = true;
            chkStrokeLeft.enabled = true;
            chkStrokeRight.enabled = true;
        }
        schedulePreviewRefresh(true);
    }

    rbStrokeNone.onClick = updateStrokeMode;
    rbStrokeAll.onClick = updateStrokeMode;
    rbStrokeCustom.onClick = updateStrokeMode;

    // 初期状態に反映
    updateStrokeMode();

    inputCellW.onChanging = schedulePreviewRefresh;
    inputCellH.onChanging = schedulePreviewRefresh;

    changeValueByArrowKey(inputCellW, { integer: false, min: 1, max: 2000 }, schedulePreviewRefresh);
    changeValueByArrowKey(inputCellH, { integer: false, min: 1, max: 2000 }, schedulePreviewRefresh);

    // ===== 書式パネル =====
    var pnlFormat = tabText.add("panel", undefined, L("panel.font"));
    pnlFormat.orientation = "column";
    pnlFormat.alignChildren = "left";
    pnlFormat.margins = PANEL_MARGINS;

    // フォントサイズグリッド用変数宣言
    var inputFontSize;
    var inputMonthFontSize;
    var inputYearFontSize;
    var stFSYear;

    // フォントサイズ（2行4列）
    var pnlFontSize = tabText.add("panel", undefined, L("panel.fontSize") + "（" + textUnitLabel + "）");
    pnlFontSize.orientation = "column";
    pnlFontSize.alignChildren = "left";
    pnlFontSize.margins = PANEL_MARGINS;

    // フォントサイズ（2行4列）
    var gFSGrid = pnlFontSize.add("group");
    gFSGrid.orientation = "column";
    gFSGrid.alignChildren = ["fill", "center"];

    var gFSLabels = gFSGrid.add("group");
    gFSLabels.orientation = "row";
    gFSLabels.alignChildren = ["center", "center"];
    gFSLabels.alignment = ["center", "top"];

    var stFSBasic = gFSLabels.add("statictext", undefined, "基本");
    var stFSWeek = gFSLabels.add("statictext", undefined, "曜日");
    var stFSMonth = gFSLabels.add("statictext", undefined, "月");
    stFSYear = gFSLabels.add("statictext", undefined, "年");

    // ラベル幅を揃えて中央寄せ
    var __fsLabelW = 42;
    stFSBasic.preferredSize.width = __fsLabelW; stFSBasic.justification = "center";
    stFSWeek.preferredSize.width = __fsLabelW; stFSWeek.justification = "center";
    stFSMonth.preferredSize.width = __fsLabelW; stFSMonth.justification = "center";
    stFSYear.preferredSize.width = __fsLabelW; stFSYear.justification = "center";

    var gFSInputs = gFSGrid.add("group");
    gFSInputs.orientation = "row";
    gFSInputs.alignChildren = ["left", "center"];
    gFSInputs.alignment = ["left", "top"];

    inputFontSize = gFSInputs.add("edittext", undefined, "12");
    inputFontSize.characters = 3;
    inputFontSize.onChanging = schedulePreviewRefresh;

    inputWeekdayFontSize = gFSInputs.add("edittext", undefined, "12");
    inputWeekdayFontSize.characters = 3;
    inputWeekdayFontSize.onChanging = schedulePreviewRefresh;
    changeValueByArrowKey(inputWeekdayFontSize, { integer: false, min: 0.1, max: 9999 }, schedulePreviewRefresh);

    inputMonthFontSize = gFSInputs.add("edittext", undefined, "12");
    inputMonthFontSize.characters = 3;
    inputMonthFontSize.onChanging = schedulePreviewRefresh;
    changeValueByArrowKey(inputMonthFontSize, { integer: false, min: 0.1, max: 9999 }, schedulePreviewRefresh);

    inputYearFontSize = gFSInputs.add("edittext", undefined, "12");
    inputYearFontSize.characters = 3;
    inputYearFontSize.onChanging = schedulePreviewRefresh;
    changeValueByArrowKey(inputYearFontSize, { integer: false, min: 0.1, max: 9999 }, schedulePreviewRefresh);

    try { __SCM_syncYearPanelDimToFontSize(); } catch (_) { }

    // フォント（インストール済み）選択
    var gFontName = pnlFormat.add("group");
    var stFontLbl = gFontName.add("statictext", undefined, L("format.font"));
    var ddFont = gFontName.add("dropdownlist", undefined, []);
    gFontName.alignChildren = ["left", "center"];
    ddFont.minimumSize = [80, 22];
    ddFont.preferredSize = [180, 22];
    ddFont.maximumSize = [180, 22];

    // お気に入り（よく使うフォント）
    var gFontFav = pnlFormat.add("group");
    gFontFav.orientation = "row";
    gFontFav.alignChildren = ["left", "left"];
    var stFavLbl = gFontFav.add("statictext", undefined, L("format.favorites"));
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
    ddFavFont.preferredSize = [180, 22];
    ddFavFont.maximumSize = [180, 22];
    try { ddFavFont.selection = 0; } catch (_) { }

    // フォント／お気に入りのラベル幅を一定に揃える（プルダウン左端を合わせる）/ Fix font & favorites label widths so the dropdowns align
    var __FONT_LABEL_W = 80;
    stFontLbl.preferredSize.width = __FONT_LABEL_W;
    stFavLbl.preferredSize.width = __FONT_LABEL_W;

    // お気に入りの解決は重いので、初回にだけ辞書化してキャッシュする
    var __SCM_FAV_LABELS = [
        "Automate OT Light",
        "Hegante Display Regular",
        "Bodoni Moda SemiBold",
        "DIN Condensed VF Regular",
        "Caflisch Script Pro Light",
        "AgencyFB Regular"
    ];

    /* お気に入りフォント名→内部名の対応表を一度だけ構築 / Build the favorite-label→internal-name map once */
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

    /* お気に入り表示名から内部フォント名を解決 / Resolve an internal font name from a favorite label */
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

    /* カスタムオプションからフォント一覧キャッシュを読む / Load the cached font list from custom options */
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

    /* フォント一覧をカスタムオプションへ保存 / Save the font list to custom options */
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
            var w = new Window("palette", title || L("loading.title"));
            w.orientation = "column";
            w.alignChildren = "fill";
            w.margins = 12;

            var msg = w.add("statictext", undefined, initialText || L("loading.text"));
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

        var __loading = __createLoadingPalette(L("loading.title"), L("loading.fontsRead"), 100);

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
                        __loading.setText(L("loading.fontsRead") + " (" + fi + "/" + __len + ")");
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
            __loading.setText(L("loading.ready"));
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
                    var w = new Window("palette", title || L("loading.title"));
                    w.orientation = "column";
                    w.alignChildren = "fill";
                    w.margins = 12;
                    var msg = w.add("statictext", undefined, initialText || L("loading.text"));
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
            __ddLoading = __createLoadingPalette(L("loading.title"), L("loading.fontsSet"), fontNames.length);
        }
    } catch (_) { }

    for (var iFont = 0; iFont < fontNames.length; iFont++) {
        ddFont.add("item", fontNames[iFont]);
        if (__ddLoading && (iFont % 50) === 0) {
            try {
                __ddLoading.setValue(iFont);
                __ddLoading.setText(L("loading.fontsSet") + " (" + iFont + "/" + fontNames.length + ")");
            } catch (_) { }
        }
    }

    if (__ddLoading) {
        try {
            __ddLoading.setValue(fontNames.length);
            __ddLoading.setText(L("loading.ready"));
        } catch (_) { }
        __ddLoading.close();
    }



    // デフォルト: DNPShueiMGoStd-B があれば選択、なければ先頭
    var defaultFontName = "DNPShueiMGoStd-B";
    var foundIndex = -1;
    for (var di = 0; di < ddFont.items.length; di++) {
        if (ddFont.items[di].text === defaultFontName) { foundIndex = di; break; }
    }
    if (ddFont.items.length > 0) {
        ddFont.selection = (foundIndex >= 0) ? foundIndex : 0;
    }



    // ===== 下部コントロール行（左：プレビュー／右：キャンセル・OK）=====
    var gBottom = dlg.add("group");
    gBottom.orientation = "row";
    gBottom.alignChildren = ["fill", "center"];
    gBottom.alignment = "fill";

    // 左側：プレビュー
    var gBottomLeft = gBottom.add("group");
    gBottomLeft.alignment = ["left", "center"];
    var previewChk = gBottomLeft.add("checkbox", undefined, L("button.preview"));
    previewChk.value = true;

    // スペーサー（左右を分離）
    var gSpacer = gBottom.add("group");
    gSpacer.alignment = ["fill", "fill"];

    // 右側：ボタン
    var gBottomRight = gBottom.add("group");
    gBottomRight.alignment = ["right", "center"];

    var cancelBtn = gBottomRight.add("button", undefined, L("button.cancel"), { name: "cancel" });
    var okBtn = gBottomRight.add("button", undefined, L("button.create"), { name: "ok" });

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

    // scheduleTask から呼ぶため global に置く（#targetengine で Illustrator 起動中だけ保持）
    $.global.__SCM_doRefreshPreview = function () {
        try {
            __PREVIEW_TASK_ID = null;
            refreshPreview();
        } catch (_) { }
    };

    /* プレビュー更新をデバウンスして予約 / Schedule a debounced preview refresh */
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
    /* ラジオ群から選択中のキーを返す / Return the selected key from a radio group */
    function __SCM_getSelectedRadioValue(map) {
        try {
            for (var k in map) {
                if (map.hasOwnProperty(k) && map[k] && map[k].value) return k;
            }
        } catch (_) { }
        return null;
    }

    /* キー指定でラジオ群を選択 / Select a radio in a group by key */
    function __SCM_setRadioByKey(map, key) {
        try {
            for (var k in map) {
                if (!map.hasOwnProperty(k) || !map[k]) continue;
                map[k].value = (k === key);
            }
        } catch (_) { }
    }

    /* ドロップダウンの選択テキストを返す / Return the selected text of a dropdown */
    function __SCM_getDropdownText(dd) {
        try { return (dd && dd.selection) ? (dd.selection.text || "") : ""; } catch (_) { }
        return "";
    }

    /* テキスト一致でドロップダウンを選択 / Select a dropdown item by matching text */
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

    /* 現在のUI状態をプリセットオブジェクトへ直列化 / Serialize the current UI state into a preset object */
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
            monthFontSize: Number(inputMonthFontSize.text),
            yearFontSize: Number(inputYearFontSize.text),

            // layout
            cellW: Number(inputCellW.text),
            cellH: Number(inputCellH.text),
            cellGapX: Number(inputCellGapX.text),
            cellGapY: Number(inputCellGapY.text),
            cellPosAdjY: Number(inputCellPosAdjY.text),
            outerMarginX: Number(inputOuterMarginX.text),
            outerMarginY: Number(inputOuterMarginY.text),

            cellFill: !!(chkCellFill && chkCellFill.value),

            // format
            fontName: __SCM_getDropdownText(ddFont),
            favFont: __SCM_getDropdownText(ddFavFont),
            fontSize: Number(inputFontSize.text),

            align: __SCM_getSelectedRadioValue({ left: rbLeft, center: rbCenter, right: rbRight }),
            sundayRed: !!(chkSundayRed && chkSundayRed.value),
            holidayRed: !!(chkHolidayRed && chkHolidayRed.value)
        };
    }

    /* プリセットオブジェクトをUIへ反映 / Apply a preset object to the UI */
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
        setTextSafe(inputMonthFontSize, obj.monthFontSize);
        setTextSafe(inputYearFontSize, obj.yearFontSize);

        // layout
        setTextSafe(inputCellW, obj.cellW);
        setTextSafe(inputCellH, obj.cellH);
        setTextSafe(inputCellGapX, (obj.cellGapX != null) ? obj.cellGapX : 0);
        setTextSafe(inputCellGapY, (obj.cellGapY != null) ? obj.cellGapY : 0);
        setTextSafe(inputOuterMarginX, obj.outerMarginX);
        setTextSafe(inputOuterMarginY, obj.outerMarginY);

        setTextSafe(inputCellPosAdjY, (obj.cellPosAdjY != null) ? obj.cellPosAdjY : 0);
        setCheckSafe(chkCellFill, (obj.cellFill != null) ? obj.cellFill : true);

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

    /* プリセットを保存用テキスト(JS/JSON)へ変換 / Convert a preset to savable text (JS/JSON) */
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

    /* 保存テキストからプリセットを復元 / Parse a preset from saved text */
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

    /* プリセットをファイルへ書き出す / Save a preset to a file */
    function __SCM_savePresetToFile() {
        try {
            var obj = __SCM_serializePreset();
            var txt = __SCM_toJsLiteral(obj);

            var f = File.saveDialog(L("preset.saveDialog") + " " + L("preset.saveHint"), "*.jsxpreset;*.json");
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
                    alert(L("preset.err") + "\n\n" + "JSON.stringify が利用できないため .json では保存できません。\n拡張子 .jsxpreset を選んでください。\n");
                    return;
                }
                txt = JSON.stringify(obj, null, 2);
            }

            f.encoding = "UTF-8";
            if (f.open("w")) {
                f.write(txt);
                f.close();
                alert(L("preset.saved"));
            }
        } catch (e) {
            alert(L("preset.err") + "\n\n" + e);
        }
    }

    /* プリセットをファイルから読み込む / Load a preset from a file */
    function __SCM_loadPresetFromFile() {
        try {
            var f = File.openDialog(L("preset.loadDialog"), "*.json;*.jsxpreset");
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
            alert(L("preset.loaded"));
        } catch (e) {
            alert(L("preset.err") + "\n\n" + e);
        }
    }

    // button handlers
    presetSaveBtn.onClick = function () { __SCM_savePresetToFile(); };
    presetLoadBtn.onClick = function () { __SCM_loadPresetFromFile(); };

    /* 入力値からプレビュー用カレンダーを再生成 / Rebuild the preview calendar from current inputs */
    function refreshPreview() {
        // UI同期（プレビューOFFでも反映）: 月数=1のとき外側マージンをディム
        try {
            var __mc = Math.round(Number(inputMonths.text));
            if (!__mc || __mc < 1) __mc = 1;
            if (__mc > 24) __mc = 24;
            pnlMonthOuter.enabled = (__mc !== 1);
            // レイアウトの月数/列数は 1ヶ月 のときディム
            try { gCount.enabled = (__mc !== 1); } catch (_) { }
            // ゴーストは 1ヶ月 のときだけ有効
            try {
                chkGhost.enabled = (__mc === 1);
                if (__mc !== 1) chkGhost.value = false;
            } catch (_) { }
        } catch (_) { }

        if (!previewChk.value) return;

        // プレビューレイヤーを全削除して作り直す
        removeLayerIfExists(doc, PREVIEW_LAYER_NAME);

        var baseDate = parseYMDFields(inputY.text, inputM.text, inputD.text);
        if (!baseDate) return; // 不正入力はプレビューしない
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
            __SCM_syncYearPanelDimToFontSize();

            try { gYearMargin.enabled = chkTopYear.value; } catch (_) { }

            // 12ヶ月 プリセット時は「1月から」を強制
            if (isP12) {
                rbStartJan.value = true;
                rbStartCurrent.value = false;
                // 12ヶ月のときは月タイトルの「年を併記」をOFF
                try { chkMonthYear.value = false; } catch (_) { }
            }
        } catch (_) { }

        var fontSize = Number(inputFontSize.text);
        if (!fontSize || fontSize <= 0) return;

        var align = rbLeft.value ? "left" : (rbRight.value ? "right" : "center");
        var monthTitleAlign = rbMonthAlignL.value ? "left" : (rbMonthAlignR.value ? "right" : "center");
        var fontName = (ddFont && ddFont.selection) ? ddFont.selection.text : null;
        var sundayRed = (chkSundayRed && chkSundayRed.value) ? true : false;
        var holidayRed = (chkHolidayRed && chkHolidayRed.value) ? true : false;

        var ghost = (chkGhost && chkGhost.enabled && chkGhost.value) ? true : false;

        var cellFill = !!(chkCellFill && chkCellFill.value);

        // 罫線設定
        var strokeMode = (rbStrokeNone && rbStrokeNone.value) ? "none"
            : ((rbStrokeCustom && rbStrokeCustom.value) ? "custom" : "all");

        var strokeTop = !!(chkStrokeTop && chkStrokeTop.value);
        var strokeBottom = !!(chkStrokeBottom && chkStrokeBottom.value);
        var strokeLeft = !!(chkStrokeLeft && chkStrokeLeft.value);
        var strokeRight = !!(chkStrokeRight && chkStrokeRight.value);

        if (strokeMode === "all") {
            strokeTop = strokeBottom = strokeLeft = strokeRight = true;
        }

        var cellW = toPtFromUI(inputCellW);
        var cellH = toPtFromUI(inputCellH);
        if (!cellW || isNaN(cellW) || cellW <= 0) return;
        if (!cellH || isNaN(cellH) || cellH <= 0) return;

        var cellGapX = toPtFromUI(inputCellGapX);
        var cellGapY = toPtFromUI(inputCellGapY);
        if (isNaN(cellGapX) || cellGapX < 0) cellGapX = 0;
        if (cellGapX > 2000) cellGapX = 2000;
        if (isNaN(cellGapY) || cellGapY < 0) cellGapY = 0;
        if (cellGapY > 2000) cellGapY = 2000;

        var cellPosAdjY = toPtFromUI(inputCellPosAdjY);
        if (isNaN(cellPosAdjY)) cellPosAdjY = 0;
        if (cellPosAdjY < -2000) cellPosAdjY = -2000;
        if (cellPosAdjY > 2000) cellPosAdjY = 2000;

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
        if (!weekdayFontSize || weekdayFontSize <= 0) weekdayFontSize = fontSize;

        var monthFontSize = Number(inputMonthFontSize.text);
        if (!monthFontSize || monthFontSize <= 0) monthFontSize = fontSize;

        var yearFontSize = Number(inputYearFontSize.text);
        if (!yearFontSize || yearFontSize <= 0) yearFontSize = fontSize;

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
            var options = {
                fontSize: fontSize,
                align: align,
                monthTitleAlign: monthTitleAlign,
                fontName: fontName,
                sundayRed: sundayRed,
                holidayRed: holidayRed,
                ghost: ghost,
                cell: {
                    w: cellW,
                    h: cellH,
                    gapX: cellGapX,
                    gapY: cellGapY,
                    bgOffsetY: cellPosAdjY,
                    fill: cellFill,
                    stroke: {
                        mode: strokeMode,
                        top: strokeTop,
                        bottom: strokeBottom,
                        left: strokeLeft,
                        right: strokeRight
                    }
                },
                year: {
                    showTop: showTopYear,
                    bottomMargin: topYearBottomMargin,
                    fontSize: yearFontSize
                },
                month: {
                    includeYear: includeYearInMonthTitle,
                    titleBottomBorder: monthTitleBottomBorder,
                    titleMode: monthTitleMode,
                    bottomMargin: monthBottomMargin,
                    fontSize: monthFontSize
                },
                weekday: {
                    startMonday: weekStartMonday,
                    labelMode: weekdayLabelMode,
                    bottomMargin: weekdayBottomMargin,
                    fontSize: weekdayFontSize
                },
                unit: {
                    outerX: outerMarginX,
                    outerY: outerMarginY
                },
                layout: {
                    monthCount: monthCount,
                    colCount: colCount,
                    startFromJanuary: startFromJanuary
                }
            };
            buildCalendarOpt(doc, baseDate, PREVIEW_LAYER_NAME, options);
        } catch (e) {
            try { removeLayerIfExists(doc, PREVIEW_LAYER_NAME); } catch (_) { }
            alert("Preview error:\n\n" + e);
            return;
        }
        try { app.redraw(); } catch (_) { }
    }

    // ===== buildCalendar: options object wrapper =====
    // 引数ズレを防ぐため、呼び出し側は options だけを組み立てて渡す。
    /* optを正規化してbuildCalendarへ委譲 / Normalize the options object and delegate to buildCalendar */
    function buildCalendarOpt(doc, baseDate, layerName, options) {
        options = options || {};
        options.cell = options.cell || {};
        options.cell.stroke = options.cell.stroke || {};
        options.month = options.month || {};
        options.year = options.year || {};
        options.weekday = options.weekday || {};
        options.unit = options.unit || {};
        options.layout = options.layout || {};

        // normalize
        var fontSize = Number(options.fontSize); if (!fontSize || fontSize <= 0) fontSize = 12;
        var align = options.align || "center";
        var monthTitleAlign = options.monthTitleAlign || align;
        var fontName = options.fontName || null;
        var sundayRed = !!options.sundayRed;
        var holidayRed = !!options.holidayRed;
        var ghost = !!options.ghost;

        var cellW = Number(options.cell.w);
        var cellH = Number(options.cell.h);
        var cellGapX = Number(options.cell.gapX) || 0;
        var cellGapY = Number(options.cell.gapY) || 0;
        var cellPosAdjY = Number(options.cell.bgOffsetY) || 0; // 背景用
        var cellFill = !!options.cell.fill;

        var st = options.cell.stroke;
        var strokeMode = st.mode || "none";
        var strokeTop = !!st.top;
        var strokeBottom = !!st.bottom;
        var strokeLeft = !!st.left;
        var strokeRight = !!st.right;

        var showTopYear = !!options.year.showTop;
        var includeYearInMonthTitle = !!options.month.includeYear;
        var monthTitleBottomBorder = !!options.month.titleBottomBorder;
        var topYearBottomMargin = Number(options.year.bottomMargin) || 0;

        var monthCount = Math.round(Number(options.layout.monthCount));
        if (!monthCount || monthCount < 1) monthCount = 1;
        var colCount = Math.round(Number(options.layout.colCount));
        if (!colCount || colCount < 1) colCount = 1;

        var monthTitleMode = options.month.titleMode || "pad";
        var monthBottomMargin = Number(options.month.bottomMargin) || 0;
        var weekdayBottomMargin = Number(options.weekday.bottomMargin) || 0;

        var outerMarginX = Number(options.unit.outerX) || 0;
        var outerMarginY = Number(options.unit.outerY) || 0;

        var startFromJanuary = !!options.layout.startFromJanuary;
        var weekStartMonday = (options.weekday.startMonday !== false);
        var weekdayLabelMode = options.weekday.labelMode || "jp";

        var weekdayFontSize = Number(options.weekday.fontSize);
        var monthFontSize = Number(options.month.fontSize);
        var yearFontSize = Number(options.year.fontSize);
        if (!weekdayFontSize || weekdayFontSize <= 0) weekdayFontSize = fontSize;
        if (!monthFontSize || monthFontSize <= 0) monthFontSize = fontSize;
        if (!yearFontSize || yearFontSize <= 0) yearFontSize = fontSize;

        // Delegate to the existing implementation (現行buildCalendarは温存)
        buildCalendar(doc, baseDate, layerName, fontSize, align, monthTitleAlign, fontName, sundayRed, holidayRed, ghost,
            cellW, cellH, cellGapX, cellGapY, cellPosAdjY, cellFill,
            strokeMode, strokeTop, strokeBottom, strokeLeft, strokeRight,
            showTopYear, includeYearInMonthTitle, monthTitleBottomBorder, topYearBottomMargin,
            monthCount, colCount, monthTitleMode, monthBottomMargin, weekdayBottomMargin,
            outerMarginX, outerMarginY, startFromJanuary, weekStartMonday, weekdayLabelMode,
            weekdayFontSize, monthFontSize, yearFontSize);
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
        var baseDate = parseYMDFields(inputY.text, inputM.text, inputD.text);
        if (!baseDate) {
            alert(L("error.badDate"));
            return;
        }
        var monthCount = Math.round(Number(inputMonths.text));
        if (!monthCount || monthCount < 1) monthCount = 1;
        if (monthCount > 24) monthCount = 24;

        var colCount = Math.round(Number(inputCols.text));
        if (!colCount || colCount < 1) colCount = 1;
        if (colCount > 12) colCount = 12;
        if (colCount > monthCount) colCount = monthCount;

        var fontSize = Number(inputFontSize.text);
        if (!fontSize || fontSize <= 0) {
            alert(L("error.badFontSize"));
            return;
        }

        var ghost = (chkGhost && chkGhost.enabled && chkGhost.value) ? true : false;

        // プレビューONで既に描画済みなら、それをそのまま確定（再描画しない）
        var __existingPreview = null;
        try { __existingPreview = getLayerByName(doc, PREVIEW_LAYER_NAME); } catch (_) { }
        var __finalName = "Calendar_" + baseDate.getFullYear() + "_" + pad2(baseDate.getMonth() + 1);
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

        var cellFill = !!(chkCellFill && chkCellFill.value);

        // 罫線設定
        var strokeMode = (rbStrokeNone && rbStrokeNone.value) ? "none"
            : ((rbStrokeCustom && rbStrokeCustom.value) ? "custom" : "all");

        var strokeTop = !!(chkStrokeTop && chkStrokeTop.value);
        var strokeBottom = !!(chkStrokeBottom && chkStrokeBottom.value);
        var strokeLeft = !!(chkStrokeLeft && chkStrokeLeft.value);
        var strokeRight = !!(chkStrokeRight && chkStrokeRight.value);

        if (strokeMode === "all") {
            strokeTop = strokeBottom = strokeLeft = strokeRight = true;
        }

        var cellW = toPtFromUI(inputCellW);
        var cellH = toPtFromUI(inputCellH);
        if (!cellW || isNaN(cellW) || cellW <= 0) return alert(L("error.badCellW"));
        if (!cellH || isNaN(cellH) || cellH <= 0) return alert(L("error.badCellH"));

        var cellGapX = toPtFromUI(inputCellGapX);
        var cellGapY = toPtFromUI(inputCellGapY);
        if (isNaN(cellGapX) || cellGapX < 0) cellGapX = 0;
        if (cellGapX > 2000) cellGapX = 2000;
        if (isNaN(cellGapY) || cellGapY < 0) cellGapY = 0;
        if (cellGapY > 2000) cellGapY = 2000;

        var cellPosAdjY = toPtFromUI(inputCellPosAdjY);
        if (isNaN(cellPosAdjY)) cellPosAdjY = 0;
        if (cellPosAdjY < -2000) cellPosAdjY = -2000;
        if (cellPosAdjY > 2000) cellPosAdjY = 2000;

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

        var weekdayFontSize = Number(inputWeekdayFontSize.text);
        if (!weekdayFontSize || weekdayFontSize <= 0) weekdayFontSize = fontSize;

        var monthFontSize = Number(inputMonthFontSize.text);
        if (!monthFontSize || monthFontSize <= 0) monthFontSize = fontSize;

        var yearFontSize = Number(inputYearFontSize.text);
        if (!yearFontSize || yearFontSize <= 0) yearFontSize = fontSize;

        var options = {
            fontSize: fontSize,
            align: align,
            monthTitleAlign: monthTitleAlign,
            fontName: fontName,
            sundayRed: sundayRed,
            holidayRed: holidayRed,
            ghost: ghost,
            cell: {
                w: cellW,
                h: cellH,
                gapX: cellGapX,
                gapY: cellGapY,
                bgOffsetY: cellPosAdjY,
                fill: cellFill,
                stroke: {
                    mode: strokeMode,
                    top: strokeTop,
                    bottom: strokeBottom,
                    left: strokeLeft,
                    right: strokeRight
                }
            },
            year: {
                showTop: showTopYear,
                bottomMargin: topYearBottomMargin,
                fontSize: yearFontSize
            },
            month: {
                includeYear: includeYearInMonthTitle,
                titleBottomBorder: monthTitleBottomBorder,
                titleMode: monthTitleMode,
                bottomMargin: monthBottomMargin,
                fontSize: monthFontSize
            },
            weekday: {
                startMonday: weekStartMonday,
                labelMode: weekdayLabelMode,
                bottomMargin: weekdayBottomMargin,
                fontSize: weekdayFontSize
            },
            unit: {
                outerX: outerMarginX,
                outerY: outerMarginY
            },
            layout: {
                monthCount: monthCount,
                colCount: colCount,
                startFromJanuary: startFromJanuary
            }
        };

        buildCalendarOpt(doc, baseDate, PREVIEW_LAYER_NAME, options);

        // 確定：プレビューレイヤー名を変更して残す
        var lyr = getLayerByName(doc, PREVIEW_LAYER_NAME);
        if (lyr) lyr.name = __finalName;

        app.redraw();
        dlg.close(1);
    };

    cancelBtn.onClick = function () {
        try {
            if (__SCM_VIEW && __SCM_ORG_ZOOM != null && __SCM_ORG_CENTER != null) {
                __SCM_VIEW.zoom = __SCM_ORG_ZOOM;
                __SCM_VIEW.centerPoint = __SCM_ORG_CENTER;
                app.redraw();
            }
        } catch (_) { }
        removeLayerIfExists(doc, PREVIEW_LAYER_NAME);
        app.redraw();
        dlg.close(0);
    };


    dlg.show();

    // ====== functions ======
    /* 値をmin/maxで丸める / Clamp a number to min/max */
    function clampNumber(v, minV, maxV) {
        if (typeof minV === "number" && v < minV) v = minV;
        if (typeof maxV === "number" && v > maxV) v = maxV;
        return v;
    }

    /**
     * ↑↓キーで数値を増減（Shift: ±10スナップ / Option: ±0.1）
     * @param {EditText} editText
     * @param {{integer?:boolean, min?:number, max?:number, allowNegative?:boolean}} options
     * @param {Function=} onChanged 変更後に呼ぶ（プレビュー更新など）
     */
    function changeValueByArrowKey(editText, options, onChanged) {
        options = options || {};
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
            if (keyboard.altKey && options.integer !== true) {
                value = Math.round(value * 10) / 10; // 小数第1位
            } else {
                value = Math.round(value); // 整数
            }

            // 下限
            if (options.allowNegative !== true && value < 0) value = 0;

            // min/max
            value = clampNumber(value, options.min, options.max);

            event.preventDefault();
            editText.text = String(value);

            if (typeof onChanged === "function") onChanged();
        });
    }
    /* カレンダー本体（年/月/曜日/日付/罫線）を描画 / Draw the calendar body (year/month/weekday/dates/borders) */
    function buildCalendar(
        doc, baseDate, layerName, fontSize, alignMode, monthTitleAlign, fontName,
        sundayRed, holidayRed, ghost,
        cellW, cellH, cellGapX, cellGapY, cellPosAdjY, cellFill,
        strokeMode, strokeTop, strokeBottom, strokeLeft, strokeRight,

        showTopYear, includeYearInMonthTitle, monthTitleBottomBorder, topYearBottomMargin,
        monthCount, colCount, monthTitleMode, monthBottomMargin, weekdayBottomMargin,
        outerMarginX, outerMarginY, startFromJanuary, weekStartMonday, weekdayLabelMode,
        weekdayFontSize, monthFontSize, yearFontSize
    ) {

        // ---- safety aliases (legacy name compatibility) ----
        ghost = !!ghost;
        cellFill = !!cellFill;
        if (!monthTitleMode) monthTitleMode = "pad";
        monthFontSize = Number(monthFontSize);
        if (!monthFontSize || monthFontSize <= 0) monthFontSize = fontSize;

        yearFontSize = Number(yearFontSize);
        if (!yearFontSize || yearFontSize <= 0) yearFontSize = fontSize;

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

        strokeMode = String(strokeMode || "none");
        strokeTop = !!strokeTop;
        strokeBottom = !!strokeBottom;
        strokeLeft = !!strokeLeft;
        strokeRight = !!strokeRight;

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

        // ===== セル塗り（仮）: K10 =====
        var __K10 = null;
        try {
            __K10 = new CMYKColor();
            __K10.cyan = 0;
            __K10.magenta = 0;
            __K10.yellow = 0;
            __K10.black = 10; // K10
        } catch (_) { __K10 = null; }

        // ===== ゴースト文字色: K30 =====
        var __K30 = null;
        try {
            __K30 = new CMYKColor();
            __K30.cyan = 0;
            __K30.magenta = 0;
            __K30.yellow = 0;
            __K30.black = 30; // K30
        } catch (_) { __K30 = null; }

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
        cellGapX = Number(cellGapX);
        if (isNaN(cellGapX) || cellGapX < 0) cellGapX = 0;
        if (cellGapX > 2000) cellGapX = 2000;

        cellGapY = Number(cellGapY);
        if (isNaN(cellGapY) || cellGapY < 0) cellGapY = 0;
        if (cellGapY > 2000) cellGapY = 2000;

        var stepX = cellW + cellGapX;
        var totalW = 7 * cellW + 6 * cellGapX;

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
            // 日付行は weeks 行なので、間隔は (weeks-1) 回だけ入る
            if (cellGapY > 0 && weeks > 1) totalH += (weeks - 1) * cellGapY;

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
            addText(layer, String(baseYear), startX0, startY0, yearFontSize, "center", blockW, fontName, false, null);
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
            addText(layer, titleText, startX, startY, monthFontSize, monthTitleAlign || alignMode, totalW, fontName, false, monthGroup);

            // タイトル行の下ボーダー（タイトルと曜日の中間）
            if (monthTitleBottomBorder) {
                var by = startY - cellH - (monthBottomMargin * 0.5);
                addHLine(layer, startX, startX + totalW, by, 0.3, monthGroup);
            }

            // 曜日ヘッダ（タイトル下にマージンを入れる）
            var headerY = startY - cellH - monthBottomMargin;
            for (var c = 0; c < 7; c++) {
                var isSunH = (sundayRed && c === sundayCol);
                addText(layer, headers[c], startX + c * stepX, headerY, weekdayFontSize || fontSize, alignMode, cellW, fontName, isSunH, monthGroup);
            }

            // 日付（必要なら前後月も表示）
            var totalCells = info.weeks * 7;

            // 前月の日数（ゴースト用）
            var prevDaysInMonth = 0;
            if (ghost) {
                try {
                    var py = info.year;
                    var pm0 = info.month0 - 1;
                    if (pm0 < 0) { pm0 = 11; py--; }
                    prevDaysInMonth = (new Date(py, pm0 + 1, 0)).getDate();
                } catch (_) { prevDaysInMonth = 0; }
            }

            for (var idx = 0; idx < totalCells; idx++) {
                var col = idx % 7;
                var weekRow = Math.floor(idx / 7);
                var row = 1 + weekRow; // 1=曜日ヘッダの次の行

                // 当月の「日」(1..daysInMonth)
                var day = (idx - info.startIndex) + 1;

                var isGhostDay = false;
                var dispDay = day;

                // ゴーストOFF時は、当月外セル（前月/翌月）は完全にスキップ（罫線も描かない）
                if (!ghost && (day < 1 || day > info.daysInMonth)) {
                    continue;
                }

                // はみ出しは前後月
                if (day < 1 || day > info.daysInMonth) {
                    if (!ghost) continue; // ゴーストOFFなら描画しない
                    isGhostDay = true;
                    if (day < 1) {
                        dispDay = prevDaysInMonth + day; // 前月
                    } else {
                        dispDay = day - info.daysInMonth; // 翌月
                    }
                }

                var x = startX + col * stepX;
                var yy = headerY - weekdayBottomMargin - row * cellH - (row - 1) * cellGapY;

                // 背景の位置調整（背景のみ上下調整）
                var __bgAdjY = Number(cellPosAdjY);
                if (isNaN(__bgAdjY)) __bgAdjY = 0;

                // 背景・罫線の基準（背景オフセット込み）
                var __cellTopY = yy + __bgAdjY;
                var __cellBotY = __cellTopY - cellH;
                var __xL = x;
                var __xR = x + cellW;

                var __cellRect = null;
                if (cellFill && __K10) {
                    try {
                        __cellRect = layer.pathItems.rectangle(__cellTopY, x, cellW, cellH);
                        __cellRect.stroked = false;
                        __cellRect.filled = true;
                        __cellRect.fillColor = __K10;
                        if (monthGroup) {
                            try { __cellRect.move(monthGroup, ElementPlacement.PLACEATBEGINNING); } catch (_) { }
                        }
                    } catch (_) { }
                }

                // 日曜の赤は当月のみ（ゴーストはK30で上書き）
                var isSunD = (sundayRed && col === sundayCol);
                var colorFlag = isSunD && !isGhostDay;

                var tf = addText(layer, String(dispDay), x, yy, fontSize, alignMode, cellW, fontName, colorFlag, monthGroup);

                // ===== 罫線（セル背景に連動 / 背景より前面）=====
                if (strokeMode !== "none") {
                    try {
                        var __strokeW = 0.3; // 固定（必要ならUI化）
                        var __strokeC = null;
                        try {
                            __strokeC = new CMYKColor();
                            __strokeC.cyan = 0; __strokeC.magenta = 0; __strokeC.yellow = 0; __strokeC.black = 100; // K100
                        } catch (_) { __strokeC = null; }

                        function __drawLine(x1, y1, x2, y2) {
                            var ln = layer.pathItems.add();
                            ln.stroked = true;
                            ln.filled = false;
                            ln.strokeWidth = __strokeW;
                            if (__strokeC) ln.strokeColor = __strokeC;
                            ln.setEntirePath([[x1, y1], [x2, y2]]);
                            if (monthGroup) {
                                try { ln.move(monthGroup, ElementPlacement.PLACEATEND); } catch (_) { }
                            }
                            return ln;
                        }

                        // モードに応じて描画
                        if (strokeMode === "all") {
                            __drawLine(__xL, __cellTopY, __xR, __cellTopY); // 上
                            __drawLine(__xL, __cellBotY, __xR, __cellBotY); // 下
                            __drawLine(__xL, __cellBotY, __xL, __cellTopY); // 左
                            __drawLine(__xR, __cellBotY, __xR, __cellTopY); // 右
                        } else {
                            if (strokeTop) __drawLine(__xL, __cellTopY, __xR, __cellTopY);
                            if (strokeBottom) __drawLine(__xL, __cellBotY, __xR, __cellBotY);
                            if (strokeLeft) __drawLine(__xL, __cellBotY, __xL, __cellTopY);
                            if (strokeRight) __drawLine(__xR, __cellBotY, __xR, __cellTopY);
                        }
                    } catch (_) { }
                }

                // 祝日（当月のみ）
                if (!isGhostDay) {
                    var hname = getJPHolidayName(info.year, info.month0 + 1, day);
                    var isHol = (holidayRed && hname);
                    if (isHol) {
                        try {
                            var oc = new RGBColor();
                            oc.red = 255;
                            oc.green = 140;
                            oc.blue = 0; // orange
                            tf.textRange.characterAttributes.fillColor = oc;
                        } catch (_) { }
                    }
                }

                // ゴースト文字色（K30）
                if (isGhostDay && __K30 && tf) {
                    try { tf.textRange.characterAttributes.fillColor = __K30; } catch (_) { }
                }

                // 前後関係を最終確定（背景→最背面、テキスト→最前面）
                try { if (__cellRect) __cellRect.zOrder(ZOrderMethod.SENDTOBACK); } catch (_) { }
                try { if (tf) tf.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (_) { }
            }
        }
    }

    /* 水平罫線を1本追加 / Add a single horizontal rule */
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

    /* 揃え指定でテキストフレームを1つ追加 / Add one text frame with the given alignment */
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

    /* 数値を2桁ゼロ埋め / Zero-pad a number to 2 digits */
    function pad2(n) {
        return (n < 10) ? ("0" + n) : String(n);
    }

    /* 年月日入力を検証しDateを返す（不正はnull）/ Validate Y/M/D fields and return a Date (null if invalid) */
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

    /* 名前でレイヤーを取得 / Get a layer by name */
    function getLayerByName(doc, name) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === name) return doc.layers[i];
        }
        return null;
    }

    /* レイヤーを取得、無ければ作成 / Get a layer, creating it if missing */
    function getOrCreateLayer(doc, name) {
        var lyr = getLayerByName(doc, name);
        if (lyr) return lyr;
        lyr = doc.layers.add();
        lyr.name = name;
        return lyr;
    }

    /* レイヤーを削除して同名で作り直す / Remove a layer and recreate it under the same name */
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

    /* 同名レイヤーがあれば削除 / Remove a layer if it exists */
    function removeLayerIfExists(doc, name) {
        var lyr = getLayerByName(doc, name);
        if (!lyr) return;

        // ロック解除して削除
        lyr.locked = false;
        lyr.visible = true;
        lyr.remove();
    }
})();