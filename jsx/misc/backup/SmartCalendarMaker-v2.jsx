#target illustrator
#targetengine "SmartCalendarMaker"

    /* =====================================================
     * SmartCalendarMaker.jsx
     * 概要: 基準日を基準に、指定月数ぶんのカレンダー（月曜はじまり）をアートボード中心に作成します。
     *      ダイアログ上の変更はプレビューで即時反映し、OKで確定レイヤーとして残します。
     * 更新日: 2026-02-14
     * ===================================================== */

    (function () {
        if (app.documents.length === 0) {
            alert("ドキュメントを開いてから実行してください。");
            return;
        }

        var doc = app.activeDocument;

        function getUnitLabel() {
            switch (doc.rulerUnits) {
                case RulerUnits.Millimeters: return "mm";
                case RulerUnits.Centimeters: return "cm";
                case RulerUnits.Inches: return "in";
                case RulerUnits.Picas: return "pc";
                case RulerUnits.Pixels: return "px";
                default: return "pt";
            }
        }
        var unitLabel = getUnitLabel();

        // --- Unit conversion helpers ---
        function unitValueToPt(v) {
            v = Number(v);
            if (isNaN(v)) return NaN;

            // Convert from current ruler units to points.
            try {
                switch (doc.rulerUnits) {
                    case RulerUnits.Millimeters: return v * 72 / 25.4;
                    case RulerUnits.Centimeters: return v * 72 / 2.54;
                    case RulerUnits.Inches: return v * 72;
                    case RulerUnits.Picas: return v * 12; // 1 pica = 12 pt
                    case RulerUnits.Pixels:
                        // px -> pt depends on document resolution
                        var res = 72;
                        try {
                            if (doc.rasterEffectSettings && doc.rasterEffectSettings.resolution) {
                                res = Number(doc.rasterEffectSettings.resolution) || 72;
                            }
                        } catch (_) { }
                        return v * 72 / res;
                    default:
                        return v; // points
                }
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
                switch (doc.rulerUnits) {
                    case RulerUnits.Millimeters: return pt * 25.4 / 72;
                    case RulerUnits.Centimeters: return pt * 2.54 / 72;
                    case RulerUnits.Inches: return pt / 72;
                    case RulerUnits.Picas: return pt / 12;
                    case RulerUnits.Pixels:
                        var res = 72;
                        try {
                            if (doc.rasterEffectSettings && doc.rasterEffectSettings.resolution) {
                                res = Number(doc.rasterEffectSettings.resolution) || 72;
                            }
                        } catch (_) { }
                        return pt * res / 72;
                    default:
                        return pt;
                }
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

        var dlg = new Window("dialog", "カレンダー作成");
        dlg.orientation = "column";
        dlg.alignChildren = "fill";

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
        var pnlBaseDate = gColL.add("panel", undefined, "基本設定");
        pnlBaseDate.orientation = "column";
        pnlBaseDate.alignChildren = "left";
        pnlBaseDate.margins = [15, 20, 15, 10];

        var g1 = pnlBaseDate.add("group");
        g1.add("statictext", undefined, "年:");
        var inputY = g1.add("edittext", undefined, String(today.getFullYear()));
        inputY.characters = 4;

        g1.add("statictext", undefined, "月:");
        var inputM = g1.add("edittext", undefined, String(today.getMonth() + 1));
        inputM.characters = 2;

        g1.add("statictext", undefined, "日:");
        var inputD = g1.add("edittext", undefined, String(today.getDate()));
        inputD.characters = 2;

        // 月数プリセット（ラジオ）
        var gMonthPreset = pnlBaseDate.add("group");
        gMonthPreset.orientation = "row";
        gMonthPreset.alignChildren = ["left", "center"];

        var rbPreset1 = gMonthPreset.add("radiobutton", undefined, "1ヶ月");
        var rbPreset3 = gMonthPreset.add("radiobutton", undefined, "3ヶ月");
        var rbPreset12 = gMonthPreset.add("radiobutton", undefined, "12ヶ月");

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
            try { refreshPreview(); } catch (_) { }
        }

        rbPreset1.onClick = function () { applyMonthPreset(1, 1); };
        rbPreset3.onClick = function () { applyMonthPreset(3, 3); };
        rbPreset12.onClick = function () { applyMonthPreset(12, 3); };

        // 月の起点
        var gMonthStart = pnlBaseDate.add("group");
        gMonthStart.orientation = "row";
        gMonthStart.alignChildren = ["left", "center"];

        var rbStartCurrent = gMonthStart.add("radiobutton", undefined, "当月基準");
        var rbStartJan = gMonthStart.add("radiobutton", undefined, "1月から");

        rbStartCurrent.value = true; // デフォルト
        rbStartCurrent.onClick = refreshPreview;
        rbStartJan.onClick = refreshPreview;




        // ===== 年パネル =====
        var pnlYear = gColR.add("panel", undefined, "年");
        pnlYear.orientation = "column";
        pnlYear.alignChildren = "left";
        pnlYear.margins = [15, 20, 15, 10];

        var gYear = pnlYear.add("group");
        gYear.orientation = "row";
        gYear.alignChildren = ["left", "center"];


        var chkTopYear = gYear.add("checkbox", undefined, "年を表示");
        chkTopYear.value = false;      // 12ヶ月のときだけONにする
        chkTopYear.enabled = false;    // 12ヶ月以外は触れない
        chkTopYear.onClick = function () {
            try { gYearMargin.enabled = chkTopYear.value; } catch (_) { }
            refreshPreview();
        };

        // 年タイトル下のマージン（pt）
        var gYearMargin = pnlYear.add("group");
        gYearMargin.orientation = "row";
        gYearMargin.alignChildren = ["left", "center"];
        gYearMargin.add("statictext", undefined, "下マージン");
        var inputTopYearBottomMargin = gYearMargin.add("edittext", undefined, "3");
        inputTopYearBottomMargin.characters = 4;
        gYearMargin.add("statictext", undefined, unitLabel);
        inputTopYearBottomMargin.onChanging = refreshPreview;

        try { gYearMargin.enabled = chkTopYear.value; } catch (_) { }

        // ===== 月パネル =====
        var pnlMonth = gColR.add("panel", undefined, "月");
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

        var chkMonthYear = gMonth.add("checkbox", undefined, "年を併記");
        chkMonthYear.value = true;
        chkMonthYear.onClick = refreshPreview;

        // 月タイトルの揃え
        var gMonthAlign = pnlMonth.add("group");
        gMonthAlign.orientation = "row";
        gMonthAlign.alignChildren = ["left", "center"];
        gMonthAlign.add("statictext", undefined, "揃え：");
        var rbMonthAlignL = gMonthAlign.add("radiobutton", undefined, "左");
        var rbMonthAlignC = gMonthAlign.add("radiobutton", undefined, "中央");
        var rbMonthAlignR = gMonthAlign.add("radiobutton", undefined, "右");
        rbMonthAlignC.value = true; // default
        rbMonthAlignL.onClick = refreshPreview;
        rbMonthAlignC.onClick = refreshPreview;
        rbMonthAlignR.onClick = refreshPreview;

        // 月タイトルの表記
        var pnlMonthFmt = pnlMonth.add("panel", undefined, "表記");
        pnlMonthFmt.orientation = "column";
        pnlMonthFmt.alignChildren = "left";
        pnlMonthFmt.margins = [15, 20, 15, 10];

        var gMonthFmt = pnlMonthFmt.add("group");
        gMonthFmt.orientation = "column";
        gMonthFmt.alignChildren = ["left", "top"];

        var rbMonthNum = gMonthFmt.add("radiobutton", undefined, "数字");
        var rbMonthPad = gMonthFmt.add("radiobutton", undefined, "数字（0埋め）");
        var rbMonthEn = gMonthFmt.add("radiobutton", undefined, "英語");
        var rbMonthEnS = gMonthFmt.add("radiobutton", undefined, "英語（短縮）");

        // デフォルトは 0埋め
        rbMonthPad.value = true;

        rbMonthNum.onClick = refreshPreview;
        rbMonthPad.onClick = refreshPreview;
        rbMonthEn.onClick = refreshPreview;
        rbMonthEnS.onClick = refreshPreview;

        // 月タイトル下のマージン（pt）
        var gMonthMargin = pnlMonth.add("group");
        gMonthMargin.orientation = "row";
        gMonthMargin.alignChildren = ["left", "center"];
        gMonthMargin.add("statictext", undefined, "下マージン");
        var inputMonthBottomMargin = gMonthMargin.add("edittext", undefined, "3");
        inputMonthBottomMargin.characters = 4;
        gMonthMargin.add("statictext", undefined, unitLabel);
        inputMonthBottomMargin.onChanging = refreshPreview;

        // 月タイトル行の下ボーダー
        var chkMonthBottomBorder = pnlMonth.add("checkbox", undefined, "下ボーダー");
        chkMonthBottomBorder.value = true; // デフォルトON
        chkMonthBottomBorder.onClick = refreshPreview;
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
        var pnlWeekdayLabel = gColR.add("panel", undefined, "曜日");
        pnlWeekdayLabel.orientation = "column";
        pnlWeekdayLabel.alignChildren = "left";
        pnlWeekdayLabel.margins = [15, 20, 15, 10];

        // 週の始まりの曜日（panel）
        var pnlWeekStart = pnlWeekdayLabel.add("panel", undefined, "週の始まりの曜日");
        pnlWeekStart.orientation = "column";
        pnlWeekStart.alignChildren = "left";
        pnlWeekStart.margins = [15, 20, 15, 10];

        var gWeekStart = pnlWeekStart.add("group");
        gWeekStart.orientation = "row";
        gWeekStart.alignChildren = ["left", "center"];

        var rbWeekMon = gWeekStart.add("radiobutton", undefined, "月曜日");
        var rbWeekSun = gWeekStart.add("radiobutton", undefined, "日曜日");

        rbWeekMon.value = true; // デフォルト
        rbWeekMon.onClick = refreshPreview;
        rbWeekSun.onClick = refreshPreview;

        // 曜日ヘッダの表記
        var pnlWeekdayFmt = pnlWeekdayLabel.add("panel", undefined, "表記");
        pnlWeekdayFmt.orientation = "column";
        pnlWeekdayFmt.alignChildren = "left";
        pnlWeekdayFmt.margins = [15, 20, 15, 10];

        var gWeekdayLabel = pnlWeekdayFmt.add("group");
        gWeekdayLabel.orientation = "row";
        gWeekdayLabel.alignChildren = ["left", "center"];

        var rbWdJP = gWeekdayLabel.add("radiobutton", undefined, "月");
        var rbWdMTW = gWeekdayLabel.add("radiobutton", undefined, "M");
        var rbWdMon = gWeekdayLabel.add("radiobutton", undefined, "Mon");

        // 曜日ヘッダ下のマージン（pt）
        var gWeekdayBottomMargin = pnlWeekdayLabel.add("group");
        gWeekdayBottomMargin.orientation = "row";
        gWeekdayBottomMargin.alignChildren = ["left", "center"];
        gWeekdayBottomMargin.add("statictext", undefined, "下マージン");
        var inputWeekdayBottomMargin = gWeekdayBottomMargin.add("edittext", undefined, "2");
        inputWeekdayBottomMargin.characters = 4;
        gWeekdayBottomMargin.add("statictext", undefined, unitLabel);
        inputWeekdayBottomMargin.onChanging = refreshPreview;

        rbWdJP.value = true; // default
        rbWdJP.onClick = refreshPreview;
        rbWdMTW.onClick = refreshPreview;
        rbWdMon.onClick = refreshPreview;

        // ===== レイアウトパネル =====
        var pnlLayout = gColL.add("panel", undefined, "レイアウト");
        pnlLayout.orientation = "column";
        pnlLayout.alignChildren = "left";
        pnlLayout.margins = [15, 20, 15, 10];

        // 月数/列数
        var gCount = pnlLayout.add("group");
        gCount.orientation = "row";
        gCount.alignChildren = ["left", "center"];
        gCount.add("statictext", undefined, "月数:");
        var inputMonths = gCount.add("edittext", undefined, "1");
        inputMonths.characters = 3;
        try { rbPreset1.value = true; } catch (_) { }
        gCount.add("statictext", undefined, "列数:");
        var inputCols = gCount.add("edittext", undefined, "1");
        inputCols.characters = 3;

        // セル（panel）
        var pnlCell = pnlLayout.add("panel", undefined, "セル");
        pnlCell.orientation = "column";
        pnlCell.alignChildren = "left";
        pnlCell.margins = [15, 20, 15, 10];

        var gCellW = pnlCell.add("group");
        var stCellW = gCellW.add("statictext", undefined, "幅");
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
        gCellW.add("statictext", undefined, "text/units");

        var gCellH = pnlCell.add("group");
        var stCellH = gCellH.add("statictext", undefined, "高さ");
        stCellH.justification = "right";
        stCellH.preferredSize.width = 30;
        var __defaultCellH_pt = Math.round(__fs0 * 1.3);
        var __defaultCellH = Math.round(ptToUnitValue(__defaultCellH_pt));
        var inputCellH = gCellH.add("edittext", undefined, String(__defaultCellH));
        inputCellH.characters = 5;
        gCellH.add("statictext", undefined, "text/units");

        inputCellW.onChanging = refreshPreview;
        inputCellH.onChanging = refreshPreview;

        changeValueByArrowKey(inputCellW, { integer: false, min: 1, max: 2000 }, refreshPreview);
        changeValueByArrowKey(inputCellH, { integer: false, min: 1, max: 2000 }, refreshPreview);

        // 月（ユニット間マージン）panel
        var pnlMonthOuter = pnlLayout.add("panel", undefined, "月");
        pnlMonthOuter.orientation = "column";
        pnlMonthOuter.alignChildren = "left";
        pnlMonthOuter.margins = [15, 20, 15, 10];
        pnlMonthOuter.enabled = false; // 月数=1 のときはディム（refreshPreviewで同期）

        var gOuterH = pnlMonthOuter.add("group");
        gOuterH.orientation = "row";
        gOuterH.alignChildren = ["left", "center"];
        gOuterH.add("statictext", undefined, "左右");
        var inputOuterMarginX = gOuterH.add("edittext", undefined, "20");
        inputOuterMarginX.characters = 4;
        gOuterH.add("statictext", undefined, unitLabel);
        inputOuterMarginX.onChanging = refreshPreview;

        var gOuterV = pnlMonthOuter.add("group");
        gOuterV.orientation = "row";
        gOuterV.alignChildren = ["left", "center"];
        gOuterV.add("statictext", undefined, "上下");
        var inputOuterMarginY = gOuterV.add("edittext", undefined, "3");
        inputOuterMarginY.characters = 4;
        gOuterV.add("statictext", undefined, unitLabel);
        inputOuterMarginY.onChanging = refreshPreview;

        // ===== 書式パネル =====
        var pnlFormat = gColL.add("panel", undefined, "書式");
        pnlFormat.orientation = "column";
        pnlFormat.alignChildren = "left";
        pnlFormat.margins = [15, 20, 15, 10];


        // フォント（インストール済み）選択
        var gFontName = pnlFormat.add("group");
        gFontName.add("statictext", undefined, "フォント");
        var ddFont = gFontName.add("dropdownlist", undefined, []);
        gFontName.alignChildren = ["left", "center"];
        ddFont.minimumSize = [80, 22];
        ddFont.preferredSize = [200, 22];
        ddFont.maximumSize = [200, 22];

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
                var w = new Window("palette", title || "準備中");
                w.orientation = "column";
                w.alignChildren = "fill";
                w.margins = 12;

                var msg = w.add("statictext", undefined, initialText || "処理中…");
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

            var __loading = __createLoadingPalette("準備中", "フォント一覧を読み込み中…", 100);

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
                            __loading.setText("フォント一覧を読み込み中… (" + fi + "/" + __len + ")");
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
                __loading.setText("準備完了");
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
                        var w = new Window("palette", title || "準備中");
                        w.orientation = "column";
                        w.alignChildren = "fill";
                        w.margins = 12;
                        var msg = w.add("statictext", undefined, initialText || "処理中…");
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
                __ddLoading = __createLoadingPalette("準備中", "フォント一覧をセット中…", fontNames.length);
            }
        } catch (_) { }

        for (var iFont = 0; iFont < fontNames.length; iFont++) {
            ddFont.add("item", fontNames[iFont]);
            if (__ddLoading && (iFont % 50) === 0) {
                try {
                    __ddLoading.setValue(iFont);
                    __ddLoading.setText("フォント一覧をセット中… (" + iFont + "/" + fontNames.length + ")");
                } catch (_) { }
            }
        }

        if (__ddLoading) {
            try {
                __ddLoading.setValue(fontNames.length);
                __ddLoading.setText("準備完了");
            } catch (_) { }
            __ddLoading.close();
        }


        var gFont = pnlFormat.add("group");
        gFont.add("statictext", undefined, "フォントサイズ");
        var inputFontSize = gFont.add("edittext", undefined, "12");
        inputFontSize.characters = 5;
        gFont.add("statictext", undefined, "pt");

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

        gAlign.add("statictext", undefined, "揃え：");

        var rbLeft = gAlign.add("radiobutton", undefined, "左");
        var rbCenter = gAlign.add("radiobutton", undefined, "中央");
        var rbRight = gAlign.add("radiobutton", undefined, "右");
        rbCenter.value = true; // デフォルト中央
        rbLeft.onClick = refreshPreview;
        rbCenter.onClick = refreshPreview;
        rbRight.onClick = refreshPreview;



        // 日曜日を赤にする
        var chkSundayRed = pnlFormat.add("checkbox", undefined, "日曜日のカラー");
        chkSundayRed.value = true;
        chkSundayRed.onClick = refreshPreview;

        // 祝日チェックボックス（色は橙色で表示）
        var chkHolidayRed = pnlFormat.add("checkbox", undefined, "祝日");
        chkHolidayRed.value = true;
        chkHolidayRed.onClick = refreshPreview;


        var g2 = dlg.add("group");
        var previewChk = g2.add("checkbox", undefined, "プレビュー");
        previewChk.value = true;

        // var hint = dlg.add("statictext", undefined, "入力した日付の当月カレンダー（月曜はじまり）をアートボード中心に配置します。");

        var btns = dlg.add("group");
        btns.alignment = "right";

        var cancelBtn = btns.add("button", undefined, "キャンセル", { name: "cancel" });
        var okBtn = btns.add("button", undefined, "作成", { name: "ok" });

        // ===== キー操作（↑↓で数値増減）=====
        changeValueByArrowKey(inputY, { integer: true, min: 1, max: 9999 }, refreshPreview);
        changeValueByArrowKey(inputM, { integer: true, min: 1, max: 12 }, refreshPreview);
        changeValueByArrowKey(inputD, { integer: true, min: 1, max: 31 }, refreshPreview);
        changeValueByArrowKey(inputMonths, { integer: true, min: 1, max: 24 }, refreshPreview);
        changeValueByArrowKey(inputCols, { integer: true, min: 1, max: 12 }, refreshPreview);
        changeValueByArrowKey(inputMonthBottomMargin, { integer: false, min: 0, max: 2000 }, refreshPreview);
        changeValueByArrowKey(inputTopYearBottomMargin, { integer: false, min: 0, max: 2000 }, refreshPreview);
        changeValueByArrowKey(inputWeekdayBottomMargin, { integer: false, min: 0, max: 2000 }, refreshPreview);
        changeValueByArrowKey(inputFontSize, { integer: false, min: 0.1, max: 9999 }, refreshPreview);
        changeValueByArrowKey(inputOuterMarginX, { integer: false, min: 0, max: 5000 }, refreshPreview);
        changeValueByArrowKey(inputOuterMarginY, { integer: false, min: 0, max: 5000 }, refreshPreview);


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
            buildCalendar(doc, base, PREVIEW_LAYER_NAME, fs, align, monthTitleAlign, fontName, sundayRed, holidayRed, cellW, cellH, showTopYear, includeYearInMonthTitle, monthTitleBottomBorder, topYearBottomMargin, monthCount, colCount, monthTitleMode, monthBottomMargin, weekdayBottomMargin, outerMarginX, outerMarginY, startFromJanuary, weekStartMonday, weekdayLabelMode);
            app.redraw();
        }

        inputY.onChanging = refreshPreview;
        inputM.onChanging = refreshPreview;
        inputD.onChanging = refreshPreview;
        inputMonths.onChanging = refreshPreview;
        inputCols.onChanging = refreshPreview;
        previewChk.onClick = refreshPreview;
        if (ddFont) ddFont.onChange = refreshPreview;

        // キーボードショートカット: L=左 / C=中央 / R=右
        dlg.addEventListener("keydown", function (event) {
            if (event.keyName === "L") {
                rbLeft.value = true;
                refreshPreview();
                event.preventDefault();
            } else if (event.keyName === "C") {
                rbCenter.value = true;
                refreshPreview();
                event.preventDefault();
            } else if (event.keyName === "R") {
                rbRight.value = true;
                refreshPreview();
                event.preventDefault();
            }
        });

        okBtn.onClick = function () {
            var base = parseYMDFields(inputY.text, inputM.text, inputD.text);
            if (!base) {
                alert("日付形式が正しくありません。例: 2026-02-14");
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
                alert("フォントサイズが正しくありません。");
                return;
            }

            // プレビューが無い場合でも確定作成
            clearLayerContents(doc, PREVIEW_LAYER_NAME);
            var align = rbLeft.value ? "left" : (rbRight.value ? "right" : "center");
            var monthTitleAlign = rbMonthAlignL.value ? "left" : (rbMonthAlignR.value ? "right" : "center");
            var fontName = (ddFont && ddFont.selection) ? ddFont.selection.text : null;
            var sundayRed = (chkSundayRed && chkSundayRed.value) ? true : false;
            var holidayRed = (chkHolidayRed && chkHolidayRed.value) ? true : false;

            var cellW = toPtFromUI(inputCellW);
            var cellH = toPtFromUI(inputCellH);
            if (!cellW || isNaN(cellW) || cellW <= 0) return alert("セル幅が正しくありません。");
            if (!cellH || isNaN(cellH) || cellH <= 0) return alert("セル高さが正しくありません。");

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
            buildCalendar(doc, base, PREVIEW_LAYER_NAME, fs, align, monthTitleAlign, fontName, sundayRed, holidayRed, cellW, cellH, showTopYear, includeYearInMonthTitle, monthTitleBottomBorder, topYearBottomMargin, monthCount, colCount, monthTitleMode, monthBottomMargin, weekdayBottomMargin, outerMarginX, outerMarginY, startFromJanuary, weekStartMonday, weekdayLabelMode);

            // 確定：プレビューレイヤー名を変更して残す
            var lyr = getLayerByName(doc, PREVIEW_LAYER_NAME);
            if (lyr) lyr.name = "Calendar_" + base.getFullYear() + "_" + pad2(base.getMonth() + 1);

            app.redraw();
            dlg.close(1);
        };

        cancelBtn.onClick = function () {
            removeLayerIfExists(doc, PREVIEW_LAYER_NAME);
            app.redraw();
            dlg.close(0);
        };

        // ダイアログ表示後に初回プレビュー（表示を優先）
        dlg.onShow = function () {
            try { refreshPreview(); } catch (_) { }
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

        function buildCalendar(doc, baseDate, layerName, fontSize, alignMode, monthTitleAlign, fontName, sundayRed, holidayRed, cellW, cellH, showTopYear, includeYearInMonthTitle, monthTitleBottomBorder, topYearBottomMargin, monthCount, colCount, monthTitleMode, monthBottomMargin, weekdayBottomMargin, outerMarginX, outerMarginY, startFromJanuary, weekStartMonday, weekdayLabelMode) {
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
                addText(layer, String(baseYear), startX0, startY0, fontSize + 6, "center", blockW, fontName, false);
            }

            for (var bi = 0; bi < monthInfos.length; bi++) {
                var info = monthInfos[bi];

                var colI = bi % colCount;
                var rowI = Math.floor(bi / colCount);

                // 各月の左上（各セルは maxCellH の上揃え）
                var startX = startX0 + colI * (totalW + gapX);
                var startY = (startY0 - yearTitleH) - rowI * (maxCellH + gapY);

                // タイトル
                var titleText = formatMonthTitle(info.year, info.month0, includeYearInMonthTitle, monthTitleMode);
                addText(layer, titleText, startX, startY, fontSize + 2, monthTitleAlign || alignMode, totalW, fontName, false);

                // タイトル行の下ボーダー（タイトルと曜日の中間）
                if (monthTitleBottomBorder) {
                    var by = startY - cellH - (monthBottomMargin * 0.5);
                    addHLine(layer, startX, startX + totalW, by, 0.3);
                }

                // 曜日ヘッダ（タイトル下にマージンを入れる）
                var headerY = startY - cellH - monthBottomMargin;
                for (var c = 0; c < 7; c++) {
                    var isSunH = (sundayRed && c === sundayCol);
                    addText(layer, headers[c], startX + c * cellW, headerY, fontSize, alignMode, cellW, fontName, isSunH);
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
                    var tf = addText(layer, String(d), x, yy, fontSize, alignMode, cellW, fontName, colorFlag);

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

        function addHLine(layer, x1, x2, y, strokeW) {
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
            return p;
        }

        function addText(layer, str, xLeft, y, size, alignMode, boxW, fontName, isRed) {
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
            if (alignMode === "right") {
                tf.textRange.justification = Justification.RIGHT;
            } else if (alignMode === "left") {
                tf.textRange.justification = Justification.LEFT;
            } else {
                tf.textRange.justification = Justification.CENTER;
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

            // ロック解除して内容のみ削除（レイヤーは残す）
            lyr.locked = false;
            lyr.visible = true;

            // pageItems を逆順で削除（TextFrame / PathItem など全て）
            try {
                for (var i = lyr.pageItems.length - 1; i >= 0; i--) {
                    try { lyr.pageItems[i].remove(); } catch (_) { }
                }
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
