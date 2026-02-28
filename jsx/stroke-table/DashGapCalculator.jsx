#target illustrator
try { app.preferences.setBooleanPreference('ShowExternalJSXWarning', false); } catch (_) { }

/*
 * 破線計算機（Gap→Dash）
 * 更新日: 2026-02-28
 * Version: v2.0
 * 概要: パス（オープン/クローズ）を複数選択して、分割数と間隔から線分長を算出して破線を適用します。
 *      「開始位置（オフセット）」で破線の開始位置（位相）も指定できます。
 */

var SCRIPT_VERSION = "v2.0";

// --- 前回値の記憶（app.putCustomOptions） ---
var PREF_KEY = "DashCalcPrefs_GapToDash_v1";

function loadPrefs() {
    try {
        var d = app.getCustomOptions(PREF_KEY);
        var p = {};

        var kSegments = stringIDToTypeID("segments");
        var kGapPt = stringIDToTypeID("gapPt");
        var kOffsetPt = stringIDToTypeID("offsetPt");
        var kCapMode = stringIDToTypeID("capMode");
        var kReverse = stringIDToTypeID("reversePath");
        var kAdjustEnds = stringIDToTypeID("adjustEnds");
        var kDashPt = stringIDToTypeID("dashPt");
        var kMode = stringIDToTypeID("mode");

        var kUseOffset = stringIDToTypeID("useOffset");

        // --- Random pattern keys ---
        var kRand0Pt = stringIDToTypeID("rand0Pt");
        var kRand1Pt = stringIDToTypeID("rand1Pt");
        var kRand2Pt = stringIDToTypeID("rand2Pt");
        var kRand3Pt = stringIDToTypeID("rand3Pt");
        var kRand4Pt = stringIDToTypeID("rand4Pt");
        var kRand5Pt = stringIDToTypeID("rand5Pt");

        if (d.hasKey(kSegments)) p.segments = d.getInteger(kSegments);
        if (d.hasKey(kGapPt)) p.gapPt = d.getDouble(kGapPt);
        if (d.hasKey(kOffsetPt)) p.offsetPt = d.getDouble(kOffsetPt);
        if (d.hasKey(kCapMode)) p.capMode = d.getInteger(kCapMode);
        if (d.hasKey(kReverse)) p.reversePath = d.getBoolean(kReverse);
        if (d.hasKey(kAdjustEnds)) p.adjustEnds = d.getBoolean(kAdjustEnds);
        if (d.hasKey(kDashPt)) p.dashPt = d.getDouble(kDashPt);
        if (d.hasKey(kMode)) p.mode = d.getInteger(kMode);
        if (d.hasKey(kUseOffset)) p.useOffset = d.getBoolean(kUseOffset);

        if (d.hasKey(kRand0Pt)) p.rand0Pt = d.getDouble(kRand0Pt);
        if (d.hasKey(kRand1Pt)) p.rand1Pt = d.getDouble(kRand1Pt);
        if (d.hasKey(kRand2Pt)) p.rand2Pt = d.getDouble(kRand2Pt);
        if (d.hasKey(kRand3Pt)) p.rand3Pt = d.getDouble(kRand3Pt);
        if (d.hasKey(kRand4Pt)) p.rand4Pt = d.getDouble(kRand4Pt);
        if (d.hasKey(kRand5Pt)) p.rand5Pt = d.getDouble(kRand5Pt);

        return p;
    } catch (_) {
        return null;
    }
}

function savePrefs(p) {
    try {
        var d = new ActionDescriptor();
        d.putInteger(stringIDToTypeID("segments"), p.segments);
        d.putDouble(stringIDToTypeID("gapPt"), p.gapPt);
        d.putDouble(stringIDToTypeID("offsetPt"), p.offsetPt);
        d.putInteger(stringIDToTypeID("capMode"), p.capMode);
        d.putBoolean(stringIDToTypeID("reversePath"), !!p.reversePath);
        d.putBoolean(stringIDToTypeID("adjustEnds"), !!p.adjustEnds);
        d.putDouble(stringIDToTypeID("dashPt"), p.dashPt);
        d.putInteger(stringIDToTypeID("mode"), p.mode);
        d.putBoolean(stringIDToTypeID("useOffset"), !!p.useOffset);
        if (p.randPt && p.randPt.length === 6) {
            d.putDouble(stringIDToTypeID("rand0Pt"), p.randPt[0]);
            d.putDouble(stringIDToTypeID("rand1Pt"), p.randPt[1]);
            d.putDouble(stringIDToTypeID("rand2Pt"), p.randPt[2]);
            d.putDouble(stringIDToTypeID("rand3Pt"), p.randPt[3]);
            d.putDouble(stringIDToTypeID("rand4Pt"), p.randPt[4]);
            d.putDouble(stringIDToTypeID("rand5Pt"), p.randPt[5]);
        }
        app.putCustomOptions(PREF_KEY, d, true);
    } catch (_) { }
}

function getCurrentLang() {
    return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: { ja: "破線計算機", en: "Dash Calculator (Gap→Dash)" },
    panelPathInfo: { ja: "選択中のパス情報", en: "Selected Path Info" },
    pathLength: { ja: "パスの長さ:", en: "Path length:" },

    panelSplit: { ja: "破線の計算", en: "Dash Calculation" },

    segmentsLabel: { ja: "分割数:", en: "Segments:" },
    gapLabel: { ja: "間隔:", en: "Gap:" },
    dashLabel: { ja: "線分:", en: "Dash:" },

    modeGapToDash: { ja: "間隔→線分", en: "Gap→Dash" },
    modeDashToGap: { ja: "線分→間隔", en: "Dash→Gap" },
    modeRandom: { ja: "ランダム", en: "Random" },
    calcMethodPanel: { ja: "計算方法", en: "Calculation" },
    showPartialOnly: { ja: "部分表示", en: "Partial Display" },

    offsetPanel: { ja: "開始位置", en: "Offset" },
    dashSplitPanel: { ja: "部分表示", en: "Partial Display" },
    adjustEnds: { ja: "両端を調整", en: "Adjust ends" },

    capPanel: { ja: "線端", en: "Cap" },
    capNone: { ja: "なし", en: "Butt" },
    capRound: { ja: "丸型", en: "Round" },
    capProject: { ja: "突出", en: "Projecting" },
    reversePath: { ja: "パスの方向反転", en: "Reverse Path Direction" },

    btnCancel: { ja: "キャンセル", en: "Cancel" },
    btnOK: { ja: "OK", en: "OK" },
    btnClearDash: { ja: "破線クリア", en: "Clear Dashes" },

    err: { ja: "エラー", en: "Error" },

    alertNoDoc: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
    alertNoSelection: { ja: "対象となるパス（線や円など）を選択してください。", en: "Select one path (open or closed)." },
    alertNeedClosedPath: { ja: "パス（オープン/クローズ）を1つ選択してください。", en: "Select exactly one path (open or closed)." },

    alertSegmentsInvalid: { ja: "分割数は1以上の整数を入力してください。", en: "Enter an integer of 1 or greater for Segments." },
    alertGapInvalid: { ja: "間隔 (Gap) は0以上の数値を入力してください。", en: "Enter a number of 0 or greater for Gap." },
    alertDashInvalid: { ja: "線分 (Dash) は0以上の数値を入力してください。", en: "Enter a number of 0 or greater for Dash." },
    alertOffsetInvalid: { ja: "開始位置 (Offset) は数値を入力してください。", en: "Enter a number for Offset." },

    alertGapTooLongDetail: {
        ja: "間隔 (Gap) が長すぎます。線分がゼロまたはマイナスになってしまいます。\n(設定可能な最大Gap: ほぼ {0} {1})",
        en: "Gap is too long; dash would be zero or negative.\n(Max allowed Gap: about {0} {1})"
    },

    alertDashTooLongDetail: {
        ja: "線分 (Dash) が長すぎます。間隔がゼロまたはマイナスになってしまいます。\n(設定可能な最大Dash: ほぼ {0} {1})",
        en: "Dash is too long; gap would be zero or negative.\n(Max allowed Dash: about {0} {1})"
    }
};

function L(key) {
    var o = LABELS[key];
    if (!o) return key;
    return (o[lang] != null) ? o[lang] : (o.en != null ? o.en : key);
}

function LF(key, args) {
    var s = L(key);
    if (!args) return s;
    for (var i = 0; i < args.length; i++) {
        s = s.split("{" + i + "}").join(String(args[i]));
    }
    return s;
}

// --- 単位ユーティリティ（strokeUnits を参照） ---
// 単位コードとラベルのマップ
var unitLabelMap = {
    0: "in",
    1: "mm",
    2: "pt",
    3: "pica",
    4: "cm",
    5: "Q/H",
    6: "px",
    7: "ft/in",
    8: "m",
    9: "yd",
    10: "ft"
};

function getStrokeUnitInfo() {
    var unitCode = 2; // fallback pt
    try { unitCode = app.preferences.getIntegerPreference("strokeUnits"); } catch (_) { }

    var label = unitLabelMap[unitCode] || "pt";

    // Q/H は環境設定（東アジア言語）に合わせて表示を切替
    if (unitCode === 5) {
        var asianUnits = 0;
        try { asianUnits = app.preferences.getIntegerPreference("text/asianunits"); } catch (_) { }
        // 一般的に 0:Q / 1:H（取得できない場合はQ扱い）
        label = (asianUnits === 1) ? "H" : "Q";
    }

    return {
        code: unitCode,
        label: label,
        factor: unitCodeToPtFactor(unitCode)
    };
}

// 単位コード → pt換算係数（1 unit あたり何 pt か）
function unitCodeToPtFactor(unitCode) {
    switch (unitCode) {
        case 0: return 72;                 // in
        case 1: return 72 / 25.4;          // mm
        case 2: return 1;                  // pt
        case 3: return 12;                 // pica
        case 4: return 72 / 2.54;          // cm
        case 5: return (72 / 25.4) * 0.25; // Q/H（1Q=0.25mm, 1H=0.25mm）
        case 6: return 1;                  // px（Illustratorは基本的に 1px=1pt扱い）
        case 7: return 72;                 // ft/in（厳密な複合表記は扱わず、in相当として扱う）
        case 8: return 72 / 0.0254;        // m
        case 9: return 72 * 36;            // yd
        case 10: return 72 * 12;            // ft
        default: return 1;
    }
}

function unitToPt(value, unitInfo) {
    return value * unitInfo.factor;
}

function ptToUnit(pt, unitInfo) {
    return pt / unitInfo.factor;
}

function main() {
    if (app.documents.length === 0) {
        alert(L("alertNoDoc"));
        return;
    }

    var doc = app.activeDocument;
    if (doc.selection.length === 0) {
        alert(L("alertNoSelection"));
        return;
    }

    var selection = doc.selection;

    // PathItem のみ対象（複数可）
    var targets = [];
    for (var iSel = 0; iSel < selection.length; iSel++) {
        if (selection[iSel] && selection[iSel].typename === "PathItem") targets.push(selection[iSel]);
    }
    if (targets.length === 0) {
        alert(L("alertNeedClosedPath"));
        return;
    }

    // 先頭をUI/表示用の代表として扱う
    var sel = targets[0];

    // 代表の長さ（UI表示用）
    var pathLength = sel.length;

    // 線の単位（strokeUnits）
    var unitInfo = getStrokeUnitInfo();

    // ダイアログを開く前の状態を保存（キャンセルで復元）
    var originalStates = [];
    for (var si = 0; si < targets.length; si++) {
        var it0 = targets[si];
        originalStates.push({
            item: it0,
            stroked: it0.stroked,
            strokeCap: it0.strokeCap,
            strokeDashes: (it0.strokeDashes && it0.strokeDashes.length) ? it0.strokeDashes.slice(0) : [],
            strokeDashOffset: (typeof it0.strokeDashOffset === "number") ? it0.strokeDashOffset : 0
        });
    }
    var closedByOK = false;
    var directionReversed = false;

    var isDashCleared = false;

    // 前回値
    var prefs = loadPrefs();
    var initSegments = (prefs && (prefs.segments >= 1)) ? prefs.segments : 3;
    var initGapUnit = (prefs && (typeof prefs.gapPt === "number") && prefs.gapPt >= 0) ? ptToUnit(prefs.gapPt, unitInfo) : 5;
    var initOffsetUnit = (prefs && (typeof prefs.offsetPt === "number") && prefs.offsetPt >= 0) ? ptToUnit(prefs.offsetPt, unitInfo) : 0;
    var initUseOffset = (prefs && (typeof prefs.useOffset === "boolean")) ? prefs.useOffset : false;
    var initCapMode = (prefs && (typeof prefs.capMode === "number")) ? prefs.capMode : 0;
    var initReverse = (prefs && prefs.reversePath === true);
    var initAdjustEnds = (prefs && (typeof prefs.adjustEnds === "boolean")) ? prefs.adjustEnds : true;
    var initMode = (prefs && (typeof prefs.mode === "number")) ? prefs.mode : 0; // 0: Gap→Dash / 1: Dash→Gap / 2: Random
    var initDashUnit = (prefs && (typeof prefs.dashPt === "number") && prefs.dashPt >= 0) ? ptToUnit(prefs.dashPt, unitInfo) : 0;

    function fmtFieldNumber(v) {
        if (v == null || isNaN(v)) return "0";
        var rounded = Math.round(v * 1000) / 1000;
        if (Math.abs(rounded - Math.round(rounded)) < 1e-10) return String(Math.round(rounded));
        return String(rounded);
    }

    // --- Random pattern state (Dash, Gap, Dash, Gap, Dash, Gap) ---
    var randomDashesPt = null; // [d1,g1,d2,g2,d3,g3] in pt

    // Random settings (in current stroke unit) - based on Graphic Arts Unit "ランダム破線.jsx" style
    // NOTE: UI for these settings is not added yet; adjust defaults here if needed.
    var randSettings = {
        dashMin: 0,
        dashMax: 40,
        gapMin: 3,
        gapMax: 3,
        rounded: false
    };

    function ensureRandomPatternPt() {
        if (randomDashesPt && randomDashesPt.length === 6) return;

        // Try restore from prefs
        if (prefs && typeof prefs.rand0Pt === "number" && typeof prefs.rand1Pt === "number" &&
            typeof prefs.rand2Pt === "number" && typeof prefs.rand3Pt === "number") {
            // v1.7以前(4要素)の保存値も受け入れる
            if (typeof prefs.rand4Pt === "number" && typeof prefs.rand5Pt === "number") {
                randomDashesPt = [prefs.rand0Pt, prefs.rand1Pt, prefs.rand2Pt, prefs.rand3Pt, prefs.rand4Pt, prefs.rand5Pt];
            } else {
                // 4要素 → 6要素へ拡張（最後のペアはコピー）
                randomDashesPt = [prefs.rand0Pt, prefs.rand1Pt, prefs.rand2Pt, prefs.rand3Pt, prefs.rand0Pt, prefs.rand1Pt];
            }
            return;
        }

        recalcRandomPatternPt();
    }

    function getRandomUnit(min, max) {
        var rd = Math.random() * (max - min) + min;
        if (randSettings.rounded) rd = Math.round(rd);
        return rd;
    }

    function recalcRandomPatternPt() {
        // ---- Dynamic dash minimum depending on cap & unit ----
        var dashMinUnit = 0;
        var dashMaxUnit = randSettings.dashMax; // still 40

        if (rbCapNone.value) {
            // Butt cap (なし)
            switch (unitInfo.code) {
                case 2: // pt
                    dashMinUnit = 2;
                    break;
                case 1: // mm
                    dashMinUnit = 1;
                    break;
                case 5: // Q/H
                    dashMinUnit = 4;
                    break;
                default:
                    dashMinUnit = 2; // safe fallback
                    break;
            }
        } else {
            // Round / Projecting
            dashMinUnit = 0;
        }

        // Generate: dash, gap, dash, gap, dash, gap (6 entries)
        var d1u = getRandomUnit(dashMinUnit, dashMaxUnit);
        var g1u = getRandomUnit(randSettings.gapMin, randSettings.gapMax);
        var d2u = getRandomUnit(dashMinUnit, dashMaxUnit);
        var g2u = getRandomUnit(randSettings.gapMin, randSettings.gapMax);
        var d3u = getRandomUnit(dashMinUnit, dashMaxUnit);
        var g3u = getRandomUnit(randSettings.gapMin, randSettings.gapMax);

        // Safety: Illustrator dislikes NaN/negative; clamp to >= 0
        d1u = (isNaN(d1u) || d1u < 0) ? 0 : d1u;
        g1u = (isNaN(g1u) || g1u < 0) ? 0 : g1u;
        d2u = (isNaN(d2u) || d2u < 0) ? 0 : d2u;
        g2u = (isNaN(g2u) || g2u < 0) ? 0 : g2u;
        d3u = (isNaN(d3u) || d3u < 0) ? 0 : d3u;
        g3u = (isNaN(g3u) || g3u < 0) ? 0 : g3u;

        randomDashesPt = [
            unitToPt(d1u, unitInfo),
            unitToPt(g1u, unitInfo),
            unitToPt(d2u, unitInfo),
            unitToPt(g2u, unitInfo),
            unitToPt(d3u, unitInfo),
            unitToPt(g3u, unitInfo)
        ];
    }

    function getRandomPatternUnitTexts() {
        ensureRandomPatternPt();
        var d1 = ptToUnit(randomDashesPt[0], unitInfo);
        var g1 = ptToUnit(randomDashesPt[1], unitInfo);
        var d2 = ptToUnit(randomDashesPt[2], unitInfo);
        var g2 = ptToUnit(randomDashesPt[3], unitInfo);
        var d3 = ptToUnit(randomDashesPt[4], unitInfo);
        var g3 = ptToUnit(randomDashesPt[5], unitInfo);
        return {
            dashText: d1.toFixed(3) + " / " + d2.toFixed(3) + " / " + d3.toFixed(3),
            gapText: g1.toFixed(3) + " / " + g2.toFixed(3) + " / " + g3.toFixed(3)
        };
    }

    function focusGapField() {
        try { inpGap.active = true; } catch (_) { }
    }

    function getGapUnitFromUI() {
        var v = parseFloat(inpGap.text);
        if (isNaN(v) || v < 0) v = 0;
        return v;
    }

    function applyRandomGapFromUI() {
        var gapUnit = getGapUnitFromUI();
        var gPt = unitToPt(gapUnit, unitInfo);

        // randomDashesPt が未生成なら先に作る（dashはランダムで良い）
        if (!randomDashesPt || randomDashesPt.length !== 6) {
            recalcRandomPatternPt();
        }

        // gap をすべてに適用
        randomDashesPt[1] = gPt;
        randomDashesPt[3] = gPt;
        randomDashesPt[5] = gPt;
    }

    // --- ダイアログの作成 ---
    var win = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
    win.orientation = "column";
    win.alignChildren = "fill";

    // 選択中のパス情報（全幅）
    var infoPanel = win.add("panel", undefined, L("panelPathInfo"));
    infoPanel.alignChildren = "center";
    infoPanel.margins = [15, 20, 15, 10];
    var infoText = L("pathLength") + " " + ptToUnit(pathLength, unitInfo).toFixed(3) + " " + unitInfo.label;
    if (targets.length > 1) infoText += "  (" + targets.length + ")";
    infoPanel.add("statictext", undefined, infoText);

    // 2カラム（上段）
    var mainColumns = win.add("group");
    mainColumns.orientation = "row";
    mainColumns.alignChildren = ["fill", "top"];
    mainColumns.alignment = "fill";
    mainColumns.spacing = 10;

    var colLeft = mainColumns.add("group");
    colLeft.orientation = "column";
    colLeft.alignChildren = "fill";

    var colRight = mainColumns.add("group");
    colRight.orientation = "column";
    colRight.alignChildren = "fill";

    // 破線の計算（左カラム）
    var splitPanel = colLeft.add("panel", undefined, L("panelSplit"));
    splitPanel.alignChildren = "left";
    splitPanel.margins = [15, 20, 15, 10];

    // 入力項目
    var inputGroup = splitPanel.add("group");
    inputGroup.orientation = "column";
    inputGroup.alignChildren = "left";

    // ラベル幅（分割数／間隔／線分）を統一
    var LABEL_W = 40;

    // 分割数
    var grpSegments = inputGroup.add("group");
    var stSegments = grpSegments.add("statictext", undefined, L("segmentsLabel"));
    stSegments.preferredSize.width = LABEL_W;
    stSegments.justify = "right";
    var inpSegments = grpSegments.add("edittext", undefined, String(initSegments));
    inpSegments.characters = 4;
    // ↑↓キーで増減（分割数は整数・最小1）
    inpSegments._forceInteger = true;
    inpSegments._minValue = 1;
    changeValueByArrowKey(inpSegments);

    // 間隔 (Gap)
    var grpGap = inputGroup.add("group");
    var stGap = grpGap.add("statictext", undefined, L("gapLabel"));
    stGap.preferredSize.width = LABEL_W;
    stGap.justify = "right";
    var grpGapField = grpGap.add("group");
    grpGapField.orientation = "stack";

    var inpGap = grpGapField.add("edittext", undefined, fmtFieldNumber(initGapUnit));
    inpGap.characters = 4;

    var outGap = grpGapField.add("statictext", undefined, "");
    outGap.justify = "right";
    outGap.preferredSize.width = inpGap.preferredSize.width;
    outGap.visible = false;

    var stGapUnit = grpGap.add("statictext", undefined, unitInfo.label);
    // ↑↓キーで増減（間隔は0以上）
    inpGap._forceInteger = false;
    inpGap._minValue = 0;
    changeValueByArrowKey(inpGap);

    // 線分 (Dash) - 入力/出力切替
    var grpDash = inputGroup.add("group");
    var stDash = grpDash.add("statictext", undefined, L("dashLabel"));
    stDash.preferredSize.width = LABEL_W;
    stDash.justify = "right";

    var grpDashField = grpDash.add("group");
    grpDashField.orientation = "stack";

    var inpDash = grpDashField.add("edittext", undefined, fmtFieldNumber(initDashUnit));
    inpDash.characters = 4;
    // ↑↓キーで増減（線分は0以上）
    inpDash._forceInteger = false;
    inpDash._minValue = 0;
    changeValueByArrowKey(inpDash);

    // 結果（Dash）を表示するテキスト
    var outDash = grpDashField.add("statictext", undefined, "");
    outDash.justify = "right";
    outDash.preferredSize.width = inpDash.preferredSize.width;

    var stDashUnit = grpDash.add("statictext", undefined, unitInfo.label);

    // 計算方法（パスの分割パネル下部）
    var methodPanel = colLeft.add("panel", undefined, L("calcMethodPanel"));
    methodPanel.alignChildren = "left";
    methodPanel.margins = [15, 20, 18, 10];

    var modeGroup = methodPanel.add("group");
    modeGroup.orientation = "column";
    modeGroup.alignChildren = "left";

    var rbModeGapToDash = modeGroup.add("radiobutton", undefined, L("modeGapToDash"));
    var rbModeDashToGap = modeGroup.add("radiobutton", undefined, L("modeDashToGap"));
    var rbModeRandom = modeGroup.add("radiobutton", undefined, L("modeRandom"));

    rbModeGapToDash.value = (initMode === 0);
    rbModeDashToGap.value = (initMode === 1);
    rbModeRandom.value = (initMode === 2);

    function setSplitPanelDim(isDim) {
        // Random時は panel 全体を disabled にすると Gap まで触れなくなるので、
        // Segments/Dash だけをディム表示にする。
        try { grpSegments.enabled = !isDim; } catch (_) { }
        try { grpDash.enabled = !isDim; } catch (_) { }
        // Gap は常に編集可（Random時に値を採用する）
        try { grpGap.enabled = true; } catch (_) { }
    }

    function setRandomRelatedUI(isRandom) {
        // Random時は入力UIを結果表示に寄せる
        if (isRandom) {
            // Partial display is incompatible with random pattern
            if (typeof cbShowPartialOnly !== "undefined" && cbShowPartialOnly) {
                cbShowPartialOnly.value = false;
                cbShowPartialOnly.enabled = false;
            }

            // Adjust ends is irrelevant for random pattern
            if (typeof cbAdjustEnds !== "undefined" && cbAdjustEnds) {
                cbAdjustEnds.enabled = false;
            }

            inpDash.visible = false;
            outDash.visible = true;

            // Gap は編集可能にする
            inpGap.visible = true;
            outGap.visible = false;

            setSplitPanelDim(true);

            // 現在の Gap 入力値を Random パターンに反映
            applyRandomGapFromUI();

            var t = getRandomPatternUnitTexts();
            outDash.text = t.dashText;

            focusGapField();
        } else {
            if (typeof cbShowPartialOnly !== "undefined" && cbShowPartialOnly) {
                cbShowPartialOnly.enabled = true;
            }
            if (typeof cbAdjustEnds !== "undefined" && cbAdjustEnds) {
                cbAdjustEnds.enabled = !sel.closed;
            }

            setSplitPanelDim(false);
        }
    }

    // 初期モードに合わせて表示を切替
    (function () {
        var isRandom = rbModeRandom.value;
        if (isRandom) {
            ensureRandomPatternPt();

            inpDash.visible = false;
            outDash.visible = true;

            inpGap.visible = true;
            outGap.visible = false;

            setSplitPanelDim(true);

            applyRandomGapFromUI();

            var t0 = getRandomPatternUnitTexts();
            outDash.text = t0.dashText;

            focusGapField();
            return;
        }

        var dashToGap = rbModeDashToGap.value;
        inpDash.visible = dashToGap;
        outDash.visible = !dashToGap;
        inpGap.visible = !dashToGap;
        outGap.visible = dashToGap;

        setRandomRelatedUI(false);
    })();

    // 開始位置パネル
    var offsetPanel = colRight.add("panel", undefined, L("offsetPanel"));
    offsetPanel.alignChildren = "left";
    offsetPanel.margins = [15, 20, 15, 10];

    // 開始位置（入力）
    var grpOffset = offsetPanel.add("group");
    grpOffset.orientation = "row";
    grpOffset.alignChildren = ["left", "center"];

    var cbUseOffset = grpOffset.add("checkbox", undefined, "");
    cbUseOffset.value = initUseOffset;
    cbUseOffset.preferredSize.width = 18;

    var inpOffset = grpOffset.add("edittext", undefined, fmtFieldNumber(initOffsetUnit));
    inpOffset.characters = 4;
    inpOffset.enabled = cbUseOffset.value;

    var stOffsetUnit = grpOffset.add("statictext", undefined, unitInfo.label);
    stOffsetUnit.enabled = cbUseOffset.value;

    // ↑↓キーで増減（開始位置は0以上）
    inpOffset._forceInteger = false;
    inpOffset._minValue = 0;
    changeValueByArrowKey(inpOffset);

    // 1周期（= dash+gap）基準のプリセット
    var offsetPresetGroup = offsetPanel.add("group");
    offsetPresetGroup.orientation = "row";
    offsetPresetGroup.alignChildren = ["left", "center"];
    offsetPresetGroup.enabled = cbUseOffset.value;

    var rbOffQ1 = offsetPresetGroup.add("radiobutton", undefined, "1/4");
    var rbOffQ2 = offsetPresetGroup.add("radiobutton", undefined, "1/2");
    var rbOffQ3 = offsetPresetGroup.add("radiobutton", undefined, "3/4");


    // 部分表示（開始位置パネルの下）
    var dashSplitPanel = colRight.add("panel", undefined, L("dashSplitPanel"));
    dashSplitPanel.alignChildren = "left";
    dashSplitPanel.margins = [15, 20, 15, 10];

    // 部分表示（部分表示パネル）
    var cbShowPartialOnly = dashSplitPanel.add("checkbox", undefined, L("showPartialOnly"));
    cbShowPartialOnly.value = false;


    // 線端
    var capPanel = colRight.add("panel", undefined, L("capPanel"));
    capPanel.alignChildren = "left";
    capPanel.margins = [15, 20, 15, 10];

    var capGroup = capPanel.add("group");
    capGroup.orientation = "row";
    capGroup.alignChildren = ["left", "center"];

    var rbCapNone = capGroup.add("radiobutton", undefined, L("capNone"));
    var rbCapRound = capGroup.add("radiobutton", undefined, L("capRound"));
    var rbCapProject = capGroup.add("radiobutton", undefined, L("capProject"));

    // デフォルト（前回値）
    if (initCapMode === 1) rbCapRound.value = true;
    else if (initCapMode === 2) rbCapProject.value = true;
    else rbCapNone.value = true;

    // 線端パネルの下：パスの方向反転（中央）
    var grpReversePath = win.add("group");
    grpReversePath.orientation = "row";
    grpReversePath.alignment = "center";
    grpReversePath.spacing = 20;

    // 両端を調整（オープンパス向け）
    var cbAdjustEnds = grpReversePath.add("checkbox", undefined, L("adjustEnds"));
    cbAdjustEnds.value = initAdjustEnds;
    // クローズパスでは意味がないためディム表示
    if (sel.closed) cbAdjustEnds.enabled = false;

    var cbReversePath = grpReversePath.add("checkbox", undefined, L("reversePath"));
    cbReversePath.value = initReverse;

    // ボタン（3カラム）
    var btnArea = win.add("group");
    btnArea.orientation = "row";
    btnArea.alignChildren = ["fill", "center"];
    btnArea.alignment = "fill";

    // 左：破線クリア
    var btnLeft = btnArea.add("group");
    btnLeft.orientation = "row";
    btnLeft.alignChildren = ["left", "center"];
    btnLeft.alignment = "left";
    var btnClearDash = btnLeft.add("button", undefined, L("btnClearDash"));

    // 中央：スペーサー（伸びる）
    var btnSpacer = btnArea.add("group");
    btnSpacer.alignment = ["fill", "fill"];
    btnSpacer.minimumSize.width = 0;

    // 右：キャンセル / OK
    var btnRight = btnArea.add("group");
    btnRight.orientation = "row";
    btnRight.alignChildren = ["right", "center"];
    btnRight.alignment = "right";
    var btnCancel = btnRight.add("button", undefined, L("btnCancel"), { name: "cancel" });
    var btnOK = btnRight.add("button", undefined, L("btnOK"), { name: "ok" });

    // 破線クリア（プレビュー）
    btnClearDash.onClick = function () {
        isDashCleared = true;
        updatePreview();
    };

    // キャンセル：ダイアログを開く前の状態に戻して閉じる
    btnCancel.onClick = function () {
        restoreOriginalState();
        win.close(0);
    };

    // ×ボタン / Esc などで閉じた場合も、OK以外は復元
    win.onClose = function () {
        if (!closedByOK) {
            restoreOriginalState();
        }
        return true;
    };

    // ↑↓キー / Shift+↑↓ / Option(Alt)+↑↓ で数値を増減
    // - ↑↓: ±1
    // - Shift+↑↓: ±10（10の倍数にスナップ）
    // - Option(Alt)+↑↓: ±0.1（小数第1位まで）
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;

            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var forceInt = !!editText._forceInteger;
            var minValue = (typeof editText._minValue === "number") ? editText._minValue : 0;
            var useDecimal = keyboard.altKey && !forceInt;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                if (event.keyName == "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                } else if (event.keyName == "Down") {
                    value = Math.floor((value - 1) / delta) * delta;
                }
                event.preventDefault();
            } else if (useDecimal) {
                delta = 0.1;
                if (event.keyName == "Up") {
                    value += delta;
                } else if (event.keyName == "Down") {
                    value -= delta;
                }
                event.preventDefault();
            } else {
                delta = 1;
                if (event.keyName == "Up") {
                    value += delta;
                } else if (event.keyName == "Down") {
                    value -= delta;
                }
                event.preventDefault();
            }

            if (useDecimal) {
                value = Math.round(value * 10) / 10; // 小数第1位まで
                if (Math.abs(value) < 0.0000001) value = 0; // -0 対策
            } else {
                value = Math.round(value); // 整数
            }

            if (value < minValue) value = minValue;

            editText.text = String(value);

            // プレビュー更新（必要な場合）
            try {
                if (typeof editText._onArrowChange === "function") {
                    editText._onArrowChange();
                }
            } catch (_) { }
        });
    }

    // 破線計算（pt）
    // - クローズ: 1周期(=dash+gap)=全長/分割数
    // - オープン: 両端をダッシュで終える（分割数=ダッシュ本数、gapは(分割数-1)回）
    function calcDashAndStepPt(segments, gapPt, pathLen, isClosed) {
        if (!(segments > 0)) return null;
        if (!(gapPt >= 0)) return null;

        if (isClosed) {
            var stepPt = pathLen / segments;
            var dashPt = stepPt - gapPt;
            return { dashPt: dashPt, stepPt: stepPt };
        }

        // open path
        // 両端を調整: 両端をダッシュで終える（分割数=ダッシュ本数、gapは(分割数-1)回）
        // OFF: クローズと同じ計算（1周期=全長/分割数、末尾は端数になり得る）
        if (!cbAdjustEnds.value) {
            var stepPtOpen = pathLen / segments;
            var dashPtOpen = stepPtOpen - gapPt;
            return { dashPt: dashPtOpen, stepPt: stepPtOpen };
        }

        if (segments === 1) {
            var dash1 = pathLen;
            return { dashPt: dash1, stepPt: dash1 + gapPt };
        }

        var dash = (pathLen - gapPt * (segments - 1)) / segments;
        return { dashPt: dash, stepPt: dash + gapPt };
    }

    // 線分(Dash)から間隔(Gap)を逆算（pt）
    function calcGapPtFromDashPt(segments, dashPt, pathLen, isClosed) {
        if (!(segments > 0)) return null;
        if (!(dashPt >= 0)) return null;

        // クローズ、または「両端を調整」OFFは 1周期=全長/分割数 で扱う
        if (isClosed || !cbAdjustEnds.value) {
            var stepPt = pathLen / segments;
            return stepPt - dashPt;
        }

        // open + 両端を調整ON
        if (segments === 1) return 0;
        return (pathLen - dashPt * segments) / (segments - 1);
    }

    // 1周期（= dash+gap）を unit にした値を返す
    function getStepUnit() {
        var segments = parseInt(inpSegments.text, 10);
        if (isNaN(segments) || segments <= 0) return null;

        // Random: step is sum of (dash,gap,dash,gap,dash,gap)
        if (rbModeRandom.value) {
            ensureRandomPatternPt();
            applyRandomGapFromUI();
            var stepPtRand = randomDashesPt[0] + randomDashesPt[1] + randomDashesPt[2] + randomDashesPt[3] + randomDashesPt[4] + randomDashesPt[5];
            return ptToUnit(stepPtRand, unitInfo);
        }

        // 一部だけを表示する：間隔0として 1周期を計算
        if (cbShowPartialOnly.value) {
            var r0 = calcDashAndStepPt(segments, 0, pathLength, sel.closed);
            if (!r0) return null;
            return ptToUnit(r0.stepPt, unitInfo);
        }

        // Dash→Gap
        if (rbModeDashToGap.value) {
            var dashUnit = parseFloat(inpDash.text);
            if (isNaN(dashUnit) || dashUnit < 0) return null;

            var dashPt = unitToPt(dashUnit, unitInfo);

            // open + 両端を調整ON + 1本は全長ダッシュ
            if (!sel.closed && cbAdjustEnds.value && segments === 1) {
                dashPt = pathLength;
            }

            var stepPt;
            if (sel.closed || !cbAdjustEnds.value) {
                stepPt = pathLength / segments;
            } else {
                var gapPt = calcGapPtFromDashPt(segments, dashPt, pathLength, sel.closed);
                if (gapPt == null) return null;
                stepPt = dashPt + gapPt;
            }

            return ptToUnit(stepPt, unitInfo);
        }

        // Gap→Dash
        var gapUnit = parseFloat(inpGap.text);
        if (isNaN(gapUnit) || gapUnit < 0) return null;

        var gapPt = unitToPt(gapUnit, unitInfo);
        var r = calcDashAndStepPt(segments, gapPt, pathLength, sel.closed);
        if (!r) return null;
        return ptToUnit(r.stepPt, unitInfo);
    }

    // 開始位置（unit）を入力欄に反映（見た目を整える）
    function setOffsetText(unitVal) {
        if (unitVal == null || isNaN(unitVal) || unitVal < 0) unitVal = 0;

        // 入力欄に収まりやすいように小数第3位まで
        var rounded = Math.round(unitVal * 1000) / 1000;
        if (Math.abs(rounded - Math.round(rounded)) < 1e-10) {
            inpOffset.text = String(Math.round(rounded));
        } else {
            inpOffset.text = String(rounded);
        }
    }

    // プリセット（1/4 / 1/2 / 3/4）と入力値を同期
    function syncOffsetPreset(offsetUnit) {
        // 一旦すべて解除（カスタム値もあり得る）
        rbOffQ1.value = rbOffQ2.value = rbOffQ3.value = false;

        var stepUnit = getStepUnit();
        if (stepUnit == null) return;

        // 単位によって誤差が出るので、ゆるめの許容
        var tol = Math.max(0.001, Math.abs(stepUnit) * 0.0005);

        var t1 = stepUnit * 0.25;
        var t2 = stepUnit * 0.50;
        var t3 = stepUnit * 0.75;

        if (Math.abs(offsetUnit - t1) <= tol) rbOffQ1.value = true;
        else if (Math.abs(offsetUnit - t2) <= tol) rbOffQ2.value = true;
        else if (Math.abs(offsetUnit - t3) <= tol) rbOffQ3.value = true;
    }

    // プリセットを適用（1周期基準）
    function applyOffsetPreset(frac) {
        var stepUnit = getStepUnit();
        if (stepUnit == null) return;
        setOffsetText(stepUnit * frac);
        updatePreviewUnclear();
    }


    // プリセット（1周期基準）
    rbOffQ1.onClick = function () { applyOffsetPreset(0.25); };
    rbOffQ2.onClick = function () { applyOffsetPreset(0.50); };
    rbOffQ3.onClick = function () { applyOffsetPreset(0.75); };


    function forEachTarget(fn) {
        for (var i = 0; i < targets.length; i++) {
            try { fn(targets[i]); } catch (_) { }
        }
    }

    function getPathLengthOf(item) {
        try { return item.length; } catch (_) { return 0; }
    }

    function setStrokeCapTo(item) {
        if (rbCapRound.value) item.strokeCap = StrokeCap.ROUNDENDCAP;
        else if (rbCapProject.value) item.strokeCap = StrokeCap.PROJECTINGENDCAP;
        else item.strokeCap = StrokeCap.BUTTENDCAP;
    }

    // 値変更時にプレビュー更新（アラート無し）
    function updatePreview() {
        // 破線クリア状態（プレビュー）
        if (isDashCleared) {
            outDash.text = "";
            outGap.text = "";
            forEachTarget(function (item) {
                item.stroked = true;
                setStrokeCapTo(item);
                item.strokeDashOffset = 0;
                item.strokeDashes = [];
            });
            app.redraw();
            return;
        }
        var offsetUnit = 0;
        if (cbUseOffset.value) {
            offsetUnit = parseFloat(inpOffset.text);
            if (isNaN(offsetUnit) || offsetUnit < 0) {
                outDash.text = "";
                outGap.text = "";
                return;
            }
        }

        var segments = parseInt(inpSegments.text, 10);
        if (isNaN(segments) || segments <= 0) {
            outDash.text = "";
            outGap.text = "";
            return;
        }

        // Random mode: set Dash/Gap/Dash/Gap pattern (per-path randomization)
        if (rbModeRandom.value) {
            var offsetPtRand = unitToPt(offsetUnit, unitInfo);

            // UI表示用に1回だけ生成（代表表示）
            recalcRandomPatternPt();
            applyRandomGapFromUI();
            var tt = getRandomPatternUnitTexts();
            outDash.text = tt.dashText;

            // 各パスごとに別乱数で適用
            for (var iR = 0; iR < targets.length; iR++) {
                var itemR = targets[iR];

                // パスごとに新しい乱数を生成
                recalcRandomPatternPt();
                applyRandomGapFromUI();

                itemR.stroked = true;
                if (rbCapRound.value) itemR.strokeCap = StrokeCap.ROUNDENDCAP;
                else if (rbCapProject.value) itemR.strokeCap = StrokeCap.PROJECTINGENDCAP;
                else itemR.strokeCap = StrokeCap.BUTTENDCAP;

                itemR.strokeDashOffset = offsetPtRand;
                itemR.strokeDashes = randomDashesPt.slice(0);
            }

            app.redraw();
            return;
        }

        var dashToGap = rbModeDashToGap.value;

        var gapPt, dashPt;

        // 一部だけを表示する：間隔0として計算（入力値は使わない）
        if (cbShowPartialOnly.value) {
            gapPt = 0;
            var r0 = calcDashAndStepPt(segments, 0, pathLength, sel.closed);
            dashPt = r0 ? r0.dashPt : -1;
            if (dashPt <= 0) {
                outDash.text = L("err");
                outGap.text = "";
                return;
            }
            // 表示（Dash / Gap=0）
            outDash.text = ptToUnit(dashPt, unitInfo).toFixed(3);
            outGap.text = ptToUnit(0, unitInfo).toFixed(3);
            // hidden field sync
            inpGap.text = "0";
            inpDash.text = fmtFieldNumber(ptToUnit(dashPt, unitInfo));
        } else if (dashToGap) {
            var dashUnit = parseFloat(inpDash.text);
            if (isNaN(dashUnit) || dashUnit < 0) {
                outGap.text = "";
                return;
            }
            dashPt = unitToPt(dashUnit, unitInfo);
            // open + 両端を調整ON + 1本は全長ダッシュ
            if (!sel.closed && cbAdjustEnds.value && segments === 1) {
                dashPt = pathLength;
                inpDash.text = fmtFieldNumber(ptToUnit(dashPt, unitInfo));
            }
            gapPt = calcGapPtFromDashPt(segments, dashPt, pathLength, sel.closed);
            if (gapPt == null || gapPt < 0) {
                outGap.text = L("err");
                return;
            }
            // 表示（Gap）
            outGap.text = ptToUnit(gapPt, unitInfo).toFixed(3);
            // hidden field sync（切替に備える）
            inpGap.text = fmtFieldNumber(ptToUnit(gapPt, unitInfo));
            // outDash も同期（表示切替に備える）
            outDash.text = ptToUnit(dashPt, unitInfo).toFixed(3);
        } else {
            var gapUnit = parseFloat(inpGap.text);
            if (isNaN(gapUnit) || gapUnit < 0) {
                outDash.text = "";
                return;
            }
            gapPt = unitToPt(gapUnit, unitInfo);
            var r = calcDashAndStepPt(segments, gapPt, pathLength, sel.closed);
            dashPt = r ? r.dashPt : -1;
            if (dashPt <= 0) {
                outDash.text = L("err");
                return;
            }
            // 表示（Dash）
            outDash.text = ptToUnit(dashPt, unitInfo).toFixed(3);
            // hidden field sync（切替に備える）
            inpDash.text = fmtFieldNumber(ptToUnit(dashPt, unitInfo));
            // outGap も同期（表示切替に備える）
            outGap.text = ptToUnit(gapPt, unitInfo).toFixed(3);
        }

        // プリセット表示を同期（手入力/分割数変更に追従）
        syncOffsetPreset(offsetUnit);

        // ---- Apply to each target (per-path) ----
        var offsetPtEach = unitToPt(offsetUnit, unitInfo);

        if (cbShowPartialOnly.value) {
            forEachTarget(function (item) {
                var len = getPathLengthOf(item);
                var rPO = calcDashAndStepPt(segments, 0, len, item.closed);
                if (!rPO || !(rPO.dashPt > 0)) return;
                var C0 = rPO.dashPt;
                var D0 = C0 / 2;
                var B0 = len + D0;
                item.stroked = true;
                setStrokeCapTo(item);
                item.strokeDashOffset = offsetPtEach;
                item.strokeDashes = [0, 0, C0, B0];
            });
        } else if (dashToGap) {
            var dashUnitIn = parseFloat(inpDash.text);
            if (!isNaN(dashUnitIn) && dashUnitIn >= 0) {
                var dashPtIn = unitToPt(dashUnitIn, unitInfo);
                forEachTarget(function (item) {
                    var len = getPathLengthOf(item);
                    var dPt = dashPtIn;
                    if (!item.closed && cbAdjustEnds.value && segments === 1) dPt = len;
                    var gPt = calcGapPtFromDashPt(segments, dPt, len, item.closed);
                    if (gPt == null || gPt < 0) return;
                    item.stroked = true;
                    setStrokeCapTo(item);
                    item.strokeDashOffset = offsetPtEach;
                    item.strokeDashes = [dPt, gPt];
                });
            }
        } else {
            var gapUnitIn = parseFloat(inpGap.text);
            if (!isNaN(gapUnitIn) && gapUnitIn >= 0) {
                var gapPtIn = unitToPt(gapUnitIn, unitInfo);
                forEachTarget(function (item) {
                    var len = getPathLengthOf(item);
                    var rGD = calcDashAndStepPt(segments, gapPtIn, len, item.closed);
                    if (!rGD || !(rGD.dashPt > 0)) return;
                    var dPt2 = rGD.dashPt;
                    item.stroked = true;
                    setStrokeCapTo(item);
                    item.strokeDashOffset = offsetPtEach;
                    item.strokeDashes = [dPt2, gapPtIn];
                });
            }
        }

        app.redraw();
    }

    function setReversePath(shouldReverse) {
        shouldReverse = !!shouldReverse;
        if (shouldReverse === directionReversed) return;

        try {
            // Reverse Path Direction は選択に対して実行されるため、対象を選択してから実行
            doc.selection = targets;
            app.executeMenuCommand('Reverse Path Direction');
            directionReversed = shouldReverse;
        } catch (_) {
            // 失敗した場合はフラグを変更しない
        }
    }

    function restoreOriginalState() {
        // パス方向を元に戻す（ダイアログ内で反転していた場合）
        if (directionReversed) {
            try {
                doc.selection = targets;
                app.executeMenuCommand('Reverse Path Direction');
            } catch (_) { }
            directionReversed = false;
        }

        // ストローク状態の復元（まとめて）
        for (var i = 0; i < originalStates.length; i++) {
            var s0 = originalStates[i];
            try {
                s0.item.strokeDashes = s0.strokeDashes.slice(0);
                s0.item.strokeDashOffset = s0.strokeDashOffset;
                s0.item.strokeCap = s0.strokeCap;
                s0.item.stroked = s0.stroked;
            } catch (_) { }
        }

        try { app.redraw(); } catch (_) { }
    }

    // ↑↓キーによる値変更でもプレビューを更新
    inpSegments._onArrowChange = updatePreviewUnclear;
    inpGap._onArrowChange = updatePreviewUnclear;
    inpDash._onArrowChange = updatePreviewUnclear;
    inpOffset._onArrowChange = updatePreviewUnclear;

    // 入力値や線端の変更でプレビュー更新
    inpSegments.onChanging = updatePreviewUnclear;
    inpGap.onChanging = updatePreviewUnclear;
    inpDash.onChanging = updatePreviewUnclear;
    inpOffset.onChanging = updatePreviewUnclear;
    inpSegments.onChange = updatePreviewUnclear;
    inpGap.onChange = updatePreviewUnclear;
    inpDash.onChange = updatePreviewUnclear;
    inpOffset.onChange = updatePreviewUnclear;
    rbCapNone.onClick = updatePreview;
    rbCapRound.onClick = updatePreview;
    rbCapProject.onClick = updatePreview;
    cbReversePath.onClick = function () { setReversePath(cbReversePath.value); updatePreview(); };
    cbAdjustEnds.onClick = updatePreviewUnclear;
    rbModeGapToDash.onClick = function () {
        inpDash.visible = false;
        outDash.visible = true;
        inpGap.visible = true;
        outGap.visible = false;
        rbModeRandom.value = false;
        setRandomRelatedUI(false);
        updatePreviewUnclear();
    };
    rbModeDashToGap.onClick = function () {
        inpDash.visible = true;
        outDash.visible = false;
        inpGap.visible = false;
        outGap.visible = true;
        rbModeRandom.value = false;
        setRandomRelatedUI(false);
        updatePreviewUnclear();
    };
    rbModeRandom.onClick = function () {
        // クリックのたびに再計算
        rbModeGapToDash.value = false;
        rbModeDashToGap.value = false;
        rbModeRandom.value = true;

        recalcRandomPatternPt();
        applyRandomGapFromUI();
        setRandomRelatedUI(true);
        updatePreviewUnclear();
    };
    cbUseOffset.onClick = function () {
        inpOffset.enabled = cbUseOffset.value;
        stOffsetUnit.enabled = cbUseOffset.value;
        offsetPresetGroup.enabled = cbUseOffset.value;
        updatePreviewUnclear();
    };

    cbShowPartialOnly.onClick = function () {
        if (cbShowPartialOnly.value) {
            // ONにしたら間隔を0へ
            inpGap.text = "0";
        }
        updatePreviewUnclear();
    };

    // --- updatePreviewUnclear をトップレベルに分離 ---
    function updatePreviewUnclear() {
        if (isDashCleared) isDashCleared = false;
        updatePreview();
    }

    // 前回値（パス方向）を反映
    setReversePath(cbReversePath.value);

    // 初期表示で Random の場合はUIを整える
    if (rbModeRandom.value) {
        ensureRandomPatternPt();
        setRandomRelatedUI(true);
    }

    // 初期値でも一度プレビュー
    updatePreview();

    // OKボタンの処理（プレビューを確定して閉じる）
    btnOK.onClick = function () {
        // 破線クリアを確定
        if (isDashCleared) {
            try {
                forEachTarget(function (item) {
                    item.stroked = true;
                    item.strokeDashOffset = 0;
                    item.strokeDashes = [];
                });
                app.redraw();
            } catch (_) { }

            // 前回値として保存（他のUI状態は残す）
            try {
                var segmentsSave = parseInt(inpSegments.text, 10);
                if (isNaN(segmentsSave) || segmentsSave <= 0) segmentsSave = initSegments;
                var capModeSave = rbCapRound.value ? 1 : (rbCapProject.value ? 2 : 0);

                savePrefs({
                    segments: segmentsSave,
                    gapPt: (prefs && typeof prefs.gapPt === "number") ? prefs.gapPt : 0,
                    dashPt: (prefs && typeof prefs.dashPt === "number") ? prefs.dashPt : 0,
                    offsetPt: 0,
                    capMode: capModeSave,
                    reversePath: cbReversePath.value,
                    adjustEnds: cbAdjustEnds.value,
                    mode: rbModeRandom.value ? 2 : (rbModeDashToGap.value ? 1 : 0),
                    randPt: (rbModeRandom.value ? (randomDashesPt ? randomDashesPt.slice(0) : null) : null),
                    useOffset: cbUseOffset.value
                });
            } catch (_) { }

            closedByOK = true;
            win.close(1);
            return;
        }
        var segments = parseInt(inpSegments.text, 10);
        var offsetUnit = 0;
        if (isNaN(segments) || segments <= 0) {
            alert(L("alertSegmentsInvalid"));
            return;
        }
        if (cbUseOffset.value) {
            offsetUnit = parseFloat(inpOffset.text);
            if (isNaN(offsetUnit) || offsetUnit < 0) {
                alert(L("alertOffsetInvalid"));
                return;
            }
        }

        // 一部だけを表示する：間隔0として確定（入力値は使わない）
        if (cbShowPartialOnly.value) {
            var r0 = calcDashAndStepPt(segments, 0, pathLength, sel.closed);
            var dashPt0 = r0 ? r0.dashPt : -1;
            if (dashPt0 <= 0) {
                alert(L("err"));
                return;
            }

            // UI同期
            inpGap.text = "0";
            inpDash.text = fmtFieldNumber(ptToUnit(dashPt0, unitInfo));

            // 選択状態に適用（＝プレビュー更新と同じ処理）
            updatePreview();

            // 保存（gapPt=0固定）
            try {
                var offsetPtSave0 = unitToPt(offsetUnit, unitInfo);
                var capModeSave0 = rbCapRound.value ? 1 : (rbCapProject.value ? 2 : 0);
                savePrefs({
                    segments: segments,
                    gapPt: 0,
                    dashPt: dashPt0,
                    offsetPt: offsetPtSave0,
                    capMode: capModeSave0,
                    reversePath: cbReversePath.value,
                    adjustEnds: cbAdjustEnds.value,
                    mode: rbModeRandom.value ? 2 : (rbModeDashToGap.value ? 1 : 0),
                    randPt: (rbModeRandom.value ? (randomDashesPt ? randomDashesPt.slice(0) : null) : null),
                    useOffset: cbUseOffset.value
                });
            } catch (_) { }

            closedByOK = true;
            win.close(1);
            return;
        }

        var dashToGap = rbModeDashToGap.value;
        var gapPt, dashPt;

        if (dashToGap) {
            var dashUnit = parseFloat(inpDash.text);
            if (isNaN(dashUnit) || dashUnit < 0) {
                alert(L("alertDashInvalid"));
                return;
            }
            dashPt = unitToPt(dashUnit, unitInfo);
            // open + 両端を調整ON + 1本は全長ダッシュ
            if (!sel.closed && cbAdjustEnds.value && segments === 1) {
                dashPt = pathLength;
                inpDash.text = fmtFieldNumber(ptToUnit(dashPt, unitInfo));
            }
            gapPt = calcGapPtFromDashPt(segments, dashPt, pathLength, sel.closed);
            if (gapPt == null || gapPt < 0) {
                var maxDashPt = pathLength / segments;
                var maxDashUnit = ptToUnit(maxDashPt, unitInfo);
                alert(LF("alertDashTooLongDetail", [maxDashUnit.toFixed(3), unitInfo.label]));
                return;
            }
            // hidden field sync（保存/切替に備える）
            inpGap.text = fmtFieldNumber(ptToUnit(gapPt, unitInfo));
        } else {
            var gapUnit = parseFloat(inpGap.text);
            if (isNaN(gapUnit) || gapUnit < 0) {
                alert(L("alertGapInvalid"));
                return;
            }
            gapPt = unitToPt(gapUnit, unitInfo);
            var r = calcDashAndStepPt(segments, gapPt, pathLength, sel.closed);
            dashPt = r ? r.dashPt : -1;
            if (dashPt <= 0) {
                // 最大Gap（概算）
                var maxGapPt;
                if (sel.closed || !cbAdjustEnds.value) {
                    maxGapPt = (pathLength / segments);
                } else {
                    maxGapPt = (segments <= 1) ? (pathLength) : (pathLength / (segments - 1));
                }
                var maxGapUnit = ptToUnit(maxGapPt, unitInfo);
                alert(LF("alertGapTooLongDetail", [maxGapUnit.toFixed(3), unitInfo.label]));
                outDash.text = L("err");
                return;
            }
            // hidden field sync（保存/切替に備える）
            inpDash.text = fmtFieldNumber(ptToUnit(dashPt, unitInfo));
        }

        // 選択状態に適用（＝プレビュー更新と同じ処理）
        updatePreview();

        // 前回値として保存（ptで保持）
        try {
            var offsetPtSave = unitToPt(offsetUnit, unitInfo);
            var capModeSave = rbCapRound.value ? 1 : (rbCapProject.value ? 2 : 0);
            savePrefs({
                segments: segments,
                gapPt: gapPt,
                dashPt: dashPt,
                offsetPt: offsetPtSave,
                capMode: capModeSave,
                reversePath: cbReversePath.value,
                adjustEnds: cbAdjustEnds.value,
                mode: rbModeRandom.value ? 2 : (dashToGap ? 1 : 0),
                randPt: (rbModeRandom.value ? (randomDashesPt ? randomDashesPt.slice(0) : null) : null),
                useOffset: cbUseOffset.value
            });
        } catch (_) { }

        // プレビューを確定してダイアログを閉じる
        closedByOK = true;
        win.close(1);
    };

    // --- フォーカスヘルパー ---
    function focusSegmentsField() {
        // 分割数のテキストフィールドをアクティブに
        inpSegments.active = true;
    }

    // 表示前の保険
    focusSegmentsField();

    // ダイアログ表示時：分割数をアクティブに
    win.onShow = function () {
        focusSegmentsField();
    };

    win.show();
}

main();