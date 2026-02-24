#target illustrator
try { app.preferences.setBooleanPreference('ShowExternalJSXWarning', false); } catch (_) { }

/*
 * 破線計算機（Gap→Dash）
 * 更新日: 2026-02-25
 * Version: v1.4
 * 概要: パス（オープン/クローズ）1つを選択し、分割数と間隔から線分長を算出して破線を適用します。
 *      「開始位置（オフセット）」で破線の開始位置（位相）も指定できます。
 */


var SCRIPT_VERSION = "v1.4";

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

        if (d.hasKey(kSegments)) p.segments = d.getInteger(kSegments);
        if (d.hasKey(kGapPt)) p.gapPt = d.getDouble(kGapPt);
        if (d.hasKey(kOffsetPt)) p.offsetPt = d.getDouble(kOffsetPt);
        if (d.hasKey(kCapMode)) p.capMode = d.getInteger(kCapMode);
        if (d.hasKey(kReverse)) p.reversePath = d.getBoolean(kReverse);
        if (d.hasKey(kAdjustEnds)) p.adjustEnds = d.getBoolean(kAdjustEnds);
        if (d.hasKey(kDashPt)) p.dashPt = d.getDouble(kDashPt);
        if (d.hasKey(kMode)) p.mode = d.getInteger(kMode);

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

    segmentsLabel: { ja: "分割数:", en: "Segments:" },
    gapLabel: { ja: "間隔:", en: "Gap:" },
    dashLabel: { ja: "線分:", en: "Dash:" },

    calcModeLabel: { ja: "計算:", en: "Mode:" },
    modeGapToDash: { ja: "間隔→線分", en: "Gap→Dash" },
    modeDashToGap: { ja: "線分→間隔", en: "Dash→Gap" },
    calcMethodPanel: { ja: "計算方法", en: "Calculation" },

    // offsetLabel: { ja: "開始位置:", en: "Offset:" },
    offsetPanel: { ja: "開始位置", en: "Offset" },
    adjustEnds: { ja: "両端を調整", en: "Adjust ends" },

    pasteApply: { ja: "反映", en: "Apply" },

    capPanel: { ja: "線端", en: "Cap" },
    capNone: { ja: "なし", en: "Butt" },
    capRound: { ja: "丸型", en: "Round" },
    capProject: { ja: "突出", en: "Projecting" },
    reversePath: { ja: "パスの方向反転", en: "Reverse Path Direction" },

    btnCancel: { ja: "キャンセル", en: "Cancel" },
    btnOK: { ja: "OK", en: "OK" },

    err: { ja: "エラー", en: "Error" },

    alertNoDoc: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
    alertNoSelection: { ja: "対象となるパス（線や円など）を選択してください。", en: "Select one path (open or closed)." },
    alertNeedClosedPath: { ja: "パス（オープン/クローズ）を1つ選択してください。", en: "Select exactly one path (open or closed)." },

    alertSegmentsInvalid: { ja: "分割数は1以上の整数を入力してください。", en: "Enter an integer of 1 or greater for Segments." },
    alertGapInvalid: { ja: "間隔 (Gap) は0以上の数値を入力してください。", en: "Enter a number of 0 or greater for Gap." },
    alertDashInvalid: { ja: "線分 (Dash) は0以上の数値を入力してください。", en: "Enter a number of 0 or greater for Dash." },
    alertOffsetInvalid: { ja: "開始位置 (Offset) は数値を入力してください。", en: "Enter a number for Offset." },

    alertPasteInvalid: { ja: "貼り付け形式が不正です。例: 0,15,20,13", en: "Invalid paste format. Example: 0,15,20,13" },

    alertGapTooLongDetail: {
        ja: "間隔 (Gap) が長すぎます。線分がゼロまたはマイナスになってしまいます。\n(設定可能な最大Gap: ほぼ {0} {1})",
        en: "Gap is too long; dash would be zero or negative.\n(Max allowed Gap: about {0} {1})"
    }
    ,
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

    var sel = doc.selection[0];

    // パスアイテムであるか確認（オープン/クローズどちらもOK）
    if (sel.typename !== "PathItem") {
        alert(L("alertNeedClosedPath"));
        return;
    }

    // パスの長さを取得（円周など）
    var pathLength = sel.length;

    // 線の単位（strokeUnits）
    var unitInfo = getStrokeUnitInfo();

    // ダイアログを開く前の状態を保存（キャンセルで復元）
    var originalState = {
        stroked: sel.stroked,
        strokeCap: sel.strokeCap,
        strokeDashes: (sel.strokeDashes && sel.strokeDashes.length) ? sel.strokeDashes.slice(0) : [],
        strokeDashOffset: (typeof sel.strokeDashOffset === "number") ? sel.strokeDashOffset : 0
    };
    var closedByOK = false;
    var directionReversed = false;

    // 貼付モード（dash array を直接適用）
    var pasteMode = false;
    var pasteDashPts = null; // pt配列
    var pasteDisplayIndex = 0; // UI表示に使う先頭インデックス

    // 前回値
    var prefs = loadPrefs();
    var initSegments = (prefs && (prefs.segments >= 1)) ? prefs.segments : 3;
    var initGapUnit = (prefs && (typeof prefs.gapPt === "number") && prefs.gapPt >= 0) ? ptToUnit(prefs.gapPt, unitInfo) : 5;
    var initOffsetUnit = (prefs && (typeof prefs.offsetPt === "number") && prefs.offsetPt >= 0) ? ptToUnit(prefs.offsetPt, unitInfo) : 0;
    var initCapMode = (prefs && (typeof prefs.capMode === "number")) ? prefs.capMode : 0;
    var initReverse = (prefs && prefs.reversePath === true);
    var initAdjustEnds = (prefs && (typeof prefs.adjustEnds === "boolean")) ? prefs.adjustEnds : true;
    var initMode = (prefs && (typeof prefs.mode === "number")) ? prefs.mode : 0; // 0: Gap→Dash / 1: Dash→Gap
    var initDashUnit = (prefs && (typeof prefs.dashPt === "number") && prefs.dashPt >= 0) ? ptToUnit(prefs.dashPt, unitInfo) : 0;

    function fmtFieldNumber(v) {
        if (v == null || isNaN(v)) return "0";
        var rounded = Math.round(v * 1000) / 1000;
        if (Math.abs(rounded - Math.round(rounded)) < 1e-10) return String(Math.round(rounded));
        return String(rounded);
    }

    // --- ダイアログの作成 ---
    var win = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
    win.orientation = "column";
    win.alignChildren = "fill";

    // 選択中のパス情報（全幅）
    var infoPanel = win.add("panel", undefined, L("panelPathInfo"));
    infoPanel.alignChildren = "center";
    infoPanel.margins = [15, 20, 15, 10];
    infoPanel.add("statictext", undefined, L("pathLength") + " " + ptToUnit(pathLength, unitInfo).toFixed(3) + " " + unitInfo.label);

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

    // パスの分割（左カラム）
    var splitPanel = colLeft.add("panel", undefined, "パスの分割");
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
    var methodPanel = splitPanel.add("panel", undefined, L("calcMethodPanel"));
    methodPanel.alignChildren = "left";
    methodPanel.margins = [15, 20, 15, 10];

    var modeGroup = methodPanel.add("group");
    modeGroup.orientation = "column";
    modeGroup.alignChildren = "left";

    var rbModeGapToDash = modeGroup.add("radiobutton", undefined, L("modeGapToDash"));
    var rbModeDashToGap = modeGroup.add("radiobutton", undefined, L("modeDashToGap"));
    rbModeGapToDash.value = (initMode !== 1);
    rbModeDashToGap.value = (initMode === 1);

    // 初期モードに合わせて表示を切替
    (function () {
        var dashToGap = rbModeDashToGap.value;
        inpDash.visible = dashToGap;
        outDash.visible = !dashToGap;
        inpGap.visible = !dashToGap;
        outGap.visible = dashToGap;
    })();

    // 開始位置パネル
    var offsetPanel = colRight.add("panel", undefined, L("offsetPanel"));
    offsetPanel.alignChildren = "left";
    offsetPanel.margins = [15, 20, 15, 10];

    // 開始位置（入力）
    var grpOffset = offsetPanel.add("group");
    grpOffset.orientation = "row";
    grpOffset.alignChildren = ["left", "center"];

    // var stOffset = grpOffset.add("statictext", undefined, L("offsetLabel"));
    // stOffset.preferredSize.width = LABEL_W;
    // stOffset.justify = "right";

    var inpOffset = grpOffset.add("edittext", undefined, fmtFieldNumber(initOffsetUnit));
    inpOffset.characters = 4;

    var stOffsetUnit = grpOffset.add("statictext", undefined, unitInfo.label);

    // ↑↓キーで増減（開始位置は0以上）
    inpOffset._forceInteger = false;
    inpOffset._minValue = 0;
    changeValueByArrowKey(inpOffset);

    // 1周期（= dash+gap）基準のプリセット
    var offsetPresetGroup = offsetPanel.add("group");
    offsetPresetGroup.orientation = "row";
    offsetPresetGroup.alignChildren = ["left", "center"];

    var rbOff0 = offsetPresetGroup.add("radiobutton", undefined, "0");
    var rbOffQ1 = offsetPresetGroup.add("radiobutton", undefined, "1/4");
    var rbOffQ2 = offsetPresetGroup.add("radiobutton", undefined, "1/2");
    var rbOffQ3 = offsetPresetGroup.add("radiobutton", undefined, "3/4");
    rbOff0.value = true;

    // 貼付入力（例: 0,15,20,13）
    var grpPaste = offsetPanel.add("group");
    grpPaste.orientation = "row";
    grpPaste.alignChildren = ["left", "center"];

    // ラベル相当の空き
    // var stPasteSpacer = grpPaste.add("statictext", undefined, "");
    // stPasteSpacer.preferredSize.width = LABEL_W;

    var inpPaste = grpPaste.add("edittext", undefined, "");
    inpPaste.characters = 12;

    var btnPasteApply = grpPaste.add("button", undefined, L("pasteApply"));
    btnPasteApply.preferredSize.width = 44;

    // 両端を調整（オープンパス向け）
    var cbAdjustEnds = offsetPanel.add("checkbox", undefined, L("adjustEnds"));
    cbAdjustEnds.value = initAdjustEnds;
    // クローズパスでは意味がないためディム表示
    if (sel.closed) cbAdjustEnds.enabled = false;


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

    var cbReversePath = grpReversePath.add("checkbox", undefined, L("reversePath"));
    cbReversePath.value = initReverse;

    // ボタン
    var btnGroup = win.add("group");
    btnGroup.alignment = "center";
    var btnCancel = btnGroup.add("button", undefined, L("btnCancel"), { name: "cancel" });
    var btnOK = btnGroup.add("button", undefined, L("btnOK"), { name: "ok" });

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
    function calcDashAndStepPt(segments, gapPt) {
        if (!(segments > 0)) return null;
        if (!(gapPt >= 0)) return null;

        if (sel.closed) {
            var stepPt = pathLength / segments;
            var dashPt = stepPt - gapPt;
            return { dashPt: dashPt, stepPt: stepPt };
        }

        // open path
        // 両端を調整: 両端をダッシュで終える（分割数=ダッシュ本数、gapは(分割数-1)回）
        // OFF: クローズと同じ計算（1周期=全長/分割数、末尾は端数になり得る）
        if (!cbAdjustEnds.value) {
            var stepPtOpen = pathLength / segments;
            var dashPtOpen = stepPtOpen - gapPt;
            return { dashPt: dashPtOpen, stepPt: stepPtOpen };
        }

        if (segments === 1) {
            var dash1 = pathLength;
            return { dashPt: dash1, stepPt: dash1 + gapPt };
        }

        var dash = (pathLength - gapPt * (segments - 1)) / segments;
        return { dashPt: dash, stepPt: dash + gapPt };
    }

    // 線分(Dash)から間隔(Gap)を逆算（pt）
    function calcGapPtFromDashPt(segments, dashPt) {
        if (!(segments > 0)) return null;
        if (!(dashPt >= 0)) return null;

        // クローズ、または「両端を調整」OFFは 1周期=全長/分割数 で扱う
        if (sel.closed || !cbAdjustEnds.value) {
            var stepPt = pathLength / segments;
            return stepPt - dashPt;
        }

        // open + 両端を調整ON
        if (segments === 1) return 0;
        return (pathLength - dashPt * segments) / (segments - 1);
    }

    // 1周期（= dash+gap）を unit にした値を返す
    function getStepUnit() {
        var segments = parseInt(inpSegments.text, 10);
        if (pasteMode && pasteDashPts && pasteDashPts.length > 0) {
            var corePts = getPasteCoreDashPts(pasteDashPts);
            var sumPt = 0;
            for (var i = 0; i < corePts.length; i++) sumPt += corePts[i];
            return ptToUnit(sumPt, unitInfo);
        }
        if (isNaN(segments) || segments <= 0) return null;

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
                var gapPt = calcGapPtFromDashPt(segments, dashPt);
                if (gapPt == null) return null;
                stepPt = dashPt + gapPt;
            }

            return ptToUnit(stepPt, unitInfo);
        }

        // Gap→Dash
        var gapUnit = parseFloat(inpGap.text);
        if (isNaN(gapUnit) || gapUnit < 0) return null;

        var gapPt = unitToPt(gapUnit, unitInfo);
        var r = calcDashAndStepPt(segments, gapPt);
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

    // プリセット（0 / 1/4 / 1/2 / 3/4）と入力値を同期
    function syncOffsetPreset(offsetUnit) {
        // 一旦すべて解除（カスタム値もあり得る）
        rbOff0.value = rbOffQ1.value = rbOffQ2.value = rbOffQ3.value = false;

        var stepUnit = getStepUnit();
        if (stepUnit == null) return;

        // 単位によって誤差が出るので、ゆるめの許容
        var tol = Math.max(0.001, Math.abs(stepUnit) * 0.0005);

        var t0 = 0;
        var t1 = stepUnit * 0.25;
        var t2 = stepUnit * 0.50;
        var t3 = stepUnit * 0.75;

        if (Math.abs(offsetUnit - t0) <= tol) rbOff0.value = true;
        else if (Math.abs(offsetUnit - t1) <= tol) rbOffQ1.value = true;
        else if (Math.abs(offsetUnit - t2) <= tol) rbOffQ2.value = true;
        else if (Math.abs(offsetUnit - t3) <= tol) rbOffQ3.value = true;
    }

    // プリセットを適用（1周期基準）
    function applyOffsetPreset(frac) {
        var stepUnit = getStepUnit();
        if (stepUnit == null) return;
        setOffsetText(stepUnit * frac);
        updatePreview();
    }

    function exitPasteMode() {
        pasteMode = false;
        pasteDashPts = null;
        pasteDisplayIndex = 0;
        inpPaste.text = "";
    }

    function getPasteDisplayIndexFromUnits(unitsArr) {
        // 0,100,50,20 のように先頭0があり、かつ4つ以上ある場合は
        // UIの「線分/間隔」は 100,50 を採用する（先頭0はパターンとして保持）
        if (!unitsArr || unitsArr.length < 2) return 0;
        if (unitsArr.length >= 4 && Math.abs(unitsArr[0]) < 1e-12) return 1;
        return 0;
    }

    function hasLeadingZeroOffsetDashPts(dashPts) {
        return !!(dashPts && dashPts.length >= 4 && Math.abs(dashPts[0]) < 1e-10);
    }

    // pasteDashPts から「0,xxx」のプレフィックスを除いたコア配列を返す
    function getPasteCoreDashPts(dashPts) {
        if (hasLeadingZeroOffsetDashPts(dashPts)) return dashPts.slice(2);
        return dashPts.slice(0);
    }

    // 開始位置（offsetUnit）に応じて [0,offset,...] を付与（offset<=0なら付与しない）
    function buildDashPtsWithOffset(corePts, offsetUnit) {
        var offsetPt = unitToPt(offsetUnit, unitInfo);
        if (!(offsetPt > 0)) return corePts.slice(0);
        return [0, offsetPt].concat(corePts);
    }

    function getDisplayIndexFromDashPts(dashPts) {
        return hasLeadingZeroOffsetDashPts(dashPts) ? 1 : 0;
    }

    // 対応例:
    // - 0,100,50,20 -> dashes=[0,100,50,20]（UIの線分/間隔は 100,50 を採用）
    // - 100,50      -> dashes=[100,50]
    // - [8,4,2,4]   -> dashes=[8,4,2,4]

    function parsePasteText(s) {
        s = (s == null) ? "" : String(s);
        s = s.replace(/[，、]/g, ",");
        // 例: [0,100,50,20] のような括弧付きにも対応
        s = s.replace(/[\[\]\(\)\{\}]/g, "");

        var parts = s.split(/[\s,]+/);
        var nums = [];

        for (var i = 0; i < parts.length; i++) {
            var t = parts[i];
            if (!t) continue;
            var n = parseFloat(t);
            if (isNaN(n)) return null;
            nums.push(n);
        }


        // 破線配列は最低2つ必要
        if (nums.length < 2) return null;

        return { dashesUnit: nums };
    }

    function applyPaste() {
        var r = parsePasteText(inpPaste.text);
        if (!r) {
            alert(L("alertPasteInvalid"));
            return;
        }

        var di = getPasteDisplayIndexFromUnits(r.dashesUnit);
        pasteDisplayIndex = di;

        // dashes（unit → pt）
        var pts = [];
        for (var i = 0; i < r.dashesUnit.length; i++) {
            var u = r.dashesUnit[i];
            if (isNaN(u) || u < 0) {
                alert(L("alertPasteInvalid"));
                return;
            }
            pts.push(unitToPt(u, unitInfo));
        }

        // UI同期（線分/間隔に入れるペアを選ぶ）
        var du0 = (r.dashesUnit.length > di) ? r.dashesUnit[di] : r.dashesUnit[0];
        var du1 = (r.dashesUnit.length > di + 1) ? r.dashesUnit[di + 1] : r.dashesUnit[1];

        // UI同期（選んだペアを「線分」「間隔」に入れる）
        inpDash.text = fmtFieldNumber(du0);
        inpGap.text = fmtFieldNumber(du1);

        // 先頭が 0 の貼付（例: 0,100,50,20）のときは、開始位置に 2番目を自動代入
        if (r.dashesUnit && r.dashesUnit.length >= 4 && Math.abs(r.dashesUnit[0]) < 1e-12) {
            inpOffset.text = fmtFieldNumber(r.dashesUnit[1]);
        }

        pasteMode = true;
        pasteDashPts = pts;

        updatePreview();
    }

    // プリセット（1周期基準）
    rbOff0.onClick = function () { applyOffsetPreset(0); };
    rbOffQ1.onClick = function () { applyOffsetPreset(0.25); };
    rbOffQ2.onClick = function () { applyOffsetPreset(0.50); };
    rbOffQ3.onClick = function () { applyOffsetPreset(0.75); };

    btnPasteApply.onClick = applyPaste;
    inpPaste.onChange = applyPaste;

    // 値変更時にプレビュー更新（アラート無し）
    function updatePreview() {
        var offsetUnit = parseFloat(inpOffset.text);
        if (isNaN(offsetUnit) || offsetUnit < 0) {
            outDash.text = "";
            outGap.text = "";
            return;
        }

        // 貼付モード：開始位置は配列で表現（strokeDashOffsetは使わない）
        if (pasteMode && pasteDashPts && pasteDashPts.length > 0) {
            // 貼付のコア配列を取り出し、開始位置フィールドに応じて [0,offset,...] を付与
            var corePts = getPasteCoreDashPts(pasteDashPts);
            var finalPts = buildDashPtsWithOffset(corePts, offsetUnit);

            // 表示は [0,offset,...] の場合、offset から
            var di = getDisplayIndexFromDashPts(finalPts);
            pasteDisplayIndex = di;

            var p0 = (finalPts.length > di) ? finalPts[di] : (finalPts.length >= 1 ? finalPts[0] : 0);
            var p1 = (finalPts.length > di + 1) ? finalPts[di + 1] : (finalPts.length >= 2 ? finalPts[1] : 0);

            outDash.text = (p0 != null) ? ptToUnit(p0, unitInfo).toFixed(3) : "";
            outGap.text = (p1 != null) ? ptToUnit(p1, unitInfo).toFixed(3) : "";

            // 入力欄も見た目を揃える（貼付モード中は計算には使わない）
            inpDash.text = fmtFieldNumber(ptToUnit(p0, unitInfo));
            inpGap.text = fmtFieldNumber(ptToUnit(p1, unitInfo));

            syncOffsetPreset(offsetUnit);

            sel.stroked = true;

            if (rbCapRound.value) sel.strokeCap = StrokeCap.ROUNDENDCAP;
            else if (rbCapProject.value) sel.strokeCap = StrokeCap.PROJECTINGENDCAP;
            else sel.strokeCap = StrokeCap.BUTTENDCAP;

            // 開始位置は dashOffset ではなく配列で表現
            sel.strokeDashOffset = 0;
            sel.strokeDashes = finalPts.slice(0);

            app.redraw();
            return;
        }

        var segments = parseInt(inpSegments.text, 10);
        if (isNaN(segments) || segments <= 0) {
            outDash.text = "";
            outGap.text = "";
            return;
        }

        var dashToGap = rbModeDashToGap.value;

        var gapPt, dashPt;

        if (dashToGap) {
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

            gapPt = calcGapPtFromDashPt(segments, dashPt);
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
            var r = calcDashAndStepPt(segments, gapPt);
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

        var offsetPt = unitToPt(offsetUnit, unitInfo);

        // プレビュー適用
        sel.stroked = true;

        // 線端
        if (rbCapRound.value) {
            sel.strokeCap = StrokeCap.ROUNDENDCAP;
        } else if (rbCapProject.value) {
            sel.strokeCap = StrokeCap.PROJECTINGENDCAP;
        } else {
            sel.strokeCap = StrokeCap.BUTTENDCAP;
        }

        // 開始位置は dashOffset ではなく配列で表現
        var corePts = [dashPt, gapPt];
        var finalPts = buildDashPtsWithOffset(corePts, offsetUnit);
        sel.strokeDashOffset = 0;
        sel.strokeDashes = finalPts;
        app.redraw();
    }

    function setReversePath(shouldReverse) {
        shouldReverse = !!shouldReverse;
        if (shouldReverse === directionReversed) return;

        try {
            // Reverse Path Direction は選択に対して実行されるため、対象を選択してから実行
            doc.selection = [sel];
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
                doc.selection = [sel];
                app.executeMenuCommand('Reverse Path Direction');
            } catch (_) { }
            directionReversed = false;
        }

        // ストローク状態の復元（まとめて）
        try {
            sel.strokeDashes = originalState.strokeDashes.slice(0);
            sel.strokeDashOffset = originalState.strokeDashOffset;
            sel.strokeCap = originalState.strokeCap;
            sel.stroked = originalState.stroked;
        } catch (_) { }

        try { app.redraw(); } catch (_) { }
    }

    function updatePreviewExitPaste() {
        if (pasteMode) exitPasteMode();
        updatePreview();
    }

    function updatePreviewOffset() {
        updatePreview();
    }

    // ↑↓キーによる値変更でもプレビューを更新
    inpSegments._onArrowChange = updatePreviewExitPaste;
    inpGap._onArrowChange = updatePreviewExitPaste;
    inpDash._onArrowChange = updatePreviewExitPaste;
    inpOffset._onArrowChange = updatePreviewOffset;

    // 入力値や線端の変更でプレビュー更新
    inpSegments.onChanging = updatePreviewExitPaste;
    inpGap.onChanging = updatePreviewExitPaste;
    inpDash.onChanging = updatePreviewExitPaste;
    inpOffset.onChanging = updatePreviewOffset;
    inpSegments.onChange = updatePreviewExitPaste;
    inpGap.onChange = updatePreviewExitPaste;
    inpDash.onChange = updatePreviewExitPaste;
    inpOffset.onChange = updatePreviewOffset;
    rbCapNone.onClick = updatePreview;
    rbCapRound.onClick = updatePreview;
    rbCapProject.onClick = updatePreview;
    cbReversePath.onClick = function () { setReversePath(cbReversePath.value); updatePreview(); };
    cbAdjustEnds.onClick = updatePreviewExitPaste;
    rbModeGapToDash.onClick = function () {
        if (pasteMode) exitPasteMode();
        inpDash.visible = false;
        outDash.visible = true;
        inpGap.visible = true;
        outGap.visible = false;
        updatePreview();
    };
    rbModeDashToGap.onClick = function () {
        if (pasteMode) exitPasteMode();
        inpDash.visible = true;
        outDash.visible = false;
        inpGap.visible = false;
        outGap.visible = true;
        updatePreview();
    };

    // 前回値（パス方向）を反映
    setReversePath(cbReversePath.value);

    // 初期値でも一度プレビュー
    updatePreview();

    // OKボタンの処理（プレビューを確定して閉じる）
    btnOK.onClick = function () {
        var segments = parseInt(inpSegments.text, 10);
        var offsetUnit = parseFloat(inpOffset.text);

        if (isNaN(segments) || segments <= 0) {
            alert(L("alertSegmentsInvalid"));
            return;
        }
        if (isNaN(offsetUnit) || offsetUnit < 0) {
            alert(L("alertOffsetInvalid"));
            return;
        }

        if (pasteMode && pasteDashPts && pasteDashPts.length > 0) {
            try {
                sel.stroked = true;
                var corePts = getPasteCoreDashPts(pasteDashPts);
                var finalPts = buildDashPtsWithOffset(corePts, offsetUnit);
                sel.strokeDashOffset = 0;
                sel.strokeDashes = finalPts.slice(0);
            } catch (_) { }

            // 保存（UIに表示していた先頭ペアを保存：finalPtsに合わせる）
            try {
                var coreSavePts = getPasteCoreDashPts(pasteDashPts);
                var finalSavePts = buildDashPtsWithOffset(coreSavePts, offsetUnit);
                var di = getDisplayIndexFromDashPts(finalSavePts);
                var dashPtSave = (finalSavePts.length > di) ? finalSavePts[di] : (finalSavePts.length >= 1 ? finalSavePts[0] : 0);
                var gapPtSave = (finalSavePts.length > di + 1) ? finalSavePts[di + 1] : (finalSavePts.length >= 2 ? finalSavePts[1] : 0);
                var offsetPtSave = unitToPt(offsetUnit, unitInfo);
                var capModeSave = rbCapRound.value ? 1 : (rbCapProject.value ? 2 : 0);

                savePrefs({
                    segments: segments,
                    gapPt: gapPtSave,
                    dashPt: dashPtSave,
                    offsetPt: offsetPtSave,
                    capMode: capModeSave,
                    reversePath: cbReversePath.value,
                    adjustEnds: cbAdjustEnds.value,
                    mode: rbModeDashToGap.value ? 1 : 0
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

            gapPt = calcGapPtFromDashPt(segments, dashPt);
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
            var r = calcDashAndStepPt(segments, gapPt);
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
                mode: dashToGap ? 1 : 0
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