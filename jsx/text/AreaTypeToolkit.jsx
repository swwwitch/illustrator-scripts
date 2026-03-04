#target illustrator
// Use a persistent session engine for dialog position memory
#targetengine "AreaTypeToolkit"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
    作成日：2025-03-06
    更新日：2026-03-04

    概要：
    ポイント文字・パス上文字・図形・エリア内文字を対象に、エリア内文字の作成と調整を行うツール。

    ・ポイント文字 / パス上文字 → エリア内文字に変換（シンプル / ボタン風）
    ・図形をエリア内文字に変換してダミーテキストを入力
    ・（テキスト＋図形）ポイント文字 / パス上文字＋図形からエリア内文字を生成
    ・エリア内文字のサイズ / 行揃え / フレームサイズ / インデント / 余白を調整
    ・テキストの配置（上 / 中央 / 下 / 均等）をダイナミックアクションで適用

    ダイアログAでエリア内文字の作成方法を選択し、
    ダイアログBでエリア内文字の詳細調整を行う。

    プレビューON時は変更内容がリアルタイムで反映される。
*/

// --- Script version / スクリプトバージョン ---

// --- Dummy text for "Use selected object + dummy" ---
var DUMMY_TEXT_JA = "今日は15:00から予定外のミーティングがありました。疲れを癒すため、夕方に近くのカフェWAVEで、お気に入りの抹茶ラテを楽しみました。短い休憩でしたが、心が“ほっと”しました。";
var DUMMY_TEXT_EN = "An unplanned meeting popped up at 3:00 PM today. To reset, I stopped by a nearby café for a matcha latte. It was a short break, but it helped me breathe and refocus.";

function getTextFontSafe(preferNames) {
    // preferNames: ["ExactFontName", ...]
    try {
        if (preferNames && preferNames.length) {
            for (var i = 0; i < preferNames.length; i++) {
                try {
                    var f = app.textFonts.getByName(preferNames[i]);
                    if (f) return f;
                } catch (e0) { }
            }
        }
        // Fallback: return first available font
        if (app.textFonts.length > 0) return app.textFonts[0];
    } catch (e) { }
    return null;
}

var SCRIPT_VERSION = "v1.1.3";

// --- Localization / ローカライズ ---
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: { ja: "AreaType Toolkit", en: "AreaType Toolkit" },

    // Top mode panel
    pnlSeparate: { ja: "テキストを分離", en: "Separate text" },
    radNoAdjust: { ja: "調整しない", en: "No adjust" },
    radSeparate: { ja: "1pt黒に", en: "Stroke 1pt black" },
    radPathOpacity: { ja: "パスをなし", en: "No path" },
    radRemovePath: { ja: "パスを削除", en: "Remove path" },

    // Convert panel
    pnlConvert: { ja: "ポイント文字→エリア内文字", en: "Point Type → Area Type" },
    radPointType: { ja: "シンプル", en: "Simple" },
    radAreaType: { ja: "ボタン風", en: "Button style" },
    radUseSelected: { ja: "選択オブジェクトを利用", en: "Use selected object" },
    radUseSelectedDummy: { ja: "選択オブジェクトにダミーテキスト", en: "Dummy text on selected object" },
    btnConvert: { ja: "変換", en: "Convert" },

    // Right column panels
    pnlJustify: { ja: "行揃え", en: "Justification" },
    justLeft: { ja: "左", en: "Left" },
    justCenter: { ja: "中央", en: "Center" },
    justRight: { ja: "右", en: "Right" },
    justFull: { ja: "均等配置（最終行左）", en: "Justify (last line left)" },
    justFullAll: { ja: "両端揃え（すべての行）", en: "Justify all lines" },

    pnlTextAlign: { ja: "テキストの配置", en: "Text alignment" },
    alignTop: { ja: "上揃え", en: "Top" },
    alignCenter: { ja: "中央揃え", en: "Center" },
    alignBottom: { ja: "下揃え", en: "Bottom" },
    alignJustify: { ja: "均等配置", en: "Justify" },
    pnlThread: { ja: "スレッドテキスト", en: "Thread text" },
    btnThreadText: { ja: "スレッドテキストに", en: "Make thread text" },

    // Left column panels
    pnlAutoSize: { ja: "フォントサイズ", en: "Font size" },
    lblFontSize: { ja: "フォントサイズ", en: "Font size" },
    btnTextSize: { ja: "文字あふれ解消", en: "Make overset" },
    btnFit: { ja: "フィット", en: "Fit" },

    pnlFrameSize: { ja: "フレームサイズ", en: "Frame size" },
    lblWidth: { ja: "幅", en: "Width" },
    lblHeight: { ja: "高さ", en: "Height" },
    lblChars: { ja: "文字", en: "chars" },
    chkFrameAuto: { ja: "自動", en: "Auto" },

    pnlIndent: { ja: "インデント", en: "Indent" },
    indentLeft: { ja: "左", en: "Left" },
    indentRight: { ja: "右", en: "Right" },
    chkSync: { ja: "連動", en: "Link" },

    pnlOptions: { ja: "オプション", en: "Options" },
    chkMargin: { ja: "外側からの間隔", en: "Spacing" },
    chkLeader: { ja: "メニュー作成用（リーダー罫）", en: "Leader tabs" },

    // Bottom
    chkPreview: { ja: "プレビュー", en: "Preview" },
    btnRun: { ja: "実行", en: "Run" },
    btnClose: { ja: "閉じる", en: "Close" },

    // Alerts
    alertSelectText: { ja: "ポイント文字・パス上文字・エリア内文字を選択してください。", en: "Please select point text, path text, or area text." },
    alertNoDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." }
};


function L(key) {
    try {
        if (LABELS[key] && LABELS[key][lang]) return LABELS[key][lang];
        if (LABELS[key] && LABELS[key].en) return LABELS[key].en;
    } catch (e) { }
    return key;
}


// --- Path Text → Point Text (Dialog A preprocess) ---
function detachPathTextToPointText(doc, pathTextFrames) {
    function safe(fn) { try { return fn(); } catch (e) { return undefined; } }

    var created = [];
    if (!doc || !pathTextFrames || !pathTextFrames.length) return created;

    // Clear selection so only newly created texts can be selected if needed
    safe(function () { doc.selection = null; });

    for (var j = pathTextFrames.length - 1; j >= 0; j--) {
        var pathText = pathTextFrames[j];
        if (!pathText || pathText.typename !== "TextFrame" || pathText.kind !== TextType.PATHTEXT) continue;

        var originalPath = null;
        safe(function () { originalPath = pathText.textPath; });
        if (!originalPath) continue;

        // 1) Snapshot per-character attributes
        var charAttrs = [];
        for (var c = 0; c < pathText.characters.length; c++) {
            var ca = pathText.characters[c].characterAttributes;
            charAttrs.push({
                font: ca.textFont,
                size: ca.size,
                fillColor: ca.fillColor,
                strokeColor: ca.strokeColor,
                strokeWeight: ca.strokeWeight,
                autoLeading: ca.autoLeading,
                leading: ca.leading
            });
        }

        var textContents = "";
        safe(function () { textContents = pathText.contents; });

        var justification = null;
        safe(function () {
            if (pathText.paragraphs && pathText.paragraphs.length > 0) {
                justification = pathText.paragraphs[0].paragraphAttributes.justification;
            }
        });

        // 2) Create new Point Text at path start anchor
        var newText = doc.textFrames.add();
        var anchorPoint = null;
        safe(function () {
            if (originalPath.pathPoints && originalPath.pathPoints.length > 0) {
                anchorPoint = originalPath.pathPoints[0].anchor;
            }
        });
        if (anchorPoint) {
            newText.position = [anchorPoint[0], anchorPoint[1]];
        }

        newText.contents = textContents;

        if (justification !== null && newText.paragraphs && newText.paragraphs.length > 0) {
            safe(function () { newText.paragraphs[0].paragraphAttributes.justification = justification; });
        }

        // Clear default stroke first; restore per-character later
        safe(function () {
            var nc = new NoColor();
            newText.textRange.characterAttributes.strokeColor = nc;
            newText.textRange.characterAttributes.strokeWeight = 0;
        });

        // Restore per-character attributes
        var n = Math.min(newText.characters.length, charAttrs.length);
        for (var k = 0; k < n; k++) {
            var targetCa = newText.characters[k].characterAttributes;
            var srcCa = charAttrs[k];

            safe(function () { targetCa.textFont = srcCa.font; });
            safe(function () { targetCa.size = srcCa.size; });
            safe(function () { targetCa.fillColor = srcCa.fillColor; });
            safe(function () {
                var sc = srcCa.strokeColor;
                targetCa.strokeColor = sc;
                targetCa.strokeWeight = (sc && sc.typename === "NoColor") ? 0 : srcCa.strokeWeight;
            });
            // safe(function () { targetCa.tracking = srcCa.tracking; }); // tracking not restored
            safe(function () { targetCa.baselineShift = 0; });
            safe(function () { targetCa.horizontalScale = 100; });
            safe(function () { targetCa.verticalScale = 100; });
            safe(function () { targetCa.autoLeading = srcCa.autoLeading; });
            if (!srcCa.autoLeading) safe(function () { targetCa.leading = srcCa.leading; });
        }

        // 3) Remove original Path Text (its path will be removed with it)
        safe(function () { pathText.remove(); });

        // 4) Select and return new text
        safe(function () { newText.selected = true; });
        created.push(newText);
    }

    return created;
}

function preprocessSelectionForDialogA(doc, sel) {
    // Convert selected Path Text items to Point Text so Dialog A can treat them like Point Type.
    if (!doc || !sel || !sel.length) return sel;

    var pathTexts = [];
    for (var i = 0; i < sel.length; i++) {
        var it = sel[i];
        try {
            if (it && it.typename === "TextFrame" && it.kind === TextType.PATHTEXT) {
                pathTexts.push(it);
            }
        } catch (e0) {
            // invalid object (e.g., already removed) — skip
        }
    }
    if (!pathTexts.length) return sel;

    var newTexts = detachPathTextToPointText(doc, pathTexts);
    if (!newTexts.length) return sel;

    // Build a new selection array: replace pathTexts with new point texts, keep other items.
    var out = [];
    for (var j = 0; j < sel.length; j++) {
        var it2 = sel[j];
        try {
            if (it2 && it2.typename === "TextFrame" && it2.kind === TextType.PATHTEXT) {
                // skip old
            } else if (it2) {
                out.push(it2);
            }
        } catch (e1) {
            // invalid object — skip
        }
    }
    for (var k = 0; k < newTexts.length; k++) out.push(newTexts[k]);

    try { doc.selection = out; } catch (e) { }
    try { app.redraw(); } catch (e2) { }

    return out;
}


// --- 自動サイズ調整ヘルパー（AutoFitTextFrame.jsx 参考） ---

function safeOverflows(tf) {
    try {
        if (tf && typeof tf.overflows !== "undefined") return !!tf.overflows;
    } catch (e) { }
    return null;
}

function isOverset(tf, lineAmt) {
    if (tf.lines.length > 0) {
        var chars = 0;
        lineAmt = (typeof lineAmt === "undefined" || lineAmt === null) ? 1 : Math.floor(lineAmt);
        if (lineAmt < 1) lineAmt = 1;
        if (lineAmt > tf.lines.length) lineAmt = tf.lines.length;
        for (var i = 0; i < lineAmt; i++) { chars += tf.lines[i].characters.length; }
        return chars < tf.characters.length;
    }
    return tf.characters.length > 0;
}

function isOversetFrame(tf) {
    if (tf && tf.kind === TextType.AREATEXT) {
        var ov = safeOverflows(tf);
        if (ov !== null) return ov;
        try { return isOverset(tf, tf.lines.length || 1); } catch (e) { return false; }
    }
    try { return isOverset(tf, tf.lines.length || 1); } catch (e) { return false; }
}

function getLeadingInfo(tf) {
    try {
        var attrs = tf.textRange.characterAttributes;
        if (attrs.autoLeading) return null;
        var size = attrs.size, leading = attrs.leading;
        if (size > 0 && leading > 0) return { ratio: leading / size };
    } catch (e) { }
    return null;
}

function applyLeading(tf, newSize, li) {
    if (!li) return;
    try { tf.textRange.characterAttributes.leading = newSize * li.ratio; } catch (e) { }
}


// --- Dynamic Action loader (generate once per session) ---
function _hexAscii(s) {
    var out = "";
    for (var i = 0; i < s.length; i++) {
        var h = s.charCodeAt(i).toString(16);
        if (h.length < 2) h = "0" + h;
        out += h;
    }
    return out;
}

function _nameBlockAscii(s) {
    // /name [ <len> <hex> ]  (ASCII)
    return "/name [ " + s.length + " " + _hexAscii(s).toUpperCase() + " ]";
}

function _buildActionSetAIA(setName, internalName, localizedNameHex, paramKeyInt, actionDefs) {
    // actionDefs: [{name: "...", value: <int>}, ...]
    var str = "/version 3" +
        _nameBlockAscii(setName) +
        "/isOpen 1" +
        "/actionCount " + actionDefs.length;

    for (var i = 0; i < actionDefs.length; i++) {
        var a = actionDefs[i];
        str += "/action-" + (i + 1) + " {" +
            " " + _nameBlockAscii(a.name) +
            " /keyIndex 0" +
            " /colorIndex 0" +
            " /isOpen 1" +
            " /eventCount 1" +
            " /event-1 {" +
            " /useRulersIn1stQuadrant 0" +
            " /internalName (" + internalName + ")" +
            (localizedNameHex ? (" /localizedName [ " + localizedNameHex + " ]") : "") +
            " /isOpen 0" +
            " /isOn 1" +
            " /hasDialog 0" +
            " /parameterCount 1" +
            " /parameter-1 {" +
            " /key " + paramKeyInt +
            " /showInPalette 4294967295" +
            " /type (integer)" +
            " /value " + a.value +
            " }" +
            " }" +
            "}";
    }
    return str;
}

function _ensureActionSetLoadedOnce(setName, aiaString, flagKey) {
    try {
        if ($.global[flagKey]) return;
        // 既存があれば一旦外して衝突回避
        try { app.unloadAction(setName, ""); } catch (e0) { }

        var f = new File(Folder.temp + "/AreaTypeToolkit_action_" + setName + ".aia");
        f.open("w");
        f.write(aiaString);
        f.close();

        app.loadAction(f);
        try { f.remove(); } catch (e1) { }

        $.global[flagKey] = true;
    } catch (e) { }
}

function ensureAreaTypeActionsLoaded() {
    // Set: AreaType / Actions: AutoSizeOn(1), AutoSizeOff(2)
    var setName = "AreaType";
    var aia = _buildActionSetAIA(
        setName,
        "adobe_SLOAreaTextDialog",
        // localizedName (same as previous) length 33 hex bytes
        "33 e382a8e383aae382a2e58685e69687e5ad97e382aae38397e382b7e383a7e383b3",
        1952539754,
        [
            { name: "AutoSizeOn", value: 1 },
            { name: "AutoSizeOff", value: 2 }
        ]
    );
    _ensureActionSetLoadedOnce(setName, aia, "__AreaTypeToolkit_loaded_AreaType");
}

function ensureAreaTextActionsLoaded() {
    // Set: AreaText / Actions: AlignTop(0), AlignCenter(1), AlignBottom(2), AlignJustify(3)
    var setName = "AreaText";
    var aia = _buildActionSetAIA(
        setName,
        "adobe_frameAlignment",
        // localizedName (same as previous) length 39 hex bytes (kept as-is)
        "39 e382a8e383aae382a2e58685e69687e5ad97e381aee38395e383ace383bce383a0e695b4e58897",
        1717660782,
        [
            { name: "AlignTop", value: 0 },
            { name: "AlignCenter", value: 1 },
            { name: "AlignBottom", value: 2 },
            { name: "AlignJustify", value: 3 }
        ]
    );
    _ensureActionSetLoadedOnce(setName, aia, "__AreaTypeToolkit_loaded_AreaText");
}
// フレームサイズ：Illustrator アクションで自動サイズ調整をON/OFFにする
function expandFrameToFit(tf) {
    app.activeDocument.selection = [tf];
    act_setAutoSizeAdjust(1);
}

function collapseFrameAuto(tf) {
    app.activeDocument.selection = [tf];
    act_setAutoSizeAdjust(2);
}

function act_setAutoSizeAdjust(valueInt) {
    // valueInt: 1 = ON, 2 = OFF
    if (valueInt !== 1 && valueInt !== 2) return;

    ensureAreaTypeActionsLoaded();

    // Action name: AutoSizeOn / AutoSizeOff  (Set name: AreaType)
    var actionName = (valueInt === 1) ? "AutoSizeOn" : "AutoSizeOff";
    try { app.doScript(actionName, "AreaType", false); } catch (e) { }
}

// テキストの配置（フレーム整列）：Illustrator アクションでエリア内文字の配置を変更する
// valueInt: 0=上, 1=中央, 2=下, 3=均等配置
function act_alignHorizontal(valueInt) {
    // valueInt: 0=上, 1=中央, 2=下, 3=均等配置
    if (valueInt !== 0 && valueInt !== 1 && valueInt !== 2 && valueInt !== 3) return;

    ensureAreaTextActionsLoaded();

    // Action names: AlignTop / AlignCenter / AlignBottom / AlignJustify (Set name: AreaText)
    var actionName = "AlignTop";
    if (valueInt === 1) actionName = "AlignCenter";
    else if (valueInt === 2) actionName = "AlignBottom";
    else if (valueInt === 3) actionName = "AlignJustify";

    try { app.doScript(actionName, "AreaText", false); } catch (e) { }
}

function applyAreaTextFrameAlignment(tf, valueInt, forPreview) {
    // app.doScript はプレビュー中にダイアログから呼ぶと不安定なためスキップ
    if (forPreview) return;
    try {
        var doc = app.activeDocument;
        doc.selection = null;
        doc.selection = [tf];
        app.redraw(); // Illustratorに選択状態を確定させる
        act_alignHorizontal(valueInt);
    } catch (e) { }
}

// ちぢむ処理（バイナリサーチ：最大 40 回でオーバーセットにならない最大サイズを探す）
function shrinkFont(tf) {
    if (tf.characters.length <= 0 || !isOversetFrame(tf)) return;
    var li = getLeadingInfo(tf);
    var hi = tf.textRange.characterAttributes.size;
    var lo = 0.1;

    // lo でもオーバーセットなら最小サイズのまま終了
    tf.textRange.characterAttributes.size = lo;
    applyLeading(tf, lo, li);
    if (isOversetFrame(tf)) return;

    // バイナリサーチ
    for (var iter = 0; iter < 40; iter++) {
        var mid = (lo + hi) / 2;
        tf.textRange.characterAttributes.size = mid;
        applyLeading(tf, mid, li);
        if (isOversetFrame(tf)) {
            hi = mid;
        } else {
            lo = mid;
        }
        if (hi - lo < 0.1) break;
    }

    // オーバーセットにならない側（lo）に確定
    tf.textRange.characterAttributes.size = lo;
    applyLeading(tf, lo, li);
}

// テキスト（フィット）：オーバーセットになるまで拡大 → ちぢむ（fitFont 相当）
function fitTextToFrame(tf) {
    if (tf.characters.length <= 0) return;
    var li = getLeadingInfo(tf);
    var original = tf.textRange.characterAttributes.size;

    // Step 1: オーバーセットが出るまで 2 倍ずつ拡大
    if (!isOversetFrame(tf)) {
        var high = original, guard = 0;
        while (!isOversetFrame(tf) && guard < 25) {
            guard++;
            high = high * 2;
            if (high > 100000) break;
            try { tf.textRange.characterAttributes.size = high; applyLeading(tf, high, li); } catch (e) { break; }
        }
    }

    // オーバーセットが出なければ元に戻して終了
    if (!isOversetFrame(tf)) {
        try { tf.textRange.characterAttributes.size = original; applyLeading(tf, original, li); } catch (e) { }
        return;
    }

    // Step 2: ちぢむ処理
    shrinkFont(tf);
}

// ↑↓キーで値を増減（プレビュー更新コールバックつき）
function changeValueByArrowKey(editText, allowNegative, onChangeCallback) {
    editText.addEventListener("keydown", function (event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;

        if (keyboard.shiftKey) {
            var delta = 10;
            if (event.keyName == "Up") {
                value = Math.ceil((value + 1) / delta) * delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value = Math.floor((value - 1) / delta) * delta;
                event.preventDefault();
            }
        } else if (keyboard.altKey) {
            if (event.keyName == "Up") {
                value += 0.1;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value -= 0.1;
                event.preventDefault();
            }
        } else {
            if (event.keyName == "Up") {
                value += 1;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value -= 1;
                event.preventDefault();
            }
        }

        value = keyboard.altKey ? Math.round(value * 10) / 10 : Math.round(value);
        if (!allowNegative && value < 0) value = 0;

        editText.text = value;
        if (typeof onChangeCallback === "function") { onChangeCallback(); }
    });
}

// メニュー作成用リーダー罫：各段落に右揃えタブ（リーダー「…」400pt）を適用
function applyLeaderTab(tf) {
    try {
        var tabStop = new TabStopInfo();
        tabStop.position = 400;
        tabStop.alignment = TabStopAlignment.Right;
        tabStop.leader = "…";
        for (var p = 0; p < tf.paragraphs.length; p++) {
            var pa = tf.paragraphs[p].paragraphAttributes;
            pa.tabStops = [tabStop];
            pa.justification = Justification.RIGHT;
        }
    } catch (e) { }
}

// ルーラー単位に応じたラベルと pt 変換係数を返す
function getRulerUnitInfo(doc) {
    var ru = doc.rulerUnits;
    if (ru === RulerUnits.Millimeters) return { label: "mm", toPt: 72 / 25.4 };
    if (ru === RulerUnits.Centimeters) return { label: "cm", toPt: 72 / 2.54 };
    if (ru === RulerUnits.Inches) return { label: "in", toPt: 72 };
    if (ru === RulerUnits.Points) return { label: "pt", toPt: 1 };
    if (ru === RulerUnits.Picas) return { label: "pica", toPt: 12 };
    if (ru === RulerUnits.Pixels) return { label: "px", toPt: 72 / 96 };
    return { label: "pt", toPt: 1 };
}

// ダイアログ位置の保存・復元（セッション中のみ保持） / Save & restore dialog position (session only)
function saveDlgPosition(dlg) {
    try {
        var loc = dlg.location;
        $.global.__AreaTypeToolkit_dlgPos = { x: loc[0], y: loc[1] };
    } catch (e) { }
}
function restoreDlgPosition(dlg) {
    try {
        var data = $.global.__AreaTypeToolkit_dlgPos;
        if (!data) return;
        if (typeof data.x === "number" && typeof data.y === "number") {
            dlg.location = [data.x, data.y];
        }
    } catch (e) { }
}

// ============================================================
// ダイアログA：ポイント文字 → エリア内文字 変換
// ============================================================
function showDialogA(doc, sel) {
    sel = preprocessSelectionForDialogA(doc, sel);
    var dlgA = new Window("dialog", "エリア内文字に変換");
    dlgA.alignChildren = "fill";
    dlgA.margins = 20;

    // ラジオボタン（パネルなし）
    var grpConvert = dlgA.add("group");
    grpConvert.orientation = "column";
    grpConvert.alignChildren = ["left", "center"];

    // grpConvert.add("statictext", undefined, L("pnlConvert"));

    var radPointType = grpConvert.add("radiobutton", undefined, L("radPointType"));
    var radAreaType = grpConvert.add("radiobutton", undefined, L("radAreaType"));
    var radUseSelected = grpConvert.add("radiobutton", undefined, L("radUseSelected"));
    var radUseSelectedDummy = grpConvert.add("radiobutton", undefined, L("radUseSelectedDummy"));
    radPointType.value = true;

    // ボタンエリア：左から［キャンセル］［変換］
    var grpBottomA = dlgA.add("group");
    grpBottomA.orientation = "row";
    grpBottomA.alignment = ["center", "center"]; // center buttons
    grpBottomA.spacing = 10;
    var btnCloseA = grpBottomA.add("button", undefined, L("btnClose"), { name: "cancel" });
    var btnConvert = grpBottomA.add("button", undefined, L("btnConvert"), { name: "ok" });

    // 選択種別フラグ
    var hasPointText = false, hasPathItem = false;
    for (var ti = 0; ti < sel.length; ti++) {
        if (sel[ti].typename === "TextFrame" && (sel[ti].kind === TextType.POINTTEXT || sel[ti].kind === TextType.PATHTEXT)) hasPointText = true;
        if (sel[ti].typename === "PathItem" || sel[ti].typename === "CompoundPathItem") hasPathItem = true;
    }

    // 選択内容に応じてモードを誘導
    if (hasPointText && hasPathItem) {
        // テキスト＋図形 → シンプル・ボタン風をディム、選択オブジェクトを利用を自動選択
        radPointType.enabled = false;
        radAreaType.enabled = false;
        radUseSelected.value = true; // default, dummy is also available
    } else if (hasPointText && !hasPathItem) {
        // ポイント文字のみ → 選択オブジェクト系をディム
        radUseSelected.enabled = false;
        radUseSelectedDummy.enabled = false;
    } else if (!hasPointText && hasPathItem) {
        // 図形のみ → 「選択オブジェクトにダミーテキスト」を自動選択し、残りをディム
        radPointType.enabled = false;
        radAreaType.enabled = false;
        radUseSelected.enabled = false;
        radUseSelectedDummy.enabled = true;
        radUseSelectedDummy.value = true;
    }

    // ラジオ排他制御
    var allConvertRadios = [radPointType, radAreaType, radUseSelected, radUseSelectedDummy];
    for (var cr = 0; cr < allConvertRadios.length; cr++) {
        (function (btn) {
            btn.onClick = function () {
                for (var j = 0; j < allConvertRadios.length; j++) {
                    allConvertRadios[j].value = (allConvertRadios[j] === btn);
                }
            };
        })(allConvertRadios[cr]);
    }

    // 変換後にダイアログBを開くための変数
    var convertedFrames = [];
    var shouldOpenB = false;
    var _openBDefaultAlignInt = null; // 0=Top,1=Center,2=Bottom,3=Justify

    btnConvert.onClick = function () {
        var currentSel = app.activeDocument.selection;
        if (!currentSel || currentSel.length === 0) { return; }
        var useSel = !!radUseSelected.value;
        var useSelDummy = !!radUseSelectedDummy.value;
        var created = runConvert(doc, currentSel,
            radPointType.value, radAreaType.value, useSel, useSelDummy);
        if (created.length > 0) {
            convertedFrames = created;
            // If Button style was used, default Dialog B alignment to Center
            _openBDefaultAlignInt = radAreaType.value ? 1 : null;
            shouldOpenB = true;
            try { app.activeDocument.selection = created; } catch (e) { }
            dlgA.close(1);
        }
    };

    btnCloseA.onClick = function () { dlgA.close(0); };

    restoreDlgPosition(dlgA);
    dlgA.onClose = function () { saveDlgPosition(dlgA); return true; };
    dlgA.show();

    // dlgA.show() はブロッキング。ダイアログが閉じた後にBを開く
    if (shouldOpenB && convertedFrames.length > 0) {
        showDialogB(doc, convertedFrames[0], convertedFrames, _openBDefaultAlignInt);
    }
}

// ============================================================
// runConvert：ポイント文字をエリア内文字に変換して作成フレームを返す
// ============================================================
function runConvert(doc, sel, useSimple, useButton, useSelected, useSelectedDummy) {
    var created = [];

    function _getClosedPathItem(it) {
        try {
            if (!it) return null;
            if (it.typename === "PathItem") {
                if (it.closed) return it;
            } else if (it.typename === "CompoundPathItem") {
                if (it.pathItems && it.pathItems.length > 0) {
                    var p = it.pathItems[0];
                    if (p && p.closed) return p;
                }
            }
        } catch (e) { }
        return null;
    }

    function _applyTextAttrs(tf, fontObj, sizePt) {
        try {
            if (sizePt > 0) tf.textRange.characterAttributes.size = sizePt;
        } catch (e0) { }
        try {
            if (fontObj) tf.textRange.characterAttributes.textFont = fontObj;
        } catch (e1) { }
    }

    if (useSelectedDummy) {
        // 選択オブジェクトにダミーテキスト：選択した図形（閉じたパス）をエリア内文字に変換してダミー文字を入れる
        var dummyText = (lang === "ja") ? DUMMY_TEXT_JA : DUMMY_TEXT_EN;
        var fontObj = null;
        if (lang === "ja") {
            // User requested: HiraginoSans-W3, 10pt
            fontObj = getTextFontSafe(["HiraginoSans-W3", "Hiragino Sans W3", "HiraginoSans-W3"]);
        } else {
            // User requested: Myriad, 10pt (name differs by environment)
            fontObj = getTextFontSafe(["MyriadPro-Regular", "Myriad Pro Regular", "MyriadPro", "Myriad"]);
        }
        var sizePt = 10;

        for (var d = 0; d < sel.length; d++) {
            var pth = _getClosedPathItem(sel[d]);
            if (!pth) continue;
            try {
                // Convert the selected path itself to Area Type
                var tfD = doc.textFrames.areaText(pth);
                try { tfD.textPath.filled = false; tfD.textPath.stroked = false; } catch (e0) { }
                try { tfD.contents = dummyText; } catch (e1) { }
                _applyTextAttrs(tfD, fontObj, sizePt);
                created.push(tfD);
            } catch (e2) { }
        }

    } else if (useSelected) {
        // 選択オブジェクトを利用
        var srcText = null, destPath = null;
        for (var k = 0; k < sel.length; k++) {
            if (!srcText && sel[k].typename === "TextFrame") { srcText = sel[k]; }
            else if (!destPath && sel[k].typename === "PathItem" && sel[k].closed) { destPath = sel[k]; }
        }
        if (srcText && destPath) {
            try {
                var srcContents = srcText.contents;
                var srcFont = null, srcSize = 0;
                try { srcFont = srcText.textRange.characterAttributes.textFont; srcSize = srcText.textRange.characterAttributes.size; } catch (e) { }
                var dupPath = destPath.duplicate();
                dupPath.filled = false; dupPath.stroked = false;
                var newTf = doc.textFrames.areaText(dupPath);
                newTf.contents = srcContents;
                try { if (srcFont) newTf.textRange.characterAttributes.textFont = srcFont; if (srcSize > 0) newTf.textRange.characterAttributes.size = srcSize; } catch (e) { }
                srcText.remove();
                destPath.remove();
                created.push(newTf);
            } catch (e) { }
        }
    } else if (useSimple) {
        // シンプル：幅+1pt・高さ+1pt の長方形
        for (var sm = sel.length - 1; sm >= 0; sm--) {
            var smItem = sel[sm];
            if (smItem.typename === "TextFrame" && smItem.kind === TextType.POINTTEXT) {
                try {
                    var smB = smItem.geometricBounds;
                    var smRect = doc.pathItems.rectangle(smB[1], smB[0], smB[2] - smB[0] + 1, smB[1] - smB[3] + 1);
                    smRect.filled = false; smRect.stroked = false;
                    var smC = smItem.contents, smFont = null, smSize = 0;
                    try { smFont = smItem.textRange.characterAttributes.textFont; smSize = smItem.textRange.characterAttributes.size; } catch (e) { }
                    var smTf = doc.textFrames.areaText(smRect);
                    smTf.contents = smC;
                    try { if (smFont) smTf.textRange.characterAttributes.textFont = smFont; if (smSize > 0) smTf.textRange.characterAttributes.size = smSize; } catch (e) { }
                    created.push(smTf);
                    smItem.remove();
                } catch (e) { }
            }
        }
    } else if (useButton) {
        // ボタン風：幅×1.2・高さ×1.8、中央揃え
        for (var bt = sel.length - 1; bt >= 0; bt--) {
            var btItem = sel[bt];
            if (btItem.typename === "TextFrame" && btItem.kind === TextType.POINTTEXT) {
                try {
                    var btB = btItem.geometricBounds;
                    var btOW = btB[2] - btB[0], btOH = btB[1] - btB[3];
                    var btW = btOW * 1.2, btH = btOH * 1.8;
                    var btRect = doc.pathItems.rectangle(
                        btB[1] + (btH - btOH) / 2, btB[0] - (btW - btOW) / 2, btW, btH);
                    btRect.filled = false; btRect.stroked = false;
                    var btC = btItem.contents, btFont = null, btSize = 0;
                    try { btFont = btItem.textRange.characterAttributes.textFont; btSize = btItem.textRange.characterAttributes.size; } catch (e) { }
                    var btTf = doc.textFrames.areaText(btRect);
                    btTf.contents = btC;
                    try { if (btFont) btTf.textRange.characterAttributes.textFont = btFont; if (btSize > 0) btTf.textRange.characterAttributes.size = btSize; } catch (e) { }
                    // Horizontal center + vertical center (frame alignment)
                    try { btTf.textRange.paragraphAttributes.justification = Justification.CENTER; } catch (e) { }
                    // NOTE: DOM `verticalAlignment` is not reliably applied for Area Type in some versions.
                    // Use the same dynamic action used in Dialog B (adobe_frameAlignment): 0=Top, 1=Center, 2=Bottom, 3=Justify
                    try { applyAreaTextFrameAlignment(btTf, 1, false); } catch (e2) { }
                    created.push(btTf);
                    btItem.remove();
                } catch (e) { }
            }
        }
    }

    if (created.length > 0) {
        app.activeDocument.selection = created;
        app.redraw();
    }
    return created;
}

// ============================================================
// ダイアログB：エリア内文字 調整
// ============================================================
function showDialogB(doc, initialTf, targetFrames, defaultAlignInt) {
    var rulerInfo = getRulerUnitInfo(doc);

    // Dialog A から受け取った変換結果を確実に対象にする（選択が変わっても崩れないように）
    if (targetFrames && targetFrames.length) {
        try { app.activeDocument.selection = targetFrames; } catch (e) { }
    } else if (initialTf) {
        try { app.activeDocument.selection = [initialTf]; } catch (e2) { }
    }
    try { app.redraw(); } catch (e3) { }

    // Dialog B の調整対象（非分離モード用に固定）。
    // モーダルダイアログ中は selection が変動/取得不能になることがあるため、Aから渡された配列を優先する。
    var _fixedTargets = null;
    if (targetFrames && targetFrames.length) {
        _fixedTargets = targetFrames.slice(0);
    } else {
        try {
            var _sel0 = app.activeDocument.selection;
            if (_sel0 && _sel0.length) {
                _fixedTargets = [];
                for (var _k0 = 0; _k0 < _sel0.length; _k0++) { _fixedTargets.push(_sel0[_k0]); }
            }
        } catch (e4) { }
    }

    function refreshFixedTargetsFromSelection() {
        try {
            var s = app.activeDocument.selection;
            var arr = [];
            if (s && s.length) {
                for (var i = 0; i < s.length; i++) {
                    if (s[i] && s[i].typename === "TextFrame" && s[i].kind === TextType.AREATEXT) {
                        arr.push(s[i]);
                    }
                }
            }
            _fixedTargets = arr.length ? arr : null;
        } catch (e) {
            _fixedTargets = null;
        }
    }

    var dlgB = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
    dlgB.alignChildren = "fill";
    dlgB.margins = 20;

    // 「テキストを分離」パネル
    var pnlAdjust = dlgB.add("panel", undefined, L("pnlSeparate"));
    pnlAdjust.orientation = "row";
    pnlAdjust.alignChildren = ["center", "center"];
    pnlAdjust.margins = [15, 20, 15, 10];

    var radNoAdjust = pnlAdjust.add("radiobutton", undefined, L("radNoAdjust"));
    var radSeparate = pnlAdjust.add("radiobutton", undefined, L("radSeparate"));
    var radPathOpacity = pnlAdjust.add("radiobutton", undefined, L("radPathOpacity"));
    var radRemovePath = pnlAdjust.add("radiobutton", undefined, L("radRemovePath"));
    radNoAdjust.value = true;

    // 2カラムレイアウト
    var grpColumns = dlgB.add("group");
    grpColumns.orientation = "row";
    grpColumns.alignChildren = ["fill", "top"];
    grpColumns.spacing = 10;

    var grpLeft = grpColumns.add("group");
    grpLeft.orientation = "column";
    grpLeft.alignChildren = "fill";

    var grpRight = grpColumns.add("group");
    grpRight.orientation = "column";
    grpRight.alignChildren = "fill";

    // 右カラム：行揃え
    var pnlJustify = grpRight.add("panel", undefined, L("pnlJustify"));
    pnlJustify.alignChildren = "left";
    pnlJustify.margins = [15, 20, 15, 10];
    var radJustLeft = pnlJustify.add("radiobutton", undefined, L("justLeft"));
    var radJustCenter = pnlJustify.add("radiobutton", undefined, L("justCenter"));
    var radJustRight = pnlJustify.add("radiobutton", undefined, L("justRight"));
    var radJustFull = pnlJustify.add("radiobutton", undefined, L("justFull"));
    var radJustFullAll = pnlJustify.add("radiobutton", undefined, L("justFullAll"));
    radJustLeft.value = true;

    // 右カラム：テキストの配置
    var pnlTextAlign = grpRight.add("panel", undefined, L("pnlTextAlign"));
    pnlTextAlign.alignChildren = "left";
    pnlTextAlign.margins = [15, 20, 15, 10];
    var radAlignTop = pnlTextAlign.add("radiobutton", undefined, L("alignTop"));
    var radAlignCenter = pnlTextAlign.add("radiobutton", undefined, L("alignCenter"));
    var radAlignBottom = pnlTextAlign.add("radiobutton", undefined, L("alignBottom"));
    var radAlignJustify = pnlTextAlign.add("radiobutton", undefined, L("alignJustify"));
    // Default alignment
    if (defaultAlignInt === 1) {
        radAlignCenter.value = true;
    } else if (defaultAlignInt === 2) {
        radAlignBottom.value = true;
    } else if (defaultAlignInt === 3) {
        radAlignJustify.value = true;
    } else {
        radAlignTop.value = true;
    }

    // 左カラム：フォントサイズ
    var pnlAutoSize = grpLeft.add("panel", undefined, L("pnlAutoSize"));
    pnlAutoSize.alignChildren = "fill";
    pnlAutoSize.margins = [15, 20, 15, 10];
    var grpFontSize = pnlAutoSize.add("group");
    grpFontSize.alignment = "left";
    grpFontSize.add("statictext", undefined, L("lblFontSize"));
    var etFontSize = grpFontSize.add("edittext", undefined, "");
    etFontSize.characters = 5;
    grpFontSize.add("statictext", undefined, "pt");
    var grpAutoSizeBtns = pnlAutoSize.add("group");
    grpAutoSizeBtns.orientation = "row";
    var btnTextSize = grpAutoSizeBtns.add("button", undefined, L("btnTextSize"));
    var btnFit = grpAutoSizeBtns.add("button", undefined, L("btnFit"));

    // 左カラム：フレームサイズ
    var pnlFrameSize = grpLeft.add("panel", undefined, L("pnlFrameSize"));
    pnlFrameSize.alignChildren = "left";
    pnlFrameSize.margins = [15, 20, 15, 10];
    var grpWidth = pnlFrameSize.add("group");
    var lblWidth = grpWidth.add("statictext", undefined, L("lblWidth"));
    lblWidth.preferredSize.width = 28;
    var etWidth = grpWidth.add("edittext", undefined, "");
    etWidth.characters = 5;
    grpWidth.add("statictext", undefined, rulerInfo.label);
    var etChars = grpWidth.add("edittext", undefined, "");
    etChars.characters = 4;
    var lblChars = grpWidth.add("statictext", undefined, L("lblChars"));
    var grpHeight = pnlFrameSize.add("group");
    var lblHeight = grpHeight.add("statictext", undefined, L("lblHeight"));
    lblHeight.preferredSize.width = 28;
    var etHeight = grpHeight.add("edittext", undefined, "");
    etHeight.characters = 5;
    grpHeight.add("statictext", undefined, rulerInfo.label);
    var btnFrameAuto = grpHeight.add("button", undefined, L("chkFrameAuto"));
    btnFrameAuto.preferredSize.width = 60;
    // 英語UIでは chars 計算は不正確なため使用不可にする
    if (lang !== "ja") {
        etChars.enabled = false;
        lblChars.enabled = false;
    }

    // 左カラム：インデント
    var pnlIndent = grpLeft.add("panel", undefined, L("pnlIndent"));
    pnlIndent.orientation = "row";
    pnlIndent.alignChildren = ["left", "top"];
    pnlIndent.margins = [15, 20, 15, 10];
    var grpIndentLeft = pnlIndent.add("group");
    grpIndentLeft.orientation = "column";
    grpIndentLeft.alignChildren = "left";
    var grpLeftIndent = grpIndentLeft.add("group");
    var chkLeftIndent = grpLeftIndent.add("checkbox", undefined, L("indentLeft"));
    chkLeftIndent.preferredSize.width = 40;
    var etLeftIndent = grpLeftIndent.add("edittext", undefined, "0");
    etLeftIndent.characters = 4;
    grpLeftIndent.add("statictext", undefined, rulerInfo.label);
    etLeftIndent.enabled = false;
    var grpRightIndent = grpIndentLeft.add("group");
    var chkRightIndent = grpRightIndent.add("checkbox", undefined, L("indentRight"));
    chkRightIndent.preferredSize.width = 40;
    var etRightIndent = grpRightIndent.add("edittext", undefined, "0");
    etRightIndent.characters = 4;
    grpRightIndent.add("statictext", undefined, rulerInfo.label);
    etRightIndent.enabled = false;
    var grpIndentRight = pnlIndent.add("group");
    grpIndentRight.orientation = "column";
    grpIndentRight.alignChildren = "left";
    grpIndentRight.alignment = ["left", "center"];
    var chkSync = grpIndentRight.add("checkbox", undefined, L("chkSync"));

    // 左カラム：オプション
    var pnlOptions = grpLeft.add("panel", undefined, L("pnlOptions"));
    pnlOptions.alignChildren = "left";
    pnlOptions.margins = [15, 20, 15, 10];
    var grpMargin = pnlOptions.add("group");
    var chkMargin = grpMargin.add("checkbox", undefined, L("chkMargin"));
    var etMargin = grpMargin.add("edittext", undefined, "0");
    etMargin.characters = 6;
    var lblUnit = grpMargin.add("statictext", undefined, rulerInfo.label);
    etMargin.enabled = false;
    lblUnit.enabled = false;
    var chkLeader = pnlOptions.add("checkbox", undefined, L("chkLeader"));

    // ボタンエリア
    var grpBottom = dlgB.add("group");
    grpBottom.orientation = "row";
    grpBottom.alignment = "fill";
    grpBottom.alignChildren = ["fill", "center"];
    var chkPreview = grpBottom.add("checkbox", undefined, L("chkPreview"));
    chkPreview.alignment = ["left", "center"];
    var btnGroupB = grpBottom.add("group");
    btnGroupB.alignment = ["right", "center"];
    var btnCloseB = btnGroupB.add("button", undefined, L("btnClose"), { name: "cancel" });
    var btnRun = btnGroupB.add("button", undefined, L("btnRun"), { name: "ok" });

    // 状態変数
    var isPreviewActive = false;
    var _autoSizeMode = "none";
    var hasMultiParagraph = false;
    var _frameAutoOn = false;
    var fontSize = 0;

    // --- 入力バリデーション（幅/高さ） ---
    // 0以下・NaN・極端値を弾いてIllustratorの不安定化を避ける
    var _lastValidWidth = null; // ruler units
    var _lastValidHeight = null; // ruler units

    function _maxSizeInRulerUnits() {
        // 上限は pt で固定（極端値を防ぐ）。表示単位に合わせて換算。
        // 100000pt は現実的に十分大きく、かつ事故りにくい上限。
        return 100000 / rulerInfo.toPt;
    }

    function validateSizeField(editText, lastValue, fieldName) {
        var raw = String(editText.text);
        var v = parseFloat(raw);
        if (isNaN(v) || !isFinite(v)) {
            if (lastValue !== null) editText.text = lastValue;
            return null;
        }
        if (v <= 0) {
            if (lastValue !== null) editText.text = lastValue;
            return null;
        }
        var maxV = _maxSizeInRulerUnits();
        if (v > maxV) {
            v = maxV;
            editText.text = Math.round(v * 100) / 100;
        }
        // 小さすぎる値も事故の元なので下限を設ける
        if (v < 0.01) {
            v = 0.01;
            editText.text = Math.round(v * 100) / 100;
        }
        return v;
    }

    var allRadios = [radSeparate, radPathOpacity, radRemovePath, radNoAdjust];
    var allJustRadios = [radJustLeft, radJustCenter, radJustRight, radJustFull, radJustFullAll];
    var allAlignRadios = [radAlignTop, radAlignCenter, radAlignBottom, radAlignJustify];

    // --- ローカル関数 ---
    function updateFrameAutoButtonLabel() {
        btnFrameAuto.text = (_frameAutoOn ? "✓ " : "") + L("chkFrameAuto");
    }

    function getAdjustmentPt() {
        var mp = chkMargin.value ? (parseFloat(etMargin.text) || 0) * rulerInfo.toPt : 0;
        var lp = (chkLeftIndent.value || chkSync.value) ? (parseFloat(etLeftIndent.text) || 0) * rulerInfo.toPt : 0;
        var rp = chkSync.value ? lp : (chkRightIndent.value ? (parseFloat(etRightIndent.text) || 0) * rulerInfo.toPt : 0);
        return 2 * mp + lp + rp;
    }

    function loadValuesFromFrame(tf0) {
        try { hasMultiParagraph = (tf0.paragraphs && tf0.paragraphs.length >= 2); } catch (e) { hasMultiParagraph = false; }
        var initW = tf0.textPath.width / rulerInfo.toPt;
        var initH = tf0.textPath.height / rulerInfo.toPt;
        fontSize = 0;
        try { fontSize = tf0.textRange.characterAttributes.size || 0; } catch (e) { }
        if (fontSize > 0) { etFontSize.text = Math.round(fontSize * 100) / 100; }
        etWidth.text = Math.round(initW * 100) / 100;
        etHeight.text = Math.round(initH * 100) / 100;
        _lastValidWidth = parseFloat(etWidth.text);
        _lastValidHeight = parseFloat(etHeight.text);
        try {
            var ij = tf0.paragraphs.length > 0 ? tf0.paragraphs[0].justification : Justification.LEFT;
            radJustLeft.value = (ij === Justification.LEFT);
            radJustCenter.value = (ij === Justification.CENTER);
            radJustRight.value = (ij === Justification.RIGHT);
            radJustFull.value = (ij === Justification.FULLJUSTIFYLASTLINELEFT);
            radJustFullAll.value = (ij === Justification.FULLJUSTIFY);
            if (!radJustLeft.value && !radJustCenter.value && !radJustRight.value &&
                !radJustFull.value && !radJustFullAll.value) { radJustLeft.value = true; }
        } catch (e) { radJustLeft.value = true; }
        try {
            var sp = tf0.spacing || 0;
            etMargin.text = Math.round((sp / rulerInfo.toPt) * 100) / 100;
            chkMargin.value = (sp !== 0);
            etMargin.enabled = chkMargin.value;
            lblUnit.enabled = chkMargin.value;
        } catch (e) { }
        try {
            var ilp = tf0.paragraphs.length > 0 ? (tf0.paragraphs[0].leftIndent || 0) : 0;
            var irp = tf0.paragraphs.length > 0 ? (tf0.paragraphs[0].rightIndent || 0) : 0;
            chkLeftIndent.value = (ilp !== 0);
            etLeftIndent.enabled = chkLeftIndent.value;
            etLeftIndent.text = chkLeftIndent.value ? Math.round((ilp / rulerInfo.toPt) * 100) / 100 : "0";
            chkRightIndent.value = (irp !== 0);
            etRightIndent.enabled = chkRightIndent.value;
            etRightIndent.text = chkRightIndent.value ? Math.round((irp / rulerInfo.toPt) * 100) / 100 : "0";
        } catch (e) { }
        if (fontSize > 0) {
            etChars.text = Math.round(((tf0.textPath.width - getAdjustmentPt()) / fontSize) * 100) / 100;
        }
        btnTextSize.enabled = !hasMultiParagraph;
        btnFit.enabled = !hasMultiParagraph;
    }

    function runAdjust(forPreview) {
        var doSeparate = radSeparate.value;
        var doPathOpacity = radPathOpacity.value;
        var doRemovePath = radRemovePath.value;
        var doTextSize = (_autoSizeMode === "textsize");
        var doFit = (_autoSizeMode === "fit");
        var doFrameSize = _frameAutoOn;
        var doLeader = chkLeader.value;

        var justValue = Justification.LEFT;
        if (radJustCenter.value) justValue = Justification.CENTER;
        else if (radJustRight.value) justValue = Justification.RIGHT;
        else if (radJustFull.value) justValue = Justification.FULLJUSTIFYLASTLINELEFT;
        else if (radJustFullAll.value) justValue = Justification.FULLJUSTIFY;

        var alignValueInt = 0;
        if (radAlignCenter.value) alignValueInt = 1;
        else if (radAlignBottom.value) alignValueInt = 2;
        else if (radAlignJustify.value) alignValueInt = 3;

        var leftIndentPt = (chkLeftIndent.value || chkSync.value)
            ? (parseFloat(etLeftIndent.text) || 0) * rulerInfo.toPt : 0;
        var rightIndentPt = chkSync.value
            ? leftIndentPt
            : (chkRightIndent.value ? (parseFloat(etRightIndent.text) || 0) * rulerInfo.toPt : 0);
        var marginPt = chkMargin.value ? (parseFloat(etMargin.text) || 0) * rulerInfo.toPt : 0;

        var isSeparationMode = doSeparate || doPathOpacity || doRemovePath;
        var savedSel = [];
        if (!isSeparationMode) {
            var origSel = app.activeDocument.selection;
            for (var s = 0; s < origSel.length; s++) { savedSel.push(origSel[s]); }
        }

        // 分離モードはオブジェクトが置き換わるため selection を使う。
        // 非分離モードは固定ターゲットを優先して、プレビューが確実に反映されるようにする。
        var targets = null;
        if (isSeparationMode) {
            targets = app.activeDocument.selection;
        } else {
            if (!(_fixedTargets && _fixedTargets.length)) { refreshFixedTargetsFromSelection(); }
            targets = _fixedTargets && _fixedTargets.length ? _fixedTargets : app.activeDocument.selection;
        }
        for (var i = targets.length - 1; i >= 0; i--) {
            var obj = targets[i];
            if (obj.typename === "TextFrame" && obj.kind === TextType.AREATEXT) {
                if (isSeparationMode) {
                    var bounds = obj.geometricBounds;
                    var left = bounds[0] - marginPt, top = bounds[1] + marginPt;
                    var right = bounds[2] + marginPt, bottom = bounds[3] - marginPt;
                    var rect = doc.pathItems.rectangle(top, left, right - left, top - bottom);
                    rect.filled = false; rect.stroked = true;
                    var newText = doc.textFrames.add();
                    newText.contents = obj.contents;
                    newText.position = [left, top - newText.textRange.characterAttributes.size];
                    var ca = obj.textRange.characterAttributes, nca = newText.textRange.characterAttributes;
                    if (ca.textFont && ca.textFont.name) { nca.textFont = ca.textFont; }
                    nca.size = ca.size; nca.leading = ca.leading;
                    obj.remove();
                    if (doSeparate) {
                        var blk = new CMYKColor(); blk.cyan = 0; blk.magenta = 0; blk.yellow = 0; blk.black = 100;
                        rect.stroked = true; rect.strokeColor = blk; rect.strokeWidth = 1; rect.filled = false;
                    }
                    if (doPathOpacity) { rect.stroked = false; rect.filled = false; }
                    if (doRemovePath) { rect.remove(); }
                    try { var pa = newText.textRange.paragraphAttributes; pa.justification = justValue; pa.leftIndent = leftIndentPt; pa.rightIndent = rightIndentPt; } catch (e) { }
                    if (doLeader) {
                        radJustLeft.value = false; radJustCenter.value = false; radJustRight.value = true;
                        radJustFull.value = false; radJustFullAll.value = false;
                        applyLeaderTab(newText);
                    }
                } else {
                    try { obj.spacing = marginPt; } catch (e) { }
                    // 幅/高さ：NaN・0以下・極端値をガード
                    var wRu = validateSizeField(etWidth, _lastValidWidth, "width");
                    var hRu = validateSizeField(etHeight, _lastValidHeight, "height");
                    if (wRu !== null) { _lastValidWidth = wRu; try { obj.textPath.width = wRu * rulerInfo.toPt; } catch (e) { } }
                    if (hRu !== null) { _lastValidHeight = hRu; try { obj.textPath.height = hRu * rulerInfo.toPt; } catch (e) { } }
                    if (doTextSize) { shrinkFont(obj); } else if (doFit) { fitTextToFrame(obj); }
                    // プレビュー中は app.doScript 経由の処理を走らせない（不安定化・クラッシュ回避）
                    if (doFrameSize && !forPreview) { expandFrameToFit(obj); }
                    try { var pa2 = obj.textRange.paragraphAttributes; pa2.justification = justValue; pa2.leftIndent = leftIndentPt; pa2.rightIndent = rightIndentPt; } catch (e) { }
                    applyAreaTextFrameAlignment(obj, alignValueInt, forPreview);
                    if (doLeader) {
                        radJustLeft.value = false; radJustCenter.value = false; radJustRight.value = true;
                        radJustFull.value = false; radJustFullAll.value = false;
                        applyLeaderTab(obj);
                    }
                }
            }
        }
        if (savedSel.length > 0) { try { app.activeDocument.selection = savedSel; } catch (e) { } }
        app.redraw();
    }

    function updatePreview() {
        if (isPreviewActive) {
            try { app.undo(); } catch (e) { }
            try { app.redraw(); } catch (e2) { }
            isPreviewActive = false;

            // undo 後は参照が無効化されるため、対象を取り直す
            refreshFixedTargetsFromSelection();
        }
        if (chkPreview.value) {
            // プレビュー適用前にも念のため取り直す（モード切替直後の選択揺れ対策）
            if (!(_fixedTargets && _fixedTargets.length)) { refreshFixedTargetsFromSelection(); }
            runAdjust(true);
            isPreviewActive = true;
        }
    }

    function updateSeparationDim() {
        var sep = radSeparate.value || radPathOpacity.value || radRemovePath.value;
        pnlJustify.enabled = !sep;
        pnlTextAlign.enabled = !sep;
        pnlAutoSize.enabled = !sep;
        pnlFrameSize.enabled = !sep;
        btnFrameAuto.enabled = !sep;
        pnlIndent.enabled = !sep;
        pnlOptions.enabled = !sep;
        if (pnlAutoSize.enabled) { btnTextSize.enabled = !hasMultiParagraph; btnFit.enabled = !hasMultiParagraph; }
        else { btnTextSize.enabled = false; btnFit.enabled = false; }
    }

    function updateMarginEnabled() {
        chkMargin.enabled = radNoAdjust.value;
        if (!radNoAdjust.value) { etMargin.enabled = false; lblUnit.enabled = false; }
        else { etMargin.enabled = chkMargin.value; lblUnit.enabled = chkMargin.value; }
    }

    function applyFontSizeFromField() {
        var newSize = parseFloat(etFontSize.text) || 0;
        if (newSize <= 0) return;
        fontSize = newSize;
        var tgts = app.activeDocument.selection;
        for (var i = 0; i < tgts.length; i++) {
            if (tgts[i].typename === "TextFrame" && tgts[i].kind === TextType.AREATEXT) {
                try { tgts[i].textRange.characterAttributes.size = newSize; } catch (e) { }
            }
        }
        if (fontSize > 0) {
            etChars.text = Math.round((((parseFloat(etWidth.text) || 0) * rulerInfo.toPt - getAdjustmentPt()) / fontSize) * 100) / 100;
        }
        app.redraw();
    }

    function onWidthChange() {
        var wRu = validateSizeField(etWidth, _lastValidWidth, "width");
        if (wRu !== null) { _lastValidWidth = wRu; }
        if (fontSize > 0) {
            etChars.text = Math.round((((parseFloat(etWidth.text) || 0) * rulerInfo.toPt - getAdjustmentPt()) / fontSize) * 100) / 100;
        }
        updatePreview();
    }
    function onCharsChange() {
        if (fontSize > 0) {
            var nextW = (((parseFloat(etChars.text) || 0) * fontSize + getAdjustmentPt()) / rulerInfo.toPt);
            if (!isNaN(nextW) && isFinite(nextW) && nextW > 0) {
                etWidth.text = Math.round(nextW * 100) / 100;
                var wRu = validateSizeField(etWidth, _lastValidWidth, "width");
                if (wRu !== null) { _lastValidWidth = wRu; }
            }
        }
        updatePreview();
    }
    function onAdjustmentChange() {
        if (chkSync.value) { etRightIndent.text = etLeftIndent.text; }
        onWidthChange();
    }

    function addJustifyKeyHandler(dialog) {
        dialog.addEventListener("keydown", function (event) {
            if (!event || !event.keyName) return;
            var k = String(event.keyName).toUpperCase();
            var target = null;
            if (k === "L") target = radJustLeft;
            else if (k === "C") target = radJustCenter;
            else if (k === "R") target = radJustRight;
            else if (k === "J") target = radJustFull;
            else if (k === "F") target = radJustFullAll;
            if (!target) return;
            for (var i = 0; i < allJustRadios.length; i++) { allJustRadios[i].value = (allJustRadios[i] === target); }
            try { if (chkLeader && chkLeader.value && target !== radJustRight) { chkLeader.value = false; } } catch (e) { }
            event.preventDefault();
            updatePreview();
        });
    }

    // --- イベントハンドラ ---
    for (var r = 0; r < allRadios.length; r++) {
        (function (btn) {
            btn.onClick = function () {
                for (var j = 0; j < allRadios.length; j++) { allRadios[j].value = (allRadios[j] === btn); }
                updateSeparationDim();
                updateMarginEnabled();
                updatePreview();
            };
        })(allRadios[r]);
    }
    for (var jr = 0; jr < allJustRadios.length; jr++) {
        (function (btn) {
            btn.onClick = function () {
                for (var j = 0; j < allJustRadios.length; j++) { allJustRadios[j].value = (allJustRadios[j] === btn); }
                try { if (chkLeader && chkLeader.value && btn !== radJustRight) { chkLeader.value = false; } } catch (e) { }
                updatePreview();
            };
        })(allJustRadios[jr]);
    }
    for (var ar = 0; ar < allAlignRadios.length; ar++) {
        (function (btn) {
            btn.onClick = function () {
                for (var j = 0; j < allAlignRadios.length; j++) { allAlignRadios[j].value = (allAlignRadios[j] === btn); }
                updatePreview();
            };
        })(allAlignRadios[ar]);
    }
    addJustifyKeyHandler(dlgB);

    btnTextSize.onClick = function () { _autoSizeMode = "textsize"; runAdjust(false); _autoSizeMode = "none"; };
    btnFit.onClick = function () { _autoSizeMode = "fit"; runAdjust(false); _autoSizeMode = "none"; };
    chkPreview.onClick = updatePreview;
    btnFrameAuto.onClick = function () {
        _frameAutoOn = !_frameAutoOn;
        updateFrameAutoButtonLabel();
        if (!_frameAutoOn) {
            try {
                var t = app.activeDocument.selection;
                for (var i = 0; i < t.length; i++) {
                    if (t[i] && t[i].typename === "TextFrame" && t[i].kind === TextType.AREATEXT) { collapseFrameAuto(t[i]); }
                }
            } catch (e) { }
        }
        updatePreview();
    };
    chkLeader.onClick = updatePreview;
    chkMargin.onClick = function () {
        etMargin.enabled = chkMargin.value;
        lblUnit.enabled = chkMargin.value;
        if (chkMargin.value) { etMargin.text = "1"; }
        onAdjustmentChange();
    };
    chkSync.onClick = function () {
        if (chkSync.value) {
            chkLeftIndent.value = true; etLeftIndent.enabled = true;
            chkRightIndent.enabled = false; etRightIndent.enabled = false;
            etRightIndent.text = etLeftIndent.text;
        } else {
            chkRightIndent.enabled = true; etRightIndent.enabled = chkRightIndent.value;
        }
        onAdjustmentChange();
    };
    chkLeftIndent.onClick = function () {
        if (!chkSync.value) { etLeftIndent.enabled = chkLeftIndent.value; }
        if (!chkLeftIndent.value) { etLeftIndent.text = "0"; }
        onAdjustmentChange();
    };
    chkRightIndent.onClick = function () {
        etRightIndent.enabled = chkRightIndent.value;
        if (!chkRightIndent.value) { etRightIndent.text = "0"; }
        onAdjustmentChange();
    };

    etFontSize.onChange = applyFontSizeFromField;
    etMargin.onChange = onAdjustmentChange;
    etWidth.onChange = onWidthChange;
    etHeight.onChange = function () {
        var hRu = validateSizeField(etHeight, _lastValidHeight, "height");
        if (hRu !== null) { _lastValidHeight = hRu; }
        updatePreview();
    };
    etChars.onChange = onCharsChange;
    etLeftIndent.onChange = onAdjustmentChange;
    etRightIndent.onChange = onAdjustmentChange;
    changeValueByArrowKey(etFontSize, false, applyFontSizeFromField);
    changeValueByArrowKey(etMargin, false, onAdjustmentChange);
    changeValueByArrowKey(etWidth, false, onWidthChange);
    changeValueByArrowKey(etHeight, false, updatePreview);
    changeValueByArrowKey(etChars, false, onCharsChange);
    changeValueByArrowKey(etLeftIndent, false, onAdjustmentChange);
    changeValueByArrowKey(etRightIndent, false, onAdjustmentChange);

    btnRun.onClick = function () {
        saveDlgPosition(dlgB);
        if (isPreviewActive) { try { app.undo(); } catch (e) { } try { app.redraw(); } catch (e2) { } isPreviewActive = false; }
        runAdjust(false);
        dlgB.close(1);
    };
    btnCloseB.onClick = function () {
        saveDlgPosition(dlgB);
        if (isPreviewActive) { app.undo(); app.redraw(); }
        dlgB.close(0);
    };

    // 初期値読み込み
    if (initialTf) { loadValuesFromFrame(initialTf); }

    updateSeparationDim();
    updateFrameAutoButtonLabel();

    // ダイアログBを開いたらプレビューON
    chkPreview.value = true;
    updatePreview();

    restoreDlgPosition(dlgB);
    dlgB.onClose = function () { saveDlgPosition(dlgB); return true; };
    dlgB.show();
}

// ============================================================
// エントリポイント
// ============================================================
if (app.documents.length > 0) {
    var doc = app.activeDocument;
    var sel = doc.selection;

    if (sel && sel.length > 0) {
        var _hasPoint = false, _hasArea = false, _hasPath = false;
        for (var _ei = 0; _ei < sel.length; _ei++) {
            if (sel[_ei].typename === "TextFrame") {
                if (sel[_ei].kind === TextType.POINTTEXT || sel[_ei].kind === TextType.PATHTEXT) _hasPoint = true;
                if (sel[_ei].kind === TextType.AREATEXT) _hasArea = true;
            }
            if (sel[_ei].typename === "PathItem" || sel[_ei].typename === "CompoundPathItem") {
                _hasPath = true;
            }
        }

        if (_hasPoint || _hasPath) {
            // ポイント文字 または 図形 → ダイアログA
            showDialogA(doc, sel);
        } else if (_hasArea) {
            // エリア内文字のみ → ダイアログB
            var _firstArea = null;
            for (var _fi = 0; _fi < sel.length; _fi++) {
                if (sel[_fi].typename === "TextFrame" && sel[_fi].kind === TextType.AREATEXT) {
                    _firstArea = sel[_fi]; break;
                }
            }
            showDialogB(doc, _firstArea, null, null);
        } else {
            alert(L("alertSelectText"));
        }
    } else {
        alert(L("alertSelectText"));
    }
} else {
    alert(L("alertNoDocument"));
}