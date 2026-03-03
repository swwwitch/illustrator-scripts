#target illustrator
// Use a persistent session engine for dialog position memory
#targetengine "AreaTypeToolkit"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
    作成日：2025-03-06
    更新日：2026-03-03
    説明：選択されたエリア内文字と同じサイズ・座標の長方形を作成し、
          そのエリア内文字のテキストを新しいポイント文字として配置し、
          元のエリア内文字を削除するスクリプト。
*/

// --- Script version / スクリプトバージョン ---
var SCRIPT_VERSION = "v1.0.1";

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
    btnCancel: { ja: "キャンセル", en: "Cancel" },
    btnOk: { ja: "OK", en: "OK" },

    // Alerts
    alertSelectAreaText: { ja: "エリア内文字を選択してください。", en: "Please select area text." },
    alertNoDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." }
};

function L(key) {
    try {
        if (LABELS[key] && LABELS[key][lang]) return LABELS[key][lang];
        if (LABELS[key] && LABELS[key].en) return LABELS[key].en;
    } catch (e) { }
    return key;
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

    var str = '/version 3' + '/name [ 8' + ' 4172656154797065' + ']' + '/isOpen 1' + '/actionCount 1' + '/action-1 {' + ' /name [ 8' + ' 4175746f53697a65' + ' ]' + ' /keyIndex 0' + ' /colorIndex 0' + ' /isOpen 1' + ' /eventCount 1' + ' /event-1 {' + ' /useRulersIn1stQuadrant 0' + ' /internalName (adobe_SLOAreaTextDialog)' + ' /localizedName [ 33' + ' e382a8e383aae382a2e58685e69687e5ad97e382aae38397e382b7e383a7e383b3' + ' ]' + ' /isOpen 1' + ' /isOn 1' + ' /hasDialog 0' + ' /parameterCount 1' + ' /parameter-1 {' + ' /key 1952539754' + ' /showInPalette 4294967295' + ' /type (integer)' + ' /value ' + valueInt + ' }' + ' }' + '}';

    var f = new File('~/ScriptAction.aia');
    f.open('w');
    f.write(str);
    f.close();
    app.loadAction(f);
    f.remove();

    // Action name: AutoSize / Set name: AreaType
    app.doScript("AutoSize", "AreaType", false);
    app.unloadAction("AreaType", "");
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

// Adobe Illustratorのドキュメントが開かれているか確認
if (app.documents.length > 0) {
    var doc = app.activeDocument;
    var sel = doc.selection;

    var rulerInfo = getRulerUnitInfo(doc);
    var fontSize = 0; // 文字サイズ（pt）、初期値ループで設定

    if (sel.length > 0) {

        // ダイアログボックスを作成
        var dlg = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
        dlg.alignChildren = "fill";
        dlg.margins = 20;

        // 「テキストを分離」パネル（全幅・最上段）
        var pnlAdjust = dlg.add("panel", undefined, L("pnlSeparate"));
        pnlAdjust.orientation = "row";
        pnlAdjust.alignChildren = ["center", "center"];
        pnlAdjust.margins = [15, 20, 15, 10];

        var radNoAdjust = pnlAdjust.add("radiobutton", undefined, L("radNoAdjust"));
        var radSeparate = pnlAdjust.add("radiobutton", undefined, L("radSeparate"));
        var radPathOpacity = pnlAdjust.add("radiobutton", undefined, L("radPathOpacity"));
        var radRemovePath = pnlAdjust.add("radiobutton", undefined, L("radRemovePath"));
        radNoAdjust.value = true; // 開いたときのデフォルト

        // 2カラムレイアウト
        var grpColumns = dlg.add("group");
        grpColumns.orientation = "row";
        grpColumns.alignChildren = ["fill", "top"];
        grpColumns.spacing = 10;

        // 左カラム
        var grpLeft = grpColumns.add("group");
        grpLeft.orientation = "column";
        grpLeft.alignChildren = "fill";

        // 右カラム
        var grpRight = grpColumns.add("group");
        grpRight.orientation = "column";
        grpRight.alignChildren = "fill";

        // 「行揃え」パネル（右）
        var pnlJustify = grpRight.add("panel", undefined, L("pnlJustify"));
        pnlJustify.alignChildren = "left";
        pnlJustify.margins = [15, 20, 15, 10];

        var radJustLeft = pnlJustify.add("radiobutton", undefined, L("justLeft"));
        var radJustCenter = pnlJustify.add("radiobutton", undefined, L("justCenter"));
        var radJustRight = pnlJustify.add("radiobutton", undefined, L("justRight"));
        var radJustFull = pnlJustify.add("radiobutton", undefined, L("justFull"));
        var radJustFullAll = pnlJustify.add("radiobutton", undefined, L("justFullAll"));
        radJustLeft.value = true; // デフォルト（初期値ループで上書き）

        // 「テキストの配置」パネル（右）
        var pnlTextAlign = grpRight.add("panel", undefined, L("pnlTextAlign"));
        pnlTextAlign.alignChildren = "left";
        pnlTextAlign.margins = [15, 20, 15, 10];

        var radAlignTop = pnlTextAlign.add("radiobutton", undefined, L("alignTop"));
        var radAlignCenter = pnlTextAlign.add("radiobutton", undefined, L("alignCenter"));
        var radAlignBottom = pnlTextAlign.add("radiobutton", undefined, L("alignBottom"));
        var radAlignJustify = pnlTextAlign.add("radiobutton", undefined, L("alignJustify"));
        radAlignTop.value = true; // デフォルト

        // 「フォントサイズの調整」パネル（左）
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

        // 「フレームサイズの調整」パネル（左）
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
        grpWidth.add("statictext", undefined, L("lblChars"));

        var grpHeight = pnlFrameSize.add("group");
        var lblHeight = grpHeight.add("statictext", undefined, L("lblHeight"));
        lblHeight.preferredSize.width = 28;
        var etHeight = grpHeight.add("edittext", undefined, "");
        etHeight.characters = 5;
        grpHeight.add("statictext", undefined, rulerInfo.label);
        var btnFrameAuto = grpHeight.add("button", undefined, L("chkFrameAuto"));
        btnFrameAuto.preferredSize.width = 60;

        // 「インデント」パネル（左）
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

        // 「オプション」パネル（左）
        var pnlOptions = grpLeft.add("panel", undefined, L("pnlOptions"));
        pnlOptions.alignChildren = "left";
        pnlOptions.margins = [15, 20, 15, 10];

        var grpMargin = pnlOptions.add("group");
        var chkMargin = grpMargin.add("checkbox", undefined, L("chkMargin"));
        var etMargin = grpMargin.add("edittext", undefined, "0");
        etMargin.characters = 6;
        var lblUnit = grpMargin.add("statictext", undefined, rulerInfo.label);

        // 初期状態：チェックOFFなのでディム
        etMargin.enabled = false;
        lblUnit.enabled = false;

        var chkLeader = pnlOptions.add("checkbox", undefined, L("chkLeader"));

        // 全ラジオボタン（排他制御用）
        var allRadios = [radSeparate, radPathOpacity, radRemovePath, radNoAdjust];
        var allJustRadios = [radJustLeft, radJustCenter, radJustRight, radJustFull, radJustFullAll];

        // ボタンエリア（プレビュー左寄せ／ボタン右寄せ）
        var grpBottom = dlg.add("group");
        grpBottom.orientation = "row";
        grpBottom.alignment = "fill";
        grpBottom.alignChildren = ["fill", "center"];

        var chkPreview = grpBottom.add("checkbox", undefined, L("chkPreview"));
        chkPreview.alignment = ["left", "center"];

        var btnGroup = grpBottom.add("group");
        btnGroup.alignment = ["right", "center"];
        var btnCancel = btnGroup.add("button", undefined, L("btnCancel"), { name: "cancel" });
        var btnOk = btnGroup.add("button", undefined, L("btnOk"), { name: "ok" });

        var isPreviewActive = false;
        var _autoSizeMode = "none"; // "textsize" | "fit" | "none"
        var hasMultiParagraph = false; // 2段落以上なら「文字あふれ」「フィット」をディム表示

        var _frameAutoOn = false; // フレームサイズ「自動」ボタンの状態

        function updateFrameAutoButtonLabel() {
            btnFrameAuto.text = (_frameAutoOn ? "✓ " : "") + L("chkFrameAuto");
        }

        // 処理本体（forPreview=true のときプレビュー向け安全モード）
        function runProcess(forPreview) {
            var doSeparate = radSeparate.value;
            var doPathOpacity = radPathOpacity.value;
            var doRemovePath = radRemovePath.value;
            var doNoAdjust = radNoAdjust.value;
            var doTextSize = (_autoSizeMode === "textsize");
            var doFit = (_autoSizeMode === "fit");
            var doFrameSize = _frameAutoOn; // フレームサイズ「自動」ボタンの状態
            var doLeader = chkLeader.value;

            // 行揃え：選択されているラジオの値を取得（明示的なif/elseで安全に処理）
            var justValue = Justification.LEFT;
            if (radJustCenter.value) {
                justValue = Justification.CENTER;
            } else if (radJustRight.value) {
                justValue = Justification.RIGHT;
            } else if (radJustFull.value) {
                justValue = Justification.FULLJUSTIFYLASTLINELEFT;
            } else if (radJustFullAll.value) {
                justValue = Justification.FULLJUSTIFY;
            }

            // インデント：チェックOFFの側は常に0として扱う（必ず適用）
            // When unchecked, treat as 0 (always apply)
            var leftIndentPt = (chkLeftIndent.value || chkSync.value)
                ? (parseFloat(etLeftIndent.text) || 0) * rulerInfo.toPt
                : 0;
            var rightIndentPt = chkSync.value
                ? leftIndentPt
                : (chkRightIndent.value ? (parseFloat(etRightIndent.text) || 0) * rulerInfo.toPt : 0);

            // 外側からの間隔：ルーラー単位 → pt 変換（チェックOFFなら0）
            var marginPt = chkMargin.value ? (parseFloat(etMargin.text) || 0) * rulerInfo.toPt : 0;

            // 分離モードは obj.remove() で参照が無効になるため選択状態を保存しない
            var isSeparationMode = doSeparate || doPathOpacity || doRemovePath;
            var savedSel = [];
            if (!isSeparationMode) {
                var origSel = app.activeDocument.selection;
                for (var s = 0; s < origSel.length; s++) { savedSel.push(origSel[s]); }
            }

            var targets = app.activeDocument.selection;
            for (var i = targets.length - 1; i >= 0; i--) {
                var obj = targets[i];
                if (obj.typename === "TextFrame" && obj.kind === TextType.AREATEXT) {

                    if (isSeparationMode) {
                        // --- テキストを分離 ---
                        var bounds = obj.geometricBounds;
                        var left = bounds[0] - marginPt;
                        var top = bounds[1] + marginPt;
                        var right = bounds[2] + marginPt;
                        var bottom = bounds[3] - marginPt;
                        var width = right - left;
                        var height = top - bottom;

                        var rect = doc.pathItems.rectangle(top, left, width, height);
                        rect.filled = false;
                        rect.stroked = true;

                        var newText = doc.textFrames.add();
                        newText.contents = obj.contents;
                        newText.position = [left, top - newText.textRange.characterAttributes.size];

                        var charAttrs = obj.textRange.characterAttributes;
                        var newCharAttrs = newText.textRange.characterAttributes;
                        if (charAttrs.textFont && charAttrs.textFont.name) {
                            newCharAttrs.textFont = charAttrs.textFont;
                        }
                        newCharAttrs.size = charAttrs.size;
                        newCharAttrs.leading = charAttrs.leading;

                        obj.remove(); // ← ここで obj の参照が無効になる

                        if (doSeparate) {
                            var blackColor = new CMYKColor();
                            blackColor.cyan = 0; blackColor.magenta = 0;
                            blackColor.yellow = 0; blackColor.black = 100;
                            rect.stroked = true; rect.strokeColor = blackColor;
                            rect.strokeWidth = 1; rect.filled = false;
                        }
                        if (doPathOpacity) { rect.stroked = false; rect.filled = false; }
                        if (doRemovePath) { rect.remove(); }

                        // 行揃え・インデント・リーダー罫を新しいポイント文字に適用
                        try {
                            var paAll = newText.textRange.paragraphAttributes;
                            paAll.justification = justValue;
                            paAll.leftIndent = leftIndentPt;
                            paAll.rightIndent = rightIndentPt;
                        } catch (e) { }
                        if (doLeader) {
                            // リーダー罫は右揃え前提のため、UIも「右」に切り替える / Leader tabs assume right align
                            radJustLeft.value = false;
                            radJustCenter.value = false;
                            radJustRight.value = true;
                            radJustFull.value = false;
                            radJustFullAll.value = false;
                            applyLeaderTab(newText);
                        }

                    } else {
                        // --- フォントサイズ / フレームサイズの調整 ---
                        try { obj.spacing = marginPt; } catch (e) { }

                        // 幅・高さを指定値に変更（空欄はスキップ）
                        var wPt = parseFloat(etWidth.text) || 0;
                        var hPt = parseFloat(etHeight.text) || 0;
                        if (wPt > 0) { try { obj.textPath.width = wPt * rulerInfo.toPt; } catch (e) { } }
                        if (hPt > 0) { try { obj.textPath.height = hPt * rulerInfo.toPt; } catch (e) { } }

                        if (doTextSize) {
                            shrinkFont(obj);
                        } else if (doFit) {
                            fitTextToFrame(obj);
                        }
                        // doNoAdjust のときはフォントサイズ変更なし

                        // 自動サイズ調整（独立チェックボックス）
                        // app.doScript はプレビュー中にダイアログから呼ぶと不安定なためスキップ
                        if (doFrameSize) {
                            expandFrameToFit(obj);
                        }

                        // 行揃え・インデント・リーダー罫をエリア内文字に適用
                        try {
                            var paAll2 = obj.textRange.paragraphAttributes;
                            paAll2.justification = justValue;
                            paAll2.leftIndent = leftIndentPt;
                            paAll2.rightIndent = rightIndentPt;
                        } catch (e) { }
                        if (doLeader) {
                            // リーダー罫は右揃え前提のため、UIも「右」に切り替える / Leader tabs assume right align
                            radJustLeft.value = false;
                            radJustCenter.value = false;
                            radJustRight.value = true;
                            radJustFull.value = false;
                            radJustFullAll.value = false;
                            applyLeaderTab(obj);
                        }
                    }
                }
            }

            // サイズ調整モードのみ選択状態を復元（分離モードは削除済み参照のためスキップ）
            if (savedSel.length > 0) {
                try { app.activeDocument.selection = savedSel; } catch (e) { }
            }
            app.redraw();
        }

        // プレビュー更新：チェックON→適用、OFF→取り消し
        function updatePreview() {
            if (isPreviewActive) {
                app.undo();
                app.redraw();
                isPreviewActive = false;
            }
            if (chkPreview.value) {
                runProcess(true); // プレビューモード
                isPreviewActive = true;
            }
        }

        // マージンUIの有効/無効（「調整しない」のときだけ有効）
        function updateSeparationDim() {
            var isSeparating = radSeparate.value || radPathOpacity.value || radRemovePath.value;
            pnlJustify.enabled = !isSeparating;
            pnlTextAlign.enabled = !isSeparating;
            pnlAutoSize.enabled = !isSeparating;
            pnlFrameSize.enabled = !isSeparating;
            btnFrameAuto.enabled = pnlFrameSize.enabled;
            pnlIndent.enabled = !isSeparating;
            pnlOptions.enabled = !isSeparating;
            // 2段落以上のときは「文字あふれ」「フィット」を無効化（ディム）
            // Disable overset/fit buttons when multiple paragraphs exist
            if (pnlAutoSize.enabled) {
                btnTextSize.enabled = !hasMultiParagraph;
                btnFit.enabled = !hasMultiParagraph;
            } else {
                btnTextSize.enabled = false;
                btnFit.enabled = false;
            }
        }

        function updateMarginEnabled() {
            chkMargin.enabled = radNoAdjust.value;
            if (!radNoAdjust.value) {
                etMargin.enabled = false;
                lblUnit.enabled = false;
            } else {
                etMargin.enabled = chkMargin.value;
                lblUnit.enabled = chkMargin.value;
            }
        }

        // キー入力で行揃えラジオを選択 / Select justification radios by keys
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

                // Apply manual exclusive selection
                for (var i = 0; i < allJustRadios.length; i++) {
                    allJustRadios[i].value = (allJustRadios[i] === target);
                }

                // Leader tabs rule: if ON and not Right, turn OFF
                try {
                    if (chkLeader && chkLeader.value && target !== radJustRight) {
                        chkLeader.value = false;
                    }
                } catch (e) { }

                event.preventDefault();
                updatePreview();
            });
        }

        // ラジオボタン：排他制御 + プレビュー更新
        for (var r = 0; r < allRadios.length; r++) {
            (function (btn) {
                btn.onClick = function () {
                    for (var j = 0; j < allRadios.length; j++) {
                        allRadios[j].value = (allRadios[j] === btn);
                    }
                    updateSeparationDim();
                    updateMarginEnabled();
                    updatePreview();
                };
            })(allRadios[r]);
        }

        // 行揃えラジオボタン：手動排他制御（onClick 時点で .value が確定していないケースへの対応）
        for (var jr = 0; jr < allJustRadios.length; jr++) {
            (function (btn) {
                btn.onClick = function () {
                    for (var j = 0; j < allJustRadios.length; j++) {
                        allJustRadios[j].value = (allJustRadios[j] === btn);
                    }

                    // メニュー作成用（リーダー罫）がONのとき、「右」以外を選んだらOFFにする
                    // If leader tabs are ON, selecting non-right justification turns them OFF
                    try {
                        if (chkLeader && chkLeader.value && btn !== radJustRight) {
                            chkLeader.value = false;
                        }
                    } catch (e) { }

                    updatePreview();
                };
            })(allJustRadios[jr]);
        }
        addJustifyKeyHandler(dlg);

        btnTextSize.onClick = function () {
            _autoSizeMode = "textsize";
            runProcess(false);
            _autoSizeMode = "none";
        };
        btnFit.onClick = function () {
            _autoSizeMode = "fit";
            runProcess(false);
            _autoSizeMode = "none";
        };

        chkPreview.onClick = updatePreview;
        btnFrameAuto.onClick = function () {
            _frameAutoOn = !_frameAutoOn;
            updateFrameAutoButtonLabel();

            // Auto OFF: when turning OFF, clear auto-size-adjust on selected area text
            if (!_frameAutoOn) {
                try {
                    var t = app.activeDocument.selection;
                    for (var i = 0; i < t.length; i++) {
                        if (t[i] && t[i].typename === "TextFrame" && t[i].kind === TextType.AREATEXT) {
                            collapseFrameAuto(t[i]);
                        }
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
                chkLeftIndent.value = true;
                etLeftIndent.enabled = true;
                chkRightIndent.enabled = false;
                etRightIndent.enabled = false;
                etRightIndent.text = etLeftIndent.text; // 連動ON時に右を左へ即時同期
            } else {
                chkRightIndent.enabled = true;
                etRightIndent.enabled = chkRightIndent.value;
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

        // フォントサイズフィールド変更時：選択フレームに適用＋文字数再計算
        function applyFontSizeFromField() {
            var newSize = parseFloat(etFontSize.text) || 0;
            if (newSize <= 0) return;
            fontSize = newSize; // グローバルを更新（文字数連動に使用）
            var targets = app.activeDocument.selection;
            for (var i = 0; i < targets.length; i++) {
                var obj = targets[i];
                if (obj.typename === "TextFrame" && obj.kind === TextType.AREATEXT) {
                    try { obj.textRange.characterAttributes.size = newSize; } catch (e) { }
                }
            }
            // 文字数を再計算（幅固定のまま）
            if (fontSize > 0) {
                var effectivePt = (parseFloat(etWidth.text) || 0) * rulerInfo.toPt - getAdjustmentPt();
                etChars.text = Math.round((effectivePt / fontSize) * 100) / 100;
            }
            app.redraw();
        }

        // インデント＋間隔の合計（pt）：文字数計算の補正値
        function getAdjustmentPt() {
            var marginPt = chkMargin.value ? (parseFloat(etMargin.text) || 0) * rulerInfo.toPt : 0;
            var leftPt = (chkLeftIndent.value || chkSync.value) ? (parseFloat(etLeftIndent.text) || 0) * rulerInfo.toPt : 0;
            var rightPt = chkSync.value
                ? leftPt
                : (chkRightIndent.value ? (parseFloat(etRightIndent.text) || 0) * rulerInfo.toPt : 0);
            return 2 * marginPt + leftPt + rightPt;
        }

        // 幅→文字数、文字数→幅 の連動更新（インデント・間隔を加味）
        function onWidthChange() {
            if (fontSize > 0) {
                var effectivePt = (parseFloat(etWidth.text) || 0) * rulerInfo.toPt - getAdjustmentPt();
                etChars.text = Math.round((effectivePt / fontSize) * 100) / 100;
            }
            updatePreview();
        }
        function onCharsChange() {
            if (fontSize > 0) {
                var widthPt = (parseFloat(etChars.text) || 0) * fontSize + getAdjustmentPt();
                etWidth.text = Math.round((widthPt / rulerInfo.toPt) * 100) / 100;
            }
            updatePreview();
        }
        // インデント・間隔が変わったら幅を固定したまま文字数を再計算
        // 連動ONのとき、左の値を右フィールドにも反映してから計算
        function onAdjustmentChange() {
            if (chkSync.value) { etRightIndent.text = etLeftIndent.text; }
            onWidthChange();
        }

        // etMargin / etWidth / etHeight / indent / chars の値変更でプレビューを更新
        etFontSize.onChange = applyFontSizeFromField;
        etMargin.onChange = onAdjustmentChange;
        etWidth.onChange = onWidthChange;
        etHeight.onChange = updatePreview;
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

        btnOk.onClick = function () {
            saveDlgPosition(dlg);

            // If preview is active, rollback it first
            if (isPreviewActive) {
                try { app.undo(); } catch (e) { }
                try { app.redraw(); } catch (e2) { }
                isPreviewActive = false;
            }

            // Always run real execution
            runProcess(false);
            dlg.close(1);
        };

        btnCancel.onClick = function () {
            saveDlgPosition(dlg);
            if (isPreviewActive) { app.undo(); app.redraw(); }
            dlg.close(0);
        };

        // 初期値：最初のエリア内文字フレームの幅・高さ・インデント・文字サイズをルーラー単位で表示
        for (var si = 0; si < sel.length; si++) {
            if (sel[si].typename === "TextFrame" && sel[si].kind === TextType.AREATEXT) {
                var tf0 = sel[si];
                try { hasMultiParagraph = (tf0.paragraphs && tf0.paragraphs.length >= 2); } catch (e) { hasMultiParagraph = false; }
                var initW = tf0.textPath.width / rulerInfo.toPt;
                var initH = tf0.textPath.height / rulerInfo.toPt;
                try { fontSize = tf0.textRange.characterAttributes.size || 0; } catch (e) { }
                if (fontSize > 0) { etFontSize.text = Math.round(fontSize * 100) / 100; }
                etWidth.text = Math.round(initW * 100) / 100;
                etHeight.text = Math.round(initH * 100) / 100;

                // 行揃え初期値（最初の段落から取得してラジオボタンに反映）
                try {
                    var initJust = tf0.paragraphs.length > 0
                        ? tf0.paragraphs[0].justification : Justification.LEFT;
                    radJustLeft.value = (initJust === Justification.LEFT);
                    radJustCenter.value = (initJust === Justification.CENTER);
                    radJustRight.value = (initJust === Justification.RIGHT);
                    radJustFull.value = (initJust === Justification.FULLJUSTIFYLASTLINELEFT);
                    radJustFullAll.value = (initJust === Justification.FULLJUSTIFY);
                    // 上記に該当しない場合（FULLJUSTIFY等）はデフォルトの左に
                    if (!radJustLeft.value && !radJustCenter.value && !radJustRight.value &&
                        !radJustFull.value && !radJustFullAll.value) {
                        radJustLeft.value = true;
                    }
                } catch (e) { radJustLeft.value = true; }

                // 外側からの間隔初期値
                try {
                    var initSpacingPt = tf0.spacing || 0;
                    etMargin.text = Math.round((initSpacingPt / rulerInfo.toPt) * 100) / 100;
                    if (initSpacingPt !== 0) {
                        chkMargin.value = true;
                        etMargin.enabled = true;
                        lblUnit.enabled = true;
                    }
                } catch (e) { }

                // インデント初期値（最初の段落から取得）
                try {
                    var initLeftPt = tf0.paragraphs.length > 0 ? (tf0.paragraphs[0].leftIndent || 0) : 0;
                    var initRightPt = tf0.paragraphs.length > 0 ? (tf0.paragraphs[0].rightIndent || 0) : 0;

                    if (initLeftPt !== 0) {
                        etLeftIndent.text = Math.round((initLeftPt / rulerInfo.toPt) * 100) / 100;
                        chkLeftIndent.value = true;
                        etLeftIndent.enabled = true;
                    } else {
                        etLeftIndent.text = "0";
                        chkLeftIndent.value = false;
                        etLeftIndent.enabled = false;
                    }

                    if (initRightPt !== 0) {
                        etRightIndent.text = Math.round((initRightPt / rulerInfo.toPt) * 100) / 100;
                        chkRightIndent.value = true;
                        etRightIndent.enabled = true;
                    } else {
                        etRightIndent.text = "0";
                        chkRightIndent.value = false;
                        etRightIndent.enabled = false;
                    }
                } catch (e) { }

                // 文字数（インデント・間隔補正後）
                if (fontSize > 0) {
                    var effectivePt = tf0.textPath.width - getAdjustmentPt();
                    etChars.text = Math.round((effectivePt / fontSize) * 100) / 100;
                }
                // 2段落以上なら「文字あふれ」「フィット」をディム表示
                btnTextSize.enabled = !hasMultiParagraph;
                btnFit.enabled = !hasMultiParagraph;
                break;
            }
        }

        updateSeparationDim();
        restoreDlgPosition(dlg);
        updateFrameAutoButtonLabel();
        dlg.onClose = function () {
            saveDlgPosition(dlg);
            return true;
        };
        dlg.show();

    } else {
        alert(L("alertSelectAreaText"));
    }
} else {
    alert(L("alertNoDocument"));
}