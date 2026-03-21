#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
    作成日：2025-03-01
    更新日：2026-03-04
    概要 / Summary:
    - マージ：テキストフレームを2つ以上選択して、内容を統合した新しいエリア内文字を作成。
      Merge: Select 2+ text frames and combine contents into a new area text frame.
    - スレッドテキスト：
      - 2つ以上選択：リンク（必要に応じて既存スレッド解除→再リンク）
      - 1つ選択：リンク解除 / スレッドから除外（置き換え）
      Thread Text:
      - 2+ selected: Link (optionally unthread first, then relink)
      - Single selection: Unlink / Release from thread (replacement)
    - 交換：テキストフレーム2つを選択して、文字列のみを入れ替え。
      Swap: Select exactly 2 text frames and exchange text only.
    - 順序：マージ時の結合順を、上→下 / 左→右 で切り替え。
      Order: Controls merge order (top-to-bottom or left-to-right).
    - スタイル：均等配置、外枠からの間隔（エリア内文字のみ）、アピアランス等を適用。
      「書式保持（段落単位）」は段落ごとに文字属性を復元（段落設定は完全ではありません）。
      Style: Apply justification, inset spacing (area text only), appearance, etc.
      “Preserve Formatting (per paragraph)” restores character attributes per paragraph (paragraph settings are not fully preserved).
    - フレームの高さ：マージ結果のエリア内文字に対して、フィット/自動サイズ調整を適用。
      Frame Height: Apply Fit/Auto Size to the merged area text frame.
*/

var SCRIPT_VERSION = "v1.0";

/* ロケール判定 / Detect locale */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: { ja: "複数のテキストフレーム", en: "Multiple Text Frames" },
    panelOption: { ja: "オプション", en: "Option" },
    rbMerge: { ja: "マージ", en: "Merge" },
    rbThread: { ja: "スレッドテキスト", en: "Thread Text" },
    rbSwap: { ja: "交換（文字列のみ）", en: "Swap (Text Only)" },
    panelOrder: { ja: "順序", en: "Order" },
    rbTopToBottom: { ja: "上から", en: "Top to Bottom" },
    rbLeftToRight: { ja: "左から", en: "Left to Right" },
    panelStyle: { ja: "スタイル", en: "Style" },
    cbRemoveCR: { ja: "改行削除", en: "Remove Line Breaks" },
    cbPreserve: { ja: "書式保持（段落単位）", en: "Preserve Formatting (per paragraph)" },
    cbSpacing: { ja: "外枠からの間隔", en: "Inset Spacing" },
    cbJustify: { ja: "均等配置（最終行左揃え）", en: "Justify (Last Line Left)" },
    cbAppearance: { ja: "アピアランス", en: "Appearance" },
    panelHeight: { ja: "フレームの高さ", en: "Frame Height" },
    rbNone: { ja: "何もしない", en: "None" },
    rbFit: { ja: "フィット", en: "Fit" },
    rbAuto: { ja: "自動サイズ調整", en: "Auto Size" },
    btnOK: { ja: "OK", en: "OK" },
    btnCancel: { ja: "キャンセル", en: "Cancel" },
    alertNoSelection: { ja: "テキストフレームを選択してください。", en: "Please select a text frame." },
    alertNeedTwo: { ja: "テキストフレームを2つ以上選択してください。", en: "Please select 2 or more text frames." },
    panelThreadText: { ja: "スレッドテキスト", en: "Thread Text" },
    rbThreadLink: { ja: "リンク", en: "Link" },
    rbThreadUnlink: { ja: "リンクを解除", en: "Unlink" },
    rbThreadRelease: { ja: "スレッドから除外", en: "Release from Thread" },
    rbThreadAdd: { ja: "スレッドに追加", en: "Add to Thread" },
    rbThreadRelease2: { ja: "スレッドから除外（2）", en: "Release from Thread (2)" },
    alertNoDoc: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
    alertNoText: { ja: "選択にテキストフレームが含まれていません。", en: "Selection does not contain any text frames." },
    alertNotThreaded: { ja: "選択されたテキストフレームはスレッドテキストではありません。", en: "The selected text frame is not threaded text." }
};

function L(key) {
    return LABELS[key][lang];
}

/*
    LE - ライブエフェクトヘルパー / Helper object for Live Effect scripts
*/
var LE = {
    functionName: 'LE',
    testMode: false,
    debug: false,
    defaults: {},
    defaultsObject: function (item, defaults, options, func) {
        LE.functionName = func.name;
        try {
            if (defaults == undefined && options == undefined) return {};
            if (defaults == undefined) return options;
            if (options == undefined) return defaults;
            if (options.debug) LE.debug = true;
            for (var key in options) {
                defaults[key] = options[key];
            }
            LE.defaults = defaults;
            return defaults;
        } catch (error) {
            throw new Error(func.name + ' failed to parse options object. ' + error)
        }
    },
    applyEffect: function (item, xml, expand) {
        if (LE.testMode) {
            LE.testResults.push({ timestamp: new Date(), functionName: LE.functionName, xml: xml });
            return xml;
        } else {
            if (item == undefined) {
                throw new Error(LE.functionName + ' failed. No item available.');
            }
            var items;
            if (item.typename != undefined) {
                items = [item];
            } else if (item.length != undefined) {
                items = item;
            }
            if (!items) throw new Error(LE.functionName + ' failed. Unexpected item type. [1]');
            for (var i = 0; i < items.length; i++) {
                if (items[i].typename == undefined) throw new Error(LE.functionName + ' failed. Unexpected item type. [2]');
                items[i].applyEffect(xml);
                if (expand) LE.expandAppearance(items[i]);
            }
            if (LE.debug) $.writeln(LE.functionName + ':\n' + xml);
        }
    },
    handleError: function (error) {
        alert(error.message);
    },
    expandAppearance: function (item) {
        app.redraw();
        app.activeDocument.selection = [item];
        app.executeMenuCommand('expandStyle');
        item = app.activeDocument.selection[0];
    },
    formatColor: function (colr) {
        if (colr == undefined) throw new Error(LE.functionName + ': No color available');
        var colorCode, breakdown;
        if (typeof colr == 'string') {
            var isPercent = (colr.search('%') != -1);
            colr = colr.match(/-?\d+\.?\d*%?/g);
            if (colr == undefined) throw new Error(String(LE.functionName) + ': Couldn\'t parse color.');
        }
        if (Object.prototype.toString.call(colr) === '[object Array]') {
            for (var i = 0; i < colr.length; i++) colr[i] = parseFloat(colr[i]);
            switch (colr.length) {
                case 1:
                    colorCode = 0;
                    breakdown = [Number(colr[0]) / 100];
                    break;
                case 3:
                    colorCode = 5;
                    var divisor = isPercent ? 100 : 255;
                    breakdown = [Number(colr[0]) / divisor, Number(colr[1]) / divisor, Number(colr[2]) / divisor];
                    break;
                case 4:
                    colorCode = 1;
                    breakdown = [colr[0] / 100, colr[1] / 100, colr[2] / 100, colr[3] / 100];
                    break;
                default:
                    throw new Error(String(LE.functionName) + ': Couldn\'t parse color (' + colr + ')');
            }
            for (var i = 0; i < breakdown.length; i++) breakdown[i] = Math.round(Number(breakdown[i]) * 1000) / 1000;
        } else if (colr.typename != undefined) {
            switch (colr.typename) {
                case 'GrayColor':
                    colorCode = 0;
                    breakdown = [colr.gray / 100];
                    break;
                case 'RGBColor':
                    colorCode = 5;
                    breakdown = [colr.red / 255, colr.green / 255, colr.blue / 255];
                    break;
                case 'CMYKColor':
                    colorCode = 1;
                    breakdown = [colr.cyan / 100, colr.magenta / 100, colr.yellow / 100, colr.black / 100];
                    break;
                default:
                    throw new Error(String(LE.functionName) + ': Couldn\'t parse color of type \'' + colr.typename + '\'');
            }
        }
        return [colorCode].concat(breakdown).join(' ');
    },
    transformPoints: [
        Transformation.TOPLEFT, Transformation.TOP, Transformation.TOPRIGHT,
        Transformation.LEFT, Transformation.CENTER, Transformation.RIGHT,
        Transformation.BOTTOMLEFT, Transformation.BOTTOM, Transformation.BOTTOMRIGHT
    ]
};

function LE_ConvertToShape(item, options) {
    try {
        var defaults = {
            shapeType: 0,
            widthPts: 10,
            heightPts: 10,
            absoluteness: false,
            cornerRadiusPts: 9,
            expandAppearance: false
        }
        var o = LE.defaultsObject(item, defaults, options, arguments.callee)
        o.shapeName = ['Rectangle', 'RoundedRectangle', 'Ellipse'][o.shapeType];
        if (o.absoluteness == true || o.absoluteness > 0) {
            o.widthPts = o.widthPts || 100;
            o.heightPts = o.heightPts || 100;
        }
        var xml = '<LiveEffect name="Adobe Shape Effects" isPre="1"><Dict data="U DisplayString #1 I Shape #2 R RelWidth #3 R RelHeight #4 R AbsWidth #5 R AbsHeight #6 R Absolute #7 R CornerRadius #8 "/></LiveEffect>'
            .replace(/#1/, o.shapeName)
            .replace(/#2/, o.shapeType)
            .replace(/#3/, o.widthPts)
            .replace(/#4/, o.heightPts)
            .replace(/#5/, o.widthPts)
            .replace(/#6/, o.heightPts)
            .replace(/#7/, Number(o.absoluteness))
            .replace(/#8/, o.cornerRadiusPts);
        LE.applyEffect(item, xml, o.expandAppearance);
    } catch (error) {
        LE.handleError(error);
    }
}

function LE_ConvertToShape_Rectangle(item, options) {
    if (options == undefined || typeof options != 'object') options = {};
    options.shapeType = 0;
    LE_ConvertToShape(item, options);
}

/* 自動サイズ調整のON/OFF / Toggle auto size adjustment via Illustrator action */
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
    var actionLoaded = false;
    try {
        f.open('w');
        f.write(str);
        f.close();
        app.loadAction(f);
        actionLoaded = true;
        app.doScript("AutoSize", "AreaType", false);
    } finally {
        try { f.remove(); } catch (e) { }
        if (actionLoaded) {
            try { app.unloadAction("AreaType", ""); } catch (e) { }
        }
    }
}

/* 単位ユーティリティ / Unit utilities */
var unitMap = {
    0: "in", 1: "mm", 2: "pt", 3: "pica", 4: "cm",
    6: "px", 7: "ft/in", 8: "m", 9: "yd", 10: "ft"
};

function getUnitLabel(code, prefKey) {
    if (code === 5) {
        var hKeys = { "text/asianunits": true, "rulerType": true, "strokeUnits": true };
        return hKeys[prefKey] ? "H" : "Q";
    }
    return unitMap[code] || "pt";
}

function getPtFactorFromUnitCode(code) {
    switch (code) {
        case 0: return 72.0;                   // in
        case 1: return 72.0 / 25.4;            // mm
        case 2: return 1.0;                    // pt
        case 3: return 12.0;                   // pica
        case 4: return 72.0 / 2.54;            // cm
        case 5: return 72.0 / 25.4 * 0.25;     // Q or H
        case 6: return 1.0;                    // px
        case 7: return 72.0 * 12.0;            // ft/in
        case 8: return 72.0 / 25.4 * 1000.0;   // m
        case 9: return 72.0 * 36.0;            // yd
        case 10: return 72.0 * 12.0;           // ft
        default: return 1.0;
    }
}

(function () {
    if (app.documents.length === 0) {
        alert(L('alertNoDoc'));
        return;
    }
    var doc = app.activeDocument;
    var selection = doc.selection;

    if (!selection || selection.length < 1) {
        alert(L('alertNoSelection'));
        return;
    }

    /* テキストフレームが含まれているか検証 / Verify selection contains text frames */
    var hasTextFrame = false;
    for (var i = 0; i < selection.length; i++) {
        if (selection[i].typename === "TextFrame") {
            hasTextFrame = true;
            break;
        }
    }
    if (!hasTextFrame) {
        alert(L('alertNoText'));
        return;
    }

    /* 1つ選択時：スレッドテキストでなければ終了 / Single selection: exit if not threaded */
    if (selection.length === 1 && selection[0].typename === "TextFrame") {
        var isThreaded = false;
        try { if (selection[0].nextFrame) isThreaded = true; } catch (e) { }
        try { if (selection[0].previousFrame) isThreaded = true; } catch (e) { }
        if (!isThreaded) {
            alert(L('alertNotThreaded'));
            return;
        }
    }

    /* ルーラー単位 / Ruler units */
    var rulerCode = app.preferences.getIntegerPreference("rulerType");
    var rulerLabel = getUnitLabel(rulerCode, "rulerType");
    var rulerToPoint = getPtFactorFromUnitCode(rulerCode);

    /* ダイアログボックス / Dialog box */
    var dlg = new Window("dialog", L('dialogTitle') + ' ' + SCRIPT_VERSION);
    var mainGroup = dlg.add("group");
    mainGroup.orientation = "row";
    mainGroup.alignChildren = ["fill", "top"];

    /* 左カラム：オプション、順序 / Left column: Option, Order */
    var colLeft = mainGroup.add("group");
    colLeft.orientation = "column";
    colLeft.alignChildren = ["fill", "top"];

    var panel = colLeft.add("panel", undefined, L('panelOption'));
    panel.orientation = "column";
    panel.alignment = ["fill", "top"];
    panel.alignChildren = ["left", "center"];
    panel.margins = [15, 20, 15, 10];
    var rbMerge = panel.add("radiobutton", undefined, L('rbMerge'));
    var rbThread = panel.add("radiobutton", undefined, L('rbThread'));
    var rbSwap = panel.add("radiobutton", undefined, L('rbSwap'));
    var isSingle = (selection.length === 1);
    if (isSingle) {
        rbMerge.value = false;
        rbThread.value = true;
    } else {
        rbMerge.value = true;
    }

    var panelOrder = colLeft.add("panel", undefined, L('panelOrder'));
    panelOrder.orientation = "column";
    panelOrder.alignment = ["fill", "top"];
    panelOrder.alignChildren = ["left", "center"];
    panelOrder.margins = [15, 20, 15, 10];
    var rbTopToBottom = panelOrder.add("radiobutton", undefined, L('rbTopToBottom'));
    var rbLeftToRight = panelOrder.add("radiobutton", undefined, L('rbLeftToRight'));
    rbTopToBottom.value = true;

    var panelThread = colLeft.add("panel", undefined, L('panelThreadText'));
    panelThread.orientation = "column";
    panelThread.alignment = ["fill", "top"];
    panelThread.alignChildren = ["left", "center"];
    panelThread.margins = [15, 20, 15, 10];
    var rbThreadLink = panelThread.add("radiobutton", undefined, L('rbThreadLink'));
    var rbThreadUnlink = panelThread.add("radiobutton", undefined, L('rbThreadUnlink'));
    var rbThreadAdd = panelThread.add("radiobutton", undefined, L('rbThreadAdd'));
    var rbThreadRelease = panelThread.add("radiobutton", undefined, L('rbThreadRelease'));
    var rbThreadRelease2 = panelThread.add("radiobutton", undefined, L('rbThreadRelease2'));
    if (isSingle) {
        rbThreadRelease.value = true;
    } else {
        rbThreadLink.value = true;
    }

    /* 右カラム：スタイル、フレームの高さ / Right column: Style, Frame Height */
    var colRight = mainGroup.add("group");
    colRight.orientation = "column";
    colRight.alignChildren = ["fill", "top"];

    var panel2 = colRight.add("panel", undefined, L('panelStyle'));
    panel2.orientation = "column";
    panel2.alignment = ["fill", "top"];
    panel2.alignChildren = ["left", "center"];
    panel2.margins = [15, 20, 15, 10];
    var cbRemoveCR = panel2.add("checkbox", undefined, L('cbRemoveCR'));
    cbRemoveCR.value = false;
    var cbPreserve = panel2.add("checkbox", undefined, L('cbPreserve'));
    cbPreserve.helpTip = "段落ごとに文字属性を復元します（段落設定は完全ではありません）";
    cbPreserve.value = true;
    cbRemoveCR.enabled = !cbPreserve.value;
    cbPreserve.onClick = function () { cbRemoveCR.enabled = !cbPreserve.value; };
    var grpSpacing = panel2.add("group");
    var cbSpacing = grpSpacing.add("checkbox", undefined, L('cbSpacing'));
    cbSpacing.value = false;
    var txtSpacing = grpSpacing.add("edittext", undefined, "1");
    txtSpacing.characters = 4;
    txtSpacing.enabled = cbSpacing.value;
    grpSpacing.add("statictext", undefined, rulerLabel);
    cbSpacing.onClick = function () { txtSpacing.enabled = cbSpacing.value; };
    var cbJustify = panel2.add("checkbox", undefined, L('cbJustify'));
    cbJustify.value = true;
    var cbAppearance = panel2.add("checkbox", undefined, L('cbAppearance'));
    cbAppearance.value = false;

    var panel3 = colRight.add("panel", undefined, L('panelHeight'));
    panel3.orientation = "column";
    panel3.alignment = ["fill", "top"];
    panel3.alignChildren = ["left", "center"];
    panel3.margins = [15, 20, 15, 10];
    var rbNone = panel3.add("radiobutton", undefined, L('rbNone'));
    var rbFit = panel3.add("radiobutton", undefined, L('rbFit'));
    var rbAuto = panel3.add("radiobutton", undefined, L('rbAuto'));
    rbFit.value = true;

    /* ポイント文字/パステキストが含まれているか / Check for non-area text */
    var hasNonAreaText = false;
    for (var j = 0; j < selection.length; j++) {
        if (selection[j].typename === "TextFrame" &&
            (selection[j].kind === TextType.POINTTEXT || selection[j].kind === TextType.PATHTEXT)) {
            hasNonAreaText = true;
            break;
        }
    }

    /* パネルの有効/無効を切り替え / Toggle panel enabled state */
    var updatePanels = function () {
        /* 1つのときマージ・交換を無効化 / Disable merge and swap when single object */
        rbMerge.enabled = !isSingle;
        rbSwap.enabled = (selection.length === 2);
        /* ポイント文字/パステキスト含有時スレッドを無効化 / Disable thread when non-area text present */
        rbThread.enabled = !hasNonAreaText;
        /* 交換・スレッドのとき順序を無効化 / Disable order when swap or thread is selected */
        var orderEnabled = !rbSwap.value && !rbThread.value;
        rbTopToBottom.enabled = orderEnabled;
        rbLeftToRight.enabled = orderEnabled;
        /* スレッドテキストパネルはスレッド選択時のみ有効 / Thread text panel enabled only when thread is selected */
        var threadPanelEnabled = rbThread.value;
        rbThreadLink.enabled = threadPanelEnabled;
        rbThreadUnlink.enabled = threadPanelEnabled;
        rbThreadRelease.enabled = threadPanelEnabled;
        rbThreadAdd.enabled = threadPanelEnabled;
        rbThreadRelease2.enabled = threadPanelEnabled;
        /* 交換・1つ選択のときスタイルを無効化 / Disable style when swap is selected or single object */
        var styleEnabled = !rbSwap.value && !isSingle;
        cbRemoveCR.enabled = styleEnabled && !cbPreserve.value;
        cbPreserve.enabled = styleEnabled;
        cbSpacing.enabled = styleEnabled;
        txtSpacing.enabled = styleEnabled && cbSpacing.value;
        cbJustify.enabled = styleEnabled;
        cbAppearance.enabled = styleEnabled;
        /* 交換・スレッド・1つ選択のときフレームの高さを無効化 / Disable frame height when swap, thread, or single object */
        var heightEnabled = !rbThread.value && !rbSwap.value && !isSingle;
        rbNone.enabled = heightEnabled;
        rbFit.enabled = heightEnabled;
        rbAuto.enabled = heightEnabled;
    };
    updatePanels();
    rbMerge.onClick = updatePanels;
    rbThread.onClick = updatePanels;
    rbSwap.onClick = updatePanels;

    var btnGroup = dlg.add("group");
    btnGroup.alignment = ["right", "center"];
    btnGroup.add("button", undefined, L('btnCancel'), { name: "cancel" });
    btnGroup.add("button", undefined, L('btnOK'), { name: "ok" });

    if (dlg.show() !== 1) return;

    /* フレームの高さを適用（エリア内文字のみ） / Apply frame height (area text only) */
    function applyFrameHeight(frames) {
        if (rbNone.value) return;
        for (var i = 0; i < frames.length; i++) {
            if (frames[i].typename !== "TextFrame" || frames[i].kind !== TextType.AREATEXT) continue;
            if (rbFit.value) {
                expandFrameToFit(frames[i]);
                collapseFrameAuto(frames[i]);
            } else if (rbAuto.value) {
                expandFrameToFit(frames[i]);
            }
        }
    }

    /* スタイルを適用（各フレームに対して実行） / Apply style to each frame */
    function applyStyle(frames) {
        for (var i = 0; i < frames.length; i++) {
            if (frames[i].typename !== "TextFrame") continue;
            try {
                if (cbJustify.value) {
                    frames[i].textRange.paragraphAttributes.justification = Justification.FULLJUSTIFYLASTLINELEFT;
                }
                if (cbSpacing.value && frames[i].kind === TextType.AREATEXT) {
                    var spacingVal = parseFloat(txtSpacing.text);
                    if (!isNaN(spacingVal)) {
                        frames[i].spacing = spacingVal * rulerToPoint;
                    }
                }
            } catch (e) { /* フレームが無効な状態の場合スキップ */ }
        }
    }

    /* スレッドテキスト / Thread text */
    if (rbThread.value) {
        if (rbThreadUnlink.value) {
            app.executeMenuCommand('removeThreading');
        } else if (rbThreadRelease.value) {
            /*
                選択フレームのテキスト・書式・座標を保存してからスレッド除外
                NOTE:
                - 先に contents を消すと、コマンド失敗時にテキスト消失のリスクがある。
                - まず release を実行し、成功を確認してから contents を消す。
            */

            function isThreadedTF(tf) {
                if (!tf || tf.typename !== "TextFrame") return false;
                var threaded = false;
                try { if (tf.nextFrame) threaded = true; } catch (e) { }
                try { if (tf.previousFrame) threaded = true; } catch (e) { }
                return threaded;
            }

            var relFrames = [];
            var targets = [];

            // Preserve original selection to restore later (best-effort)
            var __origSelection = [];
            try {
                for (var osi = 0; osi < doc.selection.length; osi++) __origSelection.push(doc.selection[osi]);
            } catch (e) { }

            for (var ri = 0; ri < selection.length; ri++) {
                if (selection[ri].typename !== "TextFrame") continue;
                var rf = selection[ri];

                // Save geometry + text + per-paragraph character attributes (best-effort)
                var rb = rf.geometricBounds;
                var rw = rb[2] - rb[0];
                var rh = rb[1] - rb[3];
                var rText = rf.contents;

                var rParaAttrs = [];
                for (var rpi = 0; rpi < rf.paragraphs.length; rpi++) {
                    try {
                        var rca = rf.paragraphs[rpi].characterAttributes;
                        rParaAttrs.push({
                            textFont: rca.textFont,
                            size: rca.size,
                            leading: rca.leading,
                            fillColor: rca.fillColor,
                            tracking: rca.tracking,
                            kerningMethod: rca.kerningMethod
                        });
                    } catch (e) {
                        break;
                    }
                }

                relFrames.push({
                    ref: rf,
                    left: rb[0],
                    top: rb[1],
                    width: rw,
                    height: rh,
                    text: rText,
                    paraAttrs: rParaAttrs
                });

                targets.push(rf);
            }

            // Execute release FIRST (no destructive changes yet)
            var releaseOk = false;
            try {
                // Ensure selection is exactly the targets
                doc.selection = targets;
                app.executeMenuCommand('releaseThreadedTextSelection');

                // Success criteria: all targets are no longer threaded
                releaseOk = true;
                for (var ck = 0; ck < relFrames.length; ck++) {
                    if (isThreadedTF(relFrames[ck].ref)) {
                        releaseOk = false;
                        break;
                    }
                }
            } catch (e) {
                releaseOk = false;
            }

            if (!releaseOk) {
                alert("スレッド解除に失敗したため、内容は変更していません。\n（releaseThreadedTextSelection が実行できませんでした）");
                return;
            }

            // Now it is safe to clear the original contents (no longer impacts the thread chain)
            for (var clr = 0; clr < relFrames.length; clr++) {
                try {
                    relFrames[clr].ref.contents = "";
                } catch (e) { }
            }

            /* 保存した情報から新しいエリア内文字を作成 / Create new area text from saved info */
            for (var rj = 0; rj < relFrames.length; rj++) {
                var info = relFrames[rj];
                var newRect = doc.pathItems.rectangle(info.top, info.left, info.width, info.height);
                var newTF = doc.textFrames.areaText(newRect);
                newTF.contents = info.text;

                /* まず先頭段落の書式をtextRange全体に適用 / Apply first para attrs to entire textRange */
                if (info.paraAttrs.length > 0) {
                    var base = info.paraAttrs[0];
                    var tr = newTF.textRange.characterAttributes;
                    tr.textFont = base.textFont;
                    tr.size = base.size;
                    tr.leading = base.leading;
                    tr.fillColor = base.fillColor;
                    tr.tracking = base.tracking;
                    tr.kerningMethod = base.kerningMethod;
                }

                /* 段落ごとに書式を適用 / Apply attributes per paragraph */
                for (var rpj = 1; rpj < info.paraAttrs.length; rpj++) {
                    try {
                        var pa = info.paraAttrs[rpj];
                        var pc = newTF.paragraphs[rpj].characterAttributes;
                        pc.textFont = pa.textFont;
                        pc.size = pa.size;
                        pc.leading = pa.leading;
                        pc.fillColor = pa.fillColor;
                        pc.tracking = pa.tracking;
                        pc.kerningMethod = pa.kerningMethod;
                    } catch (e) {
                        break;
                    }
                }
            }

            // Remove now-empty original frames (confirmed non-threaded)
            for (var rm = 0; rm < relFrames.length; rm++) {
                try {
                    var rref = relFrames[rm].ref;
                    if (!rref || rref.typename !== "TextFrame") continue;
                    // Safety: remove only if empty and no longer threaded
                    var isEmpty = false;
                    try { isEmpty = (rref.contents === ""); } catch (e) { isEmpty = false; }
                    if (!isEmpty) continue;
                    if (isThreadedTF(rref)) continue;
                    rref.remove();
                } catch (e) { }
            }

            // Restore selection (best-effort). Skip items that were removed.
            try {
                if (__origSelection && __origSelection.length > 0) {
                    var __sel2 = [];
                    for (var si2 = 0; si2 < __origSelection.length; si2++) {
                        try {
                            var it2 = __origSelection[si2];
                            // Accessing typename on a removed pageItem can throw
                            if (it2 && it2.typename) __sel2.push(it2);
                        } catch (e) { }
                    }
                    if (__sel2.length > 0) doc.selection = __sel2;
                }
            } catch (e) { }
        } else if (rbThreadRelease2.value) {
            /* 選択オブジェクトを複製→独立確認→元のテキストを削除→元オブジェクトを削除 */
            for (var r2i = 0; r2i < selection.length; r2i++) {
                if (selection[r2i].typename !== "TextFrame") continue;
                var origFrame = selection[r2i];
                /* 複製（スレッドに属さない独立コピーが作られる） */
                var dupFrame = origFrame.duplicate();
                /* 複製がスレッドから独立しているか簡易検証 */
                var dupOk = true;
                try { if (dupFrame.nextFrame) dupOk = false; } catch (e) { }
                try { if (dupFrame.previousFrame) dupOk = false; } catch (e) { }
                if (!dupOk) {
                    dupFrame.remove();
                    continue;
                }
                /* 元フレームのテキストをスレッドから削除 */
                origFrame.contents = "";
                /* 元フレームを削除 */
                origFrame.remove();
            }
        } else if (rbThreadAdd.value) {
            app.executeMenuCommand('removeThreading');
            app.executeMenuCommand('threadTextCreate');
        } else {
            /* スレッドテキストが混在しているか判定 / Check if selection has threaded frames */
            var hasThreaded = false;
            for (var ti = 0; ti < selection.length; ti++) {
                if (selection[ti].typename !== "TextFrame") continue;
                try { if (selection[ti].nextFrame) { hasThreaded = true; break; } } catch (e) { }
                try { if (selection[ti].previousFrame) { hasThreaded = true; break; } } catch (e) { }
            }
            if (hasThreaded) {
                app.executeMenuCommand('removeThreading');
            }
            app.executeMenuCommand('threadTextCreate');
        }
        return;
    }

    /* 交換 / Swap */
    if (rbSwap.value) {
        if (selection.length !== 2) {
            alert(L('alertNeedTwo'));
            return;
        }
        var temp = selection[0].contents;
        selection[0].contents = selection[1].contents;
        selection[1].contents = temp;
        return;
    }

    /* マージ / Merge */
    var textFrames = [];
    for (var i = 0; i < selection.length; i++) {
        if (selection[i].typename === "TextFrame" &&
            (selection[i].kind === TextType.AREATEXT || selection[i].kind === TextType.POINTTEXT || selection[i].kind === TextType.PATHTEXT)) {
            textFrames.push(selection[i]);
        }
    }

    if (textFrames.length < 2) {
        alert(L('alertNeedTwo'));
        return;
    }

    /* ソート（左から右、または上から下） / Sort (left-to-right or top-to-bottom) */
    if (rbLeftToRight.value) {
        textFrames.sort(function (a, b) {
            return a.geometricBounds[0] - b.geometricBounds[0];
        });
    } else {
        textFrames.sort(function (a, b) {
            return b.geometricBounds[1] - a.geometricBounds[1];
        });
    }

    var minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
    var combinedText = [];
    var firstTextFrame = textFrames[0];
    var textAttributes = firstTextFrame.textRange.characterAttributes;
    var CR = String.fromCharCode(13);

    for (var i = 0; i < textFrames.length; i++) {
        var tf = textFrames[i];
        var bounds = tf.geometricBounds;

        minX = Math.min(minX, bounds[0]);
        maxY = Math.max(maxY, bounds[1]);
        maxX = Math.max(maxX, bounds[2]);
        minY = Math.min(minY, bounds[3]);

        combinedText.push(tf.contents);
    }

    /* 書式保持用：フレームごとの属性と段落数を収集 / Collect frame-level attributes and paragraph counts */
    var frameAttrs = [];
    var frameParagraphCounts = [];
    if (cbPreserve.value) {
        for (var ci = 0; ci < textFrames.length; ci++) {
            var ca = textFrames[ci].textRange.characterAttributes;
            frameAttrs.push({
                textFont: ca.textFont,
                size: ca.size,
                leading: ca.leading,
                fillColor: ca.fillColor,
                tracking: ca.tracking,
                kerningMethod: ca.kerningMethod
            });
            frameParagraphCounts.push(textFrames[ci].paragraphs.length);
        }
    }

    var rectWidth = maxX - minX;
    var rectHeight = maxY - minY;

    var rect = doc.pathItems.rectangle(maxY, minX, rectWidth, rectHeight);
    var newTextFrame = doc.textFrames.areaText(rect);
    var joined = combinedText.join(CR);
    if (cbRemoveCR.value) {
        joined = joined.replace(/[\r\n]/g, '');
    }
    newTextFrame.contents = joined;

    /* 書式の適用 / Apply character attributes */
    if (cbPreserve.value && frameAttrs.length > 0 && !cbRemoveCR.value) {
        /* 書式保持：先頭フレームの書式をtextRange全体に適用後、段落ごとに上書き */
        var baseAttr = frameAttrs[0];
        var baseRange = newTextFrame.textRange.characterAttributes;
        baseRange.textFont = baseAttr.textFont;
        baseRange.size = baseAttr.size;
        baseRange.leading = baseAttr.leading;
        baseRange.fillColor = baseAttr.fillColor;
        baseRange.tracking = baseAttr.tracking;
        baseRange.kerningMethod = baseAttr.kerningMethod;
        var paraIdx = 0;
        for (var fi = 0; fi < frameAttrs.length; fi++) {
            for (var pi = 0; pi < frameParagraphCounts[fi]; pi++) {
                try {
                    var pr = newTextFrame.paragraphs[paraIdx].characterAttributes;
                    pr.textFont = frameAttrs[fi].textFont;
                    pr.size = frameAttrs[fi].size;
                    pr.leading = frameAttrs[fi].leading;
                    pr.fillColor = frameAttrs[fi].fillColor;
                    pr.tracking = frameAttrs[fi].tracking;
                    pr.kerningMethod = frameAttrs[fi].kerningMethod;
                } catch (e) { /* 末尾の空段落はスキップ */ }
                paraIdx++;
            }
        }
    } else {
        /* 先頭フレームの書式を一括適用 / Apply first frame attributes uniformly */
        var newTextRange = newTextFrame.textRange;
        newTextRange.characterAttributes.textFont = textAttributes.textFont;
        newTextRange.characterAttributes.size = textAttributes.size;
        newTextRange.characterAttributes.leading = textAttributes.leading;
        newTextRange.characterAttributes.fillColor = textAttributes.fillColor;
        newTextRange.characterAttributes.tracking = textAttributes.tracking;
        newTextRange.characterAttributes.kerningMethod = textAttributes.kerningMethod;
    }

    for (var i = textFrames.length - 1; i >= 0; i--) {
        textFrames[i].remove();
    }

    /* スタイルを適用 / Apply style */
    applyStyle([newTextFrame]);

    /* 新しいエリア内文字を選択 / Select new area text frame */
    newTextFrame.selected = true;

    /* フレームの高さを適用 / Apply frame height */
    applyFrameHeight([newTextFrame]);

    /* アピアランスを適用 / Apply appearance */
    if (cbAppearance.value) {
        // app.executeMenuCommand('Adobe New Fill Shortcut');
        app.executeMenuCommand('Adobe New Stroke Shortcut');
        LE_ConvertToShape_Rectangle(newTextFrame);
        // app.executeMenuCommand("Live Pathfinder Add");
    }
    app.redraw();
})();