#target illustrator
#targetengine "smartDistributerPalette"

// 外部 JSX 実行時の警告ダイアログを抑制 / Suppress the external-JSX warning dialog
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

// =========================================
// 概要 / Overview
// =========================================
/*
SmartDistributer.jsx

DistributeDownFromTop / DistributeUpFromTop の統合パレット。
十字ボタン（↑ / ← 0 → / ↓）を押すたびに、その時点の選択へ 1 ステップ適用する。
実際のドキュメント操作は BridgeTalk でメインエンジンへ送って実行する（1 クリック = 取り消し 1 回）。

複数オブジェクトを選択しているとき:
  縦（↑/↓）… 最上部を固定し、以降を「移動距離」ぶんずつ上下に等間隔調整
  横（←/→）… 最左部を固定し、以降を「移動距離」ぶんずつ左右に等間隔調整
  0       … オブジェクト間の隙間を 0 にする（並びの向きで自動判定：縦長は上から、横長は左から密着）

テキストを 1 つだけ選択しているとき:
  縦（↑/↓）… 行送りを「移動距離」ぶん、↓ で加算／↑ で減算
  横（←/→）・0 … 無効（ディム表示）

Shift を押しながらボタンで 10 倍。
Option + 矢印キーでも十字ボタンと同じ操作ができる。

「基準点」パネル（5点の十字ラジオ）で固定する位置を指定する:
  ・上端/下端 … 縦方向の固定端（中段3点の縦成分は中央＝対称）
  ・左端/右端 … 横方向の固定端（上下点の横成分は中央）
  ・中央      … 縦横とも中央固定（対称分配・0 詰めは広がりの大きい軸）

移動距離は「移動距離」パネルのラジオで選ぶ:
  ・環境設定のテキスト/行送り（text/sizeIncrement × 表示単位 text/units）
  ・環境設定のキー増加（cursorKeyLength）
  ・カスタム（pt 指定）
いずれも内部では pt 換算して処理する。
*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartDistributer";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ローカライズ / Localization
// =========================================

var currentLanguage = ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";

var LABELS = {
    dialog: {
        title: { ja: "スマート均等配置", en: "Smart Distributer" }
    },
    panel: {
        anchor: { ja: "基準点", en: "Reference point" },
        distance: { ja: "移動距離", en: "Distance" }
    },
    radio: {
        sourceTextLeading: { ja: "テキスト{slash}行送り（環境設定）", en: "Type Size{slash}Leading (Pref)" },
        sourceKeyInput: { ja: "キー増加（環境設定）", en: "Keyboard Increment (Pref)" },
        sourceCustom: { ja: "カスタム", en: "Custom" }
    },
    tooltip: {
        anchorTop: { ja: "上端を固定", en: "Fix top" },
        anchorBottom: { ja: "下端を固定", en: "Fix bottom" },
        anchorLeft: { ja: "左端を固定", en: "Fix left" },
        anchorRight: { ja: "右端を固定", en: "Fix right" },
        anchorCenter: { ja: "中央を固定", en: "Fix center" }
    }
};

/* ドット区切りキー（例 "panel.anchor"）でカテゴリ分けされた LABELS を引く / Look up a categorized label by dot-separated key */
// 末尾の {slash} はスラッシュ "/" に置換する（ソース内に裸の "/" を書かないため）
function L(keyPath) {
    var node = LABELS;
    var parts = keyPath.split(".");
    for (var i = 0; i < parts.length; i++) {
        if (!node || node[parts[i]] === undefined) {
            return keyPath;
        }
        node = node[parts[i]];
    }
    var text = (node && node[currentLanguage] !== undefined) ? node[currentLanguage] : keyPath;
    return String(text).replace(/\{slash\}/g, "/");
}

// =========================================
// パネル・グループ共通レイアウト / Shared panel & group layout
// =========================================

/* パネルの余白と間隔 / Panel margins and spacing */
var PANEL_MARGINS = [16, 20, 16, 12];
var PANEL_SPACING = 8;

/* パネルの共通設定 / Apply shared panel layout */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = ["fill", "top"];
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

// =========================================
// 矢印キーで値を増減 / Arrow-key value change
// =========================================
// ↑↓ = ±1 / Shift = ±10（10の倍数にスナップ）/ Option = ±0.1
function changeValueByArrowKey(editText) {
    editText.addEventListener("keydown", function (event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;
        var delta = 1;

        if (keyboard.shiftKey) {
            delta = 10;
            // Shiftキー押下時は10の倍数にスナップ
            if (event.keyName === "Up") {
                value = Math.ceil((value + 1) / delta) * delta;
                event.preventDefault();
            } else if (event.keyName === "Down") {
                value = Math.floor((value - 1) / delta) * delta;
                if (value < 0) value = 0;
                event.preventDefault();
            }
        } else if (keyboard.altKey) {
            delta = 0.1;
            // Optionキー押下時は0.1単位で増減
            if (event.keyName === "Up") {
                value += delta;
                event.preventDefault();
            } else if (event.keyName === "Down") {
                value -= delta;
                event.preventDefault();
            }
        } else {
            delta = 1;
            if (event.keyName === "Up") {
                value += delta;
                event.preventDefault();
            } else if (event.keyName === "Down") {
                value -= delta;
                if (value < 0) value = 0;
                event.preventDefault();
            }
        }

        if (keyboard.altKey) {
            // 小数第1位までに丸め
            value = Math.round(value * 10) / 10;
        } else {
            // 整数に丸め
            value = Math.round(value);
        }

        editText.text = value;
    });
}

// =========================================
// メイン処理 / Main
// =========================================

(function () {
    var SCRIPT_NAME = L("dialog.title") + " " + SCRIPT_VERSION;
    var PREF_FILE = new File(Folder.userData + "/SmartDistributer/palette-position.txt");
    var SETTINGS_FILE = new File(Folder.userData + "/SmartDistributer/settings.txt");

    var textUnit = app.preferences.getIntegerPreference("text/units");
    var savedSettings = loadSettings();
    var initialMode = savedSettings.mode !== undefined ? savedSettings.mode : "leading";
    var initialCustom = savedSettings.custom !== undefined ? savedSettings.custom : "0.1";
    var initialAnchor = savedSettings.anchor !== undefined ? savedSettings.anchor : "top";
    var currentAnchor = initialAnchor;  // "top" / "bottom" / "left" / "right" / "center"

    // BridgeTalk ジョブ保持（送信完了前に GC されないよう保持） / Hold pending BridgeTalk jobs
    if (!$.global.smartDistributerJobs) {
        $.global.smartDistributerJobs = [];
    }

    // 多重起動ガード：既存パレットがあれば閉じてから作り直す
    try {
        if ($.global.smartDistributerWindow) {
            $.global.smartDistributerWindow.close();
            $.global.smartDistributerWindow = null;
        }
    } catch (e) {
    }

    // =========================================================
    // パレット
    // =========================================================

    var paletteWindow = new Window("palette", SCRIPT_NAME);
    $.global.smartDistributerWindow = paletteWindow;
    paletteWindow.orientation = "column";
    paletteWindow.alignChildren = ["fill", "top"];
    paletteWindow.margins = 14;
    paletteWindow.spacing = 10;
    paletteWindow.opacity = 0.97;  // パレットの不透明度 / Palette opacity

    // 2カラム：左＝基準点 / 右＝ボタン類
    var mainRow = paletteWindow.add("group");
    mainRow.orientation = "row";
    mainRow.alignChildren = ["fill", "center"];
    mainRow.spacing = 12;

    // 基準点パネル（左カラム・固定する位置を5点の十字から指定）
    var anchorPanel = mainRow.add("panel", undefined, L("panel.anchor"));
    setupPanel(anchorPanel, 4);
    anchorPanel.alignChildren = ["center", "center"];  // 十字を中央寄せ / center the cross
    anchorPanel.alignment = ["left", "fill"];

    var anchorTopRow = anchorPanel.add("group");
    var anchorTopRadio = anchorTopRow.add("radiobutton", undefined, "");
    anchorTopRadio.helpTip = L("tooltip.anchorTop");

    var anchorMidRow = anchorPanel.add("group");
    anchorMidRow.spacing = 18;
    var anchorLeftRadio = anchorMidRow.add("radiobutton", undefined, "");
    var anchorCenterRadio = anchorMidRow.add("radiobutton", undefined, "");
    var anchorRightRadio = anchorMidRow.add("radiobutton", undefined, "");
    anchorLeftRadio.helpTip = L("tooltip.anchorLeft");
    anchorCenterRadio.helpTip = L("tooltip.anchorCenter");
    anchorRightRadio.helpTip = L("tooltip.anchorRight");

    var anchorBottomRow = anchorPanel.add("group");
    var anchorBottomRadio = anchorBottomRow.add("radiobutton", undefined, "");
    anchorBottomRadio.helpTip = L("tooltip.anchorBottom");

    var anchorRadios = {
        top: anchorTopRadio,
        bottom: anchorBottomRadio,
        left: anchorLeftRadio,
        right: anchorRightRadio,
        center: anchorCenterRadio
    };

    anchorTopRadio.onClick = function () { selectAnchor("top"); };
    anchorBottomRadio.onClick = function () { selectAnchor("bottom"); };
    anchorLeftRadio.onClick = function () { selectAnchor("left"); };
    anchorRightRadio.onClick = function () { selectAnchor("right"); };
    anchorCenterRadio.onClick = function () { selectAnchor("center"); };

    applyAnchorSelection(initialAnchor);  // 初期選択（保存はしない）

    /* 5点を手動排他で選択（別コンテナのため自動排他が効かない） / Select one of the 5 points with manual exclusivity */
    function applyAnchorSelection(anchorKey) {
        for (var key in anchorRadios) {
            anchorRadios[key].value = (key === anchorKey);
        }
        currentAnchor = anchorKey;
    }

    /* 基準点を選び、設定を保存 / Select a reference point and persist it */
    function selectAnchor(anchorKey) {
        applyAnchorSelection(anchorKey);
        saveSettings();
    }

    // 十字ボタン（右カラム）   ↑ / ← 0 → / ↓
    var padGroup = mainRow.add("group");
    padGroup.orientation = "column";
    padGroup.alignChildren = ["center", "center"];
    padGroup.alignment = ["fill", "center"];
    padGroup.spacing = 4;

    var rowTop = padGroup.add("group");
    var upButton = rowTop.add("button", undefined, "↑");

    var rowMid = padGroup.add("group");
    rowMid.spacing = 4;
    var leftButton = rowMid.add("button", undefined, "←");
    var zeroButton = rowMid.add("button", undefined, "0");
    var rightButton = rowMid.add("button", undefined, "→");

    var rowBottom = padGroup.add("group");
    var downButton = rowBottom.add("button", undefined, "↓");

    var allButtons = [upButton, downButton, leftButton, rightButton, zeroButton];
    for (var b = 0; b < allButtons.length; b++) allButtons[b].preferredSize = [32, 24];

    // dx / dy は座標差分（→/↓ を正）
    upButton.onClick = function () { applyNudge(0, -1); };
    downButton.onClick = function () { applyNudge(0, 1); };
    leftButton.onClick = function () { applyNudge(-1, 0); };
    rightButton.onClick = function () { applyNudge(1, 0); };
    zeroButton.onClick = function () { collapseSpacing(); };

    // 移動距離パネル（ラジオ3択）— 十字ボタンの下に配置
    var distancePanel = paletteWindow.add("panel", undefined, L("panel.distance"));
    setupPanel(distancePanel, 6);
    distancePanel.alignChildren = ["left", "top"];  // ラジオは左寄せ / left-align radios

    var sourceTextLeadingRadio = distancePanel.add("radiobutton", undefined, L("radio.sourceTextLeading"));
    var sourceKeyInputRadio = distancePanel.add("radiobutton", undefined, L("radio.sourceKeyInput"));
    refreshSourceLabels();  // 環境設定の現在値をラベルに反映

    var customRow = distancePanel.add("group");
    customRow.spacing = 6;
    var sourceCustomRadio = customRow.add("radiobutton", undefined, L("radio.sourceCustom"));
    var customField = customRow.add("edittext", undefined, initialCustom);
    customField.characters = 5;
    changeValueByArrowKey(customField);
    customRow.add("statictext", undefined, "pt");

    // 保存値からラジオの初期状態を復元
    sourceTextLeadingRadio.value = (initialMode === "leading");
    sourceKeyInputRadio.value = (initialMode === "keyinput");
    sourceCustomRadio.value = (initialMode === "custom");
    if (!sourceTextLeadingRadio.value && !sourceKeyInputRadio.value && !sourceCustomRadio.value) {
        sourceTextLeadingRadio.value = true;
    }

    // カスタムだけ別コンテナ（customRow）にあり ScriptUI の自動排他が効かないので手動で排他にする
    sourceTextLeadingRadio.onClick = function () { selectDistanceMode(sourceTextLeadingRadio); };
    sourceKeyInputRadio.onClick = function () { selectDistanceMode(sourceKeyInputRadio); };
    sourceCustomRadio.onClick = function () { selectDistanceMode(sourceCustomRadio); };
    customField.onChange = function () { saveSettings(); };

    function selectDistanceMode(selectedRadio) {
        sourceTextLeadingRadio.value = (selectedRadio === sourceTextLeadingRadio);
        sourceKeyInputRadio.value = (selectedRadio === sourceKeyInputRadio);
        sourceCustomRadio.value = (selectedRadio === sourceCustomRadio);
        customField.enabled = sourceCustomRadio.value;  // カスタム以外はディム
        saveSettings();
    }
    customField.enabled = sourceCustomRadio.value;

    restoreWindowPosition(paletteWindow);

    paletteWindow.onMove = function () {
        saveWindowPosition(paletteWindow);
    };

    paletteWindow.onClose = function () {
        saveWindowPosition(paletteWindow);
        saveSettings();
        $.global.smartDistributerWindow = null;
    };

    // 選択が変わるたびにボタンの有効状態とラベルを更新
    paletteWindow.onActivate = function () {
        updateButtonStates();
        refreshSourceLabels();
    };

    // Option + 矢印キーで十字ボタンと同じ操作（カスタム値の編集中は field 側に任せる）
    try {
        paletteWindow.addEventListener("keydown", function (kbEvent) {
            if (!kbEvent.altKey) return;
            if (kbEvent.target === customField) return;
            if (kbEvent.keyName === "Up" && upButton.enabled) { applyNudge(0, -1); kbEvent.preventDefault(); }
            else if (kbEvent.keyName === "Down" && downButton.enabled) { applyNudge(0, 1); kbEvent.preventDefault(); }
            else if (kbEvent.keyName === "Left" && leftButton.enabled) { applyNudge(-1, 0); kbEvent.preventDefault(); }
            else if (kbEvent.keyName === "Right" && rightButton.enabled) { applyNudge(1, 0); kbEvent.preventDefault(); }
        });
    } catch (e) {
    }

    updateButtonStates();
    paletteWindow.show();

    // =========================================================
    // ボタン状態・呼び出し
    // =========================================================

    /* 現在の選択を取得（無ければ null） / Get the current selection (null if none) */
    function currentSelection() {
        if (app.documents.length < 1) return null;
        var doc = app.activeDocument;
        return doc.selection;
    }

    /* テキストを1つだけ選択しているか / Whether exactly one text frame is selected */
    function isSingleTextSelection(selection) {
        return selection && selection.length === 1 && selection[0].typename === "TextFrame";
    }

    /* 選択内容に応じてボタンの有効・無効（ディム表示）を更新 / Update button enabled state from the selection */
    function updateButtonStates() {
        var selection = currentSelection();
        var single = isSingleTextSelection(selection);
        var hasMultiple = selection && selection.length >= 2;

        // 縦：テキスト1つ または 複数選択で有効
        upButton.enabled = single || hasMultiple;
        downButton.enabled = single || hasMultiple;
        // 横・0：複数選択のときだけ有効
        leftButton.enabled = hasMultiple;
        rightButton.enabled = hasMultiple;
        zeroButton.enabled = hasMultiple;
        // 基準点は複数選択時のみ／テキスト1つではキー増加も使えないのでディム
        anchorPanel.enabled = hasMultiple;
        sourceKeyInputRadio.enabled = !single;
    }

    /* 選択したラジオに応じた移動距離（pt） / Move distance (pt) for the selected radio */
    function stepPoints() {
        if (sourceKeyInputRadio.value) {
            return app.preferences.getRealPreference("cursorKeyLength");
        }
        if (sourceCustomRadio.value) {
            var value = parseFloat(customField.text);
            if (isNaN(value)) value = 0.1;
            return value;  // カスタムは pt 指定
        }
        // 既定：環境設定のテキスト/行送り（表示単位込みで pt 換算）
        return app.preferences.getRealPreference("text/sizeIncrement") * unitInfo(textUnit).pt;
    }

    /* 矢印1回ぶんをメインエンジンへ送って適用（Shift で 10 倍） / Apply one step via the main engine (Shift = x10) */
    function applyNudge(dx, dy) {
        var multiplier = ScriptUI.environment.keyboardState.shiftKey ? 10 : 1;
        runInMainEngine("nudge", dx, dy, stepPoints() * multiplier, currentAnchor);
    }

    /* オブジェクト間の隙間を 0 にするをメインエンジンへ送る / Send the collapse-to-zero op to the main engine */
    function collapseSpacing() {
        runInMainEngine("collapse", 0, 0, 0, currentAnchor);
    }

    // =========================================================
    // BridgeTalk でメインエンジンへ処理を送る / Dispatch to the main engine via BridgeTalk
    // =========================================================

    /* worker 本体を自己完結で毎回送信（常駐させないので古い版が残らない） / Send the self-contained worker each time (no resident copy, never stale) */
    function runInMainEngine(action, dx, dy, step, anchor) {
        if (typeof BridgeTalk === "undefined") return;
        // 不正な基準点は top にフォールバック（中央への誤フォールバックを防ぐ） / Fall back to top for invalid anchors
        var safeAnchor = (anchor === "top" || anchor === "bottom" || anchor === "left"
            || anchor === "right" || anchor === "center") ? anchor : "top";
        var bridgeTalk = newIllustratorBridgeTalk();
        bridgeTalk.body = "(" + workerMain.toString() + ")(" +
            quoteForScript(action) + "," + dx + "," + dy + "," + step + "," + quoteForScript(safeAnchor) + ");";
        bridgeTalk.onResult = function () { removeJob(bridgeTalk); };
        bridgeTalk.onError = function () { removeJob(bridgeTalk); };
        $.global.smartDistributerJobs.push(bridgeTalk);
        trimJobs();
        bridgeTalk.send();
    }

    /* 文字列を安全に引用符で囲む / Quote a string safely for embedding */
    function quoteForScript(value) {
        return '"' + String(value).replace(/\\/g, "\\\\").replace(/"/g, '\\"') + '"';
    }

    /* Illustrator 宛ての BridgeTalk を生成 / Create a BridgeTalk addressed to Illustrator */
    function newIllustratorBridgeTalk() {
        var bridgeTalk = new BridgeTalk();
        try {
            bridgeTalk.target = BridgeTalk.getSpecifier("illustrator");
        } catch (e) {
            bridgeTalk.target = "illustrator";
        }
        return bridgeTalk;
    }

    /* 保留中のジョブを一定数に保つ / Keep pending jobs under a limit */
    function trimJobs() {
        var pendingJobs = $.global.smartDistributerJobs;
        while (pendingJobs.length > 12) {
            pendingJobs.shift();
        }
    }

    /* 完了したジョブを配列から取り除く / Remove a finished job from the array */
    function removeJob(bridgeTalk) {
        var pendingJobs = $.global.smartDistributerJobs;
        for (var i = pendingJobs.length - 1; i >= 0; i--) {
            if (pendingJobs[i] === bridgeTalk) {
                pendingJobs.splice(i, 1);
            }
        }
    }

    // =========================================================
    // メインエンジンで実行される本体（文字列化して送信） / Worker run in the main engine (sent as a string)
    // ※ 外側の変数を参照せず、引数とアプリ DOM だけで完結させる
    // =========================================================

    /* メインエンジンで調整を実行する本体 / Worker that performs the adjustment in the main engine
       anchor: "top" / "bottom" / "left" / "right" / "center"（固定する基準点）*/
    function workerMain(action, dx, dy, stepPoints, anchor) {
        try {
            if (!app.documents || app.documents.length < 1) return "ドキュメントがありません。";
            var doc = app.activeDocument;
            var selection = doc.selection;
            if (!selection || selection.length < 1) return "選択がありません。";

            function isSingleText(sel) {
                return sel.length === 1 && sel[0].typename === "TextFrame";
            }

            function copySelectionToArray(sel) {
                var items = [];
                for (var i = 0; i < sel.length; i++) {
                    items.push(sel[i]);
                }
                return items;
            }

            function sortByY(sel) {  // 上→下
                var arr = copySelectionToArray(sel);
                arr.sort(function (a, b) { return b.position[1] - a.position[1]; });
                return arr;
            }

            function sortByX(sel) {  // 左→右
                var arr = copySelectionToArray(sel);
                arr.sort(function (a, b) { return a.position[0] - b.position[0]; });
                return arr;
            }

            function getCollapseAxis(sel, anchorKey) {
                if (anchorKey === "top" || anchorKey === "bottom") return "v";
                if (anchorKey === "left" || anchorKey === "right") return "h";

                var minPX = null;
                var maxPX = null;
                var minPY = null;
                var maxPY = null;

                for (var i = 0; i < sel.length; i++) {
                    var pos = sel[i].position;
                    if (minPX === null || pos[0] < minPX) minPX = pos[0];
                    if (maxPX === null || pos[0] > maxPX) maxPX = pos[0];
                    if (minPY === null || pos[1] < minPY) minPY = pos[1];
                    if (maxPY === null || pos[1] > maxPY) maxPY = pos[1];
                }

                return ((maxPX - minPX) > (maxPY - minPY)) ? "h" : "v";
            }

            function collapseVertical(sel, anchorKey) {
                var itemsByY = sortByY(sel);
                var heights = [];
                var totalHeight = 0;
                var maxTop = null;
                var minBottom = null;

                for (var i = 0; i < itemsByY.length; i++) {
                    var bounds = itemsByY[i].geometricBounds;
                    var height = bounds[1] - bounds[3];
                    heights.push(height);
                    totalHeight += height;
                    if (maxTop === null || bounds[1] > maxTop) maxTop = bounds[1];
                    if (minBottom === null || bounds[3] < minBottom) minBottom = bounds[3];
                }

                var startTop = (anchorKey === "bottom") ? (minBottom + totalHeight)
                    : (anchorKey === "top") ? maxTop
                        : ((maxTop + minBottom) / 2 + totalHeight / 2);
                var currentTop = startTop;

                for (var j = 0; j < itemsByY.length; j++) {
                    itemsByY[j].translate(0, currentTop - itemsByY[j].geometricBounds[1]);
                    currentTop -= heights[j];
                }
            }

            function collapseHorizontal(sel, anchorKey) {
                var itemsByX = sortByX(sel);
                var widths = [];
                var totalWidth = 0;
                var minLeft = null;
                var maxRight = null;

                for (var i = 0; i < itemsByX.length; i++) {
                    var bounds = itemsByX[i].geometricBounds;
                    var width = bounds[2] - bounds[0];
                    widths.push(width);
                    totalWidth += width;
                    if (minLeft === null || bounds[0] < minLeft) minLeft = bounds[0];
                    if (maxRight === null || bounds[2] > maxRight) maxRight = bounds[2];
                }

                var startLeft = (anchorKey === "right") ? (maxRight - totalWidth)
                    : (anchorKey === "left") ? minLeft
                        : ((minLeft + maxRight) / 2 - totalWidth / 2);
                var currentLeft = startLeft;

                for (var j = 0; j < itemsByX.length; j++) {
                    itemsByX[j].translate(currentLeft - itemsByX[j].geometricBounds[0], 0);
                    currentLeft += widths[j];
                }
            }

            function collapseSpacingInWorker(sel, anchorKey) {
                if (sel.length < 2) return "複数選択が必要です。";

                if (getCollapseAxis(sel, anchorKey) === "v") {
                    collapseVertical(sel, anchorKey);
                } else {
                    collapseHorizontal(sel, anchorKey);
                }
                app.redraw();
                return "ok";
            }

            function adjustSingleTextLeading(textFrame, dyValue, stepValue) {
                if (dyValue === 0) return "テキストの横方向は無効です。";

                var attributes = textFrame.textRange.characterAttributes;
                attributes.leading = attributes.leading + dyValue * stepValue;
                app.redraw();
                return "ok";
            }

            function distributeVertically(sel, dyValue, stepValue, anchorKey) {
                var itemsByY = sortByY(sel);
                var fixedAtEnd = (anchorKey === "top" || anchorKey === "bottom");
                var anchorIndex = (anchorKey === "top") ? 0
                    : (anchorKey === "bottom") ? (itemsByY.length - 1)
                        : (itemsByY.length - 1) / 2;

                for (var i = 0; i < itemsByY.length; i++) {
                    var distanceIndex = i - anchorIndex;
                    if (fixedAtEnd) distanceIndex = Math.abs(distanceIndex);
                    itemsByY[i].translate(0, -distanceIndex * dyValue * stepValue);
                }
            }

            function distributeHorizontally(sel, dxValue, stepValue, anchorKey) {
                var itemsByX = sortByX(sel);
                var fixedAtEnd = (anchorKey === "left" || anchorKey === "right");
                var anchorIndex = (anchorKey === "left") ? 0
                    : (anchorKey === "right") ? (itemsByX.length - 1)
                        : (itemsByX.length - 1) / 2;

                for (var i = 0; i < itemsByX.length; i++) {
                    var distanceIndex = i - anchorIndex;
                    if (fixedAtEnd) distanceIndex = Math.abs(distanceIndex);
                    itemsByX[i].translate(distanceIndex * dxValue * stepValue, 0);
                }
            }

            function distributeSelection(sel, dxValue, dyValue, stepValue, anchorKey) {
                if (sel.length < 2) return "複数選択が必要です。";

                if (dyValue !== 0) {
                    distributeVertically(sel, dyValue, stepValue, anchorKey);
                }
                if (dxValue !== 0) {
                    distributeHorizontally(sel, dxValue, stepValue, anchorKey);
                }
                app.redraw();
                return "ok";
            }

            if (action === "collapse") {
                return collapseSpacingInWorker(selection, anchor);
            }

            if (isSingleText(selection)) {
                return adjustSingleTextLeading(selection[0], dy, stepPoints);
            }

            return distributeSelection(selection, dx, dy, stepPoints, anchor);
        } catch (e) {
            return "エラー: " + e.message;
        }
    }

    // =========================================================
    // 単位ヘルパー
    // =========================================================

    /* ラジオラベルに環境設定の現在値を反映（表示単位込み） / Reflect current preference values in the radio labels */
    function refreshSourceLabels() {
        textUnit = app.preferences.getIntegerPreference("text/units");
        // 行送りはテキスト単位（text/units）、キー増加は一般単位（rulerType）を参照
        var rulerType = app.preferences.getIntegerPreference("rulerType");

        // 行送り：値は text/units 単位そのまま
        var leadingValue = app.preferences.getRealPreference("text/sizeIncrement");
        sourceTextLeadingRadio.text = L("radio.sourceTextLeading") + "：" + leadingValue + unitInfo(textUnit).label;

        // キー増加：cursorKeyLength は pt で返るので一般単位へ換算して表示
        var keyValue = app.preferences.getRealPreference("cursorKeyLength") / unitInfo(rulerType).pt;
        keyValue = Math.round(keyValue * 1000) / 1000;
        sourceKeyInputRadio.text = L("radio.sourceKeyInput") + "：" + keyValue + unitInfo(rulerType).label;
    }

    /* 単位コード（text/units・ruler 共通）→ ラベルと pt 換算係数 / Unit code to label and points-per-unit
       0=inch, 1=mm, 2=pt, 3=pica, 4=cm, 5=Q, 6=px */
    function unitInfo(unitType) {
        if (unitType === 0) return { label: "inch", pt: 72 };
        if (unitType === 1) return { label: "mm", pt: 72 / 25.4 };
        if (unitType === 3) return { label: "pica", pt: 12 };
        if (unitType === 4) return { label: "cm", pt: 72 / 2.54 };
        if (unitType === 5) return { label: "Q", pt: 72 / 25.4 * 0.25 };  // 1Q = 0.25mm
        if (unitType === 6) return { label: "px", pt: 1 };
        return { label: "pt", pt: 1 };                                    // pt / 既定
    }

    // =========================================================
    // ウィンドウ位置の保存・復元
    // =========================================================

    /* 保存したウィンドウ位置を復元 / Restore the saved window position */
    function restoreWindowPosition(windowRef) {
        if (!PREF_FILE.exists) return;
        try {
            if (!PREF_FILE.open("r")) return;
            var savedPositionText = PREF_FILE.read();
            PREF_FILE.close();

            var positionParts = savedPositionText.split(",");
            if (positionParts.length !== 2) return;

            var x = parseInt(positionParts[0], 10);
            var y = parseInt(positionParts[1], 10);
            if (isNaN(x) || isNaN(y)) return;

            windowRef.location = [x, y];
        } catch (e) {
        }
    }

    /* 現在のウィンドウ位置を保存 / Save the current window position */
    function saveWindowPosition(windowRef) {
        try {
            var prefFolder = PREF_FILE.parent;
            if (!prefFolder.exists) prefFolder.create();
            if (!PREF_FILE.open("w")) return;
            PREF_FILE.write(windowRef.location[0] + "," + windowRef.location[1]);
            PREF_FILE.close();
        } catch (e) {
        }
    }

    // =========================================================
    // 設定の保存・復元（key=value 形式）
    // =========================================================

    /* 設定ファイルを読み込み key=value を連想配列に / Load settings into a key=value map */
    function loadSettings() {
        var result = {};
        if (!SETTINGS_FILE.exists) return result;
        try {
            if (!SETTINGS_FILE.open("r")) return result;
            var content = SETTINGS_FILE.read();
            SETTINGS_FILE.close();

            var lines = content.split(/\r\n|\r|\n/);
            for (var i = 0; i < lines.length; i++) {
                var separatorIndex = lines[i].indexOf("=");
                if (separatorIndex < 1) continue;
                var key = lines[i].substring(0, separatorIndex);
                var value = lines[i].substring(separatorIndex + 1);
                result[key] = value;
            }
        } catch (e) {
        }
        return result;
    }

    /* 現在の設定を key=value 形式で保存 / Save current settings as key=value */
    function saveSettings() {
        try {
            var settingsFolder = SETTINGS_FILE.parent;
            if (!settingsFolder.exists) settingsFolder.create();

            var mode = sourceKeyInputRadio.value ? "keyinput"
                : (sourceCustomRadio.value ? "custom" : "leading");

            var lines = [
                "mode=" + mode,
                "custom=" + customField.text,
                "anchor=" + currentAnchor
            ];

            if (!SETTINGS_FILE.open("w")) return;
            SETTINGS_FILE.write(lines.join("\n"));
            SETTINGS_FILE.close();
        } catch (e) {
        }
    }

})();
