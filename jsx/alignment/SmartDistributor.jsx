#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);
#targetengine "smartDistributorPalette"

/*

### 概要

DistributeDownFromTop.jsx / DistributeUpFromTop.jsx を統合した常駐パレットです。
十字ボタン（↑ / ← 0 → / ↓）を押すたびに、その時点の選択へ1ステップぶん適用します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartDistributor.md

### Overview

A persistent palette that combines DistributeDownFromTop.jsx and DistributeUpFromTop.jsx.
Each press of the cross buttons (↑ / ← 0 → / ↓) applies one step to whatever is selected at that moment.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartDistributor.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartDistributor";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartDistributor.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartDistributor.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 概要 / Overview
    // =========================================
    /*
    SmartDistributor.jsx

    DistributeDownFromTop / DistributeUpFromTop の統合パレット。
    十字ボタン（↑ / ← 0 → / ↓）を押すたびに、その時点の選択へ 1 ステップ適用する。
    実際のドキュメント操作は BridgeTalk でメインエンジンへ送って実行する（1 クリック = 取り消し 1 回）。

    複数オブジェクトを選択しているとき:
      縦（↑/↓）… 基準点で決めた端（既定=上端）を固定し、以降を「移動距離」ぶんずつ上下に等間隔調整
      横（←/→）… 基準点で決めた端（既定=左端）を固定し、以降を「移動距離」ぶんずつ左右に等間隔調整
      0       … オブジェクト間の隙間を 0 にする（基準点で固定端を決定。中央は並びの向きで自動判定）

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
    // ユーザー設定 / User Settings
    // =========================================

    /* 保存した設定が無いときの移動距離 / Distance source when nothing is saved
       "leading" = サイズ／行送り、"keyinput" = キー増加、"custom" = カスタム */
    var DEFAULT_DISTANCE_MODE = "leading";

    /* カスタムの移動距離の初期値（pt）。入力が数値でないときもこの値を使う
       Initial custom step in points; also used when the field is not a number */
    var DEFAULT_CUSTOM_STEP_PT = 0.1;

    /* 保存した設定が無いときの基準点 / Reference point when nothing is saved
       "top" / "bottom" / "left" / "right" / "center" */
    var DEFAULT_ANCHOR = "top";

    // =========================================
    // 保存ファイル / Saved files
    // =========================================

    /* パレット位置と設定の保存先 / Where the palette position and settings are stored */
    var POSITION_FILE = new File(Folder.userData + "/SmartDistributor/palette-position.txt");
    var SETTINGS_FILE = new File(Folder.userData + "/SmartDistributor/settings.txt");

    // =========================================
    // レイアウト / Layout
    // =========================================

    var PALETTE_MARGINS = 14;               /* パレットの余白 / palette margins */
    var PALETTE_SPACING = 10;               /* パレット内の間隔 / spacing inside the palette */
    var PALETTE_OPACITY = 0.97;             /* パレットの不透明度 / palette opacity */
    var COLUMN_SPACING = 12;                /* 基準点と十字ボタンの間隔 / gap between the two columns */
    var PANEL_MARGINS = [16, 20, 16, 12];   /* パネルの余白 / panel margins */
    var PANEL_SPACING = 8;                  /* パネル内の既定の間隔 / default spacing inside panels */
    var ANCHOR_PANEL_SPACING = 4;           /* 基準点の行間 / spacing between the reference point rows */
    var ANCHOR_RADIO_GAP = 18;              /* 基準点の中段ラジオの間隔 / gap between the middle-row radios */
    var NUDGE_BUTTON_SIZE = [32, 24];       /* 十字ボタンの大きさ / size of the nudge buttons */
    var NUDGE_BUTTON_GAP = 4;               /* 十字ボタンの間隔 / gap between the nudge buttons */
    var DISTANCE_PANEL_SPACING = 6;         /* 移動距離パネルの行間 / spacing inside the distance panel */
    var CUSTOM_ROW_SPACING = 6;             /* カスタム行の間隔 / spacing inside the custom row */
    var CUSTOM_FIELD_CHARS = 5;             /* カスタム値の入力欄の幅（文字数）/ width of the custom field */

    /**
     * パネルの共通設定
     * @param {Panel} targetPanel - 対象パネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    var uiLang = ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        dialog: {
            title: { ja: "スマート均等配置", en: "Smart Distributor" }
        },
        panel: {
            anchor: { ja: "基準点", en: "Reference point" },
            distance: { ja: "移動距離", en: "Distance" }
        },
        radio: {
            sourceTextLeading: { ja: "サイズ{slash}行送り（環境設定）", en: "Size{slash}Leading (Pref)" },
            sourceKeyInput: { ja: "キー増加（環境設定）", en: "Keyboard Increment (Pref)" },
            sourceCustom: { ja: "カスタム", en: "Custom" }
        },
        tooltip: {
            nudgeUp: {
                ja: "基準点で決めた端を固定し、上方向へ「移動距離」ぶん広げます（Shiftで10倍）。テキスト1つの選択では行送りを詰めます。",
                en: "Spreads the objects upward by the distance, keeping the anchored edge fixed (Shift for x10). With a single text object it tightens the leading."
            },
            nudgeDown: {
                ja: "基準点で決めた端を固定し、下方向へ「移動距離」ぶん広げます（Shiftで10倍）。テキスト1つの選択では行送りを広げます。",
                en: "Spreads the objects downward by the distance, keeping the anchored edge fixed (Shift for x10). With a single text object it opens up the leading."
            },
            nudgeLeft: {
                ja: "基準点で決めた端を固定し、左方向へ「移動距離」ぶん広げます（Shiftで10倍）。",
                en: "Spreads the objects to the left by the distance, keeping the anchored edge fixed (Shift for x10)."
            },
            nudgeRight: {
                ja: "基準点で決めた端を固定し、右方向へ「移動距離」ぶん広げます（Shiftで10倍）。",
                en: "Spreads the objects to the right by the distance, keeping the anchored edge fixed (Shift for x10)."
            },
            collapse: {
                ja: "オブジェクト間の隙間を 0 にして密着させます。基準点で固定する端を決めます。",
                en: "Closes the gaps between the objects. The reference point decides which edge stays put."
            },
            sourceTextLeading: {
                ja: "環境設定［テキスト］の「サイズ／行送り」の値を1回ぶんの移動距離に使います。",
                en: "Uses the Size/Leading increment from the Type preferences as one step."
            },
            sourceKeyInput: {
                ja: "環境設定［一般］の「キー入力」の値を1回ぶんの移動距離に使います。",
                en: "Uses the Keyboard Increment from the General preferences as one step."
            },
            sourceCustom: { ja: "1回ぶんの移動距離を pt で直接指定します。", en: "Sets one step directly, in points." },
            anchorTop: { ja: "上端を固定", en: "Fix top" },
            anchorBottom: { ja: "下端を固定", en: "Fix bottom" },
            anchorLeft: { ja: "左端を固定", en: "Fix left" },
            anchorRight: { ja: "右端を固定", en: "Fix right" },
            anchorCenter: { ja: "中央を固定", en: "Fix center" }
        }
    };

    /**
     * ドット区切りのパス（例 "panel.anchor"）で LABELS から現在の言語のラベルを取り出す
     * {slash} はスラッシュ "/" に置き換える（ソース内に裸の "/" を書かないため）
     * @param {string} labelPath - "panel.anchor" のようなドット区切りのキー
     * @returns {string} 現在の言語のラベル（見つからなければ labelPath そのもの）
     */
    function getLabel(labelPath) {
        var labelNode = LABELS;
        var pathKeys = labelPath.split(".");
        for (var i = 0; i < pathKeys.length; i++) {
            if (!labelNode || labelNode[pathKeys[i]] === undefined) {
                return labelPath;
            }
            labelNode = labelNode[pathKeys[i]];
        }
        var localizedText = (labelNode && labelNode[uiLang] !== undefined) ? labelNode[uiLang] : labelPath;
        return String(localizedText).replace(/\{slash\}/g, "/");
    }

    /**
     * 項目名と値の間に入れるコロンを返す（日本語は全角、英語は半角＋空白）
     * @returns {string} コロン
     */
    function labelColon() {
        return (uiLang === "ja") ? "：" : ": ";
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コードに対応する表示ラベルと、1単位あたりのポイント数
       Unit code -> display label and points per unit */
    var UNITS = [
        { label: "in",    pointsPerUnit: 72 },                /* 0 */
        { label: "mm",    pointsPerUnit: 72 / 25.4 },         /* 1 */
        { label: "pt",    pointsPerUnit: 1 },                 /* 2 */
        { label: "pica",  pointsPerUnit: 12 },                /* 3 */
        { label: "cm",    pointsPerUnit: 72 / 2.54 },         /* 4 */
        { label: "Q",     pointsPerUnit: 72 / 25.4 * 0.25 },  /* 5 */
        { label: "px",    pointsPerUnit: 1 },                 /* 6 */
        { label: "ft/in", pointsPerUnit: 72 * 12 },           /* 7 */
        { label: "m",     pointsPerUnit: 72 / 25.4 * 1000 },  /* 8 */
        { label: "yd",    pointsPerUnit: 72 * 36 },           /* 9 */
        { label: "ft",    pointsPerUnit: 72 * 12 }            /* 10 */
    ];

    /* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
       Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        /* 級（Q）と歯（H）は同じ長さだが、文字サイズは「Q」、距離は「H」と呼び分ける */
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    // =========================================
    // 矢印キーで値を増減 / Arrow-key value change
    // =========================================

    /**
     * 入力欄で ↑↓ キーによる値の増減を有効にする
     * ↑↓ = ±1 / Shift = ±10（10の倍数にスナップ）/ Option = ±0.1。0 未満にはしない
     * @param {EditText} editText - 対象の入力欄
     * @returns {void}
     */
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;

            var currentValue = Number(editText.text);
            if (isNaN(currentValue)) return;

            var keyboardState = ScriptUI.environment.keyboardState;
            var isUp = (event.keyName === "Up");
            var newValue;

            if (keyboardState.shiftKey) {
                newValue = isUp ? Math.ceil((currentValue + 1) / 10) * 10 : Math.floor((currentValue - 1) / 10) * 10;
            } else if (keyboardState.altKey) {
                newValue = Math.round((currentValue + (isUp ? 0.1 : -0.1)) * 10) / 10;
            } else {
                newValue = Math.round(currentValue) + (isUp ? 1 : -1);
            }

            if (newValue < 0) newValue = 0;
            editText.text = newValue;
            event.preventDefault();
        });
    }

    // =========================================
    // パレットの状態 / Palette state
    // =========================================

    /* パレットとコントロール（showPalette() の中で代入）/ Palette and controls, assigned in showPalette() */
    var paletteWindow = null;
    var anchorPanel = null;
    var anchorRadios = {};              /* 基準点キー → ラジオ / reference point key -> radio */
    var upButton = null;
    var downButton = null;
    var leftButton = null;
    var rightButton = null;
    var zeroButton = null;
    var sourceTextLeadingRadio = null;
    var sourceKeyInputRadio = null;
    var sourceCustomRadio = null;
    var customField = null;

    var currentAnchor = DEFAULT_ANCHOR; /* "top" / "bottom" / "left" / "right" / "center" */
    var workerInstalled = false;        /* BridgeTalk ワーカーを登録済みか / whether the worker is installed */

    // =========================================
    // パレットの構築 / Palette construction
    // =========================================

    /**
     * 前回のパレットが開いていれば閉じる（多重起動ガード）
     * @returns {void}
     */
    function closeExistingPalette() {
        /* 前回のウィンドウが破棄済みだと close() が失敗することがある / close() may fail on a disposed window */
        try {
            if ($.global.smartDistributorWindow) {
                $.global.smartDistributorWindow.close();
                $.global.smartDistributorWindow = null;
            }
        } catch (e) {
        }
    }

    /**
     * パレットのウィンドウを作る
     * @returns {Window} 作ったパレット
     */
    function createPaletteWindow() {
        var newPalette = new Window("palette", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        newPalette.orientation = "column";
        newPalette.alignChildren = ["fill", "top"];
        newPalette.margins = PALETTE_MARGINS;
        newPalette.spacing = PALETTE_SPACING;
        newPalette.opacity = PALETTE_OPACITY;
        return newPalette;
    }

    /**
     * 基準点のラジオを1つ追加し、anchorRadios に登録する
     * @param {Group} anchorRow - 追加先の行
     * @param {string} anchorKey - 基準点キー（"top" など）
     * @param {string} tooltipText - ツールチップ
     * @returns {void}
     */
    function addAnchorRadio(anchorRow, anchorKey, tooltipText) {
        var anchorRadio = anchorRow.add("radiobutton", undefined, "");
        anchorRadio.helpTip = tooltipText;
        anchorRadio.onClick = function () { selectAnchor(anchorKey); };
        anchorRadios[anchorKey] = anchorRadio;
    }

    /**
     * 基準点パネル（固定する位置を5点の十字から選ぶ）を追加する
     * @param {Group} parentGroup - 追加先のグループ
     * @param {string} initialAnchor - 最初に選ぶ基準点キー
     * @returns {void}
     */
    function addAnchorPanel(parentGroup, initialAnchor) {
        anchorPanel = parentGroup.add("panel", undefined, getLabel("panel.anchor"));
        setupPanel(anchorPanel, ANCHOR_PANEL_SPACING);
        anchorPanel.alignChildren = ["center", "center"];  /* 十字を中央寄せ / center the cross */
        anchorPanel.alignment = ["left", "fill"];

        var anchorTopRow = anchorPanel.add("group");
        addAnchorRadio(anchorTopRow, "top", getLabel("tooltip.anchorTop"));

        var anchorMidRow = anchorPanel.add("group");
        anchorMidRow.spacing = ANCHOR_RADIO_GAP;
        addAnchorRadio(anchorMidRow, "left", getLabel("tooltip.anchorLeft"));
        addAnchorRadio(anchorMidRow, "center", getLabel("tooltip.anchorCenter"));
        addAnchorRadio(anchorMidRow, "right", getLabel("tooltip.anchorRight"));

        var anchorBottomRow = anchorPanel.add("group");
        addAnchorRadio(anchorBottomRow, "bottom", getLabel("tooltip.anchorBottom"));

        applyAnchorSelection(initialAnchor);  /* 初期選択（保存はしない）/ initial selection, not saved */
    }

    /**
     * 5点のうち1つだけを選択状態にする（別コンテナのため自動排他が効かない）
     * @param {string} anchorKey - 選ぶ基準点キー
     * @returns {void}
     */
    function applyAnchorSelection(anchorKey) {
        for (var radioKey in anchorRadios) {
            anchorRadios[radioKey].value = (radioKey === anchorKey);
        }
        currentAnchor = anchorKey;
    }

    /**
     * 基準点を選び、設定を保存する
     * @param {string} anchorKey - 選ぶ基準点キー
     * @returns {void}
     */
    function selectAnchor(anchorKey) {
        applyAnchorSelection(anchorKey);
        saveSettings();
    }

    /**
     * 十字ボタンを1つ追加する
     * @param {Group} buttonRow - 追加先の行
     * @param {string} buttonText - ボタンの表示文字（矢印・0）
     * @param {string} tooltipText - ツールチップ
     * @param {function} clickHandler - クリック時の処理
     * @returns {Button} 追加したボタン
     */
    function addNudgeButton(buttonRow, buttonText, tooltipText, clickHandler) {
        var nudgeButton = buttonRow.add("button", undefined, buttonText);
        nudgeButton.helpTip = tooltipText;
        nudgeButton.preferredSize = NUDGE_BUTTON_SIZE;
        nudgeButton.onClick = clickHandler;
        return nudgeButton;
    }

    /**
     * 十字ボタン（↑ / ← 0 → / ↓）を追加する
     * dx / dy は座標差分（→ / ↓ を正）
     * @param {Group} parentGroup - 追加先のグループ
     * @returns {void}
     */
    function addNudgePad(parentGroup) {
        var nudgePadGroup = parentGroup.add("group");
        nudgePadGroup.orientation = "column";
        nudgePadGroup.alignChildren = ["center", "center"];
        nudgePadGroup.alignment = ["fill", "center"];
        nudgePadGroup.spacing = NUDGE_BUTTON_GAP;

        var nudgeTopRow = nudgePadGroup.add("group");
        upButton = addNudgeButton(nudgeTopRow, "↑", getLabel("tooltip.nudgeUp"), function () { applyNudge(0, -1); });

        var nudgeMidRow = nudgePadGroup.add("group");
        nudgeMidRow.spacing = NUDGE_BUTTON_GAP;
        leftButton = addNudgeButton(nudgeMidRow, "←", getLabel("tooltip.nudgeLeft"), function () { applyNudge(-1, 0); });
        zeroButton = addNudgeButton(nudgeMidRow, "0", getLabel("tooltip.collapse"), function () { collapseSpacing(); });
        rightButton = addNudgeButton(nudgeMidRow, "→", getLabel("tooltip.nudgeRight"), function () { applyNudge(1, 0); });

        var nudgeBottomRow = nudgePadGroup.add("group");
        downButton = addNudgeButton(nudgeBottomRow, "↓", getLabel("tooltip.nudgeDown"), function () { applyNudge(0, 1); });
    }

    /**
     * 移動距離パネル（ラジオ3択とカスタム値）を十字ボタンの下に追加する
     * @param {Window} parentWindow - 追加先のパレット
     * @param {string} initialMode - 最初に選ぶ移動距離（"leading" / "keyinput" / "custom"）
     * @param {string} initialCustom - カスタム値の初期表示
     * @returns {void}
     */
    function addDistancePanel(parentWindow, initialMode, initialCustom) {
        var distancePanel = parentWindow.add("panel", undefined, getLabel("panel.distance"));
        setupPanel(distancePanel, DISTANCE_PANEL_SPACING);
        distancePanel.alignChildren = ["left", "top"];  /* ラジオは左寄せ / left-align radios */

        sourceTextLeadingRadio = distancePanel.add("radiobutton", undefined, getLabel("radio.sourceTextLeading"));
        sourceTextLeadingRadio.helpTip = getLabel("tooltip.sourceTextLeading");
        sourceKeyInputRadio = distancePanel.add("radiobutton", undefined, getLabel("radio.sourceKeyInput"));
        sourceKeyInputRadio.helpTip = getLabel("tooltip.sourceKeyInput");
        refreshSourceLabels();  /* 環境設定の現在値をラベルに反映 / show the current preference values */

        var customRow = distancePanel.add("group");
        customRow.spacing = CUSTOM_ROW_SPACING;
        sourceCustomRadio = customRow.add("radiobutton", undefined, getLabel("radio.sourceCustom"));
        sourceCustomRadio.helpTip = getLabel("tooltip.sourceCustom");
        customField = customRow.add("edittext", undefined, initialCustom);
        customField.helpTip = getLabel("tooltip.sourceCustom");
        customField.characters = CUSTOM_FIELD_CHARS;
        changeValueByArrowKey(customField);
        customRow.add("statictext", undefined, "pt");

        /* 保存値からラジオの初期状態を復元 / Restore the radios from the saved mode */
        sourceTextLeadingRadio.value = (initialMode === "leading");
        sourceKeyInputRadio.value = (initialMode === "keyinput");
        sourceCustomRadio.value = (initialMode === "custom");
        if (!sourceTextLeadingRadio.value && !sourceKeyInputRadio.value && !sourceCustomRadio.value) {
            sourceTextLeadingRadio.value = true;
        }

        /* カスタムだけ別コンテナ（customRow）にあり ScriptUI の自動排他が効かないので手動で排他にする
           The custom radio lives in its own row, so exclusivity is handled by hand */
        sourceTextLeadingRadio.onClick = function () { selectDistanceMode(sourceTextLeadingRadio); };
        sourceKeyInputRadio.onClick = function () { selectDistanceMode(sourceKeyInputRadio); };
        sourceCustomRadio.onClick = function () { selectDistanceMode(sourceCustomRadio); };
        customField.onChange = function () { saveSettings(); };
        customField.enabled = sourceCustomRadio.value;
    }

    /**
     * 移動距離のラジオを1つだけ選び、カスタム値の入力欄を切り替えて設定を保存する
     * @param {RadioButton} selectedRadio - 選んだラジオ
     * @returns {void}
     */
    function selectDistanceMode(selectedRadio) {
        sourceTextLeadingRadio.value = (selectedRadio === sourceTextLeadingRadio);
        sourceKeyInputRadio.value = (selectedRadio === sourceKeyInputRadio);
        sourceCustomRadio.value = (selectedRadio === sourceCustomRadio);
        customField.enabled = sourceCustomRadio.value;  /* カスタム以外はディム / dim unless custom */
        saveSettings();
    }

    /**
     * 移動距離ラジオのラベルに環境設定の現在値を表示する（表示単位込み）
     * @returns {void}
     */
    function refreshSourceLabels() {
        /* 行送りは text/units 単位の値そのまま / the leading increment is already in text/units */
        var leadingValue = app.preferences.getRealPreference("text/sizeIncrement");
        sourceTextLeadingRadio.text = getLabel("radio.sourceTextLeading") + labelColon() + leadingValue + getUnitInfo("text/units").label;

        /* cursorKeyLength は pt で返るので一般単位（rulerType）へ換算して表示 / cursorKeyLength is in points */
        var rulerUnit = getUnitInfo("rulerType");
        var keyValue = app.preferences.getRealPreference("cursorKeyLength") / rulerUnit.pointsPerUnit;
        keyValue = Math.round(keyValue * 1000) / 1000;
        sourceKeyInputRadio.text = getLabel("radio.sourceKeyInput") + labelColon() + keyValue + rulerUnit.label;
    }

    /**
     * Option + 矢印キーで十字ボタンと同じ操作をする（カスタム値の編集中は入力欄に任せる）
     * @param {KeyboardEvent} keyEvent - キーイベント
     * @returns {void}
     */
    function onPaletteKeyDown(keyEvent) {
        if (!keyEvent.altKey) return;
        if (keyEvent.target === customField) return;
        if (keyEvent.keyName === "Up" && upButton.enabled) {
            applyNudge(0, -1);
            keyEvent.preventDefault();
        } else if (keyEvent.keyName === "Down" && downButton.enabled) {
            applyNudge(0, 1);
            keyEvent.preventDefault();
        } else if (keyEvent.keyName === "Left" && leftButton.enabled) {
            applyNudge(-1, 0);
            keyEvent.preventDefault();
        } else if (keyEvent.keyName === "Right" && rightButton.enabled) {
            applyNudge(1, 0);
            keyEvent.preventDefault();
        }
    }

    /**
     * パレットの移動・終了・アクティブ化・キー操作のイベントを登録する
     * @returns {void}
     */
    function attachPaletteEvents() {
        paletteWindow.onMove = function () {
            saveWindowPosition(paletteWindow);
        };

        paletteWindow.onClose = function () {
            saveWindowPosition(paletteWindow);
            saveSettings();
            $.global.smartDistributorWindow = null;
        };

        /* 選択が変わるたびにボタンの有効状態とラベルを更新 / Refresh buttons and labels on activation */
        paletteWindow.onActivate = function () {
            updateButtonStates();
            refreshSourceLabels();
        };

        paletteWindow.addEventListener("keydown", onPaletteKeyDown);
    }

    // =========================================
    // ボタン状態・呼び出し / Button state and actions
    // =========================================

    /**
     * 現在の選択を返す
     * @returns {Array|null} 選択（ドキュメントが無ければ null）
     */
    function getCurrentSelection() {
        if (app.documents.length < 1) return null;
        return app.activeDocument.selection;
    }

    /**
     * テキストを1つだけ選択しているかを返す
     * @param {Array|null} selectedObjects - 選択
     * @returns {boolean} テキスト1つなら true
     */
    function isSingleTextSelection(selectedObjects) {
        return selectedObjects && selectedObjects.length === 1 && selectedObjects[0].typename === "TextFrame";
    }

    /**
     * 選択内容に応じてボタンの有効・無効（ディム表示）を更新する
     * @returns {void}
     */
    function updateButtonStates() {
        var selectedObjects = getCurrentSelection();
        var hasSingleText = isSingleTextSelection(selectedObjects);
        var hasMultiple = selectedObjects && selectedObjects.length >= 2;

        /* 縦：テキスト1つ または 複数選択で有効 / vertical: single text or multiple objects */
        upButton.enabled = hasSingleText || hasMultiple;
        downButton.enabled = hasSingleText || hasMultiple;
        /* 横・0：複数選択のときだけ有効 / horizontal and 0: multiple objects only */
        leftButton.enabled = hasMultiple;
        rightButton.enabled = hasMultiple;
        zeroButton.enabled = hasMultiple;
        /* 基準点は複数選択時のみ／テキスト1つではキー増加も使えないのでディム
           Reference point needs multiple objects; the key increment is dimmed for a single text */
        anchorPanel.enabled = hasMultiple;
        sourceKeyInputRadio.enabled = !hasSingleText;
    }

    /**
     * 選択したラジオに応じた1回ぶんの移動距離を返す
     * @returns {number} 移動距離（pt）
     */
    function getStepPoints() {
        if (sourceKeyInputRadio.value) {
            return app.preferences.getRealPreference("cursorKeyLength");
        }
        if (sourceCustomRadio.value) {
            /* カスタムは pt 指定 / the custom value is in points */
            var customStepPt = parseFloat(customField.text);
            if (isNaN(customStepPt)) customStepPt = DEFAULT_CUSTOM_STEP_PT;
            return customStepPt;
        }
        /* 既定：環境設定のテキスト/行送り（表示単位込みで pt 換算）/ default: the Size/Leading increment in points */
        return app.preferences.getRealPreference("text/sizeIncrement") * getUnitInfo("text/units").pointsPerUnit;
    }

    /**
     * 矢印1回ぶんをワーカーへ送って適用する（Shift で 10 倍）
     * 選択の検証はワーカー側に任せる（collapse と同じ流れ）
     * @param {number} dx - 横方向（→ が 1、← が -1）
     * @param {number} dy - 縦方向（↓ が 1、↑ が -1）
     * @returns {void}
     */
    function applyNudge(dx, dy) {
        var multiplier = 1;
        /* keyboardState は環境によって例外になるため保護 / keyboardState may throw on some setups */
        try {
            if (ScriptUI.environment.keyboardState.shiftKey) multiplier = 10;
        } catch (e) {
        }
        callWorker("nudge", dx, dy, getStepPoints() * multiplier, currentAnchor);
    }

    /**
     * オブジェクト間の隙間を 0 にする
     * @returns {void}
     */
    function collapseSpacing() {
        callWorker("collapse", 0, 0, 0, currentAnchor);
    }

    // =========================================
    // BridgeTalk でメインエンジンへ送って実行 / Dispatch to the main engine via BridgeTalk
    // パレットエンジンからの直接編集は不安定なため、実処理はメインエンジン側で行う
    // =========================================

    var MAX_PENDING_JOBS = 12;  /* 保持する BridgeTalk ジョブの上限 / cap on pending BridgeTalk jobs */

    /**
     * ワーカーを（必要なら登録してから）短い呼び出しで実行する
     * @param {string} workerAction - "nudge" / "collapse"
     * @param {number} dx - 横方向
     * @param {number} dy - 縦方向
     * @param {number} step - 1回ぶんの移動距離（pt）
     * @param {string} anchor - 基準点キー
     * @returns {void}
     */
    function callWorker(workerAction, dx, dy, step, anchor) {
        if (typeof BridgeTalk === "undefined") return;
        if (!workerInstalled) installWorker();
        /* 不正な基準点は top にフォールバック / Fall back to top for invalid anchors */
        var safeAnchor = (anchor === "top" || anchor === "bottom" || anchor === "left"
            || anchor === "right" || anchor === "center") ? anchor : "top";
        sendWorkerCall('$.global.smartDistributorWorker("' + workerAction + '",' + dx + ',' + dy + ',' + step + ',"' + safeAnchor + '");', true);
    }

    /**
     * ワーカー本体をメインエンジンへ一度だけ登録する
     * @returns {void}
     */
    function installWorker() {
        var bridgeTalk = createIllustratorBridgeTalk();
        bridgeTalk.body = "$.global.smartDistributorWorker = (" + adjustSelectionWorker.toString() + ");";
        bridgeTalk.onResult = function () { removePendingJob(bridgeTalk); };
        bridgeTalk.onError = function () { removePendingJob(bridgeTalk); };
        $.global.smartDistributorJobs.push(bridgeTalk);
        trimPendingJobs();
        bridgeTalk.send();
        workerInstalled = true;
    }

    /**
     * 登録済みのワーカーを短い呼び出しで実行する（失敗時は1回だけ再登録して送り直す）
     * @param {string} callBody - メインエンジンで評価する呼び出し
     * @param {boolean} allowReinstall - 失敗時に再登録するか
     * @returns {void}
     */
    function sendWorkerCall(callBody, allowReinstall) {
        var bridgeTalk = createIllustratorBridgeTalk();
        bridgeTalk.body = callBody;
        bridgeTalk.onResult = function () { removePendingJob(bridgeTalk); };
        bridgeTalk.onError = function () {
            removePendingJob(bridgeTalk);
            if (allowReinstall) {
                workerInstalled = false;
                installWorker();
                sendWorkerCall(callBody, false);
            }
        };
        $.global.smartDistributorJobs.push(bridgeTalk);
        trimPendingJobs();
        bridgeTalk.send();
    }

    /**
     * Illustrator 宛ての BridgeTalk を作る
     * @returns {BridgeTalk} 宛先を設定した BridgeTalk
     */
    function createIllustratorBridgeTalk() {
        var bridgeTalk = new BridgeTalk();
        /* 指定子を引けない環境では素の名前で送る / fall back to the bare name when no specifier is found */
        try {
            bridgeTalk.target = BridgeTalk.getSpecifier("illustrator");
        } catch (e) {
            bridgeTalk.target = "illustrator";
        }
        return bridgeTalk;
    }

    /**
     * 保留中のジョブを MAX_PENDING_JOBS 件までに保つ
     * @returns {void}
     */
    function trimPendingJobs() {
        var pendingJobs = $.global.smartDistributorJobs;
        while (pendingJobs.length > MAX_PENDING_JOBS) pendingJobs.shift();
    }

    /**
     * 完了したジョブを保留リストから除く
     * @param {BridgeTalk} bridgeTalk - 完了したジョブ
     * @returns {void}
     */
    function removePendingJob(bridgeTalk) {
        var pendingJobs = $.global.smartDistributorJobs;
        for (var i = pendingJobs.length - 1; i >= 0; i--) {
            if (pendingJobs[i] === bridgeTalk) pendingJobs.splice(i, 1);
        }
    }

    // =========================================
    // メインエンジンで実行される本体（文字列化して登録）/ Worker run in the main engine (registered as a string)
    // ※ 外側の変数を参照せず、引数とアプリ DOM だけで完結させる
    // ※ toString() で送るため JSDoc を付けず、関数内のコメントは /* */ だけにする
    //    Sent through toString(): no JSDoc, and only block comments inside
    // =========================================

    /* 選択の間隔・行送りを1ステップ調整する / Adjust the spacing or leading of the selection by one step
       workerAction: "nudge" / "collapse"、anchor: "top" / "bottom" / "left" / "right" / "center"（固定する基準点） */
    function adjustSelectionWorker(workerAction, dx, dy, step, anchor) {
        var selectedItems;

        /* 配列へ写してから並べ替える / Copy into an array, then sort */
        function sortItemsBy(targetItems, compareItems) {
            var sortedItems = [];
            for (var i = 0; i < targetItems.length; i++) sortedItems.push(targetItems[i]);
            sortedItems.sort(compareItems);
            return sortedItems;
        }

        /* 並べ替えは見た目のボックス（geometricBounds）で判定する。TextFrame の position はベースライン（≒下端）を
           指すため、position で並べると上端側の固定要素が見た目の最上段とずれる（top / left の分配が崩れる原因） */
        /* 上→下（geometricBounds[1] = 上端 y）/ top to bottom */
        function sortTopToBottom(targetItems) {
            return sortItemsBy(targetItems, function (itemA, itemB) { return itemB.geometricBounds[1] - itemA.geometricBounds[1]; });
        }

        /* 左→右（geometricBounds[0] = 左端 x）/ left to right */
        function sortLeftToRight(targetItems) {
            return sortItemsBy(targetItems, function (itemA, itemB) { return itemA.geometricBounds[0] - itemB.geometricBounds[0]; });
        }

        /* 詰める向きを決める（中央は position の広がりが大きい軸）/ Pick the axis; center uses the wider spread */
        function resolveCollapseAxis() {
            if (anchor === "top" || anchor === "bottom") return "v";
            if (anchor === "left" || anchor === "right") return "h";
            var minX = null, maxX = null, minY = null, maxY = null;
            for (var i = 0; i < selectedItems.length; i++) {
                var itemPosition = selectedItems[i].position;
                if (minX === null || itemPosition[0] < minX) minX = itemPosition[0];
                if (maxX === null || itemPosition[0] > maxX) maxX = itemPosition[0];
                if (minY === null || itemPosition[1] < minY) minY = itemPosition[1];
                if (maxY === null || itemPosition[1] > maxY) maxY = itemPosition[1];
            }
            return ((maxX - minX) > (maxY - minY)) ? "h" : "v";
        }

        /* 縦の隙間を 0 に詰める / Close the vertical gaps */
        function collapseVertical() {
            var itemsTopToBottom = sortTopToBottom(selectedItems);
            var heights = [], totalHeight = 0, maxTop = null, minBottom = null;
            for (var i = 0; i < itemsTopToBottom.length; i++) {
                var bounds = itemsTopToBottom[i].geometricBounds;
                heights.push(bounds[1] - bounds[3]);
                totalHeight += bounds[1] - bounds[3];
                if (maxTop === null || bounds[1] > maxTop) maxTop = bounds[1];
                if (minBottom === null || bounds[3] < minBottom) minBottom = bounds[3];
            }
            /* 入れ子三項は必ず括弧で右結合を明示する（ExtendScript は括弧なしだと左結合に誤評価する） */
            var nextTop = (anchor === "bottom") ? (minBottom + totalHeight)
                : ((anchor === "top") ? maxTop : ((maxTop + minBottom) / 2 + totalHeight / 2));
            for (var j = 0; j < itemsTopToBottom.length; j++) {
                itemsTopToBottom[j].translate(0, nextTop - itemsTopToBottom[j].geometricBounds[1]);
                nextTop = nextTop - heights[j];
            }
        }

        /* 横の隙間を 0 に詰める / Close the horizontal gaps */
        function collapseHorizontal() {
            var itemsLeftToRight = sortLeftToRight(selectedItems);
            var widths = [], totalWidth = 0, minLeft = null, maxRight = null;
            for (var i = 0; i < itemsLeftToRight.length; i++) {
                var bounds = itemsLeftToRight[i].geometricBounds;
                widths.push(bounds[2] - bounds[0]);
                totalWidth += bounds[2] - bounds[0];
                if (minLeft === null || bounds[0] < minLeft) minLeft = bounds[0];
                if (maxRight === null || bounds[2] > maxRight) maxRight = bounds[2];
            }
            var nextLeft = (anchor === "right") ? (maxRight - totalWidth)
                : ((anchor === "left") ? minLeft : ((minLeft + maxRight) / 2 - totalWidth / 2));
            for (var j = 0; j < itemsLeftToRight.length; j++) {
                itemsLeftToRight[j].translate(nextLeft - itemsLeftToRight[j].geometricBounds[0], 0);
                nextLeft = nextLeft + widths[j];
            }
        }

        /* 隙間を 0 に詰める（基準点で固定端を決定）/ Close the gaps; the anchor decides the fixed edge */
        function collapseGaps() {
            if (selectedItems.length < 2) return "need 2+";
            if (resolveCollapseAxis() === "v") {
                collapseVertical();
            } else {
                collapseHorizontal();
            }
            return "ok";
        }

        /* テキスト1つの行送り増減（横は無効）/ Change the leading of a single text (vertical only) */
        function shiftTextLeading() {
            if (dy === 0) return "h disabled for text";
            var charAttributes = selectedItems[0].textRange.characterAttributes;
            charAttributes.leading = charAttributes.leading + dy * step;
            return "ok";
        }

        /* 1つの軸に沿って分配する（端固定＝矢印方向／中央＝対称）/ Distribute along one axis */
        function distributeAlongAxis(isVertical) {
            var sortedItems = isVertical ? sortTopToBottom(selectedItems) : sortLeftToRight(selectedItems);
            var startAnchor = isVertical ? "top" : "left";
            var endAnchor = isVertical ? "bottom" : "right";
            var isEdgeAnchor = (anchor === startAnchor || anchor === endAnchor);
            /* 入れ子三項は必ず括弧で右結合を明示する（括弧なしだと左結合に誤評価し、端が中央扱いになる） */
            var pivotIndex = (anchor === startAnchor) ? 0 : ((anchor === endAnchor) ? (sortedItems.length - 1) : (sortedItems.length - 1) / 2);
            for (var i = 0; i < sortedItems.length; i++) {
                var stepCount = i - pivotIndex;
                if (isEdgeAnchor) stepCount = Math.abs(stepCount);
                if (isVertical) {
                    sortedItems[i].translate(0, -stepCount * dy * step);
                } else {
                    sortedItems[i].translate(stepCount * dx * step, 0);
                }
            }
        }

        /* 例外はメインエンジンへ持ち込まず、文字列で返す / Report errors as strings instead of throwing */
        try {
            if (!app.documents || app.documents.length < 1) return "no document";
            selectedItems = app.activeDocument.selection;
            if (!selectedItems || selectedItems.length < 1) return "no selection";

            var result;
            if (workerAction === "collapse") {
                result = collapseGaps();
            } else if (selectedItems.length === 1 && selectedItems[0].typename === "TextFrame") {
                result = shiftTextLeading();
            } else if (selectedItems.length < 2) {
                return "need 2+";
            } else {
                if (dy !== 0) distributeAlongAxis(true);
                if (dx !== 0) distributeAlongAxis(false);
                result = "ok";
            }
            app.redraw();
            return result;
        } catch (e) {
            return "error: " + e.message;
        }
    }

    // =========================================
    // ウィンドウ位置の保存・復元 / Window position
    // =========================================

    /**
     * 保存したウィンドウ位置を復元する
     * @param {Window} targetWindow - 対象のウィンドウ
     * @returns {void}
     */
    function restoreWindowPosition(targetWindow) {
        if (!POSITION_FILE.exists) return;
        try {
            if (!POSITION_FILE.open("r")) return;
            var savedPositionText = POSITION_FILE.read();
            POSITION_FILE.close();

            var positionParts = savedPositionText.split(",");
            if (positionParts.length !== 2) return;

            var savedX = parseInt(positionParts[0], 10);
            var savedY = parseInt(positionParts[1], 10);
            if (isNaN(savedX) || isNaN(savedY)) return;

            targetWindow.location = [savedX, savedY];
        } catch (e) {
        }
    }

    /**
     * 現在のウィンドウ位置を保存する
     * @param {Window} targetWindow - 対象のウィンドウ
     * @returns {void}
     */
    function saveWindowPosition(targetWindow) {
        try {
            var positionFolder = POSITION_FILE.parent;
            if (!positionFolder.exists) positionFolder.create();
            if (!POSITION_FILE.open("w")) return;
            POSITION_FILE.write(targetWindow.location[0] + "," + targetWindow.location[1]);
            POSITION_FILE.close();
        } catch (e) {
        }
    }

    // =========================================
    // 設定の保存・復元（key=value 形式）/ Settings (key=value)
    // =========================================

    /**
     * 設定ファイルを読み込み、key=value を連想配列にする
     * @returns {object} 読み込んだ設定（無ければ空のオブジェクト）
     */
    function loadSettings() {
        var savedSettings = {};
        if (!SETTINGS_FILE.exists) return savedSettings;
        try {
            if (!SETTINGS_FILE.open("r")) return savedSettings;
            var settingsText = SETTINGS_FILE.read();
            SETTINGS_FILE.close();

            var settingLines = settingsText.split(/\r\n|\r|\n/);
            for (var i = 0; i < settingLines.length; i++) {
                var separatorIndex = settingLines[i].indexOf("=");
                if (separatorIndex < 1) continue;
                savedSettings[settingLines[i].substring(0, separatorIndex)] = settingLines[i].substring(separatorIndex + 1);
            }
        } catch (e) {
        }
        return savedSettings;
    }

    /**
     * 現在の設定を key=value 形式で保存する
     * @returns {void}
     */
    function saveSettings() {
        try {
            var settingsFolder = SETTINGS_FILE.parent;
            if (!settingsFolder.exists) settingsFolder.create();

            var distanceMode = sourceKeyInputRadio.value ? "keyinput"
                : (sourceCustomRadio.value ? "custom" : "leading");

            var settingLines = [
                "mode=" + distanceMode,
                "custom=" + customField.text,
                "anchor=" + currentAnchor
            ];

            if (!SETTINGS_FILE.open("w")) return;
            SETTINGS_FILE.write(settingLines.join("\n"));
            SETTINGS_FILE.close();
        } catch (e) {
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 保存した設定を読み込み、前回のパレットを閉じてから新しいパレットを表示する
     * @returns {void}
     */
    function showPalette() {
        var savedSettings = loadSettings();
        var initialMode = (savedSettings.mode !== undefined) ? savedSettings.mode : DEFAULT_DISTANCE_MODE;
        var initialCustom = (savedSettings.custom !== undefined) ? savedSettings.custom : String(DEFAULT_CUSTOM_STEP_PT);
        var initialAnchor = (savedSettings.anchor !== undefined) ? savedSettings.anchor : DEFAULT_ANCHOR;
        currentAnchor = initialAnchor;

        /* BridgeTalk ジョブの保持先 / Holder for pending BridgeTalk jobs */
        if (!$.global.smartDistributorJobs) {
            $.global.smartDistributorJobs = [];
        }

        closeExistingPalette();

        paletteWindow = createPaletteWindow();
        $.global.smartDistributorWindow = paletteWindow;

        /* 2カラム：左＝基準点 / 右＝十字ボタン / Two columns: reference point | nudge buttons */
        var mainRowGroup = paletteWindow.add("group");
        mainRowGroup.orientation = "row";
        mainRowGroup.alignChildren = ["fill", "center"];
        mainRowGroup.spacing = COLUMN_SPACING;

        addAnchorPanel(mainRowGroup, initialAnchor);
        addNudgePad(mainRowGroup);
        addDistancePanel(paletteWindow, initialMode, initialCustom);

        restoreWindowPosition(paletteWindow);
        attachPaletteEvents();

        updateButtonStates();
        paletteWindow.show();

        /* 起動直後にワーカーを登録（最初のクリックも短い呼び出しだけで済む）/ Install the worker at startup */
        if (typeof BridgeTalk !== "undefined") {
            installWorker();
        }
    }

    showPalette();

})();
