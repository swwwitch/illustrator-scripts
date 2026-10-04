#target illustrator
#targetengine "SwwwitchPalettes"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ビデオ定規・すべてのドキュメントでの定規・ガイド・スマートガイド・アートボード・境界線・カンバスカラー・バウンディングボックスの表示を、常駐パレットのボタンで切り替えます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ViewTogglePalette.md

### Overview

A persistent palette whose buttons toggle the video ruler, rulers in all documents, guides, smart guides, artboards, edges, canvas color and bounding box.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ViewTogglePalette.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ViewTogglePalette";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-03";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ViewTogglePalette.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ViewTogglePalette.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    var DIALOG_OPACITY = 0.98;   /* パレットの不透明度 / Palette opacity */

    // =========================================
    // レイアウト / Layout
    // =========================================

    var BUTTON_PADDING = 12;            /* ボタンの左右の余白（片側） / side padding of the buttons (each side) */
    var BUTTON_HEIGHT_TRIM = 4;         /* ボタンの高さを詰める量 / how much to trim the button height */

    // UIレイアウト（再利用パーツ） / UI layout (reusable)

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
    var TAB_MARGINS    = [15, 20, 5, 10];    /* タブ余白 [左,上,右,下] / tab margins */

    /**
     * ウィンドウの共通設定
     * @param {Window} targetWindow - 対象のウィンドウ
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(targetWindow, spacing) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルの共通設定（子は幅いっぱい。ボタンは alignment = "left" で広げない）
     * @param {Panel} targetPanel - 対象のパネル
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

    /**
     * タブの共通設定
     * @param {Tab} targetTab - 対象のタブ
     * @param {number} [spacing] - 要素間隔（省略時は変えない）
     * @returns {void}
     */
    function setupTab(targetTab, spacing) {
        targetTab.orientation = "column";
        targetTab.alignChildren = "fill";
        targetTab.margins = TAB_MARGINS;
        if (typeof spacing === "number") targetTab.spacing = spacing;
    }

    /**
     * 横並びの行グループの共通設定（ボタン列など）。
     * alignment と alignChildren を対で指定し、中のボタンが横に伸びたり天地がずれたりしないようにする
     * @param {Group} rowGroup - 対象のグループ
     * @param {string|string[]} [rowAlignment] - 横方向の alignment（省略時は "left"）。配列ならそのまま使う
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(rowGroup, rowAlignment, spacing) {
        rowGroup.orientation = "row";
        rowGroup.alignment = (rowAlignment instanceof Array) ? rowAlignment : [rowAlignment || "left", "center"];
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ボタンの高さを指定した px だけ詰める（レイアウトが決まったあとに呼ぶ）
     * @param {Button} targetButton - 対象のボタン
     * @param {number} trimPixels - 詰める量（px）
     * @returns {void}
     */
    function trimButtonHeight(targetButton, trimPixels) {
        /* レイアウト前は size が無い / size is not set until the layout runs */
        if (!targetButton.size) return;
        targetButton.size = [targetButton.size.width, targetButton.size.height - trimPixels];
    }

    // UIレイアウト（再利用パーツ）ここまで / End of the reusable UI layout

    // =========================================
    // ローカライズ / Localization
    // =========================================
    // ローカライズ（再利用パーツ） / Localization (reusable)

    /**
     * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /**
     * LABELS から今の UI 言語の文言を取り出す。
     * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
     * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
     * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
     */
    function getLabel(labelRef, placeholderValues) {
        var labelEntry = labelRef;
        if (typeof labelRef === "string") {
            var labelPathKeys = labelRef.split(".");
            labelEntry = LABELS;
            for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
                labelEntry = labelEntry[labelPathKeys[i]];
            }
        }
        var labelString;
        if (typeof labelEntry === "string") labelString = labelEntry;
        else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
        else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
        else return (typeof labelRef === "string") ? labelRef : "";
        return fillLabelPlaceholders(String(labelString), placeholderValues);
    }

    /**
     * 項目名の文言の末尾にコロンを付ける（日本語は半角スペース＋半角コロン「 :」、英語は「:」。Illustrator の線パネルなどの項目名に合わせる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? " :" : ":");
    }

    /**
     * 「項目名 : 値」の1行を返す（日本語は「件数 : 5」、英語は「Count: 5」。どちらもコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + " " + value;
    }

    /**
     * 文言の {name} や %1 に値を差し込む
     * @param {string} labelString - 文言
     * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
     * @returns {string} 差し込んだ文言
     */
    function fillLabelPlaceholders(labelString, placeholderValues) {
        if (placeholderValues == null) return labelString;
        if (placeholderValues instanceof Array) {
            /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
            for (var i = placeholderValues.length; i >= 1; i--) {
                labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
            }
            return labelString;
        }
        for (var placeholderKey in placeholderValues) {
            if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
            labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
        }
        return labelString;
    }

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

    /* 日英ラベル定義（カテゴリ分け）/ Japanese-English label definitions (categorized) */
    var LABELS = {
        dialog: {
            title: { ja: "表示の切り替え", en: "View Toggles" }
        },
        panel: {
            artboard: { ja: "アートボード", en: "Artboards" },
            objectSelection: { ja: "オブジェクト選択", en: "Object Selection" },
            ruler: { ja: "定規", en: "Rulers" },
            guide: { ja: "ガイド", en: "Guides" }
        },
        tooltip: {
            videoRuler: { ja: "ビデオ定規の表示を切り替え", en: "Toggle video ruler visibility" },
            globalRulers: { ja: "環境設定［すべてのドキュメントで定規を表示］を切り替え", en: "Toggle the Show Rulers in All Documents preference" },
            showGuides: { ja: "ガイドの表示を切り替え", en: "Toggle guide visibility" },
            lockGuides: { ja: "ガイドのロックを切り替え", en: "Toggle guide lock" },
            smartGuides: { ja: "スマートガイドのオン／オフを切り替え", en: "Toggle Smart Guides on/off" },
            artboard: { ja: "アートボードの表示を切り替え", en: "Toggle artboard visibility" },
            edges: { ja: "エッジの表示を切り替え", en: "Toggle edge visibility" },
            canvasColor: { ja: "カンバスカラーを「UIに合わせる」と「ホワイト」で切り替え", en: "Switch the canvas color between Match Brightness and White" },
            boundingBox: { ja: "バウンディングボックスの表示を切り替え", en: "Toggle bounding box visibility" }
        },
        button: {
            videoRuler: { ja: "ビデオ定規", en: "Video Ruler" },
            globalRulers: { ja: "すべてのドキュメントで表示", en: "Show in All Documents" },
            showGuides: { ja: "ガイドを表示", en: "Show Guides" },
            lockGuides: { ja: "ガイドをロック", en: "Lock Guides" },
            smartGuides: { ja: "スマートガイド", en: "Smart Guides" },
            artboard: { ja: "アートボード", en: "Artboards" },
            edges: { ja: "境界線", en: "Edges" },
            canvasColor: { ja: "カンバスカラー", en: "Canvas Color" },
            boundingBox: { ja: "バウンディングボックス", en: "Bounding Box" }
        }
    };

    // =========================================
    // BridgeTalk 委譲 / BridgeTalk delegation
    // =========================================

    /**
     * メインエンジン（target="illustrator"）へコードを送って実行する
     * @param {string} bodyCode - 実行するコード
     * @returns {void}
     */
    function runInMainEngine(bodyCode) {
        try {
            var bridge = new BridgeTalk();
            bridge.target = "illustrator"; /* #targetengine 指定なし＝メインエンジン / no engine = main engine */
            bridge.body = bodyCode;
            bridge.onError = function (message) {
                /* エラーは意図的に握りつぶす（常駐パレットなので alert は出さない）。
                   完全に無音だと失敗に気づけないため、デバッグ用に $.writeln にだけ残す。
                   Intentionally swallowed (no alert in a persistent palette);
                   logged to $.writeln only so failures stay noticeable while debugging. */
                $.writeln(SCRIPT_NAME + " BridgeTalk error: " + message.body);
            };
            bridge.send();
        } catch (e) {
            /* BridgeTalk 不可時は同一エンジンで直接実行 / Fallback: run directly in this engine */
            try {
                eval(bodyCode);
            } catch (e2) {
                // no-op
            }
        }
    }

    // =========================================
    // UI ヘルパー / UI helpers
    // =========================================

    /**
     * ボタンの幅を文字幅＋左右の余白に詰める（高さは既定のまま）
     * @param {Button} targetButton - 対象のボタン
     * @returns {void}
     */
    function fitButtonWidth(targetButton) {
        try {
            var measuredSize = targetButton.graphics.measureString(targetButton.text);
            var textWidth = (measuredSize.width !== undefined) ? measuredSize.width : measuredSize[0];
            /* 実測できないときは既定の幅のまま / Keep the default width when the text cannot be measured */
            if (isNaN(textWidth) || textWidth <= 0) return;
            targetButton.preferredSize.width = Math.ceil(textWidth) + BUTTON_PADDING * 2;
        } catch (e) {}
    }

    /**
     * トグルボタンを追加する（クリックでメインエンジンにコードを送る）
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelKey - LABELS.button / LABELS.tooltip のキー
     * @param {string} bodyCode - クリックで実行するコード
     * @returns {Button} 追加したボタン
     */
    function addToggleButton(parentPanel, labelKey, bodyCode) {
        var toggleButton = parentPanel.add('button', undefined, getLabel('button.' + labelKey));
        toggleButton.alignment = "left"; /* パネル幅に伸ばさない / Do not stretch to the panel width */
        toggleButton.helpTip = getLabel('tooltip.' + labelKey);
        toggleButton.onClick = function () {
            runInMainEngine(bodyCode);
        };
        fitButtonWidth(toggleButton);
        return toggleButton;
    }

    /**
     * ボタンを縦に並べる：［アートボード］パネルにアートボード／カンバスカラー、［オブジェクト選択］パネルに境界線／バウンディングボックス、［定規］パネルにビデオ定規／すべてのドキュメントで表示、［ガイド］パネルにガイドを表示／ガイドをロック／スマートガイド
     * @param {Window} parent - 追加先
     * @returns {Button[]} 追加したボタン（レイアウト確定後の trimButtonHeight 用）
     */
    function buildToggleButtons(parent) {
        /* アートボードとカンバスカラーはパネルにまとめる / Group Artboards and Canvas Color in a panel */
        var artboardPanel = parent.add('panel', undefined, getLabel('panel.artboard'));
        setupPanel(artboardPanel, 6);

        /* 境界線とバウンディングボックスはパネルにまとめる / Group Edges and Bounding Box in a panel */
        var objectSelectionPanel = parent.add('panel', undefined, getLabel('panel.objectSelection'));
        setupPanel(objectSelectionPanel, 6);

        /* ビデオ定規と定規の環境設定は［定規］パネルに入れる / Put Video Ruler and the ruler preference in the Rulers panel */
        var rulerPanel = parent.add('panel', undefined, getLabel('panel.ruler'));
        setupPanel(rulerPanel, 6);

        /* ガイドの表示・ロックとスマートガイドは［ガイド］パネルに入れる / Put the guide toggles and Smart Guides in the Guides panel */
        var guidePanel = parent.add('panel', undefined, getLabel('panel.guide'));
        setupPanel(guidePanel, 6);

        /* カンバスカラー：uiCanvasIsWhite を 0（UIに合わせる）と 1（ホワイト）で交互に切替。
           読み出しから反転・保存・再描画までをメインエンジン側で完結（エンジン間で値がずれないように）。
           反映には画面の強制再描画が必要。redraw() だけでは足りないため、ズーム操作で描き直す（PresetManager と同じ手当て）
           Canvas Color: flip uiCanvasIsWhite between 0 (Match Brightness) and 1 (White).
           Read, flip, save, and repaint all in the main engine so the two engines can't disagree on the value.
           Applying it needs a forced repaint; redraw() alone is not enough, so toggle the zoom (same workaround as PresetManager) */
        var canvasColorCode =
            'var isWhite = 0; ' +
            'try { isWhite = app.preferences.getIntegerPreference("uiCanvasIsWhite"); } catch (e) {} ' +
            'app.preferences.setIntegerPreference("uiCanvasIsWhite", (isWhite === 1) ? 0 : 1); ' +
            'try { app.redraw(); app.executeMenuCommand("zoomout"); app.executeMenuCommand("zoomin"); } catch (e) {}';

        /* すべてのドキュメントで定規を表示：useGlobalRulers を反転。読み出しから保存までメインエンジン側で完結
           Show Rulers in All Documents: flip useGlobalRulers, read and saved in the main engine */
        var globalRulersCode =
            'var useGlobal = false; ' +
            'try { useGlobal = app.preferences.getBooleanPreference("useGlobalRulers"); } catch (e) {} ' +
            'app.preferences.setBooleanPreference("useGlobalRulers", !useGlobal); ' +
            'try { app.redraw(); } catch (e) {}';

        return [
            addToggleButton(artboardPanel, 'artboard', "try{app.executeMenuCommand('artboard');}catch(e){}"),
            addToggleButton(artboardPanel, 'canvasColor', canvasColorCode),
            addToggleButton(objectSelectionPanel, 'edges', "try{app.executeMenuCommand('edge');}catch(e){}"),
            addToggleButton(objectSelectionPanel, 'boundingBox', "try{app.executeMenuCommand('AI Bounding Box Toggle');}catch(e){}"),
            addToggleButton(rulerPanel, 'videoRuler', "try{app.executeMenuCommand('videoruler');}catch(e){}"),
            addToggleButton(rulerPanel, 'globalRulers', globalRulersCode),
            addToggleButton(guidePanel, 'showGuides', "try{app.executeMenuCommand('showguide');}catch(e){}"),
            addToggleButton(guidePanel, 'lockGuides', "try{app.executeMenuCommand('lockguide');}catch(e){}"),
            addToggleButton(guidePanel, 'smartGuides', "try{app.executeMenuCommand('Snapomatic on-off menu item');}catch(e){}")
        ];
    }

    // =========================================
    // メイン処理 / Main process
    // =========================================

    /**
     * パレットを構築して表示する
     * @returns {void}
     */
    function main() {

        /* すでにパレットが開いていれば前面に出して終了 / If a palette already exists, bring it forward and return */
        try {
            if ($.global.__viewTogglePalette) {
                $.global.__viewTogglePalette.show();
                return;
            }
        } catch (e) {
            $.global.__viewTogglePalette = null;
        }

        var palette = new Window('palette', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
        setupWindow(palette, 6);
        palette.opacity = DIALOG_OPACITY;
        $.global.__viewTogglePalette = palette;

        /* 閉じたら参照をクリア（次回は再構築）/ Clear the reference on close (rebuild next time) */
        palette.onClose = function () {
            $.global.__viewTogglePalette = null;
        };

        /* パレットがアクティブなとき esc キーで閉じる / Close the palette with Esc while it is active */
        palette.addEventListener('keydown', function (event) {
            if (event.keyName === 'Escape') {
                palette.close();
            }
        });

        var toggleButtons = buildToggleButtons(palette);

        /* 表示時にレイアウトを再計算（描画欠け防止）/ Recalc the layout on show (avoid partial rendering) */
        palette.onShow = function () {
            palette.layout.layout(true);
            palette.layout.resize();
        };

        palette.show();

        /* ボタンの高さを詰める（レイアウト確定後に1回）/ Trim button heights (once, after layout) */
        for (var i = 0; i < toggleButtons.length; i++) {
            trimButtonHeight(toggleButtons[i], BUTTON_HEIGHT_TRIM);
        }
    }

    main();

}());
