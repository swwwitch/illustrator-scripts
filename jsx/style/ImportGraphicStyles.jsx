#target illustrator
#targetengine "ImportGraphicStylesEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

指定したAIファイルからグラフィックスタイルを取り込み、スタイル名のラジオボタンを自動生成して選択オブジェクトへ適用します。
指定したファイルは記憶され、次回以降は自動で参照します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ImportGraphicStyles.md

### Overview

Imports the graphic styles from an AI file you choose, builds a radio button per style name, and applies the selected one to the selection.
The chosen file is remembered and reused on later runs.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ImportGraphicStyles.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ImportGraphicStyles";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.7.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-14";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ImportGraphicStyles.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ImportGraphicStyles.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 現在のロケールから言語コードを取得 / Get language code from locale */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    /**
     * グラフィックスタイルのラジオボタンを追加する
     * @param {Group|Panel} parentContainer - 追加先のコンテナ
     * @param {string} styleName - スタイル名
     * @returns {RadioButton} 追加したラジオボタン
     */
    function addStyleRadio(parentContainer, styleName) {
        var radioButton = parentContainer.add("radiobutton", undefined, styleName);
        radioButton.helpTip = getLabel("tooltip.style");
        return radioButton;
    }

    var LABELS = {
        /* ダイアログ / Dialog */
        dialog: {
            title: { ja: "スタイルを適用", en: "Apply Style" },
            pickFile: { ja: "スタイルの AI ファイルを選択", en: "Select a style AI file" }
        },
        /* スタイル選択 / Style choices */
        style: {
            panelTitle: { ja: "スタイル", en: "Style" }
        },
        /* スタイルの読み込み / Load styles */
        load: {
            panelTitle: { ja: "スタイルの読み込み", en: "Load Styles" }
        },
        /* ボタン / Buttons */
        tooltip: {
            loadFile: { ja: "グラフィックスタイルを読み込む .ai ファイルを選びます。", en: "Picks the .ai file to load the graphic styles from." },
            style: { ja: "選択オブジェクトに適用するグラフィックスタイルです。", en: "The graphic style applied to the selected objects." }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            apply: { ja: "適用", en: "Apply" },
            loadFile: { ja: "読み込み", en: "Load" }
        },
        /* 実行時メッセージ / Runtime messages */
        message: {
            openDocFirst: {
                ja: "元のドキュメントを開いてから実行してください。",
                en: "Please open the destination document first."
            },
            selectObjectFirst: {
                ja: "スタイルを適用するオブジェクトを選択してから実行してください。",
                en: "Please select the object(s) to apply the style to."
            },
            fileNotFoundTitle: { ja: "ファイルが見つかりません", en: "File Not Found" },
            fileNotFoundBody: {
                ja: "指定したファイルが見つかりません：\n",
                en: "The specified file was not found:\n"
            },
            styleNotFound: {
                ja: "指定したグラフィックスタイルが見つかりません：\n",
                en: "The graphic style was not found:\n"
            },
            noFileSelected: { ja: "ファイル未選択", en: "No file selected" },
            noStylesHint: {
                ja: "「スタイルを読み込み」でファイルを選択してください。",
                en: "Click “Load Styles” to choose a file."
            }
        }
    };

    /* 指定キーのローカライズ文字列を取得（ドット区切りパス対応）/ Resolve localized string by dotted key path */
    function getLabel(key) {
        var parts = key.split(".");
        var node = LABELS;
        for (var i = 0; i < parts.length; i++) {
            if (node == null) break;
            node = node[parts[i]];
        }
        return (node && node[currentLanguage] != null) ? node[currentLanguage] : ("[" + key + "]");
    }

    // =========================================
    // ユーティリティ / Utilities
    // =========================================

    /* 取り込み先レイヤーを取得（無ければ作成）/ Get or create the destination layer "// _imported" */
    function getOrCreateImportLayer(destinationDoc) {
        var importLayerName = "// _imported";
        var importLayer;
        try {
            importLayer = destinationDoc.layers.getByName(importLayerName);
        } catch (e) {
            importLayer = destinationDoc.layers.add();
            importLayer.name = importLayerName;
        }
        importLayer.locked = false;
        importLayer.visible = true;
        return importLayer;
    }

    /* パスから表示用ファイル名を取得（文字化け対策で decodeURI）/ Get a display-friendly filename from a path */
    function getDisplayFileName(filePath) {
        try {
            return decodeURI(new File(filePath).name);
        } catch (e) {
            return filePath;
        }
    }

    /* 名前でグラフィックスタイルを取得（無ければ null）/ Get a graphic style by name (null if none) */
    function findGraphicStyle(destinationDoc, styleName) {
        try {
            return destinationDoc.graphicStyles.getByName(styleName);
        } catch (e) {
            return null;
        }
    }

    // =========================================
    // 設定の保存・読込 / Preferences (remember last file)
    // =========================================

    /* 設定ファイル（前回のスタイルファイルのパスを記憶）/ Prefs file that remembers the last style file */
    function getPrefsFile() {
        return new File(Folder.userData + "/swwwitch_ImportGraphicStyles.txt");
    }

    /* 記憶しているスタイルファイルのパスを読み込む（無ければ空文字）/ Load the remembered style-file path */
    function loadSavedFilePath() {
        var prefsFile = getPrefsFile();
        if (!prefsFile.exists) return "";
        var savedPath = "";
        try {
            prefsFile.encoding = "UTF-8";
            prefsFile.open("r");
            var content = prefsFile.read();
            prefsFile.close();
            var lines = content.split(/\r\n|\r|\n/);
            for (var i = 0; i < lines.length; i++) {
                var separatorIndex = lines[i].indexOf("=");
                if (separatorIndex < 0) continue;
                if (lines[i].substring(0, separatorIndex) === "styleFilePath") {
                    savedPath = lines[i].substring(separatorIndex + 1);
                }
            }
        } catch (e) {}
        return savedPath;
    }

    /* スタイルファイルのパスを記憶する（key=value 形式）/ Remember the style-file path (key=value) */
    function saveFilePath(filePath) {
        var prefsFile = getPrefsFile();
        try {
            prefsFile.encoding = "UTF-8";
            prefsFile.open("w");
            prefsFile.write("styleFilePath=" + filePath + "\n");
            prefsFile.close();
        } catch (e) {}
    }

    /* スタイル用 AI ファイルを選ばせる（キャンセルで空文字）/ Let the user pick a style AI file */
    function pickStyleFile() {
        var picked = File.openDialog(getLabel("dialog.pickFile"), function (candidate) {
            return (candidate instanceof Folder) || /\.ai$/i.test(candidate.name);
        });
        return picked ? picked.fsName : "";
    }

    // =========================================
    // スタイル取り込み / Import Styles
    // =========================================

    /* 対象AIを開いてスタイル名を取得→コピー→一時レイヤーへ貼り付け→レイヤーごと削除（アセットのみ登録）
       戻り値: 取り込んだグラフィックスタイル名の配列（ファイル未検出時は警告して null）
       Open the AI, read its style names, copy, paste to a temp layer, then remove it (assets only stay registered).
       Returns imported graphic-style names (null if the file is missing) */
    function importStylesFrom(destinationDoc, filePath) {
        var styleFile = new File(filePath);
        if (!styleFile.exists) {
            alert(getLabel("message.fileNotFoundTitle") + "\n" + getLabel("message.fileNotFoundBody") + getDisplayFileName(filePath));
            return null;
        }
        var styleSourceDoc = app.open(styleFile);

        // 元ファイルのグラフィックスタイル名を取得（index 0 の既定スタイルは除外）
        // Collect style names from the source (skip the default style at index 0)
        var sourceStyleNames = [];
        for (var i = 1; i < styleSourceDoc.graphicStyles.length; i++) {
            sourceStyleNames.push(styleSourceDoc.graphicStyles[i].name);
        }

        // 作業アートボード内の全てをコピーして保存せず閉じる / Copy in-artboard objects, close without saving
        app.executeMenuCommand("selectallinartboard");
        app.executeMenuCommand("copy");
        styleSourceDoc.close(SaveOptions.DONOTSAVECHANGES);

        // 一時レイヤーへ貼り付け→パネル登録後にレイヤーごと削除 / Paste to temp layer, then remove it
        app.activeDocument = destinationDoc;
        var importLayer = getOrCreateImportLayer(destinationDoc);
        destinationDoc.activeLayer = importLayer;
        app.executeMenuCommand("paste");
        try {
            importLayer.remove();
        } catch (e) {}

        // 実際に書類へ登録されたスタイル名だけを返す / Keep only names actually registered in the destination
        var importedStyleNames = [];
        for (var k = 0; k < sourceStyleNames.length; k++) {
            if (findGraphicStyle(destinationDoc, sourceStyleNames[k])) importedStyleNames.push(sourceStyleNames[k]);
        }
        return importedStyleNames;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子は DIALOG_* / prepareDialogWindow / *DialogLeft* / getSelectionViewSpan の名前
    // 2. スクリプトの先頭（#target の次の行）に #targetengine "<SCRIPT_NAME>Engine" を置く。
    //    #targetengine が無いと $.global が実行ごとに消え、位置を覚えられない。すでにあればそのまま使う
    // 3. ダイアログの show() の直前で prepareDialogWindow(dialog, SCRIPT_NAME) を呼ぶ。
    //    それまでに入れた onShow / onMove / onClose はそのまま生かし、あとに位置の復元・記録をつなぐ
    //      prepareDialogWindow(mainDialog, SCRIPT_NAME);
    //      var dialogResult = mainDialog.show();
    //    同じスクリプトで複数のダイアログを開くときは、2つ目以降のキーを変える（SCRIPT_NAME + "_colorPicker" など）
    //    同じダイアログを何度も開くときも、毎回 show() の直前で呼んでよい（2回目からは選択範囲を測り直すだけ）
    // 4. 初めて開くとき（記録が無いとき）は、スクリプト側の配置（中央・オフセットなど）がそのまま効く
    // 5. 開く位置が選択中のオブジェクトに重なりそうなら左右の反対側へずらす（Illustrator のみ）。
    //    ずらした位置は記録せず、ユーザーが動かしたときだけ記録する
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var DIALOG_OPACITY = 0.97;       /* ダイアログの不透明度 / dialog opacity */
    var DIALOG_AVOID_MARGIN = 60;    /* 選択範囲の推定位置の両側に取る余裕（px）/ margin on each side of the estimated selection (px) */
    var DIALOG_AVOID_MAX_ITEMS = 100; /* 選択範囲を測るオブジェクトの上限 / max items measured for the selection bounds */

    /**
     * ダイアログの不透明度を設定し、前回閉じた位置で開いて、動かした位置を記録するようにする。
     * 開く位置が選択中のオブジェクトに重なりそうなときは、左右の反対側へずらす（Illustrator のみ）。
     * 既存の onShow / onMove / onClose は先に呼んでから、位置の復元・記録を行う。
     * @param {Window} dialog - 対象のダイアログ
     * @param {string} storageKey - 位置を覚えるキー（ふつうは SCRIPT_NAME）
     * @returns {void}
     */
    function prepareDialogWindow(dialog, storageKey) {
        /* 同じダイアログを開き直すときは、選択範囲を測り直すだけにする（ハンドラーを重ねない）
           When the same dialog is shown again, only re-measure the selection (don't stack handlers) */
        if (dialog.dialogWindowState) {
            dialog.dialogWindowState.selectionSpan = getSelectionViewSpan();
            dialog.dialogWindowState.avoidedLocation = null;
            return;
        }
        var locationKey = "__" + storageKey + "_DialogLocation";
        var previousOnShow = dialog.onShow;
        var previousOnMove = dialog.onMove;
        var previousOnClose = dialog.onClose;
        var windowState = {
            selectionSpan: getSelectionViewSpan(), /* 選択範囲は show() の前に測る / measured before show() */
            screenWidth: null,                     /* 最初に開いたときに推定する / estimated on the first show */
            avoidedLocation: null                  /* 避けるためにずらした位置（記録しない）/ location set to avoid the selection (not remembered) */
        };
        dialog.dialogWindowState = windowState;

        dialog.opacity = DIALOG_OPACITY;

        /* 今の位置を記録する / Remember the current location */
        function rememberDialogLocation() {
            var currentLocation = [dialog.location[0], dialog.location[1]];
            var avoidedLocation = windowState.avoidedLocation;
            if (avoidedLocation && currentLocation[0] === avoidedLocation[0] && currentLocation[1] === avoidedLocation[1]) return;
            $.global[locationKey] = currentLocation;
        }

        dialog.onShow = function () {
            /* 最初に開くときの既定の位置は画面の横中央なので、画面の幅を逆算できる。2回目からは前回の位置なので使い回す
               On the first show the default location is centered horizontally, which gives the screen width; reuse it afterwards */
            if (windowState.screenWidth === null) windowState.screenWidth = dialog.location[0] * 2 + dialog.bounds.width;
            if (previousOnShow) previousOnShow.apply(this, arguments);
            /* $.screens は実際の画面の大きさと合わない（Mac で 1280×524 など）ので、画面内かは判定しない
               $.screens does not match the real display (e.g. 1280x524 on a Mac), so no on-screen check */
            var savedLocation = $.global[locationKey];
            if (savedLocation) dialog.location = [savedLocation[0], savedLocation[1]];
            if (windowState.selectionSpan) {
                var avoidLeft = findDialogLeftAvoidingSelection(dialog.location[0], dialog.bounds.width, windowState.screenWidth, windowState.selectionSpan);
                if (avoidLeft !== null) {
                    dialog.location = [avoidLeft, dialog.location[1]];
                    /* 代入後の値で比べる（丸められることがある）/ Compare with the value after assignment, which may be rounded */
                    windowState.avoidedLocation = [dialog.location[0], dialog.location[1]];
                }
            }
        };
        dialog.onMove = function () {
            if (previousOnMove) previousOnMove.apply(this, arguments);
            rememberDialogLocation();
        };
        dialog.onClose = function () {
            rememberDialogLocation();
            /* false を返すと閉じるのを取りやめるので、戻り値は元の onClose のものを返す
               Returning false cancels the close, so pass the original onClose result through */
            if (previousOnClose) return previousOnClose.apply(this, arguments);
        };
    }

    /**
     * 選択中のオブジェクトが、ドキュメントの表示域の左端から画面上で何 px の範囲にあるかを返す。
     * @returns {{left: number, right: number, viewWidth: number}|null} 選択が無い・測れないときは null
     */
    function getSelectionViewSpan() {
        try {
            if (app.name !== "Adobe Illustrator" || !app.documents.length) return null;
            var targetDoc = app.activeDocument;
            var selectedItems = targetDoc.selection;
            if (!selectedItems || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
            var itemCount = Math.min(selectedItems.length, DIALOG_AVOID_MAX_ITEMS);
            var spanLeft = Infinity;
            var spanRight = -Infinity;
            for (var i = 0; i < itemCount; i++) {
                var itemBounds = selectedItems[i].visibleBounds;
                if (itemBounds[0] < spanLeft) spanLeft = itemBounds[0];
                if (itemBounds[2] > spanRight) spanRight = itemBounds[2];
            }
            var activeView = targetDoc.activeView; /* 複数ウィンドウで開いていても今のウィンドウ / the current window even with multiple windows */
            var viewBounds = activeView.bounds;
            var zoom = activeView.zoom;
            var viewWidth = (viewBounds[2] - viewBounds[0]) * zoom;
            /* 表示域の外にはみ出した部分は数えない / Ignore the part outside the view */
            var left = Math.max(0, (spanLeft - viewBounds[0]) * zoom);
            var right = Math.min(viewWidth, (spanRight - viewBounds[0]) * zoom);
            if (right <= left) return null;
            return { left: left, right: right, viewWidth: viewWidth };
        } catch (e) {
            /* テキスト編集中など測れないときは避けない / Do not avoid when it cannot be measured, e.g. while editing text */
            return null;
        }
    }

    /**
     * ダイアログが選択範囲に重なるなら、重ならない左端の位置を返す。
     * 表示域は画面の横中央にあるとみなし、ずれは DIALOG_AVOID_MARGIN で吸収する。
     * @param {number} dialogLeft - 今のダイアログの左端
     * @param {number} dialogWidth - ダイアログの幅
     * @param {number} screenWidth - 画面の幅
     * @param {{left: number, right: number, viewWidth: number}} selectionSpan - getSelectionViewSpan() の結果
     * @returns {number|null} ずらした左端。重ならない・どちらにも収まらないときは null
     */
    function findDialogLeftAvoidingSelection(dialogLeft, dialogWidth, screenWidth, selectionSpan) {
        var viewLeft = (screenWidth - selectionSpan.viewWidth) / 2;
        var avoidLeft = viewLeft + selectionSpan.left - DIALOG_AVOID_MARGIN;
        var avoidRight = viewLeft + selectionSpan.right + DIALOG_AVOID_MARGIN;
        if (dialogLeft + dialogWidth <= avoidLeft || dialogLeft >= avoidRight) return null;

        var leftSideLeft = avoidLeft - dialogWidth;   /* 選択範囲の左に置くとき / placed left of the selection */
        var rightSideLeft = avoidRight;               /* 選択範囲の右に置くとき / placed right of the selection */
        var fitsLeft = leftSideLeft >= 0;
        var fitsRight = rightSideLeft + dialogWidth <= screenWidth;
        /* 選択範囲が画面の右寄りなら左へ、左寄りなら右へ逃がす / Move away from the side the selection leans to */
        var preferLeft = (avoidLeft + avoidRight) / 2 > screenWidth / 2;
        if (preferLeft && fitsLeft) return leftSideLeft;
        if (fitsRight) return rightSideLeft;
        if (fitsLeft) return leftSideLeft;
        return null;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /* スタイルを選ばせ、選んだグラフィックスタイル名を返す（キャンセルで null）
       「スタイルを読み込み」ボタンで別ファイルを指定すると、その場でラジオを組み直す
       Ask which style, return its name. The "Load Styles" button re-imports and rebuilds the radios in place */
    function runStyleDialog(destinationDoc, initialFilePath, initialStyleNames) {
        var state = {
            filePath: initialFilePath || "",
            styleNames: initialStyleNames || []
        };

        var styleDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        styleDialog.orientation = "column";
        styleDialog.alignChildren = ["fill", "top"];

        // スタイル選択パネル / Style panel
        var stylePanel = styleDialog.add("panel", undefined, getLabel("style.panelTitle"));
        stylePanel.orientation = "column";
        stylePanel.alignChildren = ["left", "top"];
        stylePanel.margins = [16, 20, 16, 12];
        stylePanel.spacing = 8;

        // ラジオを差し替えるためのコンテナ / Container whose radios get rebuilt on reload
        var radioGroup = stylePanel.add("group");
        radioGroup.orientation = "column";
        radioGroup.alignChildren = ["left", "top"];
        radioGroup.spacing = 8;
        var styleRadios = [];

        // スタイルの読み込みパネル（ボタンの下にファイルパスを表示）/ Load-styles panel (path shown below the button)
        var loadPanel = styleDialog.add("panel", undefined, getLabel("load.panelTitle"));
        loadPanel.orientation = "column";
        loadPanel.alignChildren = ["left", "top"];
        loadPanel.margins = [16, 20, 16, 12];
        loadPanel.spacing = 8;
        var loadFileButton = loadPanel.add("button", undefined, getLabel("button.loadFile"));
        loadFileButton.helpTip = getLabel("tooltip.loadFile");
        var fileNameText = loadPanel.add("statictext", undefined, "", { truncate: "middle" });
        fileNameText.preferredSize.width = 240;

        // ボタン行（Mac 規約: キャンセル → 適用）/ Buttons (Mac order: Cancel → Apply)
        var dialogButtonGroup = styleDialog.add("group");
        dialogButtonGroup.alignment = ["right", "bottom"];
        dialogButtonGroup.alignChildren = ["right", "center"];
        dialogButtonGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var applyButton = dialogButtonGroup.add("button", undefined, getLabel("button.apply"), { name: "ok" });

        /* 選択中のファイル名表示を更新 / Update the file-name label */
        function refreshFileLabel() {
            fileNameText.text = state.filePath ? getDisplayFileName(state.filePath) : getLabel("message.noFileSelected");
        }

        /* ラジオを現在のスタイル名で組み直す / Rebuild radios from the current style names */
        function rebuildRadios() {
            for (var i = radioGroup.children.length - 1; i >= 0; i--) radioGroup.remove(radioGroup.children[i]);
            styleRadios = [];
            if (state.styleNames.length === 0) {
                radioGroup.add("statictext", undefined, getLabel("message.noStylesHint"));
                applyButton.enabled = false;
            } else {
                for (var j = 0; j < state.styleNames.length; j++) {
                    styleRadios.push(addStyleRadio(radioGroup, state.styleNames[j]));
                }
                styleRadios[0].value = true; // 既定 / Default
                applyButton.enabled = true;
            }
            styleDialog.layout.layout(true);
            styleDialog.layout.resize();
        }

        // onClick で連結（addEventListener は発火しない環境があるため）/ Use onClick, not addEventListener
        loadFileButton.onClick = function () {
            var pickedPath = pickStyleFile();
            if (!pickedPath) return;
            var importedStyleNames = importStylesFrom(destinationDoc, pickedPath);
            if (importedStyleNames === null) return; // ファイル未検出は importStylesFrom 側で警告済み
            state.filePath = pickedPath;
            state.styleNames = importedStyleNames;
            saveFilePath(pickedPath); // 次回以降このファイルを参照 / Remember for next runs
            refreshFileLabel();
            rebuildRadios();
        };

        refreshFileLabel();
        rebuildRadios();

        prepareDialogWindow(styleDialog, SCRIPT_NAME);
        if (styleDialog.show() !== 1) return null;
        for (var k = 0; k < styleRadios.length; k++) {
            if (styleRadios[k].value) return state.styleNames[k];
        }
        return null;
    }

    // =========================================
    // メイン処理 / Main Process
    // =========================================

    /* 記憶したファイルから取り込み→スタイル選択→選択オブジェクトへ適用 / Import from the remembered file, choose, apply */
    function main() {
        // ドキュメントの存在を確認 / Check for an open document
        if (app.documents.length === 0) {
            alert(getLabel("message.openDocFirst"));
            return;
        }
        var destinationDoc = app.activeDocument;

        // 選択オブジェクトを確保 / Capture current selection
        var currentSelection = destinationDoc.selection;
        if (!currentSelection || currentSelection.length === 0) {
            alert(getLabel("message.selectObjectFirst"));
            return;
        }
        var selectedItems = [];
        for (var i = 0; i < currentSelection.length; i++) selectedItems.push(currentSelection[i]);

        // 記憶しているファイルがあればダイアログ前に取り込む / If a file is remembered, import it before the dialog
        var savedFilePath = loadSavedFilePath();
        var initialFilePath = "";
        var initialStyleNames = [];
        if (savedFilePath) {
            initialFilePath = savedFilePath;
            if (new File(savedFilePath).exists) {
                var importedStyleNames = importStylesFrom(destinationDoc, savedFilePath);
                if (importedStyleNames !== null) initialStyleNames = importedStyleNames;
            }
            // ファイルが見つからない場合はダイアログの「スタイルを読み込み」で選び直してもらう
            // If the file is missing, the user re-picks it via "Load Styles"
        }

        // スタイルを選択（読み込みボタンで別ファイルにも切替可）/ Choose the style (Load Styles can switch files)
        var styleName = runStyleDialog(destinationDoc, initialFilePath, initialStyleNames);
        if (!styleName) return;

        // スタイルを取得 / Get the style
        var graphicStyle = findGraphicStyle(destinationDoc, styleName);
        if (!graphicStyle) {
            alert(getLabel("message.styleNotFound") + styleName);
            return;
        }

        // 選択オブジェクトへ適用 / Apply to the selected objects
        for (var j = 0; j < selectedItems.length; j++) {
            try {
                graphicStyle.applyTo(selectedItems[j]);
            } catch (e) {}
        }

        // 選択状態を復帰 / Restore selection
        destinationDoc.selection = selectedItems;
    }

    main();

})();
