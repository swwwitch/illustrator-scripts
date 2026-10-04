#target illustrator
#targetengine "FitAndConvertToPointTypeEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したエリア内文字のうち、あふれているものだけ自動サイズ調整をONにして解消してから、ポイント文字に変換します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FitAndConvertToPointType.md

### Overview

Turns on auto-sizing for whichever of the selected area texts are overset, resolves the overflow, and then converts them to point text.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FitAndConvertToPointType.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FitAndConvertToPointType";     /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.9";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-08-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FitAndConvertToPointType.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FitAndConvertToPointType.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 「強制改行を削除」の初期状態 / Initial state of "Remove forced line breaks" */
    var DEFAULT_REMOVE_LINE_BREAKS = false;

    /* 強制改行（ソフトリターン）と見なす文字。段落改行（\r）は残す
       Characters treated as forced line breaks (soft returns); paragraph returns (\r) are kept */
    var FORCED_BREAK_PATTERN = /[\u0003\n\u2028]/;

    // =========================================
    // レイアウト / Layout
    // =========================================

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

    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)

    var DIALOG_OPACITY = 0.98;       /* ダイアログの不透明度 / dialog opacity */
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
            /* 文字ツールで文字を選択しているときは TextRange が返り、[0] が無い / Selecting characters with the Type tool returns a TextRange, which has no [0] */
            if (!selectedItems || selectedItems.typename === "TextRange" || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
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

    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity

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

    /* 日英ラベル定義（カテゴリ別）/ Japanese-English labels grouped by category */
    var LABELS = {
        dialog: {
            title: { ja: "ポイント文字に変換", en: "Convert to Point Type" }
        },
        checkbox: {
            removeLineBreaks: { ja: "強制改行を削除", en: "Remove forced line breaks" }
        },
        tooltip: {
            removeLineBreaks: {
                ja: "変換後に残る強制改行（ソフトリターン）を取り除き、行をつなげます。段落改行は残ります。",
                en: "Strips the forced line breaks (soft returns) left after the conversion, joining the lines. Paragraph returns are kept."
            }
        },
        button: {
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectAreaText: { ja: "エリア内文字を選択してください。", en: "Please select area type." },
            noTarget: {
                ja: "変換できるエリア内文字がありませんでした。ロックや非表示になっていないか確認してください。",
                en: "No area type could be converted. Check whether the frames are locked or hidden."
            },
            notSupported: {
                ja: "お使いのIllustratorはポイント文字への変換に対応していません。",
                en: "This version of Illustrator cannot convert area type to point text."
            },
            partialFailure: {
                ja: "{done}件を変換しました。{failed}件は変換できませんでした（連結されたテキストなどは対象外です）。",
                en: "Converted {done}. {failed} could not be converted (threaded text and the like are not supported)."
            },
            actionFailed: { ja: "アクションを実行できませんでした。", en: "Could not run the action." }
        }
    };

    // 一時アクション（再利用パーツ） / Temporary action (reusable)

    /**
     * 文字列を UTF-8 のバイト列の16進にする（アクション定義の /name・/localizedName 用）
     * @param {string} sourceText - 変換する文字列
     * @returns {string} 16進の文字列（2文字で1バイト）
     */
    function toActionHex(sourceText) {
        var utf8Text = unescape(encodeURIComponent(String(sourceText)));
        var hexText = "";
        for (var i = 0; i < utf8Text.length; i++) {
            var hexByte = utf8Text.charCodeAt(i).toString(16);
            hexText += (hexByte.length < 2 ? "0" : "") + hexByte;
        }
        return hexText;
    }

    /**
     * アクション定義の「/name [ バイト数 16進 ]」の3行を返す
     * @param {string} indent - 行頭の字下げ（"\t" など）
     * @param {string} nameText - 名前
     * @param {string} [fieldName] - 項目名（既定は "name"。"localizedName" など）
     * @returns {string[]} 3行ぶんの配列
     */
    function buildActionNameLines(indent, nameText, fieldName) {
        var nameHex = toActionHex(nameText);
        return [
            indent + "/" + (fieldName || "name") + " [ " + (nameHex.length / 2),
            indent + "\t" + nameHex,
            indent + "]"
        ];
    }

    /**
     * アクション定義を一時ファイルに書き出してセットを読み込む。読み込んだら一時ファイルは消す
     * （読み込んだ時点で解釈済みなので、以降の失敗でファイルが残らない）
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @returns {boolean} 読み込めたら true
     */
    function loadTemporaryActionSet(actionSource, setName) {
        var actionFile = new File(Folder.temp + "/" + setName + "_" + new Date().getTime() + ".aia");
        try {
            actionFile.encoding = "UTF-8";
            if (!actionFile.open("w")) throw new Error("cannot open " + actionFile.fsName);
            actionFile.write(actionSource);
            actionFile.close();
            /* 前回の失敗で同じ名前のセットが残っていれば外す / Remove a same-name set left by an earlier failure */
            unloadTemporaryActionSet(setName);
            app.loadAction(actionFile);
            return true;
        } catch (e) {
            $.writeln("loadTemporaryActionSet: " + e);
            return false;
        } finally {
            try { actionFile.close(); } catch (closeError) { /* 閉じ済み / already closed */ }
            try { actionFile.remove(); } catch (removeError) { /* 消せなくても続ける / keep going */ }
        }
    }

    /**
     * 一時アクションのセットを解除する（読み込まれていなくてもエラーにしない）
     * @param {string} setName - アクションセット名
     * @returns {void}
     */
    function unloadTemporaryActionSet(setName) {
        try {
            app.unloadAction(setName, "");
        } catch (e) {
            /* 読み込まれていない / not loaded */
        }
    }

    /**
     * アクション定義を読み込んで1回実行し、解除する。途中で失敗しても解除は必ず試みる
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @param {string} actionName - 実行するアクション名
     * @returns {boolean} 実行できたら true
     */
    function runTemporaryAction(actionSource, setName, actionName) {
        if (!loadTemporaryActionSet(actionSource, setName)) return false;
        try {
            app.doScript(actionName, setName);
            return true;
        } catch (e) {
            $.writeln("runTemporaryAction: " + e);
            return false;
        } finally {
            unloadTemporaryActionSet(setName);
        }
    }

    // 一時アクション（再利用パーツ）ここまで / End of the reusable temporary action

    // =========================================
    // ダイナミックアクション / Dynamic actions
    //   自動サイズ調整はDOMから設定できないため、アクション経由で切り替える
    //   Auto-sizing cannot be set from the DOM, so it is toggled through an action
    // =========================================

    /**
     * アクション名ブロック /name [ <len> <hex> ] を生成する
     * @param {string} actionName - アクション名またはセット名
     * @returns {string} 名前ブロックの文字列
     */
    function buildActionNameBlock(actionName) {
        return "/name [ " + actionName.length + " " + toActionHex(actionName).toUpperCase() + " ]";
    }

    /**
     * 自動サイズ調整アクションセットの定義（.aia 文字列）を組み立てる
     * @param {string} setName - アクションセット名
     * @returns {string} .aia 形式のアクションセット定義
     */
    function buildAutoSizeActionSetAia(setName) {
        return "/version 3" +
            buildActionNameBlock(setName) +
            "/isOpen 1" +
            "/actionCount 1" +
            "/action-1 {" +
            " " + buildActionNameBlock("AutoSizeOn") +
            " /keyIndex 0" +
            " /colorIndex 0" +
            " /isOpen 1" +
            " /eventCount 1" +
            " /event-1 {" +
            " /useRulersIn1stQuadrant 0" +
            " /internalName (adobe_SLOAreaTextDialog)" +
            " /localizedName [ 33 e382a8e383aae382a2e58685e69687e5ad97e382aae38397e382b7e383a7e383b3 ]" +
            " /isOpen 0" +
            " /isOn 1" +
            " /hasDialog 0" +
            " /parameterCount 1" +
            " /parameter-1 {" +
            " /key 1952539754" +
            " /showInPalette 4294967295" +
            " /type (integer)" +
            " /value 1" +
            " }" +
            " }" +
            "}";
    }

    var ACTION_SET_AUTO_SIZE = SCRIPT_NAME + "_AutoSize";

    /**
     * 自動サイズ調整アクションを一時ファイル経由で読み込む（スクリプト開始時に1回）
     * @returns {boolean} 読み込めたら true
     */
    function loadAutoSizeAction() {
        return loadTemporaryActionSet(buildAutoSizeActionSetAia(ACTION_SET_AUTO_SIZE), ACTION_SET_AUTO_SIZE);
    }

    /**
     * 読み込んだアクションセットを破棄する（スクリプト終了時）
     * @returns {void}
     */
    function unloadAutoSizeAction() {
        unloadTemporaryActionSet(ACTION_SET_AUTO_SIZE);
    }

    /**
     * 選択中のエリア内文字に自動サイズ調整をかける（アクションは選択に効く）
     * @returns {void}
     */
    function runAutoSizeAction() {
        try { app.doScript("AutoSizeOn", ACTION_SET_AUTO_SIZE, false); } catch (e) { }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_BOTTOM_MARGIN = 14; /* ボタン行の下の余白。ダイアログの下余白と合わせて約30px（Illustrator 標準のダイアログに合わせる） / bottom margin; with the dialog margin about 30px, like Illustrator's own dialogs */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
    var BUTTON_ROW_CENTER_MAX_WIDTH = 200; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

    /**
     * ダイアログ下部のボタン行を作る。
     * 通常は「左のグループ・伸びるスペーサー・右のグループ」、centered なら行そのものを左右中央に置く
     * @param {Window|Group|Panel} parent - 行を足す先（ふつうはダイアログ）
     * @param {Object} [rowOptions] - { centered: true } で左右中央に並べる
     * @returns {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} 行と左右のグループ（centered のときは左右が null）
     */
    function addButtonRow(parent, rowOptions) {
        var isCentered = !!(rowOptions && rowOptions.centered);
        var btnRowGroup = parent.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, BUTTON_ROW_BOTTOM_MARGIN];
        btnRowGroup.spacing = BUTTON_ROW_SPACING;

        if (isCentered) {
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            return { rowGroup: btnRowGroup, leftGroup: null, rightGroup: null };
        }

        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        btnLeftGroup.spacing = BUTTON_ROW_SPACING;

        /* 余りの幅を吸って、右のグループを右端に寄せる / Absorbs the extra width so the right group sits at the right edge */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.spacing = BUTTON_ROW_SPACING;

        return { rowGroup: btnRowGroup, leftGroup: btnLeftGroup, rightGroup: btnRightGroup };
    }

    /**
     * 左のグループにボタンが無い（右のボタンだけの）行を、ダイアログの幅に合わせて揃える。
     * 内側の幅（左右の余白を除く）が BUTTON_ROW_CENTER_MAX_WIDTH 以下なら左右中央、それより広ければ右揃えのまま。
     * 幅はレイアウトが決まるまで分からないので、ダイアログを表示した時点（show イベント）で判定する。
     * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
     * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
     * @returns {void}
     */
    function alignRightOnlyButtonRow(buttonRow) {
        if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
        var dialogWindow = buttonRow.rowGroup.window;
        dialogWindow.addEventListener("show", function () {
            if (!buttonRow.leftGroup) return;
            var btnRowGroup = buttonRow.rowGroup;
            /* 行の幅＝ダイアログの内側の幅（左右の余白を除く）/ The row spans the dialog's inner width (margins excluded) */
            if (!btnRowGroup.size || btnRowGroup.size.width > BUTTON_ROW_CENTER_MAX_WIDTH) return;
            /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
            btnRowGroup.remove(buttonRow.leftGroup);
            btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            buttonRow.leftGroup = null;
            dialogWindow.layout.layout(true);
        });
    }

    // ボタン行（再利用パーツ）ここまで / End of the reusable button row

    /**
     * オプションダイアログを表示する
     * @returns {{removeLineBreaks: boolean}|null} 設定（キャンセル時は null）
     */
    function showOptionsDialog() {
        var optionsDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(optionsDialog);

        var optionRow = optionsDialog.add("group");
        setupRow(optionRow);
        var removeLineBreaksCheckbox = optionRow.add("checkbox", undefined, getLabel("checkbox.removeLineBreaks"));
        removeLineBreaksCheckbox.value = DEFAULT_REMOVE_LINE_BREAKS;
        removeLineBreaksCheckbox.helpTip = getLabel("tooltip.removeLineBreaks");

        /* ボタンバー（Mac 規約で キャンセル → OK）/ Button bar (Cancel → OK per macOS) */
        var buttonRow = addButtonRow(optionsDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(optionsDialog, SCRIPT_NAME);
        if (optionsDialog.show() !== 1) return null;
        return { removeLineBreaks: removeLineBreaksCheckbox.value };
    }

    // =========================================
    // 変換 / Conversion
    // =========================================

    /**
     * エリア内文字があふれているかを返す
     * @param {TextFrame} areaTextFrame - 対象のエリア内文字
     * @returns {boolean} あふれていれば true
     */
    function isFrameOverset(areaTextFrame) {
        /* overflows を持たない環境がある / Some versions lack overflows */
        try {
            if (typeof areaTextFrame.overflows !== "undefined") return !!areaTextFrame.overflows;
        } catch (e) { }
        /* overflows が読めない環境では、行に入っている文字数と全文字数を突き合わせる
           Where overflows cannot be read, the characters inside the lines are counted against the total */
        try {
            var visibleCount = 0;
            for (var i = 0; i < areaTextFrame.lines.length; i++) {
                visibleCount += areaTextFrame.lines[i].characters.length;
            }
            return visibleCount < areaTextFrame.characters.length;
        } catch (e0) { return false; }
    }

    /**
     * 変換できるエリア内文字かを返す（ロック・非表示のものは対象外）
     * @param {PageItem} pageItem - 調べるオブジェクト
     * @returns {boolean} 変換できれば true
     */
    function isConvertibleAreaTextFrame(pageItem) {
        try {
            if (!pageItem || pageItem.typename !== "TextFrame" || pageItem.kind !== TextType.AREATEXT) return false;
            if (pageItem.locked || pageItem.hidden) return false;
            var ownerLayer = pageItem.layer;
            if (ownerLayer && (ownerLayer.locked || !ownerLayer.visible)) return false;
        } catch (e) { return false; }
        return true;
    }

    /**
     * 強制改行を1文字ずつ取り除く（contents の一括置換は文字ごとの書式を失うため使わない）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {void}
     */
    function removeForcedLineBreaks(textFrame) {
        var textCharacters;
        try { textCharacters = textFrame.characters; } catch (e) { return; }
        /* 削除するとインデックスがずれるので後ろから処理する / Deleting shifts the indices, so walk from the back */
        for (var i = textCharacters.length - 1; i >= 0; i--) {
            /* 読めない文字・消せない文字は飛ばす / Skip characters that cannot be read or removed */
            try {
                if (FORCED_BREAK_PATTERN.test(textCharacters[i].contents)) textCharacters[i].remove();
            } catch (eChar) { }
        }
    }

    /**
     * 選択からエリア内文字だけを拾う
     * @param {Array} selectedItems - ドキュメントの選択内容
     * @returns {TextFrame[]} 変換できるエリア内文字
     */
    function collectAreaTextFrames(selectedItems) {
        var areaTextFrames = [];
        for (var i = 0; i < selectedItems.length; i++) {
            if (isConvertibleAreaTextFrame(selectedItems[i])) { areaTextFrames.push(selectedItems[i]); }
        }
        return areaTextFrames;
    }

    /**
     * あふれているフレームだけ自動サイズ調整で解消してから、ポイント文字へ変換する
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrame[]} areaTextFrames - 変換するエリア内文字
     * @param {boolean} removeLineBreaks - 変換後に強制改行を取り除くか
     * @returns {{converted: TextFrame[], failedCount: number}} 変換できたポイント文字と、変換できなかった数
     */
    function convertAreaTextFramesToPointText(doc, areaTextFrames, removeLineBreaks) {
        /* あふれていないフレームに自動サイズ調整をかけると、テキストの配置が中央・下のときに文字が動くため
           Auto-sizing a frame with no overset would move the text when it is centred or bottom-aligned
           convertAreaObjectToPointObject() は「その場変換・戻り値 null」で、変換後は元の参照が stale になり
           kind を AREATEXT のまま報告することがある。そこで変換前に目印の名前を付け、変換後に
           doc.textFrames を1回だけ走査して回収する
           (the API converts in place and returns null, and the old wrappers can go stale and still report
            AREATEXT, so the frames are tagged with a marker name first and collected afterwards
            by a single fresh scan of doc.textFrames) */
        var markerPrefix = "__" + SCRIPT_NAME + "_marker_";
        var previousNames = [];

        for (var i = 0; i < areaTextFrames.length; i++) {
            var previousName = "";
            try { previousName = areaTextFrames[i].name; } catch (eRead) { }
            previousNames.push(previousName);
            try { areaTextFrames[i].name = markerPrefix + i; } catch (eTag) { }
        }

        /* 変換するとオブジェクトが置き換わるので後ろから処理する
           Each frame is replaced as it goes, so the list is walked from the back */
        for (var j = areaTextFrames.length - 1; j >= 0; j--) {
            /* 1フレームで失敗しても残りを処理できるようにする / One failing frame must not stop the rest */
            try {
                if (isFrameOverset(areaTextFrames[j])) {
                    doc.selection = null;
                    areaTextFrames[j].selected = true;
                    app.redraw(); /* Illustratorに選択状態を確定させる / Let Illustrator settle the selection */
                    runAutoSizeAction();
                }
                areaTextFrames[j].convertAreaObjectToPointObject();
            } catch (e) { }
        }

        /* 目印で回収して名前を戻す（走査は1回だけ）/ Collect by marker and restore the names (a single scan) */
        var convertedFrames = [], failedCount = 0;
        var allTextFrames = doc.textFrames;
        for (var k = 0; k < allTextFrames.length; k++) {
            var frameName = "";
            try { frameName = allTextFrames[k].name; } catch (eName) { continue; }
            if (frameName.indexOf(markerPrefix) !== 0) continue;

            var markerIndex = parseInt(frameName.substring(markerPrefix.length), 10);
            var restoredName = (!isNaN(markerIndex) && typeof previousNames[markerIndex] === "string")
                ? previousNames[markerIndex] : "";
            try { allTextFrames[k].name = restoredName; } catch (eRestore) { }
            if (allTextFrames[k].kind === TextType.POINTTEXT) {
                /* 変換すると折り返し位置に強制改行が残るので、ポイント文字になってから取り除く
                   The conversion leaves a forced break at every wrap, so they are stripped once it is point text */
                if (removeLineBreaks) removeForcedLineBreaks(allTextFrames[k]);
                convertedFrames.push(allTextFrames[k]);
            } else { failedCount++; }
        }
        return { converted: convertedFrames, failedCount: failedCount };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択を確かめて変換する（ダイナミックアクションは終了時に必ずアンロードする）
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var selectedItems = doc.selection;

        /* 文字ツールで文字を選択中は TextRange が返り、length が文字数になるため配列かどうかで判定する
           With the type tool the selection is a TextRange whose length counts characters */
        if (!(selectedItems instanceof Array) || selectedItems.length === 0) {
            alert(getLabel("alert.selectAreaText"));
            return;
        }

        var areaTextFrames = collectAreaTextFrames(selectedItems);
        if (!areaTextFrames.length) {
            alert(getLabel("alert.noTarget"));
            return;
        }

        /* 古いIllustratorには変換APIが無いので、何も触らずに知らせる / Older Illustrator has no conversion API, so nothing is touched */
        var supportsConversion = false;
        try { supportsConversion = !!areaTextFrames[0].convertAreaObjectToPointObject; } catch (eApi) { }
        if (!supportsConversion) {
            alert(getLabel("alert.notSupported"));
            return;
        }

        var dialogOptions = showOptionsDialog();
        if (!dialogOptions) return;

        /* 読み込めないと、あふれたフレームを広げずに変換して文字が欠けるので中止する
           Without the set, overset frames would be converted unexpanded and lose text, so stop */
        if (!loadAutoSizeAction()) {
            alert(getLabel("alert.actionFailed"));
            return;
        }
        var conversionResult;
        try {
            conversionResult = convertAreaTextFramesToPointText(doc, areaTextFrames, dialogOptions.removeLineBreaks);
        } finally {
            unloadAutoSizeAction();
        }

        /* 変換後のポイント文字を選び直す（削除済みの参照を選択に残さない）
           Re-select the resulting point text, so no stale reference lingers in the selection */
        try { doc.selection = conversionResult.converted.length ? conversionResult.converted : null; } catch (eSelect) { }
        app.redraw();

        if (!conversionResult.converted.length) {
            alert(getLabel("alert.noTarget"));
        } else if (conversionResult.failedCount > 0) {
            alert(getLabel("alert.partialFailure")
                .replace("{done}", conversionResult.converted.length)
                .replace("{failed}", conversionResult.failedCount));
        }
    }

    main();

})();
