#target illustrator
#targetengine "SmartLayerManageEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

オブジェクトを指定したレイヤーへ一括で移動します。
対象は選択オブジェクト、全テキスト、すべて、すべて（強制）から選べます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartLayerManage.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n95ec4929ae9d

### Overview

Moves objects to a chosen layer in bulk.
The scope can be the selection, all text, everything, or everything (forced).

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartLayerManage.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartLayerManage";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.13";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-03";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartLayerManage.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartLayerManage.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n95ec4929ae9d"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var DEFAULT_NEW_LAYER_NAME = "New Layer";    /* 新規レイヤー名の初期値 / default name for a new layer */
    var DEFAULT_DELETE_EMPTY = true;             /* 空レイヤーを削除するかの初期値 / delete empty layers by default */
    var PROTECTED_LAYER_NAME = "bg";             /* 空でも削除しないレイヤー名 / layer name kept even when empty */
    var PROTECTED_LAYER_PREFIX = "//";           /* この接頭辞のレイヤーは空でも削除しない / prefix of layers kept even when empty */
    var TARGET_LAYER_COLOR = [79, 128, 255];     /* 移動先レイヤーのカラー（RGB）/ layer color given to the destination */

    // =========================================
    // レイアウト / Layout
    // =========================================
    var NEW_LAYER_FIELD_CHARS = 12;              /* 新規レイヤー名欄の文字数 / width of the new layer name field */

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

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "オブジェクトをレイヤーへ移動", en: "Move Objects to Layer" }
        },
        panel: {
            scope: { ja: "対象オブジェクト", en: "Target Objects" },
            targetLayer: { ja: "移動先レイヤー", en: "Target Layer" }
        },
        radio: {
            selected: { ja: "選択中", en: "Selected Objects" },
            allText: { ja: "全テキスト", en: "All Text" },
            allObjects: { ja: "すべて", en: "All Objects" },
            allForce: { ja: "すべて（強制）", en: "All (Force)" },
            newLayer: { ja: "新規", en: "New" }
        },
        checkbox: {
            deleteEmpty: { ja: "空レイヤーを削除", en: "Delete Empty Layers" }
        },
        tooltip: {
            selected: { ja: "選択しているオブジェクトだけを移動します。", en: "Moves only the selected objects." },
            allText: { ja: "ドキュメント内のテキストをすべて移動します。", en: "Moves every text object in the document." },
            allObjects: {
                ja: "ロックと非表示をすべて解除してから、すべてのオブジェクトを移動します。",
                en: "Unlocks and shows everything, then moves every object."
            },
            allForce: {
                ja: "［すべてのレイヤーを結合］を実行し、すべてを一番上のレイヤーにまとめます。",
                en: "Runs Flatten Artwork, merging everything into the top layer."
            },
            deleteEmpty: {
                ja: "移動したあと、中身が無くなったレイヤーを削除します（「bg」と「//」で始まるレイヤーは残します）。",
                en: "Deletes the layers left empty after the move (layers named \"bg\" or starting with \"//\" are kept)."
            },
            newLayer: {
                ja: "新しいレイヤーを作って、そこへ移動します。名前は右の欄で決めます。",
                en: "Creates a new layer and moves the objects there. The field on the right names it."
            },
            layerRadio: { ja: "このレイヤーへ移動します。", en: "Moves the objects to this layer." }
        },
        button: {
            move: { ja: "移動", en: "Move" },
            close: { ja: "閉じる", en: "Close" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noLayerName: { ja: "新しいレイヤー名を入力してください。", en: "Enter a name for the new layer." },
            noLayerSelected: { ja: "移動先レイヤーを選択してください。", en: "Please select a target layer." },
            noSelection: { ja: "オブジェクトが選択されていません。", en: "No objects selected." },
            error: { ja: "エラーが発生しました: ", en: "An error occurred: " }
        }
    };

    // =========================================
    // 一時アクション / Temporary action
    // =========================================

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

    /**
     * 一時アクションで［すべてのレイヤーを結合］を実行する
     * @returns {void}
     */
    function flattenArtwork() {
        var actionSetName = "layer";
        var actionName = "flattenLayers";

        /* アクション定義 / Action definition */
        var actionCode = [
            "/version 3 /name [ 5 6c61796572 ] /isOpen 1 /actionCount 1 /action-1 { /name [ 13 666c617474656e4c6179657273 ] /keyIndex 0 /colorIndex 0 /isOpen 1 /eventCount 1 /event-1 { /useRulersIn1stQuadrant 0 /internalName (ai_plugin_Layer) /localizedName [ 9 e8a1a8e7a4ba203a20 ] /isOpen 0 /isOn 1 /hasDialog 0 /parameterCount 2 /parameter-1 { /key 1836411236 /showInPalette 4294967295 /type (integer) /value 14 } /parameter-2 { /key 1851878757 /showInPalette 4294967295 /type (ustring) /value [ 33 e38199e381b9e381a6e381aee383ace382a4e383a4e383bce38292e7b590e590 88 ] } } }"
        ].join("\n");

        if (!runTemporaryAction(actionCode, actionSetName, actionName)) {
            alert(getLabel("alert.error") + actionSetName + " / " + actionName);
        }
    }

    // =========================================
    // 収集・移動 / Collect and move
    // =========================================

    /**
     * コンテナ（レイヤー／グループ）の pageItems を再帰的に集める（グループの中身、レイヤーならサブレイヤーも）
     * @param {Layer|GroupItem} container - 対象のレイヤーまたはグループ
     * @param {PageItem[]} collectedItems - 追加先の配列
     * @returns {void}
     */
    function collectItemsRecursive(container, collectedItems) {
        for (var i = 0; i < container.pageItems.length; i++) {
            var item = container.pageItems[i];
            collectedItems.push(item);
            if (item.typename === "GroupItem") {
                collectItemsRecursive(item, collectedItems);
            }
        }
        if (container.typename === "Layer") {
            for (var j = 0; j < container.layers.length; j++) {
                collectItemsRecursive(container.layers[j], collectedItems);
            }
        }
    }

    /**
     * レイヤーとサブレイヤーからテキストフレームを再帰的に集める
     * @param {Layer} layer - 対象レイヤー
     * @param {TextFrame[]} collectedItems - 追加先の配列
     * @returns {void}
     */
    function collectTextFramesRecursive(layer, collectedItems) {
        for (var i = 0; i < layer.textFrames.length; i++) {
            collectedItems.push(layer.textFrames[i]);
        }
        for (var j = 0; j < layer.layers.length; j++) {
            collectTextFramesRecursive(layer.layers[j], collectedItems);
        }
    }

    /**
     * アイテムのロック・非表示を解除して移動先レイヤーへ移す
     * @param {PageItem} item - 移動するアイテム
     * @param {Layer} targetLayer - 移動先レイヤー
     * @returns {void}
     */
    function prepareAndMoveItem(item, targetLayer) {
        try {
            item.locked = false;
            item.hidden = false;
            item.move(targetLayer, ElementPlacement.PLACEATEND);
        } catch (e) {
            /* 移動済みの親と一緒に動いた子や、ロックされたレイヤー内のものは飛ばす / Skip items that cannot be moved */
        }
    }

    /**
     * 空のレイヤーを再帰的に削除する（保護するレイヤー名・接頭辞は残す）
     * @param {Layer} layer - 対象レイヤー
     * @returns {void}
     */
    function deleteEmptyLayersRecursive(layer) {
        for (var i = layer.layers.length - 1; i >= 0; i--) {
            deleteEmptyLayersRecursive(layer.layers[i]);
        }
        if (layer.pageItems.length === 0 && layer.layers.length === 0 &&
            layer.name !== PROTECTED_LAYER_NAME && layer.name.indexOf(PROTECTED_LAYER_PREFIX) !== 0) {
            layer.remove();
        }
    }

    /**
     * レイヤーカラーを RGB で変更する（表示中でロックされていないときだけ）
     * @param {Layer} layer - 対象レイヤー
     * @param {number[]} rgbValues - [r, g, b]
     * @returns {void}
     */
    function setLayerColor(layer, rgbValues) {
        if (!layer.visible || layer.locked) return;
        var newColor = new RGBColor();
        newColor.red = rgbValues[0];
        newColor.green = rgbValues[1];
        newColor.blue = rgbValues[2];
        layer.color = newColor;
    }

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
    // ダイアログ / Dialog
    // =========================================

    /**
     * ラジオボタンを追加して helpTip を付ける
     * @param {Panel|Group} parent - 追加先
     * @param {string} labelKey - radio / tooltip 共通のキー
     * @returns {RadioButton} 追加したラジオボタン
     */
    function addScopeRadio(parent, labelKey) {
        var radio = parent.add("radiobutton", undefined, getLabel("radio." + labelKey));
        radio.helpTip = getLabel("tooltip." + labelKey);
        return radio;
    }

    /**
     * ダイアログを組み立てる
     * @param {Document} doc - 対象ドキュメント
     * @returns {Object} ダイアログと主要なコントロール
     */
    function buildDialog(doc) {
        var layers = doc.layers;
        var hasSelection = doc.selection && doc.selection.length > 0;

        var dialog = new Window("dialog", getLabel("dialog.title"));
        setupWindow(dialog);

        /* 左右の列を横に並べる / Lay out the left and right columns side by side */
        var columnsGroup = dialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.spacing = COLUMN_SPACING;

        /* 左列：対象オブジェクト / Left column: scope */
        var leftColumn = columnsGroup.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = ["left", "top"];

        var scopePanel = leftColumn.add("panel", undefined, getLabel("panel.scope"));
        setupPanel(scopePanel, 6);

        var controls = { dialog: dialog };
        controls.radioSelected = addScopeRadio(scopePanel, "selected");
        controls.radioAllText = addScopeRadio(scopePanel, "allText");
        controls.radioAllObjects = addScopeRadio(scopePanel, "allObjects");
        controls.radioAllForce = addScopeRadio(scopePanel, "allForce");

        /* 初期値は選択の有無で決める / Default depends on whether anything is selected */
        controls.radioSelected.value = hasSelection;
        controls.radioAllObjects.value = !hasSelection;

        var deleteEmptyRow = leftColumn.add("group");
        deleteEmptyRow.orientation = "row";
        deleteEmptyRow.alignChildren = ["center", "center"];
        deleteEmptyRow.margins = [5, 5, 0, 0];

        controls.deleteEmptyCheckbox = deleteEmptyRow.add("checkbox", undefined, getLabel("checkbox.deleteEmpty"));
        controls.deleteEmptyCheckbox.helpTip = getLabel("tooltip.deleteEmpty");
        controls.deleteEmptyCheckbox.value = DEFAULT_DELETE_EMPTY;

        /* 右列：移動先レイヤー / Right column: destination layer */
        var rightColumn = columnsGroup.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = ["fill", "top"];

        var layerPanel = rightColumn.add("panel", undefined, getLabel("panel.targetLayer"));
        setupPanel(layerPanel, 6);

        /* 先頭に「新規」ラジオと名前欄 / "New" radio and name field first */
        var newLayerRow = layerPanel.add("group");
        newLayerRow.orientation = "row";
        newLayerRow.alignChildren = ["left", "center"]; /* パネル幅に広がっても左に寄せる / stay left when the row fills the panel */
        controls.radioNewLayer = addScopeRadio(newLayerRow, "newLayer");
        controls.newLayerNameInput = newLayerRow.add("edittext", undefined, DEFAULT_NEW_LAYER_NAME);
        controls.newLayerNameInput.helpTip = getLabel("tooltip.newLayer");
        controls.newLayerNameInput.characters = NEW_LAYER_FIELD_CHARS;

        /* 既存レイヤー（ロック中は選べない）/ Existing layers; locked ones are disabled */
        controls.layerRadios = [controls.radioNewLayer];
        var firstAvailableIndex = -1;
        for (var i = 0; i < layers.length; i++) {
            var layerRadio = layerPanel.add("radiobutton", undefined, layers[i].name);
            layerRadio.helpTip = getLabel("tooltip.layerRadio");
            controls.layerRadios.push(layerRadio);
            if (layers[i].locked) {
                layerRadio.enabled = false;
            } else if (firstAvailableIndex === -1) {
                firstAvailableIndex = i + 1; /* 「新規」の分だけずらす / Offset by the "New" radio */
            }
        }
        if (firstAvailableIndex !== -1) {
            controls.layerRadios[firstAvailableIndex].value = true;
        }

        bindLayerRadios(controls);

        /* 下段：ボタン行 / Bottom: button row */
        var buttonRow = addButtonRow(dialog);
        controls.btnClose = buttonRow.rightGroup.add("button", undefined, getLabel("button.close"));
        controls.btnMove = buttonRow.rightGroup.add("button", undefined, getLabel("button.move"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        return controls;
    }

    /**
     * 移動先ラジオの排他と名前欄の有効状態をつなぐ（「新規」は別グループにあるため手動で排他する）
     * @param {Object} controls - buildDialog() のコントロール
     * @returns {void}
     */
    function bindLayerRadios(controls) {
        var layerRadios = controls.layerRadios;

        function selectLayerRadio(selectedIndex) {
            for (var i = 0; i < layerRadios.length; i++) {
                layerRadios[i].value = (i === selectedIndex);
            }
            controls.newLayerNameInput.enabled = controls.radioNewLayer.value;
        }

        for (var i = 0; i < layerRadios.length; i++) {
            (function (radioIndex) {
                layerRadios[radioIndex].onClick = function () {
                    selectLayerRadio(radioIndex);
                };
            })(i);
        }
        controls.newLayerNameInput.enabled = controls.radioNewLayer.value;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログの指定から移動先レイヤーを決める（「新規」なら一番上に作る）
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} controls - buildDialog() のコントロール
     * @returns {Layer|null} 移動先レイヤー。指定が足りなければ null
     */
    function resolveTargetLayer(doc, controls) {
        if (controls.radioNewLayer.value) {
            var newLayerName = controls.newLayerNameInput.text;
            if (!newLayerName) {
                alert(getLabel("alert.noLayerName"));
                return null;
            }
            var newLayer = doc.layers.add();
            newLayer.name = newLayerName;
            newLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
            return newLayer;
        }
        for (var i = 1; i < controls.layerRadios.length; i++) {
            if (controls.layerRadios[i].value) {
                return doc.layers.getByName(controls.layerRadios[i].text);
            }
        }
        alert(getLabel("alert.noLayerSelected"));
        return null;
    }

    /**
     * 対象の指定に応じて移動するアイテムを集める（「すべて（強制）」はレイヤーを結合して空を返す）
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} controls - buildDialog() のコントロール
     * @returns {PageItem[]|null} 移動するアイテム。選択が無ければ null
     */
    function collectItemsToMove(doc, controls) {
        var itemsToMove = [];
        if (controls.radioSelected.value) {
            var currentSelection = doc.selection;
            if (!currentSelection || currentSelection.length === 0) {
                alert(getLabel("alert.noSelection"));
                return null;
            }
            return currentSelection;
        }
        if (controls.radioAllText.value) {
            for (var i = 0; i < doc.layers.length; i++) {
                collectTextFramesRecursive(doc.layers[i], itemsToMove);
            }
        } else if (controls.radioAllForce.value) {
            flattenArtwork();
        } else {
            app.executeMenuCommand("unlockAll");
            app.executeMenuCommand("showAll");
            for (var j = 0; j < doc.layers.length; j++) {
                collectItemsRecursive(doc.layers[j], itemsToMove);
            }
        }
        return itemsToMove;
    }

    /**
     * ［移動］の処理：移動先を決めて移し、空レイヤーを削除して移動先のカラーを変える
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} controls - buildDialog() のコントロール
     * @returns {void}
     */
    function runMove(doc, controls) {
        var targetLayer = resolveTargetLayer(doc, controls);
        if (!targetLayer) return;

        var itemsToMove = collectItemsToMove(doc, controls);
        if (!itemsToMove) return;

        for (var i = 0; i < itemsToMove.length; i++) {
            prepareAndMoveItem(itemsToMove[i], targetLayer);
        }

        if (controls.deleteEmptyCheckbox.value) {
            for (var j = doc.layers.length - 1; j >= 0; j--) {
                deleteEmptyLayersRecursive(doc.layers[j]);
            }
        }

        setLayerColor(targetLayer, TARGET_LAYER_COLOR);
        controls.dialog.close();
    }

    /**
     * メイン処理
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var controls = buildDialog(doc);

        controls.btnMove.onClick = function () {
            runMove(doc, controls);
        };
        controls.btnClose.onClick = function () {
            controls.dialog.close();
        };

        prepareDialogWindow(controls.dialog, SCRIPT_NAME);
        controls.dialog.show();
    }

    main();

})();
