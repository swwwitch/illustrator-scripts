#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択に応じて、ポイント文字・パス上文字とエリア内文字を相互に変換します。
順変換は見た目を保ったままエリア内文字へ、逆変換は枠サイズを保ったままポイント文字へ変換します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ConvertAreaAndPointType.md

### Overview

Converts between point text or text on a path and area text, depending on what is selected.
The forward conversion preserves the appearance, and the reverse keeps the frame size when returning to point text.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ConvertAreaAndPointType.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ConvertAreaAndPointType";      /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-07-02";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ConvertAreaAndPointType.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ConvertAreaAndPointType.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_MARGINS = 15;                /* ダイアログの余白 / Dialog margins */
    var PANEL_MARGINS  = [16, 20, 16, 12];  /* パネル余白 [左,上,右,下] / Panel margins */
    var PANEL_SPACING  = 8;                 /* パネル内の要素間隔 / Panel spacing */

    /**
     * パネルの共通設定を適用する
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

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * Illustrator の UI 言語から表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILanguage();

    /* 日英ラベル定義（UI パーツ別）/ Japanese-English label definitions (by UI part) */
    var LABELS = {
        dialog: {
            title: { ja: "テキストの変換", en: "Convert Text" }
        },
        panel: {
            selection: { ja: "選択オブジェクト", en: "Selection" },
            style: { ja: "スタイル", en: "Style" }
        },
        radio: {
            keepStyle: { ja: "保持する", en: "Keep" },
            dontKeepStyle: { ja: "保持しない", en: "Don't keep" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        /* 順/逆で意味が異なるため補足 / Meaning differs per direction */
        tooltip: {
            keepForward: {
                ja: "元テキストの見た目（塗り・線・効果）を引き継いでエリア内文字にします。",
                en: "Carry over the source text's appearance (fill/stroke/effects) into the area type."
            },
            dontKeepForward: {
                ja: "見た目を引き継がず、文字だけをエリア内文字にします。",
                en: "Transfer only the text (no appearance) into the area type."
            },
            keepReverse: {
                ja: "変換後の見た目のまま。追加処理はしません。",
                en: "Leave the converted look as-is; no extra processing."
            },
            dontKeepReverse: {
                ja: "塗りを2枚追加し、元の枠サイズの長方形背景（ボタン風）を付与します。",
                en: "Add two fills + a rectangle background at the original frame size (button-like)."
            }
        },
        /* 選択の内訳に出すテキスト種別 / Text-type names for the selection summary */
        textType: {
            pointText: { ja: "ポイント文字", en: "Point text" },
            areaText: { ja: "エリア内文字", en: "Area type" },
            pathText: { ja: "パス上文字", en: "Path text" },
            other: { ja: "テキスト以外", en: "Non-text" }
        },
        alert: {
            selectText: {
                ja: "ポイント文字・パス上文字、またはエリア内文字を選択してください。",
                en: "Please select point text, path text, or area type."
            },
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            conversionFailed: { ja: "エリア内文字に変換できませんでした。", en: "Could not convert to area type." },
            reverseFailed: { ja: "ポイント文字に変換できませんでした。", en: "Could not convert to point text." }
        }
    };

    /**
     * ドット区切りのパスで LABELS から表示言語の文字列を引く
     * @param {string} labelPath - "alert.noDocument" のようなドット区切りのキー
     * @returns {string} 表示言語の文字列（無ければ英語、それも無ければ labelPath）
     */
    function getLabel(labelPath) {
        var pathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            if (!labelNode || typeof labelNode[pathKeys[i]] === "undefined") return labelPath;
            labelNode = labelNode[pathKeys[i]];
        }
        if (labelNode[uiLang]) return labelNode[uiLang];
        if (labelNode.en) return labelNode.en;
        return labelPath;
    }

    // =========================================
    // テキストの種類 / Text kinds
    // =========================================

    /**
     * テキストフレームの種類を返す
     * @param {PageItem} pageItem - 調べるオブジェクト
     * @returns {TextType|null} テキストの種類。テキストフレーム以外は null
     */
    function getTextKind(pageItem) {
        return (pageItem.typename === "TextFrame") ? pageItem.kind : null;
    }

    /**
     * ポイント文字またはパス上文字かどうか
     * @param {PageItem} pageItem - 調べるオブジェクト
     * @returns {boolean} ポイント文字・パス上文字なら true
     */
    function isPointOrPathText(pageItem) {
        var textKind = getTextKind(pageItem);
        return textKind === TextType.POINTTEXT || textKind === TextType.PATHTEXT;
    }

    // =========================================
    // パス上文字 → ポイント文字（変換前処理）/ Path text → Point text (pre-process)
    // =========================================

    /**
     * 関数を実行し、例外は握りつぶす
     * @param {Function} fn - 実行する関数
     * @returns {*} 関数の戻り値。失敗したら undefined
     */
    function tryQuietly(fn) { try { return fn(); } catch (e) { return undefined; } }

    /**
     * パス上文字の文字ごとの属性を退避する
     * @param {TextFrame} textFrame - パス上文字
     * @returns {Object[]} 文字ごとの属性（font / size / fillColor / strokeColor / strokeWeight / autoLeading / leading）
     */
    function snapshotCharAttrs(textFrame) {
        var charAttrs = [];
        for (var i = 0; i < textFrame.characters.length; i++) {
            var charAttr = textFrame.characters[i].characterAttributes;
            charAttrs.push({
                font: charAttr.textFont,
                size: charAttr.size,
                fillColor: charAttr.fillColor,
                strokeColor: charAttr.strokeColor,
                strokeWeight: charAttr.strokeWeight,
                autoLeading: charAttr.autoLeading,
                leading: charAttr.leading
            });
        }
        return charAttrs;
    }

    /**
     * 退避した文字属性を新しいポイント文字へ戻す（ベースラインシフト・比率は既定に戻す）
     * @param {TextFrame} pointText - 新しいポイント文字
     * @param {Object[]} charAttrs - snapshotCharAttrs() の結果
     * @returns {void}
     */
    function restoreCharAttrs(pointText, charAttrs) {
        var restoreCount = Math.min(pointText.characters.length, charAttrs.length);
        for (var i = 0; i < restoreCount; i++) {
            var targetAttr = pointText.characters[i].characterAttributes;
            var sourceAttr = charAttrs[i];

            tryQuietly(function () { targetAttr.textFont = sourceAttr.font; });
            tryQuietly(function () { targetAttr.size = sourceAttr.size; });
            tryQuietly(function () { targetAttr.fillColor = sourceAttr.fillColor; });
            tryQuietly(function () {
                var sourceStroke = sourceAttr.strokeColor;
                targetAttr.strokeColor = sourceStroke;
                targetAttr.strokeWeight = (sourceStroke && sourceStroke.typename === "NoColor") ? 0 : sourceAttr.strokeWeight;
            });
            tryQuietly(function () { targetAttr.baselineShift = 0; });
            tryQuietly(function () { targetAttr.horizontalScale = 100; });
            tryQuietly(function () { targetAttr.verticalScale = 100; });
            tryQuietly(function () { targetAttr.autoLeading = sourceAttr.autoLeading; });
            if (!sourceAttr.autoLeading) tryQuietly(function () { targetAttr.leading = sourceAttr.leading; });
        }
    }

    /**
     * パス上文字と同じ内容・行揃え・文字属性のポイント文字を、パスの始点に作る
     * @param {Document} doc - 対象のドキュメント
     * @param {TextFrame} pathText - 元のパス上文字
     * @returns {TextFrame|null} 作ったポイント文字。パスを取れなければ null
     */
    function createPointTextFromPathText(doc, pathText) {
        var originalPath = null;
        tryQuietly(function () { originalPath = pathText.textPath; });
        if (!originalPath) return null;

        /* 1) 文字属性・内容・行揃えを退避 / Snapshot char attributes, contents, justification */
        var charAttrs = snapshotCharAttrs(pathText);

        var textContents = "";
        tryQuietly(function () { textContents = pathText.contents; });

        var justification = null;
        tryQuietly(function () {
            if (pathText.paragraphs && pathText.paragraphs.length > 0) {
                justification = pathText.paragraphs[0].paragraphAttributes.justification;
            }
        });

        /* 2) パス始点にポイント文字を新規作成 / Create new point text at path start anchor */
        var pointText = doc.textFrames.add();
        var anchorPoint = null;
        tryQuietly(function () {
            if (originalPath.pathPoints && originalPath.pathPoints.length > 0) {
                anchorPoint = originalPath.pathPoints[0].anchor;
            }
        });
        if (anchorPoint) {
            pointText.position = [anchorPoint[0], anchorPoint[1]];
        }

        pointText.contents = textContents;

        if (justification !== null && pointText.paragraphs && pointText.paragraphs.length > 0) {
            tryQuietly(function () { pointText.paragraphs[0].paragraphAttributes.justification = justification; });
        }

        /* 既定の線を一旦消し、後で文字ごとに復元 / Clear default stroke, restore per-character later */
        tryQuietly(function () {
            var noColor = new NoColor();
            pointText.textRange.characterAttributes.strokeColor = noColor;
            pointText.textRange.characterAttributes.strokeWeight = 0;
        });

        /* 3) 文字属性を復元 / Restore per-character attributes */
        restoreCharAttrs(pointText, charAttrs);
        return pointText;
    }

    /**
     * パス上文字を字形を保ったままポイント文字へ分離する（元のパス上文字は削除）
     * @param {Document} doc - 対象のドキュメント
     * @param {TextFrame[]} pathTextFrames - パス上文字
     * @returns {TextFrame[]} 作ったポイント文字（選択状態）
     */
    function detachPathTextToPointText(doc, pathTextFrames) {
        var createdTexts = [];
        if (!doc || !pathTextFrames || !pathTextFrames.length) return createdTexts;

        /* 新規テキストだけ選べるよう選択を解除 / Clear selection */
        tryQuietly(function () { doc.selection = null; });

        for (var i = pathTextFrames.length - 1; i >= 0; i--) {
            var pathText = pathTextFrames[i];
            if (!pathText || pathText.typename !== "TextFrame" || pathText.kind !== TextType.PATHTEXT) continue;

            var pointText = createPointTextFromPathText(doc, pathText);
            if (!pointText) continue;

            /* 4) 元のパス上文字を削除（パスも一緒に消える）/ Remove original path text */
            tryQuietly(function () { pathText.remove(); });

            /* 5) 新規テキストを選択して返す / Select and return new text */
            tryQuietly(function () { pointText.selected = true; });
            createdTexts.push(pointText);
        }

        return createdTexts;
    }

    /**
     * 選択内のパス上文字をポイント文字へ置き換え、選択し直す
     * @param {Document} doc - 対象のドキュメント
     * @param {PageItem[]} currentSelection - 現在の選択
     * @returns {PageItem[]} 置き換えた選択（パス上文字が無ければ元の選択）
     */
    function preprocessPathTextSelection(doc, currentSelection) {
        if (!doc || !currentSelection || !currentSelection.length) return currentSelection;

        var pathTexts = [];
        for (var i = 0; i < currentSelection.length; i++) {
            var selectedItem = currentSelection[i];
            try {
                if (selectedItem && selectedItem.typename === "TextFrame" && selectedItem.kind === TextType.PATHTEXT) {
                    pathTexts.push(selectedItem);
                }
            } catch (e0) {
                /* 無効オブジェクト（削除済み等）はスキップ / Skip invalid objects */
            }
        }
        if (!pathTexts.length) return currentSelection;

        var newTexts = detachPathTextToPointText(doc, pathTexts);
        if (!newTexts.length) return currentSelection;

        /* パス上文字を新ポイント文字に差し替えた新しい選択配列を構築。削除した旧オブジェクトは参照で例外になるので飛ばす
           Build the replaced selection; the removed originals throw on access and are skipped */
        var replacedSelection = [];
        for (var j = 0; j < currentSelection.length; j++) {
            var keptItem = currentSelection[j];
            try {
                if (keptItem && !(keptItem.typename === "TextFrame" && keptItem.kind === TextType.PATHTEXT)) {
                    replacedSelection.push(keptItem);
                }
            } catch (e1) {
                /* 無効オブジェクトはスキップ / Skip invalid objects */
            }
        }
        for (var k = 0; k < newTexts.length; k++) replacedSelection.push(newTexts[k]);

        try { doc.selection = replacedSelection; app.redraw(); } catch (e) { }

        return replacedSelection;
    }

    // =========================================
    // ダイナミックアクション / Dynamic actions
    // =========================================

    /**
     * 文字列を16進数表現にする
     * @param {string} text - ASCII 文字列
     * @returns {string} 16進数の文字列
     */
    function asciiToHex(text) {
        var hexText = "";
        for (var i = 0; i < text.length; i++) {
            var hexPair = text.charCodeAt(i).toString(16);
            if (hexPair.length < 2) hexPair = "0" + hexPair;
            hexText += hexPair;
        }
        return hexText;
    }

    /**
     * アクション定義の /name ブロック（ASCII）を作る
     * @param {string} actionName - アクション名またはセット名
     * @returns {string} /name ブロック
     */
    function buildActionNameBlock(actionName) {
        return "/name [ " + actionName.length + " " + asciiToHex(actionName).toUpperCase() + " ]";
    }

    /**
     * アクションセット定義（.aia）の文字列を組み立てる
     * @param {string} setName - アクションセット名
     * @param {string} internalName - イベントの内部名
     * @param {string} localizedNameHex - ローカライズ名（"長さ 16進" 形式。空なら省略）
     * @param {number} parameterKey - パラメーターのキー
     * @param {Object[]} actionDefinitions - { name, value } のアクション定義
     * @returns {string} .aia の文字列
     */
    function buildActionSetAia(setName, internalName, localizedNameHex, parameterKey, actionDefinitions) {
        var aiaString = "/version 3" +
            buildActionNameBlock(setName) +
            "/isOpen 1" +
            "/actionCount " + actionDefinitions.length;

        for (var i = 0; i < actionDefinitions.length; i++) {
            var actionDef = actionDefinitions[i];
            aiaString += "/action-" + (i + 1) + " {" +
                " " + buildActionNameBlock(actionDef.name) +
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
                " /key " + parameterKey +
                " /showInPalette 4294967295" +
                " /type (integer)" +
                " /value " + actionDef.value +
                " }" +
                " }" +
                "}";
        }
        return aiaString;
    }

    /* フレーム整列アクションセット名 / Frame-alignment action set name */
    var AREA_TEXT_ACTION_SET = "AreaText";

    /**
     * アクションセットを読み込む（一時フォルダーに .aia を書き出して loadAction）
     * 同名のセットがあれば先に外して衝突を避ける
     * @param {string} setName - アクションセット名
     * @param {string} aiaString - .aia の文字列
     * @returns {void}
     */
    function loadActionSet(setName, aiaString) {
        unloadActionSet(setName);
        try {
            var actionFile = new File(Folder.temp + "/AreaTypeToolkit_action_" + setName + ".aia");
            actionFile.open("w");
            actionFile.write(aiaString);
            actionFile.close();

            app.loadAction(actionFile);
            actionFile.remove();
        } catch (e) { }
    }

    /**
     * アクションセットを破棄する
     * @param {string} setName - アクションセット名
     * @returns {void}
     */
    function unloadActionSet(setName) {
        try { app.unloadAction(setName, ""); } catch (e) { }
    }

    /**
     * フレーム整列（縦中央）アクションを読み込む
     * @returns {void}
     */
    function loadFrameAlignmentAction() {
        var aiaString = buildActionSetAia(
            AREA_TEXT_ACTION_SET,
            "adobe_frameAlignment",
            "39 e382a8e383aae382a2e58685e69687e5ad97e381aee38395e383ace383bce383a0e695b4e58897",
            1717660782,
            [{ name: "AlignCenter", value: 1 }]
        );
        loadActionSet(AREA_TEXT_ACTION_SET, aiaString);
    }

    /**
     * フレーム整列アクションを破棄する
     * @returns {void}
     */
    function unloadFrameAlignmentAction() {
        unloadActionSet(AREA_TEXT_ACTION_SET);
    }

    /**
     * エリア内文字を縦方向中央に配置する（DOM では不安定なためアクション再生）
     * @param {TextFrame} areaType - 対象のエリア内文字
     * @returns {void}
     */
    function applyVerticalCenter(areaType) {
        try {
            var doc = app.activeDocument;
            doc.selection = null;
            doc.selection = [areaType];
            app.redraw(); /* 選択状態を確定 / Commit the selection */
            app.doScript("AlignCenter", AREA_TEXT_ACTION_SET, false);
        } catch (e) { }
    }

    // =========================================
    // グラフィックスタイル登録 / Graphic style registration
    // =========================================

    /* 一時グラフィックスタイル名とアクション / Temp graphic style name and action */
    var TEMP_STYLE_NAME = "temp_style";
    var TEMP_STYLE_ACTION_SET = "GraphicStyle";
    var TEMP_STYLE_ACTION_NAME = "AddNewWithoutName";

    /* 選択オブジェクトの見た目を無名グラフィックスタイルとして末尾に追加するアクション定義(.aia)
       Dynamic action (.aia) that appends the selection's appearance as an unnamed graphic style */
    var GRAPHIC_STYLE_AIA = '/version 3 /name [ 12 477261706869635374796c65 ] /isOpen 1 /actionCount 1 /action-1 { /name [ 17 4164644e6577576974686f75744e616d65 ] /keyIndex 0 /colorIndex 0 /isOpen 1 /eventCount 1 /event-1 { /useRulersIn1stQuadrant 0 /internalName (ai_plugin_styles) /localizedName [ 30 e382b0e383a9e38395e382a3e38383e382afe382b9e382bfe382a4e383ab ] /isOpen 1 /isOn 1 /hasDialog 1 /showDialog 0 /parameterCount 1 /parameter-1 { /key 1835363957 /showInPalette 4294967295 /type (enumerated) /name [ 36 e696b0e8a68fe382b0e383a9e38395e382a3e38383e382afe382b9e382bfe382a4e383ab ] /value 1 } } }';

    /**
     * テキストフレームの見た目を temp_style として登録する（既存があれば作り直す）
     * @param {TextFrame} textFrame - 見た目の元になるテキストフレーム
     * @returns {string|null} 登録したスタイル名。登録できなければ null
     */
    function registerTextFrameAsTempGraphicStyle(textFrame) {
        if (!textFrame) return null;
        var doc = app.activeDocument;
        var graphicStyles = doc.graphicStyles;

        /* 既存の temp_style を削除 / Remove any existing temp_style */
        removeGraphicStyleByName(TEMP_STYLE_NAME);

        /* 登録用に対象だけを選択 / Select only the target for registration */
        doc.selection = null;
        try { textFrame.selected = true; } catch (selectError) { return null; }

        var countBefore = graphicStyles.length;
        loadActionSet(TEMP_STYLE_ACTION_SET, GRAPHIC_STYLE_AIA);
        try { app.doScript(TEMP_STYLE_ACTION_NAME, TEMP_STYLE_ACTION_SET, false); } catch (runError) { }
        unloadActionSet(TEMP_STYLE_ACTION_SET);

        /* 末尾に増えたスタイルを temp_style に改名 / Rename the newly appended style to temp_style */
        if (graphicStyles.length <= countBefore) return null;
        graphicStyles[graphicStyles.length - 1].name = TEMP_STYLE_NAME;
        return TEMP_STYLE_NAME;
    }

    /**
     * 名前でグラフィックスタイルを適用する
     * @param {string} styleName - グラフィックスタイル名
     * @param {PageItem} targetItem - 適用先
     * @returns {boolean} 適用できたら true
     */
    function applyGraphicStyleByName(styleName, targetItem) {
        if (!styleName || !targetItem) return false;
        try {
            app.activeDocument.graphicStyles.getByName(styleName).applyTo(targetItem);
            return true;
        } catch (e) { return false; }
    }

    /**
     * 名前でグラフィックスタイルを削除する（無ければ何もしない）
     * @param {string} styleName - グラフィックスタイル名
     * @returns {void}
     */
    function removeGraphicStyleByName(styleName) {
        if (!styleName) return;
        try { app.activeDocument.graphicStyles.getByName(styleName).remove(); } catch (e) { }
    }

    // =========================================
    // 正確なサイズ計測 / Accurate size measurement
    // =========================================

    /**
     * 選択オブジェクト群の可視バウンディングボックスの和を返す
     * @param {PageItem[]} targetItems - 対象のオブジェクト
     * @returns {number[]|null} [左, 上, 右, 下]。測れるものが無ければ null
     */
    function getSelectionVisibleBounds(targetItems) {
        if (!targetItems || !targetItems.length) return null;
        var left = null, top = null, right = null, bottom = null;
        for (var i = 0; i < targetItems.length; i++) {
            var itemBounds;
            try { itemBounds = targetItems[i].visibleBounds; } catch (e) { continue; }
            if (!itemBounds) continue;
            if (left === null || itemBounds[0] < left) left = itemBounds[0];
            if (top === null || itemBounds[1] > top) top = itemBounds[1];
            if (right === null || itemBounds[2] > right) right = itemBounds[2];
            if (bottom === null || itemBounds[3] < bottom) bottom = itemBounds[3];
        }
        if (left === null) return null;
        return [left, top, right, bottom];
    }

    /**
     * 複製 → アピアランス分割 → アウトラインで正確な可視サイズを測り、複製を破棄する
     * @param {PageItem} sourceItem - 測るオブジェクト
     * @returns {Object|null} { left, top, right, bottom, width, height }。測れなければ null
     */
    function measureAccurateBounds(sourceItem) {
        var doc = app.activeDocument;
        var savedSelection = doc.selection;
        var measuredBounds = null;
        var duplicatedItem = null;
        try {
            duplicatedItem = sourceItem.duplicate();
            doc.selection = null;
            duplicatedItem.selected = true;
            app.redraw();

            /* アピアランスを分割 / Expand appearance */
            try { app.executeMenuCommand('expandStyle'); } catch (e) { }
            /* テキストのアウトライン / Create outlines */
            try { app.executeMenuCommand('outline'); } catch (e) { }
            app.redraw();

            /* 分割・アウトライン後の選択全体のサイズを計測 / Measure bounds of the resulting selection */
            var expandedItems = doc.selection;
            var bounds = getSelectionVisibleBounds(expandedItems);
            if (bounds) {
                measuredBounds = { left: bounds[0], top: bounds[1], right: bounds[2], bottom: bounds[3], width: bounds[2] - bounds[0], height: bounds[1] - bounds[3] };
            }

            /* 複製（分割・アウトライン結果）を削除 / Remove the duplicate (expanded/outlined result) */
            for (var i = expandedItems.length - 1; i >= 0; i--) {
                try { expandedItems[i].remove(); } catch (e2) { }
            }
        } catch (e0) {
            if (duplicatedItem) { try { duplicatedItem.remove(); } catch (e3) { } }
        }

        try { doc.selection = savedSelection; } catch (e4) { }
        return measuredBounds;
    }

    // =========================================
    // シェイプ効果 / Shape effect
    // =========================================

    /**
     * ［形状に変換：長方形］固定モードのライブエフェクト定義を組み立てる
     * @param {number} width - 幅（pt）
     * @param {number} height - 高さ（pt）
     * @returns {string} ライブエフェクトの XML
     */
    function buildRectangleShapeEffectXML(width, height) {
        return '<LiveEffect name="Adobe Shape Effects" isPre="1"><Dict data="U DisplayString Rectangle I Shape 0 R RelWidth 0 R RelHeight 0 R AbsWidth ' + width + ' R AbsHeight ' + height + ' R Absolute 1 R CornerRadius 9 "/></LiveEffect>';
    }

    /**
     * 対象だけを選択し直す
     * @param {PageItem} targetItem - 選択するオブジェクト
     * @returns {void}
     */
    function reselectOnly(targetItem) {
        app.activeDocument.selection = null;
        targetItem.selected = true;
        app.redraw();
    }

    /**
     * 対象を選択し直してから New Fill を1枚追加する（毎回選択し直さないと2枚目が効かない）
     * @param {PageItem} targetItem - 塗りを足すオブジェクト
     * @returns {void}
     */
    function addNewFill(targetItem) {
        reselectOnly(targetItem);
        app.executeMenuCommand('Adobe New Fill Shortcut');
        app.redraw();
    }

    /**
     * 塗りを2枚追加し、固定サイズ（幅×高さ）の長方形シェイプ効果でボタン状の背景にする（塗り色は設定しない）
     * エリア→ポイント変換直後のテキストは最初の New Fill が吸収されるため、1回“空打ち”してから2枚追加する（計3回）
     * @param {PageItem} targetItem - 背景を付けるオブジェクト
     * @param {number} width - 背景の幅（pt）
     * @param {number} height - 背景の高さ（pt）
     * @returns {void}
     */
    function applyButtonBackgroundSized(targetItem, width, height) {
        try {
            addNewFill(targetItem); /* 空打ち（吸収される）/ primer (gets absorbed) */
            addNewFill(targetItem); /* 塗り1 / fill 1 */
            addNewFill(targetItem); /* 塗り2 / fill 2 */
            reselectOnly(targetItem);
            targetItem.applyEffect(buildRectangleShapeEffectXML(width, height));
        } catch (e) { }
    }

    // =========================================
    // 変換 / Conversion
    // =========================================

    /**
     * geometricBounds を { left, top, width, height } で返す
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @returns {Object} { left, top, width, height }
     */
    function geometricBoundsOf(pageItem) {
        var bounds = pageItem.geometricBounds;
        return { left: bounds[0], top: bounds[1], width: bounds[2] - bounds[0], height: bounds[1] - bounds[3] };
    }

    /**
     * 元テキストの自動カーニングと文字組みアキ量設定を取得する
     * @param {TextFrame} sourceText - 元テキスト
     * @returns {Object} { kerningMethod, mojikumi }（取れなければ null）
     */
    function readKerningAndMojikumi(sourceText) {
        var snapshot = { kerningMethod: null, mojikumi: null };
        try { snapshot.kerningMethod = sourceText.textRange.characterAttributes.kerningMethod; } catch (e) { }
        /* 文字組みアキ量設定「なし」は読むと例外 / Reading mojikumi "None" throws */
        try {
            if (sourceText.paragraphs.length > 0) {
                snapshot.mojikumi = sourceText.paragraphs[0].paragraphAttributes.mojikumi;
            }
        } catch (e2) { }
        return snapshot;
    }

    /**
     * 自動カーニングと文字組みアキ量設定をエリア内文字へ適用する
     * @param {TextFrame} areaType - 適用先のエリア内文字
     * @param {Object} snapshot - readKerningAndMojikumi() の結果
     * @returns {void}
     */
    function applyKerningAndMojikumi(areaType, snapshot) {
        if (snapshot.kerningMethod !== null) {
            try { areaType.textRange.characterAttributes.kerningMethod = snapshot.kerningMethod; } catch (e) { }
        }
        if (snapshot.mojikumi !== null && snapshot.mojikumi !== undefined) {
            var paragraphs = areaType.paragraphs;
            for (var i = 0; i < paragraphs.length; i++) {
                try { paragraphs[i].paragraphAttributes.mojikumi = snapshot.mojikumi; } catch (e2) { }
            }
        }
    }

    /**
     * 図形パスをエリア内文字にして、元テキストの内容・フォント・カーニング・文字組み・グラフィックスタイルを移す
     * @param {Document} doc - 対象のドキュメント
     * @param {PathItem} framePath - 枠にするパス
     * @param {TextFrame} sourceText - 元テキスト
     * @param {string|null} graphicStyleName - 適用するグラフィックスタイル名（null なら適用しない）
     * @returns {TextFrame} 作ったエリア内文字
     */
    function fillAreaTypeFromSourceText(doc, framePath, sourceText, graphicStyleName) {
        var sourceFont = null, sourceSize = 0;
        try {
            var sourceAttributes = sourceText.textRange.characterAttributes;
            sourceFont = sourceAttributes.textFont;
            sourceSize = sourceAttributes.size;
        } catch (e) { }
        var typeSnapshot = readKerningAndMojikumi(sourceText);

        var areaType = doc.textFrames.areaText(framePath);
        areaType.contents = sourceText.contents;
        try {
            if (sourceFont) areaType.textRange.characterAttributes.textFont = sourceFont;
            if (sourceSize > 0) areaType.textRange.characterAttributes.size = sourceSize;
        } catch (e2) { }
        applyKerningAndMojikumi(areaType, typeSnapshot);

        /* 登録したグラフィックスタイルを適用（削除は呼び出し側が finally で行う）
           Apply the registered graphic style (the caller removes it in a finally block) */
        if (graphicStyleName) { applyGraphicStyleByName(graphicStyleName, areaType); }
        return areaType;
    }

    /**
     * エリア内文字の内容を水平・垂直とも中央にする
     * @param {TextFrame} areaType - 対象のエリア内文字
     * @returns {void}
     */
    function centerAreaTypeContents(areaType) {
        try { areaType.textRange.paragraphAttributes.justification = Justification.CENTER; } catch (e) { }
        applyVerticalCenter(areaType); /* 縦位置は中央（アクション再生）/ Vertical center (via action) */
    }

    /* 変換中に握り潰した最後の例外（0件時にアラートで理由を添える）/ Last swallowed exception during conversion (surfaced on a 0-result alert) */
    var lastConversionError = null;

    /**
     * 変換結果を選択する。0件なら理由を添えてアラートを出す
     * @param {TextFrame[]} convertedItems - 変換で作ったテキスト
     * @param {string} alertKey - 0件のときのアラートの LABELS パス
     * @returns {void}
     */
    function finishConversion(convertedItems, alertKey) {
        if (convertedItems.length > 0) {
            try { app.activeDocument.selection = convertedItems; app.redraw(); } catch (e) { }
        } else {
            /* 1件も変換できなかった＝内部で例外を握り潰している。無言終了せず理由を伝える
               Nothing converted = an exception was swallowed internally. Surface it instead of exiting silently */
            var failMessage = getLabel(alertKey);
            if (lastConversionError) { failMessage += "\n" + lastConversionError; }
            alert(failMessage);
        }
    }

    /**
     * 選択内のテキストをエリア内文字へ変換する（パス上文字は先にポイント文字へ分離）
     * @param {Document} doc - 対象のドキュメント
     * @param {PageItem[]} currentSelection - 現在の選択
     * @param {Object} styleOptions - ダイアログの結果 { keepStyle }
     * @returns {void}
     */
    function convertSelectionToAreaType(doc, currentSelection, styleOptions) {
        lastConversionError = null;
        preprocessPathTextSelection(doc, currentSelection);

        /* 前処理で選択が差し替わるので取り直す / The pre-process replaces the selection, so read it again */
        var refreshedSelection = doc.selection;
        if (!refreshedSelection || refreshedSelection.length === 0) { return; }

        /* 変換元テキストの有無を判定 / Detect source text */
        var hasSourceText = false;
        for (var i = 0; i < refreshedSelection.length; i++) {
            if (isPointOrPathText(refreshedSelection[i])) hasSourceText = true;
        }
        if (!hasSourceText) { return; }

        /* 計測した実寸でフレーム化してエリア内文字へ / Frame at the measured real size, then convert to area type */
        finishConversion(convertPointTextToMeasuredArea(doc, refreshedSelection, styleOptions), "alert.conversionFailed");
    }

    /**
     * 計測した実寸でフレームを作り、テキストを流し込んで中央配置する（元のポイント文字は削除）
     * @param {Document} doc - 対象のドキュメント
     * @param {PageItem[]} targetItems - 対象（ポイント文字だけを変換）
     * @param {Object} styleOptions - ダイアログの結果 { keepStyle }
     * @returns {TextFrame[]} 作ったエリア内文字
     */
    function convertPointTextToMeasuredArea(doc, targetItems, styleOptions) {
        var createdAreaTypes = [];

        /* スタイル「保持する」なら見た目を一時スタイルで引き継ぐ（既定）/ Keep appearance via a temp style when "Keep" (default) */
        var keepStyle = !styleOptions || styleOptions.keepStyle !== false;

        for (var i = targetItems.length - 1; i >= 0; i--) {
            var sourceText = targetItems[i];
            if (getTextKind(sourceText) !== TextType.POINTTEXT) continue;
            /* 「保持する」のときだけ、元テキストの見た目を一時グラフィックスタイルとして登録
               Register the source text's appearance as a temp graphic style only when "Keep" */
            var tempStyleName = keepStyle ? registerTextFrameAsTempGraphicStyle(sourceText) : null;
            try {
                /* 複製→アピアランス分割→アウトラインで正確な可視サイズを計測（保持する/しない共通）
                   Measure accurate visible size via expand appearance → outline (same for keep / don't keep) */
                var measuredBounds = measureAccurateBounds(sourceText) || geometricBoundsOf(sourceText);
                /* 計測した実寸そのままの長方形をフレームに / Frame at the measured size as-is */
                var framePath = doc.pathItems.rectangle(measuredBounds.top, measuredBounds.left, measuredBounds.width, measuredBounds.height);
                framePath.filled = false; framePath.stroked = false;
                var areaType = fillAreaTypeFromSourceText(doc, framePath, sourceText, tempStyleName);
                centerAreaTypeContents(areaType);
                createdAreaTypes.push(areaType);
                sourceText.remove();
            } catch (e) {
                lastConversionError = e; /* 失敗理由を保持（0件時にアラートで表示）/ Keep the reason (shown on a 0-result alert) */
            } finally {
                /* アピアランスを引き継いだら一時グラフィックスタイルは削除 / Remove the temp style once its appearance is carried over */
                removeGraphicStyleByName(tempStyleName);
            }
        }

        return createdAreaTypes;
    }

    // =========================================
    // 逆変換：エリア内文字 → ポイント文字 / Reverse: area type → point text
    // =========================================

    /**
     * エリア内文字をポイント文字へ変換し、変換後のポイント文字を返す
     * convertAreaObjectToPointObject() は「その場変換・戻り値 null」で、変換後は元のラッパー（areaType や選択）が
     * stale になり kind を AREATEXT のまま報告することがある。そこで変換前に一時的な名前（マーカー）を付け、
     * 変換後に doc.textFrames を新規走査してその名前で確実に回収する
     * @param {Document} doc - 対象のドキュメント
     * @param {TextFrame} areaType - 変換するエリア内文字
     * @returns {TextFrame|null} 変換後のポイント文字。回収できなければ null
     */
    function convertAreaObjectToPointText(doc, areaType) {
        var MARKER_NAME = "__AreaTypeToPointText_marker__";
        var previousName = "";
        try { previousName = areaType.name; } catch (eName) { }

        try {
            try { areaType.name = MARKER_NAME; } catch (eSet) { }
            /* 対象だけを選択してから変換 / Select only the target, then convert */
            doc.selection = null;
            areaType.selected = true;
            app.redraw();
            areaType.convertAreaObjectToPointObject();
        } catch (e) {
            lastConversionError = e; /* 例外後も変換済みの場合があるので下で回収を試みる / May have converted anyway; try to recover below */
        }

        /* マーカー名で新規走査して回収（stale ラッパーに依存しない）/ Recover by marker name via a fresh scan (no stale wrappers) */
        var pointText = null;
        var textFrames = doc.textFrames;
        for (var i = 0; i < textFrames.length; i++) {
            var frameName = "";
            try { frameName = textFrames[i].name; } catch (eGet) { continue; }
            if (frameName === MARKER_NAME) { pointText = textFrames[i]; break; }
        }

        if (pointText) {
            try { pointText.name = previousName; } catch (eRestore) { } /* 元の名前へ戻す / Restore the original name */
            return pointText;
        }
        return null;
    }

    /**
     * 選択中のエリア内文字をポイント文字へ変換する（保持しないときは元の枠サイズ A×B のボタン背景を付ける）
     * @param {Document} doc - 対象のドキュメント
     * @param {PageItem[]} currentSelection - 現在の選択
     * @param {Object} styleOptions - ダイアログの結果 { keepStyle }
     * @returns {TextFrame[]} 変換後のポイント文字
     */
    function convertAreaTypeToPointText(doc, currentSelection, styleOptions) {
        var createdPointTexts = [];

        /* スタイル「保持する」なら変換のみ、「保持しない」なら塗り＋長方形背景を付与 / "Keep" = convert only; "Don't keep" = add fill + rectangle background */
        var keepStyle = !styleOptions || styleOptions.keepStyle !== false;

        /* 変換で選択が変わるため、対象のエリア内文字を先に確保 / Collect targets first (conversion changes the selection) */
        var areaTexts = [];
        for (var i = 0; i < currentSelection.length; i++) {
            if (getTextKind(currentSelection[i]) === TextType.AREATEXT) areaTexts.push(currentSelection[i]);
        }

        for (var j = areaTexts.length - 1; j >= 0; j--) {
            var areaType = areaTexts[j];
            try {
                /* 1) 元のエリア枠の実寸 A×B（「保持しない」の背景サイズに使う）/ Original frame size A×B (used for the "Don't keep" background) */
                var frameBounds = geometricBoundsOf(areaType);

                /* 2) ポイント文字へ変換（変換後の参照を取り直す）/ Convert to point text (re-acquire the reference) */
                var pointText = convertAreaObjectToPointText(doc, areaType);
                if (!pointText) continue;

                /* 3) 「保持しない」→ 塗り2枚＋固定サイズ（A×B）の長方形でボタン背景を付与。「保持する」→ 変換後の見た目のまま
                   "Don't keep" → add two fills + an absolute A×B rectangle background; "Keep" → leave as-is */
                if (!keepStyle) {
                    applyButtonBackgroundSized(pointText, frameBounds.width, frameBounds.height);
                }

                createdPointTexts.push(pointText);
            } catch (e) {
                lastConversionError = e; /* 失敗理由を保持（0件時にアラートで表示）/ Keep the reason (shown on a 0-result alert) */
            }
        }

        return createdPointTexts;
    }

    /**
     * 選択（エリア内文字）をポイント文字へ変換して結果を選択する
     * @param {Document} doc - 対象のドキュメント
     * @param {PageItem[]} currentSelection - 現在の選択
     * @param {Object} styleOptions - ダイアログの結果 { keepStyle }
     * @returns {void}
     */
    function convertSelectionToPointText(doc, currentSelection, styleOptions) {
        lastConversionError = null;
        finishConversion(convertAreaTypeToPointText(doc, currentSelection, styleOptions), "alert.reverseFailed");
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * テキストフレームの種類の表示名を返す
     * @param {PageItem} pageItem - 調べるオブジェクト
     * @returns {string|null} 種類の表示名。テキスト以外は null
     */
    function describeTextType(pageItem) {
        try {
            var textKind = getTextKind(pageItem);
            if (textKind === TextType.POINTTEXT) return getLabel("textType.pointText");
            if (textKind === TextType.AREATEXT) return getLabel("textType.areaText");
            if (textKind === TextType.PATHTEXT) return getLabel("textType.pathText");
        } catch (e) { }
        return null;
    }

    /**
     * 選択の内訳を「種類 ×個数」で要約する（出てきた順）
     * @param {PageItem[]} currentSelection - 現在の選択
     * @returns {string} 内訳の文字列
     */
    function summarizeSelection(currentSelection) {
        var typeCounts = {}, typeOrder = [];
        for (var i = 0; i < currentSelection.length; i++) {
            var typeLabel = describeTextType(currentSelection[i]) || getLabel("textType.other");
            if (!typeCounts.hasOwnProperty(typeLabel)) { typeCounts[typeLabel] = 0; typeOrder.push(typeLabel); }
            typeCounts[typeLabel]++;
        }
        var summaryParts = [];
        for (var k = 0; k < typeOrder.length; k++) { summaryParts.push(typeOrder[k] + " ×" + typeCounts[typeOrder[k]]); }
        return summaryParts.join("、");
    }

    /**
     * テキスト変換のオプションダイアログ（選択の内訳＋スタイル：保持する/保持しない）を表示する
     * @param {PageItem[]} currentSelection - 現在の選択
     * @param {boolean} isReverse - 逆変換（エリア内文字 → ポイント文字）なら true。ツールチップを切り替える
     * @returns {Object|null} OK なら { keepStyle }、キャンセルなら null
     */
    function showStyleDialog(currentSelection, isReverse) {
        var styleDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        styleDialog.orientation = "column";
        styleDialog.alignChildren = "fill";
        styleDialog.margins = DIALOG_MARGINS;

        /* 現在の選択オブジェクトの内訳を表示 / Show a summary of the current selection */
        var selectionPanel = styleDialog.add("panel", undefined, getLabel("panel.selection"));
        setupPanel(selectionPanel, 6);
        selectionPanel.add("statictext", undefined, summarizeSelection(currentSelection));

        /* スタイル：アピアランス（見た目）を保持するか / Style: keep the appearance or not */
        var stylePanel = styleDialog.add("panel", undefined, getLabel("panel.style"));
        setupPanel(stylePanel, 6);
        var styleChoiceGroup = stylePanel.add("group");
        var styleKeepRadio = styleChoiceGroup.add("radiobutton", undefined, getLabel("radio.keepStyle"));
        var styleDontKeepRadio = styleChoiceGroup.add("radiobutton", undefined, getLabel("radio.dontKeepStyle"));
        styleKeepRadio.value = true; /* 既定は「保持する」/ Default: keep */
        /* 順/逆で意味が異なるためツールチップで補足 / Tooltips clarify the per-direction meaning */
        styleKeepRadio.helpTip = getLabel(isReverse ? "tooltip.keepReverse" : "tooltip.keepForward");
        styleDontKeepRadio.helpTip = getLabel(isReverse ? "tooltip.dontKeepReverse" : "tooltip.dontKeepForward");

        /* ボタン（左右中央、Mac規約：キャンセル → OK）/ Buttons (centered; Mac order: Cancel → OK) */
        var btnRowGroup = styleDialog.add("group");
        btnRowGroup.alignment = "center";
        var btnCancel = btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        var dialogResult = null;
        btnOK.onClick = function () {
            dialogResult = { keepStyle: styleKeepRadio.value };
            styleDialog.close();
        };
        btnCancel.onClick = function () { dialogResult = null; styleDialog.close(); };

        /* 表示直後にレイアウトを再計算して描画欠けを防ぐ / Recalculate layout on show to avoid partial rendering */
        styleDialog.onShow = function () {
            styleDialog.layout.layout(true);
            styleDialog.layout.resize();
        };

        styleDialog.show();
        return dialogResult;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択にエリア内文字、ポイント文字・パス上文字が含まれるかを調べる
     * @param {PageItem[]} currentSelection - 現在の選択
     * @returns {{hasAreaText: boolean, hasPointOrPathText: boolean}} 含まれる種類
     */
    function detectSelectionKinds(currentSelection) {
        var selectionKinds = { hasAreaText: false, hasPointOrPathText: false };
        for (var i = 0; i < currentSelection.length; i++) {
            var textKind = getTextKind(currentSelection[i]);
            if (textKind === TextType.AREATEXT) selectionKinds.hasAreaText = true;
            else if (textKind === TextType.POINTTEXT || textKind === TextType.PATHTEXT) selectionKinds.hasPointOrPathText = true;
        }
        return selectionKinds;
    }

    /**
     * 選択に応じて、エリア内文字なら逆変換、ポイント文字・パス上文字なら順変換する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;
        var currentSelection = doc.selection;
        if (!currentSelection || currentSelection.length === 0) {
            alert(getLabel("alert.selectText"));
            return;
        }

        var selectionKinds = detectSelectionKinds(currentSelection);
        if (selectionKinds.hasAreaText) {
            /* 逆変換：ダイアログ（選択表示＋保持する/保持しない）→ ポイント文字へ変換
               Reverse: dialog (selection summary + keep / don't keep) → convert to point text */
            var reverseOptions = showStyleDialog(currentSelection, true);
            if (reverseOptions) {
                convertSelectionToPointText(doc, currentSelection, reverseOptions);
            }
        } else if (selectionKinds.hasPointOrPathText) {
            /* 順変換：ダイアログを表示（キャンセルで中止）/ Forward: show the dialog (Cancel aborts) */
            var forwardOptions = showStyleDialog(currentSelection, false);
            if (!forwardOptions) return;
            /* フレーム整列アクションを読み込み、終了時に破棄 / Load the frame-alignment action, unload on exit */
            loadFrameAlignmentAction();
            try {
                convertSelectionToAreaType(doc, currentSelection, forwardOptions);
            } finally {
                unloadFrameAlignmentAction();
            }
        } else {
            alert(getLabel("alert.selectText"));
        }
    }

    main();

})();
