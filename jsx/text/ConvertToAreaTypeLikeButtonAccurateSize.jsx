#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ポイント文字・パス上文字・図形＋テキストを、見た目を保ったままエリア内文字へ変換する。

処理の流れ：

1. 変換前に、選択テキストの見た目をグラフィックスタイルとして一時登録（ダイナミックアクション）
2. ポイント文字 / パス上文字 → 計測した実寸のエリア内文字へ変換
   テキスト＋図形 → 図形をエリア内文字にしてテキストを流し込む
3. 変換後のエリア内文字に登録したグラフィックスタイルを適用し、その一時スタイルを削除
4. 行揃え・テキストの配置は常に中央

#### 補足

- フレームサイズは、複製→アピアランス分割→アウトラインで計測した実寸を基準に決定する。
- 縦方向の中央配置とグラフィックスタイル登録にはダイナミックアクションを使用する。実行時に一時的に読み込み、終了時に自動で破棄するため、アクションパネルに残骸は残らない。

### Overview

Convert point text, path text, or text + shape into area type while preserving appearance.

Flow:

1. Before converting, register the selected text's appearance as a temporary graphic style (dynamic action).
2. Point / path text → area type at the measured real size.
   Text + shape → fill the shape (as area type) with the text.
3. Apply the registered graphic style to the converted area type, then remove the temporary style.
4. Justification and vertical alignment are always centered.

#### Notes

- Frame size is based on the real size measured via duplicate → expand appearance → create outlines.
- Vertical centering and graphic-style registration use dynamic actions that are loaded temporarily and removed automatically on exit, so nothing is left behind in the Actions panel.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ConvertToAreaTypeLikeButtonAccurateSize";  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                                   /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";              /* 作者 / author */
var SCRIPT_RELEASED = "";                                         /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                                         /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 表示言語を判定 / Detect display language */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    /* 日英ラベル定義（カテゴリ構造）/ Japanese-English label definitions (categorized) */
    var LABELS = {
        /* 警告メッセージ / Alerts */
        alert: {
            selectText: {
                ja: "ポイント文字・パス上文字、またはテキストと図形を選択してください。",
                en: "Please select point text, path text, or text and a shape."
            },
            noDocument: {
                ja: "ドキュメントが開かれていません。",
                en: "No document is open."
            }
        }
    };

    /* ドット区切りキーから LABELS のエントリを取得 / Resolve a LABELS entry from a dot-separated key */
    function getLabelEntry(key) {
        var parts = key.split(".");
        var cur = LABELS;
        for (var i = 0; i < parts.length; i++) {
            if (cur && typeof cur[parts[i]] !== "undefined") { cur = cur[parts[i]]; }
            else { return null; }
        }
        return cur;
    }

    /* キーからローカライズ文字列を取得 / Get a localized string by key */
    function L(key) {
        var entry = getLabelEntry(key);
        if (entry) {
            if (entry[currentLanguage]) return entry[currentLanguage];
            if (entry.en) return entry.en;
        }
        return key;
    }

    // =========================================
    // パス上文字 → ポイント文字（変換前処理）/ Path text → Point text (pre-process)
    // =========================================

    /* パス上文字を字形を保ったままポイント文字へ分離 / Detach path text into point text keeping attributes */
    function detachPathTextToPointText(doc, pathTextFrames) {
        function safe(fn) { try { return fn(); } catch (e) { return undefined; } }

        var created = [];
        if (!doc || !pathTextFrames || !pathTextFrames.length) return created;

        // 新規テキストだけ選べるよう選択を解除 / Clear selection
        safe(function () { doc.selection = null; });

        for (var j = pathTextFrames.length - 1; j >= 0; j--) {
            var pathText = pathTextFrames[j];
            if (!pathText || pathText.typename !== "TextFrame" || pathText.kind !== TextType.PATHTEXT) continue;

            var originalPath = null;
            safe(function () { originalPath = pathText.textPath; });
            if (!originalPath) continue;

            // 1) 文字ごとの属性を退避 / Snapshot per-character attributes
            var charAttrs = [];
            for (var c = 0; c < pathText.characters.length; c++) {
                var charAttr = pathText.characters[c].characterAttributes;
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

            var textContents = "";
            safe(function () { textContents = pathText.contents; });

            var justification = null;
            safe(function () {
                if (pathText.paragraphs && pathText.paragraphs.length > 0) {
                    justification = pathText.paragraphs[0].paragraphAttributes.justification;
                }
            });

            // 2) パス始点にポイント文字を新規作成 / Create new point text at path start anchor
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

            // 既定の線を一旦消し、後で文字ごとに復元 / Clear default stroke, restore per-character later
            safe(function () {
                var noColor = new NoColor();
                newText.textRange.characterAttributes.strokeColor = noColor;
                newText.textRange.characterAttributes.strokeWeight = 0;
            });

            // 文字ごとの属性を復元 / Restore per-character attributes
            var restoreCount = Math.min(newText.characters.length, charAttrs.length);
            for (var k = 0; k < restoreCount; k++) {
                var targetAttr = newText.characters[k].characterAttributes;
                var sourceAttr = charAttrs[k];

                safe(function () { targetAttr.textFont = sourceAttr.font; });
                safe(function () { targetAttr.size = sourceAttr.size; });
                safe(function () { targetAttr.fillColor = sourceAttr.fillColor; });
                safe(function () {
                    var sourceStroke = sourceAttr.strokeColor;
                    targetAttr.strokeColor = sourceStroke;
                    targetAttr.strokeWeight = (sourceStroke && sourceStroke.typename === "NoColor") ? 0 : sourceAttr.strokeWeight;
                });
                safe(function () { targetAttr.baselineShift = 0; });
                safe(function () { targetAttr.horizontalScale = 100; });
                safe(function () { targetAttr.verticalScale = 100; });
                safe(function () { targetAttr.autoLeading = sourceAttr.autoLeading; });
                if (!sourceAttr.autoLeading) safe(function () { targetAttr.leading = sourceAttr.leading; });
            }

            // 3) 元のパス上文字を削除（パスも一緒に消える）/ Remove original path text
            safe(function () { pathText.remove(); });

            // 4) 新規テキストを選択して返す / Select and return new text
            safe(function () { newText.selected = true; });
            created.push(newText);
        }

        return created;
    }

    /* 選択内のパス上文字をポイント文字へ置き換えた選択配列を返す / Replace path text in selection with point text */
    function preprocessPathTextSelection(doc, sel) {
        if (!doc || !sel || !sel.length) return sel;

        var pathTexts = [];
        for (var i = 0; i < sel.length; i++) {
            var item = sel[i];
            try {
                if (item && item.typename === "TextFrame" && item.kind === TextType.PATHTEXT) {
                    pathTexts.push(item);
                }
            } catch (e0) {
                // 無効オブジェクト（削除済み等）はスキップ / Skip invalid objects
            }
        }
        if (!pathTexts.length) return sel;

        var newTexts = detachPathTextToPointText(doc, pathTexts);
        if (!newTexts.length) return sel;

        // パス上文字を新ポイント文字に差し替えた新しい選択配列を構築 / Build replaced selection array
        var replacedSelection = [];
        for (var j = 0; j < sel.length; j++) {
            var keepItem = sel[j];
            try {
                if (keepItem && keepItem.typename === "TextFrame" && keepItem.kind === TextType.PATHTEXT) {
                    // 旧オブジェクトは除外 / skip old
                } else if (keepItem) {
                    replacedSelection.push(keepItem);
                }
            } catch (e1) {
                // 無効オブジェクトはスキップ / Skip invalid objects
            }
        }
        for (var k = 0; k < newTexts.length; k++) replacedSelection.push(newTexts[k]);

        try { doc.selection = replacedSelection; } catch (e) { }
        try { app.redraw(); } catch (e2) { }

        return replacedSelection;
    }

    // =========================================
    // ダイナミックアクション / Dynamic actions
    // =========================================

    /* 文字列を16進数表現へ / Convert a string to hex */
    function _hexAscii(text) {
        var hex = "";
        for (var i = 0; i < text.length; i++) {
            var pair = text.charCodeAt(i).toString(16);
            if (pair.length < 2) pair = "0" + pair;
            hex += pair;
        }
        return hex;
    }

    /* /name ブロック（ASCII）を生成 / Build a /name block (ASCII) */
    function _nameBlockAscii(text) {
        return "/name [ " + text.length + " " + _hexAscii(text).toUpperCase() + " ]";
    }

    /* アクションセット定義(.aia)文字列を組み立て / Build an action set (.aia) definition string */
    function _buildActionSetAIA(setName, internalName, localizedNameHex, paramKeyInt, actionDefs) {
        var aia = "/version 3" +
            _nameBlockAscii(setName) +
            "/isOpen 1" +
            "/actionCount " + actionDefs.length;

        for (var i = 0; i < actionDefs.length; i++) {
            var actionDef = actionDefs[i];
            aia += "/action-" + (i + 1) + " {" +
                " " + _nameBlockAscii(actionDef.name) +
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
                " /value " + actionDef.value +
                " }" +
                " }" +
                "}";
        }
        return aia;
    }

    /* フレーム整列アクションセット名 / Frame-alignment action set name */
    var AREA_TEXT_ACTION_SET = "AreaText";

    /* アクションセットを読み込む（temp に .aia を書き出して loadAction）/ Load an action set (write .aia to temp, then loadAction) */
    function loadActionSet(setName, aiaString) {
        try {
            // 既存があれば一旦外して衝突回避 / Unload existing set first to avoid conflicts
            try { app.unloadAction(setName, ""); } catch (e0) { }

            var f = new File(Folder.temp + "/AreaTypeToolkit_action_" + setName + ".aia");
            f.open("w");
            f.write(aiaString);
            f.close();

            app.loadAction(f);
            try { f.remove(); } catch (e1) { }
        } catch (e) { }
    }

    /* アクションセットを破棄 / Unload an action set */
    function unloadActionSet(setName) {
        try { app.unloadAction(setName, ""); } catch (e) { }
    }

    /* フレーム整列アクション（AlignTop/Center/Bottom/Justify）を読み込む / Load frame-alignment actions */
    function loadAreaTextActions() {
        var aia = _buildActionSetAIA(
            AREA_TEXT_ACTION_SET,
            "adobe_frameAlignment",
            "39 e382a8e383aae382a2e58685e69687e5ad97e381aee38395e383ace383bce383a0e695b4e58897",
            1717660782,
            [
                { name: "AlignTop", value: 0 },
                { name: "AlignCenter", value: 1 },
                { name: "AlignBottom", value: 2 },
                { name: "AlignJustify", value: 3 }
            ]
        );
        loadActionSet(AREA_TEXT_ACTION_SET, aia);
    }

    /* フレーム整列アクションを破棄 / Unload frame-alignment actions */
    function unloadAreaTextActions() {
        unloadActionSet(AREA_TEXT_ACTION_SET);
    }

    /* エリア内文字のフレーム整列（縦方向の配置）アクションを実行 / Run the area-text frame-alignment action (vertical placement)
       valueInt: 0=上/top, 1=中央/center, 2=下/bottom, 3=均等/justify */
    function act_areaTextFrameAlignment(valueInt) {
        if (valueInt !== 0 && valueInt !== 1 && valueInt !== 2 && valueInt !== 3) return;

        var actionName = "AlignTop";
        if (valueInt === 1) actionName = "AlignCenter";
        else if (valueInt === 2) actionName = "AlignBottom";
        else if (valueInt === 3) actionName = "AlignJustify";

        try { app.doScript(actionName, AREA_TEXT_ACTION_SET, false); } catch (e) { }
    }

    /* 指定フレームにフレーム整列（縦方向の配置）を適用 / Apply frame alignment (vertical placement) to a frame */
    function applyAreaTextFrameAlignment(areaType, valueInt) {
        try {
            var doc = app.activeDocument;
            doc.selection = null;
            doc.selection = [areaType];
            app.redraw(); // 選択状態を確定 / Commit the selection
            act_areaTextFrameAlignment(valueInt);
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

    /* TextFrame の見た目を temp_style として登録（既存があれば上書き）。登録名を返す
       Register the TextFrame's appearance as temp_style (overwrites if it exists); returns the registered name */
    function registerTextFrameAsTempGraphicStyle(textFrame) {
        if (!textFrame) return null;
        var doc = app.activeDocument;
        if (!doc) return null;
        var graphicStyles = doc.graphicStyles;
        if (!graphicStyles) return null;

        // 既存の temp_style を削除 / Remove any existing temp_style
        try { graphicStyles.getByName(TEMP_STYLE_NAME).remove(); } catch (e) { }

        // 登録用に対象だけを選択 / Select only the target for registration
        doc.selection = null;
        try { textFrame.selected = true; } catch (selectError) { return null; }

        var countBefore = graphicStyles.length;
        loadActionSet(TEMP_STYLE_ACTION_SET, GRAPHIC_STYLE_AIA);
        try { app.doScript(TEMP_STYLE_ACTION_NAME, TEMP_STYLE_ACTION_SET, false); } catch (runError) { }
        unloadActionSet(TEMP_STYLE_ACTION_SET);

        // 末尾に増えたスタイルを temp_style に改名 / Rename the newly appended style to temp_style
        if (graphicStyles.length <= countBefore) return null;
        graphicStyles[graphicStyles.length - 1].name = TEMP_STYLE_NAME;
        return TEMP_STYLE_NAME;
    }

    /* 名前でグラフィックスタイルを適用 / Apply a graphic style by name */
    function applyGraphicStyleByName(name, item) {
        if (!name || !item) return false;
        try {
            app.activeDocument.graphicStyles.getByName(name).applyTo(item);
            return true;
        } catch (e) { return false; }
    }

    /* 名前でグラフィックスタイルを削除 / Remove a graphic style by name */
    function removeGraphicStyleByName(name) {
        if (!name) return;
        try { app.activeDocument.graphicStyles.getByName(name).remove(); } catch (e) { }
    }

    // =========================================
    // 正確なサイズ計測 / Accurate size measurement
    // =========================================

    /* 選択オブジェクト群の可視バウンディングボックスの和を返す / Union of visibleBounds over a selection */
    function getSelectionVisibleBounds(sel) {
        if (!sel || !sel.length) return null;
        var left = null, top = null, right = null, bottom = null;
        for (var i = 0; i < sel.length; i++) {
            var vb;
            try { vb = sel[i].visibleBounds; } catch (e) { continue; }
            if (!vb) continue;
            if (left === null || vb[0] < left) left = vb[0];
            if (top === null || vb[1] > top) top = vb[1];
            if (right === null || vb[2] > right) right = vb[2];
            if (bottom === null || vb[3] < bottom) bottom = vb[3];
        }
        if (left === null) return null;
        return [left, top, right, bottom];
    }

    /* 複製→アピアランス分割→テキストのアウトラインで正確な可視サイズを計測し、複製を破棄
       Measure accurate visible size via duplicate → expand appearance → create outlines, then discard the copy */
    function measureAccurateBounds(item) {
        var doc = app.activeDocument;
        var savedSel = doc.selection;
        var result = null;
        var dup = null;
        try {
            dup = item.duplicate();
            doc.selection = null;
            dup.selected = true;
            app.redraw();

            // アピアランスを分割 / Expand appearance
            try { app.executeMenuCommand('expandStyle'); } catch (e) { }
            // テキストのアウトライン / Create outlines
            try { app.executeMenuCommand('outline'); } catch (e) { }
            app.redraw();

            // 分割・アウトライン後の選択全体のサイズを計測 / Measure bounds of the resulting selection
            var resultSel = doc.selection;
            var b = getSelectionVisibleBounds(resultSel);
            if (b) {
                result = { left: b[0], top: b[1], right: b[2], bottom: b[3], width: b[2] - b[0], height: b[1] - b[3] };
            }

            // 複製（分割・アウトライン結果）を削除 / Remove the duplicate (expanded/outlined result)
            for (var d = resultSel.length - 1; d >= 0; d--) {
                try { resultSel[d].remove(); } catch (e2) { }
            }
        } catch (e0) {
            if (dup) { try { dup.remove(); } catch (e3) { } }
        }

        try { doc.selection = savedSel; } catch (e4) { }
        return result;
    }

    // =========================================
    // シェイプ効果 / Shape effect
    // =========================================

    /* ［形状に変換：長方形］ライブエフェクトの定義 / "Convert to Shape: Rectangle" live-effect definition */
    var RECTANGLE_SHAPE_EFFECT_XML = '<LiveEffect name="Adobe Shape Effects" isPre="1"><Dict data="U DisplayString Rectangle I Shape 0 R RelWidth 0 R RelHeight 0 R AbsWidth 0 R AbsHeight 0 R Absolute 0 R CornerRadius 9 "/></LiveEffect>';

    /* 塗りを2つ追加し、長方形のシェイプ効果でボタン状の背景にする / Add two fills and a rectangle shape effect for a button-like background */
    function applyButtonShapeAppearance(areaType) {
        try {
            app.activeDocument.selection = null;
            areaType.selected = true;
            app.redraw();
            app.executeMenuCommand('Adobe New Fill Shortcut');
            app.executeMenuCommand('Adobe New Fill Shortcut');
            areaType.applyEffect(RECTANGLE_SHAPE_EFFECT_XML);
        } catch (e) { }
    }

    // =========================================
    // 変換 / Conversion
    // =========================================

    /* 直線4辺・直角の長方形パスか（回転は許容、角丸は除外）
       Whether the path is a rectangle: 4 straight sides at right angles (rotation allowed, rounded corners excluded) */
    function isRectanglePath(item) {
        if (!item || item.typename !== "PathItem" || !item.closed) return false;
        var pathPoints = item.pathPoints;
        if (!pathPoints || pathPoints.length !== 4) return false;

        // すべて直線セグメント（ハンドルがアンカーと一致）/ All segments straight (handles coincide with anchors)
        for (var i = 0; i < 4; i++) {
            var point = pathPoints[i];
            if (Math.abs(point.leftDirection[0] - point.anchor[0]) > 0.01 ||
                Math.abs(point.leftDirection[1] - point.anchor[1]) > 0.01 ||
                Math.abs(point.rightDirection[0] - point.anchor[0]) > 0.01 ||
                Math.abs(point.rightDirection[1] - point.anchor[1]) > 0.01) {
                return false;
            }
        }

        // 隣接する辺が直角か（正規化した内積が0付近）/ Adjacent edges perpendicular (normalized dot near zero)
        for (var j = 0; j < 4; j++) {
            var a = pathPoints[j].anchor;
            var b = pathPoints[(j + 1) % 4].anchor;
            var c = pathPoints[(j + 2) % 4].anchor;
            var edge1x = b[0] - a[0], edge1y = b[1] - a[1];
            var edge2x = c[0] - b[0], edge2y = c[1] - b[1];
            var length1 = Math.sqrt(edge1x * edge1x + edge1y * edge1y);
            var length2 = Math.sqrt(edge2x * edge2x + edge2y * edge2y);
            if (length1 < 0.0001 || length2 < 0.0001) return false;
            if (Math.abs((edge1x * edge2x + edge1y * edge2y) / (length1 * length2)) > 0.02) return false;
        }
        return true;
    }

    /* エリア内文字のフレームに使える図形か（長方形のみ。areaText は複合パス非対応）
       Can this item serve as an area-type frame (rectangles only; areaText rejects compound paths) */
    function isAreaFrameCandidate(item) {
        return isRectanglePath(item);
    }

    /* geometricBounds を {left, top, width, height} で返す / geometricBounds as {left, top, width, height} */
    function geometricBoundsOf(item) {
        var bounds = item.geometricBounds;
        return { left: bounds[0], top: bounds[1], width: bounds[2] - bounds[0], height: bounds[1] - bounds[3] };
    }

    /* 元テキストの自動カーニングと文字組みアキ量設定を取得 / Capture auto-kerning and mojikumi from the source text */
    function readKerningAndMojikumi(sourceText) {
        var snapshot = { kerningMethod: null, mojikumi: null };
        try { snapshot.kerningMethod = sourceText.textRange.characterAttributes.kerningMethod; } catch (e) { }
        try {
            if (sourceText.paragraphs.length > 0) {
                snapshot.mojikumi = sourceText.paragraphs[0].paragraphAttributes.mojikumi;
            }
        } catch (e2) { }
        return snapshot;
    }

    /* 自動カーニングと文字組みアキ量設定をエリア内文字へ適用 / Apply auto-kerning and mojikumi to the area type */
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

    /* 図形パスをエリア内文字にして、元テキストの内容・フォント・カーニング・文字組み・グラフィックスタイルを移す
       Turn a path into area type carrying the source text's contents, font, kerning, mojikumi, and graphic style */
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

        // 登録したグラフィックスタイルを適用（削除は呼び出し側が finally で行う）
        // Apply the registered graphic style (the caller removes it in a finally block)
        if (graphicStyleName) { applyGraphicStyleByName(graphicStyleName, areaType); }
        return areaType;
    }

    /* エリア内文字を水平・垂直とも中央に / Center area-type contents horizontally and vertically */
    function centerAreaTypeContents(areaType) {
        try { areaType.textRange.paragraphAttributes.justification = Justification.CENTER; } catch (e) { }
        // 縦位置はDOMで不安定なためアクションで中央 / Vertical center via dynamic action (unreliable via DOM)
        applyAreaTextFrameAlignment(areaType, 1);
    }

    /* 選択からモードを自動判定してエリア内文字へ変換 / Auto-detect mode and convert to area type */
    function convertSelectionToAreaType(doc, sel) {
        sel = preprocessPathTextSelection(doc, sel);

        var selection = app.activeDocument.selection;
        if (!selection || selection.length === 0) { return; }

        // 変換元テキストと、フレームに使える図形の有無を判定 / Detect source text and a usable frame shape
        var hasSourceText = false, hasFrameShape = false;
        for (var i = 0; i < selection.length; i++) {
            var item = selection[i];
            if (item.typename === "TextFrame" && (item.kind === TextType.POINTTEXT || item.kind === TextType.PATHTEXT)) hasSourceText = true;
            if (isAreaFrameCandidate(item)) hasFrameShape = true;
        }
        if (!hasSourceText) { return; }

        // テキスト＋図形 → 図形をフレームに / テキストのみ → 計測した実寸でフレーム
        // Text + shape → use the shape as the frame; text only → frame at the measured real size
        var created = hasFrameShape
            ? convertTextIntoShape(doc, selection)
            : convertPointTextToMeasuredArea(doc, selection);

        if (created.length > 0) {
            try { app.activeDocument.selection = created; app.redraw(); } catch (e) { }
        }
    }

    /* テキスト＋図形：図形を複製してエリア内文字にする / Text + shape: duplicate the shape into area type */
    function convertTextIntoShape(doc, selection) {
        var created = [];

        var sourceText = null, sourceShape = null;
        for (var i = 0; i < selection.length; i++) {
            var item = selection[i];
            if (!sourceText && item.typename === "TextFrame") { sourceText = item; }
            else if (!sourceShape && isAreaFrameCandidate(item)) { sourceShape = item; }
        }
        if (!sourceText || !sourceShape) return created;

        // 元テキストの見た目をグラフィックスタイルに登録 / Register source appearance as a graphic style
        var graphicStyleName = registerTextFrameAsTempGraphicStyle(sourceText);
        try {
            var framePath = sourceShape.duplicate();
            framePath.filled = false; framePath.stroked = false;
            var areaType = fillAreaTypeFromSourceText(doc, framePath, sourceText, graphicStyleName);
            centerAreaTypeContents(areaType);
            // 塗り＋長方形シェイプ効果でボタン状の背景を付与 / Add a fill + rectangle shape effect for a button-like background
            applyButtonShapeAppearance(areaType);
            sourceText.remove();
            sourceShape.remove();
            created.push(areaType);
        } catch (e) {
        } finally {
            // 一時グラフィックスタイルは成否に関わらず削除 / Remove the temp graphic style regardless of success
            removeGraphicStyleByName(graphicStyleName);
        }

        return created;
    }

    /* ポイント文字のみ：計測した実寸でフレームを作り、中央配置 / Point text only: build a frame at the measured size, centered */
    function convertPointTextToMeasuredArea(doc, selection) {
        var created = [];

        for (var i = selection.length - 1; i >= 0; i--) {
            var sourceText = selection[i];
            if (!(sourceText.typename === "TextFrame" && sourceText.kind === TextType.POINTTEXT)) continue;
            // 元テキストの見た目をグラフィックスタイルに登録 / Register source appearance as a graphic style
            var graphicStyleName = registerTextFrameAsTempGraphicStyle(sourceText);
            try {
                // 複製→アピアランス分割→アウトラインで正確な可視サイズを計測 / Measure accurate visible size
                var measuredBounds = measureAccurateBounds(sourceText) || geometricBoundsOf(sourceText);
                var framePath = doc.pathItems.rectangle(measuredBounds.top, measuredBounds.left, measuredBounds.width, measuredBounds.height);
                framePath.filled = false; framePath.stroked = false;
                var areaType = fillAreaTypeFromSourceText(doc, framePath, sourceText, graphicStyleName);
                centerAreaTypeContents(areaType);
                created.push(areaType);
                sourceText.remove();
            } catch (e) {
            } finally {
                // 一時グラフィックスタイルは成否に関わらず削除 / Remove the temp graphic style regardless of success
                removeGraphicStyleByName(graphicStyleName);
            }
        }

        return created;
    }

    // =========================================
    // エントリポイント / Entry point
    // =========================================
    if (app.documents.length > 0) {
        var doc = app.activeDocument;
        var sel = doc.selection;

        if (sel && sel.length > 0) {
            var hasConvertibleText = false;
            for (var i = 0; i < sel.length; i++) {
                if (sel[i].typename === "TextFrame" && (sel[i].kind === TextType.POINTTEXT || sel[i].kind === TextType.PATHTEXT)) {
                    hasConvertibleText = true;
                }
            }

            if (hasConvertibleText) {
                // フレーム整列アクションを読み込み、終了時に破棄 / Load frame-alignment actions, unload on exit
                loadAreaTextActions();
                try {
                    convertSelectionToAreaType(doc, sel);
                } finally {
                    unloadAreaTextActions();
                }
            } else {
                alert(L("alert.selectText"));
            }
        } else {
            alert(L("alert.selectText"));
        }
    } else {
        alert(L("alert.noDocument"));
    }

})();
