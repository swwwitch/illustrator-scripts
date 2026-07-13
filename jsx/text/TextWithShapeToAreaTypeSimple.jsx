#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ポイント文字・パス上文字を、見た目（アピアランス）を保ったままエリア内文字へ変換する。

処理の流れ：

1. 実行時にオプションダイアログを表示（大きさ調整の する/しない と 幅%・高さ%）
2. 変換前に、選択テキストの見た目を一時グラフィックスタイルとして登録
3. ポイント文字 / パス上文字 → 計測した実寸（大きさ調整ONなら幅・高さに倍率）でエリア内文字へ変換し、登録した一時スタイルを適用してアピアランスを引き継ぐ（適用後に一時スタイルは削除）
4. 大きさ調整ONのときはボタン背景（New Fill を2枚追加＋長方形シェイプ効果）を付与
5. 行揃え・テキストの配置は常に中央

#### 補足

- フレームサイズは、複製→アピアランス分割→アウトラインで計測した実寸を基準に決定する。
- 縦方向の中央配置とグラフィックスタイル登録にはダイナミックアクションを使用する。実行時に一時的に読み込み、終了時に自動で破棄するため、アクションパネルに残骸は残らない。
- 1件も変換できなかった場合は無言で終わらず警告を表示する。

#### 対応している選択

- ポイント文字 … そのまま計測してエリア内文字化
- パス上文字 … 内部で一度ポイント文字に分離（detachPathTextToPointText）してから同じ経路で処理
- 複数選択時はそれぞれ個別に変換

#### オプションの効き方

| | 大きさ調整の倍率 | ボタン背景（塗り2枚＋長方形効果） |
|---|---|---|
| する | ○（幅×1.2 / 高さ×1.6 等） | ○ |
| しない | 等倍 | ✕ |

- ボタン背景は New Fill を2枚追加し、長方形シェイプ効果で背景化する。塗り色は設定しない（追加した塗りは既定のまま）。

### Overview

Convert point text or path text into area type while preserving its appearance.

Flow:

1. On launch, show the options dialog (size adjustment On/Off with width% / height%).
2. Before converting, register the selected text's appearance as a temporary graphic style.
3. Point / path text → area type at the measured real size (scaled by the width/height ratios when size adjustment is On), then apply the temp style to carry over the appearance (the temp style is removed afterwards).
4. When size adjustment is On, add a button background (two New Fills + a rectangle shape effect).
5. Justification and vertical alignment are always centered.

#### Notes

- Frame size is based on the real size measured via duplicate → expand appearance → create outlines.
- Vertical centering and graphic-style registration use dynamic actions that are loaded temporarily and removed automatically on exit, so nothing is left behind in the Actions panel.
- If nothing could be converted, the script alerts instead of exiting silently.

#### Supported selections

- Point text … measured directly and converted to area type
- Path text … first detached into point text internally (detachPathTextToPointText), then converted through the same path
- With a multiple selection, each is converted individually

#### How the options behave

| | Size ratio | Button background (2 fills + rectangle effect) |
|---|---|---|
| On | Yes (e.g. W×1.2 / H×1.6) | Yes |
| Off | 1× | No |

- The button background adds two fills via New Fill and reshapes them with the rectangle shape effect. Fill colors are not set (the added fills keep their defaults).

### 更新履歴 / Changelog

- v1.3.0 (2026-07-02): 「スタイルの読み込み」機能（読み込み／再読み込みボタン、外部 AI ファイルからのスタイル取り込み、参照ファイルの記憶）、グラフィックスタイルの選択 UI（元の見た目／読み込んだスタイルのラジオ）、テキスト＋長方形の合体ロジック（長方形をフレーム化してテキストを流し込む処理）を撤去。ポイント文字・パス上文字を実寸計測でエリア内文字へ変換する処理に単純化。アピアランスの引き継ぎ（選択テキストの見た目を一時グラフィックスタイルとして登録・適用し、適用後に削除）は従来どおり維持 ／ Removed the "Load Styles" feature (Load/Reload buttons, importing styles from an external AI file, remembering the source file), the graphic-style selection UI (original/loaded-style radios), and the text + rectangle merge logic (turning a rectangle into the frame and flowing the text in). Simplified to converting point/path text into area type at the measured real size, while keeping the appearance inheritance as before (register the source text's look as a temp graphic style, apply it, then remove it)
- v1.2.0 (2026-07-02 追記): グラフィックスタイルのラジオ（元の見た目／読み込んだスタイル）を排他選択に修正（別コンテナのため自動排他が効いていなかった）。ダイアログで読み込み／再読み込みした後に変換すると選択が復帰されず無言で何も起きない問題を修正（選択復帰を変換直前に常時実行）。再読み込み後にどのラジオも未選択になる問題を修正（同名復元／なければ元の見た目へ）。変換が0件のとき無言終了せず、握り潰していた例外理由を添えて警告を表示 ／ Made the graphic-style radios (original appearance / loaded styles) mutually exclusive (they lived in separate containers, so ScriptUI's auto-exclusion didn't apply); fixed a silent no-op when converting after a Load/Reload (selection is now always restored right before converting); fixed a no-selection state after Reload (restore by name, else fall back to original); a 0-result conversion now alerts with the previously swallowed error instead of exiting silently
- v1.2.0 (2026-07-01): 固定パス（TARGET_FILE_PATH）と固定スタイル名（文字白抜き／枠のみ）を撤去。ダイアログに「スタイルの読み込み」ボタンを追加し、選んだ AI ファイルとスタイル名を Folder.userData に記憶。取り込んだスタイル名からラジオを自動生成する構成へ変更（ImportGraphicStyles v1.7.0 より移植）／ Removed the hardcoded path (TARGET_FILE_PATH) and fixed style names (white text / frame only); added a "Load Styles" button that remembers the picked AI file and style names in Folder.userData; radios are now generated automatically from the imported style names (ported from ImportGraphicStyles v1.7.0)
- v1.1.0 (2026-07-01): ダイアログにグラフィックスタイル選択（元の見た目／文字白抜き／枠のみ）を追加。外部 AI ファイル（TARGET_FILE_PATH）から未登録スタイルを取り込んで適用（ImportGraphicStyles より移植）。角丸オプションを削除／ Added a graphic-style choice (original appearance / white text / frame only); imports the chosen style from an external AI file when it is not yet registered (ported from ImportGraphicStyles). Removed the round-corners option
- v1.0.0 : 初期バージョン / Initial release

*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.3.0";

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
                ja: "ポイント文字・パス上文字を選択してください。",
                en: "Please select point text or path text."
            },
            noDocument: {
                ja: "ドキュメントが開かれていません。",
                en: "No document is open."
            },
            conversionFailed: {
                ja: "エリア内文字に変換できませんでした。",
                en: "Could not convert to area type."
            }
        },
        /* ダイアログUI / Dialog UI */
        ui: {
            dialogTitle: { ja: "エリア内文字に変換", en: "Convert to Area Type" },
            sizeAdjustPanel: { ja: "大きさ調整", en: "Size adjustment" },
            doAdjust: { ja: "する", en: "On" },
            dontAdjust: { ja: "しない", en: "Off" },
            widthRatio: { ja: "幅：", en: "Width:" },
            heightRatio: { ja: "高さ：", en: "Height:" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        }
    };

    /* ドット区切りキーから LABELS のエントリを取得 / Resolve a LABELS entry from a dot-separated key */
    function getLabelEntry(key) {
        var parts = key.split(".");
        var node = LABELS;
        for (var i = 0; i < parts.length; i++) {
            if (node && typeof node[parts[i]] !== "undefined") { node = node[parts[i]]; }
            else { return null; }
        }
        return node;
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
        var aiaString = "/version 3" +
            _nameBlockAscii(setName) +
            "/isOpen 1" +
            "/actionCount " + actionDefs.length;

        for (var i = 0; i < actionDefs.length; i++) {
            var actionDef = actionDefs[i];
            aiaString += "/action-" + (i + 1) + " {" +
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
        return aiaString;
    }

    /* フレーム整列アクションセット名 / Frame-alignment action set name */
    var AREA_TEXT_ACTION_SET = "AreaText";

    /* アクションセットを読み込む（temp に .aia を書き出して loadAction）/ Load an action set (write .aia to temp, then loadAction) */
    function loadActionSet(setName, aiaString) {
        try {
            // 既存があれば一旦外して衝突回避 / Unload existing set first to avoid conflicts
            try { app.unloadAction(setName, ""); } catch (e0) { }

            var actionFile = new File(Folder.temp + "/AreaTypeToolkit_action_" + setName + ".aia");
            actionFile.open("w");
            actionFile.write(aiaString);
            actionFile.close();

            app.loadAction(actionFile);
            try { actionFile.remove(); } catch (e1) { }
        } catch (e) { }
    }

    /* アクションセットを破棄 / Unload an action set */
    function unloadActionSet(setName) {
        try { app.unloadAction(setName, ""); } catch (e) { }
    }

    /* フレーム整列アクション（AlignTop/Center/Bottom/Justify）を読み込む / Load frame-alignment actions */
    function loadAreaTextActions() {
        var aiaString = _buildActionSetAIA(
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
        loadActionSet(AREA_TEXT_ACTION_SET, aiaString);
    }

    /* フレーム整列アクションを破棄 / Unload frame-alignment actions */
    function unloadAreaTextActions() {
        unloadActionSet(AREA_TEXT_ACTION_SET);
    }

    /* エリア内文字のフレーム整列（縦方向の配置）アクションを実行 / Run the area-text frame-alignment action (vertical placement)
       valueInt: 0=上/top, 1=中央/center, 2=下/bottom, 3=均等/justify */
    function runAreaTextFrameAlignmentAction(valueInt) {
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
            runAreaTextFrameAlignmentAction(valueInt);
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
            var itemBounds;
            try { itemBounds = sel[i].visibleBounds; } catch (e) { continue; }
            if (!itemBounds) continue;
            if (left === null || itemBounds[0] < left) left = itemBounds[0];
            if (top === null || itemBounds[1] > top) top = itemBounds[1];
            if (right === null || itemBounds[2] > right) right = itemBounds[2];
            if (bottom === null || itemBounds[3] < bottom) bottom = itemBounds[3];
        }
        if (left === null) return null;
        return [left, top, right, bottom];
    }

    /* 複製→アピアランス分割→テキストのアウトラインで正確な可視サイズを計測し、複製を破棄
       Measure accurate visible size via duplicate → expand appearance → create outlines, then discard the copy */
    function measureAccurateBounds(item) {
        var doc = app.activeDocument;
        var savedSel = doc.selection;
        var measured = null;
        var duplicatedItem = null;
        try {
            duplicatedItem = item.duplicate();
            doc.selection = null;
            duplicatedItem.selected = true;
            app.redraw();

            // アピアランスを分割 / Expand appearance
            try { app.executeMenuCommand('expandStyle'); } catch (e) { }
            // テキストのアウトライン / Create outlines
            try { app.executeMenuCommand('outline'); } catch (e) { }
            app.redraw();

            // 分割・アウトライン後の選択全体のサイズを計測 / Measure bounds of the resulting selection
            var resultSel = doc.selection;
            var bounds = getSelectionVisibleBounds(resultSel);
            if (bounds) {
                measured = { left: bounds[0], top: bounds[1], right: bounds[2], bottom: bounds[3], width: bounds[2] - bounds[0], height: bounds[1] - bounds[3] };
            }

            // 複製（分割・アウトライン結果）を削除 / Remove the duplicate (expanded/outlined result)
            for (var d = resultSel.length - 1; d >= 0; d--) {
                try { resultSel[d].remove(); } catch (e2) { }
            }
        } catch (e0) {
            if (duplicatedItem) { try { duplicatedItem.remove(); } catch (e3) { } }
        }

        try { doc.selection = savedSel; } catch (e4) { }
        return measured;
    }

    // =========================================
    // シェイプ効果 / Shape effect
    // =========================================

    /* ［形状に変換：長方形］ライブエフェクトの定義 / "Convert to Shape: Rectangle" live-effect definition */
    var RECTANGLE_SHAPE_EFFECT_XML = '<LiveEffect name="Adobe Shape Effects" isPre="1"><Dict data="U DisplayString Rectangle I Shape 0 R RelWidth 0 R RelHeight 0 R AbsWidth 0 R AbsHeight 0 R Absolute 0 R CornerRadius 9 "/></LiveEffect>';

    /* 塗りを2枚追加し、長方形のシェイプ効果でボタン状の背景にする（塗り色は設定しない）
       Add two fills and a rectangle shape effect for a button-like background (colors are left as-is) */
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

    /* ボタン風の拡大倍率 / Button-style expansion ratios */
    var BUTTON_WIDTH_RATIO = 1.2;   // 元の幅に対する倍率 / Ratio of original width
    var BUTTON_HEIGHT_RATIO = 1.6;  // 元の高さに対する倍率 / Ratio of original height

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

    /* 変換中に握り潰した最後の例外（0件時にアラートで理由を添える）/ Last swallowed exception during conversion (surfaced on a 0-result alert) */
    var lastConversionError = null;

    /* 選択内のテキストをエリア内文字へ変換 / Convert the selected text into area type */
    function convertSelectionToAreaType(doc, sel, options) {
        lastConversionError = null;
        sel = preprocessPathTextSelection(doc, sel);

        var selection = app.activeDocument.selection;
        if (!selection || selection.length === 0) { return; }

        // 変換元テキストの有無を判定 / Detect source text
        var hasSourceText = false;
        for (var i = 0; i < selection.length; i++) {
            var item = selection[i];
            if (item.typename === "TextFrame" && (item.kind === TextType.POINTTEXT || item.kind === TextType.PATHTEXT)) hasSourceText = true;
        }
        if (!hasSourceText) { return; }

        // 計測した実寸でフレーム化してエリア内文字へ / Frame at the measured real size, then convert to area type
        var created = convertPointTextToMeasuredArea(doc, selection, options);

        if (created.length > 0) {
            try { app.activeDocument.selection = created; app.redraw(); } catch (e) { }
        } else {
            // 1件も変換できなかった＝内部で例外を握り潰している。無言終了せず理由を伝える
            // Nothing converted = an exception was swallowed internally. Surface it instead of exiting silently
            var failMessage = L("alert.conversionFailed");
            if (lastConversionError) { failMessage += "\n" + lastConversionError; }
            alert(failMessage);
        }
    }

    /* 計測した実寸でフレームを作り、テキストを流し込んで中央配置 / Build a frame at the measured size, flow the text in, centered */
    function convertPointTextToMeasuredArea(doc, selection, options) {
        var created = [];

        // 大きさ調整が有効なら倍率、無効なら等倍 / Ratios when size adjustment is on, 1x when off
        var widthRatio = (options && options.adjust) ? options.widthRatio : 1;
        var heightRatio = (options && options.adjust) ? options.heightRatio : 1;

        for (var i = selection.length - 1; i >= 0; i--) {
            var sourceText = selection[i];
            if (!(sourceText.typename === "TextFrame" && sourceText.kind === TextType.POINTTEXT)) continue;
            // 元テキストの見た目を一時グラフィックスタイルとして登録（アピアランスを引き継ぐため）
            // Register the source text's appearance as a temp graphic style (to carry over its appearance)
            var tempStyleName = registerTextFrameAsTempGraphicStyle(sourceText);
            try {
                // 複製→アピアランス分割→アウトラインで正確な可視サイズを計測 / Measure accurate visible size
                var measuredBounds = measureAccurateBounds(sourceText) || geometricBoundsOf(sourceText);
                // 大きさ調整の倍率を掛け、中心を保ったままフレーム化 / Apply the size-adjustment ratios, keeping the center
                var frameWidth = measuredBounds.width * widthRatio;
                var frameHeight = measuredBounds.height * heightRatio;
                var centerX = measuredBounds.left + measuredBounds.width / 2;
                var centerY = measuredBounds.top - measuredBounds.height / 2;
                var frameLeft = centerX - frameWidth / 2;
                var frameTop = centerY + frameHeight / 2;
                var framePath = doc.pathItems.rectangle(frameTop, frameLeft, frameWidth, frameHeight);
                framePath.filled = false; framePath.stroked = false;
                var areaType = fillAreaTypeFromSourceText(doc, framePath, sourceText, tempStyleName);
                centerAreaTypeContents(areaType);
                // 「する」のときは塗り2枚＋長方形シェイプ効果でボタン状の背景を付与
                // When size adjustment is on, add a button-like background
                if (options && options.adjust) { applyButtonShapeAppearance(areaType); }
                created.push(areaType);
                sourceText.remove();
            } catch (e) {
                lastConversionError = e; // 失敗理由を保持（0件時にアラートで表示）/ Keep the reason (shown on a 0-result alert)
            } finally {
                // アピアランスを引き継いだら一時グラフィックスタイルは削除 / Remove the temp style once its appearance is carried over
                removeGraphicStyleByName(tempStyleName);
            }
        }

        return created;
    }

    // =========================================
    // ダイアログ / Dialog
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

    /* オプションダイアログ。OKなら{adjust, widthRatio, heightRatio}、キャンセルならnull
       Options dialog. Returns {adjust, widthRatio, heightRatio} on OK, null on Cancel */
    function showOptionsDialog() {
        var dialog = new Window("dialog", L("ui.dialogTitle") + " " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = "fill";
        dialog.margins = 15;

        var sizeAdjustPanel = dialog.add("panel", undefined, L("ui.sizeAdjustPanel"));
        setupPanel(sizeAdjustPanel, 6);

        // 大きさ調整 する / しない / Size adjustment on / off
        var adjustModeGroup = sizeAdjustPanel.add("group");
        var adjustOnRadio = adjustModeGroup.add("radiobutton", undefined, L("ui.doAdjust"));
        var adjustOffRadio = adjustModeGroup.add("radiobutton", undefined, L("ui.dontAdjust"));
        adjustOffRadio.value = true; // 既定は「しない」/ Default: off

        // 幅・高さの倍率（別々の行、百分率 % で入力）/ Width and height ratios (separate rows, entered as %)
        var widthRow = sizeAdjustPanel.add("group");
        var widthLabel = widthRow.add("statictext", undefined, L("ui.widthRatio"));
        widthLabel.preferredSize.width = 44;
        var widthInput = widthRow.add("edittext", undefined, String(Math.round(BUTTON_WIDTH_RATIO * 100)));
        widthInput.characters = 5;
        widthRow.add("statictext", undefined, "%");

        var heightRow = sizeAdjustPanel.add("group");
        var heightLabel = heightRow.add("statictext", undefined, L("ui.heightRatio"));
        heightLabel.preferredSize.width = 44;
        var heightInput = heightRow.add("edittext", undefined, String(Math.round(BUTTON_HEIGHT_RATIO * 100)));
        heightInput.characters = 5;
        heightRow.add("statictext", undefined, "%");

        function updateRatioInputsEnabled() {
            widthInput.enabled = adjustOnRadio.value;
            heightInput.enabled = adjustOnRadio.value;
        }
        adjustOnRadio.onClick = updateRatioInputsEnabled;
        adjustOffRadio.onClick = updateRatioInputsEnabled;
        updateRatioInputsEnabled();

        // ボタン（Mac規約：キャンセル → OK）/ Buttons (Mac order: Cancel → OK)
        var buttonGroup = dialog.add("group");
        buttonGroup.alignment = "right";
        var cancelButton = buttonGroup.add("button", undefined, L("ui.cancel"), { name: "cancel" });
        var okButton = buttonGroup.add("button", undefined, "OK", { name: "ok" });

        var dialogResult = null;
        okButton.onClick = function () {
            // 幅・高さは百分率 % 入力を倍率へ換算 / Width/height: convert the % input to a ratio
            var wPercent = parseFloat(widthInput.text);
            var hPercent = parseFloat(heightInput.text);
            dialogResult = {
                adjust: adjustOnRadio.value,
                widthRatio: (isNaN(wPercent) || wPercent <= 0) ? BUTTON_WIDTH_RATIO : wPercent / 100,
                heightRatio: (isNaN(hPercent) || hPercent <= 0) ? BUTTON_HEIGHT_RATIO : hPercent / 100
            };
            dialog.close();
        };
        cancelButton.onClick = function () { dialogResult = null; dialog.close(); };

        dialog.show();
        return dialogResult;
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
                // ダイアログを表示（キャンセルで中止）/ Show the dialog (Cancel aborts)
                var options = showOptionsDialog();
                if (options) {
                    // フレーム整列アクションを読み込み、終了時に破棄 / Load frame-alignment actions, unload on exit
                    loadAreaTextActions();
                    try {
                        convertSelectionToAreaType(doc, sel, options);
                    } finally {
                        unloadAreaTextActions();
                    }
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
