#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択に応じて、ポイント文字・パス上文字 ⇔ エリア内文字を相互変換する（順変換は見た目を保ったままエリア内文字へ、逆変換はエリア内文字を枠サイズを保ったままポイント文字へ）。ダイアログは順/逆共通（タイトル「テキストの変換」）で、現在の選択の内訳と「スタイル：保持する/保持しない」を表示し、ラジオの挙動はツールチップで補足する。

順変換（ポイント/パス文字 → エリア内文字）の流れ：

1. 変換ダイアログを表示（選択の内訳＋スタイル：保持する/保持しない）
2. スタイル「保持する」なら、変換前に選択テキストの見た目を一時グラフィックスタイルとして登録
3. ポイント文字 / パス上文字 → 複製→アピアランス分割→アウトラインで計測した実寸でエリア内文字へ変換（保持する/しない共通で、この計測サイズをフレームに使う）。「保持する」なら一時スタイルを適用してアピアランスを引き継ぐ（適用後に一時スタイルは削除）
4. 行揃え・テキストの配置は常に中央

逆変換（エリア内文字 → ポイント文字）の流れ：

1. 変換ダイアログを表示（選択の内訳＋スタイル：保持する/保持しない）
2. 元のエリア枠の実寸（幅A・高さB＝geometricBounds）を取得
3. ポイント文字へ変換（変換後の参照を取り直す）
4. スタイル「保持する」→ 何もしない（変換後の見た目のまま）／「保持しない」→ 塗りを2枚追加し、固定サイズ（AbsWidth=A / AbsHeight=B）の［形状に変換：長方形］効果でボタン背景を付与

#### 補足

- フレームサイズは、複製→アピアランス分割→アウトラインで計測した実寸を基準に決定する。
- 縦方向の中央配置とグラフィックスタイル登録にはダイナミックアクションを使用する。実行時に一時的に読み込み、終了時に自動で破棄するため、アクションパネルに残骸は残らない。
- 選択にエリア内文字が含まれていれば逆変換、含まれていなければ順変換を行う。
- 1件も変換できなかった場合は無言で終わらず警告を表示する。

#### 対応している選択

- ポイント文字 … そのまま計測してエリア内文字化
- パス上文字 … 内部で一度ポイント文字に分離（detachPathTextToPointText）してから同じ経路で処理
- 複数選択時はそれぞれ個別に変換

#### スタイル（保持する/保持しない）の効き方

| | 順変換（→エリア内文字） | 逆変換（→ポイント文字） |
|---|---|---|
| 保持する | 一時グラフィックスタイルで元の見た目を引き継ぐ | 何もしない（変換後の見た目のまま） |
| 保持しない | グラフィックスタイル適用なし（テキストのみ移す） | 塗り2枚＋固定サイズ A×B の長方形でボタン背景を付与 |

- ボタン背景（逆変換・保持しない）は New Fill を2枚追加し、固定サイズ（AbsWidth=A / AbsHeight=B）の長方形シェイプ効果で背景化する。塗り色は設定しない（追加した塗りは既定のまま）。
- エリア→ポイント変換直後のテキストは最初の New Fill が吸収されるため、実際には New Fill を3回呼ぶ（1回“空打ち”＋2枚）。毎回オブジェクトを選択し直してから実行する。

### Overview

Convert point/path text ⇔ area type depending on the selection, with a Keep/Don't-keep style choice for each direction. Both directions share one dialog (title "Convert Text") that shows a summary of the current selection and the "Style: Keep / Don't keep" choice, with tooltips clarifying each radio's behavior.

Forward flow (point/path text → area type):

1. Show the conversion dialog (selection summary + Style: Keep / Don't keep).
2. When "Keep", register the selected text's appearance as a temporary graphic style before converting.
3. Point / path text → area type at the size measured via duplicate → expand appearance → create outlines (same measurement for Keep / Don't keep). When "Keep", apply the temp style to carry over the appearance (removed afterwards).
4. Justification and vertical alignment are always centered.

Reverse flow (area type → point text):

1. Show the conversion dialog (selection summary + Style: Keep / Don't keep).
2. Read the original area frame size (width A / height B = geometricBounds).
3. Convert to point text (re-acquiring the reference afterwards).
4. "Keep" → leave as-is; "Don't keep" → add two fills + an absolute-size (AbsWidth=A / AbsHeight=B) "Convert to Shape: Rectangle" effect as a button background.

#### Notes

- Frame size is based on the real size measured via duplicate → expand appearance → create outlines.
- Vertical centering and graphic-style registration use dynamic actions that are loaded temporarily and removed automatically on exit, so nothing is left behind in the Actions panel.
- If the selection contains area type, the reverse conversion runs; otherwise the forward conversion runs.
- If nothing could be converted, the script alerts instead of exiting silently.

#### Supported selections

- Point text … measured directly and converted to area type
- Path text … first detached into point text internally (detachPathTextToPointText), then converted through the same path
- With a multiple selection, each is converted individually

#### How the Keep / Don't-keep style option behaves

| | Forward (→ area type) | Reverse (→ point text) |
|---|---|---|
| Keep | Carry over the original appearance via a temp graphic style | Do nothing (leave the converted look as-is) |
| Don't keep | Apply no graphic style (transfer text only) | Add two fills + an absolute A×B rectangle button background |

- The button background (reverse · Don't keep) adds two fills via New Fill and reshapes them with an absolute-size (AbsWidth=A / AbsHeight=B) rectangle shape effect. Fill colors are not set (the added fills keep their defaults).
- Right after area→point conversion the first New Fill is absorbed, so New Fill is actually invoked three times (one primer + two), re-selecting the object before each call.

### 更新履歴 / Changelog

- v1.0.0 (2026-07-02): 初期バージョン。選択に応じてポイント文字・パス上文字 ⇔ エリア内文字を相互変換する。共通ダイアログ「テキストの変換」で現在の選択内容とスタイル保持（保持する/保持しない）を選ぶ。順変換はアピアランス分割で計測した実寸でエリア内文字化（保持する＝一時グラフィックスタイルで見た目を引き継ぐ／保持しない＝文字のみ）。逆変換はポイント文字化（保持する＝そのまま／保持しない＝塗り2枚＋固定サイズの長方形背景）。縦中央配置とグラフィックスタイル登録はダイナミックアクションで一時的に処理 ／ Initial release. Converts point/path text ⇔ area type depending on the selection. A shared "Convert Text" dialog shows the current selection and a Keep/Don't-keep style choice. Forward: area type at the appearance-expanded measured size (Keep = carry over the look via a temp graphic style; Don't keep = text only). Reverse: point text (Keep = as-is; Don't keep = two fills + an absolute-size rectangle background). Vertical centering and graphic-style registration use temporary dynamic actions.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ConvertAreaAndPointType";      /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

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
                ja: "ポイント文字・パス上文字、またはエリア内文字を選択してください。",
                en: "Please select point text, path text, or area type."
            },
            noDocument: {
                ja: "ドキュメントが開かれていません。",
                en: "No document is open."
            },
            conversionFailed: {
                ja: "エリア内文字に変換できませんでした。",
                en: "Could not convert to area type."
            },
            reverseFailed: {
                ja: "ポイント文字に変換できませんでした。",
                en: "Could not convert to point text."
            }
        },
        /* ダイアログUI / Dialog UI */
        ui: {
            dialogTitle: { ja: "テキストの変換", en: "Convert Text" },
            selectionPanel: { ja: "選択オブジェクト", en: "Selection" },
            stylePanel: { ja: "スタイル", en: "Style" },
            keepStyle: { ja: "保持する", en: "Keep" },
            dontKeepStyle: { ja: "保持しない", en: "Don't keep" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        /* ツールチップ（順/逆で意味が異なるため補足）/ Tooltips (meaning differs per direction) */
        tooltip: {
            keepForward: { ja: "元テキストの見た目（塗り・線・効果）を引き継いでエリア内文字にします。", en: "Carry over the source text's appearance (fill/stroke/effects) into the area type." },
            dontKeepForward: { ja: "見た目を引き継がず、文字だけをエリア内文字にします。", en: "Transfer only the text (no appearance) into the area type." },
            keepReverse: { ja: "変換後の見た目のまま。追加処理はしません。", en: "Leave the converted look as-is; no extra processing." },
            dontKeepReverse: { ja: "塗りを2枚追加し、元の枠サイズの長方形背景（ボタン風）を付与します。", en: "Add two fills + a rectangle background at the original frame size (button-like)." }
        },
        /* テキスト種別 / Text-type names */
        textType: {
            pointText: { ja: "ポイント文字", en: "Point text" },
            areaText: { ja: "エリア内文字", en: "Area type" },
            pathText: { ja: "パス上文字", en: "Path text" },
            other: { ja: "テキスト以外", en: "Non-text" }
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

    /* fn を実行し、例外は握り潰す（失敗時 undefined）/ Run fn, swallowing errors (undefined on failure) */
    function tryQuietly(fn) { try { return fn(); } catch (e) { return undefined; } }

    /* パス上文字の文字ごとの属性を退避 / Snapshot per-character attributes of a path text */
    function snapshotCharAttrs(textFrame) {
        var charAttrs = [];
        for (var c = 0; c < textFrame.characters.length; c++) {
            var charAttr = textFrame.characters[c].characterAttributes;
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

    /* 退避した文字属性を新規テキストへ復元 / Restore snapshotted per-character attributes onto the new text */
    function restoreCharAttrs(newText, charAttrs) {
        var restoreCount = Math.min(newText.characters.length, charAttrs.length);
        for (var k = 0; k < restoreCount; k++) {
            var targetAttr = newText.characters[k].characterAttributes;
            var sourceAttr = charAttrs[k];

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

    /* パス上文字を字形を保ったままポイント文字へ分離 / Detach path text into point text keeping attributes */
    function detachPathTextToPointText(doc, pathTextFrames) {
        var created = [];
        if (!doc || !pathTextFrames || !pathTextFrames.length) return created;

        // 新規テキストだけ選べるよう選択を解除 / Clear selection
        tryQuietly(function () { doc.selection = null; });

        for (var j = pathTextFrames.length - 1; j >= 0; j--) {
            var pathText = pathTextFrames[j];
            if (!pathText || pathText.typename !== "TextFrame" || pathText.kind !== TextType.PATHTEXT) continue;

            var originalPath = null;
            tryQuietly(function () { originalPath = pathText.textPath; });
            if (!originalPath) continue;

            // 1) 文字属性・内容・行揃えを退避 / Snapshot char attributes, contents, justification
            var charAttrs = snapshotCharAttrs(pathText);

            var textContents = "";
            tryQuietly(function () { textContents = pathText.contents; });

            var justification = null;
            tryQuietly(function () {
                if (pathText.paragraphs && pathText.paragraphs.length > 0) {
                    justification = pathText.paragraphs[0].paragraphAttributes.justification;
                }
            });

            // 2) パス始点にポイント文字を新規作成 / Create new point text at path start anchor
            var newText = doc.textFrames.add();
            var anchorPoint = null;
            tryQuietly(function () {
                if (originalPath.pathPoints && originalPath.pathPoints.length > 0) {
                    anchorPoint = originalPath.pathPoints[0].anchor;
                }
            });
            if (anchorPoint) {
                newText.position = [anchorPoint[0], anchorPoint[1]];
            }

            newText.contents = textContents;

            if (justification !== null && newText.paragraphs && newText.paragraphs.length > 0) {
                tryQuietly(function () { newText.paragraphs[0].paragraphAttributes.justification = justification; });
            }

            // 既定の線を一旦消し、後で文字ごとに復元 / Clear default stroke, restore per-character later
            tryQuietly(function () {
                var noColor = new NoColor();
                newText.textRange.characterAttributes.strokeColor = noColor;
                newText.textRange.characterAttributes.strokeWeight = 0;
            });

            // 3) 文字属性を復元 / Restore per-character attributes
            restoreCharAttrs(newText, charAttrs);

            // 4) 元のパス上文字を削除（パスも一緒に消える）/ Remove original path text
            tryQuietly(function () { pathText.remove(); });

            // 5) 新規テキストを選択して返す / Select and return new text
            tryQuietly(function () { newText.selected = true; });
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

        try { doc.selection = replacedSelection; app.redraw(); } catch (e) { }

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

    /* フレーム整列（縦中央）アクションを読み込む / Load the frame-alignment (vertical center) action */
    function loadFrameAlignmentAction() {
        var aiaString = _buildActionSetAIA(
            AREA_TEXT_ACTION_SET,
            "adobe_frameAlignment",
            "39 e382a8e383aae382a2e58685e69687e5ad97e381aee38395e383ace383bce383a0e695b4e58897",
            1717660782,
            [{ name: "AlignCenter", value: 1 }]
        );
        loadActionSet(AREA_TEXT_ACTION_SET, aiaString);
    }

    /* フレーム整列アクションを破棄 / Unload the frame-alignment action */
    function unloadFrameAlignmentAction() {
        unloadActionSet(AREA_TEXT_ACTION_SET);
    }

    /* エリア内文字を縦方向中央に配置（DOMでは不安定なためアクション再生）/ Vertically center the area type via the action (unreliable via DOM) */
    function applyVerticalCenter(areaType) {
        try {
            var doc = app.activeDocument;
            doc.selection = null;
            doc.selection = [areaType];
            app.redraw(); // 選択状態を確定 / Commit the selection
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

    /* ［形状に変換：長方形］固定モードのライブエフェクト定義を組み立て（幅・高さは pt）
       Build the "Convert to Shape: Rectangle" live-effect definition in Absolute mode (width/height in pt) */
    function buildRectangleShapeEffectXML(width, height) {
        return '<LiveEffect name="Adobe Shape Effects" isPre="1"><Dict data="U DisplayString Rectangle I Shape 0 R RelWidth 0 R RelHeight 0 R AbsWidth ' + width + ' R AbsHeight ' + height + ' R Absolute 1 R CornerRadius 9 "/></LiveEffect>';
    }

    /* 対象を選択し直してから New Fill を1枚追加（毎回選択し直さないと2枚目が効かない）
       Add one New Fill after re-selecting the target (without re-selecting, the second call is a no-op) */
    function addNewFill(item) {
        app.activeDocument.selection = null;
        item.selected = true;
        app.redraw();
        app.executeMenuCommand('Adobe New Fill Shortcut');
        app.redraw();
    }

    /* 塗りを2枚追加し、固定サイズ（幅×高さ）の長方形シェイプ効果でボタン状の背景にする（塗り色は設定しない）
       エリア→ポイント変換直後のテキストは最初の New Fill が吸収されるため、1回“空打ち”してから2枚追加する（計3回）
       Add two fills + an absolute-size rectangle shape effect for a button-like background (colors are left as-is).
       Right after area→point conversion the first New Fill is absorbed, so prime once, then add two (3 calls total). */
    function applyButtonBackgroundSized(item, width, height) {
        try {
            addNewFill(item); // 空打ち（吸収される）/ primer (gets absorbed)
            addNewFill(item); // 塗り1 / fill 1
            addNewFill(item); // 塗り2 / fill 2
            app.activeDocument.selection = null;
            item.selected = true;
            app.redraw();
            item.applyEffect(buildRectangleShapeEffectXML(width, height));
        } catch (e) { }
    }

    // =========================================
    // 変換 / Conversion
    // =========================================

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
        applyVerticalCenter(areaType); // 縦位置は中央（アクション再生）/ Vertical center (via action)
    }

    /* 変換中に握り潰した最後の例外（0件時にアラートで理由を添える）/ Last swallowed exception during conversion (surfaced on a 0-result alert) */
    var lastConversionError = null;

    /* 変換結果を選択に反映。0件なら理由付きでアラート / Select the results, or alert (with reason) when nothing converted */
    function finishConversion(created, alertKey) {
        if (created.length > 0) {
            try { app.activeDocument.selection = created; app.redraw(); } catch (e) { }
        } else {
            // 1件も変換できなかった＝内部で例外を握り潰している。無言終了せず理由を伝える
            // Nothing converted = an exception was swallowed internally. Surface it instead of exiting silently
            var failMessage = L(alertKey);
            if (lastConversionError) { failMessage += "\n" + lastConversionError; }
            alert(failMessage);
        }
    }

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
        finishConversion(convertPointTextToMeasuredArea(doc, selection, options), "alert.conversionFailed");
    }

    /* 計測した実寸でフレームを作り、テキストを流し込んで中央配置 / Build a frame at the measured size, flow the text in, centered */
    function convertPointTextToMeasuredArea(doc, selection, options) {
        var created = [];

        // スタイル「保持する」なら見た目を一時スタイルで引き継ぐ（既定）/ Keep appearance via a temp style when "Keep" (default)
        var keepStyle = !options || options.keepStyle !== false;

        for (var i = selection.length - 1; i >= 0; i--) {
            var sourceText = selection[i];
            if (!(sourceText.typename === "TextFrame" && sourceText.kind === TextType.POINTTEXT)) continue;
            // 「保持する」のときだけ、元テキストの見た目を一時グラフィックスタイルとして登録
            // Register the source text's appearance as a temp graphic style only when "Keep"
            var tempStyleName = keepStyle ? registerTextFrameAsTempGraphicStyle(sourceText) : null;
            try {
                // 複製→アピアランス分割→アウトラインで正確な可視サイズを計測（保持する/しない共通）
                // Measure accurate visible size via expand appearance → outline (same for keep / don't keep)
                var measuredBounds = measureAccurateBounds(sourceText) || geometricBoundsOf(sourceText);
                // 計測した実寸そのままの長方形をフレームに / Frame at the measured size as-is
                var framePath = doc.pathItems.rectangle(measuredBounds.top, measuredBounds.left, measuredBounds.width, measuredBounds.height);
                framePath.filled = false; framePath.stroked = false;
                var areaType = fillAreaTypeFromSourceText(doc, framePath, sourceText, tempStyleName);
                centerAreaTypeContents(areaType);
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
    // 逆変換：エリア内文字 → ポイント文字 / Reverse: area type → point text
    // =========================================

    /* エリア内文字をポイント文字へ変換し、変換後のポイント文字参照を返す（無ければ null）
       Convert an area type to point text and return the resulting point-text frame (null if it fails).
       convertAreaObjectToPointObject() は「その場変換・戻り値 null」で、変換後は元のラッパー（areaType や選択）が
       stale になり kind を AREATEXT のまま報告することがある。そこで変換前に一時的な名前（マーカー）を付け、
       変換後に doc.textFrames を新規走査してその名前で確実に回収する
       (the API converts in place and returns null; afterwards the old wrappers can go stale and still report
        AREATEXT, so we tag the frame with a temporary name before converting and re-find it by that name
        via a fresh doc.textFrames scan) */
    function convertAreaObjectToPointText(doc, areaType) {
        var MARKER_NAME = "__AreaTypeToPointText_marker__";
        var previousName = "";
        try { previousName = areaType.name; } catch (eName) { }

        try {
            try { areaType.name = MARKER_NAME; } catch (eSet) { }
            // 対象だけを選択してから変換 / Select only the target, then convert
            doc.selection = null;
            areaType.selected = true;
            app.redraw();
            areaType.convertAreaObjectToPointObject();
        } catch (e) {
            lastConversionError = e; // 例外後も変換済みの場合があるので下で回収を試みる / May have converted anyway; try to recover below
        }

        // マーカー名で新規走査して回収（stale ラッパーに依存しない）/ Recover by marker name via a fresh scan (no stale wrappers)
        var pointText = null;
        var textFrames = doc.textFrames;
        for (var i = 0; i < textFrames.length; i++) {
            var frameName = "";
            try { frameName = textFrames[i].name; } catch (eGet) { continue; }
            if (frameName === MARKER_NAME) { pointText = textFrames[i]; break; }
        }

        if (pointText) {
            try { pointText.name = previousName; } catch (eRestore) { } // 元の名前へ戻す / Restore the original name
            return pointText;
        }
        return null;
    }

    /* 選択中のエリア内文字をポイント文字へ変換（保持しないときは元の枠サイズ A×B のボタン背景を付与）
       Convert selected area type into point text (adding an A×B button background when "Don't keep") */
    function convertAreaTypeToPointText(doc, selection, options) {
        var created = [];

        // スタイル「保持する」なら変換のみ、「保持しない」なら塗り＋長方形背景を付与 / "Keep" = convert only; "Don't keep" = add fill + rectangle background
        var keepStyle = !options || options.keepStyle !== false;

        // 変換で選択が変わるため、対象のエリア内文字を先に確保 / Collect targets first (conversion changes the selection)
        var areaTexts = [];
        for (var i = 0; i < selection.length; i++) {
            var item = selection[i];
            if (item.typename === "TextFrame" && item.kind === TextType.AREATEXT) areaTexts.push(item);
        }

        for (var j = areaTexts.length - 1; j >= 0; j--) {
            var areaType = areaTexts[j];
            try {
                // 1) 元のエリア枠の実寸 A×B（「保持しない」の背景サイズに使う）/ Original frame size A×B (used for the "Don't keep" background)
                var frameBounds = geometricBoundsOf(areaType);
                var frameWidth = frameBounds.width;   // A
                var frameHeight = frameBounds.height;  // B

                // 2) ポイント文字へ変換（変換後の参照を取り直す）/ Convert to point text (re-acquire the reference)
                var pointText = convertAreaObjectToPointText(doc, areaType);
                if (!pointText) continue;

                // 3) 「保持しない」→ 塗り2枚＋固定サイズ（A×B）の長方形でボタン背景を付与
                //    「保持する」→ 何もしない（変換後の見た目をそのまま）
                //    "Don't keep" → add two fills + an absolute A×B rectangle background; "Keep" → leave as-is
                if (!keepStyle) {
                    applyButtonBackgroundSized(pointText, frameWidth, frameHeight);
                }

                created.push(pointText);
            } catch (e) {
                lastConversionError = e; // 失敗理由を保持（0件時にアラートで表示）/ Keep the reason (shown on a 0-result alert)
            }
        }

        return created;
    }

    /* 選択（エリア内文字）をポイント文字へ変換 / Convert the selected area type into point text */
    function convertSelectionToPointText(doc, sel, options) {
        lastConversionError = null;
        finishConversion(convertAreaTypeToPointText(doc, sel, options), "alert.reverseFailed");
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

    /* テキストフレームの種類ラベルを返す（テキスト以外は null）/ Return the text-type label (null for non-text) */
    function describeTextType(item) {
        try {
            if (item.typename !== "TextFrame") return null;
            if (item.kind === TextType.POINTTEXT) return L("textType.pointText");
            if (item.kind === TextType.AREATEXT) return L("textType.areaText");
            if (item.kind === TextType.PATHTEXT) return L("textType.pathText");
        } catch (e) { }
        return null;
    }

    /* 選択の内訳を「種類 ×個数」で要約 / Summarize the selection as "type ×count" */
    function summarizeSelection(sel) {
        var counts = {}, order = [];
        for (var i = 0; i < sel.length; i++) {
            var label = describeTextType(sel[i]) || L("textType.other");
            if (!counts.hasOwnProperty(label)) { counts[label] = 0; order.push(label); }
            counts[label]++;
        }
        var parts = [];
        for (var k = 0; k < order.length; k++) { parts.push(order[k] + " ×" + counts[order[k]]); }
        return parts.join("、");
    }

    /* テキスト変換のオプションダイアログ（選択の内訳＋スタイル：保持する/保持しない）。OKなら{keepStyle}、キャンセルならnull
       isReverse で順/逆のツールチップを切替
       Text-conversion options dialog (selection summary + Style: keep / don't keep). Returns {keepStyle} on OK, null on Cancel.
       isReverse switches the forward/reverse tooltips */
    function showStyleDialog(sel, isReverse) {
        var dialog = new Window("dialog", L("ui.dialogTitle") + " " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = "fill";
        dialog.margins = 15;

        // 現在の選択オブジェクトの内訳を表示 / Show a summary of the current selection
        var selectionPanel = dialog.add("panel", undefined, L("ui.selectionPanel"));
        setupPanel(selectionPanel, 6);
        selectionPanel.add("statictext", undefined, summarizeSelection(sel));

        // スタイル：アピアランス（見た目）を保持するか / Style: keep the appearance or not
        var stylePanel = dialog.add("panel", undefined, L("ui.stylePanel"));
        setupPanel(stylePanel, 6);
        var styleChoiceGroup = stylePanel.add("group");
        var styleKeepRadio = styleChoiceGroup.add("radiobutton", undefined, L("ui.keepStyle"));
        var styleDontKeepRadio = styleChoiceGroup.add("radiobutton", undefined, L("ui.dontKeepStyle"));
        styleKeepRadio.value = true; // 既定は「保持する」/ Default: keep
        // 順/逆で意味が異なるためツールチップで補足 / Tooltips clarify the per-direction meaning
        styleKeepRadio.helpTip = L(isReverse ? "tooltip.keepReverse" : "tooltip.keepForward");
        styleDontKeepRadio.helpTip = L(isReverse ? "tooltip.dontKeepReverse" : "tooltip.dontKeepForward");

        // ボタン（左右中央、Mac規約：キャンセル → OK）/ Buttons (centered; Mac order: Cancel → OK)
        var buttonGroup = dialog.add("group");
        buttonGroup.alignment = "center";
        var cancelButton = buttonGroup.add("button", undefined, L("ui.cancel"), { name: "cancel" });
        var okButton = buttonGroup.add("button", undefined, "OK", { name: "ok" });

        var dialogResult = null;
        okButton.onClick = function () {
            dialogResult = { keepStyle: styleKeepRadio.value };
            dialog.close();
        };
        cancelButton.onClick = function () { dialogResult = null; dialog.close(); };

        // 表示直後にレイアウトを再計算して描画欠けを防ぐ / Recalculate layout on show to avoid partial rendering
        dialog.onShow = function () {
            dialog.layout.layout(true);
            dialog.layout.resize();
        };

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
            // 選択の種類を判定：エリア内文字なら逆変換、ポイント/パス文字なら順変換
            // Detect the selection: area type → reverse; point/path text → forward
            var hasAreaText = false, hasPointOrPathText = false;
            for (var i = 0; i < sel.length; i++) {
                if (sel[i].typename !== "TextFrame") continue;
                if (sel[i].kind === TextType.AREATEXT) hasAreaText = true;
                else if (sel[i].kind === TextType.POINTTEXT || sel[i].kind === TextType.PATHTEXT) hasPointOrPathText = true;
            }

            if (hasAreaText) {
                // 逆変換：ダイアログ（選択表示＋保持する/保持しない）→ ポイント文字へ変換
                // Reverse: dialog (selection summary + keep / don't keep) → convert to point text
                var reverseOptions = showStyleDialog(sel, true);
                if (reverseOptions) {
                    convertSelectionToPointText(doc, sel, reverseOptions);
                }
            } else if (hasPointOrPathText) {
                // 順変換：ダイアログを表示（キャンセルで中止）/ Forward: show the dialog (Cancel aborts)
                var options = showStyleDialog(sel, false);
                if (options) {
                    // フレーム整列アクションを読み込み、終了時に破棄 / Load the frame-alignment action, unload on exit
                    loadFrameAlignmentAction();
                    try {
                        convertSelectionToAreaType(doc, sel, options);
                    } finally {
                        unloadFrameAlignmentAction();
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
