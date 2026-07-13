#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ポイント文字・パス上文字・図形・エリア内文字を対象に、エリア内文字の作成と調整を行うツール。

選択内容に応じて自動でエリア内文字へ変換し、調整ダイアログを開く。

- ポイント文字 / パス上文字のみ → ボタン風（幅×1.2・高さ×1.6）に変換
- テキスト＋図形 → 図形をエリア内文字にしてテキストを流し込む
- エリア内文字のみ → そのまま調整ダイアログへ

#### 調整ダイアログ

- フォントサイズの指定、「文字あふれ解消」で枠に収まる最大サイズへ自動縮小
- フレームサイズ（幅・高さ）の変更
- 左右インデント（連動可）と外側からの間隔
- 行揃え・テキストの配置は常に中央
- プレビューは常時ONで、変更を即座に反映

#### 補足

- 縦方向の中央配置にはフレーム整列アクションを使用する。実行時に一時的に読み込み、終了時に自動で破棄するため、アクションパネルに残骸は残らない。

### Overview

A tool to create and adjust area type from point text, path text, shapes, or existing area type.

Depending on the selection it converts to area type automatically and opens the adjust dialog.

- Point / path text only → convert to button style (width ×1.2, height ×1.6)
- Text + shape → fill the shape (as area type) with the text
- Area type only → open the adjust dialog directly

#### Adjust dialog

- Set font size, or auto-shrink to the largest fitting size via "Make overset"
- Change frame size (width / height)
- Left / right indent (linkable) and outer spacing
- Justification and vertical alignment are always centered
- Preview is always on and reflects changes instantly

#### Notes

- Vertical centering uses a frame-alignment action that is loaded temporarily and removed automatically on exit, so nothing is left behind in the Actions panel.

*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.0.0";

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* ボタン風の拡大倍率 / Button-style expansion ratios */
    var BUTTON_WIDTH_RATIO = 1.2;   // 元の幅に対する倍率 / Ratio of original width
    var BUTTON_HEIGHT_RATIO = 1.6;  // 元の高さに対する倍率 / Ratio of original height

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
        /* ダイアログ / Dialog */
        dialog: {
            title: { ja: "AreaType Toolkit", en: "AreaType Toolkit" }
        },
        /* パネル見出し / Panel titles */
        panel: {
            fontSize: { ja: "フォントサイズ", en: "Font size" },
            frameSize: { ja: "フレームサイズ", en: "Frame size" },
            indent: { ja: "インデント", en: "Indent" },
            options: { ja: "オプション", en: "Options" }
        },
        /* 入力ラベル / Field labels */
        label: {
            fontSize: { ja: "フォントサイズ", en: "Font size" },
            width: { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" }
        },
        /* チェックボックス / Checkboxes */
        checkbox: {
            indentLeft: { ja: "左", en: "Left" },
            indentRight: { ja: "右", en: "Right" },
            sync: { ja: "連動", en: "Link" },
            margin: { ja: "外側からの間隔", en: "Spacing" }
        },
        /* ボタン / Buttons */
        button: {
            overset: { ja: "文字あふれ解消", en: "Make overset" },
            run: { ja: "実行", en: "Run" },
            close: { ja: "閉じる", en: "Close" }
        },
        /* 警告メッセージ / Alerts */
        alert: {
            selectText: {
                ja: "ポイント文字・パス上文字・エリア内文字を選択してください。",
                en: "Please select point text, path text, or area text."
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

    /* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
    function labelText(key) {
        return L(key) + (currentLanguage === "ja" ? "：" : ":");
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* ルーラー単位に応じたラベルと pt 変換係数を返す / Return unit label and pt conversion factor */
    function getRulerUnitInfo(doc) {
        var ru = doc.rulerUnits;
        if (ru === RulerUnits.Millimeters) return { label: "mm", toPt: 72 / 25.4 };
        if (ru === RulerUnits.Centimeters) return { label: "cm", toPt: 72 / 2.54 };
        if (ru === RulerUnits.Inches) return { label: "in", toPt: 72 };
        if (ru === RulerUnits.Points) return { label: "pt", toPt: 1 };
        if (ru === RulerUnits.Picas) return { label: "pica", toPt: 12 };
        if (ru === RulerUnits.Pixels) return { label: "px", toPt: 72 / 96 };
        return { label: "pt", toPt: 1 };
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
                var ca = pathText.characters[c].characterAttributes;
                charAttrs.push({
                    font: ca.textFont,
                    size: ca.size,
                    fillColor: ca.fillColor,
                    strokeColor: ca.strokeColor,
                    strokeWeight: ca.strokeWeight,
                    autoLeading: ca.autoLeading,
                    leading: ca.leading
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
                var nc = new NoColor();
                newText.textRange.characterAttributes.strokeColor = nc;
                newText.textRange.characterAttributes.strokeWeight = 0;
            });

            // 文字ごとの属性を復元 / Restore per-character attributes
            var n = Math.min(newText.characters.length, charAttrs.length);
            for (var k = 0; k < n; k++) {
                var targetCa = newText.characters[k].characterAttributes;
                var srcCa = charAttrs[k];

                safe(function () { targetCa.textFont = srcCa.font; });
                safe(function () { targetCa.size = srcCa.size; });
                safe(function () { targetCa.fillColor = srcCa.fillColor; });
                safe(function () {
                    var sc = srcCa.strokeColor;
                    targetCa.strokeColor = sc;
                    targetCa.strokeWeight = (sc && sc.typename === "NoColor") ? 0 : srcCa.strokeWeight;
                });
                safe(function () { targetCa.baselineShift = 0; });
                safe(function () { targetCa.horizontalScale = 100; });
                safe(function () { targetCa.verticalScale = 100; });
                safe(function () { targetCa.autoLeading = srcCa.autoLeading; });
                if (!srcCa.autoLeading) safe(function () { targetCa.leading = srcCa.leading; });
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
            var it = sel[i];
            try {
                if (it && it.typename === "TextFrame" && it.kind === TextType.PATHTEXT) {
                    pathTexts.push(it);
                }
            } catch (e0) {
                // 無効オブジェクト（削除済み等）はスキップ / Skip invalid objects
            }
        }
        if (!pathTexts.length) return sel;

        var newTexts = detachPathTextToPointText(doc, pathTexts);
        if (!newTexts.length) return sel;

        // パス上文字を新ポイント文字に差し替えた新しい選択配列を構築 / Build replaced selection array
        var out = [];
        for (var j = 0; j < sel.length; j++) {
            var it2 = sel[j];
            try {
                if (it2 && it2.typename === "TextFrame" && it2.kind === TextType.PATHTEXT) {
                    // 旧オブジェクトは除外 / skip old
                } else if (it2) {
                    out.push(it2);
                }
            } catch (e1) {
                // 無効オブジェクトはスキップ / Skip invalid objects
            }
        }
        for (var k = 0; k < newTexts.length; k++) out.push(newTexts[k]);

        try { doc.selection = out; } catch (e) { }
        try { app.redraw(); } catch (e2) { }

        return out;
    }

    // =========================================
    // オーバーセット判定・文字サイズ調整 / Overset detection & font sizing
    // =========================================

    /* overflows プロパティを安全に取得 / Safely read the overflows property */
    function safeOverflows(tf) {
        try {
            if (tf && typeof tf.overflows !== "undefined") return !!tf.overflows;
        } catch (e) { }
        return null;
    }

    /* 指定行数までに全文字が収まらないか判定 / Whether text overflows within the given line count */
    function isOverset(tf, lineAmt) {
        if (tf.lines.length > 0) {
            var chars = 0;
            lineAmt = (typeof lineAmt === "undefined" || lineAmt === null) ? 1 : Math.floor(lineAmt);
            if (lineAmt < 1) lineAmt = 1;
            if (lineAmt > tf.lines.length) lineAmt = tf.lines.length;
            for (var i = 0; i < lineAmt; i++) { chars += tf.lines[i].characters.length; }
            return chars < tf.characters.length;
        }
        return tf.characters.length > 0;
    }

    /* テキストフレームがあふれているか判定 / Whether a text frame is overset */
    function isOversetFrame(tf) {
        if (tf && tf.kind === TextType.AREATEXT) {
            var ov = safeOverflows(tf);
            if (ov !== null) return ov;
            try { return isOverset(tf, tf.lines.length || 1); } catch (e) { return false; }
        }
        try { return isOverset(tf, tf.lines.length || 1); } catch (e) { return false; }
    }

    /* 行送り比率（行送り/サイズ）を取得 / Get leading ratio (leading / size) */
    function getLeadingInfo(tf) {
        try {
            var attrs = tf.textRange.characterAttributes;
            if (attrs.autoLeading) return null;
            var size = attrs.size, leading = attrs.leading;
            if (size > 0 && leading > 0) return { ratio: leading / size };
        } catch (e) { }
        return null;
    }

    /* 比率を保ったまま行送りを更新 / Update leading keeping the ratio */
    function applyLeading(tf, newSize, li) {
        if (!li) return;
        try { tf.textRange.characterAttributes.leading = newSize * li.ratio; } catch (e) { }
    }

    /* あふれなくなる最大サイズをバイナリサーチで探して縮小 / Shrink to the largest non-overset size */
    function shrinkFont(tf) {
        if (tf.characters.length <= 0 || !isOversetFrame(tf)) return;
        var li = getLeadingInfo(tf);
        var hi = tf.textRange.characterAttributes.size;
        var lo = 0.1;

        // 最小でもあふれるならそのまま終了 / If even lo overflows, keep min size
        tf.textRange.characterAttributes.size = lo;
        applyLeading(tf, lo, li);
        if (isOversetFrame(tf)) return;

        // バイナリサーチ / Binary search
        for (var iter = 0; iter < 40; iter++) {
            var mid = (lo + hi) / 2;
            tf.textRange.characterAttributes.size = mid;
            applyLeading(tf, mid, li);
            if (isOversetFrame(tf)) {
                hi = mid;
            } else {
                lo = mid;
            }
            if (hi - lo < 0.1) break;
        }

        // あふれない側（lo）に確定 / Settle on the non-overset side
        tf.textRange.characterAttributes.size = lo;
        applyLeading(tf, lo, li);
    }

    // =========================================
    // ダイナミックアクション / Dynamic actions
    // =========================================

    /* 文字列を16進数表現へ / Convert a string to hex */
    function _hexAscii(s) {
        var out = "";
        for (var i = 0; i < s.length; i++) {
            var h = s.charCodeAt(i).toString(16);
            if (h.length < 2) h = "0" + h;
            out += h;
        }
        return out;
    }

    /* /name ブロック（ASCII）を生成 / Build a /name block (ASCII) */
    function _nameBlockAscii(s) {
        return "/name [ " + s.length + " " + _hexAscii(s).toUpperCase() + " ]";
    }

    /* アクションセット定義(.aia)文字列を組み立て / Build an action set (.aia) definition string */
    function _buildActionSetAIA(setName, internalName, localizedNameHex, paramKeyInt, actionDefs) {
        var str = "/version 3" +
            _nameBlockAscii(setName) +
            "/isOpen 1" +
            "/actionCount " + actionDefs.length;

        for (var i = 0; i < actionDefs.length; i++) {
            var a = actionDefs[i];
            str += "/action-" + (i + 1) + " {" +
                " " + _nameBlockAscii(a.name) +
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
                " /value " + a.value +
                " }" +
                " }" +
                "}";
        }
        return str;
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

    /* テキストの配置（フレーム整列）を変更 / Change vertical alignment (frame alignment)
       valueInt: 0=上/top, 1=中央/center, 2=下/bottom, 3=均等/justify */
    function act_alignHorizontal(valueInt) {
        if (valueInt !== 0 && valueInt !== 1 && valueInt !== 2 && valueInt !== 3) return;

        var actionName = "AlignTop";
        if (valueInt === 1) actionName = "AlignCenter";
        else if (valueInt === 2) actionName = "AlignBottom";
        else if (valueInt === 3) actionName = "AlignJustify";

        try { app.doScript(actionName, AREA_TEXT_ACTION_SET, false); } catch (e) { }
    }

    /* 指定フレームにフレーム整列を適用（プレビュー中はスキップ）/ Apply frame alignment (skipped during preview) */
    function applyAreaTextFrameAlignment(tf, valueInt, forPreview) {
        // app.doScript はプレビュー中に呼ぶと不安定なためスキップ / Unstable during preview, so skip
        if (forPreview) return;
        try {
            var doc = app.activeDocument;
            doc.selection = null;
            doc.selection = [tf];
            app.redraw(); // 選択状態を確定 / Commit the selection
            act_alignHorizontal(valueInt);
        } catch (e) { }
    }

    // =========================================
    // UIユーティリティ / UI utilities
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

    /* グループの共通設定（row/column で整列を切り替え）/ Apply shared group layout (alignChildren switches by orientation) */
    function setupGroup(group, orientation, spacing) {
        var groupOrientation = orientation || "column";
        group.orientation = groupOrientation;
        /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
        group.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
        group.alignment = "fill";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* ↑↓キーで値を増減（Shift=10/Alt=0.1）/ Increment value with arrow keys (Shift=10, Alt=0.1) */
    function changeValueByArrowKey(editText, allowNegative, onChangeCallback) {
        editText.addEventListener("keydown", function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;

            if (keyboard.shiftKey) {
                var delta = 10;
                if (event.keyName == "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value = Math.floor((value - 1) / delta) * delta;
                    event.preventDefault();
                }
            } else if (keyboard.altKey) {
                if (event.keyName == "Up") {
                    value += 0.1;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= 0.1;
                    event.preventDefault();
                }
            } else {
                if (event.keyName == "Up") {
                    value += 1;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= 1;
                    event.preventDefault();
                }
            }

            value = keyboard.altKey ? Math.round(value * 10) / 10 : Math.round(value);
            if (!allowNegative && value < 0) value = 0;

            editText.text = value;
            if (typeof onChangeCallback === "function") { onChangeCallback(); }
        });
    }

    // =========================================
    // 変換 / Conversion
    // =========================================

    /* 選択からモードを自動判定して変換し、調整ダイアログを開く / Auto-detect mode, convert, then open the dialog */
    function convertToAreaTypeAndAdjust(doc, sel) {
        sel = preprocessPathTextSelection(doc, sel);

        var currentSel = app.activeDocument.selection;
        if (!currentSel || currentSel.length === 0) { return; }

        // 選択種別フラグ / Selection type flags
        var hasPointText = false, hasPathItem = false;
        for (var ti = 0; ti < currentSel.length; ti++) {
            if (currentSel[ti].typename === "TextFrame" && (currentSel[ti].kind === TextType.POINTTEXT || currentSel[ti].kind === TextType.PATHTEXT)) hasPointText = true;
            if (currentSel[ti].typename === "PathItem" || currentSel[ti].typename === "CompoundPathItem") hasPathItem = true;
        }

        // 選択内容に応じてモードを決定（ポイント文字は常にボタン風）/ Decide mode (point text is always button style)
        var useButton = false, useSelected = false;
        if (hasPointText && hasPathItem) {
            useSelected = true;
        } else if (hasPointText && !hasPathItem) {
            useButton = true;
        } else {
            return;
        }

        var created = runConvert(doc, currentSel, useButton, useSelected);
        if (created.length > 0) {
            try { app.activeDocument.selection = created; } catch (e) { }
            showDialogB(doc, created[0], created);
        }
    }

    /* モードに応じてエリア内文字を作成し、作成フレーム配列を返す / Create area type per mode; return created frames */
    function runConvert(doc, sel, useButton, useSelected) {
        var created = [];

        if (useSelected) {
            // 選択オブジェクトを利用：テキスト＋図形からエリア内文字を生成 / Text + shape → area type
            var srcText = null, destPath = null;
            for (var k = 0; k < sel.length; k++) {
                if (!srcText && sel[k].typename === "TextFrame") { srcText = sel[k]; }
                else if (!destPath && sel[k].typename === "PathItem" && sel[k].closed) { destPath = sel[k]; }
            }
            if (srcText && destPath) {
                try {
                    var srcContents = srcText.contents;
                    var srcFont = null, srcSize = 0;
                    try { srcFont = srcText.textRange.characterAttributes.textFont; srcSize = srcText.textRange.characterAttributes.size; } catch (e) { }
                    var dupPath = destPath.duplicate();
                    dupPath.filled = false; dupPath.stroked = false;
                    var newTf = doc.textFrames.areaText(dupPath);
                    newTf.contents = srcContents;
                    try { if (srcFont) newTf.textRange.characterAttributes.textFont = srcFont; if (srcSize > 0) newTf.textRange.characterAttributes.size = srcSize; } catch (e) { }
                    srcText.remove();
                    destPath.remove();
                    created.push(newTf);
                } catch (e) { }
            }
        } else if (useButton) {
            // ボタン風：幅×BUTTON_WIDTH_RATIO・高さ×BUTTON_HEIGHT_RATIO、中央揃え / Button style, centered
            for (var bt = sel.length - 1; bt >= 0; bt--) {
                var btItem = sel[bt];
                if (btItem.typename === "TextFrame" && btItem.kind === TextType.POINTTEXT) {
                    try {
                        var btB = btItem.geometricBounds;
                        var btOW = btB[2] - btB[0], btOH = btB[1] - btB[3];
                        var btW = btOW * BUTTON_WIDTH_RATIO, btH = btOH * BUTTON_HEIGHT_RATIO;
                        var btRect = doc.pathItems.rectangle(
                            btB[1] + (btH - btOH) / 2, btB[0] - (btW - btOW) / 2, btW, btH);
                        btRect.filled = false; btRect.stroked = false;
                        var btC = btItem.contents, btFont = null, btSize = 0;
                        try { btFont = btItem.textRange.characterAttributes.textFont; btSize = btItem.textRange.characterAttributes.size; } catch (e) { }
                        var btTf = doc.textFrames.areaText(btRect);
                        btTf.contents = btC;
                        try { if (btFont) btTf.textRange.characterAttributes.textFont = btFont; if (btSize > 0) btTf.textRange.characterAttributes.size = btSize; } catch (e) { }
                        // 行揃えを中央に / Horizontal center
                        try { btTf.textRange.paragraphAttributes.justification = Justification.CENTER; } catch (e) { }
                        // 縦位置はDOMで不安定なためアクションで中央 / Vertical center via dynamic action (unreliable via DOM)
                        try { applyAreaTextFrameAlignment(btTf, 1, false); } catch (e2) { }
                        created.push(btTf);
                        btItem.remove();
                    } catch (e) { }
                }
            }
        }

        if (created.length > 0) {
            app.activeDocument.selection = created;
            app.redraw();
        }
        return created;
    }

    // =========================================
    // 調整ダイアログ / Adjust dialog
    // =========================================

    /* エリア内文字の調整ダイアログを表示 / Show the area-type adjust dialog */
    function showDialogB(doc, initialTf, targetFrames) {
        var rulerInfo = getRulerUnitInfo(doc);

        // 受け取った変換結果を確実に対象にする / Make the passed frames the active target
        if (targetFrames && targetFrames.length) {
            try { app.activeDocument.selection = targetFrames; } catch (e) { }
        } else if (initialTf) {
            try { app.activeDocument.selection = [initialTf]; } catch (e2) { }
        }
        try { app.redraw(); } catch (e3) { }

        // モーダル中は selection が変動するため、渡された配列を優先して固定 / Pin targets (selection drifts in modal)
        var _fixedTargets = null;
        if (targetFrames && targetFrames.length) {
            _fixedTargets = targetFrames.slice(0);
        } else {
            try {
                var _sel0 = app.activeDocument.selection;
                if (_sel0 && _sel0.length) {
                    _fixedTargets = [];
                    for (var _k0 = 0; _k0 < _sel0.length; _k0++) { _fixedTargets.push(_sel0[_k0]); }
                }
            } catch (e4) { }
        }

        /* 現在の選択からエリア内文字を取り直して固定対象を更新 / Refresh pinned targets from selection */
        function refreshFixedTargetsFromSelection() {
            try {
                var s = app.activeDocument.selection;
                var arr = [];
                if (s && s.length) {
                    for (var i = 0; i < s.length; i++) {
                        if (s[i] && s[i].typename === "TextFrame" && s[i].kind === TextType.AREATEXT) {
                            arr.push(s[i]);
                        }
                    }
                }
                _fixedTargets = arr.length ? arr : null;
            } catch (e) {
                _fixedTargets = null;
            }
        }

        var dlgB = new Window("dialog", L("dialog.title") + " " + SCRIPT_VERSION);
        dlgB.alignChildren = "fill";
        dlgB.margins = 20;

        var grpLeft = dlgB.add("group");
        setupGroup(grpLeft, "column");

        /* フォントサイズ / Font size */
        var pnlAutoSize = grpLeft.add("panel", undefined, L("panel.fontSize"));
        setupPanel(pnlAutoSize);
        var grpFontSize = pnlAutoSize.add("group");
        grpFontSize.alignment = "left";
        grpFontSize.add("statictext", undefined, labelText("label.fontSize"));
        var etFontSize = grpFontSize.add("edittext", undefined, "");
        etFontSize.characters = 5;
        grpFontSize.add("statictext", undefined, "pt");
        var grpAutoSizeBtns = pnlAutoSize.add("group");
        grpAutoSizeBtns.orientation = "row";
        var btnTextSize = grpAutoSizeBtns.add("button", undefined, L("button.overset"));

        /* フレームサイズ / Frame size */
        var pnlFrameSize = grpLeft.add("panel", undefined, L("panel.frameSize"));
        setupPanel(pnlFrameSize);
        var grpWidth = pnlFrameSize.add("group");
        var lblWidth = grpWidth.add("statictext", undefined, labelText("label.width"));
        lblWidth.preferredSize.width = 44;
        var etWidth = grpWidth.add("edittext", undefined, "");
        etWidth.characters = 5;
        grpWidth.add("statictext", undefined, rulerInfo.label);
        var grpHeight = pnlFrameSize.add("group");
        var lblHeight = grpHeight.add("statictext", undefined, labelText("label.height"));
        lblHeight.preferredSize.width = 44;
        var etHeight = grpHeight.add("edittext", undefined, "");
        etHeight.characters = 5;
        grpHeight.add("statictext", undefined, rulerInfo.label);

        /* インデント / Indent */
        var pnlIndent = grpLeft.add("panel", undefined, L("panel.indent"));
        setupPanel(pnlIndent);
        // 連動チェックを右側に並べるため行方向へ上書き / Override to row so "Link" sits to the right
        pnlIndent.orientation = "row";
        pnlIndent.alignChildren = ["left", "top"];
        var grpIndentLeft = pnlIndent.add("group");
        grpIndentLeft.orientation = "column";
        grpIndentLeft.alignChildren = "left";
        var grpLeftIndent = grpIndentLeft.add("group");
        var chkLeftIndent = grpLeftIndent.add("checkbox", undefined, labelText("checkbox.indentLeft"));
        chkLeftIndent.preferredSize.width = 52;
        var etLeftIndent = grpLeftIndent.add("edittext", undefined, "0");
        etLeftIndent.characters = 4;
        grpLeftIndent.add("statictext", undefined, rulerInfo.label);
        etLeftIndent.enabled = false;
        var grpRightIndent = grpIndentLeft.add("group");
        var chkRightIndent = grpRightIndent.add("checkbox", undefined, labelText("checkbox.indentRight"));
        chkRightIndent.preferredSize.width = 52;
        var etRightIndent = grpRightIndent.add("edittext", undefined, "0");
        etRightIndent.characters = 4;
        grpRightIndent.add("statictext", undefined, rulerInfo.label);
        etRightIndent.enabled = false;
        var grpIndentRight = pnlIndent.add("group");
        grpIndentRight.orientation = "column";
        grpIndentRight.alignChildren = "left";
        grpIndentRight.alignment = ["left", "center"];
        var chkSync = grpIndentRight.add("checkbox", undefined, L("checkbox.sync"));

        /* オプション / Options */
        var pnlOptions = grpLeft.add("panel", undefined, L("panel.options"));
        setupPanel(pnlOptions);
        var grpMargin = pnlOptions.add("group");
        var chkMargin = grpMargin.add("checkbox", undefined, L("checkbox.margin"));
        var etMargin = grpMargin.add("edittext", undefined, "0");
        etMargin.characters = 6;
        var lblUnit = grpMargin.add("statictext", undefined, rulerInfo.label);
        etMargin.enabled = false;
        lblUnit.enabled = false;

        /* ボタンエリア / Button area */
        var grpBottom = dlgB.add("group");
        setupGroup(grpBottom, "row");
        grpBottom.alignChildren = ["right", "center"];
        var btnGroupB = grpBottom.add("group");
        btnGroupB.alignment = ["right", "center"];
        var btnCloseB = btnGroupB.add("button", undefined, L("button.close"), { name: "cancel" });
        var btnRun = btnGroupB.add("button", undefined, L("button.run"), { name: "ok" });

        // 状態変数 / State
        var isPreviewActive = false;
        var _autoSizeMode = "none";
        var hasMultiParagraph = false;
        var fontSize = 0;

        // 入力バリデーション用の最終正常値（ルーラー単位）/ Last valid values for validation (ruler units)
        var _lastValidWidth = null;
        var _lastValidHeight = null;

        /* 表示単位での上限値（極端値防止）/ Max size in ruler units (guards extreme values) */
        function _maxSizeInRulerUnits() {
            return 100000 / rulerInfo.toPt;
        }

        /* 幅/高さ入力を検証して有効値を返す（不正は最終正常値へ戻す）/ Validate width/height input */
        function validateSizeField(editText, lastValue) {
            var raw = String(editText.text);
            var v = parseFloat(raw);
            if (isNaN(v) || !isFinite(v)) {
                if (lastValue !== null) editText.text = lastValue;
                return null;
            }
            if (v <= 0) {
                if (lastValue !== null) editText.text = lastValue;
                return null;
            }
            var maxV = _maxSizeInRulerUnits();
            if (v > maxV) {
                v = maxV;
                editText.text = Math.round(v * 100) / 100;
            }
            if (v < 0.01) {
                v = 0.01;
                editText.text = Math.round(v * 100) / 100;
            }
            return v;
        }

        /* 対象フレームから現在値をUIに読み込む / Load current values from the frame into the UI */
        function loadValuesFromFrame(tf0) {
            try { hasMultiParagraph = (tf0.paragraphs && tf0.paragraphs.length >= 2); } catch (e) { hasMultiParagraph = false; }
            var initW = tf0.textPath.width / rulerInfo.toPt;
            var initH = tf0.textPath.height / rulerInfo.toPt;
            fontSize = 0;
            try { fontSize = tf0.textRange.characterAttributes.size || 0; } catch (e) { }
            if (fontSize > 0) { etFontSize.text = Math.round(fontSize * 100) / 100; }
            etWidth.text = Math.round(initW * 100) / 100;
            etHeight.text = Math.round(initH * 100) / 100;
            _lastValidWidth = parseFloat(etWidth.text);
            _lastValidHeight = parseFloat(etHeight.text);
            try {
                var sp = tf0.spacing || 0;
                etMargin.text = Math.round((sp / rulerInfo.toPt) * 100) / 100;
                chkMargin.value = (sp !== 0);
                etMargin.enabled = chkMargin.value;
                lblUnit.enabled = chkMargin.value;
            } catch (e) { }
            try {
                var ilp = tf0.paragraphs.length > 0 ? (tf0.paragraphs[0].leftIndent || 0) : 0;
                var irp = tf0.paragraphs.length > 0 ? (tf0.paragraphs[0].rightIndent || 0) : 0;
                chkLeftIndent.value = (ilp !== 0);
                etLeftIndent.enabled = chkLeftIndent.value;
                etLeftIndent.text = chkLeftIndent.value ? Math.round((ilp / rulerInfo.toPt) * 100) / 100 : "0";
                chkRightIndent.value = (irp !== 0);
                etRightIndent.enabled = chkRightIndent.value;
                etRightIndent.text = chkRightIndent.value ? Math.round((irp / rulerInfo.toPt) * 100) / 100 : "0";
            } catch (e) { }
            btnTextSize.enabled = !hasMultiParagraph;
        }

        /* UIの値をフレームへ適用（行揃え・配置は常に中央）/ Apply UI values (justify & alignment always centered) */
        function runAdjust(forPreview) {
            var doTextSize = (_autoSizeMode === "textsize");

            var justValue = Justification.CENTER;
            var alignValueInt = 1;

            var leftIndentPt = (chkLeftIndent.value || chkSync.value)
                ? (parseFloat(etLeftIndent.text) || 0) * rulerInfo.toPt : 0;
            var rightIndentPt = chkSync.value
                ? leftIndentPt
                : (chkRightIndent.value ? (parseFloat(etRightIndent.text) || 0) * rulerInfo.toPt : 0);
            var marginPt = chkMargin.value ? (parseFloat(etMargin.text) || 0) * rulerInfo.toPt : 0;

            var savedSel = [];
            var origSel = app.activeDocument.selection;
            for (var s = 0; s < origSel.length; s++) { savedSel.push(origSel[s]); }

            // 固定ターゲットを優先 / Prefer pinned targets
            if (!(_fixedTargets && _fixedTargets.length)) { refreshFixedTargetsFromSelection(); }
            var targets = _fixedTargets && _fixedTargets.length ? _fixedTargets : app.activeDocument.selection;

            for (var i = targets.length - 1; i >= 0; i--) {
                var obj = targets[i];
                if (obj.typename === "TextFrame" && obj.kind === TextType.AREATEXT) {
                    try { obj.spacing = marginPt; } catch (e) { }
                    // 幅/高さ：NaN・0以下・極端値をガード / Guard NaN, non-positive, extreme values
                    var wRu = validateSizeField(etWidth, _lastValidWidth);
                    var hRu = validateSizeField(etHeight, _lastValidHeight);
                    if (wRu !== null) { _lastValidWidth = wRu; try { obj.textPath.width = wRu * rulerInfo.toPt; } catch (e) { } }
                    if (hRu !== null) { _lastValidHeight = hRu; try { obj.textPath.height = hRu * rulerInfo.toPt; } catch (e) { } }
                    if (doTextSize) { shrinkFont(obj); }
                    try { var pa2 = obj.textRange.paragraphAttributes; pa2.justification = justValue; pa2.leftIndent = leftIndentPt; pa2.rightIndent = rightIndentPt; } catch (e) { }
                    applyAreaTextFrameAlignment(obj, alignValueInt, forPreview);
                }
            }
            if (savedSel.length > 0) { try { app.activeDocument.selection = savedSel; } catch (e) { } }
            app.redraw();
        }

        /* プレビューを更新（常時ON。直前のプレビューはundoで戻す）/ Update preview (always on; undo the previous one first) */
        function updatePreview() {
            if (isPreviewActive) {
                try { app.undo(); } catch (e) { }
                try { app.redraw(); } catch (e2) { }
                isPreviewActive = false;

                // undo 後は参照が無効化されるため対象を取り直す / Refresh targets after undo
                refreshFixedTargetsFromSelection();
            }
            if (!(_fixedTargets && _fixedTargets.length)) { refreshFixedTargetsFromSelection(); }
            runAdjust(true);
            isPreviewActive = true;
        }

        /* フォントサイズ入力を選択中のエリア内文字へ即時反映 / Apply font size field immediately */
        function applyFontSizeFromField() {
            var newSize = parseFloat(etFontSize.text) || 0;
            if (newSize <= 0) return;
            var tgts = app.activeDocument.selection;
            for (var i = 0; i < tgts.length; i++) {
                if (tgts[i].typename === "TextFrame" && tgts[i].kind === TextType.AREATEXT) {
                    try { tgts[i].textRange.characterAttributes.size = newSize; } catch (e) { }
                }
            }
            app.redraw();
        }

        /* 幅入力の検証＋プレビュー / Validate width then preview */
        function onWidthChange() {
            var wRu = validateSizeField(etWidth, _lastValidWidth);
            if (wRu !== null) { _lastValidWidth = wRu; }
            updatePreview();
        }

        /* 連動時に右インデントを同期してから幅変更扱い / Sync right indent (when linked) then treat as width change */
        function onAdjustmentChange() {
            if (chkSync.value) { etRightIndent.text = etLeftIndent.text; }
            onWidthChange();
        }

        // --- イベントハンドラ / Event handlers ---
        btnTextSize.onClick = function () { _autoSizeMode = "textsize"; runAdjust(false); _autoSizeMode = "none"; };
        chkMargin.onClick = function () {
            etMargin.enabled = chkMargin.value;
            lblUnit.enabled = chkMargin.value;
            if (chkMargin.value) { etMargin.text = "1"; }
            onAdjustmentChange();
        };
        chkSync.onClick = function () {
            if (chkSync.value) {
                chkLeftIndent.value = true; etLeftIndent.enabled = true;
                chkRightIndent.enabled = false; etRightIndent.enabled = false;
                etRightIndent.text = etLeftIndent.text;
            } else {
                chkRightIndent.enabled = true; etRightIndent.enabled = chkRightIndent.value;
            }
            onAdjustmentChange();
        };
        chkLeftIndent.onClick = function () {
            if (!chkSync.value) { etLeftIndent.enabled = chkLeftIndent.value; }
            if (!chkLeftIndent.value) { etLeftIndent.text = "0"; }
            onAdjustmentChange();
        };
        chkRightIndent.onClick = function () {
            etRightIndent.enabled = chkRightIndent.value;
            if (!chkRightIndent.value) { etRightIndent.text = "0"; }
            onAdjustmentChange();
        };

        etFontSize.onChange = applyFontSizeFromField;
        etMargin.onChange = onAdjustmentChange;
        etWidth.onChange = onWidthChange;
        etHeight.onChange = function () {
            var hRu = validateSizeField(etHeight, _lastValidHeight);
            if (hRu !== null) { _lastValidHeight = hRu; }
            updatePreview();
        };
        etLeftIndent.onChange = onAdjustmentChange;
        etRightIndent.onChange = onAdjustmentChange;
        changeValueByArrowKey(etFontSize, false, applyFontSizeFromField);
        changeValueByArrowKey(etMargin, false, onAdjustmentChange);
        changeValueByArrowKey(etWidth, false, onWidthChange);
        changeValueByArrowKey(etHeight, false, updatePreview);
        changeValueByArrowKey(etLeftIndent, false, onAdjustmentChange);
        changeValueByArrowKey(etRightIndent, false, onAdjustmentChange);

        btnRun.onClick = function () {
            if (isPreviewActive) { try { app.undo(); } catch (e) { } try { app.redraw(); } catch (e2) { } isPreviewActive = false; }
            runAdjust(false);
            dlgB.close(1);
        };
        btnCloseB.onClick = function () {
            if (isPreviewActive) { app.undo(); app.redraw(); }
            dlgB.close(0);
        };

        // 初期値読み込み / Load initial values
        if (initialTf) { loadValuesFromFrame(initialTf); }

        // 開いたらプレビュー実行（常時ON）/ Run preview on open (always on)
        updatePreview();

        dlgB.show();
    }

    // =========================================
    // エントリポイント / Entry point
    // =========================================
    if (app.documents.length > 0) {
        var doc = app.activeDocument;
        var sel = doc.selection;

        if (sel && sel.length > 0) {
            var _hasPoint = false, _hasArea = false, _hasPath = false;
            for (var _ei = 0; _ei < sel.length; _ei++) {
                if (sel[_ei].typename === "TextFrame") {
                    if (sel[_ei].kind === TextType.POINTTEXT || sel[_ei].kind === TextType.PATHTEXT) _hasPoint = true;
                    if (sel[_ei].kind === TextType.AREATEXT) _hasArea = true;
                }
                if (sel[_ei].typename === "PathItem" || sel[_ei].typename === "CompoundPathItem") {
                    _hasPath = true;
                }
            }

            if (_hasPoint || _hasPath || _hasArea) {
                // アクションを実行時に読み込み、終了時に破棄 / Load actions at start, unload on exit
                loadAreaTextActions();
                try {
                    if (_hasPoint || _hasPath) {
                        // ポイント文字 または 図形 → 変換（常にボタン風）してダイアログ / Point or shape → convert then dialog
                        convertToAreaTypeAndAdjust(doc, sel);
                    } else {
                        // エリア内文字のみ → 調整ダイアログ / Area type only → adjust dialog
                        var _firstArea = null;
                        for (var _fi = 0; _fi < sel.length; _fi++) {
                            if (sel[_fi].typename === "TextFrame" && sel[_fi].kind === TextType.AREATEXT) {
                                _firstArea = sel[_fi]; break;
                            }
                        }
                        showDialogB(doc, _firstArea, null);
                    }
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
