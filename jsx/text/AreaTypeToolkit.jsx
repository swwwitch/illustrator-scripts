#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ポイント文字・パス上文字・図形・エリア内文字を対象に、エリア内文字の作成と調整を行うツール。

- ポイント文字 / パス上文字 → エリア内文字に変換（シンプル / ボタン風）
- 図形をエリア内文字に変換してダミーテキストを入力
- （テキスト＋図形）ポイント文字 / パス上文字＋図形からエリア内文字を生成
- エリア内文字のサイズ / 行揃え / フレームサイズ / インデント / 余白を調整
- テキストの配置（上 / 中央 / 下 / 均等）をダイナミックアクションで適用

ダイアログAで作成方法を選択し、ダイアログBで詳細を調整する。プレビューONで変更内容をリアルタイム反映。


*/

/*

### Overview

A toolkit that creates and adjusts Area Type from point text, path text, shapes, and existing area text.

- Convert point/path text to Area Type (Simple / Button style)
- Convert a shape to Area Type and fill it with dummy text
- Build Area Type from (text + shape) point/path text plus a shape
- Adjust size / justification / frame size / indent / spacing of Area Type
- Apply text placement (top / center / bottom / justify) via dynamic actions

Dialog A picks the creation method; Dialog B handles detailed adjustment. Preview reflects changes in real time.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AreaTypeToolkit";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-03-04";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 「選択オブジェクトにダミーテキスト」で流し込むダミー文字 / Dummy text for "Dummy text on selected object" */
    var DUMMY_TEXT_JA = "今日は15:00から予定外のミーティングがありました。疲れを癒すため、夕方に近くのカフェWAVEで、お気に入りの抹茶ラテを楽しみました。短い休憩でしたが、心が“ほっと”しました。";
    var DUMMY_TEXT_EN = "An unplanned meeting popped up at 3:00 PM today. To reset, I stopped by a nearby café for a matcha latte. It was a short break, but it helped me breathe and refocus.";

    /* ダミーテキストの優先フォント候補とサイズ / Preferred fonts and size for dummy text */
    var DUMMY_FONT_JA = ["HiraginoSans-W3", "Hiragino Sans W3"];
    var DUMMY_FONT_EN = ["MyriadPro-Regular", "Myriad Pro Regular", "MyriadPro", "Myriad"];
    var DUMMY_FONT_SIZE = 10;

    /* 候補名から使用可能なフォントを返す（無ければ先頭フォント）/ Return first available font from candidates (fallback: first font) */
    function getTextFontSafe(preferNames) {
        try {
            if (preferNames && preferNames.length) {
                for (var i = 0; i < preferNames.length; i++) {
                    try {
                        var font = app.textFonts.getByName(preferNames[i]);
                        if (font) return font;
                    } catch (e0) { }
                }
            }
            if (app.textFonts.length > 0) return app.textFonts[0];
        } catch (e) { }
        return null;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 現在の言語（ja / en）/ Current language (ja / en) */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    /* 日英ラベル定義（カテゴリ別）/ Japanese-English labels grouped by category */
    var LABELS = {
        dialog: {
            title: { ja: "AreaType Toolkit", en: "AreaType Toolkit" },
            convertTitle: { ja: "エリア内文字に変換", en: "Convert to Area Type" }
        },
        panel: {
            separate: { ja: "テキストを分離", en: "Separate text" },
            convert: { ja: "ポイント文字→エリア内文字", en: "Point Type → Area Type" },
            justify: { ja: "行揃え", en: "Justification" },
            textAlign: { ja: "テキストの配置", en: "Text alignment" },
            thread: { ja: "スレッドテキスト", en: "Thread text" },
            fontSize: { ja: "フォントサイズ", en: "Font size" },
            frameSize: { ja: "フレームサイズ", en: "Frame size" },
            indent: { ja: "インデント", en: "Indent" },
            options: { ja: "オプション", en: "Options" }
        },
        radio: {
            noAdjust: { ja: "調整しない", en: "No adjust" },
            strokeBlack: { ja: "1pt黒に", en: "Stroke 1pt black" },
            noPath: { ja: "パスをなし", en: "No path" },
            removePath: { ja: "パスを削除", en: "Remove path" },
            pointSimple: { ja: "シンプル", en: "Simple" },
            buttonStyle: { ja: "ボタン風", en: "Button style" },
            useSelected: { ja: "選択オブジェクトを利用", en: "Use selected object" },
            useSelectedDummy: { ja: "選択オブジェクトにダミーテキスト", en: "Dummy text on selected object" },
            justifyLeft: { ja: "左", en: "Left" },
            justifyCenter: { ja: "中央", en: "Center" },
            justifyRight: { ja: "右", en: "Right" },
            justifyFull: {
                ja: "均等配置（最終行左）",
                en: "Justify (last line left)"
            },
            justifyFullAll: {
                ja: "両端揃え（すべての行）",
                en: "Justify all lines"
            },
            alignTop: { ja: "上揃え", en: "Top" },
            alignCenter: { ja: "中央揃え", en: "Center" },
            alignBottom: { ja: "下揃え", en: "Bottom" },
            alignJustify: { ja: "均等配置", en: "Justify" }
        },
        checkbox: {
            sync: { ja: "連動", en: "Link" },
            margin: { ja: "外側からの間隔", en: "Spacing" },
            leader: { ja: "メニュー作成用（リーダー罫）", en: "Leader tabs" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        button: {
            convert: { ja: "変換", en: "Convert" },
            thread: { ja: "スレッドテキストに", en: "Make thread text" },
            overset: { ja: "文字あふれ解消", en: "Make overset" },
            fit: { ja: "フィット", en: "Fit" },
            frameAuto: { ja: "自動", en: "Auto" },
            run: { ja: "実行", en: "Run" },
            close: { ja: "閉じる", en: "Close" }
        },
        label: {
            fontSize: { ja: "フォントサイズ", en: "Font size" },
            width: { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" },
            chars: { ja: "文字", en: "chars" },
            indentLeft: { ja: "左", en: "Left" },
            indentRight: { ja: "右", en: "Right" }
        },
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

    /* ラベル取得（"category.key" 形式、{slash} は / に展開）/ Resolve "category.key" label, expand {slash} to / */
    function getLabel(key) {
        var parts = key.split(".");
        var node = LABELS;
        for (var i = 0; i < parts.length; i++) {
            if (!node) break;
            node = node[parts[i]];
        }
        var text = key;
        if (node) {
            if (typeof node[currentLanguage] === "string") text = node[currentLanguage];
            else if (typeof node.en === "string") text = node.en;
        }
        return text.replace(/\{slash\}/g, "/");
    }

    /* getLabel の別名 / Short alias for getLabel */
    function L(key) {
        return getLabel(key);
    }

    /* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
    function labelText(key) {
        return getLabel(key) + (currentLanguage === "ja" ? "：" : ":");
    }

    /* 件数付きラベル（日本語は全角括弧、英語は半角括弧）/ Label with count (full-width JA parentheses, half-width EN parentheses) */
    function labelWithCount(key, count) {
        if (currentLanguage === "ja") {
            return getLabel(key) + "（" + count + "）";
        }
        return getLabel(key) + " (" + count + ")";
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* ルーラー単位に応じたラベルと pt 変換係数を返す / Ruler unit label and pt conversion factor */
    function getRulerUnitInfo(doc) {
        var rulerUnit = doc.rulerUnits;
        if (rulerUnit === RulerUnits.Millimeters) return { label: "mm", toPt: 72 / 25.4 };
        if (rulerUnit === RulerUnits.Centimeters) return { label: "cm", toPt: 72 / 2.54 };
        if (rulerUnit === RulerUnits.Inches) return { label: "in", toPt: 72 };
        if (rulerUnit === RulerUnits.Points) return { label: "pt", toPt: 1 };
        if (rulerUnit === RulerUnits.Picas) return { label: "pica", toPt: 12 };
        if (rulerUnit === RulerUnits.Pixels) return { label: "px", toPt: 72 / 96 };
        return { label: "pt", toPt: 1 };
    }

    // =========================================
    // パネルレイアウト / Panel layout
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


    /* パス上文字 → ポイント文字（ダイアログA前処理）/ Path text → point text (Dialog A preprocess) */
    function detachPathTextToPointText(doc, pathTextFrames) {
        function safe(fn) { try { return fn(); } catch (e) { return undefined; } }

        var created = [];
        if (!doc || !pathTextFrames || !pathTextFrames.length) return created;

        // Clear selection so only newly created texts can be selected if needed
        safe(function () { doc.selection = null; });

        for (var j = pathTextFrames.length - 1; j >= 0; j--) {
            var pathText = pathTextFrames[j];
            if (!pathText || pathText.typename !== "TextFrame" || pathText.kind !== TextType.PATHTEXT) continue;

            var originalPath = null;
            safe(function () { originalPath = pathText.textPath; });
            if (!originalPath) continue;

            // 1) Snapshot per-character attributes
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

            // 2) Create new Point Text at path start anchor
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

            // Clear default stroke first; restore per-character later
            safe(function () {
                var nc = new NoColor();
                newText.textRange.characterAttributes.strokeColor = nc;
                newText.textRange.characterAttributes.strokeWeight = 0;
            });

            // Restore per-character attributes
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
                // safe(function () { targetCa.tracking = srcCa.tracking; }); // tracking not restored
                safe(function () { targetCa.baselineShift = 0; });
                safe(function () { targetCa.horizontalScale = 100; });
                safe(function () { targetCa.verticalScale = 100; });
                safe(function () { targetCa.autoLeading = srcCa.autoLeading; });
                if (!srcCa.autoLeading) safe(function () { targetCa.leading = srcCa.leading; });
            }

            // 3) Remove original Path Text (its path will be removed with it)
            safe(function () { pathText.remove(); });

            // 4) Select and return new text
            safe(function () { newText.selected = true; });
            created.push(newText);
        }

        return created;
    }

    function preprocessSelectionForDialogA(doc, sel) {
        // Convert selected Path Text items to Point Text so Dialog A can treat them like Point Type.
        if (!doc || !sel || !sel.length) return sel;

        var pathTexts = [];
        for (var i = 0; i < sel.length; i++) {
            var it = sel[i];
            try {
                if (it && it.typename === "TextFrame" && it.kind === TextType.PATHTEXT) {
                    pathTexts.push(it);
                }
            } catch (e0) {
                // invalid object (e.g., already removed) — skip
            }
        }
        if (!pathTexts.length) return sel;

        var newTexts = detachPathTextToPointText(doc, pathTexts);
        if (!newTexts.length) return sel;

        // Build a new selection array: replace pathTexts with new point texts, keep other items.
        var out = [];
        for (var j = 0; j < sel.length; j++) {
            var it2 = sel[j];
            try {
                if (it2 && it2.typename === "TextFrame" && it2.kind === TextType.PATHTEXT) {
                    // skip old
                } else if (it2) {
                    out.push(it2);
                }
            } catch (e1) {
                // invalid object — skip
            }
        }
        for (var k = 0; k < newTexts.length; k++) out.push(newTexts[k]);

        try { doc.selection = out; } catch (e) { }
        try { app.redraw(); } catch (e2) { }

        return out;
    }


    /* 自動サイズ調整ヘルパー（AutoFitTextFrame.jsx 参考）/ Auto-fit helpers (based on AutoFitTextFrame.jsx) */

    function safeOverflows(tf) {
        try {
            if (tf && typeof tf.overflows !== "undefined") return !!tf.overflows;
        } catch (e) { }
        return null;
    }

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

    function isOversetFrame(tf) {
        if (tf && tf.kind === TextType.AREATEXT) {
            var ov = safeOverflows(tf);
            if (ov !== null) return ov;
            try { return isOverset(tf, tf.lines.length || 1); } catch (e) { return false; }
        }
        try { return isOverset(tf, tf.lines.length || 1); } catch (e) { return false; }
    }

    function getLeadingInfo(tf) {
        try {
            var attrs = tf.textRange.characterAttributes;
            if (attrs.autoLeading) return null;
            var size = attrs.size, leading = attrs.leading;
            if (size > 0 && leading > 0) return { ratio: leading / size };
        } catch (e) { }
        return null;
    }

    function applyLeading(tf, newSize, leadingInfo) {
        if (!leadingInfo) return;
        try { tf.textRange.characterAttributes.leading = newSize * leadingInfo.ratio; } catch (e) { }
    }


    // =========================================
    // ダイナミックアクション / Dynamic actions
    //   スクリプト実行時に読み込み、終了時にアンロードする
    //   Loaded at startup, unloaded on exit
    // =========================================

    /* 文字列を ASCII 16進に変換 / Convert a string to ASCII hex */
    function asciiToHex(s) {
        var out = "";
        for (var i = 0; i < s.length; i++) {
            var h = s.charCodeAt(i).toString(16);
            if (h.length < 2) h = "0" + h;
            out += h;
        }
        return out;
    }

    /* アクション名ブロック /name [ <len> <hex> ] を生成 / Build the /name [ <len> <hex> ] block */
    function actionNameBlock(s) {
        return "/name [ " + s.length + " " + asciiToHex(s).toUpperCase() + " ]";
    }

    /* アクションセット定義（.aia 文字列）を組み立てる / Build an action set definition (.aia string) */
    function buildActionSetAia(setName, internalName, localizedNameHex, paramKeyInt, actionDefs) {
        // actionDefs: [{name: "...", value: <int>}, ...]
        var str = "/version 3" +
            actionNameBlock(setName) +
            "/isOpen 1" +
            "/actionCount " + actionDefs.length;

        for (var i = 0; i < actionDefs.length; i++) {
            var a = actionDefs[i];
            str += "/action-" + (i + 1) + " {" +
                " " + actionNameBlock(a.name) +
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

    /* .aia 文字列を一時ファイル経由で読み込む（既存があれば先に外す）/ Load an .aia string via a temp file (unload any existing set first) */
    function loadActionSet(setName, aiaString) {
        try { app.unloadAction(setName, ""); } catch (e0) { }
        var f = new File(Folder.temp + "/AreaTypeToolkit_action_" + setName + ".aia");
        f.open("w");
        f.write(aiaString);
        f.close();
        app.loadAction(f);
        try { f.remove(); } catch (e1) { }
    }

    /* 使用するアクションセットを読み込む（スクリプト開始時に1回）/ Load the action sets used (once at startup) */
    function loadDynamicActions() {
        // AreaType: AutoSizeOn(1) / AutoSizeOff(2)
        loadActionSet("AreaType", buildActionSetAia(
            "AreaType",
            "adobe_SLOAreaTextDialog",
            "33 e382a8e383aae382a2e58685e69687e5ad97e382aae38397e382b7e383a7e383b3",
            1952539754,
            [
                { name: "AutoSizeOn", value: 1 },
                { name: "AutoSizeOff", value: 2 }
            ]
        ));
        // AreaText: AlignTop(0) / AlignCenter(1) / AlignBottom(2) / AlignJustify(3)
        loadActionSet("AreaText", buildActionSetAia(
            "AreaText",
            "adobe_frameAlignment",
            "39 e382a8e383aae382a2e58685e69687e5ad97e381aee38395e383ace383bce383a0e695b4e58897",
            1717660782,
            [
                { name: "AlignTop", value: 0 },
                { name: "AlignCenter", value: 1 },
                { name: "AlignBottom", value: 2 },
                { name: "AlignJustify", value: 3 }
            ]
        ));
    }

    /* 読み込んだアクションセットを破棄する（スクリプト終了時）/ Discard the loaded action sets (on exit) */
    function unloadDynamicActions() {
        try { app.unloadAction("AreaType", ""); } catch (e) { }
        try { app.unloadAction("AreaText", ""); } catch (e) { }
    }

    /* フレームサイズ：自動サイズ調整を ON にして文字に合わせて広げる / Frame size: turn auto-size on to expand to fit text */
    function expandFrameToFit(tf) {
        app.activeDocument.selection = [tf];
        applyAutoSizeAction(1);
    }

    /* フレームサイズ：自動サイズ調整を OFF にする / Frame size: turn auto-size off */
    function collapseFrameAuto(tf) {
        app.activeDocument.selection = [tf];
        applyAutoSizeAction(2);
    }

    /* 自動サイズ調整アクションを実行（1=ON, 2=OFF）/ Run the auto-size action (1=ON, 2=OFF) */
    function applyAutoSizeAction(valueInt) {
        if (valueInt !== 1 && valueInt !== 2) return;
        var actionName = (valueInt === 1) ? "AutoSizeOn" : "AutoSizeOff";
        try { app.doScript(actionName, "AreaType", false); } catch (e) { }
    }

    /* テキストの配置アクションを実行（0=上, 1=中央, 2=下, 3=均等）/ Run the frame-alignment action (0=top, 1=center, 2=bottom, 3=justify) */
    function applyFrameAlignmentAction(valueInt) {
        if (valueInt !== 0 && valueInt !== 1 && valueInt !== 2 && valueInt !== 3) return;
        var actionName = "AlignTop";
        if (valueInt === 1) actionName = "AlignCenter";
        else if (valueInt === 2) actionName = "AlignBottom";
        else if (valueInt === 3) actionName = "AlignJustify";
        try { app.doScript(actionName, "AreaText", false); } catch (e) { }
    }

    function applyAreaTextFrameAlignment(tf, valueInt, forPreview) {
        // app.doScript はプレビュー中にダイアログから呼ぶと不安定なためスキップ
        if (forPreview) return;
        try {
            var doc = app.activeDocument;
            doc.selection = null;
            doc.selection = [tf];
            app.redraw(); // Illustratorに選択状態を確定させる
            applyFrameAlignmentAction(valueInt);
        } catch (e) { }
    }

    // ちぢむ処理（バイナリサーチ：最大 40 回でオーバーセットにならない最大サイズを探す）
    function shrinkFont(tf) {
        if (tf.characters.length <= 0 || !isOversetFrame(tf)) return;
        var leadingInfo = getLeadingInfo(tf);
        var hi = tf.textRange.characterAttributes.size;
        var lo = 0.1;

        // lo でもオーバーセットなら最小サイズのまま終了
        tf.textRange.characterAttributes.size = lo;
        applyLeading(tf, lo, leadingInfo);
        if (isOversetFrame(tf)) return;

        // バイナリサーチ
        for (var iter = 0; iter < 40; iter++) {
            var mid = (lo + hi) / 2;
            tf.textRange.characterAttributes.size = mid;
            applyLeading(tf, mid, leadingInfo);
            if (isOversetFrame(tf)) {
                hi = mid;
            } else {
                lo = mid;
            }
            if (hi - lo < 0.1) break;
        }

        // オーバーセットにならない側（lo）に確定
        tf.textRange.characterAttributes.size = lo;
        applyLeading(tf, lo, leadingInfo);
    }

    // テキスト（フィット）：オーバーセットになるまで拡大 → ちぢむ（fitFont 相当）
    function fitTextToFrame(tf) {
        if (tf.characters.length <= 0) return;
        var leadingInfo = getLeadingInfo(tf);
        var original = tf.textRange.characterAttributes.size;

        // Step 1: オーバーセットが出るまで 2 倍ずつ拡大
        if (!isOversetFrame(tf)) {
            var high = original, guard = 0;
            while (!isOversetFrame(tf) && guard < 25) {
                guard++;
                high = high * 2;
                if (high > 100000) break;
                try { tf.textRange.characterAttributes.size = high; applyLeading(tf, high, leadingInfo); } catch (e) { break; }
            }
        }

        // オーバーセットが出なければ元に戻して終了
        if (!isOversetFrame(tf)) {
            try { tf.textRange.characterAttributes.size = original; applyLeading(tf, original, leadingInfo); } catch (e) { }
            return;
        }

        // Step 2: ちぢむ処理
        shrinkFont(tf);
    }

    // ↑↓キーで値を増減（プレビュー更新コールバックつき）
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

    // メニュー作成用リーダー罫：各段落に右揃えタブ（リーダー「…」400pt）を適用
    function applyLeaderTab(tf) {
        try {
            var tabStop = new TabStopInfo();
            tabStop.position = 400;
            tabStop.alignment = TabStopAlignment.Right;
            tabStop.leader = "…";
            for (var p = 0; p < tf.paragraphs.length; p++) {
                var pa = tf.paragraphs[p].paragraphAttributes;
                pa.tabStops = [tabStop];
                pa.justification = Justification.RIGHT;
            }
        } catch (e) { }
    }

    // ============================================================
    // ダイアログA：ポイント文字 → エリア内文字 変換
    // ============================================================
    function showDialogA(doc, sel) {
        sel = preprocessSelectionForDialogA(doc, sel);
        var dlgA = new Window("dialog", L("dialog.convertTitle") + " " + SCRIPT_VERSION);
        dlgA.alignChildren = "fill";
        dlgA.margins = 20;

        // 作成方法のラジオボタン（パネルなし）/ Creation-method radios (no panel)
        var convertOptionsColumn = dlgA.add("group");
        setupGroup(convertOptionsColumn, "column");

        var radPointType = convertOptionsColumn.add("radiobutton", undefined, L("radio.pointSimple"));
        var radAreaType = convertOptionsColumn.add("radiobutton", undefined, L("radio.buttonStyle"));
        var radUseSelected = convertOptionsColumn.add("radiobutton", undefined, L("radio.useSelected"));
        var radUseSelectedDummy = convertOptionsColumn.add("radiobutton", undefined, L("radio.useSelectedDummy"));
        radPointType.value = true;

        // ボタンエリア：左から［キャンセル］［変換］
        var dialogAButtonRow = dlgA.add("group");
        dialogAButtonRow.orientation = "row";
        dialogAButtonRow.alignment = ["center", "center"]; // center buttons
        dialogAButtonRow.spacing = 10;
        var btnCloseA = dialogAButtonRow.add("button", undefined, L("button.close"), { name: "cancel" });
        var btnConvert = dialogAButtonRow.add("button", undefined, L("button.convert"), { name: "ok" });

        // 選択種別フラグ
        var hasPointText = false, hasPathItem = false;
        for (var ti = 0; ti < sel.length; ti++) {
            if (sel[ti].typename === "TextFrame" && (sel[ti].kind === TextType.POINTTEXT || sel[ti].kind === TextType.PATHTEXT)) hasPointText = true;
            if (sel[ti].typename === "PathItem" || sel[ti].typename === "CompoundPathItem") hasPathItem = true;
        }

        // 選択内容に応じてモードを誘導
        if (hasPointText && hasPathItem) {
            // テキスト＋図形 → シンプル・ボタン風をディム、選択オブジェクトを利用を自動選択
            radPointType.enabled = false;
            radAreaType.enabled = false;
            radUseSelected.value = true; // default, dummy is also available
        } else if (hasPointText && !hasPathItem) {
            // ポイント文字のみ → 選択オブジェクト系をディム
            radUseSelected.enabled = false;
            radUseSelectedDummy.enabled = false;
        } else if (!hasPointText && hasPathItem) {
            // 図形のみ → 「選択オブジェクトにダミーテキスト」を自動選択し、残りをディム
            radPointType.enabled = false;
            radAreaType.enabled = false;
            radUseSelected.enabled = false;
            radUseSelectedDummy.enabled = true;
            radUseSelectedDummy.value = true;
        }

        // ラジオ排他制御
        var allConvertRadios = [radPointType, radAreaType, radUseSelected, radUseSelectedDummy];
        for (var cr = 0; cr < allConvertRadios.length; cr++) {
            (function (btn) {
                btn.onClick = function () {
                    for (var j = 0; j < allConvertRadios.length; j++) {
                        allConvertRadios[j].value = (allConvertRadios[j] === btn);
                    }
                };
            })(allConvertRadios[cr]);
        }

        // 変換後にダイアログBを開くための変数
        var convertedFrames = [];
        var shouldOpenB = false;
        var _openBDefaultAlignInt = null; // 0=Top,1=Center,2=Bottom,3=Justify

        btnConvert.onClick = function () {
            var currentSel = app.activeDocument.selection;
            if (!currentSel || currentSel.length === 0) { return; }
            var useSel = !!radUseSelected.value;
            var useSelDummy = !!radUseSelectedDummy.value;
            var created = runConvert(doc, currentSel,
                radPointType.value, radAreaType.value, useSel, useSelDummy);
            if (created.length > 0) {
                convertedFrames = created;
                // If Button style was used, default Dialog B alignment to Center
                _openBDefaultAlignInt = radAreaType.value ? 1 : null;
                shouldOpenB = true;
                try { app.activeDocument.selection = created; } catch (e) { }
                dlgA.close(1);
            }
        };

        btnCloseA.onClick = function () { dlgA.close(0); };

        dlgA.show();

        // dlgA.show() はブロッキング。ダイアログが閉じた後にBを開く
        if (shouldOpenB && convertedFrames.length > 0) {
            showDialogB(doc, convertedFrames[0], convertedFrames, _openBDefaultAlignInt);
        }
    }

    // ============================================================
    // runConvert：ポイント文字をエリア内文字に変換して作成フレームを返す
    // ============================================================
    function runConvert(doc, sel, useSimple, useButton, useSelected, useSelectedDummy) {
        var created = [];

        /* 閉じたパスを取り出す（複合パスは先頭を見る）/ Return a closed path (for compound paths, inspect the first) */
        function getClosedPathItem(it) {
            try {
                if (!it) return null;
                if (it.typename === "PathItem") {
                    if (it.closed) return it;
                } else if (it.typename === "CompoundPathItem") {
                    if (it.pathItems && it.pathItems.length > 0) {
                        var p = it.pathItems[0];
                        if (p && p.closed) return p;
                    }
                }
            } catch (e) { }
            return null;
        }

        /* フォントとサイズを適用 / Apply font and size */
        function applyTextAttrs(tf, fontObj, sizePt) {
            try {
                if (sizePt > 0) tf.textRange.characterAttributes.size = sizePt;
            } catch (e0) { }
            try {
                if (fontObj) tf.textRange.characterAttributes.textFont = fontObj;
            } catch (e1) { }
        }

        if (useSelectedDummy) {
            // 選択オブジェクトにダミーテキスト：選択した図形（閉じたパス）をエリア内文字に変換してダミー文字を入れる
            var dummyText = (currentLanguage === "ja") ? DUMMY_TEXT_JA : DUMMY_TEXT_EN;
            var dummyFont = getTextFontSafe((currentLanguage === "ja") ? DUMMY_FONT_JA : DUMMY_FONT_EN);

            for (var d = 0; d < sel.length; d++) {
                var closedPath = getClosedPathItem(sel[d]);
                if (!closedPath) continue;
                try {
                    // 選択した閉じたパス自体をエリア内文字に変換 / Convert the selected closed path itself to Area Type
                    var dummyFrame = doc.textFrames.areaText(closedPath);
                    try { dummyFrame.textPath.filled = false; dummyFrame.textPath.stroked = false; } catch (e0) { }
                    try { dummyFrame.contents = dummyText; } catch (e1) { }
                    applyTextAttrs(dummyFrame, dummyFont, DUMMY_FONT_SIZE);
                    created.push(dummyFrame);
                } catch (e2) { }
            }

        } else if (useSelected) {
            // 選択オブジェクトを利用：選択テキストの内容を、選択した閉じたパスのエリア内文字に流し込む
            var sourceText = null, targetPath = null;
            for (var k = 0; k < sel.length; k++) {
                if (!sourceText && sel[k].typename === "TextFrame") { sourceText = sel[k]; }
                else if (!targetPath && sel[k].typename === "PathItem" && sel[k].closed) { targetPath = sel[k]; }
            }
            if (sourceText && targetPath) {
                try {
                    var sourceContents = sourceText.contents;
                    var sourceFont = null, sourceFontSize = 0;
                    try { sourceFont = sourceText.textRange.characterAttributes.textFont; sourceFontSize = sourceText.textRange.characterAttributes.size; } catch (e) { }
                    var duplicatedPath = targetPath.duplicate();
                    duplicatedPath.filled = false; duplicatedPath.stroked = false;
                    var selectedFrame = doc.textFrames.areaText(duplicatedPath);
                    selectedFrame.contents = sourceContents;
                    try { if (sourceFont) selectedFrame.textRange.characterAttributes.textFont = sourceFont; if (sourceFontSize > 0) selectedFrame.textRange.characterAttributes.size = sourceFontSize; } catch (e) { }
                    sourceText.remove();
                    targetPath.remove();
                    created.push(selectedFrame);
                } catch (e) { }
            }
        } else if (useSimple) {
            // シンプル：幅+1pt・高さ+1pt の長方形
            for (var sm = sel.length - 1; sm >= 0; sm--) {
                var smItem = sel[sm];
                if (smItem.typename === "TextFrame" && smItem.kind === TextType.POINTTEXT) {
                    try {
                        var smB = smItem.geometricBounds;
                        var smRect = doc.pathItems.rectangle(smB[1], smB[0], smB[2] - smB[0] + 1, smB[1] - smB[3] + 1);
                        smRect.filled = false; smRect.stroked = false;
                        var smC = smItem.contents, smFont = null, smSize = 0;
                        try { smFont = smItem.textRange.characterAttributes.textFont; smSize = smItem.textRange.characterAttributes.size; } catch (e) { }
                        var smTf = doc.textFrames.areaText(smRect);
                        smTf.contents = smC;
                        try { if (smFont) smTf.textRange.characterAttributes.textFont = smFont; if (smSize > 0) smTf.textRange.characterAttributes.size = smSize; } catch (e) { }
                        created.push(smTf);
                        smItem.remove();
                    } catch (e) { }
                }
            }
        } else if (useButton) {
            // ボタン風：幅×1.2・高さ×1.8、中央揃え
            for (var bt = sel.length - 1; bt >= 0; bt--) {
                var btItem = sel[bt];
                if (btItem.typename === "TextFrame" && btItem.kind === TextType.POINTTEXT) {
                    try {
                        var btB = btItem.geometricBounds;
                        var btOW = btB[2] - btB[0], btOH = btB[1] - btB[3];
                        var btW = btOW * 1.2, btH = btOH * 1.8;
                        var btRect = doc.pathItems.rectangle(
                            btB[1] + (btH - btOH) / 2, btB[0] - (btW - btOW) / 2, btW, btH);
                        btRect.filled = false; btRect.stroked = false;
                        var btC = btItem.contents, btFont = null, btSize = 0;
                        try { btFont = btItem.textRange.characterAttributes.textFont; btSize = btItem.textRange.characterAttributes.size; } catch (e) { }
                        var btTf = doc.textFrames.areaText(btRect);
                        btTf.contents = btC;
                        try { if (btFont) btTf.textRange.characterAttributes.textFont = btFont; if (btSize > 0) btTf.textRange.characterAttributes.size = btSize; } catch (e) { }
                        // Horizontal center + vertical center (frame alignment)
                        try { btTf.textRange.paragraphAttributes.justification = Justification.CENTER; } catch (e) { }
                        // NOTE: DOM `verticalAlignment` is not reliably applied for Area Type in some versions.
                        // Use the same dynamic action used in Dialog B (adobe_frameAlignment): 0=Top, 1=Center, 2=Bottom, 3=Justify
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

    // ============================================================
    // ダイアログB：エリア内文字 調整
    // ============================================================
    function showDialogB(doc, initialTf, targetFrames, defaultAlignInt) {
        var rulerInfo = getRulerUnitInfo(doc);

        // Dialog A から受け取った変換結果を確実に対象にする（選択が変わっても崩れないように）
        if (targetFrames && targetFrames.length) {
            try { app.activeDocument.selection = targetFrames; } catch (e) { }
        } else if (initialTf) {
            try { app.activeDocument.selection = [initialTf]; } catch (e2) { }
        }
        try { app.redraw(); } catch (e3) { }

        // Dialog B の調整対象（非分離モード用に固定）。
        // モーダルダイアログ中は selection が変動/取得不能になることがあるため、Aから渡された配列を優先する。
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

        // 「テキストを分離」パネル / "Separate text" panel
        var separatePanel = dlgB.add("panel", undefined, L("panel.separate"));
        setupPanel(separatePanel);
        separatePanel.orientation = "row";
        separatePanel.alignChildren = ["center", "center"];

        var radNoAdjust = separatePanel.add("radiobutton", undefined, L("radio.noAdjust"));
        var radSeparate = separatePanel.add("radiobutton", undefined, L("radio.strokeBlack"));
        var radPathOpacity = separatePanel.add("radiobutton", undefined, L("radio.noPath"));
        var radRemovePath = separatePanel.add("radiobutton", undefined, L("radio.removePath"));
        radNoAdjust.value = true;

        // 2カラムレイアウト（左右カラムは上揃えで横いっぱいに）/ Two-column layout (columns fill width, top-aligned)
        var columnsGroup = dlgB.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.spacing = 10;

        var leftColumn = columnsGroup.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = "fill";

        var rightColumn = columnsGroup.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = "fill";

        // 右カラム：行揃え / Right column: justification
        var justifyPanel = rightColumn.add("panel", undefined, L("panel.justify"));
        setupPanel(justifyPanel);
        var radJustLeft = justifyPanel.add("radiobutton", undefined, L("radio.justifyLeft"));
        var radJustCenter = justifyPanel.add("radiobutton", undefined, L("radio.justifyCenter"));
        var radJustRight = justifyPanel.add("radiobutton", undefined, L("radio.justifyRight"));
        var radJustFull = justifyPanel.add("radiobutton", undefined, L("radio.justifyFull"));
        var radJustFullAll = justifyPanel.add("radiobutton", undefined, L("radio.justifyFullAll"));
        radJustLeft.value = true;

        // 右カラム：テキストの配置 / Right column: text alignment
        var textAlignPanel = rightColumn.add("panel", undefined, L("panel.textAlign"));
        setupPanel(textAlignPanel);
        var radAlignTop = textAlignPanel.add("radiobutton", undefined, L("radio.alignTop"));
        var radAlignCenter = textAlignPanel.add("radiobutton", undefined, L("radio.alignCenter"));
        var radAlignBottom = textAlignPanel.add("radiobutton", undefined, L("radio.alignBottom"));
        var radAlignJustify = textAlignPanel.add("radiobutton", undefined, L("radio.alignJustify"));
        // Default alignment
        if (defaultAlignInt === 1) {
            radAlignCenter.value = true;
        } else if (defaultAlignInt === 2) {
            radAlignBottom.value = true;
        } else if (defaultAlignInt === 3) {
            radAlignJustify.value = true;
        } else {
            radAlignTop.value = true;
        }

        // 左カラム：フォントサイズ / Left column: font size
        var fontSizePanel = leftColumn.add("panel", undefined, L("panel.fontSize"));
        setupPanel(fontSizePanel);
        var fontSizeRow = fontSizePanel.add("group");
        fontSizeRow.alignment = "left";
        fontSizeRow.add("statictext", undefined, labelText("label.fontSize"));
        var etFontSize = fontSizeRow.add("edittext", undefined, "");
        etFontSize.characters = 5;
        fontSizeRow.add("statictext", undefined, "pt");
        var fontSizeButtonRow = fontSizePanel.add("group");
        fontSizeButtonRow.orientation = "row";
        var btnTextSize = fontSizeButtonRow.add("button", undefined, L("button.overset"));
        var btnFit = fontSizeButtonRow.add("button", undefined, L("button.fit"));

        // 左カラム：フレームサイズ / Left column: frame size
        var frameSizePanel = leftColumn.add("panel", undefined, L("panel.frameSize"));
        setupPanel(frameSizePanel);
        var widthRow = frameSizePanel.add("group");
        var lblWidth = widthRow.add("statictext", undefined, labelText("label.width"));
        lblWidth.preferredSize.width = 28;
        var etWidth = widthRow.add("edittext", undefined, "");
        etWidth.characters = 5;
        widthRow.add("statictext", undefined, rulerInfo.label);
        var etChars = widthRow.add("edittext", undefined, "");
        etChars.characters = 4;
        var lblChars = widthRow.add("statictext", undefined, L("label.chars"));
        var heightRow = frameSizePanel.add("group");
        var lblHeight = heightRow.add("statictext", undefined, labelText("label.height"));
        lblHeight.preferredSize.width = 28;
        var etHeight = heightRow.add("edittext", undefined, "");
        etHeight.characters = 5;
        heightRow.add("statictext", undefined, rulerInfo.label);
        var btnFrameAuto = heightRow.add("button", undefined, L("button.frameAuto"));
        btnFrameAuto.preferredSize.width = 60;
        // 英語UIでは chars 計算は不正確なため使用不可にする
        if (currentLanguage !== "ja") {
            etChars.enabled = false;
            lblChars.enabled = false;
        }

        // 左カラム：インデント / Left column: indent
        var indentPanel = leftColumn.add("panel", undefined, L("panel.indent"));
        setupPanel(indentPanel);
        indentPanel.orientation = "row";
        indentPanel.alignChildren = ["left", "top"];
        var indentFieldsColumn = indentPanel.add("group");
        indentFieldsColumn.orientation = "column";
        indentFieldsColumn.alignChildren = "left";
        var leftIndentRow = indentFieldsColumn.add("group");
        var chkLeftIndent = leftIndentRow.add("checkbox", undefined, L("label.indentLeft"));
        chkLeftIndent.preferredSize.width = 40;
        var etLeftIndent = leftIndentRow.add("edittext", undefined, "0");
        etLeftIndent.characters = 4;
        leftIndentRow.add("statictext", undefined, rulerInfo.label);
        etLeftIndent.enabled = false;
        var rightIndentRow = indentFieldsColumn.add("group");
        var chkRightIndent = rightIndentRow.add("checkbox", undefined, L("label.indentRight"));
        chkRightIndent.preferredSize.width = 40;
        var etRightIndent = rightIndentRow.add("edittext", undefined, "0");
        etRightIndent.characters = 4;
        rightIndentRow.add("statictext", undefined, rulerInfo.label);
        etRightIndent.enabled = false;
        var syncColumn = indentPanel.add("group");
        syncColumn.orientation = "column";
        syncColumn.alignChildren = "left";
        syncColumn.alignment = ["left", "center"];
        var chkSync = syncColumn.add("checkbox", undefined, L("checkbox.sync"));

        // 左カラム：オプション / Left column: options
        var optionsPanel = leftColumn.add("panel", undefined, L("panel.options"));
        setupPanel(optionsPanel);
        var marginRow = optionsPanel.add("group");
        var chkMargin = marginRow.add("checkbox", undefined, L("checkbox.margin"));
        var etMargin = marginRow.add("edittext", undefined, "0");
        etMargin.characters = 6;
        var lblUnit = marginRow.add("statictext", undefined, rulerInfo.label);
        etMargin.enabled = false;
        lblUnit.enabled = false;
        var chkLeader = optionsPanel.add("checkbox", undefined, L("checkbox.leader"));

        // ボタンエリア
        var bottomBar = dlgB.add("group");
        bottomBar.orientation = "row";
        bottomBar.alignment = "fill";
        bottomBar.alignChildren = ["fill", "center"];
        var chkPreview = bottomBar.add("checkbox", undefined, L("checkbox.preview"));
        chkPreview.alignment = ["left", "center"];
        var bottomButtonRow = bottomBar.add("group");
        bottomButtonRow.alignment = ["right", "center"];
        var btnCloseB = bottomButtonRow.add("button", undefined, L("button.close"), { name: "cancel" });
        var btnRun = bottomButtonRow.add("button", undefined, L("button.run"), { name: "ok" });

        // 状態変数
        var isPreviewActive = false;
        var _autoSizeMode = "none";
        var hasMultiParagraph = false;
        var _frameAutoOn = false;
        var fontSize = 0;

        /* 入力バリデーション（幅/高さ）/ Input validation (width/height) */
        // 0以下・NaN・極端値を弾いてIllustratorの不安定化を避ける
        var _lastValidWidth = null; // ruler units
        var _lastValidHeight = null; // ruler units

        function _maxSizeInRulerUnits() {
            // 上限は pt で固定（極端値を防ぐ）。表示単位に合わせて換算。
            // 100000pt は現実的に十分大きく、かつ事故りにくい上限。
            return 100000 / rulerInfo.toPt;
        }

        function validateSizeField(editText, lastValue, fieldName) {
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
            // 小さすぎる値も事故の元なので下限を設ける
            if (v < 0.01) {
                v = 0.01;
                editText.text = Math.round(v * 100) / 100;
            }
            return v;
        }

        var allRadios = [radSeparate, radPathOpacity, radRemovePath, radNoAdjust];
        var allJustRadios = [radJustLeft, radJustCenter, radJustRight, radJustFull, radJustFullAll];
        var allAlignRadios = [radAlignTop, radAlignCenter, radAlignBottom, radAlignJustify];

        /* ローカル関数 / Local functions */
        function updateFrameAutoButtonLabel() {
            btnFrameAuto.text = (_frameAutoOn ? "✓ " : "") + L("button.frameAuto");
        }

        function getAdjustmentPt() {
            var mp = chkMargin.value ? (parseFloat(etMargin.text) || 0) * rulerInfo.toPt : 0;
            var lp = (chkLeftIndent.value || chkSync.value) ? (parseFloat(etLeftIndent.text) || 0) * rulerInfo.toPt : 0;
            var rp = chkSync.value ? lp : (chkRightIndent.value ? (parseFloat(etRightIndent.text) || 0) * rulerInfo.toPt : 0);
            return 2 * mp + lp + rp;
        }

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
                var ij = tf0.paragraphs.length > 0 ? tf0.paragraphs[0].justification : Justification.LEFT;
                radJustLeft.value = (ij === Justification.LEFT);
                radJustCenter.value = (ij === Justification.CENTER);
                radJustRight.value = (ij === Justification.RIGHT);
                radJustFull.value = (ij === Justification.FULLJUSTIFYLASTLINELEFT);
                radJustFullAll.value = (ij === Justification.FULLJUSTIFY);
                if (!radJustLeft.value && !radJustCenter.value && !radJustRight.value &&
                    !radJustFull.value && !radJustFullAll.value) { radJustLeft.value = true; }
            } catch (e) { radJustLeft.value = true; }
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
            if (fontSize > 0) {
                etChars.text = Math.round(((tf0.textPath.width - getAdjustmentPt()) / fontSize) * 100) / 100;
            }
            btnTextSize.enabled = !hasMultiParagraph;
            btnFit.enabled = !hasMultiParagraph;
        }

        function runAdjust(forPreview) {
            var doSeparate = radSeparate.value;
            var doPathOpacity = radPathOpacity.value;
            var doRemovePath = radRemovePath.value;
            var doTextSize = (_autoSizeMode === "textsize");
            var doFit = (_autoSizeMode === "fit");
            var doFrameSize = _frameAutoOn;
            var doLeader = chkLeader.value;

            var justValue = Justification.LEFT;
            if (radJustCenter.value) justValue = Justification.CENTER;
            else if (radJustRight.value) justValue = Justification.RIGHT;
            else if (radJustFull.value) justValue = Justification.FULLJUSTIFYLASTLINELEFT;
            else if (radJustFullAll.value) justValue = Justification.FULLJUSTIFY;

            var alignValueInt = 0;
            if (radAlignCenter.value) alignValueInt = 1;
            else if (radAlignBottom.value) alignValueInt = 2;
            else if (radAlignJustify.value) alignValueInt = 3;

            var leftIndentPt = (chkLeftIndent.value || chkSync.value)
                ? (parseFloat(etLeftIndent.text) || 0) * rulerInfo.toPt : 0;
            var rightIndentPt = chkSync.value
                ? leftIndentPt
                : (chkRightIndent.value ? (parseFloat(etRightIndent.text) || 0) * rulerInfo.toPt : 0);
            var marginPt = chkMargin.value ? (parseFloat(etMargin.text) || 0) * rulerInfo.toPt : 0;

            var isSeparationMode = doSeparate || doPathOpacity || doRemovePath;
            var savedSel = [];
            if (!isSeparationMode) {
                var origSel = app.activeDocument.selection;
                for (var s = 0; s < origSel.length; s++) { savedSel.push(origSel[s]); }
            }

            // 分離モードはオブジェクトが置き換わるため selection を使う。
            // 非分離モードは固定ターゲットを優先して、プレビューが確実に反映されるようにする。
            var targets = null;
            if (isSeparationMode) {
                targets = app.activeDocument.selection;
            } else {
                if (!(_fixedTargets && _fixedTargets.length)) { refreshFixedTargetsFromSelection(); }
                targets = _fixedTargets && _fixedTargets.length ? _fixedTargets : app.activeDocument.selection;
            }
            for (var i = targets.length - 1; i >= 0; i--) {
                var obj = targets[i];
                if (obj.typename === "TextFrame" && obj.kind === TextType.AREATEXT) {
                    if (isSeparationMode) {
                        var bounds = obj.geometricBounds;
                        var left = bounds[0] - marginPt, top = bounds[1] + marginPt;
                        var right = bounds[2] + marginPt, bottom = bounds[3] - marginPt;
                        var rect = doc.pathItems.rectangle(top, left, right - left, top - bottom);
                        rect.filled = false; rect.stroked = true;
                        var newText = doc.textFrames.add();
                        newText.contents = obj.contents;
                        newText.position = [left, top - newText.textRange.characterAttributes.size];
                        var ca = obj.textRange.characterAttributes, nca = newText.textRange.characterAttributes;
                        if (ca.textFont && ca.textFont.name) { nca.textFont = ca.textFont; }
                        nca.size = ca.size; nca.leading = ca.leading;
                        obj.remove();
                        if (doSeparate) {
                            var blk = new CMYKColor(); blk.cyan = 0; blk.magenta = 0; blk.yellow = 0; blk.black = 100;
                            rect.stroked = true; rect.strokeColor = blk; rect.strokeWidth = 1; rect.filled = false;
                        }
                        if (doPathOpacity) { rect.stroked = false; rect.filled = false; }
                        if (doRemovePath) { rect.remove(); }
                        try { var pa = newText.textRange.paragraphAttributes; pa.justification = justValue; pa.leftIndent = leftIndentPt; pa.rightIndent = rightIndentPt; } catch (e) { }
                        if (doLeader) {
                            radJustLeft.value = false; radJustCenter.value = false; radJustRight.value = true;
                            radJustFull.value = false; radJustFullAll.value = false;
                            applyLeaderTab(newText);
                        }
                    } else {
                        try { obj.spacing = marginPt; } catch (e) { }
                        // 幅/高さ：NaN・0以下・極端値をガード
                        var wRu = validateSizeField(etWidth, _lastValidWidth, "width");
                        var hRu = validateSizeField(etHeight, _lastValidHeight, "height");
                        if (wRu !== null) { _lastValidWidth = wRu; try { obj.textPath.width = wRu * rulerInfo.toPt; } catch (e) { } }
                        if (hRu !== null) { _lastValidHeight = hRu; try { obj.textPath.height = hRu * rulerInfo.toPt; } catch (e) { } }
                        if (doTextSize) { shrinkFont(obj); } else if (doFit) { fitTextToFrame(obj); }
                        // プレビュー中は app.doScript 経由の処理を走らせない（不安定化・クラッシュ回避）
                        if (doFrameSize && !forPreview) { expandFrameToFit(obj); }
                        try { var pa2 = obj.textRange.paragraphAttributes; pa2.justification = justValue; pa2.leftIndent = leftIndentPt; pa2.rightIndent = rightIndentPt; } catch (e) { }
                        applyAreaTextFrameAlignment(obj, alignValueInt, forPreview);
                        if (doLeader) {
                            radJustLeft.value = false; radJustCenter.value = false; radJustRight.value = true;
                            radJustFull.value = false; radJustFullAll.value = false;
                            applyLeaderTab(obj);
                        }
                    }
                }
            }
            if (savedSel.length > 0) { try { app.activeDocument.selection = savedSel; } catch (e) { } }
            app.redraw();
        }

        function updatePreview() {
            if (isPreviewActive) {
                try { app.undo(); } catch (e) { }
                try { app.redraw(); } catch (e2) { }
                isPreviewActive = false;

                // undo 後は参照が無効化されるため、対象を取り直す
                refreshFixedTargetsFromSelection();
            }
            if (chkPreview.value) {
                // プレビュー適用前にも念のため取り直す（モード切替直後の選択揺れ対策）
                if (!(_fixedTargets && _fixedTargets.length)) { refreshFixedTargetsFromSelection(); }
                runAdjust(true);
                isPreviewActive = true;
            }
        }

        function updateSeparationDim() {
            var sep = radSeparate.value || radPathOpacity.value || radRemovePath.value;
            justifyPanel.enabled = !sep;
            textAlignPanel.enabled = !sep;
            fontSizePanel.enabled = !sep;
            frameSizePanel.enabled = !sep;
            btnFrameAuto.enabled = !sep;
            indentPanel.enabled = !sep;
            optionsPanel.enabled = !sep;
            if (fontSizePanel.enabled) { btnTextSize.enabled = !hasMultiParagraph; btnFit.enabled = !hasMultiParagraph; }
            else { btnTextSize.enabled = false; btnFit.enabled = false; }
        }

        function updateMarginEnabled() {
            chkMargin.enabled = radNoAdjust.value;
            if (!radNoAdjust.value) { etMargin.enabled = false; lblUnit.enabled = false; }
            else { etMargin.enabled = chkMargin.value; lblUnit.enabled = chkMargin.value; }
        }

        function applyFontSizeFromField() {
            var newSize = parseFloat(etFontSize.text) || 0;
            if (newSize <= 0) return;
            fontSize = newSize;
            var tgts = app.activeDocument.selection;
            for (var i = 0; i < tgts.length; i++) {
                if (tgts[i].typename === "TextFrame" && tgts[i].kind === TextType.AREATEXT) {
                    try { tgts[i].textRange.characterAttributes.size = newSize; } catch (e) { }
                }
            }
            if (fontSize > 0) {
                etChars.text = Math.round((((parseFloat(etWidth.text) || 0) * rulerInfo.toPt - getAdjustmentPt()) / fontSize) * 100) / 100;
            }
            app.redraw();
        }

        function onWidthChange() {
            var wRu = validateSizeField(etWidth, _lastValidWidth, "width");
            if (wRu !== null) { _lastValidWidth = wRu; }
            if (fontSize > 0) {
                etChars.text = Math.round((((parseFloat(etWidth.text) || 0) * rulerInfo.toPt - getAdjustmentPt()) / fontSize) * 100) / 100;
            }
            updatePreview();
        }
        function onCharsChange() {
            if (fontSize > 0) {
                var nextW = (((parseFloat(etChars.text) || 0) * fontSize + getAdjustmentPt()) / rulerInfo.toPt);
                if (!isNaN(nextW) && isFinite(nextW) && nextW > 0) {
                    etWidth.text = Math.round(nextW * 100) / 100;
                    var wRu = validateSizeField(etWidth, _lastValidWidth, "width");
                    if (wRu !== null) { _lastValidWidth = wRu; }
                }
            }
            updatePreview();
        }
        function onAdjustmentChange() {
            if (chkSync.value) { etRightIndent.text = etLeftIndent.text; }
            onWidthChange();
        }

        function addJustifyKeyHandler(dialog) {
            dialog.addEventListener("keydown", function (event) {
                if (!event || !event.keyName) return;
                var k = String(event.keyName).toUpperCase();
                var target = null;
                if (k === "L") target = radJustLeft;
                else if (k === "C") target = radJustCenter;
                else if (k === "R") target = radJustRight;
                else if (k === "J") target = radJustFull;
                else if (k === "F") target = radJustFullAll;
                if (!target) return;
                for (var i = 0; i < allJustRadios.length; i++) { allJustRadios[i].value = (allJustRadios[i] === target); }
                try { if (chkLeader && chkLeader.value && target !== radJustRight) { chkLeader.value = false; } } catch (e) { }
                event.preventDefault();
                updatePreview();
            });
        }

        /* イベントハンドラ / Event handlers */
        for (var r = 0; r < allRadios.length; r++) {
            (function (btn) {
                btn.onClick = function () {
                    for (var j = 0; j < allRadios.length; j++) { allRadios[j].value = (allRadios[j] === btn); }
                    updateSeparationDim();
                    updateMarginEnabled();
                    updatePreview();
                };
            })(allRadios[r]);
        }
        for (var jr = 0; jr < allJustRadios.length; jr++) {
            (function (btn) {
                btn.onClick = function () {
                    for (var j = 0; j < allJustRadios.length; j++) { allJustRadios[j].value = (allJustRadios[j] === btn); }
                    try { if (chkLeader && chkLeader.value && btn !== radJustRight) { chkLeader.value = false; } } catch (e) { }
                    updatePreview();
                };
            })(allJustRadios[jr]);
        }
        for (var ar = 0; ar < allAlignRadios.length; ar++) {
            (function (btn) {
                btn.onClick = function () {
                    for (var j = 0; j < allAlignRadios.length; j++) { allAlignRadios[j].value = (allAlignRadios[j] === btn); }
                    updatePreview();
                };
            })(allAlignRadios[ar]);
        }
        addJustifyKeyHandler(dlgB);

        btnTextSize.onClick = function () { _autoSizeMode = "textsize"; runAdjust(false); _autoSizeMode = "none"; };
        btnFit.onClick = function () { _autoSizeMode = "fit"; runAdjust(false); _autoSizeMode = "none"; };
        chkPreview.onClick = updatePreview;
        btnFrameAuto.onClick = function () {
            _frameAutoOn = !_frameAutoOn;
            updateFrameAutoButtonLabel();
            if (!_frameAutoOn) {
                try {
                    var t = app.activeDocument.selection;
                    for (var i = 0; i < t.length; i++) {
                        if (t[i] && t[i].typename === "TextFrame" && t[i].kind === TextType.AREATEXT) { collapseFrameAuto(t[i]); }
                    }
                } catch (e) { }
            }
            updatePreview();
        };
        chkLeader.onClick = updatePreview;
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
            var hRu = validateSizeField(etHeight, _lastValidHeight, "height");
            if (hRu !== null) { _lastValidHeight = hRu; }
            updatePreview();
        };
        etChars.onChange = onCharsChange;
        etLeftIndent.onChange = onAdjustmentChange;
        etRightIndent.onChange = onAdjustmentChange;
        changeValueByArrowKey(etFontSize, false, applyFontSizeFromField);
        changeValueByArrowKey(etMargin, false, onAdjustmentChange);
        changeValueByArrowKey(etWidth, false, onWidthChange);
        changeValueByArrowKey(etHeight, false, updatePreview);
        changeValueByArrowKey(etChars, false, onCharsChange);
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

        // 初期値読み込み
        if (initialTf) { loadValuesFromFrame(initialTf); }

        updateSeparationDim();
        updateFrameAutoButtonLabel();

        // ダイアログBを開いたらプレビューON
        chkPreview.value = true;
        updatePreview();

        dlgB.show();
    }

    // ============================================================
    // エントリポイント / Entry point
    //   ダイナミックアクションを読み込み、終了時に必ずアンロードする
    //   Load dynamic actions and always unload them on exit
    // ============================================================
    if (app.documents.length === 0) {
        alert(L("alert.noDocument"));
        return;
    }

    var doc = app.activeDocument;
    var sel = doc.selection;

    if (!sel || sel.length === 0) {
        alert(L("alert.selectText"));
        return;
    }

    loadDynamicActions();
    try {
        // 選択内容の種別を判定 / Detect what kinds of objects are selected
        var hasPointText = false, hasAreaText = false, hasPath = false;
        for (var selIndex = 0; selIndex < sel.length; selIndex++) {
            var selItem = sel[selIndex];
            if (selItem.typename === "TextFrame") {
                if (selItem.kind === TextType.POINTTEXT || selItem.kind === TextType.PATHTEXT) hasPointText = true;
                if (selItem.kind === TextType.AREATEXT) hasAreaText = true;
            }
            if (selItem.typename === "PathItem" || selItem.typename === "CompoundPathItem") {
                hasPath = true;
            }
        }

        if (hasPointText || hasPath) {
            // ポイント文字 または 図形 → ダイアログA / Point text or shape → Dialog A
            showDialogA(doc, sel);
        } else if (hasAreaText) {
            // エリア内文字のみ → ダイアログB / Area text only → Dialog B
            var firstAreaFrame = null;
            for (var areaIndex = 0; areaIndex < sel.length; areaIndex++) {
                if (sel[areaIndex].typename === "TextFrame" && sel[areaIndex].kind === TextType.AREATEXT) {
                    firstAreaFrame = sel[areaIndex];
                    break;
                }
            }
            showDialogB(doc, firstAreaFrame, null, null);
        } else {
            alert(L("alert.selectText"));
        }
    } finally {
        unloadDynamicActions();
    }

})();
