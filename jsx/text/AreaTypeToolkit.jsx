#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ポイント文字・パス上文字・図形からエリア内文字をつくり、そのまま体裁を調整するツール。

- 変換ダイアログで作成方法（シンプル／ボタン風／選択オブジェクトを利用／ダミーテキスト）を選ぶ
- 調整ダイアログでフォントサイズ・行送り・行揃え・テキストの配置・フレームサイズ・インデント・間隔を設定
- エリア内文字を選択して実行すると、調整ダイアログから始まる
- プレビューONで結果を確認しながら操作できる

詳細はREADMEを参照。

*/

/*

### Overview

Builds Area Type from point text, path text, or shapes, then tunes its typesetting in place.

- The convert dialog picks the creation method (Simple / Button style / Use selected object / Dummy text)
- The adjust dialog sets font size, leading, justification, text placement, frame size, indents, and spacing
- Running it with Area Type selected starts straight from the adjust dialog
- Preview shows the result while you work

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AreaTypeToolkit";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-03";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-07-28";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AreaTypeToolkit.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AreaTypeToolkit.md
var SCRIPT_ARTICLE_URL = ""; /* 紹介記事 / article URL */

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
    function findAvailableTextFont(candidateFontNames) {
        try {
            if (candidateFontNames && candidateFontNames.length) {
                for (var i = 0; i < candidateFontNames.length; i++) {
                    try {
                        var font = app.textFonts.getByName(candidateFontNames[i]);
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
    function getCurrentLanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLanguage();

    /* 日英ラベル定義（カテゴリ別）/ Japanese-English labels grouped by category */
    var LABELS = {
        dialog: {
            title: { ja: "エリア内文字を調整", en: "Adjust Area Type" },
            convertTitle: { ja: "エリア内文字に変換", en: "Convert to Area Type" }
        },
        panel: {
            createMethod: { ja: "作成方法", en: "Creation method" },
            separateText: { ja: "テキストを分離", en: "Separate text" },
            leading: { ja: "行送り", en: "Leading" },
            justification: { ja: "行揃え", en: "Justification" },
            textAlign: { ja: "テキストの配置", en: "Text alignment" },
            fontSize: { ja: "フォントサイズ", en: "Font size" },
            frameSize: { ja: "フレームサイズ", en: "Frame size" },
            indent: { ja: "インデント", en: "Indent" },
            options: { ja: "オプション", en: "Options" }
        },
        radio: {
            noAdjust: { ja: "分離しない", en: "Don't separate" },
            strokeBlack: { ja: "枠を1pt黒に", en: "Frame: 1pt black" },
            hidePath: { ja: "枠を塗り・線なしに", en: "Frame: unpainted" },
            removePath: { ja: "枠を削除", en: "Delete frame" },
            styleSimple: { ja: "シンプル", en: "Simple" },
            styleButton: { ja: "ボタン風", en: "Button style" },
            useShape: { ja: "選択した図形に流し込む", en: "Pour into selected shape" },
            useShapeDummy: { ja: "選択した図形にダミーテキスト", en: "Dummy text in selected shape" },
            justifyLeft: { ja: "左揃え", en: "Left" },
            justifyCenter: { ja: "中央揃え", en: "Center" },
            justifyRight: { ja: "右揃え", en: "Right" },
            justifyLastLineLeft: {
                ja: "均等配置（最終行左揃え）",
                en: "Justify (last line left)"
            },
            justifyAllLines: { ja: "両端揃え", en: "Justify all lines" },
            alignTop: { ja: "上揃え", en: "Top" },
            alignCenter: { ja: "中央揃え", en: "Center" },
            alignBottom: { ja: "下揃え", en: "Bottom" },
            alignJustify: { ja: "均等配置", en: "Justify" }
        },
        checkbox: {
            linkIndents: { ja: "連動", en: "Link" },
            spacing: { ja: "内側のマージン", en: "Inset spacing" },
            autoSize: { ja: "自動サイズ", en: "Auto-size" },
            leaderTabs: { ja: "メニュー作成用（リーダー罫）", en: "Leader tabs" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        button: {
            convert: { ja: "変換", en: "Convert" },
            shrinkToFit: { ja: "文字あふれ解消", en: "Clear overset" },
            fitFontSize: { ja: "枠にフィット", en: "Fit to frame" },
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        label: {
            fontSize: { ja: "フォントサイズ", en: "Font size" },
            leadingPercent: { ja: "行送り", en: "Leading" },
            leadingEffective: { ja: "実寸", en: "Actual" },
            width: { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" },
            charsPerLine: { ja: "字詰め", en: "chars/line" },
            indentLeft: { ja: "左", en: "Left" },
            indentRight: { ja: "右", en: "Right" }
        },
        tip: {
            styleSimple: {
                ja: "ポイント文字の実寸＋1ptの長方形をフレームにします。",
                en: "Uses a rectangle the size of the point text plus 1pt as the frame."
            },
            styleButton: {
                ja: "幅1.2倍・高さ1.8倍の長方形をフレームにし、行揃えとテキストの配置を中央にします。",
                en: "Uses a rectangle 1.2x wide and 1.8x tall, with the text centered both ways."
            },
            useShape: {
                ja: "選択した図形を複製してフレームにし、選択したテキストを流し込みます。複数組を選べます。",
                en: "Duplicates each selected shape as a frame and pours the selected text into it. Several pairs at once."
            },
            useShapeDummy: {
                ja: "選択した図形自体をエリア内文字にして、ダミー文字を流し込みます。",
                en: "Turns each selected shape itself into Area Type and fills it with dummy text."
            },
            shrinkToFit: {
                ja: "あふれが消えるまでフォントサイズを縮小します。段落が2つ以上あるときは使えません。",
                en: "Shrinks the font until the overset clears. Unavailable when the frame has more than one paragraph."
            },
            fitFontSize: {
                ja: "あふれるまで拡大してから縮小し、枠いっぱいに収めます。",
                en: "Grows until it oversets, then shrinks to fill the frame."
            },
            autoSize: {
                ja: "エリア内文字の自動サイズ調整。プレビュー中は適用されません。",
                en: "Area Type auto-sizing. Not applied during preview."
            },
            leading: {
                ja: "自動行送り量（%）。実寸＝フォントサイズ×%。実寸に入力すると%を逆算します。固定行送りのテキストでは空で開きます。",
                en: "Auto-leading amount in %. Actual = font size x %. Enter an actual value to back-calculate the %. Empty for text with a fixed leading."
            },
            textAlign: {
                ja: "フレーム内でのテキストの縦位置。ダイナミックアクションで適用するため、プレビューには反映されません。",
                en: "Vertical placement inside the frame. Applied via a dynamic action, so it does not show in the preview."
            },
            charsPerLine: {
                ja: "1行の文字数から幅を逆算します。字幅が一定でないため日本語UIのみ。",
                en: "Works the width back out from the characters per line. Japanese UI only, since Roman glyph widths vary."
            },
            linkIndents: {
                ja: "左インデントの値を右にも適用します。",
                en: "Applies the left indent value to the right as well."
            },
            leaderTabs: {
                ja: "各段落に右揃えタブ（リーダー「…」／400pt）を設定します。ONにすると行揃えが「右揃え」になります。",
                en: "Sets a right-aligned tab with a leader at 400pt on every paragraph. Turning it on switches justification to right."
            }
        },
        alert: {
            selectText: {
                ja: "ポイント文字・パス上文字・エリア内文字・図形を選択してください。",
                en: "Please select point text, path text, area text, or a shape."
            },
            noDocument: {
                ja: "ドキュメントが開かれていません。",
                en: "No document is open."
            },
            convertFailed: {
                ja: "変換できる対象がありませんでした。ポイント文字や閉じたパスを選択してください。",
                en: "Nothing could be converted. Select point text or a closed path."
            }
        }
    };

    /* ラベル取得（"category.key" 形式、{slash} は / に展開）/ Resolve "category.key" label, expand {slash} to / */
    function getLabel(key) {
        var keyParts = key.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < keyParts.length; i++) {
            if (!labelNode) break;
            labelNode = labelNode[keyParts[i]];
        }
        var resolvedText = key;
        if (labelNode) {
            if (typeof labelNode[currentLanguage] === "string") resolvedText = labelNode[currentLanguage];
            else if (typeof labelNode.en === "string") resolvedText = labelNode.en;
        }
        return resolvedText.replace(/\{slash\}/g, "/");
    }

    /* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
    function labelWithColon(key) {
        return getLabel(key) + (currentLanguage === "ja" ? "：" : ":");
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
        /* Q（歯）= 0.25mm。古いバージョンでは RulerUnits.Qs が未定義なので最後に判定 / 1Q = 0.25mm; checked last because RulerUnits.Qs is undefined on older versions */
        if (rulerUnit === RulerUnits.Qs) return { label: "Q", toPt: (72 / 25.4) * 0.25 };
        return { label: "pt", toPt: 1 };
    }

    // =========================================
    // パネルレイアウト / Panel layout
    // =========================================

    /* パネルの余白と間隔 / Panel margins and spacing */
    var PANEL_MARGINS = [16, 20, 16, 12];
    var PANEL_SPACING = 8;

    /* パネルの共通設定 / Apply shared panel layout */
    function applyPanelLayout(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* グループの共通設定（row/column で整列を切り替え）/ Apply shared group layout (alignChildren switches by orientation) */
    function applyGroupLayout(group, orientation, spacing) {
        var groupOrientation = orientation || "column";
        group.orientation = groupOrientation;
        /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
        group.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
        group.alignment = "fill";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* ラジオ群から1つだけを選択状態にする / Turn on exactly one radio in the group */
    function selectRadio(radios, targetRadio) {
        for (var i = 0; i < radios.length; i++) {
            radios[i].value = (radios[i] === targetRadio);
        }
    }

    /* ラジオ群にクリック時の排他制御を割り当てる / Wire exclusive selection onto a radio group */
    function bindExclusiveRadios(radios, onSelect) {
        for (var i = 0; i < radios.length; i++) {
            (function (radio) {
                radio.onClick = function () {
                    selectRadio(radios, radio);
                    if (typeof onSelect === "function") { onSelect(radio); }
                };
            })(radios[i]);
        }
    }


    /* パス上文字 → ポイント文字（変換ダイアログの前処理）/ Path text → point text (convert-dialog preprocess) */
    function detachPathTextToPointText(doc, pathTextFrames) {
        var createdPointTexts = [];
        if (!doc || !pathTextFrames || !pathTextFrames.length) return createdPointTexts;

        // 新しく作ったテキストだけを選べるように、いったん選択を解除する
        doc.selection = null;

        for (var i = pathTextFrames.length - 1; i >= 0; i--) {
            var pathText = pathTextFrames[i];
            if (!pathText || pathText.typename !== "TextFrame" || pathText.kind !== TextType.PATHTEXT) continue;
            try {
                var pointText = replacePathTextWithPointText(doc, pathText);
                if (pointText) { createdPointTexts.push(pointText); }
            } catch (e) { }
        }

        return createdPointTexts;
    }

    /* パス上文字1件をポイント文字に置き換える / Replace one path text frame with point text */
    function replacePathTextWithPointText(doc, pathText) {
        var originalPath = pathText.textPath;
        if (!originalPath) return null;

        var charAttrSnapshots = snapshotCharacterAttributes(pathText);
        var textContents = pathText.contents;
        var justification = pathText.paragraphs.length > 0
            ? pathText.paragraphs[0].paragraphAttributes.justification
            : null;

        var pointText = doc.textFrames.add();

        // パスの始点にそろえる / Align with the start point of the path
        if (originalPath.pathPoints.length > 0) {
            var anchorPoint = originalPath.pathPoints[0].anchor;
            pointText.position = [anchorPoint[0], anchorPoint[1]];
        }

        pointText.contents = textContents;
        if (justification !== null && pointText.paragraphs.length > 0) {
            pointText.paragraphs[0].paragraphAttributes.justification = justification;
        }

        // 既定の線をいったん外し、このあと文字単位で復元する / Clear the default stroke, then restore per character
        pointText.textRange.characterAttributes.strokeColor = new NoColor();
        pointText.textRange.characterAttributes.strokeWeight = 0;
        restoreCharacterAttributes(pointText, charAttrSnapshots);

        pathText.remove(); // パスも一緒に削除される / the path is removed along with it
        pointText.selected = true;
        return pointText;
    }

    /* 文字単位の属性を控える / Snapshot per-character attributes */
    function snapshotCharacterAttributes(textFrame) {
        var snapshots = [];
        for (var i = 0; i < textFrame.characters.length; i++) {
            var charAttr = textFrame.characters[i].characterAttributes;
            snapshots.push({
                font: charAttr.textFont,
                size: charAttr.size,
                fillColor: charAttr.fillColor,
                strokeColor: charAttr.strokeColor,
                strokeWeight: charAttr.strokeWeight,
                autoLeading: charAttr.autoLeading,
                leading: charAttr.leading
            });
        }
        return snapshots;
    }

    /* 控えた文字属性を書き戻す（パス上文字で付いた変形は捨てる）/ Write snapshotted attributes back, dropping path-induced transforms */
    function restoreCharacterAttributes(textFrame, snapshots) {
        var copyCount = Math.min(textFrame.characters.length, snapshots.length);
        for (var i = 0; i < copyCount; i++) {
            var destAttrs = textFrame.characters[i].characterAttributes;
            var srcAttrs = snapshots[i];

            // 未インストールのフォントは失敗しうるので、他の属性と分けて適用する
            try { destAttrs.textFont = srcAttrs.font; } catch (e0) { }
            try {
                destAttrs.size = srcAttrs.size;
                destAttrs.fillColor = srcAttrs.fillColor;
                destAttrs.strokeColor = srcAttrs.strokeColor;
                destAttrs.strokeWeight = (srcAttrs.strokeColor && srcAttrs.strokeColor.typename === "NoColor") ? 0 : srcAttrs.strokeWeight;
                destAttrs.baselineShift = 0;
                destAttrs.horizontalScale = 100;
                destAttrs.verticalScale = 100;
                destAttrs.autoLeading = srcAttrs.autoLeading;
                if (!srcAttrs.autoLeading) { destAttrs.leading = srcAttrs.leading; }
            } catch (e1) { }
        }
    }

    /* 選択中のパス上文字をポイント文字に置き換え、新しい選択を返す / Swap selected path text for point text and return the new selection */
    function preprocessSelectionForConvertDialog(doc, selection) {
        if (!doc || !selection || !selection.length) return selection;

        // 削除で参照が無効になる前に振り分ける / Sort before any removal invalidates references
        var pathTextItems = [], otherItems = [];
        for (var i = 0; i < selection.length; i++) {
            var item = selection[i];
            if (!item) continue;
            if (item.typename === "TextFrame" && item.kind === TextType.PATHTEXT) { pathTextItems.push(item); }
            else { otherItems.push(item); }
        }
        if (!pathTextItems.length) return selection;

        var pointTextItems = detachPathTextToPointText(doc, pathTextItems);
        if (!pointTextItems.length) return selection;

        var nextSelection = otherItems.concat(pointTextItems);
        doc.selection = nextSelection;
        app.redraw();
        return nextSelection;
    }


    /* 自動サイズ調整ヘルパー（AutoFitTextFrame.jsx 参考）/ Auto-fit helpers (based on AutoFitTextFrame.jsx) */

    /* overflows プロパティを安全に読む（取れなければ null）/ Read the overflows property safely (null when unavailable) */
    function getOverflowState(textFrame) {
        try {
            if (textFrame && typeof textFrame.overflows !== "undefined") return !!textFrame.overflows;
        } catch (e) { }
        return null;
    }

    /* 指定行数までに全文字が収まっているかで文字あふれを判定 / Decide overset by whether every character fits within the given number of lines */
    function isTextOverset(textFrame, lineLimit) {
        if (textFrame.lines.length > 0) {
            var charCount = 0;
            lineLimit = (typeof lineLimit === "undefined" || lineLimit === null) ? 1 : Math.floor(lineLimit);
            if (lineLimit < 1) lineLimit = 1;
            if (lineLimit > textFrame.lines.length) lineLimit = textFrame.lines.length;
            for (var i = 0; i < lineLimit; i++) { charCount += textFrame.lines[i].characters.length; }
            return charCount < textFrame.characters.length;
        }
        return textFrame.characters.length > 0;
    }

    /* テキストフレームが文字あふれしているか / Whether the text frame is overset */
    function isFrameOverset(textFrame) {
        if (textFrame && textFrame.kind === TextType.AREATEXT) {
            var overflowState = getOverflowState(textFrame);
            if (overflowState !== null) return overflowState;
            try { return isTextOverset(textFrame, textFrame.lines.length || 1); } catch (e) { return false; }
        }
        try { return isTextOverset(textFrame, textFrame.lines.length || 1); } catch (e) { return false; }
    }

    /* 固定行送りのとき、行送り÷フォントサイズの比率を返す / Ratio of leading to font size, for a fixed (non-auto) leading */
    function getLeadingRatioInfo(textFrame) {
        try {
            var attrs = textFrame.textRange.characterAttributes;
            if (attrs.autoLeading) return null;
            var size = attrs.size, leading = attrs.leading;
            if (size > 0 && leading > 0) return { ratio: leading / size };
        } catch (e) { }
        return null;
    }

    /* 比率を保ったまま、新しいフォントサイズに行送りを合わせる / Rescale the leading to a new font size, keeping the ratio */
    function applyProportionalLeading(textFrame, newSize, leadingRatioInfo) {
        if (!leadingRatioInfo) return;
        try { textFrame.textRange.characterAttributes.leading = newSize * leadingRatioInfo.ratio; } catch (e) { }
    }


    // =========================================
    // ダイナミックアクション / Dynamic actions
    //   スクリプト実行時に読み込み、終了時にアンロードする
    //   Loaded at startup, unloaded on exit
    // =========================================

    /* 文字列を ASCII 16進に変換 / Convert a string to ASCII hex */
    function asciiToHex(text) {
        var hex = "";
        for (var i = 0; i < text.length; i++) {
            var hexPair = text.charCodeAt(i).toString(16);
            if (hexPair.length < 2) hexPair = "0" + hexPair;
            hex += hexPair;
        }
        return hex;
    }

    /* アクション名ブロック /name [ <len> <hex> ] を生成 / Build the /name [ <len> <hex> ] block */
    function buildActionNameBlock(name) {
        return "/name [ " + name.length + " " + asciiToHex(name).toUpperCase() + " ]";
    }

    /* アクションセット定義（.aia 文字列）を組み立てる / Build an action set definition (.aia string) */
    function buildActionSetAia(setName, internalName, localizedNameHex, parameterKey, actionDefinitions) {
        // actionDefinitions: [{name: "...", value: <int>}, ...]
        var aia = "/version 3" +
            buildActionNameBlock(setName) +
            "/isOpen 1" +
            "/actionCount " + actionDefinitions.length;

        for (var i = 0; i < actionDefinitions.length; i++) {
            var actionDef = actionDefinitions[i];
            aia += "/action-" + (i + 1) + " {" +
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
        return aia;
    }

    /* .aia 文字列を一時ファイル経由で読み込む（既存があれば先に外す）/ Load an .aia string via a temp file (unload any existing set first) */
    function loadActionSet(setName, aiaString) {
        try { app.unloadAction(setName, ""); } catch (e0) { }
        var tempFile = new File(Folder.temp + "/AreaTypeToolkit_action_" + setName + ".aia");
        tempFile.open("w");
        tempFile.write(aiaString);
        tempFile.close();
        app.loadAction(tempFile);
        try { tempFile.remove(); } catch (e1) { }
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
    function enableFrameAutoSize(textFrame) {
        app.activeDocument.selection = [textFrame];
        runAutoSizeAction(1);
    }

    /* フレームサイズ：自動サイズ調整を OFF にする / Frame size: turn auto-size off */
    function disableFrameAutoSize(textFrame) {
        app.activeDocument.selection = [textFrame];
        runAutoSizeAction(2);
    }

    /* 自動サイズ調整アクションを実行（1=ON, 2=OFF）/ Run the auto-size action (1=ON, 2=OFF) */
    function runAutoSizeAction(autoSizeValue) {
        if (autoSizeValue !== 1 && autoSizeValue !== 2) return;
        var actionName = (autoSizeValue === 1) ? "AutoSizeOn" : "AutoSizeOff";
        try { app.doScript(actionName, "AreaType", false); } catch (e) { }
    }

    /* テキストの配置アクションを実行（0=上, 1=中央, 2=下, 3=均等）/ Run the frame-alignment action (0=top, 1=center, 2=bottom, 3=justify) */
    function runFrameAlignmentAction(alignmentValue) {
        if (alignmentValue !== 0 && alignmentValue !== 1 && alignmentValue !== 2 && alignmentValue !== 3) return;
        var actionName = "AlignTop";
        if (alignmentValue === 1) actionName = "AlignCenter";
        else if (alignmentValue === 2) actionName = "AlignBottom";
        else if (alignmentValue === 3) actionName = "AlignJustify";
        try { app.doScript(actionName, "AreaText", false); } catch (e) { }
    }

    /* 対象を選択してテキストの配置アクションを実行する（プレビュー中は実行しない）/ Select the frame and run the placement action (skipped during preview) */
    function applyAreaTextFrameAlignment(textFrame, alignmentValue, forPreview) {
        // app.doScript はプレビュー中にダイアログから呼ぶと不安定なためスキップ
        if (forPreview) return;
        try {
            var doc = app.activeDocument;
            doc.selection = null;
            doc.selection = [textFrame];
            app.redraw(); // Illustratorに選択状態を確定させる
            runFrameAlignmentAction(alignmentValue);
        } catch (e) { }
    }

    // ちぢむ処理（バイナリサーチ：最大 40 回でオーバーセットにならない最大サイズを探す）
    function shrinkFontToFit(textFrame) {
        if (textFrame.characters.length <= 0 || !isFrameOverset(textFrame)) return;
        var leadingRatioInfo = getLeadingRatioInfo(textFrame);
        var upperSize = textFrame.textRange.characterAttributes.size;
        var lowerSize = 0.1;

        // lowerSize でもオーバーセットなら最小サイズのまま終了
        textFrame.textRange.characterAttributes.size = lowerSize;
        applyProportionalLeading(textFrame, lowerSize, leadingRatioInfo);
        if (isFrameOverset(textFrame)) return;

        // バイナリサーチ
        for (var attempt = 0; attempt < 40; attempt++) {
            var midSize = (lowerSize + upperSize) / 2;
            textFrame.textRange.characterAttributes.size = midSize;
            applyProportionalLeading(textFrame, midSize, leadingRatioInfo);
            if (isFrameOverset(textFrame)) {
                upperSize = midSize;
            } else {
                lowerSize = midSize;
            }
            if (upperSize - lowerSize < 0.1) break;
        }

        // オーバーセットにならない側（lowerSize）に確定
        textFrame.textRange.characterAttributes.size = lowerSize;
        applyProportionalLeading(textFrame, lowerSize, leadingRatioInfo);
    }

    // テキスト（フィット）：オーバーセットになるまで拡大 → ちぢむ（fitFont 相当）
    function fitFontSizeToFrame(textFrame) {
        if (textFrame.characters.length <= 0) return;
        var leadingRatioInfo = getLeadingRatioInfo(textFrame);
        var originalSize = textFrame.textRange.characterAttributes.size;

        // Step 1: オーバーセットが出るまで 2 倍ずつ拡大
        if (!isFrameOverset(textFrame)) {
            var trialSize = originalSize, guardCount = 0;
            while (!isFrameOverset(textFrame) && guardCount < 25) {
                guardCount++;
                trialSize = trialSize * 2;
                if (trialSize > 100000) break;
                try { textFrame.textRange.characterAttributes.size = trialSize; applyProportionalLeading(textFrame, trialSize, leadingRatioInfo); } catch (e) { break; }
            }
        }

        // オーバーセットが出なければ元に戻して終了
        if (!isFrameOverset(textFrame)) {
            try { textFrame.textRange.characterAttributes.size = originalSize; applyProportionalLeading(textFrame, originalSize, leadingRatioInfo); } catch (e) { }
            return;
        }

        // Step 2: ちぢむ処理
        shrinkFontToFit(textFrame);
    }

    // ↑↓キーで値を増減（プレビュー更新コールバックつき）
    function changeValueByArrowKey(editText, allowNegative, onChangeCallback) {
        editText.addEventListener("keydown", function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboardState = ScriptUI.environment.keyboardState;

            if (keyboardState.shiftKey) {
                var delta = 10;
                if (event.keyName == "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value = Math.floor((value - 1) / delta) * delta;
                    event.preventDefault();
                }
            } else if (keyboardState.altKey) {
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

            value = keyboardState.altKey ? Math.round(value * 10) / 10 : Math.round(value);
            if (!allowNegative && value < 0) value = 0;

            editText.text = value;
            if (typeof onChangeCallback === "function") { onChangeCallback(); }
        });
    }

    // メニュー作成用リーダー罫：各段落に右揃えタブ（リーダー「…」400pt）を適用
    function applyLeaderTabStops(textFrame) {
        try {
            var tabStop = new TabStopInfo();
            tabStop.position = 400;
            tabStop.alignment = TabStopAlignment.Right;
            tabStop.leader = "…";
            for (var i = 0; i < textFrame.paragraphs.length; i++) {
                var paraAttrs = textFrame.paragraphs[i].paragraphAttributes;
                paraAttrs.tabStops = [tabStop];
                paraAttrs.justification = Justification.RIGHT;
            }
        } catch (e) { }
    }

    // ============================================================
    // 変換ダイアログ：ポイント文字 → エリア内文字 変換
    // ============================================================
    function showConvertDialog(doc, selection) {
        selection = preprocessSelectionForConvertDialog(doc, selection);
        var convertDialog = new Window("dialog", getLabel("dialog.convertTitle") + " " + SCRIPT_VERSION);
        convertDialog.alignChildren = "fill";
        convertDialog.margins = 20;

        // 作成方法のラジオボタン / Creation-method radios
        var createMethodPanel = convertDialog.add("panel", undefined, getLabel("panel.createMethod"));
        applyPanelLayout(createMethodPanel);

        var radStyleSimple = createMethodPanel.add("radiobutton", undefined, getLabel("radio.styleSimple"));
        var radStyleButton = createMethodPanel.add("radiobutton", undefined, getLabel("radio.styleButton"));
        var radUseShape = createMethodPanel.add("radiobutton", undefined, getLabel("radio.useShape"));
        var radUseShapeDummy = createMethodPanel.add("radiobutton", undefined, getLabel("radio.useShapeDummy"));
        radStyleSimple.value = true;
        radStyleSimple.helpTip = getLabel("tip.styleSimple");
        radStyleButton.helpTip = getLabel("tip.styleButton");
        radUseShape.helpTip = getLabel("tip.useShape");
        radUseShapeDummy.helpTip = getLabel("tip.useShapeDummy");

        // ボタンエリア：左から［キャンセル］［変換］
        var convertButtonRow = convertDialog.add("group");
        convertButtonRow.orientation = "row";
        convertButtonRow.alignment = ["center", "center"]; // center buttons
        convertButtonRow.spacing = 10;
        var btnCancelConvert = convertButtonRow.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnConvert = convertButtonRow.add("button", undefined, getLabel("button.convert"), { name: "ok" });

        // 選択種別フラグ
        var hasPointText = false, hasPathItem = false;
        for (var i = 0; i < selection.length; i++) {
            if (selection[i].typename === "TextFrame" && (selection[i].kind === TextType.POINTTEXT || selection[i].kind === TextType.PATHTEXT)) hasPointText = true;
            if (selection[i].typename === "PathItem" || selection[i].typename === "CompoundPathItem") hasPathItem = true;
        }

        // 選択内容に応じてモードを誘導
        if (hasPointText && hasPathItem) {
            // テキスト＋図形 → シンプル・ボタン風をディム、選択オブジェクトを利用を自動選択
            radStyleSimple.enabled = false;
            radStyleButton.enabled = false;
            radUseShape.value = true; // default, dummy is also available
        } else if (hasPointText && !hasPathItem) {
            // ポイント文字のみ → 選択オブジェクト系をディム
            radUseShape.enabled = false;
            radUseShapeDummy.enabled = false;
        } else if (!hasPointText && hasPathItem) {
            // 図形のみ → 「選択オブジェクトにダミーテキスト」を自動選択し、残りをディム
            radStyleSimple.enabled = false;
            radStyleButton.enabled = false;
            radUseShape.enabled = false;
            radUseShapeDummy.enabled = true;
            radUseShapeDummy.value = true;
        }

        bindExclusiveRadios([radStyleSimple, radStyleButton, radUseShape, radUseShapeDummy]);

        // 変換後に調整ダイアログを開くための変数
        var convertedFrames = [];
        var shouldOpenAdjustDialog = false;
        var initialAlignmentValue = null; // 0=Top,1=Center,2=Bottom,3=Justify

        btnConvert.onClick = function () {
            var currentSelection = app.activeDocument.selection;
            if (!currentSelection || currentSelection.length === 0) { return; }
            var useShape = !!radUseShape.value;
            var useShapeDummy = !!radUseShapeDummy.value;
            var createdFrames = convertSelectionToAreaText(doc, currentSelection,
                radStyleSimple.value, radStyleButton.value, useShape, useShapeDummy);
            if (createdFrames.length > 0) {
                convertedFrames = createdFrames;
                // If Button style was used, default the adjust dialog alignment to Center
                initialAlignmentValue = radStyleButton.value ? 1 : null;
                shouldOpenAdjustDialog = true;
                try { app.activeDocument.selection = createdFrames; } catch (e) { }
                convertDialog.close(1);
            } else {
                alert(getLabel("alert.convertFailed"));
            }
        };

        btnCancelConvert.onClick = function () { convertDialog.close(0); };

        convertDialog.show();

        // convertDialog.show() はブロッキング。閉じた後に調整ダイアログを開く
        if (shouldOpenAdjustDialog && convertedFrames.length > 0) {
            showAdjustDialog(doc, convertedFrames[0], convertedFrames, initialAlignmentValue);
        }
    }

    // ============================================================
    // エリア内文字への変換 / Conversion to Area Type
    // ============================================================

    /* フレームの起こし方（シンプル＝実寸＋1pt、ボタン風＝幅1.2倍・高さ1.8倍で中央）/ Frame styles (simple: +1pt; button: 1.2x / 1.8x, centered) */
    var AREA_TEXT_STYLE_SIMPLE = { widthScale: 1, heightScale: 1, widthPadPt: 1, heightPadPt: 1, centerOnText: false, centerText: false };
    var AREA_TEXT_STYLE_BUTTON = { widthScale: 1.2, heightScale: 1.8, widthPadPt: 0, heightPadPt: 0, centerOnText: true, centerText: true };

    /* 閉じたパスを取り出す（複合パスは先頭を見る）/ Return a closed path (for compound paths, inspect the first) */
    function getClosedPathItem(item) {
        if (!item) return null;
        if (item.typename === "PathItem") {
            return item.closed ? item : null;
        }
        if (item.typename === "CompoundPathItem" && item.pathItems.length > 0) {
            var firstPath = item.pathItems[0];
            return (firstPath && firstPath.closed) ? firstPath : null;
        }
        return null;
    }

    /* 閉じたパスを塗り・線なしのエリア内文字フレームにする / Turn a closed path into an unpainted Area Type frame */
    function createAreaTextFromPath(doc, pathItem) {
        pathItem.filled = false;
        pathItem.stroked = false;
        return doc.textFrames.areaText(pathItem);
    }

    /* 先頭文字のフォントとサイズを読み取る / Read font and size from the head of the text */
    function getFontAndSize(textFrame) {
        try {
            var attrs = textFrame.textRange.characterAttributes;
            return { font: attrs.textFont, size: attrs.size };
        } catch (e) {
            return { font: null, size: 0 };
        }
    }

    /* フォントとサイズを適用する / Apply font and size */
    function applyFontAndSize(textFrame, fontAndSize) {
        if (!fontAndSize) return;
        var attrs = textFrame.textRange.characterAttributes;
        if (fontAndSize.size > 0) { attrs.size = fontAndSize.size; }
        // 未インストールのフォントは適用に失敗しうる / An uninstalled font can fail to apply
        if (fontAndSize.font) { try { attrs.textFont = fontAndSize.font; } catch (e) { } }
    }

    /* 段落に自動行送りを設定する（Illustrator 上は「自動」表示になり、行送りがフォントサイズに追従する）
       Set auto-leading on the paragraphs so Illustrator shows "Auto" and the leading follows the font size */
    function applyAutoLeading(textFrame, percent) {
        if (!(percent > 0)) return;
        for (var i = 0; i < textFrame.paragraphs.length; i++) {
            try {
                textFrame.paragraphs[i].paragraphAttributes.autoLeadingAmount = percent;
                textFrame.paragraphs[i].characterAttributes.autoLeading = true;
            } catch (e) { }
        }
    }

    /* 自動行送りが有効ならその量（％）を返す。固定の行送りなら 0 を返し、欄を空のままにする
       Return the auto-leading amount in %, or 0 when the frame uses a fixed leading (leaves the field empty) */
    function getAutoLeadingPercent(textFrame) {
        try {
            if (!textFrame.textRange.characterAttributes.autoLeading) return 0;
            if (textFrame.paragraphs.length > 0) {
                return textFrame.paragraphs[0].paragraphAttributes.autoLeadingAmount || 0;
            }
        } catch (e) { }
        return 0;
    }

    /* geometricBounds の中心座標を返す / Center point of geometricBounds */
    function getBoundsCenter(item) {
        var bounds = item.geometricBounds; // [left, top, right, bottom]
        return [(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2];
    }

    /* テキストの中心を含む図形、無ければ中心が最も近い図形を返す（未使用のものだけ）/ Shape containing the text center, else the nearest unused one */
    function findShapeForText(sourceText, shapePaths, usedShapes) {
        var textCenter = getBoundsCenter(sourceText);
        var nearestShape = null, nearestDistance = -1;

        for (var i = 0; i < shapePaths.length; i++) {
            var shapePath = shapePaths[i];

            var isUsed = false;
            for (var j = 0; j < usedShapes.length; j++) {
                if (usedShapes[j] === shapePath) { isUsed = true; break; }
            }
            if (isUsed) continue;

            var shapeBounds = shapePath.geometricBounds;
            if (textCenter[0] >= shapeBounds[0] && textCenter[0] <= shapeBounds[2] &&
                textCenter[1] <= shapeBounds[1] && textCenter[1] >= shapeBounds[3]) {
                return shapePath;
            }

            var shapeCenter = getBoundsCenter(shapePath);
            var dx = shapeCenter[0] - textCenter[0], dy = shapeCenter[1] - textCenter[1];
            var distance = dx * dx + dy * dy;
            if (nearestDistance < 0 || distance < nearestDistance) {
                nearestDistance = distance;
                nearestShape = shapePath;
            }
        }
        return nearestShape;
    }

    /* 選択内容を指定の方法でエリア内文字に変換し、作成したフレームを返す / Convert the selection to Area Type and return the new frames */
    function convertSelectionToAreaText(doc, selection, useSimple, useButton, useShape, useShapeDummy) {
        var createdFrames = [];

        if (useShapeDummy) { createdFrames = fillShapesWithDummyText(doc, selection); }
        else if (useShape) { createdFrames = pourTextsIntoShapes(doc, selection); }
        else if (useSimple) { createdFrames = convertPointTexts(doc, selection, AREA_TEXT_STYLE_SIMPLE); }
        else if (useButton) { createdFrames = convertPointTexts(doc, selection, AREA_TEXT_STYLE_BUTTON); }

        if (createdFrames.length > 0) {
            app.activeDocument.selection = createdFrames;
            app.redraw();
        }
        return createdFrames;
    }

    /* 選択した閉じたパスをエリア内文字にしてダミー文字を流し込む / Turn selected closed paths into Area Type filled with dummy text */
    function fillShapesWithDummyText(doc, selection) {
        var createdFrames = [];
        var dummyText = (currentLanguage === "ja") ? DUMMY_TEXT_JA : DUMMY_TEXT_EN;
        var dummyFont = findAvailableTextFont((currentLanguage === "ja") ? DUMMY_FONT_JA : DUMMY_FONT_EN);

        for (var i = 0; i < selection.length; i++) {
            var closedPath = getClosedPathItem(selection[i]);
            if (!closedPath) continue;
            try {
                var dummyFrame = createAreaTextFromPath(doc, closedPath);
                dummyFrame.contents = dummyText;
                applyFontAndSize(dummyFrame, { font: dummyFont, size: DUMMY_FONT_SIZE });
                createdFrames.push(dummyFrame);
            } catch (e) { }
        }
        return createdFrames;
    }

    /* 選択テキストを、対応する閉じたパスのエリア内文字に流し込む / Pour each selected text into its matching closed path */
    function pourTextsIntoShapes(doc, selection) {
        var createdFrames = [];
        var sourceTexts = [], shapePaths = [];
        for (var i = 0; i < selection.length; i++) {
            if (selection[i].typename === "TextFrame") { sourceTexts.push(selection[i]); }
            else if (selection[i].typename === "PathItem" && selection[i].closed) { shapePaths.push(selection[i]); }
        }

        var usedShapes = [];
        for (var i = 0; i < sourceTexts.length; i++) {
            var sourceText = sourceTexts[i];
            var targetPath = findShapeForText(sourceText, shapePaths, usedShapes);
            if (!targetPath) continue;
            usedShapes.push(targetPath);
            try {
                var contents = sourceText.contents;
                var fontAndSize = getFontAndSize(sourceText);
                var areaFrame = createAreaTextFromPath(doc, targetPath.duplicate());
                areaFrame.contents = contents;
                applyFontAndSize(areaFrame, fontAndSize);
                sourceText.remove();
                targetPath.remove();
                createdFrames.push(areaFrame);
            } catch (e) { }
        }
        return createdFrames;
    }

    /* ポイント文字を、その大きさから起こした長方形のエリア内文字に置き換える / Replace point text with Area Type built from a rectangle around it */
    function convertPointTexts(doc, selection, style) {
        var createdFrames = [];
        for (var i = selection.length - 1; i >= 0; i--) {
            var pointText = selection[i];
            if (pointText.typename !== "TextFrame" || pointText.kind !== TextType.POINTTEXT) continue;
            try {
                var box = buildFrameBox(pointText.geometricBounds, style);
                var contents = pointText.contents;
                var fontAndSize = getFontAndSize(pointText);

                var frameRect = doc.pathItems.rectangle(box.top, box.left, box.width, box.height);
                var areaFrame = createAreaTextFromPath(doc, frameRect);
                areaFrame.contents = contents;
                applyFontAndSize(areaFrame, fontAndSize);

                if (style.centerText) {
                    areaFrame.textRange.paragraphAttributes.justification = Justification.CENTER;
                    // DOM の verticalAlignment はエリア内文字で効かないことがあるため、
                    // 調整ダイアログと同じダイナミックアクション（adobe_frameAlignment）で中央に寄せる
                    applyAreaTextFrameAlignment(areaFrame, 1, false);
                }

                createdFrames.push(areaFrame);
                pointText.remove();
            } catch (e) { }
        }
        return createdFrames;
    }

    /* スタイル定義からフレーム矩形の位置とサイズを求める / Work out the frame rectangle from a style definition */
    function buildFrameBox(bounds, style) {
        var origWidth = bounds[2] - bounds[0], origHeight = bounds[1] - bounds[3];
        var width = origWidth * style.widthScale + style.widthPadPt;
        var height = origHeight * style.heightScale + style.heightPadPt;
        return {
            top: style.centerOnText ? bounds[1] + (height - origHeight) / 2 : bounds[1],
            left: style.centerOnText ? bounds[0] - (width - origWidth) / 2 : bounds[0],
            width: width,
            height: height
        };
    }

    // ============================================================
    // 調整ダイアログ：エリア内文字 調整
    // ============================================================
    function showAdjustDialog(doc, initialFrame, targetFrames, initialAlignmentValue) {
        var rulerInfo = getRulerUnitInfo(doc);

        // 変換ダイアログから受け取った変換結果を確実に対象にする（選択が変わっても崩れないように）
        if (targetFrames && targetFrames.length) {
            try { app.activeDocument.selection = targetFrames; } catch (e) { }
        } else if (initialFrame) {
            try { app.activeDocument.selection = [initialFrame]; } catch (e2) { }
        }
        try { app.redraw(); } catch (e3) { }

        // 調整対象（非分離モード用に固定）。
        // モーダルダイアログ中は selection が変動/取得不能になることがあるため、変換ダイアログから渡された配列を優先する。
        var targetAreaFrames = null;
        if (targetFrames && targetFrames.length) {
            targetAreaFrames = targetFrames.slice(0);
        } else {
            try {
                var currentSelection = app.activeDocument.selection;
                if (currentSelection && currentSelection.length) {
                    targetAreaFrames = [];
                    for (var i = 0; i < currentSelection.length; i++) { targetAreaFrames.push(currentSelection[i]); }
                }
            } catch (e4) { }
        }

        /* 現在の選択からエリア内文字だけを拾い直す / Re-pick just the Area Type frames from the current selection */
        function refreshTargetAreaFrames() {
            try {
                var selectedItems = app.activeDocument.selection;
                var areaFrames = [];
                if (selectedItems && selectedItems.length) {
                    for (var i = 0; i < selectedItems.length; i++) {
                        if (selectedItems[i] && selectedItems[i].typename === "TextFrame" && selectedItems[i].kind === TextType.AREATEXT) {
                            areaFrames.push(selectedItems[i]);
                        }
                    }
                }
                targetAreaFrames = areaFrames.length ? areaFrames : null;
            } catch (e) {
                targetAreaFrames = null;
            }
        }

        /* 調整対象を解決する（固定ターゲット優先、無ければ選択）/ Resolve the frames to adjust (fixed targets first, selection as fallback) */
        function getTargetAreaFrames() {
            if (!(targetAreaFrames && targetAreaFrames.length)) { refreshTargetAreaFrames(); }
            return (targetAreaFrames && targetAreaFrames.length) ? targetAreaFrames : app.activeDocument.selection;
        }

        var adjustDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        adjustDialog.alignChildren = "fill";
        adjustDialog.margins = 20;

        // 「テキストを分離」パネル / "Separate text" panel
        var separateTextPanel = adjustDialog.add("panel", undefined, getLabel("panel.separateText"));
        applyPanelLayout(separateTextPanel);
        separateTextPanel.orientation = "row";
        separateTextPanel.alignChildren = ["center", "center"];

        var radNoAdjust = separateTextPanel.add("radiobutton", undefined, getLabel("radio.noAdjust"));
        var radStrokeBlack = separateTextPanel.add("radiobutton", undefined, getLabel("radio.strokeBlack"));
        var radHidePath = separateTextPanel.add("radiobutton", undefined, getLabel("radio.hidePath"));
        var radRemovePath = separateTextPanel.add("radiobutton", undefined, getLabel("radio.removePath"));
        radNoAdjust.value = true;

        // 2カラムレイアウト（左右カラムは上揃えで横いっぱいに）/ Two-column layout (columns fill width, top-aligned)
        var columnsGroup = adjustDialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.spacing = 10;

        var leftColumn = columnsGroup.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = "fill";

        var rightColumn = columnsGroup.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = "fill";

        // 右カラム：行送り（自動行送り量％と、その実寸）/ Right column: leading (auto-leading % and its effective size)
        var leadingPanel = rightColumn.add("panel", undefined, getLabel("panel.leading"));
        applyPanelLayout(leadingPanel);
        var leadingPercentRow = leadingPanel.add("group");
        var lblLeadingPercent = leadingPercentRow.add("statictext", undefined, labelWithColon("label.leadingPercent"));
        lblLeadingPercent.preferredSize.width = 48;
        var etLeadingPercent = leadingPercentRow.add("edittext", undefined, "");
        etLeadingPercent.characters = 4;
        leadingPercentRow.add("statictext", undefined, "%");
        var leadingEffectiveRow = leadingPanel.add("group");
        var lblLeadingEffective = leadingEffectiveRow.add("statictext", undefined, labelWithColon("label.leadingEffective"));
        lblLeadingEffective.preferredSize.width = 48;
        var etLeadingEffective = leadingEffectiveRow.add("edittext", undefined, "");
        etLeadingEffective.characters = 4;
        leadingEffectiveRow.add("statictext", undefined, "pt");

        // 右カラム：行揃え / Right column: justification
        var justificationPanel = rightColumn.add("panel", undefined, getLabel("panel.justification"));
        applyPanelLayout(justificationPanel);
        var radJustifyLeft = justificationPanel.add("radiobutton", undefined, getLabel("radio.justifyLeft"));
        var radJustifyCenter = justificationPanel.add("radiobutton", undefined, getLabel("radio.justifyCenter"));
        var radJustifyRight = justificationPanel.add("radiobutton", undefined, getLabel("radio.justifyRight"));
        var radJustifyLastLineLeft = justificationPanel.add("radiobutton", undefined, getLabel("radio.justifyLastLineLeft"));
        var radJustifyAllLines = justificationPanel.add("radiobutton", undefined, getLabel("radio.justifyAllLines"));
        radJustifyLeft.value = true;

        // 右カラム：テキストの配置 / Right column: text alignment
        var textAlignPanel = rightColumn.add("panel", undefined, getLabel("panel.textAlign"));
        applyPanelLayout(textAlignPanel);
        var radAlignTop = textAlignPanel.add("radiobutton", undefined, getLabel("radio.alignTop"));
        var radAlignCenter = textAlignPanel.add("radiobutton", undefined, getLabel("radio.alignCenter"));
        var radAlignBottom = textAlignPanel.add("radiobutton", undefined, getLabel("radio.alignBottom"));
        var radAlignJustify = textAlignPanel.add("radiobutton", undefined, getLabel("radio.alignJustify"));
        // Default alignment
        if (initialAlignmentValue === 1) {
            radAlignCenter.value = true;
        } else if (initialAlignmentValue === 2) {
            radAlignBottom.value = true;
        } else if (initialAlignmentValue === 3) {
            radAlignJustify.value = true;
        } else {
            radAlignTop.value = true;
        }

        // 左カラム：フォントサイズ / Left column: font size
        var fontSizePanel = leftColumn.add("panel", undefined, getLabel("panel.fontSize"));
        applyPanelLayout(fontSizePanel);
        var fontSizeRow = fontSizePanel.add("group");
        fontSizeRow.alignment = "left";
        fontSizeRow.add("statictext", undefined, labelWithColon("label.fontSize"));
        var etFontSize = fontSizeRow.add("edittext", undefined, "");
        etFontSize.characters = 4;
        fontSizeRow.add("statictext", undefined, "pt");
        var fontSizeButtonRow = fontSizePanel.add("group");
        fontSizeButtonRow.orientation = "row";
        var btnShrinkToFit = fontSizeButtonRow.add("button", undefined, getLabel("button.shrinkToFit"));
        var btnFitFontSize = fontSizeButtonRow.add("button", undefined, getLabel("button.fitFontSize"));

        // 左カラム：フレームサイズ / Left column: frame size
        var frameSizePanel = leftColumn.add("panel", undefined, getLabel("panel.frameSize"));
        applyPanelLayout(frameSizePanel);
        var widthRow = frameSizePanel.add("group");
        var lblWidth = widthRow.add("statictext", undefined, labelWithColon("label.width"));
        lblWidth.preferredSize.width = 28;
        var etWidth = widthRow.add("edittext", undefined, "");
        etWidth.characters = 5;
        widthRow.add("statictext", undefined, rulerInfo.label);
        var etCharsPerLine = widthRow.add("edittext", undefined, "");
        etCharsPerLine.characters = 4;
        var lblCharsPerLine = widthRow.add("statictext", undefined, getLabel("label.chars"));
        var heightRow = frameSizePanel.add("group");
        var lblHeight = heightRow.add("statictext", undefined, labelWithColon("label.height"));
        lblHeight.preferredSize.width = 28;
        var etHeight = heightRow.add("edittext", undefined, "");
        etHeight.characters = 5;
        heightRow.add("statictext", undefined, rulerInfo.label);
        var btnAutoSize = heightRow.add("button", undefined, getLabel("button.autoSize"));
        btnAutoSize.preferredSize.width = 60;
        // 英語UIでは chars 計算は不正確なため使用不可にする
        if (currentLanguage !== "ja") {
            etCharsPerLine.enabled = false;
            lblCharsPerLine.enabled = false;
        }

        // 左カラム：インデント / Left column: indent
        var indentPanel = leftColumn.add("panel", undefined, getLabel("panel.indent"));
        applyPanelLayout(indentPanel);
        indentPanel.orientation = "row";
        indentPanel.alignChildren = ["left", "top"];
        var indentFieldsColumn = indentPanel.add("group");
        indentFieldsColumn.orientation = "column";
        indentFieldsColumn.alignChildren = "left";
        var leftIndentRow = indentFieldsColumn.add("group");
        var chkLeftIndent = leftIndentRow.add("checkbox", undefined, labelWithColon("label.indentLeft"));
        chkLeftIndent.preferredSize.width = 40;
        var etLeftIndent = leftIndentRow.add("edittext", undefined, "0");
        etLeftIndent.characters = 4;
        leftIndentRow.add("statictext", undefined, rulerInfo.label);
        etLeftIndent.enabled = false;
        var rightIndentRow = indentFieldsColumn.add("group");
        var chkRightIndent = rightIndentRow.add("checkbox", undefined, labelWithColon("label.indentRight"));
        chkRightIndent.preferredSize.width = 40;
        var etRightIndent = rightIndentRow.add("edittext", undefined, "0");
        etRightIndent.characters = 4;
        rightIndentRow.add("statictext", undefined, rulerInfo.label);
        etRightIndent.enabled = false;
        var linkColumn = indentPanel.add("group");
        linkColumn.orientation = "column";
        linkColumn.alignChildren = "left";
        linkColumn.alignment = ["left", "center"];
        var chkLinkIndents = linkColumn.add("checkbox", undefined, getLabel("checkbox.linkIndents"));

        // 左カラム：オプション / Left column: options
        var optionsPanel = leftColumn.add("panel", undefined, getLabel("panel.options"));
        applyPanelLayout(optionsPanel);
        var spacingRow = optionsPanel.add("group");
        var chkSpacing = spacingRow.add("checkbox", undefined, getLabel("checkbox.spacing"));
        var etSpacing = spacingRow.add("edittext", undefined, "0");
        etSpacing.characters = 4;
        var lblSpacingUnit = spacingRow.add("statictext", undefined, rulerInfo.label);
        etSpacing.enabled = false;
        lblSpacingUnit.enabled = false;
        var chkLeaderTabs = optionsPanel.add("checkbox", undefined, getLabel("checkbox.leaderTabs"));

        // ボタンエリア
        var bottomBar = adjustDialog.add("group");
        bottomBar.orientation = "row";
        bottomBar.alignment = "fill";
        bottomBar.alignChildren = ["fill", "center"];
        var chkPreview = bottomBar.add("checkbox", undefined, getLabel("checkbox.preview"));
        chkPreview.alignment = ["left", "center"];
        var bottomButtonRow = bottomBar.add("group");
        bottomButtonRow.alignment = ["right", "center"];
        var btnCancelAdjust = bottomButtonRow.add("button", undefined, getLabel("button.close"), { name: "cancel" });
        var btnRun = bottomButtonRow.add("button", undefined, getLabel("button.run"), { name: "ok" });

        // 状態変数
        var isPreviewActive = false;
        var fontFitMode = "none";
        var hasMultiParagraph = false;
        var isAutoSizeOn = false;
        var currentFontSize = 0;

        /* 入力バリデーション（幅/高さ）/ Input validation (width/height) */
        // 0以下・NaN・極端値を弾いてIllustratorの不安定化を避ける
        var lastValidWidth = null; // ruler units
        var lastValidHeight = null; // ruler units

        /* 幅・高さの上限を定規単位で返す / Upper limit for width and height, in ruler units */
        function getMaxSizeInRulerUnits() {
            // 上限は pt で固定（極端値を防ぐ）。表示単位に合わせて換算。
            // 100000pt は現実的に十分大きく、かつ事故りにくい上限。
            return 100000 / rulerInfo.toPt;
        }

        /* 幅・高さ欄の値を検証して返す（不正なら null）/ Validate a width/height field and return its value (null when invalid) */
        function validateSizeField(editText, lastValue) {
            var raw = String(editText.text);
            var value = parseFloat(raw);
            if (isNaN(value) || !isFinite(value)) {
                if (lastValue !== null) editText.text = lastValue;
                return null;
            }
            if (value <= 0) {
                if (lastValue !== null) editText.text = lastValue;
                return null;
            }
            var maxValue = getMaxSizeInRulerUnits();
            if (value > maxValue) {
                value = maxValue;
                editText.text = Math.round(value * 100) / 100;
            }
            // 小さすぎる値も事故の元なので下限を設ける
            if (value < 0.01) {
                value = 0.01;
                editText.text = Math.round(value * 100) / 100;
            }
            return value;
        }

        var separationRadios = [radStrokeBlack, radHidePath, radRemovePath, radNoAdjust];
        var justificationRadios = [radJustifyLeft, radJustifyCenter, radJustifyRight, radJustifyLastLineLeft, radJustifyAllLines];
        var alignmentRadios = [radAlignTop, radAlignCenter, radAlignBottom, radAlignJustify];

        /* ローカル関数 / Local functions */

        /* 行揃えラジオを選び、右揃え以外ならリーダー罫を解除する / Select a justification radio; drop leader tabs unless right-aligned */
        function selectJustificationRadio(targetRadio) {
            selectRadio(justificationRadios, targetRadio);
            if (chkLeaderTabs.value && targetRadio !== radJustifyRight) { chkLeaderTabs.value = false; }
        }

        /* 自動サイズボタンのON/OFF表示を更新する / Refresh the on/off marker on the auto-size button */
        function updateAutoSizeButtonLabel() {
            btnAutoSize.text = (isAutoSizeOn ? "✓ " : "") + getLabel("button.autoSize");
        }

        /* 幅から差し引く余白（間隔×2＋左右インデント）/ Horizontal space taken out of the width (spacing x2 + both indents) */
        function getWidthAdjustmentPt() {
            var spacingPt = chkSpacing.value ? (parseFloat(etSpacing.text) || 0) * rulerInfo.toPt : 0;
            var leftIndentPt = (chkLeftIndent.value || chkLinkIndents.value) ? (parseFloat(etLeftIndent.text) || 0) * rulerInfo.toPt : 0;
            var rightIndentPt = chkLinkIndents.value ? leftIndentPt : (chkRightIndent.value ? (parseFloat(etRightIndent.text) || 0) * rulerInfo.toPt : 0);
            return 2 * spacingPt + leftIndentPt + rightIndentPt;
        }

        /* 選択フレームの現在値をダイアログに読み込む / Load the frame's current values into the dialog */
        function loadValuesFromFrame(sourceFrame) {
            try { hasMultiParagraph = (sourceFrame.paragraphs && sourceFrame.paragraphs.length >= 2); } catch (e) { hasMultiParagraph = false; }
            var frameWidth = sourceFrame.textPath.width / rulerInfo.toPt;
            var frameHeight = sourceFrame.textPath.height / rulerInfo.toPt;
            currentFontSize = 0;
            try { currentFontSize = sourceFrame.textRange.characterAttributes.size || 0; } catch (e) { }
            if (currentFontSize > 0) { etFontSize.text = Math.round(currentFontSize * 100) / 100; }
            etWidth.text = Math.round(frameWidth * 100) / 100;
            etHeight.text = Math.round(frameHeight * 100) / 100;
            lastValidWidth = parseFloat(etWidth.text);
            lastValidHeight = parseFloat(etHeight.text);
            try {
                var justification = sourceFrame.paragraphs.length > 0
                    ? sourceFrame.paragraphs[0].paragraphAttributes.justification
                    : Justification.LEFT;
                radJustifyLeft.value = (justification === Justification.LEFT);
                radJustifyCenter.value = (justification === Justification.CENTER);
                radJustifyRight.value = (justification === Justification.RIGHT);
                radJustifyLastLineLeft.value = (justification === Justification.FULLJUSTIFYLASTLINELEFT);
                radJustifyAllLines.value = (justification === Justification.FULLJUSTIFY);
                if (!radJustifyLeft.value && !radJustifyCenter.value && !radJustifyRight.value &&
                    !radJustifyLastLineLeft.value && !radJustifyAllLines.value) { radJustifyLeft.value = true; }
            } catch (e) { radJustifyLeft.value = true; }
            try {
                var spacingPt = sourceFrame.spacing || 0;
                etSpacing.text = Math.round((spacingPt / rulerInfo.toPt) * 100) / 100;
                chkSpacing.value = (spacingPt !== 0);
                etSpacing.enabled = chkSpacing.value;
                lblSpacingUnit.enabled = chkSpacing.value;
            } catch (e) { }
            try {
                var firstParaAttrs = sourceFrame.paragraphs.length > 0 ? sourceFrame.paragraphs[0].paragraphAttributes : null;
                var leftIndentPt = firstParaAttrs ? (firstParaAttrs.leftIndent || 0) : 0;
                var rightIndentPt = firstParaAttrs ? (firstParaAttrs.rightIndent || 0) : 0;
                chkLeftIndent.value = (leftIndentPt !== 0);
                etLeftIndent.enabled = chkLeftIndent.value;
                etLeftIndent.text = chkLeftIndent.value ? Math.round((leftIndentPt / rulerInfo.toPt) * 100) / 100 : "0";
                chkRightIndent.value = (rightIndentPt !== 0);
                etRightIndent.enabled = chkRightIndent.value;
                etRightIndent.text = chkRightIndent.value ? Math.round((rightIndentPt / rulerInfo.toPt) * 100) / 100 : "0";
            } catch (e) { }
            var leadingPercent = getAutoLeadingPercent(sourceFrame);
            etLeadingPercent.text = (leadingPercent > 0) ? Math.round(leadingPercent * 10) / 10 : "";
            updateLeadingEffective();

            if (currentFontSize > 0) {
                etCharsPerLine.text = Math.round(((sourceFrame.textPath.width - getWidthAdjustmentPt()) / currentFontSize) * 100) / 100;
            }
            btnShrinkToFit.enabled = !hasMultiParagraph;
            btnFitFontSize.enabled = !hasMultiParagraph;
        }

        /* ダイアログの入力を1つの設定オブジェクトにまとめる / Collect the dialog inputs into one settings object */
        function readAdjustmentSettings() {
            var justificationValue = Justification.LEFT;
            if (radJustifyCenter.value) justificationValue = Justification.CENTER;
            else if (radJustifyRight.value) justificationValue = Justification.RIGHT;
            else if (radJustifyLastLineLeft.value) justificationValue = Justification.FULLJUSTIFYLASTLINELEFT;
            else if (radJustifyAllLines.value) justificationValue = Justification.FULLJUSTIFY;

            var alignmentValue = 0;
            if (radAlignCenter.value) alignmentValue = 1;
            else if (radAlignBottom.value) alignmentValue = 2;
            else if (radAlignJustify.value) alignmentValue = 3;

            var leftIndentPt = (chkLeftIndent.value || chkLinkIndents.value)
                ? (parseFloat(etLeftIndent.text) || 0) * rulerInfo.toPt : 0;

            return {
                strokeBlack: radStrokeBlack.value,
                hidePath: radHidePath.value,
                removePath: radRemovePath.value,
                separate: (radStrokeBlack.value || radHidePath.value || radRemovePath.value),
                shrinkFont: (fontFitMode === "shrink"),
                fitFont: (fontFitMode === "fit"),
                autoSize: isAutoSizeOn,
                leaderTabs: chkLeaderTabs.value,
                justification: justificationValue,
                alignment: alignmentValue,
                leadingPercent: parseFloat(etLeadingPercent.text),
                leftIndentPt: leftIndentPt,
                rightIndentPt: chkLinkIndents.value ? leftIndentPt
                    : (chkRightIndent.value ? (parseFloat(etRightIndent.text) || 0) * rulerInfo.toPt : 0),
                spacingPt: chkSpacing.value ? (parseFloat(etSpacing.text) || 0) * rulerInfo.toPt : 0
            };
        }

        /* 段落の行揃えとインデントを適用する / Apply justification and indents to every paragraph */
        function applyParagraphSettings(textFrame, settings) {
            try {
                var paraAttrs = textFrame.textRange.paragraphAttributes;
                paraAttrs.justification = settings.justification;
                paraAttrs.leftIndent = settings.leftIndentPt;
                paraAttrs.rightIndent = settings.rightIndentPt;
            } catch (e) { }
            applyAutoLeading(textFrame, settings.leadingPercent);
        }

        /* エリア内文字を、囲み罫とポイント文字に分解する / Replace Area Type with a rectangle plus point text */
        function separateAreaTextFrame(areaTextFrame, settings) {
            var bounds = areaTextFrame.geometricBounds;
            var spacingPt = settings.spacingPt;
            var left = bounds[0] - spacingPt, top = bounds[1] + spacingPt;
            var right = bounds[2] + spacingPt, bottom = bounds[3] - spacingPt;

            var frameRect = doc.pathItems.rectangle(top, left, right - left, top - bottom);
            frameRect.filled = false;
            frameRect.stroked = true;

            var pointTextFrame = doc.textFrames.add();
            pointTextFrame.contents = areaTextFrame.contents;

            // 1行目のベースラインにそろえる。空フレームではサイズを取得できないことがある
            var firstLineOffset = 0;
            try { firstLineOffset = pointTextFrame.textRange.characterAttributes.size; } catch (e0) { }
            pointTextFrame.position = [left, top - firstLineOffset];

            try {
                var srcAttrs = areaTextFrame.textRange.characterAttributes;
                var destAttrs = pointTextFrame.textRange.characterAttributes;
                if (srcAttrs.textFont && srcAttrs.textFont.name) { destAttrs.textFont = srcAttrs.textFont; }
                destAttrs.size = srcAttrs.size;
                destAttrs.leading = srcAttrs.leading;
            } catch (e1) { }

            areaTextFrame.remove();

            if (settings.strokeBlack) {
                var blackColor = new CMYKColor();
                blackColor.cyan = 0; blackColor.magenta = 0; blackColor.yellow = 0; blackColor.black = 100;
                frameRect.strokeColor = blackColor;
                frameRect.strokeWidth = 1;
            } else if (settings.hidePath) {
                frameRect.stroked = false;
            } else if (settings.removePath) {
                frameRect.remove();
            }

            applyParagraphSettings(pointTextFrame, settings);
            if (settings.leaderTabs) { applyLeaderTabStops(pointTextFrame); }
        }

        /* エリア内文字のフレームサイズ・行揃え・配置などを適用する / Apply frame size, justification and placement to Area Type */
        function adjustAreaTextFrame(areaTextFrame, settings, forPreview) {
            // 幅/高さ：NaN・0以下・極端値をガード
            var widthValue = validateSizeField(etWidth, lastValidWidth);
            var heightValue = validateSizeField(etHeight, lastValidHeight);
            if (widthValue !== null) { lastValidWidth = widthValue; }
            if (heightValue !== null) { lastValidHeight = heightValue; }
            try {
                areaTextFrame.spacing = settings.spacingPt;
                if (widthValue !== null) { areaTextFrame.textPath.width = widthValue * rulerInfo.toPt; }
                if (heightValue !== null) { areaTextFrame.textPath.height = heightValue * rulerInfo.toPt; }
            } catch (e) { }

            if (settings.shrinkFont) { shrinkFontToFit(areaTextFrame); }
            else if (settings.fitFont) { fitFontSizeToFrame(areaTextFrame); }

            // プレビュー中は app.doScript 経由の処理を走らせない（不安定化・クラッシュ回避）
            if (settings.autoSize && !forPreview) { enableFrameAutoSize(areaTextFrame); }

            applyParagraphSettings(areaTextFrame, settings);
            applyAreaTextFrameAlignment(areaTextFrame, settings.alignment, forPreview);
            if (settings.leaderTabs) { applyLeaderTabStops(areaTextFrame); }
        }

        /* 対象のエリア内文字すべてに現在の設定を適用する / Apply the current settings to every target Area Type frame */
        function applyAdjustments(forPreview) {
            var settings = readAdjustmentSettings();

            // 分離モードはオブジェクトが置き換わるため selection を使う。
            // 非分離モードは固定ターゲットを優先して、プレビューが確実に反映されるようにする。
            var savedSelection = [];
            var framesToAdjust;
            if (settings.separate) {
                framesToAdjust = app.activeDocument.selection;
            } else {
                var previousSelection = app.activeDocument.selection;
                for (var i = 0; i < previousSelection.length; i++) { savedSelection.push(previousSelection[i]); }
                framesToAdjust = getTargetAreaFrames();
            }

            for (var i = framesToAdjust.length - 1; i >= 0; i--) {
                var areaTextFrame = framesToAdjust[i];
                if (areaTextFrame.typename !== "TextFrame" || areaTextFrame.kind !== TextType.AREATEXT) continue;
                // 1フレームで失敗しても残りを処理できるようにする
                try {
                    if (settings.separate) { separateAreaTextFrame(areaTextFrame, settings); }
                    else { adjustAreaTextFrame(areaTextFrame, settings, forPreview); }
                } catch (e) { }
            }

            if (savedSelection.length > 0) { app.activeDocument.selection = savedSelection; }
            app.redraw();
        }

        /* 適用中のプレビューを取り消して未適用の状態に戻す / Revert the active preview back to the unapplied state */
        function revertPreview() {
            if (!isPreviewActive) return;
            try { app.undo(); } catch (e) { }
            try { app.redraw(); } catch (e2) { }
            isPreviewActive = false;

            // undo 後は参照が無効化されるため、対象を取り直す
            refreshTargetAreaFrames();
        }

        /* プレビューを貼り直す / Re-apply the preview */
        function updatePreview() {
            revertPreview();
            if (chkPreview.value) {
                applyAdjustments(true);
                isPreviewActive = true;
            }
        }

        /* 分離モードのあいだは調整系パネルを使用不可にする / Disable the adjustment panels while separation mode is on */
        function updateSeparationModeEnabled() {
            var isSeparationMode = radStrokeBlack.value || radHidePath.value || radRemovePath.value;
            justificationPanel.enabled = !isSeparationMode;
            textAlignPanel.enabled = !isSeparationMode;
            fontSizePanel.enabled = !isSeparationMode;
            frameSizePanel.enabled = !isSeparationMode;
            btnAutoSize.enabled = !isSeparationMode;
            indentPanel.enabled = !isSeparationMode;
            optionsPanel.enabled = !isSeparationMode;
            if (fontSizePanel.enabled) { btnShrinkToFit.enabled = !hasMultiParagraph; btnFitFontSize.enabled = !hasMultiParagraph; }
            else { btnShrinkToFit.enabled = false; btnFitFontSize.enabled = false; }
        }

        /* 「外側からの間隔」の使用可否を更新する / Update whether the spacing field can be used */
        function updateSpacingEnabled() {
            chkSpacing.enabled = radNoAdjust.value;
            if (!radNoAdjust.value) { etSpacing.enabled = false; lblSpacingUnit.enabled = false; }
            else { etSpacing.enabled = chkSpacing.value; lblSpacingUnit.enabled = chkSpacing.value; }
        }

        /* 実質行送り（フォントサイズ×％）の表示を更新する / Refresh the effective-leading display (font size × %) */
        function updateLeadingEffective() {
            var size = parseFloat(etFontSize.text);
            var percent = parseFloat(etLeadingPercent.text);
            etLeadingEffective.text = (isNaN(size) || isNaN(percent))
                ? "" : Math.round(size * percent / 100 * 10) / 10;
        }

        /* 行送り％の変更を反映する / Reflect a change to the leading percentage */
        function onLeadingPercentChange() {
            updateLeadingEffective();
            updatePreview();
        }

        /* 実質行送りの入力から行送り％を逆算する / Back-calculate the leading % from the effective value */
        function onLeadingEffectiveChange() {
            var effective = parseFloat(etLeadingEffective.text);
            var size = parseFloat(etFontSize.text);
            if (isNaN(effective) || isNaN(size) || size <= 0) return;
            etLeadingPercent.text = Math.round((effective / size) * 100 * 10) / 10;
            updatePreview();
        }

        /* フォントサイズ欄の値を対象フレームに適用する / Apply the font-size field to the target frames */
        function applyFontSizeFromField() {
            var newSize = parseFloat(etFontSize.text) || 0;
            if (newSize <= 0) return;
            currentFontSize = newSize;
            var framesToAdjust = getTargetAreaFrames();
            for (var i = 0; i < framesToAdjust.length; i++) {
                if (framesToAdjust[i].typename === "TextFrame" && framesToAdjust[i].kind === TextType.AREATEXT) {
                    try { framesToAdjust[i].textRange.characterAttributes.size = newSize; } catch (e) { }
                }
            }
            if (currentFontSize > 0) {
                etCharsPerLine.text = Math.round((((parseFloat(etWidth.text) || 0) * rulerInfo.toPt - getWidthAdjustmentPt()) / currentFontSize) * 100) / 100;
            }
            updateLeadingEffective();
            app.redraw();
        }

        /* 幅の変更を検証し、文字数表示とプレビューを更新する / Validate a width change, then refresh the chars-per-line field and the preview */
        function onWidthChange() {
            var widthValue = validateSizeField(etWidth, lastValidWidth);
            if (widthValue !== null) { lastValidWidth = widthValue; }
            if (currentFontSize > 0) {
                etCharsPerLine.text = Math.round((((parseFloat(etWidth.text) || 0) * rulerInfo.toPt - getWidthAdjustmentPt()) / currentFontSize) * 100) / 100;
            }
            updatePreview();
        }
        /* 1行の文字数から幅を逆算する / Work the width back out from the characters-per-line value */
        function onCharsPerLineChange() {
            if (currentFontSize > 0) {
                var nextWidth = (((parseFloat(etCharsPerLine.text) || 0) * currentFontSize + getWidthAdjustmentPt()) / rulerInfo.toPt);
                if (!isNaN(nextWidth) && isFinite(nextWidth) && nextWidth > 0) {
                    etWidth.text = Math.round(nextWidth * 100) / 100;
                    var widthValue = validateSizeField(etWidth, lastValidWidth);
                    if (widthValue !== null) { lastValidWidth = widthValue; }
                }
            }
            updatePreview();
        }
        /* インデント・間隔の変更を幅の再計算に回す / Feed indent and spacing changes into the width recalculation */
        function onIndentOrSpacingChange() {
            if (chkLinkIndents.value) { etRightIndent.text = etLeftIndent.text; }
            onWidthChange();
        }

        /* L/C/R/J/F キーで行揃えを切り替える / Switch justification with the L / C / R / J / F keys */
        function addJustificationShortcuts(dialog) {
            dialog.addEventListener("keydown", function (event) {
                if (!event || !event.keyName) return;
                // 入力欄にフォーカスがあるときは通常の文字入力を優先する
                try { if (event.target && event.target.type === "edittext") return; } catch (e0) { }
                var keyName = String(event.keyName).toUpperCase();
                var targetRadio = null;
                if (keyName === "L") targetRadio = radJustifyLeft;
                else if (keyName === "C") targetRadio = radJustifyCenter;
                else if (keyName === "R") targetRadio = radJustifyRight;
                else if (keyName === "J") targetRadio = radJustifyLastLineLeft;
                else if (keyName === "F") targetRadio = radJustifyAllLines;
                if (!targetRadio) return;
                selectJustificationRadio(targetRadio);
                event.preventDefault();
                updatePreview();
            });
        }

        /* イベントハンドラ / Event handlers */
        bindExclusiveRadios(separationRadios, function () {
            updateSeparationModeEnabled();
            updateSpacingEnabled();
            updatePreview();
        });
        bindExclusiveRadios(justificationRadios, function (radio) {
            selectJustificationRadio(radio);
            updatePreview();
        });
        bindExclusiveRadios(alignmentRadios, updatePreview);
        addJustificationShortcuts(adjustDialog);

        /* フォントサイズ自動調整を本適用する（プレビュー分は先に取り消す）/ Commit a font-fit pass (revert the preview first) */
        function runFontFit(mode) {
            revertPreview();
            fontFitMode = mode;
            applyAdjustments(false);
            fontFitMode = "none";
        }

        btnShrinkToFit.onClick = function () { runFontFit("shrink"); };
        btnFitFontSize.onClick = function () { runFontFit("fit"); };
        chkPreview.onClick = updatePreview;
        btnAutoSize.onClick = function () {
            isAutoSizeOn = !isAutoSizeOn;
            updateAutoSizeButtonLabel();
            if (!isAutoSizeOn) {
                try {
                    var framesToAdjust = getTargetAreaFrames();
                    for (var i = 0; i < framesToAdjust.length; i++) {
                        if (framesToAdjust[i] && framesToAdjust[i].typename === "TextFrame" && framesToAdjust[i].kind === TextType.AREATEXT) { disableFrameAutoSize(framesToAdjust[i]); }
                    }
                } catch (e) { }
            }
            updatePreview();
        };
        chkLeaderTabs.onClick = function () {
            // リーダー罫は右揃え前提なので、ONにしたら行揃えも右にそろえる
            if (chkLeaderTabs.value) { selectRadio(justificationRadios, radJustifyRight); }
            updatePreview();
        };
        chkSpacing.onClick = function () {
            etSpacing.enabled = chkSpacing.value;
            lblSpacingUnit.enabled = chkSpacing.value;
            if (chkSpacing.value) { etSpacing.text = "1"; }
            onIndentOrSpacingChange();
        };
        chkLinkIndents.onClick = function () {
            if (chkLinkIndents.value) {
                chkLeftIndent.value = true; etLeftIndent.enabled = true;
                chkRightIndent.enabled = false; etRightIndent.enabled = false;
                etRightIndent.text = etLeftIndent.text;
            } else {
                chkRightIndent.enabled = true; etRightIndent.enabled = chkRightIndent.value;
            }
            onIndentOrSpacingChange();
        };
        chkLeftIndent.onClick = function () {
            if (!chkLinkIndents.value) { etLeftIndent.enabled = chkLeftIndent.value; }
            if (!chkLeftIndent.value) { etLeftIndent.text = "0"; }
            onIndentOrSpacingChange();
        };
        chkRightIndent.onClick = function () {
            etRightIndent.enabled = chkRightIndent.value;
            if (!chkRightIndent.value) { etRightIndent.text = "0"; }
            onIndentOrSpacingChange();
        };

        etFontSize.onChange = applyFontSizeFromField;
        etLeadingPercent.onChange = onLeadingPercentChange;
        etLeadingEffective.onChange = onLeadingEffectiveChange;
        etSpacing.onChange = onIndentOrSpacingChange;
        etWidth.onChange = onWidthChange;
        etHeight.onChange = function () {
            var heightValue = validateSizeField(etHeight, lastValidHeight);
            if (heightValue !== null) { lastValidHeight = heightValue; }
            updatePreview();
        };
        etCharsPerLine.onChange = onCharsPerLineChange;
        etLeftIndent.onChange = onIndentOrSpacingChange;
        etRightIndent.onChange = onIndentOrSpacingChange;
        changeValueByArrowKey(etFontSize, false, applyFontSizeFromField);
        changeValueByArrowKey(etLeadingPercent, false, onLeadingPercentChange);
        changeValueByArrowKey(etLeadingEffective, false, onLeadingEffectiveChange);
        changeValueByArrowKey(etSpacing, false, onIndentOrSpacingChange);
        changeValueByArrowKey(etWidth, false, onWidthChange);
        changeValueByArrowKey(etHeight, false, updatePreview);
        changeValueByArrowKey(etCharsPerLine, false, onCharsPerLineChange);
        changeValueByArrowKey(etLeftIndent, false, onIndentOrSpacingChange);
        changeValueByArrowKey(etRightIndent, false, onIndentOrSpacingChange);

        btnRun.onClick = function () {
            revertPreview();
            applyAdjustments(false);
            adjustDialog.close(1);
        };
        btnCancelAdjust.onClick = function () {
            revertPreview();
            adjustDialog.close(0);
        };

        // 初期値読み込み
        if (initialFrame) { loadValuesFromFrame(initialFrame); }

        updateSeparationModeEnabled();
        updateAutoSizeButtonLabel();

        // 調整ダイアログを開いたらプレビューON
        chkPreview.value = true;
        updatePreview();

        adjustDialog.show();
    }

    // ============================================================
    // エントリポイント / Entry point
    //   ダイナミックアクションを読み込み、終了時に必ずアンロードする
    //   Load dynamic actions and always unload them on exit
    // ============================================================
    if (app.documents.length === 0) {
        alert(getLabel("alert.noDocument"));
        return;
    }

    var doc = app.activeDocument;
    var selection = doc.selection;

    if (!selection || selection.length === 0) {
        alert(getLabel("alert.selectText"));
        return;
    }

    loadDynamicActions();
    try {
        // 選択内容の種別を判定 / Detect what kinds of objects are selected
        var hasPointText = false, hasAreaText = false, hasPath = false;
        for (var i = 0; i < selection.length; i++) {
            var item = selection[i];
            if (item.typename === "TextFrame") {
                if (item.kind === TextType.POINTTEXT || item.kind === TextType.PATHTEXT) hasPointText = true;
                if (item.kind === TextType.AREATEXT) hasAreaText = true;
            }
            if (item.typename === "PathItem" || item.typename === "CompoundPathItem") {
                hasPath = true;
            }
        }

        if (hasPointText || hasPath) {
            // ポイント文字 または 図形 → ダイアログA / Point text or shape → Dialog A
            showConvertDialog(doc, selection);
        } else if (hasAreaText) {
            // エリア内文字のみ → ダイアログB / Area text only → Dialog B
            var firstAreaFrame = null;
            for (var i = 0; i < selection.length; i++) {
                if (selection[i].typename === "TextFrame" && selection[i].kind === TextType.AREATEXT) {
                    firstAreaFrame = selection[i];
                    break;
                }
            }
            showAdjustDialog(doc, firstAreaFrame, null, null);
        } else {
            alert(getLabel("alert.selectText"));
        }
    } finally {
        unloadDynamicActions();
    }

})();
