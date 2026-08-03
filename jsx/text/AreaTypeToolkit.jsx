#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ポイント文字・パス上文字・図形からエリア内文字をつくり、そのまま体裁（サイズ・行送り・行揃え・日本語の組版など）を調整するツール。

詳細はREADMEを参照。

*/

/*

### Overview

Builds Area Type from point text, path text, or shapes, then tunes its typesetting (size, leading, justification, Japanese composition, and so on) in place.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AreaTypeToolkit";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-03";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-03";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AreaTypeToolkit.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AreaTypeToolkit.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nfd6cc5e13654"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 「選択した図形にダミーテキスト」で流し込むダミー文字 / Dummy text for "Dummy text in selected shape" */
    var DUMMY_TEXT_JA = "今日は15:00から予定外のミーティングがありました。疲れを癒すため、夕方に近くのカフェWAVEで、お気に入りの抹茶ラテを楽しみました。短い休憩でしたが、心が“ほっと”しました。";
    var DUMMY_TEXT_EN = "An unplanned meeting popped up at 3:00 PM today. To reset, I stopped by a nearby café for a matcha latte. It was a short break, but it helped me breathe and refocus.";

    /* ダミーテキストの優先フォント候補とサイズ / Preferred fonts and size for dummy text */
    var DUMMY_FONT_JA = ["HiraginoSans-W3", "Hiragino Sans W3"];
    var DUMMY_FONT_EN = ["MyriadPro-Regular", "Myriad Pro Regular", "MyriadPro", "Myriad"];
    var DUMMY_FONT_SIZE = 10;

    /* 日本語の組版の初期値（テキスト側で未設定のときに使う）/ Defaults for Japanese composition (used when the text has none set) */
    var DEFAULT_KINSOKU = "Soft_v2";   /* 弱い禁則 v2 / Loose v2 */
    var DEFAULT_MOJIKUMI_INDEX = 6;    /* ベタ組み / Solid */

    /* 種別プリセット（本文／見出し）。ラジオを押すと行送り・行揃え・禁則・文字組みをまとめて設定する
       Role presets (Body / Heading): one click sets leading, justification, kinsoku and mojikumi */
    var ROLE_PRESETS = {
        body: { leadingPercent: 160, justifyId: "lastLineLeft", kinsoku: "Soft_v2", mojikumiIndex: 6, alignId: "top", tabMode: "clear" },
        heading: { leadingPercent: 120, justifyId: "left", kinsoku: "Soft_v2", mojikumiIndex: 5, alignId: "top", tabMode: "clear" },
        menu: { leadingPercent: 150, justifyId: "right", kinsoku: "Soft_v2", mojikumiIndex: 5, alignId: "top", tabMode: "leader" }
    };

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
            convertTitle: { ja: "エリア内文字に変換", en: "Convert to Area Type" },
            separateTitle: { ja: "テキストを分離", en: "Separate text" }
        },
        panel: {
            createMethod: { ja: "作成方法", en: "Creation method" },
            role: { ja: "種別", en: "Role" },
            frameHandling: { ja: "枠の処理", en: "Frame handling" },
            leading: { ja: "行送り", en: "Leading" },
            justification: { ja: "行揃え", en: "Justification" },
            textAlign: { ja: "垂直方向の配置", en: "Vertical alignment" },
            fontSize: { ja: "フォントサイズ", en: "Font size" },
            frameSize: { ja: "フレームサイズ", en: "Frame size" },
            indent: { ja: "インデント", en: "Indent" },
            jpComposition: { ja: "日本語の組版", en: "Japanese composition" },
            offset: { ja: "オフセット", en: "Offset" }
        },
        radio: {
            strokeBlack: { ja: "枠を1pt黒に", en: "Frame: 1pt black" },
            hidePath: { ja: "枠を塗り・線なしに", en: "Frame: unpainted" },
            removePath: { ja: "枠を削除", en: "Delete frame" },
            roleBody: { ja: "本文", en: "Body" },
            roleHeading: { ja: "見出し", en: "Heading" },
            roleMenu: { ja: "メニュー", en: "Menu" },
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
        kinsoku: {
            none: { ja: "なし", en: "None" },
            hard: { ja: "強い禁則", en: "Strict" },
            soft: { ja: "弱い禁則", en: "Loose" },
            softV2: { ja: "弱い禁則 v2", en: "Loose v2" }
        },
        mojikumi: {
            none: { ja: "なし", en: "None" },
            lineEndFullHalf: { ja: "行末約物全角/半角", en: "Line-end punct full/half" },
            punctHalf: { ja: "約物半角", en: "Half-width punctuation" },
            lineEndHalf: { ja: "行末約物半角", en: "Line-end punct half" },
            lineEndFull: { ja: "行末約物全角", en: "Line-end punct full" },
            punctFull: { ja: "約物全角", en: "Full-width punctuation" },
            tight: { ja: "ツメ組み", en: "Tight" },
            solid: { ja: "ベタ組み", en: "Solid" }
        },
        checkbox: {
            linkIndents: { ja: "連動", en: "Link" },
            autoSize: { ja: "自動サイズ調整", en: "Auto-size" }
        },
        button: {
            convert: { ja: "変換", en: "Convert" },
            shrinkToFit: { ja: "文字あふれ解消", en: "Clear overset" },
            fitFontSize: { ja: "枠にフィット", en: "Fit to frame" },
            separateText: { ja: "テキストを分離...", en: "Separate text..." },
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        label: {
            fontSize: { ja: "フォントサイズ", en: "Font size" },
            leadingPercent: { ja: "行送り", en: "Leading" },
            leadingEffective: { ja: "実寸", en: "Actual" },
            width: { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" },
            charsPerLine: { ja: "文字", en: "chars" },
            kinsoku: { ja: "禁則", en: "Kinsoku" },
            mojikumi: { ja: "文字組み", en: "Mojikumi" },
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
                ja: "あふれが消えるまでフォントサイズを縮小します。改行を含むテキストには使えません。サイズが混在していると1つのサイズに揃います。",
                en: "Shrinks the font until the overset clears. Not available for text with line breaks. Mixed sizes are flattened to a single size."
            },
            fitFontSize: {
                ja: "あふれるまで拡大してから縮小し、枠いっぱいに収めます。改行を含むテキストには使えません。",
                en: "Grows until it oversets, then shrinks to fill the frame. Not available for text with line breaks."
            },
            autoSize: {
                ja: "エリア内文字の自動サイズ調整。クリックするとその場で反映されます。ONのまま［文字あふれ解消］［枠にフィット］を押すと、いったんOFFにして実行し、またONに戻します。",
                en: "Area Type auto-sizing. Applied as soon as you click. Clear overset and Fit to frame switch it off for the pass and back on afterwards."
            },
            leading: {
                ja: "自動行送り量（%）。実寸＝フォントサイズ×%。実寸に入力すると%を逆算します。固定行送りのテキストでは空で開きます。",
                en: "Auto-leading amount in %. Actual = font size x %. Enter an actual value to back-calculate the %. Empty for text with a fixed leading."
            },
            textAlign: {
                ja: "フレーム内でのテキストの縦位置。クリックするとその場で反映されます。",
                en: "Vertical placement inside the frame. Applied as soon as you click."
            },
            charsPerLine: {
                ja: "1行の文字数から幅を逆算します。字幅が一定でないため日本語UIのみ。",
                en: "Works the width back out from the characters per line. Japanese UI only, since Roman glyph widths vary."
            },
            roleBody: {
                ja: "本文向けの設定をまとめて適用します（行送り160%・均等配置（最終行左揃え）・弱い禁則 v2・ベタ組み・上揃え）。タブ設定は削除します。",
                en: "Applies the body-text preset (160% leading, justify with last line left, Loose v2 kinsoku, Solid mojikumi, top alignment). Tab stops are cleared."
            },
            roleHeading: {
                ja: "見出し向けの設定をまとめて適用します（行送り120%・左揃え・弱い禁則 v2・ツメ組み・上揃え）。タブ設定は削除します。",
                en: "Applies the heading preset (120% leading, left, Loose v2 kinsoku, Tight mojikumi, top alignment). Tab stops are cleared."
            },
            roleMenu: {
                ja: "メニュー向けの設定をまとめて適用します（行送り150%・右揃え・弱い禁則 v2・ツメ組み・上揃え）。各段落に右揃えタブ（リーダー「…」／400pt）を設定します。行揃えを変えると指定が解除され、タブ設定も削除されます。",
                en: "Applies the menu preset (150% leading, right, Loose v2 kinsoku, Tight mojikumi, top alignment) and sets a right-aligned tab with a leader at 400pt on every paragraph. Changing the justification drops this role and clears the tab stops."
            },
            kinsoku: {
                ja: "段落の禁則処理（なし／強い禁則／弱い禁則など）をまとめて適用します。",
                en: "Applies a kinsoku (line-break) set to the paragraphs (None / Strict / Loose and so on)."
            },
            mojikumi: {
                ja: "文字組みアキ量設定を全段落に適用します。設定が混在しているときは選択が空になり、そのままなら変更しません。",
                en: "Applies a mojikumi spacing set to every paragraph. With mixed settings the menu opens empty and nothing is changed."
            },
            linkIndents: {
                ja: "左インデントの値を右にも適用します。",
                en: "Applies the left indent value to the right as well."
            },
            separateText: {
                ja: "エリア内文字を、囲み罫（長方形）とポイント文字に分解します。別ダイアログで枠の処理を選びます。",
                en: "Breaks Area Type apart into a rectangle plus point text. A separate dialog picks how the frame is handled."
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
            lineBreakNotSupported: {
                ja: "改行を含むテキストには対応していません。改行のないテキストを選択してください。",
                en: "Text containing line breaks is not supported. Select text without line breaks."
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

    /* 禁則・文字組みポップアップの幅（英語は語が長いので広め）/ Width of the kinsoku and mojikumi popups (wider in English) */
    var JP_DROPDOWN_WIDTH = (currentLanguage === "ja") ? 140 : 190;

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

    /* ボタンの高さを指定 px 詰める（レイアウト確定後に呼ぶ）/ Trim a button's height by the given px (call after layout)
       size を変えても location は動かないので、上下から均等に詰めるには位置も半分ずらす
       Changing size keeps the location, so the position shifts by half to trim top and bottom evenly */
    function trimButtonHeight(button, px) {
        try {
            var buttonLocation = [button.location[0], button.location[1]];
            button.size = [button.size.width, button.size.height - px];
            button.location = [buttonLocation[0], buttonLocation[1] + px / 2];
        } catch (e) { }
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

    // =========================================
    // アイコンボタン（行揃え／テキストの配置）/ Icon buttons (justification and text alignment)
    //   UnifiedTypePanel.jsx（原典は Keep_TextPosition.jsx）から移植
    //   Ported from UnifiedTypePanel.jsx (originally Keep_TextPosition.jsx)
    // =========================================

    /* 行揃えの選択肢（id・ラベル・Justification 値・アイコン種別・ショートカット）
       Justification options (id, label, Justification value, icon type, shortcut key) */
    var JUSTIFY_OPTIONS = [
        { id: "left", labelKey: "radio.justifyLeft", value: Justification.LEFT, icon: "left", shortcut: "L" },
        { id: "center", labelKey: "radio.justifyCenter", value: Justification.CENTER, icon: "center", shortcut: "C" },
        { id: "right", labelKey: "radio.justifyRight", value: Justification.RIGHT, icon: "right", shortcut: "R" },
        { id: "lastLineLeft", labelKey: "radio.justifyLastLineLeft", value: Justification.FULLJUSTIFYLASTLINELEFT, icon: "justifyLeft", shortcut: "J" },
        { id: "allLines", labelKey: "radio.justifyAllLines", value: Justification.FULLJUSTIFY, icon: "justifyAll", shortcut: "F" }
    ];

    /* 行揃え id から Justification 値を引く / Resolve a justification id to a Justification value */
    function getJustificationValue(justifyId) {
        for (var i = 0; i < JUSTIFY_OPTIONS.length; i++) {
            if (JUSTIFY_OPTIONS[i].id === justifyId) return JUSTIFY_OPTIONS[i].value;
        }
        return Justification.LEFT;
    }

    /* Justification 値から行揃え id を引く / Resolve a Justification value to a justification id */
    function getJustificationId(justificationValue) {
        for (var i = 0; i < JUSTIFY_OPTIONS.length; i++) {
            if (JUSTIFY_OPTIONS[i].value === justificationValue) return JUSTIFY_OPTIONS[i].id;
        }
        return "left";
    }

    /* テキストの配置の選択肢（id・ラベル・ダイナミックアクションの値・アイコン種別）
       Text-alignment options (id, label, dynamic-action value, icon type) */
    var ALIGN_OPTIONS = [
        { id: "top", labelKey: "radio.alignTop", value: 0, icon: "top" },
        { id: "center", labelKey: "radio.alignCenter", value: 1, icon: "center" },
        { id: "bottom", labelKey: "radio.alignBottom", value: 2, icon: "bottom" },
        { id: "justify", labelKey: "radio.alignJustify", value: 3, icon: "justify" }
    ];

    /* 配置 id からアクションの値を引く / Resolve an alignment id to its action value */
    function getAlignmentValue(alignId) {
        for (var i = 0; i < ALIGN_OPTIONS.length; i++) {
            if (ALIGN_OPTIONS[i].id === alignId) return ALIGN_OPTIONS[i].value;
        }
        return 0;
    }

    /* アクションの値から配置 id を引く / Resolve an action value to an alignment id */
    function getAlignmentId(alignmentValue) {
        for (var i = 0; i < ALIGN_OPTIONS.length; i++) {
            if (ALIGN_OPTIONS[i].value === alignmentValue) return ALIGN_OPTIONS[i].id;
        }
        return "top";
    }

    /* 環境設定のUI明るさが明るい側か（0=最暗〜1=最明。0.5 は暗い側に含める）
       Whether the UI brightness preference is on the light side (0=darkest to 1=lightest; 0.5 counts as dark) */
    function isLightUI() {
        try {
            return app.preferences.getRealPreference("uiBrightness") > 0.5;
        } catch (e) {
            return false;
        }
    }

    /* テーマ＋選択状態に応じたボタン配色 / Button colors per theme and active state */
    function getIconButtonColors(isLight, isActive) {
        if (isLight) {
            return {
                bg: isActive ? [0.40, 0.40, 0.40, 1] : [1, 1, 1, 1],
                border: isActive ? [0.30, 0.30, 0.30, 1] : [0.62, 0.62, 0.62, 1],
                line: isActive ? [1, 1, 1, 1] : [0.25, 0.25, 0.25, 1]
            };
        }
        return {
            bg: isActive ? [0.92, 0.92, 0.92, 1] : [0.30, 0.30, 0.30, 1],
            border: null,
            line: isActive ? [0.16, 0.16, 0.16, 1] : [0.82, 0.82, 0.82, 1]
        };
    }

    /* 行ごとの線幅（アイコン種別ごと）/ Per-line widths for an icon type */
    function getJustifyLineWidths(iconType, longWidth, shortWidth) {
        if (iconType === "justifyAll") return [longWidth, longWidth, longWidth, longWidth];
        if (iconType === "justifyLeft") return [longWidth, longWidth, longWidth, shortWidth];
        return [longWidth, shortWidth, longWidth, shortWidth];
    }

    /* 行の開始 X（アイコン種別ごとに左／中央／右）/ Line start X per icon type (left / center / right) */
    function getJustifyLineX(iconType, buttonWidth, lineWidth) {
        var margin = 5;
        if (iconType === "right") return buttonWidth - margin - lineWidth;
        if (iconType === "center") return Math.round((buttonWidth - lineWidth) / 2);
        return margin;
    }

    /* アイコンの罫線を描く / Draw the icon's lines */
    function drawJustifyIconLines(graphics, iconType, buttonWidth, lineColor) {
        var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, lineColor, 1.2);
        var rowYs = [7, 11, 15, 19];
        var lineWidths = getJustifyLineWidths(iconType, 15, 10);
        for (var i = 0; i < rowYs.length; i++) {
            var lineWidth = lineWidths[i];
            var lineStartX = getJustifyLineX(iconType, buttonWidth, lineWidth);
            graphics.newPath();
            graphics.moveTo(lineStartX, rowYs[i]);
            graphics.lineTo(lineStartX + lineWidth, rowYs[i]);
            graphics.strokePath(pen);
        }
    }

    /* 行の Y 位置（配置種別ごと。上寄せ／中央／下寄せ／均等）
       Row Y positions per alignment icon type (top / center / bottom / evenly spread) */
    function getAlignRowYs(iconType) {
        if (iconType === "center") return [9, 13, 17];
        if (iconType === "bottom") return [11, 15, 19];
        if (iconType === "justify") return [7, 13, 19];
        return [7, 11, 15];
    }

    /* テキストの配置アイコンの罫線を描く（幅は3本とも同じで左右中央）
       Draw the lines for a text-alignment icon (three equal-width lines, horizontally centered) */
    function drawAlignIconLines(graphics, iconType, buttonWidth, lineColor) {
        var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, lineColor, 1.2);
        var lineWidth = 14;
        var lineStartX = Math.round((buttonWidth - lineWidth) / 2);
        var rowYs = getAlignRowYs(iconType);
        for (var i = 0; i < rowYs.length; i++) {
            graphics.newPath();
            graphics.moveTo(lineStartX, rowYs[i]);
            graphics.lineTo(lineStartX + lineWidth, rowYs[i]);
            graphics.strokePath(pen);
        }
    }

    /* アイコンボタンの背景を描く / Draw an icon button's background */
    function drawIconButtonBackground(graphics, buttonWidth, buttonHeight, colors) {
        try {
            graphics.rectPath(0, 0, buttonWidth, buttonHeight);
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, colors.bg));
            if (colors.border) {
                graphics.rectPath(0, 0, buttonWidth, buttonHeight);
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, colors.border, 1));
            }
        } catch (e) {
            try { graphics.drawOSControl(); } catch (e0) { }
        }
    }

    /* 行揃えボタンを描く（テーマ・選択状態に対応）/ Draw a justification button (theme and active aware) */
    function drawJustifyIcon(button, isActive, isLight) {
        var colors = getIconButtonColors(isLight, isActive);
        drawIconButtonBackground(button.graphics, button.size[0], button.size[1], colors);
        drawJustifyIconLines(button.graphics, button.iconType, button.size[0], colors.line);
    }

    /* テキストの配置ボタンを描く（テーマ・選択状態に対応）/ Draw a text-alignment button (theme and active aware) */
    function drawAlignIcon(button, isActive, isLight) {
        var colors = getIconButtonColors(isLight, isActive);
        drawIconButtonBackground(button.graphics, button.size[0], button.size[1], colors);
        drawAlignIconLines(button.graphics, button.iconType, button.size[0], colors.line);
    }

    // =========================================
    // 日本語の組版（禁則・文字組みアキ量設定）/ Japanese composition (kinsoku and mojikumi)
    //   UnifiedTypePanel.jsx から移植 / Ported from UnifiedTypePanel.jsx
    // =========================================

    /* 禁則の選択肢（id は paragraphAttributes.kinsoku に渡す値）
       Kinsoku choices (id is the value passed to paragraphAttributes.kinsoku) */
    var KINSOKU_CHOICES = [
        { id: "None", labelKey: "kinsoku.none" },
        { id: "Hard", labelKey: "kinsoku.hard" },
        { id: "Soft", labelKey: "kinsoku.soft" },
        { id: "Soft_v2", labelKey: "kinsoku.softV2" }
    ];

    /* 文字組みアキ量設定の選択肢（index は mojikumiSet の添字。-1 は「なし」）
       Mojikumi choices (index is the mojikumiSet index; -1 means "None") */
    var MOJIKUMI_CHOICES = [
        { index: -1, labelKey: "mojikumi.none" },
        { index: 0, labelKey: "mojikumi.lineEndFullHalf" },
        { index: 1, labelKey: "mojikumi.punctHalf" },
        { index: 2, labelKey: "mojikumi.lineEndHalf" },
        { index: 3, labelKey: "mojikumi.lineEndFull" },
        { index: 4, labelKey: "mojikumi.punctFull" },
        { index: 5, labelKey: "mojikumi.tight" },
        { index: 6, labelKey: "mojikumi.solid" }
    ];

    /* 選択肢テーブルからドロップダウン用のラベル配列を作る / Build the dropdown item list from a choice table */
    function buildChoiceLabels(choices) {
        var labels = [];
        for (var i = 0; i < choices.length; i++) { labels.push(getLabel(choices[i].labelKey)); }
        return labels;
    }

    /* 選択肢テーブルの値に一致する項目をドロップダウンで選ぶ（無ければ未選択のままにする）
       Select the dropdown item matching a value (left unselected when there is no match) */
    function selectChoiceByValue(dropdown, choices, keyName, value) {
        for (var i = 0; i < choices.length; i++) {
            if (choices[i][keyName] === value) { dropdown.selection = i; return; }
        }
        dropdown.selection = null;
    }

    /* 「なし」を表す文字組み設定名の候補（UI言語で変わるため）/ Names that mean "none" for mojikumi (they vary by UI language) */
    var MOJIKUMI_NONE_NAMES = ["なし", "None"];

    /* 文字組み設定名が「なし」を指すか / Whether a mojikumi name means "none" */
    function isMojikumiNoneName(name) {
        for (var i = 0; i < MOJIKUMI_NONE_NAMES.length; i++) {
            if (name === MOJIKUMI_NONE_NAMES[i]) return true;
        }
        return false;
    }

    /* 段落の文字組みを「なし」にする（設定名の候補を順に試す）/ Clear a paragraph's mojikumi, trying each candidate name */
    function clearParagraphMojikumi(paragraph) {
        for (var i = 0; i < MOJIKUMI_NONE_NAMES.length; i++) {
            try {
                paragraph.paragraphAttributes.mojikumi = MOJIKUMI_NONE_NAMES[i];
                return;
            } catch (e) { }
        }
    }

    /* 文字組み設定名から mojikumiSet の添字を引く（見つからなければ -2）
       Look up a mojikumiSet index by name (-2 when not found) */
    function matchMojikumiName(name) {
        try {
            var mojikumiSets = app.activeDocument.mojikumiSet;
            for (var i = 0; i < mojikumiSets.length; i++) {
                if (mojikumiSets[i].name === name) return i;
            }
        } catch (e) { }
        return -2;
    }

    /* 先頭段落の文字組み設定を添字で返す（-1＝なし、-2＝不明・混在）
       Return the first paragraph's mojikumi as an index (-1 = none, -2 = unknown or mixed) */
    function getMojikumiIndex(textFrame) {
        var mojikumiAttr;
        try {
            if (textFrame.paragraphs.length === 0) return -2;
            mojikumiAttr = textFrame.paragraphs[0].paragraphAttributes.mojikumi;
        } catch (e) { return -2; }
        if (mojikumiAttr === undefined || mojikumiAttr === null) return -1;
        if (typeof mojikumiAttr === "string") {
            return (mojikumiAttr === "" || isMojikumiNoneName(mojikumiAttr)) ? -1 : matchMojikumiName(mojikumiAttr);
        }
        var mojikumiName;
        try { mojikumiName = (mojikumiAttr.name === undefined || mojikumiAttr.name === null) ? "" : mojikumiAttr.name; } catch (e0) { return -2; }
        if (mojikumiName === "") return -2;
        if (isMojikumiNoneName(mojikumiName)) return -1;
        return matchMojikumiName(mojikumiName);
    }

    /* 先頭段落の禁則処理を id で返す / Return the first paragraph's kinsoku set as an id */
    function getKinsokuId(textFrame) {
        try {
            if (textFrame.paragraphs.length === 0) return "None";
            var kinsokuValue = textFrame.paragraphs[0].paragraphAttributes.kinsoku;
            if (kinsokuValue === undefined || kinsokuValue === null || kinsokuValue === "") return "None";
            return String(kinsokuValue);
        } catch (e) {
            return "None";
        }
    }

    /* 全段落に文字組みアキ量設定を適用する（-2＝不明のときは触らない）
       Apply a mojikumi set to every paragraph (-2 means unknown, so nothing is touched) */
    function applyMojikumiToTextFrame(textFrame, mojikumiIndex) {
        if (mojikumiIndex === undefined || mojikumiIndex === null || mojikumiIndex === -2) return;
        if (mojikumiIndex === -1) {
            for (var i = 0; i < textFrame.paragraphs.length; i++) { clearParagraphMojikumi(textFrame.paragraphs[i]); }
            return;
        }
        var mojikumiValue;
        try { mojikumiValue = app.activeDocument.mojikumiSet[mojikumiIndex]; } catch (e) { return; }
        for (var j = 0; j < textFrame.paragraphs.length; j++) {
            try { textFrame.paragraphs[j].paragraphAttributes.mojikumi = mojikumiValue; } catch (e0) { }
        }
    }

    /* 禁則処理を適用する（「なし」も含めて設定できる）
       Apply a kinsoku set ("None" included) */
    function applyKinsokuToTextFrame(textFrame, kinsokuId) {
        if (!kinsokuId) return;
        try { textFrame.textRange.paragraphAttributes.kinsoku = kinsokuId; } catch (e) { }
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
        restoreCharacterAttributes(pointText, charAttrSnapshots, true);

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
                leading: charAttr.leading,
                baselineShift: charAttr.baselineShift,
                horizontalScale: charAttr.horizontalScale,
                verticalScale: charAttr.verticalScale
            });
        }
        return snapshots;
    }

    /* 控えた文字属性を書き戻す / Write snapshotted attributes back
       dropTransforms が true のときは、パス上文字で付いたベースライン移動・拡大縮小を捨てる
       When dropTransforms is true, the baseline shift and scaling baked in by path text are dropped */
    function restoreCharacterAttributes(textFrame, snapshots, dropTransforms) {
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
                destAttrs.baselineShift = dropTransforms ? 0 : srcAttrs.baselineShift;
                destAttrs.horizontalScale = dropTransforms ? 100 : srcAttrs.horizontalScale;
                destAttrs.verticalScale = dropTransforms ? 100 : srcAttrs.verticalScale;
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

    /* エリア内文字の枠の高さを返す（取れなければ 0）/ Height of the Area Type frame (0 when unavailable) */
    function getFramePathHeight(textFrame) {
        try { return textFrame.textPath.height || 0; } catch (e) { return 0; }
    }

    /* 改行を含むか / Whether the text contains a hard line break */
    function hasLineBreak(textFrame) {
        try { return /[\r\n\u2028\u2029]/.test(textFrame.contents); } catch (e) { return false; }
    }

    /* 代表となるフォントサイズを返す（混在しているときは先頭文字の値）/ Representative font size (falls back to the first character when sizes are mixed) */
    function getRepresentativeFontSize(textFrame) {
        try {
            var size = textFrame.textRange.characterAttributes.size;
            if (size > 0) return size;
        } catch (e) { }
        try { return textFrame.characters[0].characterAttributes.size || 0; } catch (e0) { }
        return 0;
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

    /* アクションセット名。ユーザーのアクションを消さないよう、スクリプト名を冠して衝突を避ける
       Action set names, prefixed with the script name so a user's own sets are never unloaded */
    var ACTION_SET_AUTO_SIZE = "AreaTypeToolkit_AutoSize";
    var ACTION_SET_ALIGNMENT = "AreaTypeToolkit_Alignment";

    /* 使用するアクションセットを読み込む（スクリプト開始時に1回）/ Load the action sets used (once at startup) */
    function loadDynamicActions() {
        // AutoSizeOn(1) / AutoSizeOff(2)
        loadActionSet(ACTION_SET_AUTO_SIZE, buildActionSetAia(
            ACTION_SET_AUTO_SIZE,
            "adobe_SLOAreaTextDialog",
            "33 e382a8e383aae382a2e58685e69687e5ad97e382aae38397e382b7e383a7e383b3",
            1952539754,
            [
                { name: "AutoSizeOn", value: 1 },
                { name: "AutoSizeOff", value: 2 }
            ]
        ));
        // AlignTop(0) / AlignCenter(1) / AlignBottom(2) / AlignJustify(3)
        loadActionSet(ACTION_SET_ALIGNMENT, buildActionSetAia(
            ACTION_SET_ALIGNMENT,
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
        try { app.unloadAction(ACTION_SET_AUTO_SIZE, ""); } catch (e) { }
        try { app.unloadAction(ACTION_SET_ALIGNMENT, ""); } catch (e) { }
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
        try { app.doScript(actionName, ACTION_SET_AUTO_SIZE, false); } catch (e) { }
    }

    /* テキストの配置アクションを実行（0=上, 1=中央, 2=下, 3=均等）/ Run the frame-alignment action (0=top, 1=center, 2=bottom, 3=justify) */
    function runFrameAlignmentAction(alignmentValue) {
        if (alignmentValue !== 0 && alignmentValue !== 1 && alignmentValue !== 2 && alignmentValue !== 3) return;
        var actionName = "AlignTop";
        if (alignmentValue === 1) actionName = "AlignCenter";
        else if (alignmentValue === 2) actionName = "AlignBottom";
        else if (alignmentValue === 3) actionName = "AlignJustify";
        try { app.doScript(actionName, ACTION_SET_ALIGNMENT, false); } catch (e) { }
    }

    /* 対象を選択してテキストの配置アクションを実行する（プレビュー中は実行しない）/ Select the frame and run the placement action (skipped during preview) */
    function applyAreaTextFrameAlignment(textFrame, alignmentValue, forPreview) {
        // app.doScript はプレビュー中にダイアログから呼ぶと不安定なためスキップ
        if (forPreview) return;
        // 配置未指定（ユーザーが触っていない）なら、いまの配置をそのままにする
        // No alignment requested (the user never touched it), so the current placement is left as is
        if (alignmentValue === null || alignmentValue === undefined) return;
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
        var upperSize = getRepresentativeFontSize(textFrame);
        if (upperSize <= 0) return;
        var leadingRatioInfo = getLeadingRatioInfo(textFrame);
        var lowerSize = 0.1;

        // 最小サイズでも収まらないなら、極小のまま残さず元のサイズに戻す
        textFrame.textRange.characterAttributes.size = lowerSize;
        applyProportionalLeading(textFrame, lowerSize, leadingRatioInfo);
        if (isFrameOverset(textFrame)) {
            textFrame.textRange.characterAttributes.size = upperSize;
            applyProportionalLeading(textFrame, upperSize, leadingRatioInfo);
            return;
        }

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
        var originalSize = getRepresentativeFontSize(textFrame);
        if (originalSize <= 0) return;
        var leadingRatioInfo = getLeadingRatioInfo(textFrame);

        /* 元のフォントサイズに戻す / Put the original font size back */
        function restoreOriginalSize() {
            try {
                textFrame.textRange.characterAttributes.size = originalSize;
                applyProportionalLeading(textFrame, originalSize, leadingRatioInfo);
            } catch (e) { }
        }

        // 自動サイズ調整がONだと枠が文字に追従してあふれないため、拡大が止まらない。
        // 枠の高さが変わったらそれと判断して中止する。
        // With auto-size on the frame follows the text and never oversets, so the growth never stops.
        // A change in the frame height means that is happening, so bail out.
        var initialFrameHeight = getFramePathHeight(textFrame);

        // Step 1: オーバーセットが出るまで 2 倍ずつ拡大
        if (!isFrameOverset(textFrame)) {
            var trialSize = originalSize, guardCount = 0;
            while (!isFrameOverset(textFrame) && guardCount < 25) {
                guardCount++;
                trialSize = trialSize * 2;
                if (trialSize > 100000) break;
                try { textFrame.textRange.characterAttributes.size = trialSize; applyProportionalLeading(textFrame, trialSize, leadingRatioInfo); } catch (e) { break; }
                if (Math.abs(getFramePathHeight(textFrame) - initialFrameHeight) > 0.01) {
                    restoreOriginalSize();
                    return;
                }
            }
        }

        // オーバーセットが出なければ元に戻して終了
        if (!isFrameOverset(textFrame)) {
            restoreOriginalSize();
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

    /* 段落のタブ設定を削除する / Clear the tab stops on every paragraph */
    function clearTabStops(textFrame) {
        try {
            for (var i = 0; i < textFrame.paragraphs.length; i++) {
                textFrame.paragraphs[i].paragraphAttributes.tabStops = [];
            }
        } catch (e) { }
    }

    // ============================================================
    // 変換ダイアログ：ポイント文字 → エリア内文字 変換
    // ============================================================
    function showConvertDialog(doc, selection) {
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
            // テキスト＋図形 → シンプル・ボタン風をディムし、「選択した図形に流し込む」を自動選択
            radStyleSimple.enabled = false;
            radStyleButton.enabled = false;
            radUseShape.value = true; // default, dummy is also available
        } else if (hasPointText && !hasPathItem) {
            // ポイント文字のみ → 図形を使う方法をディム
            radUseShape.enabled = false;
            radUseShapeDummy.enabled = false;
        } else if (!hasPointText && hasPathItem) {
            // 図形のみ → 「選択した図形にダミーテキスト」を自動選択し、残りをディム
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
            // パス上文字の置き換えは、変換を確定してから行う（キャンセル時に元のパス上文字を壊さないため）
            // Path text is swapped out only once the conversion is confirmed, so Cancel leaves it intact
            currentSelection = preprocessSelectionForConvertDialog(doc, currentSelection);
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

    /* 中身が無くなった複合パスの殻を片付ける / Drop a compound path shell left empty */
    function removeEmptyCompoundPath(item) {
        try {
            if (item && item.typename === "CompoundPathItem" && item.pathItems.length === 0) { item.remove(); }
        } catch (e) { }
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
    function findShapeForText(sourceText, shapeItems, usedShapes) {
        var textCenter = getBoundsCenter(sourceText);
        var nearestShape = null, nearestDistance = -1;

        for (var i = 0; i < shapeItems.length; i++) {
            var shapeItem = shapeItems[i];

            var isUsed = false;
            for (var j = 0; j < usedShapes.length; j++) {
                if (usedShapes[j] === shapeItem) { isUsed = true; break; }
            }
            if (isUsed) continue;

            var shapeBounds = shapeItem.geometricBounds;
            if (textCenter[0] >= shapeBounds[0] && textCenter[0] <= shapeBounds[2] &&
                textCenter[1] <= shapeBounds[1] && textCenter[1] >= shapeBounds[3]) {
                return shapeItem;
            }

            var shapeCenter = getBoundsCenter(shapeItem);
            var dx = shapeCenter[0] - textCenter[0], dy = shapeCenter[1] - textCenter[1];
            var distance = dx * dx + dy * dy;
            if (nearestDistance < 0 || distance < nearestDistance) {
                nearestDistance = distance;
                nearestShape = shapeItem;
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
            var sourceShape = selection[i];
            var closedPath = getClosedPathItem(sourceShape);
            if (!closedPath) continue;
            try {
                var dummyFrame = createAreaTextFromPath(doc, closedPath);
                dummyFrame.contents = dummyText;
                applyFontAndSize(dummyFrame, { font: dummyFont, size: DUMMY_FONT_SIZE });
                createdFrames.push(dummyFrame);
                // 複合パスは先頭のパスだけを枠にするので、空になった殻を残さない
                removeEmptyCompoundPath(sourceShape);
            } catch (e) { }
        }
        return createdFrames;
    }

    /* 選択テキストを、対応する閉じたパスのエリア内文字に流し込む / Pour each selected text into its matching closed path */
    function pourTextsIntoShapes(doc, selection) {
        var createdFrames = [];
        var sourceTexts = [], shapeItems = [];
        for (var i = 0; i < selection.length; i++) {
            // 複合パスも図形として受け付ける（変換ダイアログのラジオ判定と合わせる）
            // Compound paths count as shapes too, matching how the convert dialog enables the radios
            if (selection[i].typename === "TextFrame") { sourceTexts.push(selection[i]); }
            else if (getClosedPathItem(selection[i])) { shapeItems.push(selection[i]); }
        }

        var usedShapes = [];
        for (var textIndex = 0; textIndex < sourceTexts.length; textIndex++) {
            var sourceText = sourceTexts[textIndex];
            var targetShape = findShapeForText(sourceText, shapeItems, usedShapes);
            if (!targetShape) continue;
            usedShapes.push(targetShape);
            try {
                var contents = sourceText.contents;
                var fontAndSize = getFontAndSize(sourceText);
                // 複合パスごと複製してから、その中の閉じたパスを枠にする
                // Duplicate the whole item, then use the closed path inside the copy as the frame
                var duplicatedShape = targetShape.duplicate();
                var areaFrame = createAreaTextFromPath(doc, getClosedPathItem(duplicatedShape));
                areaFrame.contents = contents;
                applyFontAndSize(areaFrame, fontAndSize);
                removeEmptyCompoundPath(duplicatedShape);
                sourceText.remove();
                targetShape.remove();
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

    /* 段落の行揃えとインデントを適用する / Apply justification and indents to every paragraph */
    function applyParagraphSettings(textFrame, settings) {
        try {
            var paraAttrs = textFrame.textRange.paragraphAttributes;
            paraAttrs.justification = settings.justification;
            paraAttrs.leftIndent = settings.leftIndentPt;
            paraAttrs.rightIndent = settings.rightIndentPt;
        } catch (e) { }
        applyAutoLeading(textFrame, settings.leadingPercent);
        applyKinsokuToTextFrame(textFrame, settings.kinsoku);
        applyMojikumiToTextFrame(textFrame, settings.mojikumiIndex);
        if (settings.tabMode === "leader") { applyLeaderTabStops(textFrame); }
        else if (settings.tabMode === "clear") { clearTabStops(textFrame); }
    }

    // ============================================================
    // 分離ダイアログ：エリア内文字 → 囲み罫＋ポイント文字
    // Separate dialog: Area Type → rectangle plus point text
    // ============================================================

    /* エリア内文字を、囲み罫とポイント文字に分解し、できたオブジェクトを返す
       Replace Area Type with a rectangle plus point text, returning what was created */
    function separateAreaTextFrame(doc, areaTextFrame, settings) {
        var bounds = areaTextFrame.geometricBounds;
        var spacingPt = settings.spacingPt;
        var left = bounds[0] - spacingPt, top = bounds[1] + spacingPt;
        var right = bounds[2] + spacingPt, bottom = bounds[3] - spacingPt;

        // 削除する前に、文字ごとの書式と1行目のベースライン位置に使うサイズを控える
        // Snapshot the per-character formatting and the first-baseline size before the frame goes away
        var charAttrSnapshots = snapshotCharacterAttributes(areaTextFrame);
        var firstLineOffset = getRepresentativeFontSize(areaTextFrame);

        var frameRect = doc.pathItems.rectangle(top, left, right - left, top - bottom);
        frameRect.filled = false;
        frameRect.stroked = true;

        var pointTextFrame = doc.textFrames.add();
        pointTextFrame.contents = areaTextFrame.contents;
        pointTextFrame.position = [left, top - firstLineOffset];

        // 既定の線をいったん外し、このあと文字単位で復元する / Clear the default stroke, then restore per character
        try {
            pointTextFrame.textRange.characterAttributes.strokeColor = new NoColor();
            pointTextFrame.textRange.characterAttributes.strokeWeight = 0;
        } catch (e0) { }
        restoreCharacterAttributes(pointTextFrame, charAttrSnapshots, false);

        areaTextFrame.remove();

        var createdItems = [pointTextFrame];

        if (settings.strokeBlack) {
            var blackColor = new CMYKColor();
            blackColor.cyan = 0; blackColor.magenta = 0; blackColor.yellow = 0; blackColor.black = 100;
            frameRect.strokeColor = blackColor;
            frameRect.strokeWidth = 1;
            createdItems.push(frameRect);
        } else if (settings.hidePath) {
            frameRect.stroked = false;
            createdItems.push(frameRect);
        } else if (settings.removePath) {
            frameRect.remove();
        } else {
            createdItems.push(frameRect);
        }

        applyParagraphSettings(pointTextFrame, settings);
        return createdItems;
    }

    /* 「テキストを分離」ダイアログ。分離を実行したら true を返す
       The "Separate text" dialog; returns true once the separation has run
       settings は調整ダイアログから受け取り、枠の処理だけを足して使う
       settings comes from the adjust dialog; only the frame handling is added here */
    function showSeparateTextDialog(doc, areaTextFrames, settings) {
        var separateDialog = new Window("dialog", getLabel("dialog.separateTitle"));
        separateDialog.alignChildren = "fill";
        separateDialog.margins = 20;

        var framePanel = separateDialog.add("panel", undefined, getLabel("panel.frameHandling"));
        applyPanelLayout(framePanel);
        var radStrokeBlack = framePanel.add("radiobutton", undefined, getLabel("radio.strokeBlack"));
        var radHidePath = framePanel.add("radiobutton", undefined, getLabel("radio.hidePath"));
        var radRemovePath = framePanel.add("radiobutton", undefined, getLabel("radio.removePath"));
        radStrokeBlack.value = true;
        bindExclusiveRadios([radStrokeBlack, radHidePath, radRemovePath]);

        // ボタンエリア：右から［OK］［キャンセル］
        var buttonRow = separateDialog.add("group");
        buttonRow.orientation = "row";
        buttonRow.alignment = ["right", "center"];
        var btnCancelSeparate = buttonRow.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnRunSeparate = buttonRow.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        var didSeparate = false;

        btnRunSeparate.onClick = function () {
            settings.strokeBlack = radStrokeBlack.value;
            settings.hidePath = radHidePath.value;
            settings.removePath = radRemovePath.value;

            // 分離するとオブジェクトが置き換わるので後ろから処理する
            var createdItems = [];
            for (var i = areaTextFrames.length - 1; i >= 0; i--) {
                var areaTextFrame = areaTextFrames[i];
                if (!areaTextFrame || areaTextFrame.typename !== "TextFrame" || areaTextFrame.kind !== TextType.AREATEXT) continue;
                // 1フレームで失敗しても残りを処理できるようにする
                try {
                    createdItems = createdItems.concat(separateAreaTextFrame(doc, areaTextFrame, settings));
                    didSeparate = true;
                } catch (e) { }
            }

            // 削除済みの参照が選択に残らないよう、できたオブジェクトを選び直す
            // Re-select what was created, so the deleted frames do not linger in the selection
            try { doc.selection = createdItems.length ? createdItems : null; } catch (e2) { }
            app.redraw();
            separateDialog.close(1);
        };
        btnCancelSeparate.onClick = function () { separateDialog.close(0); };

        separateDialog.show();
        return didSeparate;
    }

    // ============================================================
    // 調整ダイアログ：エリア内文字 調整
    // ============================================================
    function showAdjustDialog(doc, initialFrame, targetFrames, initialAlignmentValue) {
        var rulerInfo = getRulerUnitInfo(doc);

        // 変換したてのフレームかどうか。既存のエリア内文字では、DOM から読み取れない項目
        // （テキストの配置・自動サイズ調整）や未設定の項目を、触っていないのに書き換えない
        // Whether these frames were just converted. For existing Area Type, settings that cannot be
        // read back from the DOM (placement, auto-size) are left alone unless the user touches them
        var isNewlyConverted = !!(targetFrames && targetFrames.length);

        /* ユーザーが操作した項目 / Which settings the user has touched */
        var userTouched = {
            alignment: isNewlyConverted,
            height: isNewlyConverted,
            kinsoku: isNewlyConverted,
            mojikumi: isNewlyConverted
        };

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

        // 右カラム：種別（本文／見出し／メニュー）。押すと関連する設定をまとめて適用する
        // Right column: text role (Body / Heading / Menu); one click applies the whole preset
        var rolePanel = rightColumn.add("panel", undefined, getLabel("panel.role"));
        applyPanelLayout(rolePanel, 4);
        rolePanel.orientation = "row";
        rolePanel.alignChildren = ["left", "center"];
        var radRoleBody = rolePanel.add("radiobutton", undefined, getLabel("radio.roleBody"));
        var radRoleHeading = rolePanel.add("radiobutton", undefined, getLabel("radio.roleHeading"));
        var radRoleMenu = rolePanel.add("radiobutton", undefined, getLabel("radio.roleMenu"));
        radRoleBody.helpTip = getLabel("tip.roleBody");
        radRoleHeading.helpTip = getLabel("tip.roleHeading");
        radRoleMenu.helpTip = getLabel("tip.roleMenu");

        var roleRadios = [radRoleBody, radRoleHeading, radRoleMenu];

        // 選択中の種別と、それが決めたタブ設定。tabMode が "none" のあいだはタブに触らない
        var roleState = { activeId: "", tabMode: "none" };

        // 右カラム：行送り（自動行送り量％と、その実寸）/ Right column: leading (auto-leading % and its effective size)
        var leadingPanel = rightColumn.add("panel", undefined, getLabel("panel.leading"));
        applyPanelLayout(leadingPanel);
        var leadingEffectiveRow = leadingPanel.add("group");
        var lblLeadingEffective = leadingEffectiveRow.add("statictext", undefined, labelWithColon("label.leadingEffective"));
        lblLeadingEffective.preferredSize.width = 48;
        var etLeadingEffective = leadingEffectiveRow.add("edittext", undefined, "");
        etLeadingEffective.characters = 4;
        leadingEffectiveRow.add("statictext", undefined, "pt");
        var leadingPercentRow = leadingPanel.add("group");
        var lblLeadingPercent = leadingPercentRow.add("statictext", undefined, labelWithColon("label.leadingPercent"));
        lblLeadingPercent.preferredSize.width = 48;
        var etLeadingPercent = leadingPercentRow.add("edittext", undefined, "");
        etLeadingPercent.characters = 4;
        leadingPercentRow.add("statictext", undefined, "%");
        leadingPanel.helpTip = getLabel("tip.leading");
        etLeadingPercent.helpTip = leadingPanel.helpTip;
        etLeadingEffective.helpTip = leadingPanel.helpTip;

        // 右カラム：行揃え（アイコンボタン。ラベルはツールチップで見せる）
        // Right column: justification (icon buttons; the labels live in the tooltips)
        var justificationPanel = rightColumn.add("panel", undefined, getLabel("panel.justification"));
        applyPanelLayout(justificationPanel, 4);
        justificationPanel.orientation = "row";
        justificationPanel.alignChildren = ["center", "center"];

        // 選択中の id と UI 明暗を共有する（onDraw のクロージャから参照）
        var isLightTheme = isLightUI();
        var justifyState = { activeId: "left", isLight: isLightTheme };
        var justifyButtons = [];
        for (var justifyIndex = 0; justifyIndex < JUSTIFY_OPTIONS.length; justifyIndex++) {
            var justifyOption = JUSTIFY_OPTIONS[justifyIndex];
            var justifyButton = justificationPanel.add("button", undefined, "");
            justifyButton.helpTip = getLabel(justifyOption.labelKey) + " (" + justifyOption.shortcut + ")";
            justifyButton.preferredSize = [26, 26];
            justifyButton.minimumSize = [26, 26];
            justifyButton.maximumSize = [26, 26];
            justifyButton.justifyId = justifyOption.id;
            justifyButton.iconType = justifyOption.icon;
            justifyButton.onDraw = function () {
                drawJustifyIcon(this, this.justifyId === justifyState.activeId, justifyState.isLight);
            };
            justifyButtons.push(justifyButton);
        }

        // 右カラム：インデント（行揃えの下）/ Right column: indent (below the justification panel)
        var indentPanel = rightColumn.add("panel", undefined, getLabel("panel.indent"));
        applyPanelLayout(indentPanel);
        indentPanel.orientation = "row";
        indentPanel.alignChildren = ["left", "top"];
        var indentFieldsColumn = indentPanel.add("group");
        indentFieldsColumn.orientation = "column";
        indentFieldsColumn.alignChildren = "left";
        var leftIndentRow = indentFieldsColumn.add("group");
        leftIndentRow.add("statictext", undefined, labelWithColon("label.indentLeft"));
        var etLeftIndent = leftIndentRow.add("edittext", undefined, "0");
        etLeftIndent.characters = 4;
        leftIndentRow.add("statictext", undefined, rulerInfo.label);
        var rightIndentRow = indentFieldsColumn.add("group");
        rightIndentRow.add("statictext", undefined, labelWithColon("label.indentRight"));
        var etRightIndent = rightIndentRow.add("edittext", undefined, "0");
        etRightIndent.characters = 4;
        rightIndentRow.add("statictext", undefined, rulerInfo.label);
        var linkColumn = indentPanel.add("group");
        linkColumn.orientation = "column";
        linkColumn.alignChildren = "left";
        linkColumn.alignment = ["left", "center"];
        var chkLinkIndents = linkColumn.add("checkbox", undefined, getLabel("checkbox.linkIndents"));
        chkLinkIndents.helpTip = getLabel("tip.linkIndents");
        chkLinkIndents.value = true;

        // 右カラム：日本語の組版（禁則・文字組みアキ量設定）/ Right column: Japanese composition (kinsoku and mojikumi)
        var jpCompositionPanel = rightColumn.add("panel", undefined, getLabel("panel.jpComposition"));
        applyPanelLayout(jpCompositionPanel, 4);
        var lblKinsoku = jpCompositionPanel.add("statictext", undefined, labelWithColon("label.kinsoku"));
        var kinsokuDropdown = jpCompositionPanel.add("dropdownlist", undefined, buildChoiceLabels(KINSOKU_CHOICES));
        selectChoiceByValue(kinsokuDropdown, KINSOKU_CHOICES, "id", DEFAULT_KINSOKU);
        // fill を打ち消して幅を指定する（既定ではパネル幅いっぱいに広がる）
        kinsokuDropdown.alignment = "left";
        kinsokuDropdown.preferredSize.width = JP_DROPDOWN_WIDTH;
        lblKinsoku.helpTip = getLabel("tip.kinsoku");
        kinsokuDropdown.helpTip = lblKinsoku.helpTip;
        // 禁則との間を少し空ける / A little breathing room after the kinsoku row
        var mojikumiGap = jpCompositionPanel.add("group");
        mojikumiGap.preferredSize.height = 5;
        var lblMojikumi = jpCompositionPanel.add("statictext", undefined, labelWithColon("label.mojikumi"));
        var mojikumiDropdown = jpCompositionPanel.add("dropdownlist", undefined, buildChoiceLabels(MOJIKUMI_CHOICES));
        selectChoiceByValue(mojikumiDropdown, MOJIKUMI_CHOICES, "index", DEFAULT_MOJIKUMI_INDEX);
        mojikumiDropdown.alignment = "left";
        mojikumiDropdown.preferredSize.width = JP_DROPDOWN_WIDTH;
        lblMojikumi.helpTip = getLabel("tip.mojikumi");
        mojikumiDropdown.helpTip = lblMojikumi.helpTip;

        // 左カラム：フォントサイズ / Left column: font size
        var fontSizePanel = leftColumn.add("panel", undefined, getLabel("panel.fontSize"));
        applyPanelLayout(fontSizePanel);
        // パネル名が「フォントサイズ」なので、行のラベルは省く / The panel title already says it, so the row label is dropped
        var fontSizeRow = fontSizePanel.add("group");
        fontSizeRow.alignment = "left";
        var etFontSize = fontSizeRow.add("edittext", undefined, "");
        etFontSize.characters = 4;
        fontSizeRow.add("statictext", undefined, "pt");
        // フォントサイズ欄との間を少し空ける / A little breathing room after the font-size field
        var fontSizeButtonGap = fontSizePanel.add("group");
        fontSizeButtonGap.preferredSize.height = 5;

        // 縦並び。ボタンはラベル幅のまま（パネル幅いっぱいに伸ばさない）
        // Stacked vertically, each button keeping its label width (not stretched to the panel)
        var fontSizeButtonRow = fontSizePanel.add("group");
        fontSizeButtonRow.orientation = "column";
        fontSizeButtonRow.alignChildren = ["left", "top"];
        var btnShrinkToFit = fontSizeButtonRow.add("button", undefined, getLabel("button.shrinkToFit"));
        var btnFitFontSize = fontSizeButtonRow.add("button", undefined, getLabel("button.fitFontSize"));
        btnShrinkToFit.helpTip = getLabel("tip.shrinkToFit");
        btnFitFontSize.helpTip = getLabel("tip.fitFontSize");

        // 左カラム：フレームサイズ / Left column: frame size
        var frameSizePanel = leftColumn.add("panel", undefined, getLabel("panel.frameSize"));
        applyPanelLayout(frameSizePanel);
        var widthRow = frameSizePanel.add("group");
        var lblWidth = widthRow.add("statictext", undefined, labelWithColon("label.width"));
        lblWidth.preferredSize.width = 28;
        var etWidth = widthRow.add("edittext", undefined, "");
        etWidth.characters = 5;
        widthRow.add("statictext", undefined, rulerInfo.label);
        // 幅の下に字詰め欄。空きラベルで幅の入力欄と左端をそろえる
        // Chars per line goes under the width, lined up with the width field via an empty label
        var charsPerLineRow = frameSizePanel.add("group");
        var lblCharsPerLineSpacer = charsPerLineRow.add("statictext", undefined, "");
        lblCharsPerLineSpacer.preferredSize.width = 28;
        var etCharsPerLine = charsPerLineRow.add("edittext", undefined, "");
        etCharsPerLine.characters = 4;
        var lblCharsPerLine = charsPerLineRow.add("statictext", undefined, getLabel("label.charsPerLine"));
        lblCharsPerLine.helpTip = getLabel("tip.charsPerLine");
        etCharsPerLine.helpTip = lblCharsPerLine.helpTip;
        var heightRow = frameSizePanel.add("group");
        var lblHeight = heightRow.add("statictext", undefined, labelWithColon("label.height"));
        lblHeight.preferredSize.width = 28;
        var etHeight = heightRow.add("edittext", undefined, "");
        etHeight.characters = 5;
        heightRow.add("statictext", undefined, rulerInfo.label);
        // 高さの下に自動サイズ調整（heightRow の外に置く。ONのあいだ heightRow はディムするため）
        // Auto-size sits under the height (outside heightRow, which gets dimmed while it is on)
        var chkAutoSize = frameSizePanel.add("checkbox", undefined, getLabel("checkbox.autoSize"));
        chkAutoSize.helpTip = getLabel("tip.autoSize");
        // 英語UIでは chars 計算は不正確なため使用不可にする
        if (currentLanguage !== "ja") {
            etCharsPerLine.enabled = false;
            lblCharsPerLine.enabled = false;
        }

        // 左カラム：オフセット（パネル名で足りるので、チェックボックスのラベルは省く）
        // Left column: offset (the panel title says it, so the checkbox label is dropped)
        var offsetPanel = leftColumn.add("panel", undefined, getLabel("panel.offset"));
        applyPanelLayout(offsetPanel);
        var spacingRow = offsetPanel.add("group");
        var chkSpacing = spacingRow.add("checkbox", undefined, "");
        var etSpacing = spacingRow.add("edittext", undefined, "0");
        etSpacing.characters = 4;
        var lblSpacingUnit = spacingRow.add("statictext", undefined, rulerInfo.label);
        etSpacing.enabled = false;
        lblSpacingUnit.enabled = false;

        // 左カラム：テキストの配置（アイコンボタン。ラベルはツールチップで見せる）
        // Left column: text alignment (icon buttons; the labels live in the tooltips)
        var textAlignPanel = leftColumn.add("panel", undefined, getLabel("panel.textAlign"));
        applyPanelLayout(textAlignPanel, 4);
        textAlignPanel.orientation = "row";
        textAlignPanel.alignChildren = ["center", "center"];
        textAlignPanel.helpTip = getLabel("tip.textAlign");

        // ボタン風で変換したときは中央、それ以外は上揃えで開く
        var alignState = { activeId: getAlignmentId(initialAlignmentValue), isLight: isLightTheme };
        var alignButtons = [];
        for (var alignIndex = 0; alignIndex < ALIGN_OPTIONS.length; alignIndex++) {
            var alignOption = ALIGN_OPTIONS[alignIndex];
            var alignButton = textAlignPanel.add("button", undefined, "");
            alignButton.helpTip = getLabel(alignOption.labelKey);
            alignButton.preferredSize = [26, 26];
            alignButton.minimumSize = [26, 26];
            alignButton.maximumSize = [26, 26];
            alignButton.alignId = alignOption.id;
            alignButton.iconType = alignOption.icon;
            alignButton.onDraw = function () {
                drawAlignIcon(this, this.alignId === alignState.activeId, alignState.isLight);
            };
            alignButtons.push(alignButton);
        }

        // ボタンエリア：左に［テキストを分離...］、右に［キャンセル］［OK］
        // Button area: "Separate text..." on the left, Cancel and OK on the right
        var bottomBar = adjustDialog.add("group");
        bottomBar.orientation = "row";
        bottomBar.alignment = "fill";
        bottomBar.alignChildren = ["fill", "center"];
        var btnSeparateText = bottomBar.add("button", undefined, getLabel("button.separateText"));
        btnSeparateText.alignment = ["left", "center"];
        btnSeparateText.helpTip = getLabel("tip.separateText");
        var bottomButtonRow = bottomBar.add("group");
        bottomButtonRow.alignment = ["right", "center"];
        var btnCancelAdjust = bottomButtonRow.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnRun = bottomButtonRow.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        // 状態変数
        var isPreviewActive = false;
        var fontFitMode = "none";
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

        /* ローカル関数 / Local functions */

        /* テキストの配置ボタンを再描画する / Repaint the text-alignment buttons */
        function redrawAlignButtons() {
            for (var i = 0; i < alignButtons.length; i++) {
                try { alignButtons[i].notify("onDraw"); } catch (e) { }
            }
            try { adjustDialog.update(); } catch (e2) { }
        }

        /* 行揃えボタンを再描画する / Repaint the justification buttons */
        function redrawJustifyButtons() {
            for (var i = 0; i < justifyButtons.length; i++) {
                try { justifyButtons[i].notify("onDraw"); } catch (e) { }
            }
            try { adjustDialog.update(); } catch (e2) { }
        }

        /* メニュー指定を解除し、設定済みのタブも削除する / Drop the Menu role and clear the tab stops it set */
        function cancelMenuRole() {
            if (roleState.activeId !== "menu") return;
            roleState.activeId = "";
            roleState.tabMode = "clear";
            selectRadio(roleRadios, null);
        }

        /* 行揃えを選ぶ（メニューの右揃え以外にしたら、メニュー指定を解除する）
           Pick a justification (anything other than Menu's right alignment drops the Menu role) */
        function setJustification(justifyId) {
            justifyState.activeId = justifyId;
            if (justifyId !== ROLE_PRESETS.menu.justifyId) { cancelMenuRole(); }
            redrawJustifyButtons();
        }

        /* 幅から差し引く余白（間隔×2＋左右インデント）/ Horizontal space taken out of the width (spacing x2 + both indents) */
        function getWidthAdjustmentPt() {
            var spacingPt = chkSpacing.value ? (parseFloat(etSpacing.text) || 0) * rulerInfo.toPt : 0;
            var leftIndentPt = (parseFloat(etLeftIndent.text) || 0) * rulerInfo.toPt;
            var rightIndentPt = chkLinkIndents.value ? leftIndentPt : ((parseFloat(etRightIndent.text) || 0) * rulerInfo.toPt);
            return 2 * spacingPt + leftIndentPt + rightIndentPt;
        }

        /* 幅とフォントサイズから字詰め欄を更新する / Refresh the chars-per-line field from the width and font size */
        function updateCharsPerLineField() {
            if (currentFontSize <= 0) return;
            var widthPt = (parseFloat(etWidth.text) || 0) * rulerInfo.toPt;
            etCharsPerLine.text = Math.round(((widthPt - getWidthAdjustmentPt()) / currentFontSize) * 100) / 100;
        }

        /* 選択フレームの現在値をダイアログに読み込む / Load the frame's current values into the dialog */
        function loadValuesFromFrame(sourceFrame) {
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
                justifyState.activeId = getJustificationId(justification);
            } catch (e) { justifyState.activeId = "left"; }
            try {
                var spacingPt = sourceFrame.spacing || 0;
                etSpacing.text = Math.round((spacingPt / rulerInfo.toPt) * 100) / 100;
                chkSpacing.value = (spacingPt !== 0);
                updateSpacingEnabled();
            } catch (e) { }
            try {
                var firstParaAttrs = sourceFrame.paragraphs.length > 0 ? sourceFrame.paragraphs[0].paragraphAttributes : null;
                var leftIndentPt = firstParaAttrs ? (firstParaAttrs.leftIndent || 0) : 0;
                var rightIndentPt = firstParaAttrs ? (firstParaAttrs.rightIndent || 0) : 0;
                etLeftIndent.text = Math.round((leftIndentPt / rulerInfo.toPt) * 100) / 100;
                etRightIndent.text = Math.round((rightIndentPt / rulerInfo.toPt) * 100) / 100;
                // 左右が違うテキストは連動を外して開く（右の値を潰さないため）
                // Open with the link off when the two differ, so the right value is not overwritten
                if (Math.abs(leftIndentPt - rightIndentPt) > 0.01) { chkLinkIndents.value = false; }
            } catch (e) { }
            // 設定済みならその値を表示して適用対象にする（表示＝実際の設定なので、当ててもフレームは変わらない）。
            // 未設定のときは、変換直後だけ初期値（DEFAULT_KINSOKU / DEFAULT_MOJIKUMI_INDEX）を当て、
            // 既存フレームでは「なし」を表示して勝手に付けない。混在（-2）のときは未選択にして触らない。
            // A value that is already set is shown and applied back (a no-op, but it carries over when separating).
            // When nothing is set, the defaults apply to freshly converted frames only; existing frames show "None".
            var frameKinsoku = getKinsokuId(sourceFrame);
            if (frameKinsoku !== "None") {
                selectChoiceByValue(kinsokuDropdown, KINSOKU_CHOICES, "id", frameKinsoku);
                userTouched.kinsoku = true;
            } else if (!isNewlyConverted) {
                selectChoiceByValue(kinsokuDropdown, KINSOKU_CHOICES, "id", "None");
                userTouched.kinsoku = true;
            }
            var frameMojikumi = getMojikumiIndex(sourceFrame);
            if (frameMojikumi >= 0) {
                selectChoiceByValue(mojikumiDropdown, MOJIKUMI_CHOICES, "index", frameMojikumi);
                userTouched.mojikumi = true;
            } else if (frameMojikumi === -2) {
                mojikumiDropdown.selection = null;
            } else if (!isNewlyConverted) {
                selectChoiceByValue(mojikumiDropdown, MOJIKUMI_CHOICES, "index", -1);
                userTouched.mojikumi = true;
            }

            var leadingPercent = getAutoLeadingPercent(sourceFrame);
            etLeadingPercent.text = (leadingPercent > 0) ? Math.round(leadingPercent * 10) / 10 : "";
            updateLeadingEffective();

            updateCharsPerLineField();
        }

        /* ダイアログの入力を1つの設定オブジェクトにまとめる / Collect the dialog inputs into one settings object */
        function readAdjustmentSettings() {
            var leftIndentPt = (parseFloat(etLeftIndent.text) || 0) * rulerInfo.toPt;

            // 幅/高さの検証はここで1回だけ行う（不正なら null にして書き込まない）
            // Width and height are validated once here; null means the value is not written back
            var widthValue = validateSizeField(etWidth, lastValidWidth);
            var heightValue = validateSizeField(etHeight, lastValidHeight);
            if (widthValue !== null) { lastValidWidth = widthValue; }
            if (heightValue !== null) { lastValidHeight = heightValue; }

            return {
                shrinkFont: (fontFitMode === "shrink"),
                fitFont: (fontFitMode === "fit"),
                autoSize: chkAutoSize.value,
                tabMode: roleState.tabMode,
                justification: getJustificationValue(justifyState.activeId),
                alignment: userTouched.alignment ? getAlignmentValue(alignState.activeId) : null,
                widthPt: (widthValue !== null) ? widthValue * rulerInfo.toPt : null,
                heightPt: (heightValue !== null && userTouched.height) ? heightValue * rulerInfo.toPt : null,
                leadingPercent: parseFloat(etLeadingPercent.text),
                kinsoku: (userTouched.kinsoku && kinsokuDropdown.selection) ? KINSOKU_CHOICES[kinsokuDropdown.selection.index].id : null,
                mojikumiIndex: (userTouched.mojikumi && mojikumiDropdown.selection) ? MOJIKUMI_CHOICES[mojikumiDropdown.selection.index].index : -2,
                leftIndentPt: leftIndentPt,
                rightIndentPt: chkLinkIndents.value ? leftIndentPt
                    : ((parseFloat(etRightIndent.text) || 0) * rulerInfo.toPt),
                spacingPt: chkSpacing.value ? (parseFloat(etSpacing.text) || 0) * rulerInfo.toPt : 0
            };
        }

        /* エリア内文字のフレームサイズ・行揃え・配置などを適用する / Apply frame size, justification and placement to Area Type */
        function adjustAreaTextFrame(areaTextFrame, settings, forPreview) {
            try {
                areaTextFrame.spacing = settings.spacingPt;
                if (settings.widthPt !== null) { areaTextFrame.textPath.width = settings.widthPt; }
                // 自動サイズ調整ONのあいだは枠が文字に追従するので、高さには触らない
                // While auto-size is on the frame follows the text, so the height is left alone
                if (settings.heightPt !== null && !settings.autoSize) { areaTextFrame.textPath.height = settings.heightPt; }
            } catch (e) { }

            if (settings.shrinkFont) { shrinkFontToFit(areaTextFrame); }
            else if (settings.fitFont) { fitFontSizeToFrame(areaTextFrame); }

            // プレビュー中は app.doScript 経由の処理を走らせない（不安定化・クラッシュ回避）
            if (settings.autoSize && !forPreview) { enableFrameAutoSize(areaTextFrame); }

            applyParagraphSettings(areaTextFrame, settings);
            applyAreaTextFrameAlignment(areaTextFrame, settings.alignment, forPreview);
        }

        /* 対象のエリア内文字すべてに現在の設定を適用する / Apply the current settings to every target Area Type frame */
        function applyAdjustments(forPreview) {
            var settings = readAdjustmentSettings();

            // 固定ターゲットを優先して、プレビューが確実に反映されるようにする
            var savedSelection = [];
            var previousSelection = app.activeDocument.selection;
            for (var savedIndex = 0; savedIndex < previousSelection.length; savedIndex++) { savedSelection.push(previousSelection[savedIndex]); }
            var framesToAdjust = getTargetAreaFrames();

            for (var i = framesToAdjust.length - 1; i >= 0; i--) {
                var areaTextFrame = framesToAdjust[i];
                if (areaTextFrame.typename !== "TextFrame" || areaTextFrame.kind !== TextType.AREATEXT) continue;
                // 1フレームで失敗しても残りを処理できるようにする
                try { adjustAreaTextFrame(areaTextFrame, settings, forPreview); } catch (e) { }
            }

            if (savedSelection.length > 0) { app.activeDocument.selection = savedSelection; }
            app.redraw();
        }

        /* 適用中のプレビューを取り消して未適用の状態に戻す / Revert the active preview back to the unapplied state */
        function revertPreview() {
            if (!isPreviewActive) return;
            // undo に失敗したときはプレビューが残るので、フラグを倒さず次の機会にやり直す
            // A failed undo leaves the preview in place, so the flag stays up for the next attempt
            try { app.undo(); } catch (e) { return; }
            try { app.redraw(); } catch (e2) { }
            isPreviewActive = false;

            // undo 後は参照が無効化されるため、対象を取り直す
            refreshTargetAreaFrames();
        }

        /* プレビューを貼り直す（プレビューは常にON）/ Re-apply the preview (preview is always on) */
        function updatePreview() {
            revertPreview();
            applyAdjustments(true);
            isPreviewActive = true;
        }

        /* 自動サイズ調整ONのあいだは高さ欄を使えなくする（枠が文字に追従して指定できないため）
           Disable the height field while auto-size is on (the frame follows the text, so it cannot be set) */
        function updateHeightEnabled() {
            heightRow.enabled = !chkAutoSize.value;
        }

        /* 連動中は右インデント欄を使えなくする（左の値をそのまま使うため）
           Disable the right indent field while linked (it just follows the left value) */
        function updateRightIndentEnabled() {
            etRightIndent.enabled = !chkLinkIndents.value;
        }

        /* オフセット欄の使用可否を更新する / Update whether the offset field can be used */
        function updateSpacingEnabled() {
            etSpacing.enabled = chkSpacing.value;
            lblSpacingUnit.enabled = chkSpacing.value;
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

        /* 対象のエリア内文字ごとにダイナミックアクションを流す
           Run a dynamic action over each target Area Type frame
           アクションは選択に効くので選択を付け替えながら実行し、最後に元の対象へ戻す
           The action works on the selection, so it is swapped per frame and restored at the end */
        function runActionOnTargetFrames(runOnFrame) {
            var framesToAdjust = getTargetAreaFrames();
            var restoreSelection = [];
            for (var restoreIndex = 0; restoreIndex < framesToAdjust.length; restoreIndex++) { restoreSelection.push(framesToAdjust[restoreIndex]); }

            for (var i = 0; i < framesToAdjust.length; i++) {
                var frame = framesToAdjust[i];
                if (frame && frame.typename === "TextFrame" && frame.kind === TextType.AREATEXT) {
                    runOnFrame(frame);
                }
            }

            try { app.activeDocument.selection = restoreSelection; } catch (e) { }
            app.redraw();
        }

        /* 現在のテキストの配置を対象フレームへ本適用する / Commit the current text alignment to the target frames */
        function applyTextAlignment() {
            var alignmentValue = getAlignmentValue(alignState.activeId);
            runActionOnTargetFrames(function (frame) {
                applyAreaTextFrameAlignment(frame, alignmentValue, false);
            });
        }

        /* 対象フレームの自動サイズ調整を切り替える / Switch auto-size on or off across the target frames */
        function setAutoSizeOnTargets(turnOn) {
            runActionOnTargetFrames(function (frame) {
                if (turnOn) { enableFrameAutoSize(frame); } else { disableFrameAutoSize(frame); }
            });
        }

        /* チェックボックスの状態を対象フレームへ本適用する / Commit the checkbox state to the target frames */
        function applyAutoSize() {
            setAutoSizeOnTargets(chkAutoSize.value);
        }

        /* フォントサイズ欄の値を対象フレームに適用する / Apply the font-size field to the target frames */
        function applyFontSizeFromField() {
            var newSize = parseFloat(etFontSize.text) || 0;
            if (newSize <= 0) return;
            // プレビュー分を取り消してから本適用する（そうしないと後の undo がサイズ変更を巻き戻す）
            revertPreview();
            currentFontSize = newSize;
            var framesToAdjust = getTargetAreaFrames();
            for (var i = 0; i < framesToAdjust.length; i++) {
                if (framesToAdjust[i].typename === "TextFrame" && framesToAdjust[i].kind === TextType.AREATEXT) {
                    try { framesToAdjust[i].textRange.characterAttributes.size = newSize; } catch (e) { }
                }
            }
            updateCharsPerLineField();
            updateLeadingEffective();
            app.redraw();
            updatePreview();
        }

        /* 幅の変更を検証し、文字数表示とプレビューを更新する / Validate a width change, then refresh the chars-per-line field and the preview */
        function onWidthChange() {
            var widthValue = validateSizeField(etWidth, lastValidWidth);
            if (widthValue !== null) { lastValidWidth = widthValue; }
            updateCharsPerLineField();
            updatePreview();
        }
        /* 高さの変更を検証してプレビューを更新する（以後、高さを枠へ書き戻す対象にする）
           Validate a height change and refresh the preview (the height is written back from now on) */
        function onHeightChange() {
            var heightValue = validateSizeField(etHeight, lastValidHeight);
            if (heightValue !== null) { lastValidHeight = heightValue; }
            userTouched.height = true;
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
                for (var i = 0; i < JUSTIFY_OPTIONS.length; i++) {
                    if (JUSTIFY_OPTIONS[i].shortcut !== keyName) continue;
                    setJustification(JUSTIFY_OPTIONS[i].id);
                    event.preventDefault();
                    updatePreview();
                    return;
                }
            });
        }

        /* イベントハンドラ / Event handlers */
        bindExclusiveRadios(roleRadios, function (radio) {
            if (radio === radRoleBody) { applyRolePreset("body"); }
            else if (radio === radRoleHeading) { applyRolePreset("heading"); }
            else { applyRolePreset("menu"); }
        });
        for (var buttonIndex = 0; buttonIndex < justifyButtons.length; buttonIndex++) {
            justifyButtons[buttonIndex].onClick = function () {
                setJustification(this.justifyId);
                updatePreview();
            };
        }
        for (var alignButtonIndex = 0; alignButtonIndex < alignButtons.length; alignButtonIndex++) {
            alignButtons[alignButtonIndex].onClick = function () {
                alignState.activeId = this.alignId;
                userTouched.alignment = true;
                redrawAlignButtons();
                // プレビューを外してから本適用し、あらためてプレビューを貼り直す
                revertPreview();
                applyTextAlignment();
                updatePreview();
            };
        }
        addJustificationShortcuts(adjustDialog);

        /* 種別プリセットをダイアログに反映して適用する / Push a role preset into the dialog and apply it */
        function applyRolePreset(roleId) {
            var preset = ROLE_PRESETS[roleId];
            if (!preset) return;
            roleState.activeId = roleId;
            roleState.tabMode = preset.tabMode;
            etLeadingPercent.text = preset.leadingPercent;
            updateLeadingEffective();
            setJustification(preset.justifyId);
            selectChoiceByValue(kinsokuDropdown, KINSOKU_CHOICES, "id", preset.kinsoku);
            selectChoiceByValue(mojikumiDropdown, MOJIKUMI_CHOICES, "index", preset.mojikumiIndex);
            alignState.activeId = preset.alignId;
            // 種別はまとめて指定するものなので、関係する項目をすべて適用対象にする
            userTouched.alignment = true;
            userTouched.kinsoku = true;
            userTouched.mojikumi = true;
            redrawAlignButtons();

            // テキストの配置はアクションで適用するため、プレビューを外した状態で流す
            revertPreview();
            applyTextAlignment();
            updatePreview();
        }

        /* 対象フレームの現在の幅・高さを欄に取り込む（自動サイズ調整やフィットで変わるため）
           Pull the target frame's current width and height into the fields (auto-size and the font-fit passes change them) */
        function refreshFrameSizeFields() {
            var framesToAdjust = getTargetAreaFrames();
            for (var i = 0; i < framesToAdjust.length; i++) {
                var frame = framesToAdjust[i];
                if (!frame || frame.typename !== "TextFrame" || frame.kind !== TextType.AREATEXT) continue;
                try {
                    etWidth.text = Math.round((frame.textPath.width / rulerInfo.toPt) * 100) / 100;
                    etHeight.text = Math.round((frame.textPath.height / rulerInfo.toPt) * 100) / 100;
                    lastValidWidth = parseFloat(etWidth.text);
                    lastValidHeight = parseFloat(etHeight.text);
                } catch (e) { }
                return;
            }
        }

        /* 対象フレームの現在のフォントサイズを欄に取り込む / Pull the target frame's current font size into the field */
        function refreshFontSizeField() {
            var framesToAdjust = getTargetAreaFrames();
            var newSize = 0;
            for (var i = 0; i < framesToAdjust.length; i++) {
                var frame = framesToAdjust[i];
                if (frame && frame.typename === "TextFrame" && frame.kind === TextType.AREATEXT) {
                    newSize = getRepresentativeFontSize(frame);
                    break;
                }
            }
            if (newSize <= 0) return;
            currentFontSize = newSize;
            etFontSize.text = Math.round(newSize * 100) / 100;
            updateCharsPerLineField();
            updateLeadingEffective();
        }

        /* フォントサイズ自動調整を本適用する（プレビュー分は先に取り消す）/ Commit a font-fit pass (revert the preview first) */
        function runFontFit(mode) {
            // 改行があると行数を保ったまま収められず、極端に小さいサイズになるため中止する
            var framesToAdjust = getTargetAreaFrames();
            for (var i = 0; i < framesToAdjust.length; i++) {
                var frame = framesToAdjust[i];
                if (frame && frame.typename === "TextFrame" && frame.kind === TextType.AREATEXT && hasLineBreak(frame)) {
                    alert(getLabel("alert.lineBreakNotSupported"));
                    return;
                }
            }
            revertPreview();

            // 自動サイズ調整がONだと枠が文字に追従してあふれないため、いったんOFFにしてから実行する。
            // ONへ戻すのは applyAdjustments 内（settings.autoSize）。新しい文字サイズで枠が引き直される。
            // With auto-size on the frame follows the text and never oversets, so it is switched off first.
            // applyAdjustments turns it back on (settings.autoSize), redrawing the frame around the new size.
            if (chkAutoSize.value) { setAutoSizeOnTargets(false); }

            fontFitMode = mode;
            applyAdjustments(false);
            fontFitMode = "none";

            // 縮小・拡大した結果をダイアログにも戻す
            refreshFontSizeField();
            refreshFrameSizeFields();
        }

        btnShrinkToFit.onClick = function () { runFontFit("shrink"); };
        btnFitFontSize.onClick = function () { runFontFit("fit"); };
        kinsokuDropdown.onChange = function () {
            userTouched.kinsoku = true;
            updatePreview();
        };
        mojikumiDropdown.onChange = function () {
            userTouched.mojikumi = true;
            updatePreview();
        };
        chkAutoSize.onClick = function () {
            updateHeightEnabled();
            // プレビューを外してから ON/OFF を本適用し、あらためてプレビューを貼り直す
            revertPreview();
            applyAutoSize();
            // 枠の高さが変わるので、欄の値も追従させる（古い値でプレビューが枠を戻さないように）
            refreshFrameSizeFields();
            updatePreview();
        };
        chkSpacing.onClick = function () {
            if (chkSpacing.value) { etSpacing.text = "1"; }
            updateSpacingEnabled();
            onIndentOrSpacingChange();
        };
        chkLinkIndents.onClick = function () {
            // 連動中は右を左に追従させるので、右の欄は触れないようにする
            if (chkLinkIndents.value) { etRightIndent.text = etLeftIndent.text; }
            updateRightIndentEnabled();
            onIndentOrSpacingChange();
        };

        etFontSize.onChange = applyFontSizeFromField;
        etLeadingPercent.onChange = onLeadingPercentChange;
        etLeadingEffective.onChange = onLeadingEffectiveChange;
        etSpacing.onChange = onIndentOrSpacingChange;
        etWidth.onChange = onWidthChange;
        etHeight.onChange = onHeightChange;
        etCharsPerLine.onChange = onCharsPerLineChange;
        etLeftIndent.onChange = onIndentOrSpacingChange;
        etRightIndent.onChange = onIndentOrSpacingChange;
        changeValueByArrowKey(etFontSize, false, applyFontSizeFromField);
        changeValueByArrowKey(etLeadingPercent, false, onLeadingPercentChange);
        changeValueByArrowKey(etLeadingEffective, false, onLeadingEffectiveChange);
        changeValueByArrowKey(etSpacing, false, onIndentOrSpacingChange);
        changeValueByArrowKey(etWidth, false, onWidthChange);
        changeValueByArrowKey(etHeight, false, onHeightChange);
        changeValueByArrowKey(etCharsPerLine, false, onCharsPerLineChange);
        changeValueByArrowKey(etLeftIndent, false, onIndentOrSpacingChange);
        changeValueByArrowKey(etRightIndent, false, onIndentOrSpacingChange);

        btnSeparateText.onClick = function () {
            // プレビューを外し、いま見えている調整を本適用してから分離する
            // （囲み罫はフレームの現在の大きさから作るため）
            // Revert the preview and commit what is on screen before separating,
            // since the rectangle is built from the frame's current size
            revertPreview();
            applyAdjustments(false);
            var currentTargets = getTargetAreaFrames();
            var framesToSeparate = [];
            for (var i = 0; i < currentTargets.length; i++) { framesToSeparate.push(currentTargets[i]); }

            if (showSeparateTextDialog(doc, framesToSeparate, readAdjustmentSettings())) {
                // 分離するとエリア内文字が無くなるので、調整ダイアログも閉じる
                adjustDialog.close(1);
            } else {
                // 分離しなかったときはプレビューを貼り直して調整を続ける
                updatePreview();
            }
        };
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

        updateSpacingEnabled();
        updateHeightEnabled();
        updateRightIndentEnabled();

        // 調整ダイアログを開いた時点からプレビューを反映する
        updatePreview();

        // ［文字あふれ解消］［枠にフィット］だけ、上下2pxずつ詰めて小ぶりにする
        // Make the two font-fit buttons a little shorter (2px off the top and bottom)
        adjustDialog.layout.layout(true);
        trimButtonHeight(btnShrinkToFit, 4);
        trimButtonHeight(btnFitFontSize, 4);

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
            // ポイント文字 または 図形 → 変換ダイアログ / Point text or shape → convert dialog
            showConvertDialog(doc, selection);
        } else if (hasAreaText) {
            // エリア内文字のみ → 調整ダイアログ / Area text only → adjust dialog
            var firstAreaFrame = null;
            for (var areaIndex = 0; areaIndex < selection.length; areaIndex++) {
                if (selection[areaIndex].typename === "TextFrame" && selection[areaIndex].kind === TextType.AREATEXT) {
                    firstAreaFrame = selection[areaIndex];
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
