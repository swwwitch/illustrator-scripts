#targetengine "FontPresetPickerEngine"
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

よく使うフォントを、サイズ・行送り・字間・日本語の文字組みとセットで「定番」として登録し、
一覧から選ぶだけで選択中のテキストへまとめて適用する常駐パレットです。
適用する設定を絞り込んだり、比率やアキを標準の状態へ戻したりもできます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FontPresetPicker.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n3d7f8b58ef88

### Overview

A persistent palette that keeps your go-to fonts together with their size, leading, letter spacing
and Japanese typesetting, and applies the whole set to the selected text with a single click in the list.
You can narrow down what gets applied, and reset scaling and aki back to their default state.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FontPresetPicker.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FontPresetPicker";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-17";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-17";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FontPresetPicker.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FontPresetPicker.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n3d7f8b58ef88"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* プリセットの保存先ファイル名（Folder.userData 直下）/ Preset file name (directly under Folder.userData) */
    var PRESET_FILE_NAME = "FontPresetPicker_presets.json";

    /* パレット位置の保存先ファイル名 / File name that remembers the palette location */
    var LOCATION_FILE_NAME = "FontPresetPicker_location.txt";

    /* 保存ファイルが無いときの初期プリセット / Initial presets used when no file exists yet */
    var DEFAULT_PRESETS = [
        {
            psName: "HiraginoSans-W3", size: 12, leadingPercent: 175, leadingType: "top",
            kern: "metrics", tsume: 0, tracking: 0, align: "roman",
            justify: "left", kinsoku: "Soft_v2", mojikumi: 5
        }
    ];

    /* 行送りが読み取れないときに使う自動行送り量（％）/ Auto-leading amount (%) used when none can be read */
    var DEFAULT_AUTO_LEADING_PERCENT = 175;

    // =========================================
    // レイアウト / Layout
    // =========================================
    var WINDOW_MARGINS = 16;                        /* ウィンドウ外周の余白 / Window margins */
    var WINDOW_SPACING = 12;                        /* ウィンドウ内の要素間隔 / Window spacing */
    var ROW_SPACING = 12;                           /* 行グループ内の間隔 / Spacing inside a row group */
    var PANEL_MARGINS = [12, 16, 12, 12];           /* パネル内の余白 / Margins inside a panel */
    var PANEL_ROW_SPACING = 6;                      /* パネル内の行間隔 / Spacing between rows inside a panel */
    var BUTTON_ROW_TOP_MARGIN = 5;                  /* ボタン行の上余白 / Top margin above the button row */
    var LIST_WIDTH = 360;                           /* 一覧の幅（上下で共通）/ Width shared by both lists */
    var LIST_ROW_HEIGHT = 22;                       /* 一覧1行の高さの目安 / Estimated height of one list row */
    var LIST_FRAME_PADDING = 4;                     /* 一覧の枠ぶんの余白 / Padding for the list frame */
    var FONT_LIST_ROW_COUNT = 5;                    /* 上段に出す行数 / Rows shown in the top list */
    var DETAIL_ROW_COUNT = 12;                      /* 下段の項目数（fillDetailList の行数）/ Detail rows (as many as fillDetailList adds) */
    var DETAIL_VISIBLE_ROWS = DETAIL_ROW_COUNT - 1; /* 下段の高さは1行ぶん詰める / The lower list is one row shorter than its contents */
    var FONT_LIST_SIZE = [LIST_WIDTH, LIST_ROW_HEIGHT * FONT_LIST_ROW_COUNT + LIST_FRAME_PADDING];
    var DETAIL_LIST_SIZE = [LIST_WIDTH, LIST_ROW_HEIGHT * DETAIL_VISIBLE_ROWS + LIST_FRAME_PADDING];
    var DETAIL_COLUMN_WIDTHS = [200, 150];          /* 下段の列幅（項目名／値）。「プロポーショナルメトリクス：」が入る幅 / Bottom columns (label / value), wide enough for the longest label */

    /**
     * パレットの余白・間隔・子要素の揃えをまとめて設定する
     * @param {Window} paletteWindow - 対象のパレット
     * @returns {void}
     */
    function setupPalette(paletteWindow) {
        paletteWindow.orientation = "column";
        paletteWindow.margins = WINDOW_MARGINS;
        paletteWindow.spacing = WINDOW_SPACING;
        paletteWindow.alignChildren = ["fill", "top"];
    }

    /**
     * 行グループの向き・揃え・間隔をまとめて設定する
     * @param {Group} targetGroup - 対象の行グループ
     * @param {string} alignment - 横方向の揃え（"left" / "right" / "fill"）
     * @param {number} spacing - 要素間の間隔（省略時は ROW_SPACING）
     * @returns {void}
     */
    function setupRow(targetGroup, alignment, spacing) {
        targetGroup.orientation = "row";
        targetGroup.alignment = [alignment || "left", "center"];
        targetGroup.alignChildren = ["left", "center"];
        targetGroup.spacing = (typeof spacing === "number") ? spacing : ROW_SPACING;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * UI 言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getUiLanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getUiLanguage();

    /* ラベル定義 / Label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "文字設定の定番", en: "Type Presets" }
        },
        panel: {
            applySettings: { ja: "適用する設定", en: "Settings to Apply" },
            clear: { ja: "クリア", en: "Clear" }
        },
        applyGroup: {
            font: { ja: "フォント", en: "Font" },
            size: { ja: "フォントサイズ", en: "Font Size" },
            leading: { ja: "行送り", en: "Leading" },
            spacing: { ja: "カーニング・ツメ・トラッキング", en: "Kerning, Tsume & Tracking" },
            japanese: { ja: "揃え・禁則・文字組み", en: "Alignment, Kinsoku & Mojikumi" }
        },
        clearGroup: {
            scale: { ja: "水平比率・垂直比率", en: "Horizontal & Vertical Scale" },
            aki: { ja: "文字前後のアキ", en: "Aki Before & After" },
            baselineShift: { ja: "ベースラインシフト", en: "Baseline Shift" }
        },
        field: {
            size: { ja: "フォントサイズ", en: "Font Size" },
            leading: { ja: "行送り", en: "Leading" },
            leadingPercent: { ja: "行送り（%）", en: "Leading (%)" },
            leadingType: { ja: "行送りの基準位置", en: "Leading Basis" },
            kern: { ja: "自動カーニング", en: "Auto Kerning" },
            tsume: { ja: "文字ツメ", en: "Tsume" },
            propMetrics: { ja: "プロポーショナルメトリクス", en: "Proportional Metrics" },
            tracking: { ja: "トラッキング", en: "Tracking" },
            align: { ja: "文字揃え", en: "Character Alignment" },
            justify: { ja: "行揃え", en: "Justification" },
            kinsoku: { ja: "禁則", en: "Kinsoku" },
            mojikumi: { ja: "文字組みアキ量設定", en: "Mojikumi" }
        },
        leadingType: {
            top: { ja: "仮想ボディの上", en: "Top-to-top (virtual body)" },
            baseline: { ja: "欧文ベースライン", en: "Baseline-to-baseline" }
        },
        align: {
            roman: { ja: "欧文ベースライン", en: "Roman Baseline" },
            top: { ja: "仮想ボディの上/右", en: "Embox Top/Right" },
            center: { ja: "中央", en: "Embox Centre" },
            bottom: { ja: "仮想ボディの下/左", en: "Embox Bottom/Left" },
            icftop: { ja: "平均字面の上/右", en: "ICF Box Top/Right" },
            icfbottom: { ja: "平均字面の下/左", en: "ICF Box Bottom/Left" }
        },
        justify: {
            left: { ja: "左揃え", en: "Left" },
            center: { ja: "中央揃え", en: "Center" },
            right: { ja: "右揃え", en: "Right" },
            justifyLeft: { ja: "均等配置（最終行左）", en: "Justify (last left)" },
            justifyCenter: { ja: "均等配置（最終行中央）", en: "Justify (last center)" },
            justifyRight: { ja: "均等配置（最終行右）", en: "Justify (last right)" },
            justifyAll: { ja: "両端揃え", en: "Justify all" }
        },
        kinsoku: {
            none: { ja: "なし", en: "None" },
            hard: { ja: "強い禁則", en: "Strict" },
            soft: { ja: "弱い禁則", en: "Loose" },
            softV2: { ja: "弱い禁則 v2", en: "Loose v2" }
        },
        /* 設定の名前は日本語組版の機能なので、英語UIでも日本語名のまま出す（「なし」だけは一般語）
           The set names are Japanese typesetting features, so they stay in Japanese even in English */
        mojikumi: {
            none: { ja: "なし", en: "None" },
            lineEndFullHalf: { ja: "行末約物全角/半角", en: "行末約物全角/半角" },
            punctHalf: { ja: "約物半角", en: "約物半角" },
            lineEndHalf: { ja: "行末約物半角", en: "行末約物半角" },
            lineEndFull: { ja: "行末約物全角", en: "行末約物全角" },
            punctFull: { ja: "約物全角", en: "約物全角" },
            tight: { ja: "ツメ組み", en: "ツメ組み" },
            solid: { ja: "ベタ組み", en: "ベタ組み" }
        },
        autoKern: {
            mono: { ja: "和文等幅", en: "Metrics - Roman Only" },
            zero: { ja: "0", en: "0" },
            metrics: { ja: "メトリクス", en: "Metrics" },
            optical: { ja: "オプティカル", en: "Optical" }
        },
        button: {
            addPreset: { ja: "追加", en: "Add" },
            overwritePreset: { ja: "上書き", en: "Overwrite" },
            deletePreset: { ja: "削除", en: "Delete" }
        },
        tip: {
            presets: { ja: "登録した定番フォント。選ぶと、選択中のテキストへ設定一式をまとめて適用します。", en: "Saved font presets. Click one to apply the whole set to the selected text." },
            detail: { ja: "一覧で選択中の定番の内訳。［適用する設定］で外したグループはディム表示になります。", en: "What the preset selected in the list holds. Rows left out by \"Settings to Apply\" are dimmed." },
            applyPanel: { ja: "チェックした項目だけを適用します。option（alt）＋クリックで、その項目だけオンと、すべてオンを切り替えます。", en: "Only the ticked items are applied. Option (Alt)-click a box to switch between just that one and all of them." },
            applyFont: { ja: "定番のフォントを適用します。", en: "Apply the preset's font." },
            applySize: { ja: "定番のフォントサイズを適用します。", en: "Apply the preset's font size." },
            applyLeading: { ja: "定番の行送り（自動行送り量）と行送りの基準位置を適用します。", en: "Apply the preset's leading (auto-leading amount) and leading basis." },
            applySpacing: { ja: "定番の自動カーニング・文字ツメ・トラッキングを適用します。", en: "Apply the preset's auto-kerning, Tsume and tracking." },
            applyJapanese: { ja: "定番の文字揃え・行揃え・禁則・文字組みアキ量設定を適用します。", en: "Apply the preset's character alignment, justification, kinsoku and mojikumi." },
            clearPanel: { ja: "チェックした項目を、定番の適用と同時に標準の状態へ戻します。option（alt）＋クリックで、その項目だけオンと、すべてオンを切り替えます。", en: "The ticked items are reset to their defaults as the preset is applied. Option (Alt)-click a box to switch between just that one and all of them." },
            clearScale: { ja: "水平比率・垂直比率を 100% に戻します。", en: "Reset the horizontal and vertical scale to 100%." },
            clearAki: { ja: "文字前のアキ・文字後のアキを「自動」に戻します。", en: "Reset the aki before and after each character to Auto." },
            clearBaselineShift: { ja: "ベースラインシフトを 0 に戻します。", en: "Reset the baseline shift to 0." },
            addPreset: { ja: "選択中のテキストの設定を定番として追加します（同じフォントは上書き）。", en: "Add the selected text's settings as a preset (same font overwrites)." },
            overwritePreset: { ja: "一覧で選択中の定番を、選択中のテキストの設定で上書きします。", en: "Overwrite the preset selected in the list with the selected text's settings." },
            deletePreset: { ja: "一覧で選択中の定番を削除します。", en: "Delete the preset selected in the list." }
        },
        alert: {
            readFailed: { ja: "選択中のテキストから設定を読み取れませんでした。", en: "Could not read the settings from the selected text." },
            applyFailed: { ja: "適用に失敗しました。", en: "Apply failed." }
        }
    };

    /* 禁則 id と表示ラベルの対応（id は paragraphAttributes.kinsoku に渡す値）
       Kinsoku ids paired with their labels (the id is what goes to paragraphAttributes.kinsoku) */
    var KINSOKU_OPTIONS = [
        { id: "None", label: LABELS.kinsoku.none },
        { id: "Hard", label: LABELS.kinsoku.hard },
        { id: "Soft", label: LABELS.kinsoku.soft },
        { id: "Soft_v2", label: LABELS.kinsoku.softV2 }
    ];

    /* 文字組みアキ量設定の対応表
       index は適用に使う doc.mojikumiSet のインデックス（-1＝なし）、
       id は paragraphAttributes.mojikumi が読み取りで返す内部 ID。
       doc.mojikumiSet[i].name は undefined で名前では引けないため、この表が唯一の橋渡しになる。
       「行末約物半角」は読み取ると「約物半角」と同じ ID になり区別できないので id は持たせない。
       index applies via doc.mojikumiSet (-1 = None); id is what the paragraph attribute returns.
       mojikumiSet[i].name is undefined, so this table is the only way across; "line-end punct half"
       reads back as "half-width punctuation", so it carries no id. */
    var MOJIKUMI_OPTIONS = [
        { index: -1, id: "None", label: LABELS.mojikumi.none },
        { index: 0, id: "GyomatsuYakumonoZenkakuOrHankaku", label: LABELS.mojikumi.lineEndFullHalf },
        { index: 1, id: "YakumonoHankaku", label: LABELS.mojikumi.punctHalf },
        { index: 2, id: null, label: LABELS.mojikumi.lineEndHalf },
        { index: 3, id: "GyomatsuYakumonoZenkaku", label: LABELS.mojikumi.lineEndFull },
        { index: 4, id: "YakumonoZenkaku", label: LABELS.mojikumi.punctFull },
        { index: 5, id: "TsumeGumi", label: LABELS.mojikumi.tight },
        { index: 6, id: "BetaGumi", label: LABELS.mojikumi.solid }
    ];

    /**
     * 言語に応じたラベル文字列を取得する
     * @param {Object} labelSet - { ja, en } 形式のラベル定義
     * @returns {string} 現在の UI 言語のラベル
     */
    function getLabel(labelSet) {
        if (!labelSet) return "";
        return labelSet[uiLang] || labelSet.ja || labelSet.en || "";
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {Object} labelSet - { ja, en } 形式のラベル定義
     * @returns {string} コロンを付けた項目名
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // メインエンジンで実行する DOM 処理 / DOM helpers run on the main engine
    //
    // toString() で連結して BridgeTalk 本文に同梱し、メインエンジンで eval される。
    // 注意: ここの関数には JSDoc を付けない（toString() の結果が壊れる）。
    // Shipped via toString(); never attach JSDoc — it corrupts the output.
    // =========================================

    /* グループは中を再帰 / Descend into groups */
    function collectTextRangesFromItem(item, selectedRanges) {
        if (!item) return;
        if (item.typename === "TextFrame") {
            selectedRanges.push(item.textRange);
        } else if (item.typename === "TextRange") {
            selectedRanges.push(item);
        } else if (item.typename === "GroupItem" && item.pageItems) {
            for (var i = 0; i < item.pageItems.length; i++) {
                collectTextRangesFromItem(item.pageItems[i], selectedRanges);
            }
        }
    }

    /* 選択中のテキスト範囲を集める / Collect the selected text ranges */
    function getSelectedTextRanges() {
        var currentSelection = app.activeDocument.selection;
        var selectedRanges = [];
        if (!currentSelection) return selectedRanges;
        /* テキスト編集モードでは selection が配列でなく TextRange / In text-edit mode the selection is a TextRange, not an array */
        if (currentSelection.typename === "TextRange") {
            selectedRanges.push(currentSelection);
            return selectedRanges;
        }
        for (var i = 0; i < currentSelection.length; i++) {
            collectTextRangesFromItem(currentSelection[i], selectedRanges);
        }
        return selectedRanges;
    }

    /* 和文等幅＝欧文のみメトリクス / Japanese equal width = metrics for Roman only */
    function resolveAutoKernType(methodId) {
        if (methodId === "mono") return AutoKernType.METRICSROMANONLY;
        if (methodId === "metrics") return AutoKernType.AUTO;
        if (methodId === "optical") return AutoKernType.OPTICAL;
        return AutoKernType.NOAUTOKERN;
    }

    /* AutoKernType を方式 ID へ / AutoKernType to a method id */
    function kernMethodToId(value) {
        var valueText = String(value);
        if (valueText === String(AutoKernType.AUTO)) return "metrics";
        if (valueText === String(AutoKernType.OPTICAL)) return "optical";
        if (valueText === String(AutoKernType.METRICSROMANONLY)) return "mono";
        if (valueText === String(AutoKernType.NOAUTOKERN)) return "zero";
        return "";
    }

    /* 文字揃え ID を StyleRunAlignmentType へ / Resolve a character-alignment id to StyleRunAlignmentType */
    function resolveAlignment(alignId) {
        if (alignId === "top") return StyleRunAlignmentType.top;
        if (alignId === "center") return StyleRunAlignmentType.center;
        if (alignId === "bottom") return StyleRunAlignmentType.bottom;
        if (alignId === "icftop") return StyleRunAlignmentType.icfTop;
        if (alignId === "icfbottom") return StyleRunAlignmentType.icfBottom;
        return StyleRunAlignmentType.ROMANBASELINE;
    }

    /* StyleRunAlignmentType を ID 文字列へ / StyleRunAlignmentType to an id string */
    function alignmentToId(value) {
        var valueText = String(value);
        if (valueText === String(StyleRunAlignmentType.top)) return "top";
        if (valueText === String(StyleRunAlignmentType.center)) return "center";
        if (valueText === String(StyleRunAlignmentType.bottom)) return "bottom";
        if (valueText === String(StyleRunAlignmentType.icfTop)) return "icftop";
        if (valueText === String(StyleRunAlignmentType.icfBottom)) return "icfbottom";
        if (valueText === String(StyleRunAlignmentType.ROMANBASELINE)) return "roman";
        return "";
    }

    /* 行送りの基準 ID を AutoLeadingType へ / Resolve a leading-basis id to AutoLeadingType */
    function resolveLeadingType(typeId) {
        if (typeId === "baseline") return AutoLeadingType.BOTTOMTOBOTTOM;
        return AutoLeadingType.TOPTOTOP;
    }

    /* AutoLeadingType を ID 文字列へ / AutoLeadingType to an id string */
    function leadingTypeToId(value) {
        return (String(value) === String(AutoLeadingType.BOTTOMTOBOTTOM)) ? "baseline" : "top";
    }

    /* 行揃え ID を Justification へ / Resolve a justification id to Justification */
    function resolveJustification(justifyId) {
        if (justifyId === "center") return Justification.CENTER;
        if (justifyId === "right") return Justification.RIGHT;
        if (justifyId === "justifyLeft") return Justification.FULLJUSTIFYLASTLINELEFT;
        if (justifyId === "justifyCenter") return Justification.FULLJUSTIFYLASTLINECENTER;
        if (justifyId === "justifyRight") return Justification.FULLJUSTIFYLASTLINERIGHT;
        if (justifyId === "justifyAll") return Justification.FULLJUSTIFY;
        return Justification.LEFT;
    }

    /* Justification を ID 文字列へ / Justification to an id string */
    function justificationToId(value) {
        var valueText = String(value);
        if (valueText === String(Justification.LEFT)) return "left";
        if (valueText === String(Justification.CENTER)) return "center";
        if (valueText === String(Justification.RIGHT)) return "right";
        if (valueText === String(Justification.FULLJUSTIFYLASTLINELEFT)) return "justifyLeft";
        if (valueText === String(Justification.FULLJUSTIFYLASTLINECENTER)) return "justifyCenter";
        if (valueText === String(Justification.FULLJUSTIFYLASTLINERIGHT)) return "justifyRight";
        if (valueText === String(Justification.FULLJUSTIFY)) return "justifyAll";
        return "";
    }

    /* 文字属性は範囲へ直接入れる。適用できない範囲は飛ばす（属性ごとに try を書かないための受け皿）
       Character attributes go straight on the range; ranges that can't take them are skipped */
    function forEachRange(ranges, applyToRange) {
        for (var i = 0; i < ranges.length; i++) {
            try { applyToRange(ranges[i]); } catch (e) { }
        }
    }

    /* 段落属性は選択が触れた段落へ入れる（自動行送り量・行揃え・禁則・文字組みが該当）
       Paragraph attributes go on the paragraphs the selection touches */
    function forEachParagraph(ranges, applyToParagraph) {
        for (var i = 0; i < ranges.length; i++) {
            try {
                var paragraphs = ranges[i].paragraphs;
                for (var p = 0; p < paragraphs.length; p++) {
                    try { applyToParagraph(paragraphs[p]); } catch (eParagraph) { }
                }
            } catch (e) { }
        }
    }

    /* 小数第2位まで（10.5pt のような端数は残す）/ Two decimals, so 10.5 pt survives */
    function roundToHundredths(value) {
        return Math.round(value * 100) / 100;
    }

    /* 未インストールのフォントは何もしない（getByName は見つからないと例外）
       Skip fonts that are not installed; getByName throws when there is no match */
    function applyFontToRanges(ranges, psName) {
        var font = null;
        try { font = app.textFonts.getByName(psName); } catch (e) { return; }
        if (!font) return;
        forEachRange(ranges, function (range) { range.characterAttributes.textFont = font; });
    }

    /* フォントサイズ（pt）/ Font size in points */
    function applySizeToRanges(ranges, sizeValue) {
        forEachRange(ranges, function (range) { range.characterAttributes.size = sizeValue; });
    }

    /* 行送りは常に「自動」で入れる。pt は焼き込まず、自動行送り量（％）を段落属性に置いて
       各文字を autoLeading=true にするので、Illustrator 上は常に「自動」表示でサイズに追従する
       Leading always goes in as "Auto": the amount (%) lives on the paragraph and every character
       stays on autoLeading, so Illustrator shows "Auto" and the leading follows the font size */
    function applyLeadingToRanges(ranges, leadingPercent) {
        if (isNaN(leadingPercent) || leadingPercent <= 0) return;
        forEachRange(ranges, function (range) { range.characterAttributes.autoLeading = true; });
        forEachParagraph(ranges, function (paragraph) {
            paragraph.characterAttributes.autoLeading = true;
            paragraph.paragraphAttributes.autoLeadingAmount = leadingPercent;
        });
    }

    /* 行送りの基準（仮想ボディの上／欧文ベースライン）/ Leading basis (virtual body top / Roman baseline) */
    function applyLeadingTypeToRanges(ranges, leadingType) {
        forEachRange(ranges, function (range) { range.leadingType = leadingType; });
    }

    /* メトリクスのときだけプロポーショナルメトリクスを ON / Proportional metrics ON only for Metrics */
    function applyKerningToRanges(ranges, kerningMethod) {
        var useProportionalMetrics = (kerningMethod === AutoKernType.AUTO);
        forEachRange(ranges, function (range) {
            range.characterAttributes.kerningMethod = kerningMethod;
            range.characterAttributes.proportionalMetrics = useProportionalMetrics;
        });
    }

    /* 文字ツメ（0〜100 の百分率）/ Tsume as a percentage from 0 to 100 */
    function applyTsumeToRanges(ranges, tsumePercent) {
        forEachRange(ranges, function (range) { range.characterAttributes.Tsume = tsumePercent; });
    }

    /* トラッキング（1/1000em）/ Tracking in 1/1000 em */
    function applyTrackingToRanges(ranges, trackingValue) {
        forEachRange(ranges, function (range) { range.characterAttributes.tracking = trackingValue; });
    }

    /* 文字揃え（縦方向の揃え）/ Character alignment (vertical) */
    function applyAlignmentToRanges(ranges, alignValue) {
        forEachRange(ranges, function (range) { range.characterAttributes.alignment = alignValue; });
    }

    /* 行揃え（段落属性）/ Justification (a paragraph attribute) */
    function applyJustificationToRanges(ranges, justifyValue) {
        forEachParagraph(ranges, function (paragraph) { paragraph.paragraphAttributes.justification = justifyValue; });
    }

    /* 禁則（段落属性）。「なし」は scripting から設定できず例外になるが forEachParagraph が握る
       Kinsoku; "None" can't be set via scripting and throws, which forEachParagraph swallows */
    function applyKinsokuToRanges(ranges, kinsokuId) {
        forEachParagraph(ranges, function (paragraph) { paragraph.paragraphAttributes.kinsoku = kinsokuId; });
    }

    /* 文字組みアキ量設定（段落属性）。-1＝なし、0以上はドキュメントの mojikumiSet のインデックス
       Mojikumi; -1 means "None", 0 and up index the document's mojikumiSet */
    function applyMojikumiToRanges(ranges, mojikumiIndex) {
        var mojikumiValue;
        try {
            mojikumiValue = (mojikumiIndex === -1) ? "なし" : app.activeDocument.mojikumiSet[mojikumiIndex];
        } catch (eSet) {
            return;
        }
        forEachParagraph(ranges, function (paragraph) { paragraph.paragraphAttributes.mojikumi = mojikumiValue; });
    }

    /* ［クリア］でチェックされた文字属性を標準値へ戻す。定番には持たせず、適用のたびに初期化する
       文字前後のアキの「自動」は -1（2026-09-17 実測。0 は「アキなし」で別物）。
       左右を1回の取得でまとめて書くと片方が効かないことがあるため、別々に流す
       Reset the character attributes ticked under "Clear"; aki "auto" is -1 (0 means no aki),
       and the left and right values go in separate passes because one can otherwise be dropped */
    function clearRangeAttributes(ranges, clearOptions) {
        if (!clearOptions) return;
        if (clearOptions.scale) {
            forEachRange(ranges, function (range) {
                range.characterAttributes.horizontalScale = 100;
                range.characterAttributes.verticalScale = 100;
            });
        }
        if (clearOptions.aki) {
            forEachRange(ranges, function (range) { range.characterAttributes.akiLeft = -1; });
            forEachRange(ranges, function (range) { range.characterAttributes.akiRight = -1; });
        }
        if (clearOptions.baselineShift) {
            forEachRange(ranges, function (range) { range.characterAttributes.baselineShift = 0; });
        }
    }

    /* 自動行送りはサイズから算出されるので、サイズ→行送りの順に流す
       持っていないプロパティは飛ばすので、「適用する設定」で絞ったプリセットもそのまま渡せる
       Auto leading follows the size, so apply size first; absent properties are skipped,
       so a preset narrowed by the "properties to apply" panel can be passed as is */
    function applyPresetToRanges(ranges, preset) {
        if (!ranges.length || !preset) return;
        if (preset.psName) applyFontToRanges(ranges, preset.psName);
        if (preset.size !== undefined && preset.size !== null) applySizeToRanges(ranges, preset.size);
        if (preset.leadingType) applyLeadingTypeToRanges(ranges, resolveLeadingType(preset.leadingType));
        if (preset.leadingPercent !== undefined && preset.leadingPercent !== null) applyLeadingToRanges(ranges, preset.leadingPercent);
        if (preset.kern) applyKerningToRanges(ranges, resolveAutoKernType(preset.kern));
        if (preset.tsume !== undefined && preset.tsume !== null) applyTsumeToRanges(ranges, preset.tsume);
        if (preset.tracking !== undefined && preset.tracking !== null) applyTrackingToRanges(ranges, preset.tracking);
        if (preset.align) applyAlignmentToRanges(ranges, resolveAlignment(preset.align));
        if (preset.justify) applyJustificationToRanges(ranges, resolveJustification(preset.justify));
        if (preset.kinsoku) applyKinsokuToRanges(ranges, preset.kinsoku);
        if (preset.mojikumi !== undefined && preset.mojikumi !== null) applyMojikumiToRanges(ranges, preset.mojikumi);
    }

    /* 行送りは「自動」前提で％として読む。自動なら段落の自動行送り量、固定値なら
       フォントサイズとの比から逆算する。どちらも読めなければ NaN（パレット側で既定に寄せる）
       Leading is read as an auto-leading percentage: the paragraph amount when it is auto,
       otherwise back-calculated against the font size; NaN when neither can be read */
    function readLeadingPercent(range, fontSize) {
        var charAttrs = range.characterAttributes;
        if (charAttrs.autoLeading === true) {
            try {
                var paragraphs = range.paragraphs;
                if (paragraphs.length) {
                    var amount = paragraphs[0].paragraphAttributes.autoLeadingAmount;
                    if (!isNaN(amount) && amount > 0) return roundToHundredths(amount);
                }
            } catch (eParagraph) { }
            return NaN;
        }
        if (!isNaN(charAttrs.leading) && fontSize > 0) return roundToHundredths(charAttrs.leading / fontSize * 100);
        return NaN;
    }

    /* 行送りの基準。読めなければ既定の「仮想ボディの上」/ Leading basis; falls back to top-to-top */
    function readLeadingTypeId(range) {
        try { return leadingTypeToId(range.leadingType); } catch (e) { return "top"; }
    }

    /* 行揃えを id 文字列へ / Justification to an id string */
    function readJustifyId(range) {
        try {
            var paragraphs = range.paragraphs;
            if (!paragraphs.length) return "";
            return justificationToId(paragraphs[0].paragraphAttributes.justification);
        } catch (e) {
            return "";
        }
    }

    /* 禁則を id 文字列へ。禁則「なし」の段落は getter が例外（Error 9563）を投げるので "None"
       Kinsoku to an id string; a no-kinsoku paragraph throws on read (Error 9563), so "None" */
    function readKinsokuId(range) {
        try {
            var paragraphs = range.paragraphs;
            if (!paragraphs.length) return "None";
            var kinsokuValue = paragraphs[0].paragraphAttributes.kinsoku;
            if (kinsokuValue === undefined || kinsokuValue === null || kinsokuValue === "") return "None";
            return String(kinsokuValue);
        } catch (e) {
            return "None";
        }
    }

    /* 文字組みアキ量設定を内部 ID へ。getter は "TsumeGumi" のような ID 文字列を返し、
       「なし」の段落は例外（the requested attribute is undefined）になるので "None" を返す
       Mojikumi to its internal id; the getter returns a string like "TsumeGumi",
       and a paragraph set to None throws, which we report as "None" */
    function readMojikumiId(range) {
        try {
            var paragraphs = range.paragraphs;
            if (!paragraphs.length) return "None";
            var mojikumiAttr = paragraphs[0].paragraphAttributes.mojikumi;
            if (mojikumiAttr === undefined || mojikumiAttr === null || mojikumiAttr === "") return "None";
            return String(mojikumiAttr);
        } catch (e) {
            return "None";
        }
    }

    /* 先頭の範囲を代表値として現在の設定を読む / The first range represents the current settings */
    function readPresetFromRanges(ranges) {
        if (!ranges.length) return null;
        try {
            var charAttrs = ranges[0].characterAttributes;
            var currentFont = charAttrs.textFont;
            if (!currentFont || !currentFont.name) return null;
            var fontSize = roundToHundredths(charAttrs.size);
            return {
                psName: currentFont.name,
                size: fontSize,
                leadingPercent: readLeadingPercent(ranges[0], fontSize),
                leadingType: readLeadingTypeId(ranges[0]),
                kern: kernMethodToId(charAttrs.kerningMethod),
                tsume: Math.round(charAttrs.Tsume),
                tracking: Math.round(charAttrs.tracking),
                align: alignmentToId(charAttrs.alignment),
                justify: readJustifyId(ranges[0]),
                kinsoku: readKinsokuId(ranges[0]),
                mojikumiId: readMojikumiId(ranges[0])
            };
        } catch (e) {
            return null;
        }
    }

    /* 戻り値は "OK:<payload>" / "ERR:<msg>"。readPreset の payload はタブ区切りを URL エンコードしたもの
       Returns "OK:<payload>" / "ERR:<msg>"; readPreset's payload is URL-encoded tab-separated fields */
    function dispatch(actionId, params) {
        var TAB = String.fromCharCode(9), LF = String.fromCharCode(10);

        /* フォント名の解決だけはドキュメント不要 / Resolving font names is the one action that needs no document */
        if (actionId === "resolveFontNames") {
            var names = params.psList ? params.psList.split(LF) : [];
            var resolved = [];
            for (var nameIndex = 0; nameIndex < names.length; nameIndex++) {
                var displayName = names[nameIndex];
                try {
                    var resolvedFont = app.textFonts.getByName(names[nameIndex]);
                    if (resolvedFont) displayName = resolvedFont.family + (resolvedFont.style ? " " + resolvedFont.style : "");
                } catch (eResolve) { }
                resolved.push(names[nameIndex] + TAB + displayName);
            }
            return "OK:" + encodeURIComponent(resolved.join(LF));
        }

        /* ドキュメントが無いのは異常ではないので「選択ゼロ」として返す / No document is a normal state, reported as zero selection */
        if (app.documents.length === 0) return "OK:0";

        var ranges = getSelectedTextRanges();

        if (actionId === "getState") return "OK:" + ranges.length;

        if (actionId === "readPreset") {
            var currentPreset = readPresetFromRanges(ranges);
            if (!currentPreset) return "OK:";
            var leadingField = isNaN(currentPreset.leadingPercent) ? "" : String(currentPreset.leadingPercent);
            return "OK:" + encodeURIComponent([
                currentPreset.psName,
                currentPreset.size,
                leadingField,
                currentPreset.leadingType,
                currentPreset.kern,
                currentPreset.tsume,
                currentPreset.tracking,
                currentPreset.align,
                currentPreset.justify,
                currentPreset.kinsoku,
                currentPreset.mojikumiId
            ].join(TAB));
        }

        if (actionId === "applyPreset") {
            clearRangeAttributes(ranges, params.clear);
            applyPresetToRanges(ranges, params.preset);
            app.redraw();
            return "OK:" + ranges.length;
        }

        return "OK:0";
    }

    // =========================================
    // メインエンジン委譲 / Main-engine delegation (BridgeTalk)
    // =========================================

    var WORKER_FUNCS = [
        collectTextRangesFromItem, getSelectedTextRanges, resolveAutoKernType, kernMethodToId,
        resolveAlignment, alignmentToId, resolveLeadingType, leadingTypeToId, resolveJustification, justificationToId,
        forEachRange, forEachParagraph, roundToHundredths,
        applyFontToRanges, applySizeToRanges, applyLeadingToRanges, applyLeadingTypeToRanges, applyKerningToRanges,
        applyTsumeToRanges, applyTrackingToRanges, applyAlignmentToRanges, applyJustificationToRanges,
        applyKinsokuToRanges, applyMojikumiToRanges, clearRangeAttributes, applyPresetToRanges,
        readLeadingPercent, readLeadingTypeId, readJustifyId, readKinsokuId, readMojikumiId, readPresetFromRanges,
        dispatch
    ];
    var WORKER_LIB_SRC = "";
    for (var workerFuncIndex = 0; workerFuncIndex < WORKER_FUNCS.length; workerFuncIndex++) {
        WORKER_LIB_SRC += WORKER_FUNCS[workerFuncIndex].toString() + "\n";
    }

    /**
     * 委譲する引数をソース文字列へ変換する
     * @param {Object} params - 委譲したい引数（preset / clear / psList）
     * @returns {string} オブジェクトリテラルのソース
     */
    function paramsToSource(params) {
        if (!params) return "{}";
        var parts = [];
        if (params.preset) parts.push("preset:" + presetToJsonText(params.preset));
        if (params.clear) parts.push("clear:" + clearOptionsToSource(params.clear));
        if (params.psList !== undefined) parts.push("psList:" + jsonQuote(params.psList));
        return "{" + parts.join(",") + "}";
    }

    /**
     * ［クリア］のチェック状態をソース文字列へ変換する
     * @param {Object} clearOptions - 項目名をキーにした真偽値
     * @returns {string} オブジェクトリテラルのソース
     */
    function clearOptionsToSource(clearOptions) {
        var parts = [];
        for (var i = 0; i < CLEAR_OPTION_DEFINITIONS.length; i++) {
            var key = CLEAR_OPTION_DEFINITIONS[i].key;
            parts.push(key + ":" + (clearOptions[key] ? "true" : "false"));
        }
        return "{" + parts.join(",") + "}";
    }

    /**
     * メインエンジンへ処理を委譲する（非同期）
     * BridgeTalk は本文のバックスラッシュをエスケープするため、コード全体を
     * encodeURIComponent で包んで送り、ターゲットで復元する
     * @param {string} actionId - dispatch に渡すアクション名
     * @param {Object} params - dispatch に渡す引数
     * @param {Function} onDone - 成功時に payload を受け取るコールバック
     * @returns {void}
     */
    function runWorker(actionId, params, onDone) {
        var code = WORKER_LIB_SRC + '\nvar __r=dispatch("' + actionId + '",' + paramsToSource(params) + ");__r;";

        var bridge = new BridgeTalk();
        bridge.target = "illustrator";
        bridge.body = "eval(decodeURIComponent(\"" + encodeURIComponent(code) + "\"));";
        bridge.onResult = function (response) {
            var payload = response.body || "";
            var colonIndex = payload.indexOf(":");
            if (payload.substring(0, colonIndex) === "OK") {
                if (onDone) onDone(payload.substring(colonIndex + 1));
                return;
            }
            /* 失敗は握りつぶさず可視化。1行目を文、2行目に内訳 / Surface failures: a sentence, then the details */
            alert(getLabel(LABELS.alert.applyFailed) + "\n[" + actionId + "] " + payload);
        };
        bridge.onError = function (response) {
            alert(getLabel(LABELS.alert.applyFailed) + "\n[" + actionId + "] " + (response && response.body ? response.body : "BridgeTalk error"));
        };
        bridge.send();
    }

    /**
     * 文字組みアキ量設定の内部 ID を doc.mojikumiSet のインデックスへ変換する
     * @param {string} mojikumiId - 段落属性が返す内部 ID（"None" は「なし」）
     * @returns {number} mojikumiSet のインデックス（-1＝なし）。対応が無ければ NaN
     */
    function mojikumiIndexOfId(mojikumiId) {
        for (var i = 0; i < MOJIKUMI_OPTIONS.length; i++) {
            if (MOJIKUMI_OPTIONS[i].id === mojikumiId) return MOJIKUMI_OPTIONS[i].index;
        }
        return NaN;
    }

    /**
     * readPreset の payload をプリセットへ復元する
     * @param {string} payload - dispatch が返したタブ区切りの payload
     * @returns {Object|null} プリセット。読み取れていない場合は null
     */
    function parsePresetPayload(payload) {
        if (!payload) return null;
        var fields = decodeURIComponent(payload).split(String.fromCharCode(9));
        if (fields.length < 11 || fields[0] === "") return null;
        var leadingPercent = parseFloat(fields[2]);
        var restored = {
            psName: fields[0],
            size: parseFloat(fields[1]),
            /* 読み取れなければ既定の自動行送り量に寄せる / Fall back to the default auto-leading amount */
            leadingPercent: isNaN(leadingPercent) ? DEFAULT_AUTO_LEADING_PERCENT : leadingPercent,
            leadingType: fields[3] || "top",
            kern: fields[4],
            tsume: parseInt(fields[5], 10),
            tracking: parseInt(fields[6], 10),
            align: fields[7],
            justify: fields[8],
            kinsoku: fields[9] || "None"
        };
        /* 文字組みは内部 ID で返るので、適用に使うインデックスへ直す / Mojikumi comes back as an internal id; map it to the index used to apply it */
        var mojikumiIndex = mojikumiIndexOfId(fields[10]);
        if (!isNaN(mojikumiIndex)) restored.mojikumi = mojikumiIndex;
        return restored;
    }

    // =========================================
    // プリセットの保存・読み込み（JSON）/ Preset persistence (JSON)
    // ExtendScript には JSON が無いため、保存は手動シリアライズ、
    // 読み込みは自前ファイルなので eval で復元する（ES3 安全）。
    // ExtendScript has no JSON, so we serialize by hand and restore our own
    // file via eval (ES3-safe).
    // =========================================

    /**
     * プリセットの保存先ファイルを返す
     * @returns {File} プリセット保存用の JSON ファイル
     */
    function getPresetFile() {
        return File(Folder.userData.fsName + "/" + PRESET_FILE_NAME);
    }

    /**
     * 文字列を JSON の文字列リテラルへ変換する
     * @param {string} str - 変換したい文字列
     * @returns {string} 引用符とエスケープを施した文字列
     */
    function jsonQuote(str) {
        var escaped = String(str);
        escaped = escaped.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
        escaped = escaped.replace(/\r/g, "\\r").replace(/\n/g, "\\n").replace(/\t/g, "\\t");
        return '"' + escaped + '"';
    }

    /**
     * 文字列のフィールドを JSON へ追加する（未設定・空文字は書かない）
     * @param {Array<string>} fields - 追加先のフィールド配列
     * @param {string} key - JSON のキー
     * @param {string} value - 書き出したい値
     * @returns {void}
     */
    function pushStringField(fields, key, value) {
        if (value === undefined || value === null || value === "") return;
        fields.push(jsonQuote(key) + ":" + jsonQuote(value));
    }

    /**
     * 数値のフィールドを JSON へ追加する（未設定・NaN は書かない）
     * @param {Array<string>} fields - 追加先のフィールド配列
     * @param {string} key - JSON のキー
     * @param {number} value - 書き出したい値
     * @returns {void}
     */
    function pushNumberField(fields, key, value) {
        if (value === undefined || value === null || isNaN(value)) return;
        fields.push(jsonQuote(key) + ":" + String(parseFloat(value)));
    }

    /**
     * プリセット1件を JSON テキスト（1行）へ変換する
     * 持っていないプロパティは書かない（「適用する設定」で絞ったプリセットもそのまま通せる）
     * @param {Object} preset - 変換したいプリセット
     * @returns {string} JSON テキスト
     */
    function presetToJsonText(preset) {
        var fields = [];
        pushStringField(fields, "psName", preset.psName);
        pushNumberField(fields, "size", preset.size);
        pushNumberField(fields, "leadingPercent", preset.leadingPercent);
        pushStringField(fields, "leadingType", preset.leadingType);
        pushStringField(fields, "kern", preset.kern);
        pushNumberField(fields, "tsume", preset.tsume);
        pushNumberField(fields, "tracking", preset.tracking);
        pushStringField(fields, "align", preset.align);
        pushStringField(fields, "justify", preset.justify);
        pushStringField(fields, "kinsoku", preset.kinsoku);
        pushNumberField(fields, "mojikumi", preset.mojikumi);
        return "{" + fields.join(", ") + "}";
    }

    /**
     * プリセット配列を JSON テキストへ変換する
     * @param {Array<Object>} presets - 変換したいプリセット配列
     * @returns {string} JSON テキスト
     */
    function presetsToJsonText(presets) {
        var jsonLines = [];
        for (var i = 0; i < presets.length; i++) {
            jsonLines.push("  " + presetToJsonText(presets[i]));
        }
        return "[\n" + jsonLines.join(",\n") + "\n]\n";
    }

    /**
     * プリセットを保存先ファイルへ保存する
     * @param {Array<Object>} presets - 保存するプリセット配列
     * @returns {boolean} 保存できたら true
     */
    function savePresets(presets) {
        var file = getPresetFile();
        try {
            file.encoding = "UTF-8";
            if (!file.open("w")) return false;
            file.write(presetsToJsonText(presets));
            file.close();
            return true;
        } catch (e) {
            try { file.close(); } catch (eClose) { }
            return false;
        }
    }

    /**
     * 既定プリセットの複製を返す（元の配列を書き換えないように）
     * @returns {Array<Object>} 既定プリセットの複製
     */
    function defaultPresets() {
        return DEFAULT_PRESETS.slice(0);
    }

    /**
     * パース結果がプリセット配列として妥当かを判定する
     * 各要素が psName（空でない文字列）を持つ配列であることを要求する
     * すべて削除した結果の空配列は妥当として扱う（破損ではない）
     * @param {Object} parsed - eval で復元した値
     * @returns {boolean} プリセット配列として使えるなら true
     */
    function isValidPresetArray(parsed) {
        if (!parsed || typeof parsed.length !== "number") return false;
        for (var i = 0; i < parsed.length; i++) {
            var entry = parsed[i];
            if (!entry || typeof entry.psName !== "string" || entry.psName === "") return false;
        }
        return true;
    }

    /**
     * 壊れた設定ファイルを .bak に退避する（上書きせず原因を追えるように）
     * @param {File} file - 退避したいファイル
     * @returns {void}
     */
    function backupCorruptPresetFile(file) {
        try {
            var backup = new File(file.fsName + ".bak");
            if (backup.exists) backup.remove();
            file.copy(backup);
        } catch (e) { }
    }

    /**
     * 旧形式（leading / autoLeading）の保存ファイルを行送り％へ寄せる
     * 行送りは常に「自動」で扱うため、pt 値はフォントサイズとの比から % に直す
     * @param {Array<Object>} presets - 読み込んだプリセット配列
     * @returns {Array<Object>} 行送り％を持たせたプリセット配列
     */
    function migrateLeadingPercent(presets) {
        for (var i = 0; i < presets.length; i++) {
            var preset = presets[i];
            if (preset.leadingPercent !== undefined && preset.leadingPercent !== null) continue;
            if (preset.autoLeading !== true && preset.leading > 0 && preset.size > 0) {
                preset.leadingPercent = roundToHundredths(preset.leading / preset.size * 100);
            } else {
                preset.leadingPercent = DEFAULT_AUTO_LEADING_PERCENT;
            }
        }
        return presets;
    }

    /**
     * プリセットをファイルから読み込む（無ければ既定値）
     * 復旧方針: (1) 読み込み失敗 → 既定 / (2) 空ファイル → 既定 /
     * (3) パース失敗・内容不正 → 破損ファイルを .bak へ退避してから既定
     * @returns {Array<Object>} プリセット配列
     */
    function loadPresets() {
        var file = getPresetFile();
        if (!file.exists) return defaultPresets();

        /* 読み込み / Read */
        var fileText = "";
        try {
            file.encoding = "UTF-8";
            if (!file.open("r")) return defaultPresets();
            fileText = file.read();
            file.close();
        } catch (eRead) {
            try { file.close(); } catch (eClose) { }
            return defaultPresets();
        }

        /* 空ファイルは既定へ（破損扱いにしない）/ Empty file → defaults (not treated as corrupt) */
        if (!fileText || fileText.replace(/^\s+|\s+$/g, "") === "") return defaultPresets();

        /* パース（自前ファイルなので eval を許容）/ Parse (eval is acceptable for our own file) */
        var parsed = null;
        try {
            parsed = eval("(" + fileText + ")");
        } catch (eParse) {
            parsed = null;
        }

        /* パース失敗・内容不正は破損として .bak 退避のうえ既定へ / Parse failure or invalid shape → back up, then defaults */
        if (!isValidPresetArray(parsed)) {
            backupCorruptPresetFile(file);
            return defaultPresets();
        }
        return migrateLeadingPercent(parsed);
    }

    // =========================================
    // パレット位置の保存・復元 / Palette position persistence
    // 移動のたびの位置を userData に "x,y" で保存し、次回起動時に同じ位置で開く。
    // The location is stored as "x,y" so the palette reopens where it was left.
    // =========================================

    /**
     * パレット位置の保存先ファイルを返す
     * @returns {File} 位置保存用のテキストファイル
     */
    function getPaletteLocationFile() {
        return File(Folder.userData.fsName + "/" + LOCATION_FILE_NAME);
    }

    /**
     * 指定した位置がいずれかのモニタ内かを判定する（構成変更で画面外に復元しないように）
     * @param {number} x - 復元したい横位置
     * @param {number} y - 復元したい縦位置
     * @returns {boolean} 画面内なら true
     */
    function isLocationOnScreen(x, y) {
        var screens;
        try { screens = $.screens; } catch (e) { return true; }
        if (!screens || !screens.length) return true;
        var margin = 40; /* タイトルバーを掴める余白 / Keep the title bar reachable */
        for (var i = 0; i < screens.length; i++) {
            var screen = screens[i];
            if (x >= screen.left - margin && x <= screen.right - margin &&
                y >= screen.top && y <= screen.bottom - margin) return true;
        }
        return false;
    }

    /**
     * 現在のパレット位置を保存する
     * @param {Window} paletteWindow - 対象のパレット
     * @returns {void}
     */
    function savePaletteLocation(paletteWindow) {
        var file = getPaletteLocationFile();
        file.encoding = "UTF-8";
        if (!file.open("w")) return;
        file.write(Math.round(paletteWindow.location[0]) + "," + Math.round(paletteWindow.location[1]));
        file.close();
    }

    /**
     * 保存しておいたパレット位置を適用する
     * @param {Window} paletteWindow - 対象のパレット
     * @returns {void}
     */
    function applySavedPaletteLocation(paletteWindow) {
        var file = getPaletteLocationFile();
        if (!file.exists) return;
        file.encoding = "UTF-8";
        if (!file.open("r")) return;
        var fields = file.read().split(",");
        file.close();
        var x = parseInt(fields[0], 10), y = parseInt(fields[1], 10);
        if (isNaN(x) || isNaN(y) || !isLocationOnScreen(x, y)) return;
        paletteWindow.location = [x, y];
    }

    // =========================================
    // 一覧の表示 / Populating the list
    // =========================================

    /* psName → 表示名（ファミリー＋スタイル）。メインエンジンで解決した結果をためる
       psName → display name (family + style), cached from the main engine */
    var fontDisplayNames = {};

    /**
     * PostScript 名の表示名を返す（未解決なら PostScript 名のまま）
     * @param {string} psName - フォントの PostScript 名
     * @returns {string} 表示名
     */
    function fontDisplayName(psName) {
        return fontDisplayNames[psName] ? fontDisplayNames[psName] : psName;
    }

    /**
     * ID に対応する表示ラベルを返す
     * @param {Object} labelGroup - ID をキーに持つラベル定義のまとまり
     * @param {string} id - 引きたい ID
     * @returns {string} 表示ラベル。未設定・判定不能は "—"
     */
    function optionLabel(labelGroup, id) {
        var labelSet = id ? labelGroup[id] : null;
        return labelSet ? getLabel(labelSet) : "—";
    }

    /**
     * 禁則 ID を表示ラベルへ変換する
     * @param {string} kinsokuId - "None" / "Hard" / "Soft" / "Soft_v2"
     * @returns {string} 表示ラベル。未設定・判定不能は "—"
     */
    function kinsokuLabel(kinsokuId) {
        for (var i = 0; i < KINSOKU_OPTIONS.length; i++) {
            if (KINSOKU_OPTIONS[i].id === kinsokuId) return getLabel(KINSOKU_OPTIONS[i].label);
        }
        return "—";
    }

    /**
     * 文字組みアキ量設定のインデックスを表示ラベルへ変換する
     * @param {number} mojikumiIndex - -1＝なし、0以上は mojikumiSet のインデックス
     * @returns {string} 表示ラベル。未設定・判定不能は "—"
     */
    function mojikumiLabel(mojikumiIndex) {
        for (var i = 0; i < MOJIKUMI_OPTIONS.length; i++) {
            if (MOJIKUMI_OPTIONS[i].index === mojikumiIndex) return getLabel(MOJIKUMI_OPTIONS[i].label);
        }
        return "—";
    }

    /**
     * 数値をそのまま文字列にする（未設定は空文字）
     * @param {number} value - 変換したい値
     * @returns {string} 表示用の文字列
     */
    function formatNumber(value) {
        return (value !== undefined && value !== null) ? String(value) : "";
    }

    /**
     * フォントサイズの表示文字列を返す（pt 付き）
     * @param {Object} preset - 表示するプリセット
     * @returns {string} 表示用の文字列
     */
    function formatFontSize(preset) {
        if (preset.size === undefined || preset.size === null) return "";
        return String(preset.size) + " pt";
    }

    /**
     * 実質の行送りの表示文字列を返す（フォントサイズ×自動行送り量）
     * @param {Object} preset - 表示するプリセット
     * @returns {string} 表示用の文字列
     */
    function formatLeading(preset) {
        if (preset.size === undefined || preset.size === null) return "";
        if (preset.leadingPercent === undefined || preset.leadingPercent === null) return "";
        return String(roundToHundredths(preset.size * preset.leadingPercent / 100)) + " pt";
    }

    /**
     * 自動行送り量の表示文字列を返す（% 付き）
     * @param {Object} preset - 表示するプリセット
     * @returns {string} 表示用の文字列
     */
    function formatLeadingPercent(preset) {
        if (preset.leadingPercent === undefined || preset.leadingPercent === null) return "";
        return String(preset.leadingPercent) + "%";
    }

    /**
     * 上段のフォント一覧へプリセットを流し込む
     * @param {ListBox} fontList - フォント一覧の listbox
     * @param {Array<Object>} presets - 表示するプリセット配列
     * @returns {void}
     */
    function fillFontList(fontList, presets) {
        fontList.removeAll();
        for (var i = 0; i < presets.length; i++) {
            fontList.add("item", fontDisplayName(presets[i].psName));
        }
    }

    /**
     * 詳細一覧へ「項目名／値」の1行を追加する
     * 「適用する設定」で外したグループの行はディム表示にする
     * @param {ListBox} detailList - 詳細一覧の listbox
     * @param {Object} labelSet - 項目名のラベル定義
     * @param {string} value - 表示する値
     * @param {boolean} isApplied - 適用対象なら true。false ならディム表示
     * @returns {void}
     */
    function addDetailRow(detailList, labelSet, value, isApplied) {
        var row = detailList.add("item", labelText(labelSet));
        row.subItems[0].text = value;
        row.enabled = isApplied ? true : false;
    }

    /**
     * 下段の詳細一覧へ、選択中のプリセットの内訳を並べる（左＝項目名／右＝値）
     * 「適用する設定」で外したグループの行はディム表示にする
     * プロポーショナルは自動カーニングに連動する値なので、kern から導いて表示する
     * 行数を変えたら DETAIL_ROW_COUNT も合わせる
     * @param {ListBox} detailList - 詳細一覧の listbox
     * @param {Object} preset - 表示するプリセット（未選択なら null）
     * @param {Object} appliedGroups - 適用する設定のチェック状態（size / leading / spacing / japanese）
     * @returns {void}
     */
    function fillDetailList(detailList, preset, appliedGroups) {
        detailList.removeAll();
        if (!preset) return;
        addDetailRow(detailList, LABELS.field.size, formatFontSize(preset), appliedGroups.size);
        addDetailRow(detailList, LABELS.field.leading, formatLeading(preset), appliedGroups.leading);
        addDetailRow(detailList, LABELS.field.leadingPercent, formatLeadingPercent(preset), appliedGroups.leading);
        addDetailRow(detailList, LABELS.field.leadingType, optionLabel(LABELS.leadingType, preset.leadingType), appliedGroups.leading);
        addDetailRow(detailList, LABELS.field.kern, optionLabel(LABELS.autoKern, preset.kern), appliedGroups.spacing);
        addDetailRow(detailList, LABELS.field.tsume, formatNumber(preset.tsume), appliedGroups.spacing);
        addDetailRow(detailList, LABELS.field.propMetrics, (preset.kern === "metrics") ? "ON" : "OFF", appliedGroups.spacing);
        addDetailRow(detailList, LABELS.field.tracking, formatNumber(preset.tracking), appliedGroups.spacing);
        addDetailRow(detailList, LABELS.field.align, optionLabel(LABELS.align, preset.align), appliedGroups.japanese);
        addDetailRow(detailList, LABELS.field.justify, optionLabel(LABELS.justify, preset.justify), appliedGroups.japanese);
        addDetailRow(detailList, LABELS.field.kinsoku, kinsokuLabel(preset.kinsoku), appliedGroups.japanese);
        addDetailRow(detailList, LABELS.field.mojikumi, mojikumiLabel(preset.mojikumi), appliedGroups.japanese);
    }

    // =========================================
    // パレット / Palette
    // =========================================

    /* ［適用する設定］のチェックボックス定義。並び順＝表示順で、row は何行目に置くか（0 始まり）
       The "settings to apply" checkboxes; array order is display order and row is the line to put it on */
    var APPLY_GROUP_DEFINITIONS = [
        { key: "font", row: 0, label: LABELS.applyGroup.font, tip: LABELS.tip.applyFont },
        { key: "size", row: 0, label: LABELS.applyGroup.size, tip: LABELS.tip.applySize },
        { key: "leading", row: 0, label: LABELS.applyGroup.leading, tip: LABELS.tip.applyLeading },
        { key: "spacing", row: 1, label: LABELS.applyGroup.spacing, tip: LABELS.tip.applySpacing },
        { key: "japanese", row: 2, label: LABELS.applyGroup.japanese, tip: LABELS.tip.applyJapanese }
    ];

    /* ［クリア］のチェックボックス定義。定番には持たせず、適用のたびに標準値へ戻す項目
       The "clear" checkboxes; these are not stored in a preset, they are reset on every apply */
    var CLEAR_OPTION_DEFINITIONS = [
        { key: "scale", row: 0, label: LABELS.clearGroup.scale, tip: LABELS.tip.clearScale },
        { key: "aki", row: 0, label: LABELS.clearGroup.aki, tip: LABELS.tip.clearAki },
        { key: "baselineShift", row: 1, label: LABELS.clearGroup.baselineShift, tip: LABELS.tip.clearBaselineShift }
    ];

    /**
     * チェックボックスを並べたパネルを組み立てる（［適用する設定］と［クリア］で共用）
     * @param {Window} paletteWindow - 追加先のパレット
     * @param {Object} titleLabel - パネル見出しのラベル定義
     * @param {Object} panelTip - パネルのツールチップのラベル定義
     * @param {Array<Object>} definitions - チェックボックス定義（key / row / label / tip）
     * @returns {Object} byGroup（キー別のチェックボックス）と all（並び順の配列）
     */
    function buildCheckboxPanel(paletteWindow, titleLabel, panelTip, definitions) {
        var panel = paletteWindow.add("panel", undefined, getLabel(titleLabel));
        panel.helpTip = getLabel(panelTip);
        panel.orientation = "column";
        panel.alignChildren = ["left", "top"];
        panel.margins = PANEL_MARGINS;
        panel.spacing = PANEL_ROW_SPACING;

        /* 行は定義の順に必要な数だけ作る / Rows are created in definition order, as many as are asked for */
        var rowGroups = [], byGroup = {}, allChecks = [];
        for (var i = 0; i < definitions.length; i++) {
            var definition = definitions[i];
            if (!rowGroups[definition.row]) {
                rowGroups[definition.row] = panel.add("group");
                setupRow(rowGroups[definition.row], "left");
            }
            var checkbox = rowGroups[definition.row].add("checkbox", undefined, getLabel(definition.label));
            checkbox.helpTip = getLabel(definition.tip);
            checkbox.value = true; /* 既定はすべてオン / Everything is on by default */
            byGroup[definition.key] = checkbox;
            allChecks.push(checkbox);
        }
        return { byGroup: byGroup, all: allChecks };
    }

    /**
     * 上段のフォント一覧と下段の詳細一覧を組み立てる
     * @param {Window} paletteWindow - 追加先のパレット
     * @returns {Object} fontList（上段）と detailList（下段）
     */
    function buildPresetLists(paletteWindow) {
        var fontList = paletteWindow.add("listbox", undefined, [], { multiselect: false });
        fontList.preferredSize = FONT_LIST_SIZE;
        fontList.helpTip = getLabel(LABELS.tip.presets);

        var detailList = paletteWindow.add("listbox", undefined, [], {
            multiselect: false,
            numberOfColumns: 2,
            columnWidths: DETAIL_COLUMN_WIDTHS
        });
        detailList.preferredSize = DETAIL_LIST_SIZE;
        detailList.helpTip = getLabel(LABELS.tip.detail);

        return { fontList: fontList, detailList: detailList };
    }

    /**
     * プリセット操作のボタン行を組み立てる
     * @param {Window} paletteWindow - 追加先のパレット
     * @returns {Object} remove / overwrite / add のボタン
     */
    function buildActionButtons(paletteWindow) {
        var btnRowGroup = paletteWindow.add("group");
        setupRow(btnRowGroup, "right");
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        var btnRemove = btnRowGroup.add("button", undefined, getLabel(LABELS.button.deletePreset));
        btnRemove.helpTip = getLabel(LABELS.tip.deletePreset);
        var btnOverwrite = btnRowGroup.add("button", undefined, getLabel(LABELS.button.overwritePreset));
        btnOverwrite.helpTip = getLabel(LABELS.tip.overwritePreset);
        var btnAdd = btnRowGroup.add("button", undefined, getLabel(LABELS.button.addPreset));
        btnAdd.helpTip = getLabel(LABELS.tip.addPreset);
        return { remove: btnRemove, overwrite: btnOverwrite, add: btnAdd };
    }

    /**
     * 定番フォントのパレットを組み立てる
     * @returns {Window} 組み立てたパレット
     */
    function createPalette() {
        var presets = loadPresets();

        var palette = new Window("palette", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupPalette(palette);

        /* 追加する順＝上からの並び順 / The order they are added is the order they stack */
        var applyChecks = buildCheckboxPanel(palette, LABELS.panel.applySettings, LABELS.tip.applyPanel, APPLY_GROUP_DEFINITIONS);
        var presetLists = buildPresetLists(palette);
        var fontList = presetLists.fontList;
        var detailList = presetLists.detailList;
        var clearChecks = buildCheckboxPanel(palette, LABELS.panel.clear, LABELS.tip.clearPanel, CLEAR_OPTION_DEFINITIONS);
        var actionButtons = buildActionButtons(palette);
        var btnDelete = actionButtons.remove;
        var btnOverwrite = actionButtons.overwrite;
        var btnAdd = actionButtons.add;

        /* 一覧を組み直す間・選択位置を揃える間は onChange（＝適用）を止める
           Suppress onChange (i.e. apply) while the list is rebuilt or the selection is mirrored */
        var suppressListEvents = false;

        /* 選択中のテキストの有無。onShow / onActivate のたびに取り直す
           Whether text is selected; re-read on every show and activation */
        var hasTextSelection = false;

        /**
         * 委譲の戻り値（テキスト範囲の数）を、現在値を取り込むボタンの可否へ反映する
         * @param {string} payload - dispatch が返したテキスト範囲の数
         * @returns {void}
         */
        function reflectTextSelectionCount(payload) {
            hasTextSelection = (parseInt(payload, 10) > 0);
            btnAdd.enabled = hasTextSelection;
            btnOverwrite.enabled = hasTextSelection;
        }

        /**
         * フォント一覧の選択位置を返す
         * @returns {number} 選択中のインデックス。未選択は -1
         */
        function selectedPresetIndex() {
            var selection = fontList.selection;
            return selection ? selection.index : -1;
        }

        /**
         * パネルのチェック状態を、定義のキー別の真偽値にして返す
         * @param {Object} checkboxPanel - buildCheckboxPanel が返したまとまり
         * @param {Array<Object>} definitions - 対応するチェックボックス定義
         * @returns {Object} キー別の真偽値
         */
        function checkedStates(checkboxPanel, definitions) {
            var states = {};
            for (var i = 0; i < definitions.length; i++) {
                states[definitions[i].key] = checkboxPanel.byGroup[definitions[i].key].value;
            }
            return states;
        }

        /**
         * ［適用する設定］のチェック状態をまとめて返す
         * @returns {Object} グループ名をキーにした真偽値
         */
        function appliedPropertyGroups() {
            return checkedStates(applyChecks, APPLY_GROUP_DEFINITIONS);
        }

        /**
         * ［クリア］のチェック状態をまとめて返す
         * @returns {Object} 項目名をキーにした真偽値
         */
        function clearedAttributes() {
            return checkedStates(clearChecks, CLEAR_OPTION_DEFINITIONS);
        }

        /**
         * フォント一覧の選択位置を決め、下段の詳細一覧をその内訳に差し替える
         * @param {number} index - 選択したいインデックス（範囲外なら選択解除）
         * @returns {void}
         */
        function selectPreset(index) {
            var target = (index >= 0 && index < presets.length) ? index : null;
            suppressListEvents = true;
            fontList.selection = target;
            suppressListEvents = false;
            fillDetailList(detailList, (target === null) ? null : presets[target], appliedPropertyGroups());
        }

        /**
         * クリックした項目だけをオンにする。すでにその状態ならすべてオンへ戻す
         * @param {Array<Checkbox>} checkboxes - 同じパネルのチェックボックス一覧
         * @param {Checkbox} clickedCheck - 基準にするチェックボックス
         * @returns {void}
         */
        function soloCheckbox(checkboxes, clickedCheck) {
            /* onClick はチェックが切り替わったあとに来るので、
               「押した項目以外がすべてオフ」ならソロ状態だったとみなす
               onClick fires after the box has toggled, so "every other one is off" means it was soloed */
            var wasSoloed = true;
            for (var i = 0; i < checkboxes.length; i++) {
                if (checkboxes[i] !== clickedCheck && checkboxes[i].value) { wasSoloed = false; break; }
            }
            for (var j = 0; j < checkboxes.length; j++) {
                checkboxes[j].value = wasSoloed ? true : (checkboxes[j] === clickedCheck);
            }
        }

        /**
         * チェックボックス1つにクリック時の処理を割り当てる
         * option（alt）併用のときだけ、その項目だけオンと、すべてオンを切り替える
         * @param {Array<Checkbox>} checkboxes - 同じパネルのチェックボックス一覧
         * @param {Checkbox} targetCheck - 対象のチェックボックス
         * @param {Function} onChanged - 切り替えたあとに呼ぶ処理（省略可）
         * @returns {void}
         */
        function setupSoloCheckbox(checkboxes, targetCheck, onChanged) {
            targetCheck.onClick = function () {
                if (ScriptUI.environment.keyboardState.altKey) soloCheckbox(checkboxes, targetCheck);
                if (onChanged) onChanged();
            };
        }

        /**
         * 「適用する設定」を変えたら、詳細一覧のディム表示を合わせる
         * @returns {void}
         */
        function refreshDetailDimming() {
            var index = selectedPresetIndex();
            fillDetailList(detailList, (index >= 0) ? presets[index] : null, appliedPropertyGroups());
        }

        /**
         * 未解決のフォント名をメインエンジンで解決し、一覧のフォント名だけを差し替える
         * @returns {void}
         */
        function refreshFontLabels() {
            var pending = [];
            for (var i = 0; i < presets.length; i++) {
                if (!fontDisplayNames[presets[i].psName]) pending.push(presets[i].psName);
            }
            if (!pending.length) return;
            runWorker("resolveFontNames", { psList: pending.join(String.fromCharCode(10)) }, function (payload) {
                var rows = decodeURIComponent(payload).split(String.fromCharCode(10));
                for (var rowIndex = 0; rowIndex < rows.length; rowIndex++) {
                    var tabIndex = rows[rowIndex].indexOf(String.fromCharCode(9));
                    if (tabIndex >= 0) fontDisplayNames[rows[rowIndex].substring(0, tabIndex)] = rows[rowIndex].substring(tabIndex + 1);
                }
                /* 行は作り直さず、フォント名だけ差し替える / Patch the font names instead of rebuilding rows */
                var fontItems = fontList.items;
                for (var presetIndex = 0; presetIndex < presets.length; presetIndex++) {
                    if (fontItems[presetIndex]) fontItems[presetIndex].text = fontDisplayName(presets[presetIndex].psName);
                }
            });
        }

        /**
         * 選択中のテキストの有無を取り直し、ボタンへ反映する
         * @returns {void}
         */
        function refreshTextSelectionState() {
            runWorker("getState", null, reflectTextSelectionCount);
        }

        /**
         * プリセットを保存し、一覧を組み直して選択位置を戻す
         * @param {number} index - 保存後に選択しておきたいインデックス
         * @returns {void}
         */
        function commitPresets(index) {
            savePresets(presets);
            suppressListEvents = true;
            fillFontList(fontList, presets);
            suppressListEvents = false;
            selectPreset(index);
            refreshFontLabels();
        }

        /**
         * 選択中のテキストの現在値を読み取り、コールバックへ渡す
         * @param {Function} onRead - 読み取れたプリセットを受け取るコールバック
         * @returns {void}
         */
        function readCurrentPreset(onRead) {
            runWorker("readPreset", null, function (payload) {
                var currentPreset = parsePresetPayload(payload);
                if (!currentPreset) { alert(getLabel(LABELS.alert.readFailed)); return; }
                onRead(currentPreset);
            });
        }

        /**
         * 「適用する設定」でチェックされている項目だけを残したプリセットを返す
         * 残さなかったプロパティは適用時に飛ばされ、適用先の設定はそのまま残る
         * @param {Object} preset - もとのプリセット
         * @returns {Object} 適用用に絞り込んだプリセット
         */
        function presetForApply(preset) {
            var groups = appliedPropertyGroups();
            var picked = {};
            if (groups.font) picked.psName = preset.psName;
            if (groups.size) picked.size = preset.size;
            if (groups.leading) {
                picked.leadingPercent = preset.leadingPercent;
                picked.leadingType = preset.leadingType;
            }
            if (groups.spacing) {
                picked.kern = preset.kern;
                picked.tsume = preset.tsume;
                picked.tracking = preset.tracking;
            }
            if (groups.japanese) {
                picked.align = preset.align;
                picked.justify = preset.justify;
                picked.kinsoku = preset.kinsoku;
                picked.mojikumi = preset.mojikumi;
            }
            return picked;
        }

        /**
         * 一覧のクリックで選択位置と詳細一覧を反映し、プリセットを一括適用する
         * @param {ListBox} list - クリックされた listbox
         * @returns {void}
         */
        function onPresetClick() {
            if (suppressListEvents || !fontList.selection) return;
            var index = fontList.selection.index;
            selectPreset(index);
            /* 適用の戻り値が選択できたテキスト範囲の数なので、状態の更新も兼ねる
               The apply result is how many text ranges were reached, so it doubles as a state refresh */
            runWorker("applyPreset", {
                preset: presetForApply(presets[index]),
                clear: clearedAttributes()
            }, reflectTextSelectionCount);
        }

        /* フォント一覧のクリックで内訳を出し、そのまま一括適用 / A click shows the settings and applies them */
        fontList.onChange = onPresetClick;

        /* 適用しないグループは詳細一覧でもディム表示にする / Groups left out are dimmed in the detail list too */
        for (var applyIndex = 0; applyIndex < applyChecks.all.length; applyIndex++) {
            setupSoloCheckbox(applyChecks.all, applyChecks.all[applyIndex], refreshDetailDimming);
        }
        for (var clearIndex = 0; clearIndex < clearChecks.all.length; clearIndex++) {
            setupSoloCheckbox(clearChecks.all, clearChecks.all[clearIndex]);
        }

        /* 追加：選択中のテキストの設定を登録（同じフォントは上書き）/ Add: capture the selection (same font overwrites) */
        btnAdd.onClick = function () {
            readCurrentPreset(function (newPreset) {
                var index = -1;
                for (var i = 0; i < presets.length; i++) {
                    if (presets[i].psName === newPreset.psName) { index = i; break; }
                }
                if (index >= 0) presets[index] = newPreset;
                else index = presets.push(newPreset) - 1;
                commitPresets(index);
            });
        };

        /* 上書き：一覧で選択中の項目を現在の設定で更新 / Overwrite: update the selected item with the current settings */
        btnOverwrite.onClick = function () {
            var index = selectedPresetIndex();
            if (index < 0) return;
            readCurrentPreset(function (newPreset) {
                presets[index] = newPreset;
                commitPresets(index);
            });
        };

        /* 削除：一覧で選択中の項目を削除 / Delete: remove the selected item */
        btnDelete.onClick = function () {
            var index = selectedPresetIndex();
            if (index < 0) return;
            presets.splice(index, 1);
            commitPresets(index < presets.length ? index : presets.length - 1);
        };

        fillFontList(fontList, presets);

        /* 表示・フォーカス復帰のたびに選択状況とフォント名を反映 / Reflect the selection and font names on show and on regaining focus */
        palette.onShow = function () { refreshFontLabels(); refreshTextSelectionState(); };
        palette.onActivate = refreshTextSelectionState;

        /* Esc でパレットを閉じる / Close on Esc */
        palette.addEventListener("keydown", function (event) {
            if (event.keyName === "Escape") palette.close();
        });

        return palette;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 常駐パレットを開く（二重起動は既存のパレットを前面化）
     * @returns {void}
     */
    function main() {
        /* 二重起動ガード：既存パレットがあれば前面化 / Single-instance guard */
        if ($.global.__FontPresetPicker && $.global.__FontPresetPicker.window) {
            try {
                var existingPalette = $.global.__FontPresetPicker.window;
                existingPalette.show();
                existingPalette.active = true;
                return;
            } catch (existingError) {
                /* 参照が死んでいたら作り直す / Stale reference, rebuild */
                $.global.__FontPresetPicker = null;
            }
        }

        var palette = createPalette();
        $.global.__FontPresetPicker = { window: palette };

        /* 閉じた位置がそのまま次回の初期位置になる / Wherever it was closed becomes next launch's spot */
        palette.onMove = function () { savePaletteLocation(palette); };
        palette.addEventListener("close", function () {
            savePaletteLocation(palette);
            $.global.__FontPresetPicker = null;
        });

        applySavedPaletteLocation(palette);
        palette.show();
        /* show 後にレイアウトが確定するため当て直す / Re-apply once the layout is final */
        applySavedPaletteLocation(palette);
    }

    main();

})();
