#targetengine "UnifiedTypePanelEngine"
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストの文字組み設定（フォント・フォントサイズ・自動カーニング・字間・文字揃え・行揃え・行送り・
文字組みアキ量設定）をまとめて行う常駐パレットスクリプトです。3カラム構成（左：ドキュメントフォント／プリセット、
中央：フォントサイズ・自動カーニング・字間調整・文字揃え、右：行揃え・行送り・行送りの基準・文字組みアキ量設定・禁則）。

- 最上部に種別（本文／見出し）と方針パネルを左右に配置

- ドキュメントフォント：書類で使用中のフォントを一覧表示。選ぶと選択テキストへ適用
- プリセット：フォント＋カーニング・文字ツメ・トラッキングをまとめて登録。選ぶと一括適用。
  「追加」で現在の選択の設定を保存、「上書き」で選択中のプリセットを更新、「削除」で除去（JSON で Folder.userData に永続化）
- フォントサイズ：サイズ・比率（水平／垂直を同値）・実質（サイズ×比率の表示）
- 自動カーニング：和文等幅／0／メトリクス／オプティカル（「メトリクス」のみプロポーショナルメトリクスON）
- 字間調整：最上部にプロポーショナルメトリクスのチェックボックス（自動カーニングを「メトリクス」にするとON、
  それ以外でOFFに連動）。その下に文字ツメ（0〜100%）とトラッキング（-100〜500）。入力欄とスライダー（Shift で粗い刻み）
- 文字揃え：欧文ベースライン／中央／その他（仮想ボディの上下・平均字面の上下をポップアップで選択）
- 行揃え：左／中央／右／均等配置（最終行左）／両端揃え（テキストの見た目の位置を保持して適用）
- 種別：本文（文字組みベタ組み）／見出し（文字組みツメ組み）。よく使う組み合わせを一括適用
- 行送り：フォントサイズ（サイズ）・実質（pt）・行送り（%）の3入力を1パネルに集約。
  % を段落の「自動行送り量」に設定し常に自動行送りにする（Illustrator 上は常に「自動」表示、行送りはサイズに自動追従）。
  実質＝サイズ×%。実質欄に値を入れると % を逆算して設定。行送りの基準は別パネル。
- 文字組みアキ量設定：なし／約物半角／約物全角／ツメ組み／ベタ組み などをポップアップで一括適用
- 禁則：なし／強い禁則／弱い禁則／弱い禁則 v2 をポップアップで一括適用（「なし」は scripting では設定不可）
- 制御文字の表示／非表示の切り替え、再読み込み（選択の現在値を読み直して反映）、
  リセット（標準値に戻す。カーニングはメトリクス・文字組みはツメ組み・字前/字後のアキは自動に設定）
- ラジオや入力を操作すると、その場で選択中のテキストへ即時適用する
- 選択は単体のテキストフレームだけでなく、グループ内のテキストやテキスト編集モードでの範囲選択にも対応（行送りを含む全機能で共通）
- パレットにフォーカスが戻るたび、または「再読み込み」で選択の現在値を読み取って UI に反映する
- 常駐エンジン（#targetengine）でパレット表示。常駐エンジンの app は
  パレット表示中に DOM 接続を失うため、DOM 処理はメインエンジンへ
  BridgeTalk で都度委譲する（コードは encodeURIComponent で包んで送信）

### 謝辞
　
古島佑起さん
BridgeTalk のワーカー登録と呼び出しの仕組み、行揃えのボタンの実装方法など、多くのアイデアとコードを提供いただきました。
https://note.com/yukifurushima/n/n9f2078dc156f

### 紹介記事（note）

https://note.com/dtp_tranist/n/n4e2b79cf2891

### Overview

A docked palette that sets text-composition attributes (font, font size, auto kerning,
letter spacing, character alignment, justification, leading, and mojikumi) for the selected text.
Three columns (left: document fonts / presets, center: font size, kerning, letter spacing
& alignment, right: justification, leading, basis, mojikumi, kinsoku).

- Top area: Type (Body / Heading) and Policy panels laid out side by side

- Document fonts: lists the fonts used in the document; clicking one applies it to the selection
- Presets: a font plus kerning / Tsume / tracking, applied together when clicked. "Add"
  saves the current selection's settings, "Overwrite" updates the selected preset, "Delete" removes it (persisted as JSON under Folder.userData)
- Font size: size, scale (horizontal/vertical set together), effective (size × scale, shown)
- Auto kerning: Metrics - Roman Only / 0 / Metrics / Optical (proportional metrics ON only for "Metrics")
- Letter spacing: a proportional-metrics checkbox at the top (ON when auto kerning is "Metrics", OFF otherwise),
  then Tsume (0–100%) and tracking (-100 to 500), via input fields and sliders (Shift = coarse steps)
- Character alignment: Roman baseline / center / Other (embox top-bottom & ICF box top-bottom via popup)
- Justification: left / center / right / justify (last left) / justify all (applied while keeping the text's visual position)
- Type: Body (solid mojikumi) / Heading (tight mojikumi); applies common combinations at once
- Leading: 115% / 150% / Other (enter a %), and leading basis.
  The chosen % is set as the paragraph's auto-leading amount and leading is always auto (Illustrator always
  shows "Auto"; the leading follows the font size). Choosing "Other" prefills the current auto-leading amount as %.
- Mojikumi: None / half-width punctuation / full-width punctuation / tight / solid, applied together via a popup
- Show/hide hidden characters, Reload (re-read the selection's current values), and
  Reset (restore defaults; kerning Metrics, mojikumi Tight, and aki before/after set to auto)
- Operating a radio or field applies it immediately to the current selection
- Selection handling covers not just standalone text frames but text nested in groups and ranges selected in text-edit mode (consistent across all features, including leading)
- Whenever the palette regains focus — or via "Reload" — the current selection's values are read back into the UI
- Runs as a persistent palette (#targetengine). The persistent engine's app
  loses its DOM connection while the palette is shown, so all DOM work is
  delegated to the main engine via BridgeTalk (code wrapped in encodeURIComponent)

*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.3.1";

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 言語判定 / Detect UI language */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    /* ラベル定義 / Label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "統合文字組みパネル", en: "Unified Type Panel" }
        },
        tab: {
            fonts: { ja: "定番", en: "Favorites" },
            type: { ja: "文字組み", en: "Type" },
            fontSize: { ja: "サイズ", en: "Size" },
            usedFonts: { ja: "使用フォント", en: "Used Fonts" },
            style: { ja: "スタイル", en: "Style" }
        },
        field: {
            docFonts: { ja: "ドキュメントフォント", en: "Document Fonts" },
            presets: { ja: "プリセット", en: "Presets" },
            fontSizePanel: { ja: "フォントサイズ", en: "Font Size" },
            fontSize: { ja: "サイズ", en: "Size" },
            scale: { ja: "比率", en: "Scale" },
            apparent: { ja: "実質", en: "Effective" },
            autoKern: { ja: "自動カーニング", en: "Auto Kerning" },
            tsume: { ja: "文字ツメ", en: "Tsume" },
            tracking: { ja: "トラッキング", en: "Tracking" },
            proportionalMetrics: { ja: "プロポーショナルメトリクス", en: "Proportional Metrics" },
            spacingAdjust: { ja: "字間調整", en: "Letter Spacing" },
            align: { ja: "文字揃え", en: "Char Alignment" },
            justify: { ja: "行揃え", en: "Justification" },
            role: { ja: "種別", en: "Type" },
            leading: { ja: "行送り", en: "Leading" },
            leadingPercent: { ja: "行送り%", en: "Leading (%)" },
            sizeAndLeading: { ja: "フォントサイズと行送り", en: "Font Size & Leading" },
            leadingType: { ja: "行送りの基準", en: "Leading basis" },
            mojikumi: { ja: "文字組みアキ量設定", en: "Mojikumi" },
            kinsoku: { ja: "禁則", en: "Kinsoku" },
            jpComposition: { ja: "日本語組版", en: "Japanese Typography" },
            filter: { ja: "フィルター", en: "Filter" },
            typeScale: { ja: "タイプスケール", en: "Type Scale" },
            typeScaleBase: { ja: "基準", en: "Base" },
            typeScaleRatio: { ja: "倍率", en: "Ratio" },
            typeScaleStep: { ja: "#", en: "#" }
        },
        policy: {
            wabun: { ja: "和文", en: "Japanese" },
            latin: { ja: "欧文", en: "Latin" },
            mixed: { ja: "和欧混在", en: "Mixed" }
        },
        // 使用フォント一覧（DocumentFontListSelector 由来）/ Used-fonts list labels
        usedFonts: {
            autoLeading: { ja: "自動", en: "Auto" },
            view: {
                fontOnly: { ja: "フォントのみ", en: "Font only" },
                detailed: { ja: "詳細", en: "Details" }
            },
            column: {
                count: { ja: "使用数", en: "Count" },
                font: { ja: "フォント", en: "Font" },
                size: { ja: "サイズ", en: "Size" },
                leading: { ja: "行送り", en: "Leading" },
                kern: { ja: "自動カーニング", en: "Auto Kerning" },
                tsume: { ja: "文字ツメ", en: "Tsume" },
                tracking: { ja: "トラッキング", en: "Tracking" },
                propMetrics: { ja: "プロポーショナル", en: "Prop. Metrics" }
            },
            control: {
                applyOnClick: { ja: "クリックで適用", en: "Apply on click" },
                currentArtboardOnly: { ja: "現在のアートボードに限定", en: "Current artboard only" },
                includeHidden: { ja: "非表示のテキストを含む", en: "Include hidden text" },
                includeLocked: { ja: "ロックされたテキストを含む", en: "Include locked text" }
            },
            button: {
                refresh: { ja: "リストを更新", en: "Refresh list" },
                selectMatching: { ja: "選択したフォントを選択", en: "Select matching text" }
            },
            status: {
                needSelection: { ja: "適用先のテキストを選択してください。", en: "Select the text to apply to." },
                needCondition: { ja: "一覧から条件を選択してください。", en: "Select a condition from the list." },
                noMatch: { ja: "一致するテキストが見つかりませんでした。", en: "No matching text was found." }
            }
        },
        kinsoku: {
            none: { ja: "なし", en: "None" },
            hard: { ja: "強い禁則", en: "Strict" },
            soft: { ja: "弱い禁則", en: "Loose" },
            softV2: { ja: "弱い禁則 v2", en: "Loose v2" }
        },
        align: {
            roman: { ja: "欧文ベースライン", en: "Roman Baseline" },
            top: { ja: "仮想ボディの上/右", en: "Embox Top/Right" },
            center: { ja: "中央", en: "Embox Centre" },
            bottom: { ja: "仮想ボディの下/左", en: "Embox Bottom/Left" },
            icfTop: { ja: "平均字面の上/右", en: "ICF Box Top/Right" },
            icfBottom: { ja: "平均字面の下/左", en: "ICF Box Bottom/Left" }
        },
        leading: {
            typeTop: { ja: "仮想ボディの上", en: "Top-to-top (virtual body)" },
            typeBaseline: { ja: "欧文ベースライン", en: "Baseline-to-baseline" }
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
        role: {
            body: { ja: "本文", en: "Body" },
            heading: { ja: "見出し", en: "Heading" }
        },
        autoKern: {
            mono: { ja: "和文等幅", en: "Metrics - Roman Only" },
            zero: { ja: "0", en: "0" },
            metrics: { ja: "メトリクス", en: "Metrics" },
            optical: { ja: "オプティカル", en: "Optical" }
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
        button: {
            addPreset: { ja: "追加", en: "Add" },
            overwritePreset: { ja: "上書き", en: "Overwrite" },
            deletePreset: { ja: "削除", en: "Delete" },
            exportPreset: { ja: "書き出し", en: "Export" },
            reset: { ja: "リセット", en: "Reset" },
            reload: { ja: "再読み込み", en: "Reload" },
            hiddenChar: { ja: "制御文字", en: "Hidden Chars" },
            toApparent: { ja: "見かけ←→実サイズ", en: "Apparent ←→ Actual" }
        },
        tip: {
            docFonts: { ja: "書類で使用中のフォント一覧。選ぶと選択テキストに適用します。", en: "Fonts used in the document. Click one to apply it to the selection." },
            presets: { ja: "登録した文字組みプリセット（フォント＋カーニング・ツメ・トラッキング）。選ぶと一括適用します。", en: "Saved type presets (font + kerning / Tsume / tracking). Click one to apply them together." },
            addPreset: { ja: "現在の選択の設定をプリセットとして保存します（同じフォントは上書き）。", en: "Save the current selection's settings as a preset (same font overwrites)." },
            overwritePreset: { ja: "リストで選択中のプリセットを、現在の選択の設定で上書きします。", en: "Overwrite the preset selected in the list with the current selection's settings." },
            deletePreset: { ja: "リストで選択中のプリセットを削除します。", en: "Delete the preset selected in the list." },
            exportPreset: { ja: "登録済みのプリセットを JSON ファイルに書き出します。", en: "Export the saved presets to a JSON file." },
            fontSize: { ja: "選択テキストのフォントサイズ。", en: "Font size of the selection." },
            scale: { ja: "水平比率・垂直比率を同じ値でまとめて設定します。", en: "Sets the horizontal and vertical scale together to the same value." },
            apparent: { ja: "フォントサイズ×比率で計算した、実質的なサイズです。", en: "The effective size, computed as font size × scale." },
            autoKern: { ja: "自動カーニング方式（和文等幅／0／メトリクス／オプティカル）。", en: "Auto-kerning method (Japanese equal width / 0 / Metrics / Optical)." },
            tsume: { ja: "文字ツメ（0〜100%）。隣接する文字の食い込み量。", en: "Tsume (0–100%): how much adjacent characters tighten." },
            tracking: { ja: "字間（1/1000em）をまとめて調整します（-100〜500）。", en: "Adjusts overall letter spacing in 1/1000 em (-100 to 500)." },
            align: { ja: "文字の縦方向の揃え基準。「その他」で残りの基準をポップアップから選びます。", en: "Vertical alignment basis. Use “Other” to pick the rest from the popup." },
            justify: { ja: "段落の行揃え。テキストの見た目の位置を保ったまま変更します。", en: "Paragraph justification, applied while keeping the text's visual position." },
            role: { ja: "本文＝和文等幅／行送り150%／均等配置（最終行左）、見出し＝メトリクス／行送り115%／左揃え をまとめて適用。", en: "Body = Japanese equal width / 150% leading / justify (last left); Heading = Metrics / 115% leading / left." },
            leading: { ja: "行送り。115%／150%／その他（%で直接入力）／自動。", en: "Leading: 115% / 150% / Other (enter a %) / Auto." },
            leadingEffective: { ja: "実質の行送り（フォントサイズ×行送り%）。ここに値を入れると % を逆算して設定します。", en: "The effective leading (font size × leading %). Enter a value here to set the leading by back-calculating the %." },
            leadingType: { ja: "行送りを測る基準位置（仮想ボディの上／欧文ベースライン）。", en: "The reference position for measuring leading (virtual body top / Roman baseline)." },
            mojikumi: { ja: "段落の文字組みアキ量設定（約物の詰め方など）をまとめて適用します。", en: "Applies a mojikumi spacing set (punctuation spacing, etc.) to the paragraphs." },
            kinsoku: { ja: "段落の禁則処理（強い禁則／弱い禁則など）をまとめて適用します。", en: "Applies a kinsoku (line-break) set to the paragraphs." },
            reset: { ja: "標準値に戻します（フォントサイズ12pt・比率100%・ツメ0・トラッキング0・自動カーニングメトリクス・欧文ベースライン・左揃え・行送り115%・文字組みツメ組み・禁則弱い禁則v2）。", en: "Reset to defaults (12pt / 100% scale / 0 Tsume / 0 tracking / Metrics kerning / Roman baseline / left / 115% leading / Tight mojikumi / Loose v2 kinsoku)." },
            reload: { ja: "選択中のテキストの現在値を読み取り直して UI に反映します。", en: "Re-read the current values from the selection and reflect them in the UI." },
            hiddenChar: { ja: "制御文字（改行・スペースなど）の表示／非表示を切り替えます。", en: "Toggle the display of hidden characters (returns, spaces, etc.)." },
            toApparent: { ja: "サイズ×比率を実フォントサイズに焼き込んで比率100%に、もう一度押すと元のサイズ・比率へ戻します（相互変換）。", en: "Bakes size × scale into the actual font size at 100%; press again to restore the original size and scale (round-trip)." },
            typeScale: { ja: "基準サイズと倍率からタイプスケールを生成します。行をクリックすると選択テキストに適用します。", en: "Generate a type scale from a base size and ratio. Click a row to apply it to the selection." },
            typeScaleBase: { ja: "タイプスケールの基準となるフォントサイズ。", en: "The base font size for the type scale." },
            typeScaleRatio: { ja: "各段の比率（音程比）。基準サイズにこの比率を掛け合わせて各サイズを生成します。", en: "The step ratio (musical interval). Each size is the base multiplied by this ratio." }
        },
        msg: {
            applyError: { ja: "適用に失敗しました", en: "Apply failed" }
        }
    };

    /* 言語に応じたラベル文字列を取得 / Resolve a label string for the current language */
    function getLocalizedText(entry) {
        if (!entry) return "";
        return entry[currentLanguage] || entry.ja || entry.en || "";
    }

    /* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
    function labelText(entry) {
        return getLocalizedText(entry) + (currentLanguage === "ja" ? "：" : ":");
    }

    // =========================================
    // メインエンジンで実行する DOM 処理 / DOM helpers run on the main engine
    //
    // 以下の関数群は toString() で連結し、BridgeTalk 本文に同梱して
    // メインエンジン（生きた DOM を持つ）で eval される。常駐エンジン側
    // からは直接呼ばず、本文への埋め込み用途のみ。
    // These are concatenated via toString() and embedded in the BridgeTalk
    // body, then eval'd on the main engine (which has a live DOM). They are
    // not called from the palette engine; they exist only to be embedded.
    // =========================================

    /* 型名を安全に取得（host オブジェクトは typename を優先、JS オブジェクトは constructor.name）
       Safely resolve a type name: typename for host objects, constructor.name for JS objects */
    function getTypeName(obj) {
        if (obj === null || obj === undefined) return "";
        if (obj.typename) return obj.typename;
        try {
            return obj.constructor ? obj.constructor.name : "";
        } catch (e) {
            return "";
        }
    }

    /* 例外を短いメッセージ文字列へ（dispatch の "ERR:" 応答用）/ Reduce an exception to a short message string (for dispatch "ERR:" replies) */
    function errMessage(e) {
        return (e && e.message) ? e.message : String(e);
    }

    /* 1 アイテムからテキスト範囲を収集（グループは中を再帰）/ Collect text ranges from one item (descend into groups) */
    function collectTextRangesFromItem(item, selectedRanges) {
        if (!item) return;
        var itemType = getTypeName(item);
        if (itemType === "TextFrame") {
            selectedRanges.push(item.textRange);
        } else if (itemType === "TextRange") {
            selectedRanges.push(item);
        } else if (itemType === "GroupItem" && item.pageItems) {
            for (var i = 0; i < item.pageItems.length; i++) collectTextRangesFromItem(item.pageItems[i], selectedRanges);
        }
    }

    /* 選択中のテキスト範囲を取得 / Get selected text ranges from current document */
    function getSelectedTextRanges() {
        var activeDoc = app.activeDocument;
        var currentSelection = activeDoc.selection;
        var selectedRanges = [];
        if (!currentSelection) return selectedRanges;
        /* テキスト編集モードでは selection が配列でなく TextRange になる / In text-edit mode the selection is a TextRange, not an array */
        if (getTypeName(currentSelection) === "TextRange") {
            selectedRanges.push(currentSelection);
            return selectedRanges;
        }
        if (currentSelection.length === 0) return selectedRanges;
        for (var i = 0; i < currentSelection.length; i++) {
            collectTextRangesFromItem(currentSelection[i], selectedRanges);
        }
        return selectedRanges;
    }

    /* 選択が触れている段落を「段落全体の範囲」で取得（文字属性を段落単位で適用するため）
       Get the full paragraphs the selection touches (so a char-attr can be applied per paragraph).
       range.paragraphs は重なった段落を返し、各段落オブジェクトは段落全体を指す（選択で切られない）
       range.paragraphs returns the overlapping paragraphs; each paragraph spans the whole paragraph (not clipped) */
    function getSelectedParagraphRanges() {
        var ranges = getSelectedTextRanges();
        var paragraphRanges = [];
        for (var i = 0; i < ranges.length; i++) {
            try {
                var paragraphs = ranges[i].paragraphs;
                for (var j = 0; j < paragraphs.length; j++) paragraphRanges.push(paragraphs[j]);
            } catch (e) { }
        }
        return paragraphRanges;
    }

    /* 方式 ID を AutoKernType の列挙値へ変換 / Resolve a method id to an AutoKernType enum value
       和文等幅は欧文のみメトリクス＝和文は等幅（METRICSROMANONLY）
       Japanese equal width = metrics for Roman only (METRICSROMANONLY) */
    function resolveAutoKernType(methodId) {
        if (methodId === "mono") return AutoKernType.METRICSROMANONLY;
        if (methodId === "metrics") return AutoKernType.AUTO;
        if (methodId === "optical") return AutoKernType.OPTICAL;
        return AutoKernType.NOAUTOKERN;
    }

    /* 選択範囲にカーニング方式を適用 / Apply a kerning method to the given ranges
       メトリクスのときのみプロポーショナルメトリクスをON、それ以外はOFF
       Proportional metrics ON only for Metrics, OFF otherwise */
    function applyKerningToRanges(ranges, kerningMethod) {
        var useProportionalMetrics = (kerningMethod === AutoKernType.AUTO);
        for (var i = 0; i < ranges.length; i++) {
            try {
                ranges[i].characterAttributes.kerningMethod = kerningMethod;
                ranges[i].characterAttributes.proportionalMetrics = useProportionalMetrics;
            } catch (e) {
                // 適用できない範囲はスキップ / Skip ranges that can't take these attributes
            }
        }
    }

    /* カーニング方式だけを適用（プロポーショナル連動なし）。combo は独立したプロポーショナル値を持つため専用
       Apply only the kerning method (no proportional coupling); used by the combo, which carries its own proportional value */
    function applyKerningMethodToRanges(ranges, kerningMethod) {
        for (var i = 0; i < ranges.length; i++) {
            try {
                ranges[i].characterAttributes.kerningMethod = kerningMethod;
            } catch (e) {
                // 適用できない範囲はスキップ / Skip ranges that can't take this attribute
            }
        }
    }

    /* 選択範囲に文字ツメを適用 / Apply Tsume to the given ranges
       0〜100 の百分率（30 = 30%）/ Percentage from 0 to 100 (30 = 30%) */
    function applyTsumeToRanges(ranges, tsumePercent) {
        for (var i = 0; i < ranges.length; i++) {
            try {
                ranges[i].characterAttributes.Tsume = tsumePercent;
            } catch (e) {
                // 適用できない範囲はスキップ / Skip ranges that can't take this attribute
            }
        }
    }

    /* 選択範囲のアキ（前／後）を自動に戻す / Reset the aki (before/after) on the given ranges to auto
       akiLeft / akiRight に -1 を入れると「自動」/ Assigning -1 to akiLeft / akiRight means "auto" */
    function clearAkiOnRanges(ranges) {
        for (var i = 0; i < ranges.length; i++) {
            try {
                ranges[i].characterAttributes.akiLeft = -1;
                ranges[i].characterAttributes.akiRight = -1;
            } catch (e) {
                // 適用できない範囲はスキップ / Skip ranges that can't take these attributes
            }
        }
    }

    /* 選択範囲にフォントを適用 / Apply a font (by PostScript name) to the given ranges */
    function applyFontToRanges(ranges, psName) {
        var font = null;
        try { font = app.textFonts.getByName(psName); } catch (e) { return; }
        if (!font) return;
        for (var i = 0; i < ranges.length; i++) {
            try { ranges[i].characterAttributes.textFont = font; } catch (e2) { }
        }
    }

    /* 文字揃え ID を StyleRunAlignmentType へ変換 / Resolve an alignment id to StyleRunAlignmentType */
    function resolveAlignment(alignId) {
        if (alignId === "top") return StyleRunAlignmentType.top;
        if (alignId === "center") return StyleRunAlignmentType.center;
        if (alignId === "bottom") return StyleRunAlignmentType.bottom;
        if (alignId === "icftop") return StyleRunAlignmentType.icfTop;
        if (alignId === "icfbottom") return StyleRunAlignmentType.icfBottom;
        return StyleRunAlignmentType.ROMANBASELINE;
    }

    /* StyleRunAlignmentType を ID 文字列へ / Convert a StyleRunAlignmentType value to an id string */
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

    /* 自動カーニング method を ID 文字列へ / Convert a kerning method (AutoKernType) to an id string */
    function kernMethodToId(value) {
        var valueText = String(value);
        if (valueText === String(AutoKernType.AUTO)) return "metrics";
        if (valueText === String(AutoKernType.OPTICAL)) return "optical";
        if (valueText === String(AutoKernType.METRICSROMANONLY)) return "mono";
        if (valueText === String(AutoKernType.NOAUTOKERN)) return "zero";
        return "";
    }

    /* 行揃え（Justification）を ID 文字列へ / Convert a paragraph justification to an id string */
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

    /* ID 文字列を行揃え（Justification）へ / Resolve an id string to a Justification value */
    function resolveJustification(justifyId) {
        if (justifyId === "center") return Justification.CENTER;
        if (justifyId === "right") return Justification.RIGHT;
        if (justifyId === "justifyLeft") return Justification.FULLJUSTIFYLASTLINELEFT;
        if (justifyId === "justifyCenter") return Justification.FULLJUSTIFYLASTLINECENTER;
        if (justifyId === "justifyRight") return Justification.FULLJUSTIFYLASTLINERIGHT;
        if (justifyId === "justifyAll") return Justification.FULLJUSTIFY;
        return Justification.LEFT;
    }

    /* テキスト範囲っぽい型か / Is this a text-range-like type */
    function isTextRangeLikeType(typeName) {
        return typeName === "TextRange" || typeName === "InsertionPoint" || typeName === "Character" ||
            typeName === "Word" || typeName === "Line" || typeName === "Paragraph";
    }

    /* 親をたどって TextFrame を返す / Walk up parents to the enclosing TextFrame */
    function findParentTextFrame(item) {
        var current = item;
        for (var i = 0; i < 20; i++) {
            if (!current) return null;
            if (getTypeName(current) === "TextFrame") return current;
            try { current = current.parent; } catch (e) { return null; }
        }
        return null;
    }

    /* TextFrame の重複判定キー / A dedup key for a TextFrame
       uuid があれば一意。無い環境では「位置＋bounds＋文字数＋先頭文字列」で近似するが、
       これらが完全一致する別フレームが重なっていると同一視され、片方がスキップされうる（レアケース）
       uuid is unique when available; otherwise we approximate with position + bounds + length +
       a content prefix. Two distinct frames sharing all of these (fully overlapping) could still
       collide and cause one to be skipped (rare) */
    function getTextFrameKey(frame) {
        try { if (frame.uuid) return frame.uuid; } catch (e) { }
        var bounds = frame.visibleBounds;
        var contents = "";
        try { contents = String(frame.contents); } catch (eContents) { }
        var contentKey = contents.length + ":" + contents.substring(0, 16);
        return frame.typename + ":" + frame.position[0] + ":" + frame.position[1] + ":" + bounds.join(":") + ":" + contentKey;
    }

    /* 選択から行揃え対象（フレーム＋範囲）を収集 / Collect justification targets (frame + range) from a selection */
    function collectJustifyTargets(selection) {
        var targets = [];
        var seenFrameKeys = {};
        var itemList = (selection && selection.typename) ? [selection] : (selection || []);
        for (var i = 0; i < itemList.length; i++) collectJustifyTargetsFromItem(itemList[i], targets, seenFrameKeys);
        return targets;
    }

    /* 1 アイテムから対象を収集（グループは中を走査）/ Collect targets from one item (descend into groups) */
    function collectJustifyTargetsFromItem(item, targets, seenFrameKeys) {
        if (!item) return;
        var typeName = getTypeName(item);
        if (typeName === "TextFrame") { addJustifyTarget(item, item.textRange, targets, seenFrameKeys); return; }
        if (isTextRangeLikeType(typeName)) {
            var frame = findParentTextFrame(item);
            if (frame) addJustifyTarget(frame, item, targets, seenFrameKeys);
            return;
        }
        if (typeName === "GroupItem" && item.pageItems) {
            for (var i = 0; i < item.pageItems.length; i++) collectJustifyTargetsFromItem(item.pageItems[i], targets, seenFrameKeys);
        }
    }

    /* 対象を重複なく追加 / Add a target, skipping duplicate frames */
    function addJustifyTarget(frame, range, targets, seenFrameKeys) {
        if (!frame || !range || getTypeName(frame) !== "TextFrame") return;
        var frameKey = getTextFrameKey(frame);
        if (seenFrameKeys[frameKey]) return;
        seenFrameKeys[frameKey] = true;
        targets.push({ frame: frame, range: range });
    }

    /* 範囲内の全段落に行揃えを適用 / Apply justification to every paragraph in the range */
    function applyJustificationToParagraphs(range, justifyValue) {
        try {
            if (range.paragraphs && range.paragraphs.length) {
                for (var i = 0; i < range.paragraphs.length; i++) {
                    range.paragraphs[i].paragraphAttributes.justification = justifyValue;
                }
            }
        } catch (e) { }
        try { range.paragraphAttributes.justification = justifyValue; } catch (e2) { }
        try {
            if (range.parent && getTypeName(range.parent) === "TextFrame") {
                range.parent.textRange.paragraphAttributes.justification = justifyValue;
            }
        } catch (e3) { }
    }

    /* 左揃えはスクリプトから無視されることがあるため、一時リサイズで段落属性を
       リフレッシュさせてから元に戻す（位置補正は呼び出し側で実施）
       Illustrator can ignore Justification.LEFT from scripts; a temporary resize
       forces the attributes to refresh, then the frame is restored (caller fixes position)

       ⚠️ 副作用リスクあり: resize(200,200)→resize(50,50) は強引な回避策で、
          既に変形（回転・シアー等）済みのテキストやエリアテキストでは、matrix/position
          の復元が完全でなく形状やパスがわずかに崩れる可能性がある。Justification.LEFT
          の既知バグ（RIGHT/CENTER は効くが LEFT は無視される）専用のフォールバック。
       ⚠️ Side-effect risk: the resize(200,200)→resize(50,50) dance is a brute-force
          workaround. On already-transformed (rotated/sheared) text or area text, the
          matrix/position restore may be imperfect and slightly distort the shape/path.
          Used only as a fallback for the known Justification.LEFT bug (RIGHT/CENTER work). */
    function forceLeftJustification(target) {
        var frame = target.frame;
        var pos = [frame.position[0], frame.position[1]];
        var matrix = null;
        try { matrix = frame.matrix; } catch (e0) { }
        try { frame.resize(200, 200); } catch (e1) { }
        applyJustificationToParagraphs(target.range, Justification.LEFT);
        try { if (frame.story && frame.story.textRange) frame.story.textRange.justification = Justification.LEFT; } catch (e2) { }
        try { if (frame.textRange) frame.textRange.justification = Justification.LEFT; } catch (e3) { }
        try { frame.resize(50, 50); } catch (e4) { }
        try { frame.position = pos; } catch (e5) { }
        try { if (matrix) frame.matrix = matrix; } catch (e6) { }
        applyJustificationToParagraphs(target.range, Justification.LEFT);
    }

    /* 行揃えを適用しつつ、フレームの見た目位置（visibleBounds 左上）を維持
       Apply justification while preserving the frame's visual position (top-left of visibleBounds) */
    function justifyKeepingPosition(target, justifyValue, justifyId) {
        var frame = target.frame;
        var before = frame.visibleBounds;
        if (justifyId === "left") {
            forceLeftJustification(target);
        } else {
            applyJustificationToParagraphs(target.range, justifyValue);
        }
        var after = frame.visibleBounds;
        var dx = before[0] - after[0];
        var dy = before[1] - after[1];
        if (Math.abs(dx) > 0.001 || Math.abs(dy) > 0.001) frame.translate(dx, dy);
    }

    /* TextFrame / TextRange から characterAttributes を取得（それ以外は null）/ Get characterAttributes from a TextFrame / TextRange */
    function getAlignmentCharAttrs(item) {
        if (item.typename === "TextFrame") return item.textRange.characterAttributes;
        if (item.typename === "TextRange") return item.characterAttributes;
        return null;
    }

    /* グループを潜って最初のテキストの characterAttributes を返す / Descend into groups for the first text's characterAttributes */
    function getFirstAlignmentCharAttrs(item) {
        if (item.typename === "GroupItem") {
            for (var j = 0; j < item.pageItems.length; j++) {
                var found = getFirstAlignmentCharAttrs(item.pageItems[j]);
                if (found !== null) return found;
            }
            return null;
        }
        return getAlignmentCharAttrs(item);
    }

    /* テキストへ再帰的に文字揃えを設定（グループは中を走査）/ Recursively set alignment on text (descends into groups) */
    function applyStyleRunAlignment(item, alignmentValue) {
        if (item.typename === "GroupItem") {
            var applied = 0;
            for (var j = 0; j < item.pageItems.length; j++) applied += applyStyleRunAlignment(item.pageItems[j], alignmentValue);
            return applied;
        }
        var attrs = getAlignmentCharAttrs(item);
        if (attrs === null) return 0;
        attrs.alignment = alignmentValue;
        return 1;
    }

    /* 選択範囲にトラッキングを適用 / Apply tracking to the given ranges（1/1000 em 単位）*/
    function applyTrackingToRanges(ranges, trackingValue) {
        for (var i = 0; i < ranges.length; i++) {
            try {
                ranges[i].characterAttributes.tracking = trackingValue;
            } catch (e) {
                // 適用できない範囲はスキップ / Skip ranges that can't take this attribute
            }
        }
    }

    /* 選択範囲にプロポーショナルメトリクスを適用（段落単位で使うため独立関数化）
       Apply proportional metrics to the given ranges (its own function so it can be applied per paragraph) */
    function applyProportionalToRanges(ranges, useProportional) {
        var on = useProportional ? true : false;
        for (var i = 0; i < ranges.length; i++) {
            try {
                ranges[i].characterAttributes.proportionalMetrics = on;
            } catch (e) {
                // 適用できない範囲はスキップ / Skip ranges that can't take this attribute
            }
        }
    }

    /* 行頭で参照する文字数 / Number of leading characters to sample per line */
    var LINE_FONT_SIZE_SAMPLE_COUNT = 5;

    /* 行から先頭数文字のフォントサイズを採取 / Sample font sizes from the first few characters of a line */
    function getLineSampleFontSizes(line, sampleCount) {
        var fontSizes = [];
        if (!line || !line.characters || line.characters.length === 0) return fontSizes;
        var maxCount = Math.min(line.characters.length, sampleCount);
        for (var i = 0; i < maxCount; i++) {
            try {
                var size = line.characters[i].characterAttributes.size;
                if (!isNaN(size)) fontSizes.push(size);
            } catch (e) { }
        }
        return fontSizes;
    }

    /* 最頻値を返す（同数なら大きい方）/ Return the most frequent value (ties favor the larger) */
    function getMostFrequentNumber(values) {
        if (!values || values.length === 0) return NaN;
        var counts = {};
        var bestValue = values[0];
        var bestCount = 0;
        for (var i = 0; i < values.length; i++) {
            var valueKey = String(values[i]);
            if (!counts[valueKey]) counts[valueKey] = { value: values[i], count: 0 };
            counts[valueKey].count++;
            if (counts[valueKey].count > bestCount) {
                bestCount = counts[valueKey].count;
                bestValue = counts[valueKey].value;
            } else if (counts[valueKey].count === bestCount && counts[valueKey].value > bestValue) {
                bestValue = counts[valueKey].value;
            }
        }
        return bestValue;
    }

    /* 行の基準フォントサイズ（先頭数文字の最頻値）/ Base font size of a line (mode of the first few chars) */
    function getBaseFontSizeFromLine(line) {
        var fontSizes = getLineSampleFontSizes(line, LINE_FONT_SIZE_SAMPLE_COUNT);
        if (!fontSizes || fontSizes.length === 0) return NaN;
        return getMostFrequentNumber(fontSizes);
    }

    /* 処理可能なテキストフレームか判定し、該当すれば返す / Return the item if it is a processable TextFrame */
    function getProcessableTextFrame(item) {
        if (!item || item.typename !== "TextFrame") return null;
        if (!item.contents) return null;
        if (!item.lines || item.lines.length === 0) return null;
        return item;
    }

    /* 選択から重複なしの処理対象テキストフレームを収集 / Collect unique processable text frames from the selection */
    function collectLeadingFrames(selectionItems) {
        var frames = [];
        var seenFrameKeys = {};
        var itemList = (selectionItems && selectionItems.typename) ? [selectionItems] : (selectionItems || []);
        for (var i = 0; i < itemList.length; i++) collectLeadingFramesFromItem(itemList[i], frames, seenFrameKeys);
        return frames;
    }

    /* 1 アイテムからフレームを収集（グループは中を走査、テキスト編集モードの範囲は親フレームへ）
       Collect frames from one item (descend into groups; resolve text-edit ranges to their parent frame) */
    function collectLeadingFramesFromItem(item, frames, seenFrameKeys) {
        if (!item) return;
        var typeName = getTypeName(item);
        if (typeName === "TextFrame") { addLeadingFrame(item, frames, seenFrameKeys); return; }
        if (isTextRangeLikeType(typeName)) { addLeadingFrame(findParentTextFrame(item), frames, seenFrameKeys); return; }
        if (typeName === "GroupItem" && item.pageItems) {
            for (var i = 0; i < item.pageItems.length; i++) collectLeadingFramesFromItem(item.pageItems[i], frames, seenFrameKeys);
        }
    }

    /* 処理可能なフレームを重複なく追加 / Add a processable frame, skipping duplicates */
    function addLeadingFrame(frame, frames, seenFrameKeys) {
        var textFrame = getProcessableTextFrame(frame);
        if (!textFrame) return;
        var frameKey = getTextFrameKey(textFrame);
        if (seenFrameKeys[frameKey]) return;
        seenFrameKeys[frameKey] = true;
        frames.push(textFrame);
    }

    /* 行送りの基準 ID を AutoLeadingType へ変換 / Resolve a leading-basis id to AutoLeadingType */
    function resolveLeadingType(typeId) {
        if (typeId === "baseline") return AutoLeadingType.BOTTOMTOBOTTOM;
        return AutoLeadingType.TOPTOTOP;
    }

    /* AutoLeadingType を ID 文字列へ / Convert an AutoLeadingType value to an id string */
    function leadingTypeToId(typeValue) {
        if (typeValue === AutoLeadingType.BOTTOMTOBOTTOM) return "baseline";
        if (typeValue === AutoLeadingType.TOPTOTOP) return "top";
        return "";
    }

    /* 段落に自動行送りを設定：自動行送り量（%）を段落属性に、各文字を autoLeading=true に
       Set auto-leading on paragraphs: the amount (%) on the paragraph attribute, autoLeading=true on the characters
       これにより Illustrator 上は常に「自動」表示となり、行送りはフォントサイズに追従する
       So Illustrator always shows "Auto" and the leading follows the font size */
    function applyAutoLeadingToParagraphs(paragraphRanges, percent) {
        for (var i = 0; i < paragraphRanges.length; i++) {
            try {
                paragraphRanges[i].paragraphAttributes.autoLeadingAmount = percent;
                paragraphRanges[i].characterAttributes.autoLeading = true;
            } catch (e) {
                // 適用できない範囲はスキップ / Skip ranges that can't take these attributes
            }
        }
    }

    /* 行送りの基準（leadingType）だけを設定し、行送り値は変更しない
       Set only the leading basis (leadingType); leaves the leading values untouched */
    function applyLeadingBasisToFrames(frames, leadingType) {
        for (var i = 0; i < frames.length; i++) {
            try { frames[i].textRange.leadingType = leadingType; } catch (e) { }
        }
    }

    /* 段落にハイフネーションを適用（段落属性）/ Apply hyphenation to the given paragraph ranges (a paragraph attribute) */
    function applyHyphenationToRanges(ranges, useHyphenation) {
        var on = useHyphenation ? true : false;
        for (var i = 0; i < ranges.length; i++) {
            try { ranges[i].paragraphAttributes.hyphenation = on; } catch (e) { }
        }
    }

    /* テキストフレーム内の全段落に文字組みアキ量設定を適用 / Apply a mojikumi set to every paragraph in a text frame */
    function applyMojikumiToTextFrame(textFrame, mojikumiValue) {
        var paragraphs = textFrame.paragraphs;
        for (var paragraphIndex = 0; paragraphIndex < paragraphs.length; paragraphIndex++) {
            try { paragraphs[paragraphIndex].paragraphAttributes.mojikumi = mojikumiValue; } catch (e) { }
        }
    }

    /* 種類で振り分け（TextFrame は直接、GroupItem は再帰）/ Dispatch by type (TextFrame directly, GroupItem recursively) */
    function applyMojikumiToItem(item, mojikumiValue) {
        var typeName = getTypeName(item);
        if (typeName === "TextFrame") {
            applyMojikumiToTextFrame(item, mojikumiValue);
        } else if (typeName === "GroupItem" && item.pageItems) {
            for (var i = 0; i < item.pageItems.length; i++) applyMojikumiToItem(item.pageItems[i], mojikumiValue);
        } else if (isTextRangeLikeType(typeName)) {
            var frame = findParentTextFrame(item);
            if (frame) applyMojikumiToTextFrame(frame, mojikumiValue);
        }
    }

    /* 選択全体に文字組みアキ量設定を適用 / Apply a mojikumi set to the whole selection
       mojikumiIndex === -1 は「なし」/ -1 means "None" */
    function applyMojikumiToSelection(selection, mojikumiIndex) {
        var mojikumiValue = (mojikumiIndex === -1) ? "なし" : app.activeDocument.mojikumiSet[mojikumiIndex];
        var itemList = (selection && selection.typename) ? [selection] : (selection || []);
        for (var i = 0; i < itemList.length; i++) applyMojikumiToItem(itemList[i], mojikumiValue);
    }

    /* mojikumiSet 名からインデックスを引く（見つからなければ -2＝不明）/ Look up a mojikumiSet index by name (-2 = unknown) */
    function matchMojikumiName(name) {
        try {
            var mojikumiSets = app.activeDocument.mojikumiSet;
            for (var i = 0; i < mojikumiSets.length; i++) {
                if (mojikumiSets[i].name === name) return i;
            }
        } catch (e) { }
        return -2;
    }

    /* 段落の文字組みアキ量設定を ID へ変換 / Resolve a paragraph's mojikumi to an id
       戻り値: 0以上=mojikumiSet のインデックス / -1=なし / -2=読み取り不可・不明（UI を変更しない）
       getter は環境により MojikumiSet オブジェクト／名前文字列のどちらも返しうるため両対応
       Returns: >=0 mojikumiSet index / -1 = none / -2 = unreadable or unknown (leave UI untouched) */
    function mojikumiToId(paragraph) {
        var mojikumiAttr;
        try { mojikumiAttr = paragraph.paragraphAttributes.mojikumi; } catch (e) { return -2; }
        if (mojikumiAttr === undefined || mojikumiAttr === null) return -1;
        if (typeof mojikumiAttr === "string") {
            return (mojikumiAttr === "" || mojikumiAttr === "なし") ? -1 : matchMojikumiName(mojikumiAttr);
        }
        var mojikumiName;
        try { mojikumiName = (mojikumiAttr.name !== undefined && mojikumiAttr.name !== null) ? mojikumiAttr.name : ""; } catch (e2) { return -2; }
        if (mojikumiName === "") return -2;
        if (mojikumiName === "なし") return -1;
        return matchMojikumiName(mojikumiName);
    }

    /* テキストフレームに禁則を適用（「なし」= "None" は scripting では設定不可で例外になる）
       Apply kinsoku to a text frame ("None" can't be set via scripting and throws) */
    function applyKinsokuToTextFrame(textFrame, kinsokuName) {
        try {
            textFrame.textRange.paragraphAttributes.kinsoku = kinsokuName;
        } catch (e) {
            /* 「なし」は scripting から設定できない（アクション再生が必要）/ "None" can't be set via scripting */
        }
    }

    /* 種類で振り分け（TextFrame は直接、GroupItem は再帰、テキスト範囲は親フレームへ）
       Dispatch by type (TextFrame directly, GroupItem recursively, text ranges to their parent frame) */
    function applyKinsokuToItem(item, kinsokuName) {
        var typeName = getTypeName(item);
        if (typeName === "TextFrame") {
            applyKinsokuToTextFrame(item, kinsokuName);
        } else if (typeName === "GroupItem" && item.pageItems) {
            for (var i = 0; i < item.pageItems.length; i++) applyKinsokuToItem(item.pageItems[i], kinsokuName);
        } else if (isTextRangeLikeType(typeName)) {
            var frame = findParentTextFrame(item);
            if (frame) applyKinsokuToTextFrame(frame, kinsokuName);
        }
    }

    /* 選択全体に禁則を適用 / Apply kinsoku to the whole selection */
    function applyKinsokuToSelection(selection, kinsokuName) {
        var itemList = (selection && selection.typename) ? [selection] : (selection || []);
        for (var i = 0; i < itemList.length; i++) applyKinsokuToItem(itemList[i], kinsokuName);
    }

    /* 段落の禁則を id 文字列へ / Resolve a paragraph's kinsoku to an id string
       禁則「なし」の段落は getter が例外（Error 9563）を投げるため "None" を返す
       A no-kinsoku paragraph throws on read (Error 9563), so return "None" */
    function kinsokuToId(paragraph) {
        try {
            var kinsokuValue = paragraph.paragraphAttributes.kinsoku;
            if (kinsokuValue === undefined || kinsokuValue === null || kinsokuValue === "") return "None";
            return String(kinsokuValue);
        } catch (e) {
            return "None";
        }
    }

    /* 選択へ文字揃え（縦方向の揃え）を適用 / Apply character alignment to the selection */
    function applyAlignmentToSelection(selection, alignValue) {
        if (getTypeName(selection) === "TextRange") { selection.characterAttributes.alignment = alignValue; return 1; }
        var count = 0;
        if (selection && selection.length) {
            for (var i = 0; i < selection.length; i++) count += applyStyleRunAlignment(selection[i], alignValue);
        }
        return count;
    }

    /* 選択の行送り状態を読み取り、初期表示用にエンコード / Read the selection's leading state for initial reflection
       戻り値: count|isAuto|leadingPt|baseFontSizePt|autoAmount|leadingTypeId|
               paragraphCount|alignmentId|fontFamily(enc)|fontStyle(enc)|fontSizePt|kernId|justifyId|hScale|mojikumiId|kinsokuId

       注意: 値系（フォント・サイズ・行送り・カーニング・揃え・禁則など）は「最初のフレームの
       先頭文字／先頭段落」だけを代表値として読む簡略化。複数フレームを選択したり、1フレーム内で
       設定が混在している場合、UI 表示は先頭の値になり実態とズレる（count だけは全体数を返す）。
       Note: the value fields (font, size, leading, kerning, alignment, kinsoku, ...) are read only
       from the FIRST frame's first character/paragraph as a representative. With a multi-frame
       selection or mixed settings within a frame, the UI shows the first value and may not reflect
       the whole selection (only `count` reports the total). */
    function readLeadingState(frames) {
        var count = frames.length;
        var isAuto = 0, leadingPt = NaN, baseFontSizePt = NaN, autoAmount = 175;
        var leadingTypeId = "", paragraphCount = 0, alignmentId = "";
        var fontFamily = "", fontStyle = "", fontSizePt = NaN, kernId = "", justifyId = "", hScale = NaN, mojikumiId = -1, kinsokuId = "None";
        for (var i = 0; i < frames.length; i++) {
            try {
                var lines = frames[i].lines;
                if (lines && lines.length > 0 && lines[0].characters.length > 0) {
                    var firstChar = lines[0].characters[0];
                    isAuto = firstChar.characterAttributes.autoLeading ? 1 : 0;
                    leadingPt = firstChar.characterAttributes.leading;
                    baseFontSizePt = getBaseFontSizeFromLine(lines[0]);
                    alignmentId = alignmentToId(firstChar.characterAttributes.alignment);
                    fontSizePt = firstChar.characterAttributes.size;
                    hScale = firstChar.characterAttributes.horizontalScale;
                    kernId = kernMethodToId(firstChar.characterAttributes.kerningMethod);
                    try {
                        var font = firstChar.characterAttributes.textFont;
                        if (font) { fontFamily = font.family; fontStyle = font.style; }
                    } catch (eFont) { }
                    var paragraphs = frames[i].paragraphs;
                    if (paragraphs && paragraphs.length > 0) {
                        autoAmount = paragraphs[0].paragraphAttributes.autoLeadingAmount;
                        justifyId = justificationToId(paragraphs[0].paragraphAttributes.justification);
                        mojikumiId = mojikumiToId(paragraphs[0]);
                        kinsokuId = kinsokuToId(paragraphs[0]);
                    }
                    leadingTypeId = leadingTypeToId(frames[i].textRange.leadingType);
                    break;
                }
            } catch (e) { }
        }
        for (var k = 0; k < frames.length; k++) {
            try { paragraphCount += frames[k].paragraphs.length; } catch (e2) { }
        }
        return [count, isAuto, leadingPt, baseFontSizePt, autoAmount, leadingTypeId,
            paragraphCount, alignmentId, encodeURIComponent(fontFamily), encodeURIComponent(fontStyle),
            fontSizePt, kernId, justifyId, hScale, mojikumiId, kinsokuId].join("|");
    }

    // =========================================
    // メインエンジン委譲 / Main-engine delegation (BridgeTalk)
    // =========================================

    // =========================================
    // 使用フォント一覧（DocumentFontListSelector 由来）/ Used-fonts list (ported from DocumentFontListSelector)
    // これらもメインエンジンへ委譲する DOM 処理 / DOM helpers also delegated to the main engine
    // =========================================

    /* 小数2桁までで末尾ゼロを落とす / Round to 2 decimals and trim trailing zeros */
    function formatNumber(numericValue) {
        if (numericValue === null || numericValue === undefined || isNaN(numericValue)) return "?";
        var roundedValue = Math.round(numericValue * 100) / 100;
        return String(roundedValue);
    }

    /* フォントの表示名（ファミリー＋スタイル）/ Display name of a font (family + style) */
    function getFontDisplayName(textFont) {
        if (!textFont) return "(no font)";
        var fontFamily = textFont.family || textFont.name || "";
        var fontStyle = textFont.style || "";
        return fontFamily + (fontStyle ? " " + fontStyle : "");
    }

    /* 1文字の属性から組み合わせオブジェクトを作る（読めなければ null）/ Build a combination object from one character (null if unreadable) */
    function readComboFromCharacter(character) {
        try {
            var charAttributes = character.characterAttributes;
            var textFont = charAttributes.textFont;
            if (!textFont) return null;
            return {
                psName: textFont.name,
                displayName: getFontDisplayName(textFont),
                size: charAttributes.size,
                leading: charAttributes.leading,
                autoLeading: charAttributes.autoLeading ? true : false,
                tsume: charAttributes.Tsume,
                tracking: charAttributes.tracking,
                kernId: kernMethodToId(charAttributes.kerningMethod),
                proportionalMetrics: charAttributes.proportionalMetrics ? true : false
            };
        } catch (e) {
            return null;
        }
    }

    /* 組み合わせの重複排除キー / Dedupe key for a combination */
    function comboKey(combo) {
        return [
            combo.psName,
            formatNumber(combo.size),
            combo.autoLeading ? "auto" : formatNumber(combo.leading),
            String(Math.round(combo.tsume)),
            String(Math.round(combo.tracking)),
            combo.kernId,
            combo.proportionalMetrics ? "p1" : "p0"
        ].join("|");
    }

    /* 現在のアートボードの矩形を取得（取れなければ null）/ Get the active artboard rect ([L,T,R,B]); null if unavailable */
    function getActiveArtboardRect(activeDocument) {
        try {
            var activeIndex = activeDocument.artboards.getActiveArtboardIndex();
            return activeDocument.artboards[activeIndex].artboardRect;
        } catch (e) {
            return null;
        }
    }

    /* テキストフレームがアートボード矩形と重なるか / Whether a text frame overlaps the artboard rect */
    function frameOverlapsArtboard(textFrame, artboardRect) {
        if (!artboardRect) return true;
        var bounds;
        try { bounds = textFrame.geometricBounds; } catch (e) { return false; }
        var frameLeft = bounds[0], frameTop = bounds[1], frameRight = bounds[2], frameBottom = bounds[3];
        var artLeft = artboardRect[0], artTop = artboardRect[1], artRight = artboardRect[2], artBottom = artboardRect[3];
        if (frameRight < artLeft || frameLeft > artRight) return false;
        if (frameBottom > artTop || frameTop < artBottom) return false;
        return true;
    }

    /* アイテムが表示されているか（自身と親グループの hidden、所属レイヤーの visible を遡って確認）/ Whether an item is visible */
    function isItemVisible(pageItem) {
        var node = pageItem;
        try {
            while (node) {
                var nodeType = node.typename;
                if (nodeType === "Document") break;
                if (nodeType === "Layer") {
                    if (!node.visible) return false;
                } else {
                    if (node.hidden) return false;
                }
                node = node.parent;
            }
        } catch (e) { }
        return true;
    }

    /* アイテムがロックされていないか（自身と親グループの locked、所属レイヤーの locked を遡って確認）/ Whether an item is unlocked */
    function isItemUnlocked(pageItem) {
        var node = pageItem;
        try {
            while (node) {
                var nodeType = node.typename;
                if (nodeType === "Document") break;
                if (node.locked) return false;
                node = node.parent;
            }
        } catch (e) { }
        return true;
    }

    /* ドキュメントを走査し、組み合わせを TAB 区切り行（LF 連結）で返す / Scan the document; return combos as TAB-joined rows (LF-joined)
       1行 = psName \t displayName \t size \t leading \t autoLeading(0/1) \t tsume \t tracking \t kernId \t proportionalMetrics(0/1) \t count */
    function collectDocumentCombosRaw(activeDocument, currentArtboardOnly, includeHidden, includeLocked) {
        var comboOrder = [], comboData = {}, comboRows = [];
        var TAB = String.fromCharCode(9), LF = String.fromCharCode(10);
        var artboardRect = currentArtboardOnly ? getActiveArtboardRect(activeDocument) : null;
        var textFrames = activeDocument.textFrames;
        for (var f = 0; f < textFrames.length; f++) {
            if (!includeHidden && !isItemVisible(textFrames[f])) continue;
            if (!includeLocked && !isItemUnlocked(textFrames[f])) continue;
            if (artboardRect && !frameOverlapsArtboard(textFrames[f], artboardRect)) continue;
            var characters;
            try { characters = textFrames[f].characters; } catch (eFrame) { continue; }
            var seenInFrame = {};
            for (var c = 0; c < characters.length; c++) {
                var combo = readComboFromCharacter(characters[c]);
                if (!combo) continue;
                var dedupeKey = comboKey(combo);
                if (!comboData[dedupeKey]) {
                    comboData[dedupeKey] = { combo: combo, count: 0 };
                    comboOrder.push(dedupeKey);
                }
                /* 同じフレーム内で同じ組み合わせを複数回使っても使用数は1だけ増やす / Count each frame once per combo */
                if (!seenInFrame[dedupeKey]) {
                    seenInFrame[dedupeKey] = true;
                    comboData[dedupeKey].count++;
                }
            }
        }
        for (var i = 0; i < comboOrder.length; i++) {
            var entry = comboData[comboOrder[i]];
            var entryCombo = entry.combo;
            comboRows.push([
                entryCombo.psName, entryCombo.displayName, String(entryCombo.size), String(entryCombo.leading),
                entryCombo.autoLeading ? "1" : "0", String(entryCombo.tsume), String(entryCombo.tracking), entryCombo.kernId,
                entryCombo.proportionalMetrics ? "1" : "0", String(entry.count)
            ].join(TAB));
        }
        return comboRows.join(LF);
    }

    /* 選択範囲に組み合わせを適用 / Apply a combination to the given ranges. 戻り値: 適用できた範囲数 / Returns applied count */
    function applyComboToRanges(selectedRanges, combo) {
        var textFont = null;
        try { textFont = app.textFonts.getByName(combo.psName); } catch (eGet) { textFont = null; }
        var appliedCount = 0;
        for (var i = 0; i < selectedRanges.length; i++) {
            /* ラン単位の属性のみ（フォント・サイズ・行送り・トラッキング）。文字ツメ・カーニング・プロポーショナルは段落単位で applyCombo 側
               Run-scoped attrs only (font/size/leading/tracking). Tsume, kerning, proportional apply per paragraph in applyCombo */
            try {
                var charAttributes = selectedRanges[i].characterAttributes;
                if (textFont) charAttributes.textFont = textFont;
                if (!isNaN(combo.size)) charAttributes.size = combo.size;
                if (combo.autoLeading) {
                    charAttributes.autoLeading = true;
                } else if (!isNaN(combo.leading)) {
                    charAttributes.autoLeading = false;
                    charAttributes.leading = combo.leading;
                }
                if (!isNaN(combo.tracking)) charAttributes.tracking = combo.tracking;
                appliedCount++;
            } catch (eApplyRange) { }
        }
        return appliedCount;
    }

    /* 組み合わせ（または matchFontOnly ならフォントのみ）に一致する文字を含むテキストフレームを選択 / Select frames whose characters match the combo (or the font only when matchFontOnly). 戻り値: 選択したフレーム数 */
    function selectFramesMatchingCombo(activeDocument, combo, currentArtboardOnly, includeHidden, includeLocked) {
        var matchFontOnly = combo.matchFontOnly ? true : false;
        var targetKey = matchFontOnly ? combo.psName : comboKey(combo);
        var artboardRect = currentArtboardOnly ? getActiveArtboardRect(activeDocument) : null;
        var textFrames = activeDocument.textFrames;
        var matchedFrames = [];
        for (var f = 0; f < textFrames.length; f++) {
            if (!includeHidden && !isItemVisible(textFrames[f])) continue;
            if (!includeLocked && !isItemUnlocked(textFrames[f])) continue;
            if (artboardRect && !frameOverlapsArtboard(textFrames[f], artboardRect)) continue;
            var characters;
            try { characters = textFrames[f].characters; } catch (eFrame) { continue; }
            for (var c = 0; c < characters.length; c++) {
                var charCombo = readComboFromCharacter(characters[c]);
                if (charCombo && (matchFontOnly ? charCombo.psName : comboKey(charCombo)) === targetKey) { matchedFrames.push(textFrames[f]); break; }
            }
        }
        /* 一致0件のときは現在の選択を保持 / Keep the current selection when nothing matched */
        if (matchedFrames.length === 0) return 0;
        /* 配列を直接代入すれば既存選択を一括置換 / Assigning the array replaces the current selection */
        activeDocument.selection = matchedFrames;
        return matchedFrames.length;
    }

    /* メインエンジンへ送る処理関数（上で定義済みのものを再利用）
       Functions shipped to the main engine (reuse of the helpers above) */
    var WORKER_FUNCS = [
        getTypeName, errMessage, runRangeAction, collectTextRangesFromItem, getSelectedTextRanges, getSelectedParagraphRanges, resolveAutoKernType, applyKerningToRanges, applyKerningMethodToRanges, applyTsumeToRanges,
        applyTrackingToRanges, applyProportionalToRanges, clearAkiOnRanges, applyFontToRanges,
        kernMethodToId, justificationToId, resolveJustification,
        isTextRangeLikeType, findParentTextFrame, getTextFrameKey, collectJustifyTargets,
        collectJustifyTargetsFromItem, addJustifyTarget, applyJustificationToParagraphs,
        forceLeftJustification, justifyKeepingPosition,
        resolveAlignment, alignmentToId, getAlignmentCharAttrs, getFirstAlignmentCharAttrs, applyStyleRunAlignment,
        getLineSampleFontSizes, getMostFrequentNumber, getBaseFontSizeFromLine,
        getProcessableTextFrame, collectLeadingFrames, collectLeadingFramesFromItem, addLeadingFrame,
        resolveLeadingType, leadingTypeToId,
        applyLeadingBasisToFrames, applyAutoLeadingToParagraphs, applyHyphenationToRanges, applyAlignmentToSelection,
        applyMojikumiToTextFrame, applyMojikumiToItem, applyMojikumiToSelection, matchMojikumiName, mojikumiToId,
        applyKinsokuToTextFrame, applyKinsokuToItem, applyKinsokuToSelection, kinsokuToId,
        readLeadingState,
        formatNumber, getFontDisplayName, readComboFromCharacter, comboKey,
        getActiveArtboardRect, frameOverlapsArtboard, isItemVisible, isItemUnlocked,
        collectDocumentCombosRaw, applyComboToRanges, selectFramesMatchingCombo
    ];

    /* 関数配列を toString() で連結してソース文字列にする / Join the function array into a source string
       関数本体のみ送られるため、参照する定数はここで先頭に定義しておく
       Only function bodies are shipped, so constants they reference are declared up front */
    function buildWorkerLibSource() {
        var src = "var LINE_FONT_SIZE_SAMPLE_COUNT=" + LINE_FONT_SIZE_SAMPLE_COUNT + ";\n";
        for (var i = 0; i < WORKER_FUNCS.length; i++) src += WORKER_FUNCS[i].toString() + "\n";
        return src;
    }
    var workerLibCache = null;
    function getWorkerLibSource() {
        if (workerLibCache === null) workerLibCache = buildWorkerLibSource();
        return workerLibCache;
    }

    /* 選択範囲へ作用する定型アクションの共通処理（空チェック→try→redraw→件数返却）
       Shared shape for range-acting actions: empty-check, try, redraw, return the count */
    function runRangeAction(ranges, applyFn) {
        if (ranges.length === 0) return "OK:0";
        try { applyFn(); } catch (e) { return "ERR:" + errMessage(e); }
        app.redraw();
        return "OK:" + ranges.length;
    }

    /* メインエンジンで実行されるディスパッチャ / Dispatcher executed on the main engine
       count: 対象数 / getState: 選択状態 / apply: カーニング / applyTsume: 文字ツメ / applyLeading: 行送り
       結果は "OK:<payload>" または "ERR:<msg>" の文字列で返す */
    function dispatch(action, params) {
        if (app.documents.length === 0) return "ERR:nodoc";
        try { app.activeDocument; } catch (e) { return "ERR:nodoc"; }

        var ranges = getSelectedTextRanges();
        if (action === "count") {
            return "OK:" + ranges.length;
        }
        if (action === "getState") {
            return "OK:" + readLeadingState(collectLeadingFrames(app.activeDocument.selection));
        }
        if (action === "toggleHiddenChar") {
            app.executeMenuCommand("showHiddenChar");
            return "OK:1";
        }
        if (action === "getCurrentFont") {
            if (ranges.length === 0) return "OK:";
            try {
                var charAttrs = ranges[0].characterAttributes;
                var currentFont = charAttrs.textFont;
                if (!currentFont || !currentFont.name) return "OK:";
                var currentKernId = kernMethodToId(charAttrs.kerningMethod);
                var currentTsume = Math.round(charAttrs.Tsume);
                var currentTracking = Math.round(charAttrs.tracking);
                return "OK:" + encodeURIComponent([currentFont.name, currentKernId, currentTsume, currentTracking].join("|"));
            } catch (errCur) {
                return "ERR:" + errMessage(errCur);
            }
        }
        if (action === "applyFont") return runRangeAction(ranges, function () { applyFontToRanges(ranges, params.psName); });
        if (action === "applyFontPreset") return runRangeAction(ranges, function () {
            applyFontToRanges(ranges, params.psName);
            if (params.kern) applyKerningToRanges(ranges, resolveAutoKernType(params.kern));
            if (params.tsume !== undefined && params.tsume !== null) applyTsumeToRanges(ranges, params.tsume);
            if (params.tracking !== undefined && params.tracking !== null) applyTrackingToRanges(ranges, params.tracking);
        });
        if (action === "resolveFontNames") {
            var names = params.psList ? params.psList.split(String.fromCharCode(10)) : [];
            var TAB = String.fromCharCode(9), LF = String.fromCharCode(10), resolved = [];
            for (var nameIndex = 0; nameIndex < names.length; nameIndex++) {
                var displayName = names[nameIndex];
                try {
                    var resolvedFont = app.textFonts.getByName(names[nameIndex]);
                    if (resolvedFont) displayName = resolvedFont.family + (resolvedFont.style ? " " + resolvedFont.style : "");
                } catch (eRes) { }
                resolved.push(names[nameIndex] + TAB + displayName);
            }
            return "OK:" + encodeURIComponent(resolved.join(LF));
        }
        // 自動カーニング・文字ツメは段落単位（選択が触れた段落全体へ）/ Auto-kerning & Tsume apply per paragraph
        if (action === "apply") { var kernParas = getSelectedParagraphRanges(); return runRangeAction(kernParas, function () { applyKerningToRanges(kernParas, resolveAutoKernType(params.method)); }); }
        if (action === "applyTsume") { var tsumeParas = getSelectedParagraphRanges(); return runRangeAction(tsumeParas, function () { applyTsumeToRanges(tsumeParas, params.value); }); }
        // トラッキングはラン（選択）単位のまま / Tracking stays per selection (run)
        if (action === "applyTracking") return runRangeAction(ranges, function () { applyTrackingToRanges(ranges, params.tracking); });
        if (action === "applyJustification") {
            var justifyTargets = collectJustifyTargets(app.activeDocument.selection);
            if (justifyTargets.length === 0) return "OK:0";
            var justifyValue = resolveJustification(params.justify);
            try {
                for (var targetIndex = 0; targetIndex < justifyTargets.length; targetIndex++) {
                    justifyKeepingPosition(justifyTargets[targetIndex], justifyValue, params.justify);
                }
            } catch (errJustify) {
                return "ERR:" + errMessage(errJustify);
            }
            app.redraw();
            return "OK:" + justifyTargets.length;
        }
        if (action === "applyFontSize") return runRangeAction(ranges, function () {
            for (var rangeIndex = 0; rangeIndex < ranges.length; rangeIndex++) ranges[rangeIndex].characterAttributes.size = params.sizePt;
        });
        // タイプスケール用：選択が触れた段落全体にフォントサイズを適用 / Type scale: apply the size to every touched paragraph in full
        if (action === "applyFontSizePara") { var sizeParas = getSelectedParagraphRanges(); return runRangeAction(sizeParas, function () {
            for (var paraIndex = 0; paraIndex < sizeParas.length; paraIndex++) sizeParas[paraIndex].characterAttributes.size = params.sizePt;
        }); }
        if (action === "applyScale") return runRangeAction(ranges, function () {
            for (var rangeIndex = 0; rangeIndex < ranges.length; rangeIndex++) {
                ranges[rangeIndex].characterAttributes.horizontalScale = params.scale;
                ranges[rangeIndex].characterAttributes.verticalScale = params.scale;
            }
        });
        // サイズと比率を1回でまとめて適用（実サイズ↔見かけトグル用。undo を1エントリに保つ）
        // Apply size and scale together in one pass (for the Actual↔Apparent toggle; keeps undo to one entry)
        if (action === "applyFontSizeAndScale") return runRangeAction(ranges, function () {
            for (var rangeIndex = 0; rangeIndex < ranges.length; rangeIndex++) {
                ranges[rangeIndex].characterAttributes.size = params.sizePt;
                ranges[rangeIndex].characterAttributes.horizontalScale = params.scale;
                ranges[rangeIndex].characterAttributes.verticalScale = params.scale;
            }
        });
        if (action === "applyAlign") {
            var alignSel = app.activeDocument.selection;
            var alignValue = resolveAlignment(params.align);
            var alignCount = 0;
            try {
                if (getTypeName(alignSel) === "TextRange") {
                    alignSel.characterAttributes.alignment = alignValue;
                    alignCount = 1;
                } else if (alignSel && alignSel.length) {
                    for (var itemIndex = 0; itemIndex < alignSel.length; itemIndex++) {
                        alignCount += applyStyleRunAlignment(alignSel[itemIndex], alignValue);
                    }
                }
            } catch (errAlign) {
                return "ERR:" + errMessage(errAlign);
            }
            app.redraw();
            return "OK:" + alignCount;
        }
        if (action === "applyLeading") {
            var leadingFrames = collectLeadingFrames(app.activeDocument.selection);
            if (leadingFrames.length === 0) return "OK:" + readLeadingState(leadingFrames);
            try {
                // 自動行送り量（%）を段落ごとに設定し、常に自動行送りに / Set the auto-leading amount (%) per paragraph; always auto-leading
                applyAutoLeadingToParagraphs(getSelectedParagraphRanges(), params.percent);
                // 行送りの基準（leadingType）はフレーム単位で設定 / The leading basis is set per frame
                if (params.leadingType !== undefined) applyLeadingBasisToFrames(leadingFrames, resolveLeadingType(params.leadingType));
            } catch (errLeading) {
                return "ERR:" + errMessage(errLeading);
            }
            app.redraw();
            return "OK:" + readLeadingState(leadingFrames);
        }
        if (action === "applyProfile") {
            var profSel = app.activeDocument.selection;
            var profParas = getSelectedParagraphRanges(); // 自動カーニング・文字ツメは段落単位 / Kerning & Tsume per paragraph
            try {
                if (params.kern !== undefined) applyKerningToRanges(profParas, resolveAutoKernType(params.kern));
                if (params.tsume !== undefined) applyTsumeToRanges(profParas, params.tsume);
                if (params.tracking !== undefined) applyTrackingToRanges(ranges, params.tracking);
                if (params.sizePt !== undefined) { for (var sizeRangeIndex = 0; sizeRangeIndex < ranges.length; sizeRangeIndex++) ranges[sizeRangeIndex].characterAttributes.size = params.sizePt; }
                if (params.scale !== undefined) {
                    for (var scaleRangeIndex = 0; scaleRangeIndex < ranges.length; scaleRangeIndex++) {
                        ranges[scaleRangeIndex].characterAttributes.horizontalScale = params.scale;
                        ranges[scaleRangeIndex].characterAttributes.verticalScale = params.scale;
                    }
                }
                if (params.align !== undefined) applyAlignmentToSelection(profSel, resolveAlignment(params.align));
                if (params.justify !== undefined) {
                    var profTargets = collectJustifyTargets(profSel);
                    var profJustify = resolveJustification(params.justify);
                    for (var targetIndex = 0; targetIndex < profTargets.length; targetIndex++) {
                        justifyKeepingPosition(profTargets[targetIndex], profJustify, params.justify);
                    }
                }
                if (params.leadingPercent !== undefined) {
                    // 自動行送り量（%）を段落ごとに設定し、基準もあわせて設定 / Set the auto-leading amount (%) per paragraph, plus the basis
                    applyAutoLeadingToParagraphs(getSelectedParagraphRanges(), params.leadingPercent);
                    if (params.leadingType !== undefined) applyLeadingBasisToFrames(collectLeadingFrames(profSel), resolveLeadingType(params.leadingType));
                }
                if (params.mojikumiIndex !== undefined) applyMojikumiToSelection(profSel, params.mojikumiIndex);
                if (params.kinsoku !== undefined) applyKinsokuToSelection(profSel, params.kinsoku);
                if (params.clearAki) clearAkiOnRanges(ranges);
            } catch (errProfile) {
                return "ERR:" + errMessage(errProfile);
            }
            app.redraw();
            return "OK:" + ranges.length;
        }
        if (action === "applyMojikumi") {
            try {
                applyMojikumiToSelection(app.activeDocument.selection, params.mojikumiIndex);
            } catch (errMojikumi) {
                return "ERR:" + errMessage(errMojikumi);
            }
            app.redraw();
            return "OK:1";
        }
        if (action === "applyKinsoku") {
            try {
                applyKinsokuToSelection(app.activeDocument.selection, params.kinsoku);
            } catch (errKinsoku) {
                return "ERR:" + errMessage(errKinsoku);
            }
            app.redraw();
            return "OK:1";
        }
        if (action === "applyPolicy") {
            var policySel = app.activeDocument.selection;
            var policyFrames = collectLeadingFrames(policySel);
            try {
                if (params.align !== undefined) applyAlignmentToSelection(policySel, resolveAlignment(params.align));
                if (params.leadingType !== undefined) applyLeadingBasisToFrames(policyFrames, resolveLeadingType(params.leadingType));
                // ハイフネーションは段落属性（触れた段落全体へ）/ Hyphenation is a paragraph attribute (whole touched paragraphs)
                if (params.hyphenation !== undefined) applyHyphenationToRanges(getSelectedParagraphRanges(), params.hyphenation);
            } catch (errPolicy) {
                return "ERR:" + errMessage(errPolicy);
            }
            app.redraw();
            return "OK:" + policyFrames.length;
        }
        if (action === "getCombos") {
            return "OK:" + encodeURIComponent(collectDocumentCombosRaw(app.activeDocument, params && params.currentArtboardOnly, params && params.includeHidden, params && params.includeLocked));
        }
        if (action === "applyCombo") {
            if (ranges.length === 0) return "OK:nosel";
            var appliedComboCount = 0;
            try {
                appliedComboCount = applyComboToRanges(ranges, params);
                // 文字ツメ・カーニング・プロポーショナルは段落単位（applyKerningMethodToRanges は proportional 非連動）
                // Tsume, kerning method, proportional apply per paragraph (kerning method only, no proportional coupling)
                var comboParas = getSelectedParagraphRanges();
                applyTsumeToRanges(comboParas, params.tsume);
                applyKerningMethodToRanges(comboParas, resolveAutoKernType(params.kernId));
                applyProportionalToRanges(comboParas, params.proportionalMetrics);
            } catch (errCombo) {
                return "ERR:" + errMessage(errCombo);
            }
            app.redraw();
            return "OK:" + appliedComboCount;
        }
        if (action === "selectMatching") {
            var matchedCount = selectFramesMatchingCombo(app.activeDocument, params, params && params.currentArtboardOnly, params && params.includeHidden, params && params.includeLocked);
            app.redraw();
            return "OK:" + matchedCount;
        }
        return "OK:0";
    }
    var DISPATCH_SRC = "(" + dispatch.toString() + ")";

    /* パラメータを安全な JS リテラル文字列へ / Serialize params to a safe JS literal */
    function paramsToSource(params) {
        if (!params) return "{}";
        var parts = [];
        if (params.method !== undefined) parts.push('method:decodeURIComponent("' + encodeURIComponent(params.method) + '")');
        if (params.value !== undefined) parts.push("value:" + parseInt(params.value, 10));
        if (params.tracking !== undefined) parts.push("tracking:" + parseInt(params.tracking, 10));
        if (params.sizePt !== undefined) parts.push("sizePt:" + params.sizePt);
        if (params.scale !== undefined) parts.push("scale:" + params.scale);
        if (params.align !== undefined) parts.push('align:decodeURIComponent("' + encodeURIComponent(params.align) + '")');
        if (params.justify !== undefined) parts.push('justify:decodeURIComponent("' + encodeURIComponent(params.justify) + '")');
        if (params.psName !== undefined) parts.push('psName:decodeURIComponent("' + encodeURIComponent(params.psName) + '")');
        if (params.psList !== undefined) parts.push('psList:decodeURIComponent("' + encodeURIComponent(params.psList) + '")');
        if (params.kern !== undefined) parts.push('kern:decodeURIComponent("' + encodeURIComponent(params.kern) + '")');
        if (params.tsume !== undefined) parts.push("tsume:" + parseInt(params.tsume, 10));
        // 行送り：自動行送り量（%）／プロファイル用の行送り%（いずれも数値）/ Leading: auto-leading amount (%) and the profile leading % (both numeric)
        if (params.percent !== undefined) parts.push("percent:" + params.percent);
        if (params.leadingPercent !== undefined) parts.push("leadingPercent:" + params.leadingPercent);
        if (params.leadingType !== undefined) parts.push('leadingType:"' + params.leadingType + '"');
        if (params.hyphenation !== undefined) parts.push("hyphenation:" + (params.hyphenation ? "true" : "false"));
        if (params.mojikumiIndex !== undefined) parts.push("mojikumiIndex:" + parseInt(params.mojikumiIndex, 10));
        if (params.kinsoku !== undefined) parts.push('kinsoku:decodeURIComponent("' + encodeURIComponent(params.kinsoku) + '")');
        if (params.clearAki !== undefined) parts.push("clearAki:" + (params.clearAki ? "true" : "false"));
        // 使用フォント一覧（combo 適用・走査）用フィールド / Fields for the used-fonts list (combo apply / scan)
        if (params.size !== undefined) parts.push("size:" + (isNaN(params.size) ? "NaN" : params.size));
        if (params.leading !== undefined) parts.push("leading:" + (isNaN(params.leading) ? "NaN" : params.leading));
        if (params.autoLeading !== undefined) parts.push("autoLeading:" + (params.autoLeading ? "true" : "false"));
        if (params.kernId !== undefined) parts.push('kernId:decodeURIComponent("' + encodeURIComponent(params.kernId) + '")');
        if (params.proportionalMetrics !== undefined) parts.push("proportionalMetrics:" + (params.proportionalMetrics ? "true" : "false"));
        if (params.currentArtboardOnly !== undefined) parts.push("currentArtboardOnly:" + (params.currentArtboardOnly ? "true" : "false"));
        if (params.includeHidden !== undefined) parts.push("includeHidden:" + (params.includeHidden ? "true" : "false"));
        if (params.includeLocked !== undefined) parts.push("includeLocked:" + (params.includeLocked ? "true" : "false"));
        if (params.matchFontOnly !== undefined) parts.push("matchFontOnly:" + (params.matchFontOnly ? "true" : "false"));
        return "{" + parts.join(",") + "}";
    }

    /*
     * 本文をメインエンジンへ送り、結果マーカーを解析して onDone(status, payload) を呼ぶ。
     * BridgeTalk は本文送信時にバックスラッシュをエスケープするため（"\r" → ターゲットで "\\r"）、
     * コード全体を encodeURIComponent で包んで送り（%エンコードに \ は出ない）、
     * ターゲットで decodeURIComponent + eval して元ソースに復元してから実行する。
     * Send the body to the main engine and parse the result marker.
     */
    function sendWorker(code, onDone) {
        var bridge = new BridgeTalk();
        bridge.target = "illustrator";
        bridge.body = "eval(decodeURIComponent(\"" + encodeURIComponent(code) + "\"));";
        bridge.onResult = function (response) {
            var payload = response.body || "";
            var colonIndex = payload.indexOf(":");
            var marker = colonIndex >= 0 ? payload.substring(0, colonIndex) : payload;
            var rest = colonIndex >= 0 ? payload.substring(colonIndex + 1) : "";
            if (marker === "OK") onDone("ok", rest);
            else onDone("error", rest);
        };
        bridge.onError = function (response) { onDone("error", response && response.body ? response.body : "BridgeTalk error"); };
        bridge.send();
    }

    /* メインエンジンへアクションを委譲（非同期）/ Delegate an action to the main engine (async) */
    function runWorker(actionId, params, onDone) {
        var code = getWorkerLibSource() + "\nvar __r=" + DISPATCH_SRC + "(\"" + actionId + "\"," + paramsToSource(params) + ");__r;";
        sendWorker(code, onDone);
    }

    // =========================================
    // 使用フォント一覧：ラベル・解析（パレット側）/ Used-fonts list: labels & parsing (palette side)
    // =========================================

    /* カーニング method id を表示ラベルへ / Kerning method id to a display label */
    function usedFontKernLabel(kernMethodId) {
        var labelEntry = LABELS.autoKern[kernMethodId];
        return labelEntry ? getLocalizedText(labelEntry) : "—";
    }

    /* 各列のセル文字列を組み立て / Build the per-column cell strings
       [使用数, フォント, サイズ, 行送り, 自動カーニング, 文字ツメ, トラッキング, プロポーショナル] */
    function comboCells(combo) {
        var countText = isNaN(combo.count) ? "?" : String(combo.count);
        var sizeText = formatNumber(combo.size) + "pt";
        var leadingText = combo.autoLeading
            ? getLocalizedText(LABELS.usedFonts.autoLeading)
            : formatNumber(combo.leading) + "pt";
        var tsumeText = String(Math.round(combo.tsume));
        var trackingText = isNaN(combo.tracking) ? "?" : String(Math.round(combo.tracking));
        var propMetricsText = combo.proportionalMetrics ? "ON" : "OFF";
        return [countText, combo.displayName, sizeText, leadingText, usedFontKernLabel(combo.kernId), tsumeText, trackingText, propMetricsText];
    }

    /* 委譲結果（デコード済み）を組み合わせ配列に復元し、表示順に並べる / Parse the worker payload into combos and sort them */
    function parseCombosPayload(rawPayload) {
        var combos = [];
        if (!rawPayload) return combos;
        var TAB = String.fromCharCode(9), LF = String.fromCharCode(10);
        var comboRows = rawPayload.split(LF);
        for (var i = 0; i < comboRows.length; i++) {
            var fields = comboRows[i].split(TAB);
            if (fields.length < 10) continue;
            var combo = {
                psName: fields[0],
                displayName: fields[1],
                size: parseFloat(fields[2]),
                leading: parseFloat(fields[3]),
                autoLeading: fields[4] === "1",
                tsume: parseFloat(fields[5]),
                tracking: parseFloat(fields[6]),
                kernId: fields[7],
                proportionalMetrics: fields[8] === "1",
                count: parseInt(fields[9], 10)
            };
            combo.cells = comboCells(combo);
            combos.push(combo);
        }
        // フォント名→サイズ→行送り→自動カーニング→ツメ→トラッキング→プロポーショナル の順 / Sort by font, size, leading, kerning, Tsume, tracking, prop-metrics
        combos.sort(function (a, b) {
            if (a.displayName !== b.displayName) return a.displayName < b.displayName ? -1 : 1;
            if ((a.size || 0) !== (b.size || 0)) return (a.size || 0) - (b.size || 0);
            if ((a.leading || 0) !== (b.leading || 0)) return (a.leading || 0) - (b.leading || 0);
            if (a.kernId !== b.kernId) return a.kernId < b.kernId ? -1 : 1;
            if ((a.tsume || 0) !== (b.tsume || 0)) return (a.tsume || 0) - (b.tsume || 0);
            if ((a.tracking || 0) !== (b.tracking || 0)) return (a.tracking || 0) - (b.tracking || 0);
            if (a.proportionalMetrics !== b.proportionalMetrics) return a.proportionalMetrics ? 1 : -1;
            return 0;
        });
        return combos;
    }

    /* 組み合わせを適用パラメータへ / Convert a combo to apply params */
    function comboToParams(combo) {
        return {
            psName: combo.psName,
            size: combo.size,
            leading: combo.leading,
            autoLeading: combo.autoLeading,
            tsume: combo.tsume,
            tracking: combo.tracking,
            kernId: combo.kernId,
            proportionalMetrics: combo.proportionalMetrics
        };
    }

    /* 組み合わせ一覧をフォント単位に畳む（フォントのみ表示用。psName で重複排除、表示名で整列）
       Collapse the combos to unique fonts (for the font-only view; dedup by psName, sorted by display name) */
    function combosToFonts(combos) {
        var seen = {}, fonts = [];
        for (var i = 0; i < combos.length; i++) {
            var combo = combos[i];
            if (seen[combo.psName]) continue;
            seen[combo.psName] = true;
            fonts.push({ psName: combo.psName, displayName: combo.displayName });
        }
        fonts.sort(function (a, b) { return a.displayName < b.displayName ? -1 : (a.displayName > b.displayName ? 1 : 0); });
        return fonts;
    }

    // =========================================
    // UI構築 / Build UI
    // =========================================

    // =========================================
    // 行送り用の単位・選択肢ヘルパー / Leading: units and choices (palette side)
    // =========================================

    /* パネルの余白と間隔 / Panel margins and spacing */
    var PANEL_MARGINS = [10, 15, 10, 10];
    var PANEL_SPACING = 8;

    /* パネルの共通設定 / Apply shared panel layout */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* グループの共通設定（row/column で整列を切り替え）/ Apply shared group layout */
    function setupGroup(group, orientation, spacing) {
        var groupOrientation = orientation || "column";
        group.orientation = groupOrientation;
        group.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
        group.alignment = "fill";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* タブを1枚追加して共通の体裁（縦積み・上左マージン15）を適用 / Add one tab with the shared layout (column, 15px top/left margins) */
    function addTab(mainTabs, labelEntry) {
        var tab = mainTabs.add("tab", undefined, getLocalizedText(labelEntry));
        tab.orientation = "column";
        tab.alignChildren = ["fill", "top"];
        tab.margins.top = 15; // タブ内上部にマージン / Top margin inside the tab
        tab.margins.left = 15; // タブ内左にマージン / Left margin inside the tab
        return tab;
    }

    /* 少し低くするボタンの共通の高さ / Shared height for the "slightly shorter" buttons */
    var COMPACT_BUTTON_HEIGHT = 20;

    /* ボタンの高さを指定 px 詰める（レイアウト確定後に1回だけ呼ぶ）/ Trim a button's height by the given px (call once, after layout) */
    function trimButtonHeight(button, px) {
        try {
            button.size = [button.size.width, button.size.height - px];
        } catch (e) {}
    }

    /* ボタンを共通の高さ（COMPACT_BUTTON_HEIGHT）に揃える / Fit a button to the shared compact height */
    function fitButtonHeight(button) {
        if (!button) return;
        try { trimButtonHeight(button, button.size.height - COMPACT_BUTTON_HEIGHT); } catch (e) {}
    }

    /* 高さを詰めるボタンをまとめて共通高さに揃える（表示＝レイアウト確定後に呼ぶ）
       Fit every height-trimmed button to the shared compact height (call after show, once layout is final) */
    function fitAllButtonHeights(ui) {
        var buttons = [
            ui.refreshCombosButton, ui.selectMatchingButton,
            ui.exportPresetButton, ui.addPresetButton, ui.overwritePresetButton, ui.deletePresetButton,
            ui.apparentToggleButton
        ];
        for (var i = 0; i < buttons.length; i++) fitButtonHeight(buttons[i]);
    }

    var UNIT_MAP = { 0: "in", 1: "mm", 2: "pt", 3: "pica", 4: "cm", 6: "px", 7: "ft/in", 8: "m", 9: "yd", 10: "ft" };
    var UNIT_TO_PT = {
        0: 72, 1: 2.8346456692913386, 2: 1, 3: 12, 4: 28.346456692913386,
        5: 0.7086614173228346, 6: 1, 7: 72, 8: 2834.6456692913386, 9: 2592, 10: 864
    };

    function getUnitLabel(unitCode) {
        if (unitCode === 5) return "Q";
        return UNIT_MAP[unitCode] || "pt";
    }

    /* 行送りの単位ラベル（Q/H 環境では行送りは「H」）/ Leading unit label (leading uses "H" when the unit is Q/H) */
    function getLeadingUnitLabel(unitCode) {
        if (unitCode === 5) return "H";
        return getUnitLabel(unitCode);
    }

    /* テキスト単位を取得（コード・ラベル・pt 換算係数）/ Resolve the text unit (code, label, pt factor) */
    function getTextUnit() {
        var unitCode = 2;
        try { unitCode = app.preferences.getIntegerPreference("text/units"); } catch (e) { }
        return { code: unitCode, label: getUnitLabel(unitCode), factor: UNIT_TO_PT[unitCode] || 1 };
    }

    /* タイプスケールの倍率（音程比）/ Type-scale ratios (musical intervals) */
    var TYPE_SCALE_RATIOS = [
        { label: "Minor Second 1.067", value: 1.067 },
        { label: "Major Second 1.125", value: 1.125 },
        { label: "Minor Third 1.2", value: 1.2 },
        { label: "Major Third 1.25", value: 1.25 },
        { label: "Golden Ratio: ½ 1.309", value: 1.309 },
        { label: "Perfect Fourth 1.333", value: 1.333 },
        { label: "Augmented Fourth 1.414", value: 1.414 },
        { label: "Golden Ratio 1.618", value: 1.618 }
    ];
    var TYPE_SCALE_DEFAULT_RATIO_INDEX = 3; // Major Third 1.25

    /* 基準サイズと倍率からタイプスケール（7段）を生成 / Generate a 7-step type scale from a base size and ratio
       index 2 が基準サイズ（基準を段番号 0、最上段を -2 とする）/ index 2 is the base (step 0; top row is -2) */
    function generateTypeScaleSizes(baseSize, ratio) {
        var sizes = [];
        var down = baseSize;
        for (var i = 0; i < 2; i++) down /= ratio;
        for (var j = 0; j < 7; j++) {
            sizes.push(Math.round(down * 10) / 10);
            down *= ratio;
        }
        return sizes;
    }

    /* 文字組みアキ量設定の選択肢（index は mojikumiSet のインデックス、-1=なし）
       Mojikumi choices (index is the mojikumiSet index; -1 = None) */
    var MOJIKUMI_CHOICES = [
        { index: -1, label: getLocalizedText(LABELS.mojikumi.none) },
        { index: 0, label: getLocalizedText(LABELS.mojikumi.lineEndFullHalf) },
        { index: 1, label: getLocalizedText(LABELS.mojikumi.punctHalf) },
        { index: 2, label: getLocalizedText(LABELS.mojikumi.lineEndHalf) },
        { index: 3, label: getLocalizedText(LABELS.mojikumi.lineEndFull) },
        { index: 4, label: getLocalizedText(LABELS.mojikumi.punctFull) },
        { index: 5, label: getLocalizedText(LABELS.mojikumi.tight) },
        { index: 6, label: getLocalizedText(LABELS.mojikumi.solid) }
    ];

    /* 禁則の選択肢（id は paragraphAttributes.kinsoku に渡す値）
       Kinsoku choices (id is the value passed to paragraphAttributes.kinsoku) */
    var KINSOKU_CHOICES = [
        { id: "None", label: getLocalizedText(LABELS.kinsoku.none) },
        { id: "Hard", label: getLocalizedText(LABELS.kinsoku.hard) },
        { id: "Soft", label: getLocalizedText(LABELS.kinsoku.soft) },
        { id: "Soft_v2", label: getLocalizedText(LABELS.kinsoku.softV2) }
    ];

    /* edittext を ↑↓ キーで増減（Shift=10 / Alt=0.1）/ Step an edittext value with Up/Down keys (Shift=10, Alt=0.1) */
    function changeValueByArrowKey(editText, allowNegative, onUpdate, decimals) {
        function roundToStep(value, step) { return Math.round(value / step) * step; }
        editText.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;
            var fieldValue = Number(editText.text);
            if (isNaN(fieldValue)) return;
            var keyboard = ScriptUI.environment.keyboardState;
            var step = 1;
            if (keyboard.shiftKey) step = 10;
            else if (keyboard.altKey) step = 0.1;
            if (keyboard.shiftKey) fieldValue = roundToStep(fieldValue, step);
            if (event.keyName === "Up") fieldValue += step;
            else fieldValue -= step;
            if (keyboard.altKey) fieldValue = Math.round(fieldValue * 10) / 10;
            else fieldValue = Math.round(fieldValue);
            if (!allowNegative && fieldValue < 0) fieldValue = 0;
            event.preventDefault();
            if (typeof decimals === "number" && decimals >= 0) {
                var decimalFactor = Math.pow(10, decimals);
                fieldValue = Math.round(fieldValue * decimalFactor) / decimalFactor;
                editText.text = fieldValue.toFixed(decimals);
            } else {
                editText.text = String(fieldValue);
            }
            if (typeof onUpdate === "function") onUpdate();
        });
    }

    /* 行送りの基準の選択肢 / Leading-basis choices */
    function getLeadingTypeChoices() {
        return [
            { id: "top", label: getLocalizedText(LABELS.leading.typeTop) },
            { id: "baseline", label: getLocalizedText(LABELS.leading.typeBaseline) }
        ];
    }

    /* 自動カーニングの選択肢を生成 / Build the auto-kerning option list */
    function createAutoKernOptions() {
        return [
            { id: "mono", label: LABELS.autoKern.mono },
            { id: "zero", label: LABELS.autoKern.zero },
            { id: "metrics", label: LABELS.autoKern.metrics },
            { id: "optical", label: LABELS.autoKern.optical }
        ];
    }

    /* プリセット（PostScript 名 ＋ 文字組み設定）。選択時にまとめて適用する
       Presets (PostScript name + type-composition settings); all are applied together on selection
       kern: 自動カーニング method id（"metrics" 等）/ tsume: 文字ツメ%（0〜100）/ tracking: トラッキング（-500〜500）*/
    var DEFAULT_PRESETS = [
        { psName: "HiraginoSans-W3", kern: "metrics", tsume: 0, tracking: 0 }
    ];

    // =========================================
    // プリセットの保存・読み込み（JSON）/ Preset persistence (JSON)
    // ExtendScript には JSON が無いため、保存は手動シリアライズ、
    // 読み込みは自前ファイルなので eval で復元する（ES3 安全）。
    // ExtendScript has no JSON, so we serialize by hand and restore our own
    // file via eval (ES3-safe).
    // =========================================

    /* プリセット保存先の File / The File where presets are stored */
    function getPresetFile() {
        return File(Folder.userData.fsName + "/UnifiedTypePanel_presets.json");
    }

    /* 文字列を JSON 文字列リテラルへ / Quote a string as a JSON string literal */
    function jsonQuote(str) {
        var escaped = String(str);
        escaped = escaped.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
        escaped = escaped.replace(/\r/g, "\\r").replace(/\n/g, "\\n").replace(/\t/g, "\\t");
        return '"' + escaped + '"';
    }

    /* プリセット配列を JSON テキストへ / Serialize the preset array to JSON text */
    function presetsToJsonText(presets) {
        var jsonLines = [];
        for (var i = 0; i < presets.length; i++) {
            var preset = presets[i];
            var fields = [];
            fields.push('"psName":' + jsonQuote(preset.psName));
            if (preset.kern !== undefined && preset.kern !== null) fields.push('"kern":' + jsonQuote(preset.kern));
            if (preset.tsume !== undefined && preset.tsume !== null) fields.push('"tsume":' + parseInt(preset.tsume, 10));
            if (preset.tracking !== undefined && preset.tracking !== null) fields.push('"tracking":' + parseInt(preset.tracking, 10));
            jsonLines.push("  {" + fields.join(", ") + "}");
        }
        return "[\n" + jsonLines.join(",\n") + "\n]\n";
    }

    /* プリセットをファイルへ保存 / Save presets to the file */
    function savePresets(presets) {
        var file = getPresetFile();
        try {
            file.encoding = "UTF-8";
            if (file.open("w")) {
                file.write(presetsToJsonText(presets));
                file.close();
                return true;
            }
        } catch (e) { }
        return false;
    }

    /* パース結果がプリセット配列として妥当か / Whether a parsed value is a valid preset array
       各要素が psName（文字列）を持つ配列であることを要求 / Requires an array whose items each have a string psName */
    function isValidPresetArray(parsed) {
        if (!parsed || typeof parsed.length !== "number" || parsed.length === 0) return false;
        for (var i = 0; i < parsed.length; i++) {
            var entry = parsed[i];
            if (!entry || typeof entry.psName !== "string" || entry.psName === "") return false;
        }
        return true;
    }

    /* 壊れた設定ファイルを .bak に退避（上書きせず原因調査できるように）
       Move a corrupt preset file aside to .bak (don't overwrite, so it can be inspected) */
    function backupCorruptPresetFile(file) {
        try {
            var backup = new File(file.fsName + ".bak");
            if (backup.exists) backup.remove();
            file.copy(backup);
        } catch (eBackup) { }
    }

    /* プリセットをファイルから読み込み（無ければ既定）/ Load presets (fall back to defaults)
       復旧方針を明示: (1) 読み込み失敗 → 既定 / (2) 空ファイル → 既定 /
       (3) パース失敗・内容不正 → 破損ファイルを .bak に退避してから既定へ復帰
       Recovery is explicit: (1) read error → defaults / (2) empty file → defaults /
       (3) parse failure or invalid shape → move the file to .bak, then defaults */
    function loadPresets() {
        var file = getPresetFile();
        if (!file.exists) return DEFAULT_PRESETS;

        /* 読み込み / Read */
        var fileText = "";
        try {
            file.encoding = "UTF-8";
            if (!file.open("r")) return DEFAULT_PRESETS;
            fileText = file.read();
            file.close();
        } catch (eRead) {
            try { file.close(); } catch (eClose) { }
            return DEFAULT_PRESETS;
        }

        /* 空ファイルは既定へ（破損扱いにしない）/ Empty file → defaults (not treated as corrupt) */
        if (!fileText || fileText.replace(/^\s+|\s+$/g, "") === "") return DEFAULT_PRESETS;

        /* パース（自前ファイルなので eval を許容）/ Parse (eval is acceptable for our own file) */
        var parsed = null;
        try {
            parsed = eval("(" + fileText + ")");
        } catch (eParse) {
            parsed = null;
        }

        /* パース失敗・内容不正は破損として .bak 退避のうえ既定へ / Parse failure or invalid shape → back up and use defaults */
        if (!isValidPresetArray(parsed)) {
            backupCorruptPresetFile(file);
            return DEFAULT_PRESETS;
        }
        return parsed;
    }

    /* プリセット配列でリストボックスを再構築（表示名は後で和文名に差し替え）
       Rebuild a listbox from a preset array (display names resolved to Japanese later) */
    /* プリセットをフォントのみ／詳細の両 listbox へ流し込む / Fill both preset listboxes (font-only + detailed)
       詳細列: フォント / 自動カーニング / 文字ツメ / プロポーショナル(kern連動) / トラッキング
       Detailed columns: font / auto-kerning / Tsume / proportional (follows kern) / tracking */
    function fillPresetLists(fontOnlyLb, detailedLb, presets) {
        fontOnlyLb.removeAll();
        detailedLb.removeAll();
        for (var i = 0; i < presets.length; i++) {
            var preset = presets[i];
            // フォントのみ（単一列）/ Font-only (single column)
            var foItem = fontOnlyLb.add("item", preset.psName);
            foItem.psName = preset.psName;
            if (preset.kern !== undefined) foItem.kern = preset.kern;
            if (preset.tsume !== undefined) foItem.tsume = preset.tsume;
            if (preset.tracking !== undefined) foItem.tracking = preset.tracking;
            // 詳細（フォント＋4列）/ Detailed (font + 4 columns)
            var dItem = detailedLb.add("item", preset.psName);
            dItem.subItems[0].text = usedFontKernLabel(preset.kern);
            dItem.subItems[1].text = (preset.tsume !== undefined && preset.tsume !== null) ? String(preset.tsume) : "";
            dItem.subItems[2].text = (preset.kern === "metrics") ? "ON" : "OFF"; // プロポーショナルは kern 連動 / Proportional follows the kerning method
            dItem.subItems[3].text = (preset.tracking !== undefined && preset.tracking !== null) ? String(preset.tracking) : "";
            dItem.psName = preset.psName;
            if (preset.kern !== undefined) dItem.kern = preset.kern;
            if (preset.tsume !== undefined) dItem.tsume = preset.tsume;
            if (preset.tracking !== undefined) dItem.tracking = preset.tracking;
        }
    }

    /* 文字揃えの選択肢を生成 / Build the character-alignment option list */
    function createAlignOptions() {
        return [
            { id: "roman", label: LABELS.align.roman },
            { id: "top", label: LABELS.align.top },
            { id: "center", label: LABELS.align.center },
            { id: "bottom", label: LABELS.align.bottom },
            { id: "icftop", label: LABELS.align.icfTop },
            { id: "icfbottom", label: LABELS.align.icfBottom }
        ];
    }

    /* 行揃えの選択肢を生成 / Build the paragraph-justification option list */
    function createJustifyOptions() {
        return [
            { id: "left", label: LABELS.justify.left, icon: "left" },
            { id: "center", label: LABELS.justify.center, icon: "center" },
            { id: "right", label: LABELS.justify.right, icon: "right" },
            { id: "justifyLeft", label: LABELS.justify.justifyLeft, icon: "justifyLeft" },
            { id: "justifyAll", label: LABELS.justify.justifyAll, icon: "justifyAll" }
        ];
    }

    // =========================================
    // 行揃えアイコンの描画（Keep_TextPosition.jsx 由来）/ Justification icon drawing (ported from Keep_TextPosition.jsx)
    // ライト／ダーク対応は ArtboardNavigator.jsx の方式 / Light/dark handling follows ArtboardNavigator.jsx
    // =========================================

    /* 環境設定のUI明るさが明るい側か / Whether the UI brightness preference is on the light side
       uiBrightness は 0（最暗）〜1（最明）。0.5（やや暗め）は暗い側に含めるため 0.5 超で判定 */
    function isLightUI() {
        try {
            return app.preferences.getRealPreference("uiBrightness") > 0.5;
        } catch (e) {
            return false;
        }
    }

    /* テーマ＋アクティブ状態に応じたボタン配色 / Button colors per theme and active state */
    function getJustifyColors(isLight, isActive) {
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
        if (iconType === "justifyLeft" || iconType === "justifyCenter" || iconType === "justifyRight") {
            return [longWidth, longWidth, longWidth, shortWidth];
        }
        return [longWidth, shortWidth, longWidth, shortWidth];
    }

    /* 行の開始 X（アイコン種別ごとに左／中央／右）/ Line start X per icon type (left/center/right) */
    function getJustifyLineX(iconType, buttonWidth, lineWidth) {
        var margin = 5;
        if (iconType === "right" || iconType === "justifyRight") return buttonWidth - margin - lineWidth;
        if (iconType === "center" || iconType === "justifyCenter") return Math.round((buttonWidth - lineWidth) / 2);
        return margin;
    }

    /* アイコンの罫線を描く / Draw the icon's lines */
    function drawJustifyIconLines(graphics, iconType, buttonWidth, buttonHeight, lineColor) {
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

    /* ボタン背景＋アイコンを描く（テーマ＋アクティブ対応）/ Draw the button background + icon (theme + active aware) */
    function drawJustifyIcon(button, isActive, isLight) {
        var graphics = button.graphics;
        var buttonWidth = button.size[0];
        var buttonHeight = button.size[1];
        var colors = getJustifyColors(isLight, isActive);
        try {
            graphics.rectPath(0, 0, buttonWidth, buttonHeight);
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, colors.bg));
            if (colors.border) {
                graphics.rectPath(0, 0, buttonWidth, buttonHeight);
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, colors.border, 1));
            }
        } catch (e1) {
            try { graphics.drawOSControl(); } catch (e2) { }
        }
        drawJustifyIconLines(graphics, button.iconType, buttonWidth, buttonHeight, colors.line);
    }

    /* 使用フォントタブの中身を構築して参照を返す（DocumentFontListSelector 由来）
       Build the used-fonts tab contents and return refs (ported from DocumentFontListSelector) */
    function buildUsedFontsTab(usedFontsTab) {
        var usedFontColumnTitles = [
            getLocalizedText(LABELS.usedFonts.column.count),
            getLocalizedText(LABELS.usedFonts.column.font),
            getLocalizedText(LABELS.usedFonts.column.size),
            getLocalizedText(LABELS.usedFonts.column.leading),
            getLocalizedText(LABELS.usedFonts.column.kern),
            getLocalizedText(LABELS.usedFonts.column.tsume),
            getLocalizedText(LABELS.usedFonts.column.tracking),
            getLocalizedText(LABELS.usedFonts.column.propMetrics)
        ];
        var usedFontColumnWidths = [30, 88, 34, 40, 48, 36, 36, 38]; // 合計350px（パレット幅を抑えるため圧縮）/ Sums to 350px (compressed to keep the palette narrow)

        // 表示モード：フォントのみ／詳細（ドキュメントフォント一覧を兼ねる）/ View mode: font-only / detailed (also serves as the document-font list)
        var viewModeRow = usedFontsTab.add("group");
        viewModeRow.orientation = "row";
        viewModeRow.alignment = ["left", "top"];
        viewModeRow.spacing = 12;
        viewModeRow.margins = [0, 0, 0, 6];
        var fontOnlyRadio = viewModeRow.add("radiobutton", undefined, getLocalizedText(LABELS.usedFonts.view.fontOnly));
        var detailedRadio = viewModeRow.add("radiobutton", undefined, getLocalizedText(LABELS.usedFonts.view.detailed));
        fontOnlyRadio.value = true; // 既定はフォントのみ（フォント適用を手早く）/ Default: font-only (quick font apply)

        // 走査オプションのチェックボックスは「フィルター」パネルとしてタブ最下部に配置（下記）
        // The scan-option checkboxes live in a titled "Filter" panel at the bottom of the tab (see below)

        // 両方の listbox を最初に生成し、orientation="stack" で同じ位置に重ねる。
        // stack のグループ高さは「子の最大値」なので2つ並べても高さは伸びない。
        // かつ表示前生成なので onChange が確実に発火する（表示後に動的生成した listbox は
        // クリック（onChange）が効かない環境があるため、この方式にする）。
        // Both listboxes are created up front and stacked (orientation="stack") in the same cell.
        // A stack group's height is the MAX of its children, so two don't add up — the panel stays compact.
        // Being created before show, their onChange fires reliably (dynamically-created listboxes may not).
        var usedFontListHost = usedFontsTab.add("group");
        usedFontListHost.orientation = "stack";
        usedFontListHost.alignChildren = ["fill", "fill"];
        usedFontListHost.alignment = ["fill", "top"];

        // フォントのみ表示：単一列のフォント名リスト / Font-only view: single-column font list
        var fontOnlyListBox = usedFontListHost.add("listbox", undefined, [], { multiselect: false });
        fontOnlyListBox.preferredSize = [300, 230];

        // 詳細表示：8列の組み合わせリスト / Detailed view: the 8-column combination list
        var comboListBox = usedFontListHost.add("listbox", undefined, [], {
            multiselect: false,
            numberOfColumns: usedFontColumnTitles.length,
            showHeaders: true,
            columnTitles: usedFontColumnTitles,
            columnWidths: usedFontColumnWidths
        });
        comboListBox.preferredSize = [300, 230];

        // 既定はフォントのみ（詳細は隠す）/ Default: font-only (hide the detailed list)
        comboListBox.visible = false;

        // ボタンエリア：右＝更新・条件一致選択（クリック適用は常時ON・UIなし）
        // Button area: right = refresh, select matching (apply-on-click always ON, no UI)
        var usedFontButtonRow = usedFontsTab.add("group");
        usedFontButtonRow.orientation = "row";
        usedFontButtonRow.alignment = ["fill", "top"];
        usedFontButtonRow.margins = [0, 10, 0, 10]; // リストとボタンの間隔 / Gap above the buttons

        var refreshCombosButton = usedFontButtonRow.add("button", undefined, getLocalizedText(LABELS.usedFonts.button.refresh));
        refreshCombosButton.alignment = ["right", "center"];
        var selectMatchingButton = usedFontButtonRow.add("button", undefined, getLocalizedText(LABELS.usedFonts.button.selectMatching));
        selectMatchingButton.alignment = ["right", "center"];
        // 高さは表示後に trimButtonHeight でまとめて詰める / Heights are trimmed together after show via trimButtonHeight

        // フィルター：走査オプションのチェックボックスをまとめてタブ最下部に配置
        // Filter: scan-option checkboxes grouped in a titled panel at the bottom of the tab
        var filterPanel = usedFontsTab.add("panel", undefined, getLocalizedText(LABELS.field.filter));
        setupPanel(filterPanel, 4);
        filterPanel.alignChildren = ["left", "top"];
        filterPanel.alignment = ["fill", "top"];
        filterPanel.margins = [10, 12, 10, 8];

        // 走査オプションのチェックボックス（key で参照を引けるよう object に格納）/ Scan-option checkboxes (kept in an object, addressable by key)
        var checkboxSpecs = [
            { key: "includeLockedCheckbox", label: LABELS.usedFonts.control.includeLocked, initial: true },
            { key: "includeHiddenCheckbox", label: LABELS.usedFonts.control.includeHidden, initial: true },
            { key: "currentArtboardCheckbox", label: LABELS.usedFonts.control.currentArtboardOnly, initial: false }
        ];
        var scanCheckboxes = {};
        for (var cbIndex = 0; cbIndex < checkboxSpecs.length; cbIndex++) {
            var cbSpec = checkboxSpecs[cbIndex];
            var scanCheckbox = filterPanel.add("checkbox", undefined, getLocalizedText(cbSpec.label));
            scanCheckbox.value = cbSpec.initial;
            scanCheckboxes[cbSpec.key] = scanCheckbox;
        }

        return {
            comboListBox: comboListBox,
            fontOnlyListBox: fontOnlyListBox,
            fontOnlyRadio: fontOnlyRadio,
            detailedRadio: detailedRadio,
            includeLockedCheckbox: scanCheckboxes.includeLockedCheckbox,
            includeHiddenCheckbox: scanCheckboxes.includeHiddenCheckbox,
            currentArtboardCheckbox: scanCheckboxes.currentArtboardCheckbox,
            refreshCombosButton: refreshCombosButton,
            selectMatchingButton: selectMatchingButton
        };
    }

    /* 左カラム（ドキュメントフォント／プリセット）を構築して参照を返す
       Build the left column (document fonts / presets) and return refs */
    function buildFontsTab(fontsTab) {
        var fontsColumn = fontsTab.add("group");
        fontsColumn.orientation = "column";
        fontsColumn.alignChildren = ["fill", "top"];
        fontsColumn.spacing = 8;

        // ドキュメントフォント一覧は「使用フォント」タブが兼ねる。プリセットはパネル枠なしで直接配置
        // The document-font list lives in the "Used fonts" tab; presets sit directly with no panel frame

        // 表示モード：フォントのみ／詳細 / View mode: font-only / detailed
        var presetViewRow = fontsColumn.add("group");
        presetViewRow.orientation = "row";
        presetViewRow.alignment = ["left", "top"];
        presetViewRow.spacing = 12;
        var presetFontOnlyRadio = presetViewRow.add("radiobutton", undefined, getLocalizedText(LABELS.usedFonts.view.fontOnly));
        var presetDetailedRadio = presetViewRow.add("radiobutton", undefined, getLocalizedText(LABELS.usedFonts.view.detailed));
        presetFontOnlyRadio.value = true; // 既定はフォントのみ / Default: font-only

        // 2つの listbox を stack で重ねる（フォントのみ＝単一列／詳細＝5列）。高さは最大値なので伸びない
        // Two stacked listboxes (font-only = single column / detailed = 5 columns); stack height = max, so no growth
        var presetListHost = fontsColumn.add("group");
        presetListHost.orientation = "stack";
        presetListHost.alignChildren = ["fill", "fill"];
        presetListHost.alignment = ["fill", "top"];

        // フォントのみ：単一列 / Font-only: single column
        var presetFontOnlyListBox = presetListHost.add("listbox", undefined, [], { multiselect: false });
        presetFontOnlyListBox.preferredSize = [300, 350];
        presetFontOnlyListBox.helpTip = getLocalizedText(LABELS.tip.presets);

        // 詳細：フォント／自動カーニング／文字ツメ／プロポーショナル／トラッキング
        // Detailed: font / auto-kerning / Tsume / proportional / tracking
        var presetDetailedListBox = presetListHost.add("listbox", undefined, [], {
            multiselect: false,
            numberOfColumns: 5,
            showHeaders: true,
            columnTitles: [
                getLocalizedText(LABELS.usedFonts.column.font),
                getLocalizedText(LABELS.usedFonts.column.kern),
                getLocalizedText(LABELS.usedFonts.column.tsume),
                getLocalizedText(LABELS.usedFonts.column.propMetrics),
                getLocalizedText(LABELS.usedFonts.column.tracking)
            ],
            columnWidths: [128, 66, 44, 46, 44]
        });
        presetDetailedListBox.preferredSize = [300, 350];
        presetDetailedListBox.helpTip = getLocalizedText(LABELS.tip.presets);
        presetDetailedListBox.visible = false; // 既定はフォントのみ / Default: font-only

        fillPresetLists(presetFontOnlyListBox, presetDetailedListBox, loadPresets());

        // プリセット操作ボタンを3カラムに（左＝書き出し／中央＝スペーサー／右＝削除・上書き・追加）
        // Preset action buttons in 3 columns (left = export / center = fill spacer / right = delete, overwrite, add)
        var presetButtonRow = fontsColumn.add("group");
        presetButtonRow.orientation = "row";
        presetButtonRow.alignment = ["fill", "top"]; // 幅いっぱいにして左右に振り分ける / Stretch to full width to split left/right
        presetButtonRow.margins = [0, 5, 0, 0]; // ボタン行の上に余白 / Extra top margin above the button row

        // 左：書き出し / Left: export
        var exportPresetButton = presetButtonRow.add("button", undefined, getLocalizedText(LABELS.button.exportPreset));
        exportPresetButton.alignment = ["left", "center"];
        exportPresetButton.helpTip = getLocalizedText(LABELS.tip.exportPreset);

        // 中央：スペーサー（余白を吸って左右に振り分ける）/ Center: spacer (absorbs slack to split left/right)
        var presetButtonSpacer = presetButtonRow.add("statictext", undefined, "");
        presetButtonSpacer.alignment = ["fill", "center"];

        // 右：削除・上書き・追加 / Right: delete, overwrite, add
        var deletePresetButton = presetButtonRow.add("button", undefined, getLocalizedText(LABELS.button.deletePreset));
        deletePresetButton.alignment = ["right", "center"];
        deletePresetButton.helpTip = getLocalizedText(LABELS.tip.deletePreset);

        var overwritePresetButton = presetButtonRow.add("button", undefined, getLocalizedText(LABELS.button.overwritePreset));
        overwritePresetButton.alignment = ["right", "center"];
        overwritePresetButton.helpTip = getLocalizedText(LABELS.tip.overwritePreset);

        var addPresetButton = presetButtonRow.add("button", undefined, getLocalizedText(LABELS.button.addPreset));
        addPresetButton.alignment = ["right", "center"];
        addPresetButton.helpTip = getLocalizedText(LABELS.tip.addPreset);

        // 高さは表示後に trimButtonHeight でまとめて詰める / Heights are trimmed together after show via trimButtonHeight

        return {
            presetFontOnlyListBox: presetFontOnlyListBox,
            presetDetailedListBox: presetDetailedListBox,
            presetFontOnlyRadio: presetFontOnlyRadio,
            presetDetailedRadio: presetDetailedRadio,
            exportPresetButton: exportPresetButton,
            deletePresetButton: deletePresetButton,
            overwritePresetButton: overwritePresetButton,
            addPresetButton: addPresetButton
        };
    }

    /* 最上部の種別（本文／見出し）＋方針（和文／欧文／和欧混在）を構築して参照を返す
       Build the top panels (Type + Policy) and return refs */
    function buildTopPanels(palette) {
        var topRow = palette.add("group");
        topRow.orientation = "row";
        topRow.alignment = "fill";
        topRow.alignChildren = ["fill", "fill"]; // パネルを左右いっぱいに / Panels fill the full width
        topRow.spacing = 8;

        // 種別（本文／見出し）：タイトルなしパネル / Type (Body / Heading): untitled panel
        var rolePanel = topRow.add("panel", undefined, "");
        setupPanel(rolePanel);
        rolePanel.margins.top = 7;
        rolePanel.margins.bottom = 5; // 上下のみ指定、左右は setupPanel 既定 / Top/bottom only, left/right stay default
        rolePanel.alignment = ["fill", "fill"]; // 左右いっぱいに / Fill the full width
        rolePanel.alignChildren = ["center", "center"];
        var roleRow = rolePanel.add("group");
        roleRow.orientation = "row";
        roleRow.spacing = 16;
        roleRow.alignment = ["center", "center"]; // 上下左右中央 / Centered
        var roleBodyRadio = roleRow.add("radiobutton", undefined, getLocalizedText(LABELS.role.body));
        var roleHeadingRadio = roleRow.add("radiobutton", undefined, getLocalizedText(LABELS.role.heading));
        rolePanel.helpTip = getLocalizedText(LABELS.tip.role);
        roleBodyRadio.helpTip = rolePanel.helpTip;
        roleHeadingRadio.helpTip = rolePanel.helpTip;

        // 方針：和文／欧文／和欧混在（横並び）。文字揃えと行送りの基準をまとめて適用（タイトルなしパネル）
        // Policy: Japanese / Western / Mixed (row). Applies char-alignment + leading-basis together (untitled panel)
        var policyPanel = topRow.add("panel", undefined, "");
        setupPanel(policyPanel);
        policyPanel.margins.top = 7;
        policyPanel.margins.bottom = 5; // 上下のみ指定、左右は setupPanel 既定 / Top/bottom only, left/right stay default
        policyPanel.alignment = ["fill", "fill"]; // 左右いっぱいに / Fill the full width
        policyPanel.alignChildren = ["center", "center"];
        var policyRow = policyPanel.add("group");
        policyRow.orientation = "row";
        policyRow.spacing = 12;
        policyRow.alignment = ["center", "center"];
        var POLICY_PRESETS = [
            { id: "wabun", label: LABELS.policy.wabun, align: "center", leadingType: "top", hyphenation: false },     // 和文：中央 ＋ 仮想ボディの上 ＋ ハイフネーションOFF
            { id: "latin", label: LABELS.policy.latin, align: "roman", leadingType: "baseline", hyphenation: true },  // 欧文：欧文ベースライン ＋ 欧文ベースライン ＋ ハイフネーションON
            { id: "mixed", label: LABELS.policy.mixed, align: "roman", leadingType: "top", hyphenation: false }        // 和欧混在：欧文ベースライン ＋ 仮想ボディの上 ＋ ハイフネーションOFF
        ];
        var policyRadios = [];
        for (var policyIndex = 0; policyIndex < POLICY_PRESETS.length; policyIndex++) {
            var policyRadio = policyRow.add("radiobutton", undefined, getLocalizedText(POLICY_PRESETS[policyIndex].label));
            policyRadio.policyPreset = POLICY_PRESETS[policyIndex];
            policyRadios.push(policyRadio);
        }

        return {
            roleBodyRadio: roleBodyRadio,
            roleHeadingRadio: roleHeadingRadio,
            policyRadios: policyRadios
        };
    }

    /* フッター（制御文字トグル／再読み込み／リセット）を構築して参照を返す
       Build the footer (hidden-chars toggle / reload / reset) and return refs */
    function buildFooter(palette) {
        // フッター（左=制御文字／中央=スペーサー／右=リセット・再読み込み）/ Footer: hidden-chars, spacer, reset, reload
        var footerGroup = palette.add("group");
        footerGroup.orientation = "row";
        footerGroup.alignment = "fill";
        var hiddenCharButton = footerGroup.add("button", undefined, getLocalizedText(LABELS.button.hiddenChar));
        hiddenCharButton.alignment = ["left", "center"];
        hiddenCharButton.helpTip = getLocalizedText(LABELS.tip.hiddenChar);
        var footerSpacer = footerGroup.add("statictext", undefined, "");
        footerSpacer.alignment = ["fill", "center"];
        var resetButton = footerGroup.add("button", undefined, getLocalizedText(LABELS.button.reset));
        resetButton.alignment = ["right", "center"];
        resetButton.helpTip = getLocalizedText(LABELS.tip.reset);
        var reloadButton = footerGroup.add("button", undefined, getLocalizedText(LABELS.button.reload));
        reloadButton.alignment = ["right", "center"];
        reloadButton.helpTip = getLocalizedText(LABELS.tip.reload);

        return {
            hiddenCharButton: hiddenCharButton,
            reloadButton: reloadButton,
            resetButton: resetButton
        };
    }

    /* 中央カラム（フォントサイズ／自動カーニング／字間調整／行揃え）を構築して参照を返す
       Build the center column (font size / auto kerning / letter spacing / justification) and return refs */
    function buildCharacterColumn(typeColumnsRow, textUnit, autoKernOptions, justifyOptions) {
        var characterColumn = typeColumnsRow.add("group");
        characterColumn.orientation = "column";
        characterColumn.alignChildren = ["fill", "top"];
        characterColumn.spacing = 8;

        // フォントサイズ（サイズ）は行送りパネルへ移動 / Font size (size) has moved to the leading panel

        var autoKernPanel = characterColumn.add("panel", undefined, getLocalizedText(LABELS.field.autoKern));
        setupPanel(autoKernPanel, 6); // spacing は行送りのラジオと同じ / Same row spacing as the leading radios
        autoKernPanel.alignChildren = ["left", "top"];
        autoKernPanel.helpTip = getLocalizedText(LABELS.tip.autoKern);

        var kernRadios = [];
        for (var i = 0; i < autoKernOptions.length; i++) {
            var kernRadio = autoKernPanel.add("radiobutton", undefined, getLocalizedText(autoKernOptions[i].label));
            kernRadio.value = false;
            kernRadio.index = i;
            kernRadios.push(kernRadio);
        }

        // 字間調整（文字ツメ・トラッキング）：各行「ラベル：［入力］」＋スライダー
        // Letter spacing (Tsume / Tracking): each row "label: [input]" then a slider
        var spacingPanel = characterColumn.add("panel", undefined, getLocalizedText(LABELS.field.spacingAdjust));
        setupPanel(spacingPanel, 6);

        // 一番上：プロポーショナルメトリクス / Top: Proportional metrics checkbox
        var proportionalMetricsCheck = spacingPanel.add("checkbox", undefined, getLocalizedText(LABELS.field.proportionalMetrics));
        proportionalMetricsCheck.value = false;

        var tsumeRow = spacingPanel.add("group");
        setupGroup(tsumeRow, "row");
        tsumeRow.add("statictext", undefined, labelText(LABELS.field.tsume));
        var tsumeInput = tsumeRow.add("edittext", undefined, "0");
        tsumeInput.characters = 3;
        tsumeInput.helpTip = getLocalizedText(LABELS.tip.tsume);
        tsumeRow.add("statictext", undefined, "%");
        var tsumeSlider = spacingPanel.add("slider", undefined, 0, 0, 100);

        // 文字ツメとトラッキングの間に少し余白 / A little gap between Tsume and Tracking
        var spacingGap = spacingPanel.add("group");
        spacingGap.preferredSize.height = 1;

        var trackingRow = spacingPanel.add("group");
        setupGroup(trackingRow, "row");
        trackingRow.add("statictext", undefined, labelText(LABELS.field.tracking));
        var trackingInput = trackingRow.add("edittext", undefined, "0");
        trackingInput.characters = 3;
        trackingInput.helpTip = getLocalizedText(LABELS.tip.tracking);
        var trackingSlider = spacingPanel.add("slider", undefined, 0, -100, 500);

        // 文字揃え（縦方向）：行揃えと位置を入れ替えて中央カラムの末尾に配置
        // Char alignment (vertical): swapped with justification, placed at the end of the center column
        var alignPanel = characterColumn.add("panel", undefined, getLocalizedText(LABELS.field.align));
        setupPanel(alignPanel, 6);
        alignPanel.alignChildren = ["left", "top"];
        alignPanel.helpTip = getLocalizedText(LABELS.tip.align);

        // 3つのラジオ：欧文ベースライン／中央／その他（その他は残りをポップアップで指定）
        // Three radios: Roman baseline / Center / Other (Other picks the rest from a popup)
        // ラジオは alignId で排他制御（親が異なるため selectAlignRadio で手動切替）
        var alignRadios = [];
        var alignRomanRadio = alignPanel.add("radiobutton", undefined, getLocalizedText(LABELS.align.roman));
        alignRomanRadio.alignId = "roman";
        alignRadios.push(alignRomanRadio);
        var alignCenterRadio = alignPanel.add("radiobutton", undefined, getLocalizedText(LABELS.align.center));
        alignCenterRadio.alignId = "center";
        alignRadios.push(alignCenterRadio);

        // 「その他」ラジオの右にポップアップ / Popup to the right of the "Other" radio
        var alignOtherRow = alignPanel.add("group");
        alignOtherRow.orientation = "row";
        alignOtherRow.alignChildren = ["left", "center"];
        alignOtherRow.spacing = 6;
        var alignOtherRadio = alignOtherRow.add("radiobutton", undefined, ""); // ラベルなし（右のポップアップで内容を示す）/ No label (the popup shows the choice)
        alignOtherRadio.alignId = "other";
        alignRadios.push(alignOtherRadio);
        var alignOtherOptions = [
            { id: "top", label: LABELS.align.top },
            { id: "bottom", label: LABELS.align.bottom },
            { id: "icftop", label: LABELS.align.icfTop },
            { id: "icfbottom", label: LABELS.align.icfBottom }
        ];
        var alignOtherDropdown = alignOtherRow.add("dropdownlist", undefined, []);
        for (var optionIndex = 0; optionIndex < alignOtherOptions.length; optionIndex++) {
            alignOtherDropdown.add("item", getLocalizedText(alignOtherOptions[optionIndex].label));
        }
        alignOtherDropdown.selection = 0;
        alignOtherDropdown.preferredSize.width = 124; // 少し狭く / A bit narrower
        alignRomanRadio.value = true;

        return {
            kernRadios: kernRadios,
            proportionalMetricsCheck: proportionalMetricsCheck,
            tsumeInput: tsumeInput,
            tsumeSlider: tsumeSlider,
            trackingInput: trackingInput,
            trackingSlider: trackingSlider,
            alignRadios: alignRadios,
            alignRomanRadio: alignRomanRadio,
            alignCenterRadio: alignCenterRadio,
            alignOtherRadio: alignOtherRadio,
            alignOtherDropdown: alignOtherDropdown,
            alignOtherOptions: alignOtherOptions
        };
    }

    /* 右カラム（文字揃え／行送り／行送りの基準／禁則／文字組みアキ量）を構築して参照を返す
       Build the right column (char alignment / leading / basis / kinsoku / mojikumi) and return refs */
    function buildParagraphColumn(typeColumnsRow, leadingBasisChoices, justifyOptions) {
        var paragraphColumn = typeColumnsRow.add("group");
        paragraphColumn.orientation = "column";
        paragraphColumn.alignChildren = ["fill", "top"];
        paragraphColumn.spacing = 8;

        // 行送り：常に自動行送り。サイズ・実質（pt）・行送り（%）の3入力を1パネルに。右カラムの先頭に配置
        // Leading: always auto. Size, effective (pt), and leading (%) in one panel, at the top of the right column
        var leadingPanel = paragraphColumn.add("panel", undefined, getLocalizedText(LABELS.field.sizeAndLeading));
        setupPanel(leadingPanel, 6);
        leadingPanel.alignChildren = "left";
        leadingPanel.helpTip = getLocalizedText(LABELS.tip.leading);

        var leadColon = currentLanguage === "ja" ? "：" : ": ";
        var leadTextUnit = getTextUnit();

        // フォントサイズ（フォントサイズパネルから移動）/ Font size (moved from the font-size panel)
        var leadSizeRow = leadingPanel.add("group");
        leadSizeRow.orientation = "row";
        leadSizeRow.alignChildren = ["left", "center"];
        var leadSizeLabel = leadSizeRow.add("statictext", undefined, getLocalizedText(LABELS.field.fontSize) + leadColon);
        var fontSizeInput = leadSizeRow.add("edittext", undefined, "");
        fontSizeInput.characters = 3;
        leadSizeRow.add("statictext", undefined, leadTextUnit.label);
        leadSizeLabel.helpTip = getLocalizedText(LABELS.tip.fontSize); fontSizeInput.helpTip = leadSizeLabel.helpTip;

        // 実質（フォントサイズ×行送り% の結果。ここに入力すると % を逆算）/ Effective leading (size × %); entering a value back-calculates the %
        var leadEffectiveRow = leadingPanel.add("group");
        leadEffectiveRow.orientation = "row";
        leadEffectiveRow.alignChildren = ["left", "center"];
        var leadEffectiveLabel = leadEffectiveRow.add("statictext", undefined, getLocalizedText(LABELS.field.leading) + leadColon);
        var leadingEffectiveInput = leadEffectiveRow.add("edittext", undefined, "");
        leadingEffectiveInput.characters = 3;
        // 実質は行送りなので Q/H 環境では単位「H」/ Effective is a leading value, so use "H" when the unit is Q/H
        leadEffectiveRow.add("statictext", undefined, getLeadingUnitLabel(leadTextUnit.code));
        leadEffectiveLabel.helpTip = getLocalizedText(LABELS.tip.leadingEffective); leadingEffectiveInput.helpTip = leadEffectiveLabel.helpTip;

        // 行送り（自動行送り量 %）/ Leading (auto-leading amount %)
        var leadPercentRow = leadingPanel.add("group");
        leadPercentRow.orientation = "row";
        leadPercentRow.alignChildren = ["left", "center"];
        var leadPercentLabel = leadPercentRow.add("statictext", undefined, getLocalizedText(LABELS.field.leadingPercent) + leadColon);
        var leadingPercentInput = leadPercentRow.add("edittext", undefined, "");
        leadingPercentInput.characters = 3;
        leadPercentRow.add("statictext", undefined, "%");
        leadPercentLabel.helpTip = getLocalizedText(LABELS.tip.leading); leadingPercentInput.helpTip = leadPercentLabel.helpTip;

        // ラベル幅を揃える / Unify label widths
        var leadLabelWidth = 65;
        leadSizeLabel.preferredSize.width = leadLabelWidth;
        leadEffectiveLabel.preferredSize.width = leadLabelWidth;
        leadPercentLabel.preferredSize.width = leadLabelWidth;

        // 行送りの基準：独立したタイトル付きパネル（行送りパネルの外）/ Leading basis: its own titled panel (outside the leading panel)
        var leadingBasisPanel = paragraphColumn.add("panel", undefined, getLocalizedText(LABELS.field.leadingType));
        setupPanel(leadingBasisPanel, 6);
        leadingBasisPanel.alignChildren = "left";
        leadingBasisPanel.helpTip = getLocalizedText(LABELS.tip.leadingType);
        var leadingBasisRadios = [];
        for (var basisIndex = 0; basisIndex < leadingBasisChoices.length; basisIndex++) {
            var leadingBasisRadio = leadingBasisPanel.add("radiobutton", undefined, leadingBasisChoices[basisIndex].label);
            leadingBasisRadio.index = basisIndex;
            leadingBasisRadios.push(leadingBasisRadio);
        }
        leadingBasisRadios[0].value = true;

        // 行揃え：行送り・基準の下に配置 / Justification: placed below leading and its basis
        var justifyPanel = paragraphColumn.add("panel", undefined, getLocalizedText(LABELS.field.justify));
        setupPanel(justifyPanel, 5);
        justifyPanel.orientation = "row";
        justifyPanel.alignChildren = ["center", "center"]; // ボタン列を左右中央に / Center the button row horizontally
        justifyPanel.helpTip = getLocalizedText(LABELS.tip.justify);

        // アクティブな行揃え id と UI 明暗を共有（onDraw のクロージャから参照）/ Shared active id + theme (read by onDraw closures)
        var justifyState = { activeId: "", isLight: isLightUI() };
        var justifyButtons = [];
        for (var buttonIndex = 0; buttonIndex < justifyOptions.length; buttonIndex++) {
            var justifyOption = justifyOptions[buttonIndex];
            var justifyButton = justifyPanel.add("button", undefined, "");
            justifyButton.helpTip = getLocalizedText(justifyOption.label);
            justifyButton.preferredSize = [26, 26];
            justifyButton.minimumSize = [26, 26];
            justifyButton.maximumSize = [26, 26];
            justifyButton.justifyId = justifyOption.id;
            justifyButton.iconType = justifyOption.icon;
            justifyButton.onDraw = function () { drawJustifyIcon(this, this.justifyId === justifyState.activeId, justifyState.isLight); };
            justifyButtons.push(justifyButton);
        }

        // 日本語組版：禁則＋文字組みアキ量設定を1枚にまとめ、各見出しはラベルとして表示
        // Japanese typography: Kinsoku + Mojikumi in one panel; each sub-heading is shown as a label
        var jpPanel = paragraphColumn.add("panel", undefined, getLocalizedText(LABELS.field.jpComposition));
        setupPanel(jpPanel, 5);
        jpPanel.alignChildren = "fill";

        // 禁則（ラベル＋ポップアップ）/ Kinsoku (label + popup)
        var kinsokuHeading = jpPanel.add("statictext", undefined, labelText(LABELS.field.kinsoku));
        kinsokuHeading.helpTip = getLocalizedText(LABELS.tip.kinsoku);
        var kinsokuItems = [];
        for (var kinsokuChoiceIndex = 0; kinsokuChoiceIndex < KINSOKU_CHOICES.length; kinsokuChoiceIndex++) kinsokuItems.push(KINSOKU_CHOICES[kinsokuChoiceIndex].label);
        var kinsokuDropdown = jpPanel.add("dropdownlist", undefined, kinsokuItems);
        kinsokuDropdown.selection = 0;
        kinsokuDropdown.helpTip = kinsokuHeading.helpTip;
        kinsokuDropdown.alignment = "left"; // fill を無効化して幅を指定 / Override fill so the explicit width applies
        kinsokuDropdown.preferredSize.width = 150;

        // 禁則と文字組みの間の余白 / Gap between the two groups
        var jpGap = jpPanel.add("group");
        jpGap.preferredSize.height = 4;

        // 文字組みアキ量設定（ラベル＋ポップアップ）/ Mojikumi spacing set (label + popup)
        var mojikumiHeading = jpPanel.add("statictext", undefined, labelText(LABELS.field.mojikumi));
        mojikumiHeading.helpTip = getLocalizedText(LABELS.tip.mojikumi);
        var mojikumiItems = [];
        for (var mojikumiChoiceIndex = 0; mojikumiChoiceIndex < MOJIKUMI_CHOICES.length; mojikumiChoiceIndex++) mojikumiItems.push(MOJIKUMI_CHOICES[mojikumiChoiceIndex].label);
        var mojikumiDropdown = jpPanel.add("dropdownlist", undefined, mojikumiItems);
        mojikumiDropdown.selection = 0;
        mojikumiDropdown.helpTip = mojikumiHeading.helpTip;
        mojikumiDropdown.alignment = "left"; // fill を無効化して幅を指定 / Override fill so the explicit width applies
        mojikumiDropdown.preferredSize.width = 150;

        return {
            justifyButtons: justifyButtons,
            justifyState: justifyState,
            fontSizeInput: fontSizeInput,
            leadingEffectiveInput: leadingEffectiveInput,
            leadingPercentInput: leadingPercentInput,
            leadingBasisRadios: leadingBasisRadios,
            kinsokuDropdown: kinsokuDropdown,
            mojikumiDropdown: mojikumiDropdown
        };
    }

    /* 「フォントサイズ」タブ：サイズ・比率・実質をまとめた独立パネル（中央カラムから複製）
       "Font Size" tab: a standalone panel with size / scale / effective (duplicated from the center column) */
    function buildFontSizeTab(fontSizeTab, textUnit) {
        // サイズタブを左右2カラムに（左：フォントサイズ／右：タイプスケール）
        // Two columns for the Size tab (left: font size, right: type scale)
        var columnsRow = fontSizeTab.add("group");
        columnsRow.orientation = "row";
        columnsRow.alignChildren = ["fill", "top"];
        columnsRow.spacing = 12;

        // ---- 左カラム：フォントサイズ / Left column: font size ----
        var fontSizeColumn = columnsRow.add("group");
        fontSizeColumn.orientation = "column";
        fontSizeColumn.alignChildren = ["fill", "top"];
        fontSizeColumn.spacing = 8;

        var fontSizePanel = fontSizeColumn.add("panel", undefined, getLocalizedText(LABELS.field.fontSizePanel));
        setupPanel(fontSizePanel, 6);
        fontSizePanel.alignChildren = ["left", "top"];
        fontSizePanel.alignment = ["left", "top"]; // タブ幅いっぱいに広げず、内容幅に合わせる / Size to content instead of filling the tab width

        var sizeColon = currentLanguage === "ja" ? "：" : ": ";

        var fsSizeRow = fontSizePanel.add("group");
        fsSizeRow.orientation = "row";
        fsSizeRow.alignChildren = ["left", "center"];
        var fsSizeLabel = fsSizeRow.add("statictext", undefined, getLocalizedText(LABELS.field.fontSize) + sizeColon);
        var fontSizeInput = fsSizeRow.add("edittext", undefined, "");
        fontSizeInput.characters = 4;
        fsSizeRow.add("statictext", undefined, textUnit.label);

        var fsScaleRow = fontSizePanel.add("group");
        fsScaleRow.orientation = "row";
        fsScaleRow.alignChildren = ["left", "center"];
        var fsScaleLabel = fsScaleRow.add("statictext", undefined, getLocalizedText(LABELS.field.scale) + sizeColon);
        var scaleInput = fsScaleRow.add("edittext", undefined, "100");
        scaleInput.characters = 4;
        fsScaleRow.add("statictext", undefined, "%");

        var fsApparentRow = fontSizePanel.add("group");
        fsApparentRow.orientation = "row";
        fsApparentRow.alignChildren = ["left", "center"];
        fsApparentRow.margins = [0, 5, 0, 0]; // 「実質」行の上に余白 / Extra top margin above the "Effective" row
        var fsApparentLabel = fsApparentRow.add("statictext", undefined, getLocalizedText(LABELS.field.apparent) + sizeColon);
        var apparentValueLabel = fsApparentRow.add("statictext", undefined, "--");
        apparentValueLabel.characters = 5;
        var fsApparentUnit = fsApparentRow.add("statictext", undefined, textUnit.label);

        fsSizeLabel.helpTip = getLocalizedText(LABELS.tip.fontSize); fontSizeInput.helpTip = fsSizeLabel.helpTip;
        fsScaleLabel.helpTip = getLocalizedText(LABELS.tip.scale); scaleInput.helpTip = fsScaleLabel.helpTip;
        fsApparentLabel.helpTip = getLocalizedText(LABELS.tip.apparent); apparentValueLabel.helpTip = fsApparentLabel.helpTip;

        // ラベル幅を揃える / Unify label widths
        var fsLabelWidth = 55;
        fsSizeLabel.preferredSize.width = fsLabelWidth;
        fsScaleLabel.preferredSize.width = fsLabelWidth;
        fsApparentLabel.preferredSize.width = fsLabelWidth;

        // ボタン上に 5px のマージン / A 5px margin above the button
        var apparentButtonGap = fontSizePanel.add("group");
        apparentButtonGap.preferredSize.height = 5;

        // 実サイズ↔見かけ：サイズ×比率を実サイズへ焼き込み比率100%に／もう一度で復元
        // Actual ↔ Apparent: bake size × scale into the actual size at 100%; press again to restore
        var apparentToggleButton = fontSizePanel.add("button", undefined, getLocalizedText(LABELS.button.toApparent));
        apparentToggleButton.helpTip = getLocalizedText(LABELS.tip.toApparent);
        apparentToggleButton.alignment = "right";
        apparentToggleButton.preferredSize.width = 150;
        // 高さは trimButtonHeight でレイアウト確定後に詰める / Height is trimmed after layout via trimButtonHeight

        // ---- 右カラム：タイプスケール（TypeScaler 由来）/ Right column: type scale (from TypeScaler) ----
        var typeScaleColumn = columnsRow.add("group");
        typeScaleColumn.orientation = "column";
        typeScaleColumn.alignChildren = ["fill", "top"];
        typeScaleColumn.spacing = 8;

        var typeScalePanel = typeScaleColumn.add("panel", undefined, getLocalizedText(LABELS.field.typeScale));
        setupPanel(typeScalePanel, 6);
        typeScalePanel.alignChildren = ["left", "top"];
        typeScalePanel.helpTip = getLocalizedText(LABELS.tip.typeScale);

        // 基準サイズ入力 / Base size input
        var tsBaseRow = typeScalePanel.add("group");
        tsBaseRow.orientation = "row";
        tsBaseRow.alignChildren = ["left", "center"];
        var tsBaseLabel = tsBaseRow.add("statictext", undefined, getLocalizedText(LABELS.field.typeScaleBase) + sizeColon);
        var typeScaleBaseInput = tsBaseRow.add("edittext", undefined, "");
        typeScaleBaseInput.characters = 4;
        tsBaseRow.add("statictext", undefined, textUnit.label);
        tsBaseLabel.helpTip = getLocalizedText(LABELS.tip.typeScaleBase); typeScaleBaseInput.helpTip = tsBaseLabel.helpTip;

        // 倍率ポップアップ / Ratio popup
        var ratioLabels = [];
        for (var ri = 0; ri < TYPE_SCALE_RATIOS.length; ri++) ratioLabels.push(TYPE_SCALE_RATIOS[ri].label);
        var typeScaleRatioPopup = typeScalePanel.add("dropdownlist", undefined, ratioLabels);
        typeScaleRatioPopup.selection = TYPE_SCALE_DEFAULT_RATIO_INDEX;
        typeScaleRatioPopup.helpTip = getLocalizedText(LABELS.tip.typeScaleRatio);
        typeScaleRatioPopup.preferredSize.width = 180;

        // サイズ一覧（段番号＋サイズの2カラム）/ Size list (step number + size, 2 columns)
        var typeScaleList = typeScalePanel.add("listbox", undefined, [], {
            multiselect: false,
            numberOfColumns: 2,
            showHeaders: true,
            columnTitles: [getLocalizedText(LABELS.field.typeScaleStep), getLocalizedText(LABELS.field.fontSize)],
            columnWidths: [40, 130]
        });
        typeScaleList.preferredSize = [180, 160];

        // 基準・倍率からサイズ一覧を再生成 / Rebuild the size list from base + ratio
        function rebuildTypeScaleList() {
            typeScaleList.removeAll();
            var baseSize = parseFloat(typeScaleBaseInput.text);
            if (isNaN(baseSize) || baseSize <= 0) return;
            var ratio = TYPE_SCALE_RATIOS[typeScaleRatioPopup.selection ? typeScaleRatioPopup.selection.index : TYPE_SCALE_DEFAULT_RATIO_INDEX].value;
            var sizes = generateTypeScaleSizes(baseSize, ratio);
            for (var j = 0; j < sizes.length; j++) {
                var item = typeScaleList.add("item", String(j - 2)); // index 2 が基準（段番号 0）/ index 2 is the base (step 0)
                item.subItems[0].text = sizes[j] + " " + textUnit.label;
            }
        }
        typeScaleBaseInput.onChanging = rebuildTypeScaleList;
        typeScaleRatioPopup.onChange = rebuildTypeScaleList;
        changeValueByArrowKey(typeScaleBaseInput, false, rebuildTypeScaleList, 1);

        return {
            fontSizeInput: fontSizeInput,
            scaleInput: scaleInput,
            apparentValueLabel: apparentValueLabel,
            fsApparentLabel: fsApparentLabel,
            fsApparentUnit: fsApparentUnit,
            apparentToggleButton: apparentToggleButton,
            typeScaleBaseInput: typeScaleBaseInput,
            typeScaleRatioPopup: typeScaleRatioPopup,
            typeScaleList: typeScaleList,
            rebuildTypeScaleList: rebuildTypeScaleList
        };
    }

    /* パレットを組み立てて参照を返す（イベント未接続）/ Build the palette and return references (events not wired yet) */
    function createPaletteUI(autoKernOptions, alignOptions, justifyOptions) {
        var palette = new Window("palette", getLocalizedText(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        palette.alignChildren = "fill";

        var textUnit = getTextUnit();
        var leadingBasisChoices = getLeadingTypeChoices();

        // ===== 上部：種別（タイトルなし）＋方針を左右に並べて構築 / Top: Type + Policy panels side by side =====
        var topPanelsUI = buildTopPanels(palette);

        // ===== レイアウト：タブ1（ドキュメントフォント）／タブ2（カーニング・ツメ・文字揃え・行送りなど）
        //       Layout: tab 1 (document fonts) / tab 2 (kerning, tsume, alignment, leading, ...) =====
        var mainTabs = palette.add("tabbedpanel");
        mainTabs.alignChildren = ["fill", "top"];

        // タブ1：文字組み・タブ2：フォントサイズ・タブ3：フォント／プリセット・タブ4：使用フォント（生成順＝表示順）
        // Tab 1: type settings, Tab 2: font size, Tab 3: fonts/presets, Tab 4: used fonts (creation order = strip order)
        var typeTab = addTab(mainTabs, LABELS.tab.type);
        var fontSizeTab = addTab(mainTabs, LABELS.tab.fontSize);
        var fontsTab = addTab(mainTabs, LABELS.tab.fonts);
        var usedFontsTab = addTab(mainTabs, LABELS.tab.usedFonts);
        var styleTab = null; // 仮に OFF（addTab(mainTabs, LABELS.tab.style) で有効化）/ Temporarily OFF

        // フォントサイズタブの中身を構築 / Build the font-size tab contents
        var fontSizeTabUI = buildFontSizeTab(fontSizeTab, textUnit);

        // 使用フォントタブの中身を構築 / Build the used-fonts tab contents
        var usedFontsUI = buildUsedFontsTab(usedFontsTab);

        var typeColumnsRow = typeTab.add("group");
        typeColumnsRow.orientation = "row";
        typeColumnsRow.alignChildren = ["fill", "top"];
        typeColumnsRow.spacing = 12;

        mainTabs.selection = typeTab; // 起動時は文字組みタブを開く / Open the Type tab on launch

        // ---- 左カラム：ドキュメントフォント／プリセットを構築 / Left column: document fonts + presets ----
        var fontsTabUI = buildFontsTab(fontsTab);

        // ---- 中央カラム：フォントサイズ（サイズのみ）・自動カーニング・文字ツメ・文字揃えを構築 / Center column ----
        var characterColumnUI = buildCharacterColumn(typeColumnsRow, textUnit, autoKernOptions, justifyOptions);

        // ---- 右カラム：行揃え・行送り・行送りの基準・禁則・文字組みアキ量設定を構築 / Right column ----
        var paragraphColumnUI = buildParagraphColumn(typeColumnsRow, leadingBasisChoices, justifyOptions);

        // フッターを構築 / Build the footer
        var footerUI = buildFooter(palette);

        return {
            palette: palette,
            mainTabs: mainTabs,
            typeTab: typeTab,
            usedFontsTab: usedFontsTab,
            styleTab: styleTab,
            comboListBox: usedFontsUI.comboListBox,
            fontOnlyListBox: usedFontsUI.fontOnlyListBox,
            fontOnlyRadio: usedFontsUI.fontOnlyRadio,
            detailedRadio: usedFontsUI.detailedRadio,
            includeLockedCheckbox: usedFontsUI.includeLockedCheckbox,
            includeHiddenCheckbox: usedFontsUI.includeHiddenCheckbox,
            currentArtboardCheckbox: usedFontsUI.currentArtboardCheckbox,
            refreshCombosButton: usedFontsUI.refreshCombosButton,
            selectMatchingButton: usedFontsUI.selectMatchingButton,
            presetFontOnlyListBox: fontsTabUI.presetFontOnlyListBox,
            presetDetailedListBox: fontsTabUI.presetDetailedListBox,
            presetFontOnlyRadio: fontsTabUI.presetFontOnlyRadio,
            presetDetailedRadio: fontsTabUI.presetDetailedRadio,
            exportPresetButton: fontsTabUI.exportPresetButton,
            addPresetButton: fontsTabUI.addPresetButton,
            overwritePresetButton: fontsTabUI.overwritePresetButton,
            deletePresetButton: fontsTabUI.deletePresetButton,
            // 文字組みタブのサイズ欄と「フォントサイズ」タブのサイズ欄を同期して扱う / Both size fields (Type tab + Font Size tab) are kept in sync
            fontSizeInputs: [paragraphColumnUI.fontSizeInput, fontSizeTabUI.fontSizeInput],
            scaleInput: fontSizeTabUI.scaleInput,
            apparentValueLabel: fontSizeTabUI.apparentValueLabel,
            fsApparentLabel: fontSizeTabUI.fsApparentLabel,
            fsApparentUnit: fontSizeTabUI.fsApparentUnit,
            apparentToggleButton: fontSizeTabUI.apparentToggleButton,
            // タイプスケール（サイズタブ右カラム）/ Type scale (right column of the Size tab)
            typeScaleBaseInput: fontSizeTabUI.typeScaleBaseInput,
            typeScaleRatioPopup: fontSizeTabUI.typeScaleRatioPopup,
            typeScaleList: fontSizeTabUI.typeScaleList,
            rebuildTypeScaleList: fontSizeTabUI.rebuildTypeScaleList,
            roleBodyRadio: topPanelsUI.roleBodyRadio,
            roleHeadingRadio: topPanelsUI.roleHeadingRadio,
            policyRadios: topPanelsUI.policyRadios,
            alignRadios: characterColumnUI.alignRadios,
            alignRomanRadio: characterColumnUI.alignRomanRadio,
            alignCenterRadio: characterColumnUI.alignCenterRadio,
            alignOtherRadio: characterColumnUI.alignOtherRadio,
            alignOtherDropdown: characterColumnUI.alignOtherDropdown,
            alignOtherOptions: characterColumnUI.alignOtherOptions,
            justifyButtons: paragraphColumnUI.justifyButtons,
            justifyState: paragraphColumnUI.justifyState,
            alignOptions: alignOptions,
            justifyOptions: justifyOptions,
            kernRadios: characterColumnUI.kernRadios,
            proportionalMetricsCheck: characterColumnUI.proportionalMetricsCheck,
            tsumeSlider: characterColumnUI.tsumeSlider,
            tsumeInput: characterColumnUI.tsumeInput,
            trackingSlider: characterColumnUI.trackingSlider,
            trackingInput: characterColumnUI.trackingInput,
            resetButton: footerUI.resetButton,
            reloadButton: footerUI.reloadButton,
            hiddenCharButton: footerUI.hiddenCharButton,
            textUnit: textUnit,
            leadingBasisChoices: leadingBasisChoices,
            leadingEffectiveInput: paragraphColumnUI.leadingEffectiveInput,
            leadingPercentInput: paragraphColumnUI.leadingPercentInput,
            leadingBasisRadios: paragraphColumnUI.leadingBasisRadios,
            mojikumiDropdown: paragraphColumnUI.mojikumiDropdown,
            mojikumiChoices: MOJIKUMI_CHOICES,
            kinsokuDropdown: paragraphColumnUI.kinsokuDropdown,
            kinsokuChoices: KINSOKU_CHOICES
        };
    }

    /* ラジオ配列の排他選択（別コンテナのため手動）/ Exclusive selection across radios in separate containers */
    function selectExclusiveRadio(radios, index) {
        for (var i = 0; i < radios.length; i++) radios[i].value = (i === index);
    }

    /* 現在選択中のラジオ index を返す（未選択は -1）/ Index of the currently selected radio (-1 if none) */
    function selectedRadioIndex(radios) {
        for (var i = 0; i < radios.length; i++) if (radios[i].value) return i;
        return -1;
    }

    /* "count|isAuto|leadingPt|basePt|autoAmount|leadingTypeId|paraCount|alignmentId|
        fontFamily|fontStyle|fontSizePt|kernId|justifyId|hScale|mojikumiId|kinsokuId" を分解
       Parse the encoded leading state string */
    function parseLeadingState(rest) {
        var fields = String(rest || "").split("|");
        function toNumber(text) { var parsed = parseFloat(text); return isNaN(parsed) ? NaN : parsed; }
        return {
            count: parseInt(fields[0], 10) || 0,
            isAuto: fields[1] === "1",
            leadingPt: toNumber(fields[2]),
            baseFontSizePt: toNumber(fields[3]),
            autoAmount: toNumber(fields[4]),
            leadingTypeId: fields[5] || "",
            paragraphCount: parseInt(fields[6], 10) || 0,
            alignmentId: fields[7] || "",
            fontFamily: decodeURIComponent(fields[8] || ""),
            fontStyle: decodeURIComponent(fields[9] || ""),
            fontSizePt: toNumber(fields[10]),
            kernId: fields[11] || "",
            justifyId: fields[12] || "",
            hScale: toNumber(fields[13]),
            mojikumiId: isNaN(parseInt(fields[14], 10)) ? -2 : parseInt(fields[14], 10),
            kinsokuId: fields[15] || "None"
        };
    }

    /* パレットにイベントを接続 / Wire palette events */
    function bindPaletteEvents(ui, autoKernOptions, alignOptions, justifyOptions) {
        // BridgeTalk は非同期なので、連打による委譲の重複を抑止 / Guard against overlapping async delegations
        var workerBusy = false;
        var textUnit = ui.textUnit;
        var suppressUiEvents = false; // 反映中のドロップダウン onChange を抑止 / Suppress dropdown onChange while reflecting state

        /* 重要処理の失敗をダイアログで通知（握りつぶさず、どのアクションで何が起きたか見えるように）
           Surface an important-op failure via an alert (don't swallow it; show which action failed and why) */
        function showWorkerError(actionId, payload) {
            var detail = payload ? (": " + String(payload)) : "";
            try { alert("⚠ " + getLocalizedText(LABELS.msg.applyError) + " [" + actionId + "]" + detail); } catch (e) { }
        }

        // 委譲する共通処理（連打抑止）/ Delegate an apply action (guarded against overlap)
        function runApply(actionId, params, onDone) {
            if (workerBusy) return;
            workerBusy = true;
            runWorker(actionId, params, function (status, payload) {
                workerBusy = false;
                // BridgeTalk エラーや dispatch の "ERR:<msg>" はここで可視化 / BridgeTalk errors and dispatch "ERR:<msg>" become visible here
                if (status === "error") showWorkerError(actionId, payload);
                if (onDone) onDone(status, payload);
            });
        }

        // ---- プリセット / Presets（フォント＋カーニング・ツメ・トラッキングをまとめて適用）----
        // どちらのビューでクリックしても「プリセット一括適用」。表示モードは列の見せ方だけ
        // Clicking in either view applies the whole preset; the mode only changes how columns are shown
        function applyPresetFromItem(selectedPreset) {
            if (!selectedPreset || !selectedPreset.psName) return;
            var params = { psName: selectedPreset.psName };
            if (selectedPreset.kern) {
                params.kern = selectedPreset.kern;
                // カーニングのラジオも UI に反映 / Reflect the kerning radio on the UI too
                for (var kernIndex = 0; kernIndex < autoKernOptions.length; kernIndex++) {
                    if (autoKernOptions[kernIndex].id === selectedPreset.kern) { selectExclusiveRadio(ui.kernRadios, kernIndex); break; }
                }
            }
            if (selectedPreset.tsume !== undefined && selectedPreset.tsume !== null) {
                params.tsume = selectedPreset.tsume;
                // ツメは UI にも反映 / Reflect Tsume on the slider too
                ui.tsumeSlider.value = selectedPreset.tsume;
                ui.tsumeInput.text = String(selectedPreset.tsume);
            }
            if (selectedPreset.tracking !== undefined && selectedPreset.tracking !== null) {
                params.tracking = selectedPreset.tracking;
                // トラッキングも UI に反映 / Reflect tracking on the slider too
                ui.trackingSlider.value = selectedPreset.tracking;
                ui.trackingInput.text = String(selectedPreset.tracking);
            }
            runApply("applyFontPreset", params);
        }
        ui.presetFontOnlyListBox.onChange = function () { applyPresetFromItem(this.selection); };
        ui.presetDetailedListBox.onChange = function () { applyPresetFromItem(this.selection); };

        // 現在アクティブなプリセット listbox（追加・上書き・削除の選択元）/ The active preset listbox (source for add/overwrite/delete)
        function activePresetList() { return ui.presetDetailedRadio.value ? ui.presetDetailedListBox : ui.presetFontOnlyListBox; }

        // 表示モード切替（stack なので visible 切替のみ）/ Switch view mode (stack: toggle visibility only)
        function applyPresetView() {
            var fontOnly = ui.presetFontOnlyRadio.value;
            ui.presetFontOnlyListBox.visible = fontOnly;
            ui.presetDetailedListBox.visible = !fontOnly;
        }
        ui.presetFontOnlyRadio.onClick = function () { applyPresetView(); };
        ui.presetDetailedRadio.onClick = function () { applyPresetView(); };
        applyPresetView(); // 初期表示（既定＝フォントのみ）/ Initial view (default = font-only)

        // 現在の選択の設定をプリセットオブジェクトへ変換 / Build a preset object from the current selection's settings
        function buildPresetFromSelection(callback) {
            runWorker("getCurrentFont", null, function (status, payload) {
                if (status !== "ok" || !payload) { callback(null); return; }
                var fontParts = decodeURIComponent(payload).split("|");
                callback({
                    psName: fontParts[0],
                    kern: fontParts[1] || undefined,
                    tsume: fontParts[2] !== "" ? parseInt(fontParts[2], 10) : undefined,
                    tracking: fontParts[3] !== "" ? parseInt(fontParts[3], 10) : undefined
                });
            });
        }

        // プリセット保存後にリストと表示名を更新 / Persist, then refresh the list and labels
        function commitPresets(presets) {
            savePresets(presets);
            fillPresetLists(ui.presetFontOnlyListBox, ui.presetDetailedListBox, presets);
            presetLabelsResolved = false; // 追加・更新直後も表示名を再解決する / Re-resolve labels after add/update
            refreshPresetLabels();
        }

        // ---- プリセットに追加（現在の選択の設定を保存）/ Add the current selection's settings as a preset ----
        ui.addPresetButton.onClick = function () {
            buildPresetFromSelection(function (newPreset) {
                if (!newPreset) return;
                var presets = loadPresets();
                // 同じ PostScript 名があれば設定を上書き / Overwrite settings if the font is already present
                var replaced = false;
                for (var presetIndex = 0; presetIndex < presets.length; presetIndex++) {
                    if (presets[presetIndex].psName === newPreset.psName) { presets[presetIndex] = newPreset; replaced = true; break; }
                }
                if (!replaced) presets.push(newPreset);
                commitPresets(presets);
            });
        };

        // ---- プリセットを上書き（リストで選択中の項目を現在の選択の設定で更新）/ Overwrite the selected list item with the current selection's settings ----
        ui.overwritePresetButton.onClick = function () {
            var index = activePresetList().selection ? activePresetList().selection.index : -1;
            if (index < 0) return; // 未選択なら何もしない / Do nothing if nothing is selected
            buildPresetFromSelection(function (newPreset) {
                if (!newPreset) return;
                var presets = loadPresets();
                if (index >= presets.length) return;
                presets[index] = newPreset;
                commitPresets(presets);
            });
        };

        // ---- プリセットを書き出し（登録済みプリセットを JSON ファイルへ）/ Export presets to a JSON file ----
        ui.exportPresetButton.onClick = function () {
            var presets = loadPresets();
            var defaultFile = File(Folder.desktop.fsName + "/UnifiedTypePanel_presets.json");
            var saveFile = defaultFile.saveDlg(getLocalizedText(LABELS.tip.exportPreset), "JSON:*.json");
            if (!saveFile) return; // キャンセル / Cancelled
            try {
                saveFile.encoding = "UTF-8";
                if (saveFile.open("w")) {
                    saveFile.write(presetsToJsonText(presets));
                    saveFile.close();
                }
            } catch (e) {
                try { saveFile.close(); } catch (eClose) { }
                alert(getLocalizedText(LABELS.msg.applyError) + "\n" + e);
            }
        };

        // ---- プリセットを削除（リストで選択中の項目を削除）/ Delete the selected list item ----
        ui.deletePresetButton.onClick = function () {
            var index = activePresetList().selection ? activePresetList().selection.index : -1;
            if (index < 0) return; // 未選択なら何もしない / Do nothing if nothing is selected
            var presets = loadPresets();
            if (index >= presets.length) return;
            presets.splice(index, 1);
            commitPresets(presets);
        };

        // ---- 文字揃え（欧文ベースライン／中央／その他＋ポップアップ）/ Character alignment (Roman / Center / Other + popup) ----
        // ラジオは親が異なるため手動で排他制御 / Radios have different parents, so toggle them manually
        function selectAlignRadio(alignId) {
            var targetAlignId = (alignId === "roman" || alignId === "center") ? alignId : "other";
            for (var i = 0; i < ui.alignRadios.length; i++) {
                ui.alignRadios[i].value = (ui.alignRadios[i].alignId === targetAlignId);
            }
        }
        function applyAlignOther() {
            if (suppressUiEvents) return;
            var selectedOption = ui.alignOtherDropdown.selection;
            if (!selectedOption) return;
            selectAlignRadio("other");
            runApply("applyAlign", { align: ui.alignOtherOptions[selectedOption.index].id });
        }
        ui.alignRomanRadio.onClick = function () { selectAlignRadio("roman"); runApply("applyAlign", { align: "roman" }); };
        ui.alignCenterRadio.onClick = function () { selectAlignRadio("center"); runApply("applyAlign", { align: "center" }); };
        ui.alignOtherRadio.onClick = applyAlignOther;
        ui.alignOtherDropdown.onChange = applyAlignOther;

        // ---- 行揃え（グラフィカルボタン）/ Paragraph justification (graphical buttons) ----
        // アクティブ表示を切り替えて再描画 / Switch the active button and redraw
        function setActiveJustify(justifyId) {
            ui.justifyState.activeId = justifyId || "";
            for (var i = 0; i < ui.justifyButtons.length; i++) {
                try { ui.justifyButtons[i].notify("onDraw"); } catch (e) { }
            }
            try { ui.palette.update(); } catch (e2) { }
        }
        for (var buttonIndex = 0; buttonIndex < ui.justifyButtons.length; buttonIndex++) {
            ui.justifyButtons[buttonIndex].onClick = function () {
                setActiveJustify(this.justifyId);
                runApply("applyJustification", { justify: this.justifyId });
            };
        }

        // ---- 種別（本文／見出し）とリセット / Type roles (Body / Heading) and Reset ----
        // id 一致するラジオを選択 / Select the radio whose option id matches
        function selectRadioById(radios, options, targetId) {
            for (var i = 0; i < options.length; i++) {
                if (options[i].id === targetId) { selectExclusiveRadio(radios, i); return; }
            }
        }
        function reflectTsume(value) {
            ui.tsumeSlider.value = value;
            ui.tsumeInput.text = String(value);
        }
        function reflectTracking(value) {
            ui.trackingSlider.value = value;
            ui.trackingInput.text = String(value);
        }
        // 文字組みアキ量設定ポップアップを mojikumiSet インデックスで選択 / Select the mojikumi popup by mojikumiSet index
        function reflectMojikumiByIndex(mojikumiIndex) {
            for (var i = 0; i < ui.mojikumiChoices.length; i++) {
                if (ui.mojikumiChoices[i].index === mojikumiIndex) { ui.mojikumiDropdown.selection = i; return; }
            }
        }
        // 禁則ポップアップを id（"None"/"Hard"/"Soft"/"Soft_v2"）で選択 / Select the kinsoku popup by id
        function reflectKinsokuById(kinsokuId) {
            for (var i = 0; i < ui.kinsokuChoices.length; i++) {
                if (ui.kinsokuChoices[i].id === kinsokuId) {
                    suppressUiEvents = true;
                    ui.kinsokuDropdown.selection = i;
                    suppressUiEvents = false;
                    return;
                }
            }
        }

        // 本文：和文等幅／行送り 150%／均等配置（最終行左）／文字組みベタ組み / Body preset (mojikumi: Solid)
        ui.roleBodyRadio.onClick = function () {
            ui.mainTabs.selection = ui.typeTab; // 文字組みタブへ切り替え / Switch to the Type tab
            runApply("applyProfile", { kern: "mono", leadingPercent: 150, justify: "justifyLeft", mojikumiIndex: 6, kinsoku: "Soft_v2" });
            selectRadioById(ui.kernRadios, autoKernOptions, "mono");
            reflectLeadingPercent(150);
            setActiveJustify("justifyLeft");
            reflectMojikumiByIndex(6); // ベタ組み / Solid
            reflectKinsokuById("Soft_v2"); // 弱い禁則 v2 / Loose v2
        };

        // 見出し：メトリクス／行送り 115%／左揃え／文字組みツメ組み／弱い禁則 v2 / Heading preset (mojikumi: Tight, kinsoku: Soft_v2)
        ui.roleHeadingRadio.onClick = function () {
            ui.mainTabs.selection = ui.typeTab; // 文字組みタブへ切り替え / Switch to the Type tab
            runApply("applyProfile", { kern: "metrics", leadingPercent: 115, justify: "left", mojikumiIndex: 5, kinsoku: "Soft_v2" });
            selectRadioById(ui.kernRadios, autoKernOptions, "metrics");
            reflectLeadingPercent(115);
            setActiveJustify("left");
            reflectMojikumiByIndex(5); // ツメ組み / Tight
            reflectKinsokuById("Soft_v2"); // 弱い禁則 v2 / Loose v2
        };

        // 方針：文字揃え＋行送りの基準をまとめて適用し、対応するラジオも反映
        // Policy: apply char-alignment + leading-basis together and reflect the corresponding radios
        for (var policyRadioIndex = 0; policyRadioIndex < ui.policyRadios.length; policyRadioIndex++) {
            (function (index) {
                ui.policyRadios[index].onClick = function () {
                    ui.mainTabs.selection = ui.typeTab; // 文字組みタブへ切り替え / Switch to the Type tab
                    var preset = ui.policyRadios[index].policyPreset;
                    selectExclusiveRadio(ui.policyRadios, index);
                    runApply("applyPolicy", { align: preset.align, leadingType: preset.leadingType, hyphenation: preset.hyphenation });
                    selectAlignRadio(preset.align);
                    selectRadioById(ui.leadingBasisRadios, ui.leadingBasisChoices, preset.leadingType);
                };
            })(policyRadioIndex);
        }

        // 制御文字の表示／非表示 / Toggle hidden characters
        ui.hiddenCharButton.onClick = function () {
            runApply("toggleHiddenChar", null);
        };

        // リセット：すべて既定値へ / Reset everything to defaults
        ui.resetButton.onClick = function () {
            runApply("applyProfile", {
                kern: "metrics", tsume: 0, tracking: 0, align: "roman", justify: "left",
                leadingPercent: 115, leadingType: "top",
                sizePt: 12, scale: 100, mojikumiIndex: 5, kinsoku: "Soft_v2", clearAki: true
            });
            selectRadioById(ui.kernRadios, autoKernOptions, "metrics");
            reflectTsume(0);
            reflectTracking(0);
            selectAlignRadio("roman");
            setActiveJustify("left");
            reflectLeadingPercent(115);
            selectRadioById(ui.leadingBasisRadios, ui.leadingBasisChoices, "top");
            // 文字組みアキ量設定を「ツメ組み」（mojikumiSet の index 5）へ / Mojikumi to Tight (index 5)
            suppressUiEvents = true;
            for (var mojikumiChoiceIndex = 0; mojikumiChoiceIndex < ui.mojikumiChoices.length; mojikumiChoiceIndex++) {
                if (ui.mojikumiChoices[mojikumiChoiceIndex].index === 5) { ui.mojikumiDropdown.selection = mojikumiChoiceIndex; break; }
            }
            // 禁則を「弱い禁則 v2」（Soft_v2）へ / Kinsoku to Loose v2 (Soft_v2)
            for (var kinsokuChoiceIndex = 0; kinsokuChoiceIndex < ui.kinsokuChoices.length; kinsokuChoiceIndex++) {
                if (ui.kinsokuChoices[kinsokuChoiceIndex].id === "Soft_v2") { ui.kinsokuDropdown.selection = kinsokuChoiceIndex; break; }
            }
            suppressUiEvents = false;
            // 文字サイズ 12pt・比率 100% / Font size 12pt, scale 100%
            syncFontSizeText(String(Math.round((12 / textUnit.factor) * 10) / 10), null);
            ui.scaleInput.text = "100";
            updateApparentDisplay();
            updateReferenceApparent();
            updateLeadingEffective(); // サイズ確定後に実質行送りを更新 / Refresh the effective leading after the size is set
            ui.roleBodyRadio.value = false;
            ui.roleHeadingRadio.value = false;
            for (var policyResetIndex = 0; policyResetIndex < ui.policyRadios.length; policyResetIndex++) ui.policyRadios[policyResetIndex].value = false;
        };

        // 自動カーニングに合わせてプロポーショナルメトリクスのチェックを更新（メトリクスのみ ON）
        // Sync the proportional-metrics checkbox with the kerning choice (ON only for Metrics)
        function syncProportionalMetricsCheck(methodId) {
            if (ui.proportionalMetricsCheck) ui.proportionalMetricsCheck.value = (methodId === "metrics");
        }

        // ---- カーニング・文字ツメ / Kerning and Tsume ----
        for (var i = 0; i < ui.kernRadios.length; i++) {
            ui.kernRadios[i].onClick = function () {
                syncProportionalMetricsCheck(autoKernOptions[this.index].id);
                runApply("apply", { method: autoKernOptions[this.index].id });
            };
        }
        // 文字ツメスライダー（Shift 併用で 10% 刻み）/ Tsume slider (Shift snaps to 10% steps)
        function tsumeSnap(value) {
            var isShiftPressed = false;
            try { isShiftPressed = ScriptUI.environment.keyboardState.shiftKey; } catch (e) { }
            return isShiftPressed ? Math.round(value / 10) * 10 : Math.round(value);
        }
        ui.tsumeSlider.onChanging = function () {
            var snappedValue = tsumeSnap(this.value);
            this.value = snappedValue;
            ui.tsumeInput.text = String(snappedValue);
        };
        ui.tsumeSlider.onChange = function () {
            var snappedValue = tsumeSnap(this.value);
            this.value = snappedValue;
            ui.tsumeInput.text = String(snappedValue);
            runApply("applyTsume", { value: snappedValue });
        };
        // 文字ツメ入力欄：0〜100 にクランプしてスライダーへ反映 / Tsume input: clamp 0-100 and sync the slider
        function applyTsumeFromInput() {
            var clampedValue = Math.round(parseFloat(ui.tsumeInput.text));
            if (isNaN(clampedValue)) return;
            if (clampedValue < 0) clampedValue = 0; else if (clampedValue > 100) clampedValue = 100;
            ui.tsumeInput.text = String(clampedValue);
            ui.tsumeSlider.value = clampedValue;
            runApply("applyTsume", { value: clampedValue });
        }
        ui.tsumeInput.onChange = applyTsumeFromInput;
        changeValueByArrowKey(ui.tsumeInput, false, applyTsumeFromInput, 0);

        // トラッキングスライダー（-100〜500、Shift 併用で 10 単位）/ Tracking slider (-100 to 500, Shift snaps to 10 steps)
        function trackingSnap(value) {
            var isShiftPressed = false;
            try { isShiftPressed = ScriptUI.environment.keyboardState.shiftKey; } catch (e) { }
            return isShiftPressed ? Math.round(value / 10) * 10 : Math.round(value);
        }
        ui.trackingSlider.onChanging = function () {
            var snappedValue = trackingSnap(this.value);
            this.value = snappedValue;
            ui.trackingInput.text = String(snappedValue);
        };
        ui.trackingSlider.onChange = function () {
            var snappedValue = trackingSnap(this.value);
            this.value = snappedValue;
            ui.trackingInput.text = String(snappedValue);
            runApply("applyTracking", { tracking: snappedValue });
        };
        // トラッキング入力欄：-100〜500 にクランプしてスライダーへ反映 / Tracking input: clamp -100..500 and sync the slider
        function applyTrackingFromInput() {
            var clampedValue = Math.round(parseFloat(ui.trackingInput.text));
            if (isNaN(clampedValue)) return;
            if (clampedValue < -100) clampedValue = -100; else if (clampedValue > 500) clampedValue = 500;
            ui.trackingInput.text = String(clampedValue);
            ui.trackingSlider.value = clampedValue;
            runApply("applyTracking", { tracking: clampedValue });
        }
        ui.trackingInput.onChange = applyTrackingFromInput;
        changeValueByArrowKey(ui.trackingInput, true, applyTrackingFromInput, 0);

        // ---- フォントサイズ（文字サイズ・比率・見かけ）/ Font size (size, scale, apparent) ----
        // 文字組みタブと「フォントサイズ」タブの2つのサイズ欄を同期させる / Keep the two size fields (Type tab + Font Size tab) in sync
        function calcApparentSize(size, scale) { return Math.round(size * scale) / 100; }
        /* 全サイズ欄へ値を反映（except は編集中の欄を除外してキャレット移動を防ぐ）/ Push a value to every size field (except the one being edited, to avoid caret jumps) */
        function syncFontSizeText(text, except) {
            for (var i = 0; i < ui.fontSizeInputs.length; i++) {
                if (ui.fontSizeInputs[i] !== except) ui.fontSizeInputs[i].text = text;
            }
        }
        function updateApparentDisplay() {
            var size = parseFloat(ui.fontSizeInputs[0].text);
            var scale = parseFloat(ui.scaleInput.text);
            if (isNaN(size) || isNaN(scale)) {
                ui.apparentValueLabel.text = "--";
            } else {
                ui.apparentValueLabel.text = String(Math.round(calcApparentSize(size, scale) * 10) / 10);
            }
            // 比率100%のときは見かけ＝実サイズなので淡色表示 / Dim when scale is 100 (apparent == actual)
            var isDimmed = (scale === 100);
            ui.fsApparentLabel.enabled = !isDimmed;
            ui.apparentValueLabel.enabled = !isDimmed;
            ui.fsApparentUnit.enabled = !isDimmed;
        }
        // サイズ変更では基準見かけサイズ（referenceApparent）を更新しない＝ボタンで元の見かけへ戻せるように保持
        // A size edit does NOT update the reference apparent size, so the button can restore the previous look
        function applyFontSizeFromInput(sourceInput) {
            syncFontSizeText(sourceInput.text, sourceInput);
            var inputValue = parseFloat(sourceInput.text);
            // 行送りは自動行送りなのでサイズ変更に自動追従する（行送りの再適用は不要）
            // Leading is auto-leading and follows the size automatically (no need to reapply leading)
            if (!isNaN(inputValue)) runApply("applyFontSize", { sizePt: inputValue * textUnit.factor });
            updateApparentDisplay();
            updateLeadingEffective(); // サイズ変更で実質行送りの表示も更新 / Refresh the effective-leading display on size change
        }
        // 比率の変更は「見た目を変える」操作なので基準見かけサイズを更新 / A scale edit changes the look, so refresh the reference
        function applyScaleFromInput() {
            var inputValue = parseFloat(ui.scaleInput.text);
            if (!isNaN(inputValue)) runApply("applyScale", { scale: inputValue });
            updateApparentDisplay();
            updateReferenceApparent();
        }
        for (var fontSizeInputIndex = 0; fontSizeInputIndex < ui.fontSizeInputs.length; fontSizeInputIndex++) {
            (function (sizeInput) {
                sizeInput.onChange = function () { applyFontSizeFromInput(sizeInput); };
                sizeInput.onChanging = function () { syncFontSizeText(sizeInput.text, sizeInput); updateApparentDisplay(); updateLeadingEffective(); };
                changeValueByArrowKey(sizeInput, false, function () { applyFontSizeFromInput(sizeInput); }, 1);
            })(ui.fontSizeInputs[fontSizeInputIndex]);
        }
        ui.scaleInput.onChange = applyScaleFromInput;
        ui.scaleInput.onChanging = function () { updateApparentDisplay(); };
        changeValueByArrowKey(ui.scaleInput, false, applyScaleFromInput, 1);

        // ---- タイプスケール（サイズタブ右カラム）/ Type scale (right column of the Size tab) ----
        // 一覧の行をクリックしたら、そのサイズを選択テキストへ適用しサイズ欄にも反映
        // Click a row to apply that size to the selection and reflect it in the size fields
        ui.typeScaleList.onChange = function () {
            var selectedItem = ui.typeScaleList.selection;
            if (!selectedItem) return;
            var sizeValue = parseFloat(selectedItem.subItems[0].text);
            if (isNaN(sizeValue) || sizeValue <= 0) return;
            syncFontSizeText(String(sizeValue), null);
            updateApparentDisplay();
            updateReferenceApparent(); // 選んだサイズを新しい見た目の基準に / The picked size becomes the new visual reference
            updateLeadingEffective(); // サイズ変更で実質行送りの表示も更新 / Refresh the effective-leading display on size change
            // タイプスケールは段落全体にサイズを適用（行送りは自動行送りなので自動追従）
            // Type scale applies the size to the whole paragraph (leading is auto and follows automatically)
            runApply("applyFontSizePara", { sizePt: sizeValue * textUnit.factor });
        };
        // 基準サイズ欄が空なら現在のフォントサイズ（なければ既定値）で初期化 / Seed the base field from the current font size (fallback to default)
        if (ui.typeScaleBaseInput.text === "") {
            var seedBase = parseFloat(ui.fontSizeInputs[0].text);
            ui.typeScaleBaseInput.text = (!isNaN(seedBase) && seedBase > 0) ? String(seedBase) : String(Math.round((12 / textUnit.factor) * 10) / 10);
            ui.rebuildTypeScaleList();
        }

        // 見かけ←→実サイズ（相互変換）：見た目の大きさ（実質サイズ＝サイズ×比率）を保ったまま、
        //   ・比率≠100%：現在の見かけを実サイズへ焼き込み、比率100%に（見かけ→実サイズ）
        //   ・比率＝100%：サイズ変更後、変更前の見かけサイズを基準に比率を計算して合わせる（実サイズ→見かけ）
        // referenceApparent に「保つべき見かけサイズ（表示単位）」を保持。サイズ変更では更新しないので元の見た目へ戻せる
        // Apparent ←→ Actual: keep the visual size (size × scale) constant.
        //   - scale != 100: bake the current apparent into the actual size, scale → 100% (apparent → actual)
        //   - scale == 100: after a size change, compute the scale that restores the pre-change apparent (actual → apparent)
        // referenceApparent holds the apparent size to preserve (display units); size edits don't touch it, so the look can be restored.
        var referenceApparent = NaN;
        function updateReferenceApparent() {
            var size = parseFloat(ui.fontSizeInputs[0].text);
            var scale = parseFloat(ui.scaleInput.text);
            referenceApparent = (isNaN(size) || isNaN(scale)) ? NaN : calcApparentSize(size, scale);
        }
        ui.apparentToggleButton.onClick = function () {
            var size = parseFloat(ui.fontSizeInputs[0].text);
            var scale = parseFloat(ui.scaleInput.text);
            if (isNaN(size) || isNaN(scale)) return;
            var EPS = 0.05;
            if (Math.abs(scale - 100) > EPS) {
                // 見かけ→実サイズ：現在の見かけ（サイズ×比率）を実サイズへ焼き込み、比率を100%に
                // Apparent → Actual: bake the current apparent into the actual size, scale → 100%
                var bakedSize = calcApparentSize(size, scale);
                syncFontSizeText(String(bakedSize), null);
                ui.scaleInput.text = "100";
                runApply("applyFontSizeAndScale", { sizePt: bakedSize * textUnit.factor, scale: 100 });
            } else {
                // 実サイズ→見かけ：サイズ変更後、変更前の見かけサイズを基準に比率を計算して合わせる
                // Actual → Apparent: after a size change, compute the scale that restores the previous apparent size
                if (isNaN(referenceApparent) || Math.abs(referenceApparent - size) < EPS) return; // 変更なし＝何もしない / no change → no-op
                var newScale = Math.round((referenceApparent / size) * 100 * 10) / 10;
                ui.scaleInput.text = String(newScale);
                runApply("applyFontSizeAndScale", { sizePt: size * textUnit.factor, scale: newScale });
            }
            updateReferenceApparent(); // 変換後の見かけを新しい基準に / The post-conversion apparent becomes the new reference
            updateApparentDisplay();
        };

        // ---- 行送り / Leading ----

        /* 現在の行送り（自動行送り量 %）を取得 / Read the current leading (auto-leading amount %) */
        function currentLeadingPercent() {
            return parseFloat(ui.leadingPercentInput.text);
        }

        /* 実質行送り（フォントサイズ×%）の表示を更新 / Update the effective-leading display (font size × %) */
        function updateLeadingEffective() {
            var size = parseFloat(ui.fontSizeInputs[0].text);
            var percent = currentLeadingPercent();
            if (isNaN(size) || isNaN(percent)) { ui.leadingEffectiveInput.text = ""; return; }
            ui.leadingEffectiveInput.text = String(Math.round(size * percent / 100 * 10) / 10);
        }

        /* 行送り（%）欄へ値を反映し、実質表示も更新 / Set the leading (%) field and refresh the effective display */
        function reflectLeadingPercent(percent) {
            ui.leadingPercentInput.text = isNaN(percent) ? "" : String(Math.round(percent * 10) / 10);
            updateLeadingEffective();
        }

        /* UI の現在状態から行送りパラメータ（自動行送り量%＋基準）を組み立てる（% が空なら null）
           Build leading params (auto-leading % + basis) from the current UI (null if the % is empty) */
        function currentLeadingParams() {
            var percent = currentLeadingPercent();
            if (isNaN(percent)) return null;
            var basisIndex = selectedRadioIndex(ui.leadingBasisRadios);
            if (basisIndex < 0) basisIndex = 0;
            return {
                percent: percent,
                leadingType: ui.leadingBasisChoices[basisIndex].id
            };
        }

        function applyLeading() {
            var params = currentLeadingParams();
            if (!params) return;
            if (workerBusy) return;
            workerBusy = true;
            runWorker("applyLeading", params, function (status, payload) {
                workerBusy = false;
            });
        }

        // 行送り（%）入力：実質表示を更新して適用 / Leading (%) input: refresh the effective display and apply
        function applyLeadingFromPercent() {
            updateLeadingEffective();
            applyLeading();
        }
        ui.leadingPercentInput.onChange = applyLeadingFromPercent;
        ui.leadingPercentInput.onChanging = function () { updateLeadingEffective(); };
        changeValueByArrowKey(ui.leadingPercentInput, false, applyLeadingFromPercent, 1);

        // 実質（pt）入力：フォントサイズから % を逆算して適用 / Effective (pt) input: back-calculate the % from the font size and apply
        function applyLeadingFromEffective() {
            var effective = parseFloat(ui.leadingEffectiveInput.text);
            var size = parseFloat(ui.fontSizeInputs[0].text);
            if (isNaN(effective) || isNaN(size) || size <= 0) return;
            var percent = Math.round((effective / size) * 100 * 10) / 10;
            ui.leadingPercentInput.text = String(percent);
            applyLeading();
        }
        ui.leadingEffectiveInput.onChange = applyLeadingFromEffective;
        changeValueByArrowKey(ui.leadingEffectiveInput, false, applyLeadingFromEffective, 1);

        for (var basisRadioIndex = 0; basisRadioIndex < ui.leadingBasisRadios.length; basisRadioIndex++) {
            (function (index) {
                ui.leadingBasisRadios[index].onClick = function () {
                    selectExclusiveRadio(ui.leadingBasisRadios, index);
                    applyLeading();
                };
            })(basisRadioIndex);
        }

        // 文字組みアキ量設定（ポップアップ）/ Mojikumi spacing set (popup)
        ui.mojikumiDropdown.onChange = function () {
            if (suppressUiEvents || !this.selection) return;
            runApply("applyMojikumi", { mojikumiIndex: ui.mojikumiChoices[this.selection.index].index });
        };

        // 禁則（ポップアップ）/ Kinsoku (popup)
        ui.kinsokuDropdown.onChange = function () {
            if (suppressUiEvents || !this.selection) return;
            runApply("applyKinsoku", { kinsoku: ui.kinsokuChoices[this.selection.index].id });
        };

        /* 選択状態を読み取って UI に反映 / Read the selection state and reflect it into the UI */
        function refreshState() {
            runWorker("getState", null, function (status, payload) {
                if (status !== "ok") return;
                var state = parseLeadingState(payload);
                if (state.count <= 0) return;
                // フォントサイズ・比率・見かけ / Font size, scale, apparent
                syncFontSizeText(isNaN(state.fontSizePt) ? "" : String(Math.round((state.fontSizePt / textUnit.factor) * 10) / 10), null);
                ui.scaleInput.text = isNaN(state.hScale) ? "100" : String(Math.round(state.hScale * 10) / 10);
                updateApparentDisplay();
                updateReferenceApparent(); // 選択を読み直したら現在の見かけを基準にする / Reloading the selection sets the current apparent as the reference
                // 行送り：常に自動行送りなので、自動行送り量（%）を % 欄に反映し実質も更新
                // Leading: always auto-leading, so reflect the auto-leading amount (%) into the % field and refresh the effective value
                reflectLeadingPercent(isNaN(state.autoAmount) ? NaN : Math.round(state.autoAmount));
                // 行送りの基準 / Leading basis
                var basisIndex = 0;
                for (var choiceIndex = 0; choiceIndex < ui.leadingBasisChoices.length; choiceIndex++) {
                    if (ui.leadingBasisChoices[choiceIndex].id === state.leadingTypeId) { basisIndex = choiceIndex; break; }
                }
                selectExclusiveRadio(ui.leadingBasisRadios, basisIndex);
                // 文字組みアキ量設定（-2＝読み取り不可・不明のときは現在の選択を維持）/ Mojikumi (leave the current selection when -2 = unreadable/unknown)
                if (state.mojikumiId !== -2) {
                    var mojikumiSelectionIndex = 0;
                    for (var mojikumiChoiceIndex = 0; mojikumiChoiceIndex < ui.mojikumiChoices.length; mojikumiChoiceIndex++) {
                        if (ui.mojikumiChoices[mojikumiChoiceIndex].index === state.mojikumiId) { mojikumiSelectionIndex = mojikumiChoiceIndex; break; }
                    }
                    suppressUiEvents = true;
                    ui.mojikumiDropdown.selection = mojikumiSelectionIndex;
                    suppressUiEvents = false;
                }
                // 禁則 / Kinsoku
                var kinsokuSelectionIndex = 0;
                for (var kinsokuChoiceIndex = 0; kinsokuChoiceIndex < ui.kinsokuChoices.length; kinsokuChoiceIndex++) {
                    if (ui.kinsokuChoices[kinsokuChoiceIndex].id === state.kinsokuId) { kinsokuSelectionIndex = kinsokuChoiceIndex; break; }
                }
                suppressUiEvents = true;
                ui.kinsokuDropdown.selection = kinsokuSelectionIndex;
                suppressUiEvents = false;
                // 文字揃え / Character alignment
                if (state.alignmentId) {
                    selectAlignRadio(state.alignmentId);
                    // その他系は対応するポップアップ項目も選択 / For "Other" ids, also select the popup item
                    if (state.alignmentId !== "roman" && state.alignmentId !== "center") {
                        for (var optionIndex = 0; optionIndex < ui.alignOtherOptions.length; optionIndex++) {
                            if (ui.alignOtherOptions[optionIndex].id === state.alignmentId) {
                                suppressUiEvents = true;
                                ui.alignOtherDropdown.selection = optionIndex;
                                suppressUiEvents = false;
                                break;
                            }
                        }
                    }
                }
                // 行揃え / Paragraph justification
                setActiveJustify(state.justifyId || "");
            });
        }

        /* プリセットの表示名を和文フォント名に差し替え（1回だけ、両ビュー）/ Replace preset labels with Japanese font names (once, both views) */
        var presetLabelsResolved = false;
        function refreshPresetLabels() {
            if (presetLabelsResolved) return;
            var foItems = ui.presetFontOnlyListBox.items, dItems = ui.presetDetailedListBox.items;
            var psList = "";
            for (var i = 0; i < foItems.length; i++) {
                psList += (i ? String.fromCharCode(10) : "") + foItems[i].psName;
            }
            runWorker("resolveFontNames", { psList: psList }, function (status, payload) {
                if (status !== "ok") return;
                var resolvedData = decodeURIComponent(payload);
                if (resolvedData === "") return;
                var rows = resolvedData.split(String.fromCharCode(10));
                for (var rowIndex = 0; rowIndex < rows.length && rowIndex < foItems.length; rowIndex++) {
                    var tabIndex = rows[rowIndex].indexOf(String.fromCharCode(9));
                    if (tabIndex >= 0) {
                        var displayName = rows[rowIndex].substring(tabIndex + 1);
                        foItems[rowIndex].text = displayName;           // フォントのみ / font-only
                        if (dItems[rowIndex]) dItems[rowIndex].text = displayName; // 詳細のフォント列 / detailed font column
                    }
                }
                presetLabelsResolved = true;
            });
        }

        // パレット表示・フォーカス復帰のたびに選択状況とフォント一覧を反映 / Reflect the selection and font lists on show and on regaining focus
        // UI 明暗を再取得して行揃えボタンを描き直す / Re-read the UI theme and repaint the justify buttons
        function refreshJustifyTheme() {
            ui.justifyState.isLight = isLightUI();
            setActiveJustify(ui.justifyState.activeId);
        }
        // ---- 使用フォント一覧（フォントのみ／詳細を stack で重ねた2つの listbox で切替）/ Used-fonts list (two stacked listboxes toggled by visibility) ----
        var currentCombos = [];
        var currentFonts = [];
        var comboApplyInProgress = false;

        /* フォントのみ表示か / Whether the font-only view is active */
        function isFontOnlyView() { return ui.fontOnlyRadio.value; }

        // 詳細：クリックで組み合わせを一括適用 / Detailed: a click applies the whole combination
        ui.comboListBox.onChange = function () {
            if (!ui.comboListBox.selection) return;
            if (comboApplyInProgress) return;
            var combo = currentCombos[ui.comboListBox.selection.comboIndex];
            if (!combo) return;
            comboApplyInProgress = true;
            runWorker("applyCombo", comboToParams(combo), function (status, payload) {
                comboApplyInProgress = false;
                // 選択が無いときだけ警告 / Warn only when nothing was selected
                if (status === "ok" && payload === "nosel") alert(getLocalizedText(LABELS.usedFonts.status.needSelection));
            });
        };

        // フォントのみ：クリックでフォントだけ適用（旧ドキュメントフォントと同じ）/ Font-only: a click applies just the font
        ui.fontOnlyListBox.onChange = function () {
            if (!ui.fontOnlyListBox.selection || !ui.fontOnlyListBox.selection.psName) return;
            if (comboApplyInProgress) return;
            comboApplyInProgress = true;
            runWorker("applyFont", { psName: ui.fontOnlyListBox.selection.psName }, function (status, payload) {
                comboApplyInProgress = false;
            });
        };

        /* 表示モードに応じて listbox を出し分け（stack なので visible 切替のみ）/ Show one listbox per mode (stack: just toggle visibility) */
        function applyUsedFontView() {
            var fontOnly = isFontOnlyView();
            ui.fontOnlyListBox.visible = fontOnly;
            ui.comboListBox.visible = !fontOnly;
        }

        /* 走査してデータ更新→両ビュー（詳細・フォントのみ）に反映 / Scan, cache, and repopulate both views */
        function refreshCombos() {
            runWorker("getCombos", { currentArtboardOnly: ui.currentArtboardCheckbox.value, includeHidden: ui.includeHiddenCheckbox.value, includeLocked: ui.includeLockedCheckbox.value }, function (status, payload) {
                currentCombos = (status === "ok") ? parseCombosPayload(decodeURIComponent(payload)) : [];
                currentFonts = combosToFonts(currentCombos);
                // 詳細（8列）/ Detailed (8 columns)
                ui.comboListBox.removeAll();
                for (var i = 0; i < currentCombos.length; i++) {
                    var cells = currentCombos[i].cells;
                    var listItem = ui.comboListBox.add("item", cells[0]);
                    for (var col = 1; col < cells.length; col++) listItem.subItems[col - 1].text = cells[col];
                    listItem.comboIndex = i;
                }
                // フォントのみ（単一列）/ Font-only (single column)
                ui.fontOnlyListBox.removeAll();
                for (var f = 0; f < currentFonts.length; f++) {
                    var fontItem = ui.fontOnlyListBox.add("item", currentFonts[f].displayName);
                    fontItem.psName = currentFonts[f].psName;
                }
            });
        }

        // 表示モード切替 / Switch view mode
        ui.fontOnlyRadio.onClick = function () { applyUsedFontView(); };
        ui.detailedRadio.onClick = function () { applyUsedFontView(); };

        // 選択中の条件に一致するテキストを選択（詳細＝組み合わせ一致／フォントのみ＝フォント一致）
        // Select matching text (detailed = full combo / font-only = same font)
        ui.selectMatchingButton.onClick = function () {
            var matchingParams;
            if (isFontOnlyView()) {
                if (!ui.fontOnlyListBox.selection || !ui.fontOnlyListBox.selection.psName) { alert(getLocalizedText(LABELS.usedFonts.status.needCondition)); return; }
                matchingParams = { psName: ui.fontOnlyListBox.selection.psName, matchFontOnly: true };
            } else {
                if (!ui.comboListBox.selection) { alert(getLocalizedText(LABELS.usedFonts.status.needCondition)); return; }
                var combo = currentCombos[ui.comboListBox.selection.comboIndex];
                if (!combo) return;
                matchingParams = comboToParams(combo);
            }
            matchingParams.currentArtboardOnly = ui.currentArtboardCheckbox.value;
            matchingParams.includeHidden = ui.includeHiddenCheckbox.value;
            matchingParams.includeLocked = ui.includeLockedCheckbox.value;
            runWorker("selectMatching", matchingParams, function (status, payload) {
                if (status === "ok" && payload === "0") alert(getLocalizedText(LABELS.usedFonts.status.noMatch));
            });
        };

        ui.refreshCombosButton.onClick = function () { refreshCombos(); };
        ui.currentArtboardCheckbox.onClick = function () { refreshCombos(); };
        ui.includeHiddenCheckbox.onClick = function () { refreshCombos(); };
        ui.includeLockedCheckbox.onClick = function () { refreshCombos(); };

        // 初期表示モードを反映（既定＝フォントのみ）/ Apply the initial view (default = font-only)
        applyUsedFontView();

        // 使用フォントタブに切り替えたら一覧を再走査（重い走査を必要時のみ）/ Rescan when switching to the used-fonts tab (scan only when needed)
        ui.mainTabs.onChange = function () {
            if (ui.mainTabs.selection === ui.usedFontsTab) refreshCombos();
        };

        ui.palette.onShow = function () { refreshJustifyTheme(); refreshState(); refreshPresetLabels(); if (ui.mainTabs.selection === ui.usedFontsTab) refreshCombos(); };
        ui.palette.onActivate = function () { refreshJustifyTheme(); refreshState(); refreshPresetLabels(); };

        // 再読み込み：選択中のテキストの現在値を読み直して UI に反映 / Reload: re-read the selection's current values into the UI
        ui.reloadButton.onClick = function () { refreshState(); refreshPresetLabels(); };

        // Esc でパレットを閉じる / Close the palette on Esc
        ui.palette.addEventListener("keydown", function (event) {
            if (event.keyName === "Escape") ui.palette.close();
        });

        return { refreshState: refreshState };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================
    function main() {
        var autoKernOptions = createAutoKernOptions();
        var alignOptions = createAlignOptions();
        var justifyOptions = createJustifyOptions();
        var ui = createPaletteUI(autoKernOptions, alignOptions, justifyOptions);
        var controller = bindPaletteEvents(ui, autoKernOptions, alignOptions, justifyOptions);

        ui.palette.show();

        // 表示（レイアウト確定）後にボタン高さをまとめて共通高さへ / Fit button heights after show (layout is final)
        fitAllButtonHeights(ui);
    }

    main();

})();
