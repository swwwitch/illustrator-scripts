#targetengine "UnifiedTypePanelEngine"
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストの文字組み設定（フォント・フォントサイズ・自動カーニング・字間・文字揃え・行揃え・行送り・
文字組みアキ量設定）をまとめて行う常駐パレットスクリプトです。インフォバーの直下に「種別」（本文／見出し・
和文／欧文／混植）を常設し、その下は4つのタブ構成（「フォント」：ドキュメントフォント、「プリセット」：プリセット、
「文字組み」＝中央：フォントサイズ・自動カーニング・字間調整／右：文字揃え・行送り・行送りの基準、
「日本語組版」：行揃え・禁則・文字組みアキ量設定）。

- 最上部にインフォバー：テキストが1つだけ選択されているとき、フォント名・スタイル・サイズ・行送り・行揃え・自動カーニング・禁則を表示

- ドキュメントフォント：書類で使用中のフォントを一覧表示。選ぶと選択テキストへ適用
- プリセット：フォント＋カーニング・文字ツメ・トラッキングをまとめて登録。選ぶと一括適用。
  「追加」で現在の選択の設定を保存、「上書き」で選択中のプリセットを更新、「削除」で除去（JSON で Folder.userData に永続化）
- フォントサイズ：サイズ・比率（水平／垂直を同値）・実質（サイズ×比率の表示）
- 自動カーニング：和文等幅／0／メトリクス／オプティカル（「メトリクス」のみプロポーショナルメトリクスON）
- 字間調整：文字ツメ（0〜100%）とトラッキング（-100〜500）。入力欄とスライダー（Shift で粗い刻み）
- 文字揃え：欧文ベースライン／中央／その他（仮想ボディの上下・平均字面の上下をポップアップで選択）
- 行揃え：左／中央／右／均等配置（最終行左）／両端揃え（テキストの見た目の位置を保持して適用）
- 種別：本文（文字組みベタ組み）／見出し（文字組みツメ組み）。よく使う組み合わせを一括適用。
  さらに 和文（中央揃え／基準＝仮想ボディの上）・欧文（欧文ベースライン揃え／基準＝欧文ベースライン）・
  混植（欧文ベースライン揃え／基準＝仮想ボディの上）を文字揃え＋行送りの基準としてまとめて適用
- 行送り：110%／125%／150%／その他（%で直接入力）／自動、個別・共通、行送りの基準。
  「その他」を選ぶと現在の行送りを % で補完。「自動の値を変更」ボタンで自動行送りの割合を変更。
  基準サイズの取り方は「個別」（各行ごと）／「共通」（フレーム全体の最頻サイズ）をラジオで選択
- 文字組みアキ量設定：なし／約物半角／約物全角／ツメ組み／ベタ組み などをラジオボタンで一括適用（「日本語組版」タブ）
- 禁則：なし／強い禁則／弱い禁則／弱い禁則 v2 をラジオボタンで一括適用（「なし」は scripting では設定不可）（「日本語組版」タブ）
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
A "Type" role box (Body/Heading and Japanese/Roman/Mixed) sits just below the info bar; below it
are four tabs ("Fonts": document fonts; "Presets": presets; "Type" = center: font size, kerning,
Tsume / right: character alignment, leading, basis; "Japanese Composition": justification,
kinsoku, mojikumi).

- Top info bar: when exactly one text is selected, shows font family, style, size, leading, justification, auto kerning, and kinsoku

- Document fonts: lists the fonts used in the document; clicking one applies it to the selection
- Presets: a font plus kerning / Tsume / tracking, applied together when clicked. "Add"
  saves the current selection's settings, "Overwrite" updates the selected preset, "Delete" removes it (persisted as JSON under Folder.userData)
- Font size: size, scale (horizontal/vertical set together), effective (size × scale, shown)
- Auto kerning: Metrics - Roman Only / 0 / Metrics / Optical (proportional metrics ON only for "Metrics")
- Letter spacing: Tsume (0–100%) and tracking (-100 to 500), via input fields and sliders (Shift = coarse steps)
- Character alignment: Roman baseline / center / Other (embox top-bottom & ICF box top-bottom via popup)
- Justification: left / center / right / justify (last left) / justify all (applied while keeping the text's visual position)
- Type: Body (solid mojikumi) / Heading (tight mojikumi); applies common combinations at once.
  Also Japanese (center align / top-to-top basis), Roman (Roman-baseline align / baseline basis),
  and Mixed (Roman-baseline align / top-to-top basis) as a one-shot alignment + leading-basis preset
- Leading: 110% / 125% / 150% / Other (enter a %) / Auto, individual vs common, and leading basis.
  Individual (per-line base size) vs common (the frame's most frequent size) is chosen with radios
  (Leading choices: 110% / 125% / 150% / Other / Auto)
  Choosing "Other" prefills the current leading as %; "Change auto value" edits the auto-leading percentage
- Mojikumi: None / half-width punctuation / full-width punctuation / tight / solid, applied together via radio buttons (on the "Japanese Composition" tab)
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
var SCRIPT_VERSION = "v1.1.0";

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
            docFonts: { ja: "フォント", en: "Fonts" },
            presets: { ja: "プリセット", en: "Presets" },
            type: { ja: "文字組み", en: "Type" },
            jpComposition: { ja: "日本語組版", en: "Japanese Composition" }
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
            spacingAdjust: { ja: "字間調整", en: "Letter Spacing" },
            align: { ja: "文字揃え", en: "Char Alignment" },
            justify: { ja: "行揃え", en: "Justification" },
            role: { ja: "種別", en: "Type" },
            scriptType: { ja: "文字種", en: "Script" },
            leading: { ja: "行送り", en: "Leading" },
            leadingType: { ja: "行送りの基準", en: "Leading basis" },
            mojikumi: { ja: "文字組みアキ量設定", en: "Mojikumi" },
            kinsoku: { ja: "禁則", en: "Kinsoku" }
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
            other: { ja: "その他", en: "Other" },
            auto: { ja: "自動", en: "Auto" },
            perLine: { ja: "個別", en: "Individual" },
            common: { ja: "共通", en: "Common value" },
            tipPerLine: {
                ja: "各行ごとに、その行の先頭数文字の最頻サイズを基準にします。サイズが混在するフレームでは、行ごとに行送りが変わります。",
                en: "Use each line's own base size (mode of its first few characters). Leading varies per line when sizes are mixed."
            },
            tipCommon: {
                ja: "フレーム全体の最頻サイズを1つの基準にして、全行に同じ行送りを適用します。",
                en: "Use a single base size for the whole frame (its most frequent size) so every line gets the same leading."
            },
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
            heading: { ja: "見出し", en: "Heading" },
            japanese: { ja: "和文", en: "Japanese" },
            roman: { ja: "欧文", en: "Roman" },
            mixed: { ja: "混植", en: "Mixed" }
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
            reset: { ja: "リセット", en: "Reset" },
            reload: { ja: "再読み込み", en: "Reload" },
            hiddenChar: { ja: "制御文字", en: "Hidden Chars" },
            changeAuto: { ja: "自動行送りの値", en: "Auto-leading value" }
        },
        tip: {
            docFonts: { ja: "書類で使用中のフォント一覧。選ぶと選択テキストに適用します。", en: "Fonts used in the document. Click one to apply it to the selection." },
            presets: { ja: "登録した文字組みプリセット（フォント＋カーニング・ツメ・トラッキング）。選ぶと一括適用します。", en: "Saved type presets (font + kerning / Tsume / tracking). Click one to apply them together." },
            addPreset: { ja: "現在の選択の設定をプリセットとして保存します（同じフォントは上書き）。", en: "Save the current selection's settings as a preset (same font overwrites)." },
            overwritePreset: { ja: "リストで選択中のプリセットを、現在の選択の設定で上書きします。", en: "Overwrite the preset selected in the list with the current selection's settings." },
            deletePreset: { ja: "リストで選択中のプリセットを削除します。", en: "Delete the preset selected in the list." },
            fontSize: { ja: "選択テキストのフォントサイズ。", en: "Font size of the selection." },
            scale: { ja: "水平比率・垂直比率を同じ値でまとめて設定します。", en: "Sets the horizontal and vertical scale together to the same value." },
            apparent: { ja: "フォントサイズ×比率で計算した、実質的なサイズです。", en: "The effective size, computed as font size × scale." },
            autoKern: { ja: "自動カーニング方式（和文等幅／0／メトリクス／オプティカル）。", en: "Auto-kerning method (Japanese equal width / 0 / Metrics / Optical)." },
            tsume: { ja: "文字ツメ（0〜100%）。隣接する文字の食い込み量。", en: "Tsume (0–100%): how much adjacent characters tighten." },
            tracking: { ja: "字間（1/1000em）をまとめて調整します（-100〜500）。", en: "Adjusts overall letter spacing in 1/1000 em (-100 to 500)." },
            align: { ja: "文字の縦方向の揃え基準。「その他」で残りの基準をポップアップから選びます。", en: "Vertical alignment basis. Use “Other” to pick the rest from the popup." },
            justify: { ja: "段落の行揃え。テキストの見た目の位置を保ったまま変更します。", en: "Paragraph justification, applied while keeping the text's visual position." },
            role: { ja: "本文＝和文等幅／行送り150%／均等配置（最終行左）、見出し＝メトリクス／行送り110%／左揃え をまとめて適用。", en: "Body = Japanese equal width / 150% leading / justify (last left); Heading = Metrics / 110% leading / left." },
            leading: { ja: "行送り。110%／125%／150%／その他（%で直接入力）／自動。", en: "Leading: 110% / 125% / 150% / Other (enter a %) / Auto." },
            leadingType: { ja: "行送りを測る基準位置（仮想ボディの上／欧文ベースライン）。", en: "The reference position for measuring leading (virtual body top / Roman baseline)." },
            mojikumi: { ja: "段落の文字組みアキ量設定（約物の詰め方など）をまとめて適用します。", en: "Applies a mojikumi spacing set (punctuation spacing, etc.) to the paragraphs." },
            kinsoku: { ja: "段落の禁則処理（強い禁則／弱い禁則など）をまとめて適用します。", en: "Applies a kinsoku (line-break) set to the paragraphs." },
            changeAuto: { ja: "自動行送りの割合（%）を変更します。クリックして値を入力してください。", en: "Change the auto-leading percentage. Click to enter a new value." },
            reset: { ja: "標準値に戻します（フォントサイズ12pt・比率100%・ツメ0・トラッキング0・自動カーニングメトリクス・欧文ベースライン・左揃え・行送り自動・文字組みツメ組み・禁則弱い禁則v2）。", en: "Reset to defaults (12pt / 100% scale / 0 Tsume / 0 tracking / Metrics kerning / Roman baseline / left / auto leading / Tight mojikumi / Loose v2 kinsoku)." },
            reload: { ja: "選択中のテキストの現在値を読み取り直して UI に反映します。", en: "Re-read the current values from the selection and reflect them in the UI." },
            hiddenChar: { ja: "制御文字（改行・スペースなど）の表示／非表示を切り替えます。", en: "Toggle the display of hidden characters (returns, spaces, etc.)." },
            scriptType: { ja: "和文＝文字揃え中央／行送りの基準は仮想ボディの上、欧文＝文字揃え・行送りの基準ともに欧文ベースライン、混植＝文字揃え欧文ベースライン／行送りの基準は仮想ボディの上 をまとめて適用。", en: "Japanese = center alignment / top-to-top leading; Roman = Roman baseline for both alignment and leading; Mixed = Roman-baseline alignment / top-to-top leading." }
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

    /* ドキュメント内で使用しているフォントを収集 / Collect the fonts used in the document
       1フレーム内でもフォントは混在しうるため、代表フォントだけでなく全文字を走査して名前で重複排除する。
       返り値は "psName\tdisplay" を改行区切り。
       注意: 文字数に比例して DOM アクセスが増えるため、極端に長いテキストでは重くなる
       A single frame can mix fonts, so scan every character (not just the run's representative
       font) and dedup by name. Returns "psName<TAB>display" joined by newlines.
       Note: cost scales with character count (one DOM read per character), so very long text is slow */
    function collectDocumentFonts() {
        var doc = app.activeDocument;
        var seenFontNames = {}, lines = [];
        var TAB = String.fromCharCode(9), LF = String.fromCharCode(10);
        var frames = doc.textFrames;
        for (var i = 0; i < frames.length; i++) {
            try {
                var characters = frames[i].textRange.characters;
                for (var charIndex = 0; charIndex < characters.length; charIndex++) {
                    var font;
                    try { font = characters[charIndex].characterAttributes.textFont; } catch (eChar) { continue; }
                    if (font && !seenFontNames[font.name]) {
                        seenFontNames[font.name] = true;
                        var display = font.family + (font.style ? " " + font.style : "");
                        lines.push(font.name + TAB + display);
                    }
                }
            } catch (e) { }
        }
        lines.sort();
        return lines.join(LF);
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

    /* テキストフレーム全体の共通基準フォントサイズ / Common base font size across the whole frame */
    function getCommonBaseFontSize(textFrame) {
        var allSizes = [];
        var lines = textFrame.lines;
        for (var j = 0; j < lines.length; j++) {
            var sizes = getLineSampleFontSizes(lines[j], LINE_FONT_SIZE_SAMPLE_COUNT);
            for (var sizeIndex = 0; sizeIndex < sizes.length; sizeIndex++) allSizes.push(sizes[sizeIndex]);
        }
        if (allSizes.length === 0) return NaN;
        return getMostFrequentNumber(allSizes);
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

    /* テキストフレーム群に行送り・基準を適用 / Apply leading and basis to frames
       directMode: その他（行送り値を直接指定）/ ratio===null: 自動 / それ以外: 倍率
       Throws on the first failing frame so the caller can report it. */
    function applyLeadingToFrames(frames, ratio, leadingType, autoAmount, common, directMode, directLeadingPt) {
        var useDirect = directMode && !isNaN(directLeadingPt);
        var useAuto = (!useDirect) && (ratio === null);
        for (var i = 0; i < frames.length; i++) {
            var textFrame = frames[i];
            var lines = textFrame.lines;
            var commonBaseSize = (!useDirect && common && !useAuto) ? getCommonBaseFontSize(textFrame) : NaN;
            for (var j = 0; j < lines.length; j++) {
                var line = lines[j];
                if (useDirect) {
                    line.characterAttributes.autoLeading = false;
                    line.characterAttributes.leading = directLeadingPt;
                } else if (useAuto) {
                    line.characterAttributes.autoLeading = true;
                } else {
                    var baseFontSizeFromLine = common ? commonBaseSize : getBaseFontSizeFromLine(line);
                    if (isNaN(baseFontSizeFromLine)) continue;
                    line.characterAttributes.autoLeading = false;
                    line.characterAttributes.leading = baseFontSizeFromLine * ratio;
                }
            }
            var paragraphs = textFrame.paragraphs;
            for (var paragraphIndex = 0; paragraphIndex < paragraphs.length; paragraphIndex++) {
                try {
                    paragraphs[paragraphIndex].paragraphAttributes.autoLeadingAmount = autoAmount;
                } catch (e3) { }
            }
            try { textFrame.textRange.leadingType = leadingType; } catch (e4) { }
        }
    }

    /* 行送り（行単位）だけを適用。自動行送り量などの段落属性には触れない
       Apply only the per-line leading (and optionally the basis); leaves paragraph attributes untouched
       ratio===null: 自動 / それ以外: 倍率 */
    function applyLineLeadingToFrames(frames, ratio, leadingType) {
        var useAuto = (ratio === null);
        for (var i = 0; i < frames.length; i++) {
            var textFrame = frames[i];
            var lines = textFrame.lines;
            for (var j = 0; j < lines.length; j++) {
                var line = lines[j];
                if (useAuto) {
                    line.characterAttributes.autoLeading = true;
                } else {
                    var baseFontSizeFromLine = getBaseFontSizeFromLine(line);
                    if (isNaN(baseFontSizeFromLine)) continue;
                    line.characterAttributes.autoLeading = false;
                    line.characterAttributes.leading = baseFontSizeFromLine * ratio;
                }
            }
            if (leadingType !== null) { try { textFrame.textRange.leadingType = leadingType; } catch (e) { } }
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

    /* メインエンジンへ送る処理関数（上で定義済みのものを再利用）
       Functions shipped to the main engine (reuse of the helpers above) */
    var WORKER_FUNCS = [
        getTypeName, collectTextRangesFromItem, getSelectedTextRanges, resolveAutoKernType, applyKerningToRanges, applyTsumeToRanges,
        applyTrackingToRanges, clearAkiOnRanges, collectDocumentFonts, applyFontToRanges,
        kernMethodToId, justificationToId, resolveJustification,
        isTextRangeLikeType, findParentTextFrame, getTextFrameKey, collectJustifyTargets,
        collectJustifyTargetsFromItem, addJustifyTarget, applyJustificationToParagraphs,
        forceLeftJustification, justifyKeepingPosition,
        resolveAlignment, alignmentToId, getAlignmentCharAttrs, getFirstAlignmentCharAttrs, applyStyleRunAlignment,
        getLineSampleFontSizes, getMostFrequentNumber, getBaseFontSizeFromLine, getCommonBaseFontSize,
        getProcessableTextFrame, collectLeadingFrames, collectLeadingFramesFromItem, addLeadingFrame,
        resolveLeadingType, leadingTypeToId,
        applyLeadingToFrames, applyLineLeadingToFrames, applyAlignmentToSelection,
        applyMojikumiToTextFrame, applyMojikumiToItem, applyMojikumiToSelection, matchMojikumiName, mojikumiToId,
        applyKinsokuToTextFrame, applyKinsokuToItem, applyKinsokuToSelection, kinsokuToId,
        readLeadingState
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
        if (action === "getFonts") {
            return "OK:" + encodeURIComponent(collectDocumentFonts());
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
                return "ERR:" + (errCur && errCur.message ? errCur.message : String(errCur));
            }
        }
        if (action === "applyFont") {
            if (ranges.length === 0) return "OK:0";
            try {
                applyFontToRanges(ranges, params.psName);
            } catch (errFont) {
                return "ERR:" + (errFont && errFont.message ? errFont.message : String(errFont));
            }
            app.redraw();
            return "OK:" + ranges.length;
        }
        if (action === "applyFontPreset") {
            if (ranges.length === 0) return "OK:0";
            try {
                applyFontToRanges(ranges, params.psName);
                if (params.kern) applyKerningToRanges(ranges, resolveAutoKernType(params.kern));
                if (params.tsume !== undefined && params.tsume !== null) applyTsumeToRanges(ranges, params.tsume);
                if (params.tracking !== undefined && params.tracking !== null) applyTrackingToRanges(ranges, params.tracking);
            } catch (errPreset) {
                return "ERR:" + (errPreset && errPreset.message ? errPreset.message : String(errPreset));
            }
            app.redraw();
            return "OK:" + ranges.length;
        }
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
        if (action === "apply") {
            if (ranges.length === 0) return "OK:0";
            try {
                applyKerningToRanges(ranges, resolveAutoKernType(params.method));
            } catch (errApply) {
                return "ERR:" + (errApply && errApply.message ? errApply.message : String(errApply));
            }
            app.redraw();
            return "OK:" + ranges.length;
        }
        if (action === "applyTsume") {
            if (ranges.length === 0) return "OK:0";
            try {
                applyTsumeToRanges(ranges, params.value);
            } catch (errTsume) {
                return "ERR:" + (errTsume && errTsume.message ? errTsume.message : String(errTsume));
            }
            app.redraw();
            return "OK:" + ranges.length;
        }
        if (action === "applyTracking") {
            if (ranges.length === 0) return "OK:0";
            try {
                applyTrackingToRanges(ranges, params.tracking);
            } catch (errTrack) {
                return "ERR:" + (errTrack && errTrack.message ? errTrack.message : String(errTrack));
            }
            app.redraw();
            return "OK:" + ranges.length;
        }
        if (action === "applyJustification") {
            var justifyTargets = collectJustifyTargets(app.activeDocument.selection);
            if (justifyTargets.length === 0) return "OK:0";
            var justifyValue = resolveJustification(params.justify);
            try {
                for (var targetIndex = 0; targetIndex < justifyTargets.length; targetIndex++) {
                    justifyKeepingPosition(justifyTargets[targetIndex], justifyValue, params.justify);
                }
            } catch (errJustify) {
                return "ERR:" + (errJustify && errJustify.message ? errJustify.message : String(errJustify));
            }
            app.redraw();
            return "OK:" + justifyTargets.length;
        }
        if (action === "applyFontSize") {
            if (ranges.length === 0) return "OK:0";
            try {
                for (var rangeIndex = 0; rangeIndex < ranges.length; rangeIndex++) ranges[rangeIndex].characterAttributes.size = params.sizePt;
            } catch (errSize) {
                return "ERR:" + (errSize && errSize.message ? errSize.message : String(errSize));
            }
            app.redraw();
            return "OK:" + ranges.length;
        }
        if (action === "applyScale") {
            if (ranges.length === 0) return "OK:0";
            try {
                for (var rangeIndex = 0; rangeIndex < ranges.length; rangeIndex++) {
                    ranges[rangeIndex].characterAttributes.horizontalScale = params.scale;
                    ranges[rangeIndex].characterAttributes.verticalScale = params.scale;
                }
            } catch (errScale) {
                return "ERR:" + (errScale && errScale.message ? errScale.message : String(errScale));
            }
            app.redraw();
            return "OK:" + ranges.length;
        }
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
                return "ERR:" + (errAlign && errAlign.message ? errAlign.message : String(errAlign));
            }
            app.redraw();
            return "OK:" + alignCount;
        }
        if (action === "applyLeading") {
            var leadingFrames = collectLeadingFrames(app.activeDocument.selection);
            if (leadingFrames.length === 0) return "OK:" + readLeadingState(leadingFrames);
            try {
                applyLeadingToFrames(
                    leadingFrames, params.ratio, resolveLeadingType(params.leadingType),
                    params.autoAmount, params.common, params.directMode, params.directLeadingPt
                );
            } catch (errLeading) {
                return "ERR:" + (errLeading && errLeading.message ? errLeading.message : String(errLeading));
            }
            app.redraw();
            return "OK:" + readLeadingState(leadingFrames);
        }
        if (action === "applyProfile") {
            var profSel = app.activeDocument.selection;
            try {
                if (params.kern !== undefined) applyKerningToRanges(ranges, resolveAutoKernType(params.kern));
                if (params.tsume !== undefined) applyTsumeToRanges(ranges, params.tsume);
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
                if (params.ratio !== undefined) {
                    var profFrames = collectLeadingFrames(profSel);
                    var profType = (params.leadingType !== undefined) ? resolveLeadingType(params.leadingType) : null;
                    applyLineLeadingToFrames(profFrames, params.ratio, profType);
                } else if (params.leadingType !== undefined) {
                    /* 行送り値は変えず、行送りの基準だけを適用 / Set only the leading basis, leaving the leading value */
                    var basisFrames = collectLeadingFrames(profSel);
                    var basisType = resolveLeadingType(params.leadingType);
                    for (var basisFrameIndex = 0; basisFrameIndex < basisFrames.length; basisFrameIndex++) {
                        try { basisFrames[basisFrameIndex].textRange.leadingType = basisType; } catch (eBasis) { }
                    }
                }
                if (params.mojikumiIndex !== undefined) applyMojikumiToSelection(profSel, params.mojikumiIndex);
                if (params.kinsoku !== undefined) applyKinsokuToSelection(profSel, params.kinsoku);
                if (params.clearAki) clearAkiOnRanges(ranges);
            } catch (errProfile) {
                return "ERR:" + (errProfile && errProfile.message ? errProfile.message : String(errProfile));
            }
            app.redraw();
            return "OK:" + ranges.length;
        }
        if (action === "applyMojikumi") {
            try {
                applyMojikumiToSelection(app.activeDocument.selection, params.mojikumiIndex);
            } catch (errMojikumi) {
                return "ERR:" + (errMojikumi && errMojikumi.message ? errMojikumi.message : String(errMojikumi));
            }
            app.redraw();
            return "OK:1";
        }
        if (action === "applyKinsoku") {
            try {
                applyKinsokuToSelection(app.activeDocument.selection, params.kinsoku);
            } catch (errKinsoku) {
                return "ERR:" + (errKinsoku && errKinsoku.message ? errKinsoku.message : String(errKinsoku));
            }
            app.redraw();
            return "OK:1";
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
        if (params.ratio !== undefined) parts.push("ratio:" + (params.ratio === null ? "null" : params.ratio));
        if (params.leadingType !== undefined) parts.push('leadingType:"' + params.leadingType + '"');
        if (params.mojikumiIndex !== undefined) parts.push("mojikumiIndex:" + parseInt(params.mojikumiIndex, 10));
        if (params.kinsoku !== undefined) parts.push('kinsoku:decodeURIComponent("' + encodeURIComponent(params.kinsoku) + '")');
        if (params.autoAmount !== undefined) parts.push("autoAmount:" + params.autoAmount);
        if (params.common !== undefined) parts.push("common:" + (params.common ? "true" : "false"));
        if (params.directMode !== undefined) parts.push("directMode:" + (params.directMode ? "true" : "false"));
        if (params.directLeadingPt !== undefined) parts.push("directLeadingPt:" + (isNaN(params.directLeadingPt) ? "NaN" : params.directLeadingPt));
        if (params.clearAki !== undefined) parts.push("clearAki:" + (params.clearAki ? "true" : "false"));
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
    // UI構築 / Build UI
    // =========================================

    // =========================================
    // 行送り用の単位・選択肢ヘルパー / Leading: units and choices (palette side)
    // =========================================

    /* パネルの余白と間隔 / Panel margins and spacing */
    var PANEL_MARGINS = [13, 20, 13, 12];
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

    var UNIT_MAP = { 0: "in", 1: "mm", 2: "pt", 3: "pica", 4: "cm", 6: "px", 7: "ft/in", 8: "m", 9: "yd", 10: "ft" };
    var UNIT_TO_PT = {
        0: 72, 1: 2.8346456692913386, 2: 1, 3: 12, 4: 28.346456692913386,
        5: 0.7086614173228346, 6: 1, 7: 72, 8: 2834.6456692913386, 9: 2592, 10: 864
    };

    function getUnitLabel(unitCode) {
        if (unitCode === 5) return "Q";
        return UNIT_MAP[unitCode] || "pt";
    }

    /* テキスト単位を取得（コード・ラベル・pt 換算係数）/ Resolve the text unit (code, label, pt factor) */
    function getTextUnit() {
        var unitCode = 2;
        try { unitCode = app.preferences.getIntegerPreference("text/units"); } catch (e) { }
        return { code: unitCode, label: getUnitLabel(unitCode), factor: UNIT_TO_PT[unitCode] || 1 };
    }

    /* 行送りの選択肢 / Leading choices
       ratio: 倍率 / undefined=その他（直接入力）/ null=自動 */
    var LEADING_CHOICES = [
        { label: "110%", ratio: 1.1, auto: false, other: false },
        { label: "125%", ratio: 1.25, auto: false, other: false },
        { label: "150%", ratio: 1.5, auto: false, other: false },
        { label: getLocalizedText(LABELS.leading.other), ratio: undefined, auto: false, other: true },
        { label: getLocalizedText(LABELS.leading.auto), ratio: null, auto: true, other: false }
    ];

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

    /* 自動行送りラベルを組み立て（例: 自動（175%）/ Auto (175%)）/ Build the auto-leading label */
    function formatAutoLabel(amount) {
        var baseLabel = getLocalizedText(LABELS.leading.auto);
        var roundedAmount = isNaN(amount) ? 175 : Math.round(amount);
        return currentLanguage === "ja" ? (baseLabel + "（" + roundedAmount + "%）") : (baseLabel + " (" + roundedAmount + "%)");
    }

    /* 現在の行送り値が定型倍率／自動／その他のどれに一致するか / Match the current leading to a choice index */
    function findMatchingLeadingChoiceIndex(isAutoLeading, leadingValuePt, baseFontSizePt) {
        if (isAutoLeading) {
            for (var i = 0; i < LEADING_CHOICES.length; i++) {
                if (LEADING_CHOICES[i].auto) return i;
            }
        }
        if (!isNaN(leadingValuePt) && !isNaN(baseFontSizePt) && baseFontSizePt > 0) {
            var ratio = leadingValuePt / baseFontSizePt;
            for (var j = 0; j < LEADING_CHOICES.length; j++) {
                var choice = LEADING_CHOICES[j];
                if (!choice.auto && !choice.other && typeof choice.ratio === "number") {
                    if (Math.abs(choice.ratio - ratio) < 0.0001) return j;
                }
            }
        }
        for (var k = 0; k < LEADING_CHOICES.length; k++) {
            if (LEADING_CHOICES[k].other) return k;
        }
        return 0;
    }

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
    function fillPresetList(listbox, presets) {
        listbox.removeAll();
        for (var i = 0; i < presets.length; i++) {
            var preset = presets[i];
            var listItem = listbox.add("item", preset.psName);
            listItem.psName = preset.psName;
            if (preset.kern !== undefined) listItem.kern = preset.kern;
            if (preset.tsume !== undefined) listItem.tsume = preset.tsume;
            if (preset.tracking !== undefined) listItem.tracking = preset.tracking;
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

    /* パレットを組み立てて参照を返す（イベント未接続）/ Build the palette and return references (events not wired yet) */
    function createPaletteUI(autoKernOptions, alignOptions, justifyOptions) {
        var palette = new Window("palette", getLocalizedText(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        palette.alignChildren = "fill";

        var textUnit = getTextUnit();
        var leadingBasisChoices = getLeadingTypeChoices();

        // ===== 上部インフォバー（カラムを貫通、テキスト1つ選択時のみ表示、1行テキスト）
        //       Top info bar spanning the columns as a single line (shown only when a single text is selected) =====
        var infoGroup = palette.add("group");
        infoGroup.orientation = "column"; // 2行を縦に積む（マルチラインの描画クセを避ける）/ Stack two lines vertically (avoids multiline quirks)
        infoGroup.alignment = "fill";
        infoGroup.alignChildren = ["fill", "center"];
        infoGroup.spacing = 2;
        infoGroup.margins = [4, 4, 4, 4];
        // 1行目：フォント・サイズ・行送り / Line 1: font, size, leading
        var infoText = infoGroup.add("statictext", undefined, "");
        infoText.justify = "center";
        // 2行目：その他（行揃え・カーニング・禁則）/ Line 2: the rest (justify, kerning, kinsoku)
        var infoTextSub = infoGroup.add("statictext", undefined, "");
        infoTextSub.justify = "center";

        // ===== 種別：タブから独立させ、インフォバーの直下に配置。左パネル＝本文／見出し、右パネル＝和文／欧文／混植（横並び）
        //       Type roles: independent of the tabs, just below the info bar. Left panel = Body/Heading, right panel = Japanese/Roman/Mixed (side by side) =====
        var roleRow = palette.add("group");
        roleRow.orientation = "row";
        roleRow.alignment = ["fill", "top"];
        roleRow.alignChildren = ["fill", "top"];
        roleRow.spacing = 6; // パネル間ギャップを詰める / Tighten the inter-panel gap

        // 種別パネル用の詰めたマージン（左右13→8・上下を薄く）でパレット幅の膨張を抑える
        // Tighter margins for the role panels (L/R 13→8, thinner top/bottom) to curb palette-width growth
        var ROLE_PANEL_MARGINS = [8, 16, 8, 8];
        var ROLE_RADIO_SPACING = 10; // ラジオ間隔を詰める / Tighten the radio spacing

        // 左パネル：種別（本文／見出し・ラジオ横並び）/ Left panel: Type (Body / Heading, radios in a row)
        var rolePanel = roleRow.add("panel", undefined, getLocalizedText(LABELS.field.role));
        setupPanel(rolePanel, 6);
        rolePanel.margins = ROLE_PANEL_MARGINS;
        rolePanel.alignment = ["fill", "top"];
        rolePanel.alignChildren = ["center", "top"];
        rolePanel.helpTip = getLocalizedText(LABELS.tip.role);
        var roleRoleRow = rolePanel.add("group");
        roleRoleRow.orientation = "row";
        roleRoleRow.spacing = ROLE_RADIO_SPACING;
        roleRoleRow.alignment = ["center", "top"]; // 左右中央 / Horizontally centered
        var roleBodyRadio = roleRoleRow.add("radiobutton", undefined, getLocalizedText(LABELS.role.body));
        var roleHeadingRadio = roleRoleRow.add("radiobutton", undefined, getLocalizedText(LABELS.role.heading));
        roleBodyRadio.helpTip = rolePanel.helpTip;
        roleHeadingRadio.helpTip = rolePanel.helpTip;

        // 右パネル：文字種（和文／欧文／混植・ラジオ横並び。文字揃え・行送りの基準をまとめて適用）/ Right panel: Script (Japanese / Roman / Mixed, radios in a row; applies alignment + leading basis together)
        var scriptTypePanel = roleRow.add("panel", undefined, getLocalizedText(LABELS.field.scriptType));
        setupPanel(scriptTypePanel, 6);
        scriptTypePanel.margins = ROLE_PANEL_MARGINS;
        scriptTypePanel.alignment = ["fill", "top"];
        scriptTypePanel.alignChildren = ["center", "top"];
        var scriptTypeTip = getLocalizedText(LABELS.tip.scriptType);
        scriptTypePanel.helpTip = scriptTypeTip;
        var scriptTypeRadioRow = scriptTypePanel.add("group");
        scriptTypeRadioRow.orientation = "row";
        scriptTypeRadioRow.spacing = ROLE_RADIO_SPACING;
        scriptTypeRadioRow.alignment = ["center", "top"]; // 左右中央 / Horizontally centered
        var roleJapaneseRadio = scriptTypeRadioRow.add("radiobutton", undefined, getLocalizedText(LABELS.role.japanese));
        var roleRomanRadio = scriptTypeRadioRow.add("radiobutton", undefined, getLocalizedText(LABELS.role.roman));
        var roleMixedRadio = scriptTypeRadioRow.add("radiobutton", undefined, getLocalizedText(LABELS.role.mixed));
        roleJapaneseRadio.helpTip = scriptTypeTip;
        roleRomanRadio.helpTip = scriptTypeTip;
        roleMixedRadio.helpTip = scriptTypeTip;

        // ===== レイアウト：タブ1（ドキュメントフォント）／タブ2（プリセット）／タブ3（文字組み：カーニング・ツメ・文字揃え・行送りなど）／タブ4（日本語組版：行揃え・禁則・文字組みアキ量設定）
        //       Layout: tab 1 (document fonts) / tab 2 (presets) / tab 3 (type: kerning, tsume, alignment, leading, ...) / tab 4 (Japanese composition: justification, kinsoku, mojikumi) =====
        var mainTabs = palette.add("tabbedpanel");
        mainTabs.alignChildren = ["fill", "top"];

        // ---- タブ1：ドキュメントフォント / Tab 1: document fonts ----
        var docFontsTab = mainTabs.add("tab", undefined, getLocalizedText(LABELS.tab.docFonts));
        docFontsTab.orientation = "column";
        docFontsTab.alignChildren = ["fill", "top"];
        docFontsTab.margins.top = 15; // タブ内上部にマージン / Top margin inside the tab
        docFontsTab.margins.left = 15; // タブ内左にマージン / Left margin inside the tab

        // ---- タブ2：プリセット / Tab 2: presets ----
        var presetsTab = mainTabs.add("tab", undefined, getLocalizedText(LABELS.tab.presets));
        presetsTab.orientation = "column";
        presetsTab.alignChildren = ["fill", "top"];
        presetsTab.margins.top = 15; // タブ内上部にマージン / Top margin inside the tab
        presetsTab.margins.left = 15; // タブ内左にマージン / Left margin inside the tab

        // ---- タブ3：文字組み設定（中央＋右カラムを横並び）/ Tab 3: type settings (center + right columns side by side) ----
        var typeTab = mainTabs.add("tab", undefined, getLocalizedText(LABELS.tab.type));
        typeTab.orientation = "column";
        typeTab.alignChildren = ["fill", "top"];
        typeTab.margins.top = 15; // タブ内上部にマージン / Top margin inside the tab
        typeTab.margins.left = 15; // タブ内左にマージン / Left margin inside the tab

        // ---- タブ4：日本語組版（行揃え・禁則・文字組みアキ量設定）/ Tab 4: Japanese composition (justification, kinsoku, mojikumi) ----
        var jpCompositionTab = mainTabs.add("tab", undefined, getLocalizedText(LABELS.tab.jpComposition));
        jpCompositionTab.orientation = "column";
        jpCompositionTab.alignChildren = ["fill", "top"];
        jpCompositionTab.margins.top = 15; // タブ内上部にマージン / Top margin inside the tab
        jpCompositionTab.margins.left = 15; // タブ内左にマージン / Left margin inside the tab
        jpCompositionTab.spacing = 8;

        var typeColumnsRow = typeTab.add("group");
        typeColumnsRow.orientation = "row";
        typeColumnsRow.alignChildren = ["fill", "top"];
        typeColumnsRow.spacing = 12;

        mainTabs.selection = docFontsTab;

        // ---- タブ1の中身：ドキュメントフォント / Tab 1 content: document fonts ----
        var fontPanel = docFontsTab.add("panel", undefined, getLocalizedText(LABELS.field.docFonts));
        setupPanel(fontPanel);

        var fontList = fontPanel.add("listbox", undefined, [], { multiselect: false });
        fontList.preferredSize = [164, 200]; // 1行分低く / One row shorter
        fontList.helpTip = getLocalizedText(LABELS.tip.docFonts);

        // ---- タブ2の中身：プリセット / Tab 2 content: presets ----
        var presetPanel = presetsTab.add("panel", undefined, getLocalizedText(LABELS.field.presets));
        setupPanel(presetPanel);

        var presetList = presetPanel.add("listbox", undefined, [], { multiselect: false });
        presetList.preferredSize = [164, 200];
        presetList.helpTip = getLocalizedText(LABELS.tip.presets);
        fillPresetList(presetList, loadPresets());

        // プリセット操作ボタンを3カラムに（フッターと同じロジック：左＝削除／中央＝スペーサー／右＝上書き・追加）
        // Preset action buttons in 3 columns (same logic as the footer: left = delete / center = fill spacer / right = overwrite, add)
        var presetButtonRow = presetPanel.add("group");
        presetButtonRow.orientation = "row";
        presetButtonRow.alignment = ["fill", "top"]; // 幅いっぱいにして左右に振り分ける / Stretch to full width to split left/right
        presetButtonRow.margins = [0, 5, 0, 0]; // ボタン行の上に余白 / Extra top margin above the button row

        // 左：削除 / Left: delete
        var deletePresetButton = presetButtonRow.add("button", undefined, getLocalizedText(LABELS.button.deletePreset));
        deletePresetButton.alignment = ["left", "center"];
        deletePresetButton.helpTip = getLocalizedText(LABELS.tip.deletePreset);

        // 中央：スペーサー（余白を吸って左右に振り分ける）/ Center: spacer (absorbs slack to split left/right)
        var presetButtonSpacer = presetButtonRow.add("statictext", undefined, "");
        presetButtonSpacer.alignment = ["fill", "center"];

        // 右：上書き・追加 / Right: overwrite, add
        var overwritePresetButton = presetButtonRow.add("button", undefined, getLocalizedText(LABELS.button.overwritePreset));
        overwritePresetButton.alignment = ["right", "center"];
        overwritePresetButton.helpTip = getLocalizedText(LABELS.tip.overwritePreset);

        var addPresetButton = presetButtonRow.add("button", undefined, getLocalizedText(LABELS.button.addPreset));
        addPresetButton.alignment = ["right", "center"];
        addPresetButton.helpTip = getLocalizedText(LABELS.tip.addPreset);

        // ---- 中央カラム：フォントサイズ・自動カーニング・文字ツメ / Center column: font size, kerning, Tsume ----
        var characterColumn = typeColumnsRow.add("group");
        characterColumn.orientation = "column";
        characterColumn.alignChildren = ["fill", "top"];
        characterColumn.spacing = 8;

        // フォントサイズ（文字サイズ・比率・見かけ＋焼き込みトグル）/ Font size (size, scale, apparent + bake toggle)
        var fontSizePanel = characterColumn.add("panel", undefined, getLocalizedText(LABELS.field.fontSizePanel));
        setupPanel(fontSizePanel, 6);
        fontSizePanel.alignChildren = ["left", "top"];

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

        // ---- 右カラム：文字揃え・行送り・行送りの基準 / Right column: char alignment, leading, basis ----
        var paragraphColumn = typeColumnsRow.add("group");
        paragraphColumn.orientation = "column";
        paragraphColumn.alignChildren = ["fill", "top"];
        paragraphColumn.spacing = 8;

        var alignPanel = paragraphColumn.add("panel", undefined, getLocalizedText(LABELS.field.align));
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

        // 行揃え：日本語組版タブへ配置 / Justification: placed on the Japanese-composition tab
        var justifyPanel = jpCompositionTab.add("panel", undefined, getLocalizedText(LABELS.field.justify));
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

        var leadingPanel = paragraphColumn.add("panel", undefined, getLocalizedText(LABELS.field.leading));
        setupPanel(leadingPanel);
        leadingPanel.alignChildren = "left";
        leadingPanel.helpTip = getLocalizedText(LABELS.tip.leading);

        // 個別／共通はタイトルなしのパネルに入れる / Individual / Common in a titleless panel
        var modePanel = leadingPanel.add("panel", undefined, "");
        modePanel.orientation = "row";
        modePanel.alignChildren = ["center", "center"];
        modePanel.alignment = ["fill", "middle"];
        modePanel.margins = [10, 6, 10, 4];
        modePanel.spacing = 12;
        var modePerLine = modePanel.add("radiobutton", undefined, getLocalizedText(LABELS.leading.perLine));
        var modeCommon = modePanel.add("radiobutton", undefined, getLocalizedText(LABELS.leading.common));
        modePerLine.helpTip = getLocalizedText(LABELS.leading.tipPerLine);
        modeCommon.helpTip = getLocalizedText(LABELS.leading.tipCommon);
        modePerLine.value = true;

        // ラジオを縦一列に並べる（「その他」「自動」は同じ行にインラインで入力欄）
        // Stack radios in one column; "Other"/"Auto" carry an inline input on the same row
        var leadingRadioGroup = leadingPanel.add("group");
        leadingRadioGroup.orientation = "column";
        leadingRadioGroup.alignChildren = "left";
        leadingRadioGroup.spacing = 6;
        leadingRadioGroup.margins = [0, 5, 0, 0];

        var leadingRadios = [];
        var leadingInput = null;
        var autoRadio = null;
        for (var choiceIndex = 0; choiceIndex < LEADING_CHOICES.length; choiceIndex++) {
            var choice = LEADING_CHOICES[choiceIndex];
            var leadingRow = leadingRadioGroup.add("group");
            leadingRow.orientation = "row";
            leadingRow.alignChildren = "center";
            leadingRow.spacing = 6;
            var leadingRadio = leadingRow.add("radiobutton", undefined, choice.label);
            leadingRadio.index = choiceIndex;
            leadingRadios.push(leadingRadio);
            if (choice.other) {
                // その他：倍率を % で直接指定 / Other: a custom ratio entered as a percentage
                leadingInput = leadingRow.add("edittext", undefined, "");
                leadingInput.characters = 4;
                leadingRow.add("statictext", undefined, "%");
            }
            if (choice.auto) {
                // 自動行送り量はラベルにテキストで表示（編集不可）/ Auto-leading amount shown as plain label text (not editable)
                autoRadio = leadingRadio;
                leadingRadio.text = formatAutoLabel(175);
            }
        }
        leadingRadios[0].value = true;

        // 自動行送りの値を変更するボタン（「自動」の下）/ Button to change the auto-leading amount (below "Auto")
        // ラッパー group で下マージンを付ける / Wrapper group adds a bottom margin
        var changeAutoGroup = leadingPanel.add("group");
        changeAutoGroup.orientation = "row";
        changeAutoGroup.alignment = ["fill", "top"];
        changeAutoGroup.alignChildren = ["center", "top"];
        changeAutoGroup.margins = [0, 0, 0, 3]; // [左, 上, 右, 下]
        var changeAutoButton = changeAutoGroup.add("button", undefined, getLocalizedText(LABELS.button.changeAuto));
        changeAutoButton.preferredSize.height = 22; // プリセットボタンと高さを揃えて小さく / Match the preset buttons' height
        changeAutoButton.helpTip = getLocalizedText(LABELS.tip.changeAuto);

        // 行送りの基準：タイトルなしパネルとして行送りパネル内に配置（個別／共通と同じ体裁）
        // Leading basis: a titleless panel inside the leading panel (same style as Individual / Common)
        var leadingBasisPanel = leadingPanel.add("panel", undefined, "");
        leadingBasisPanel.orientation = "column";
        leadingBasisPanel.alignChildren = "left";
        leadingBasisPanel.alignment = ["fill", "top"];
        leadingBasisPanel.margins = [10, 8, 10, 8];
        leadingBasisPanel.spacing = 6;
        leadingBasisPanel.helpTip = getLocalizedText(LABELS.tip.leadingType);
        var leadingBasisRadios = [];
        for (var basisIndex = 0; basisIndex < leadingBasisChoices.length; basisIndex++) {
            var leadingBasisRadio = leadingBasisPanel.add("radiobutton", undefined, leadingBasisChoices[basisIndex].label);
            leadingBasisRadio.index = basisIndex;
            leadingBasisRadios.push(leadingBasisRadio);
        }
        leadingBasisRadios[0].value = true;

        // 禁則・文字組みアキ量設定を横並びに / Kinsoku and Mojikumi side by side
        var jpCompositionRow = jpCompositionTab.add("group");
        jpCompositionRow.orientation = "row";
        jpCompositionRow.alignChildren = ["fill", "top"];
        jpCompositionRow.spacing = 12;

        // 禁則（ラジオボタン）/ Kinsoku (radio buttons)
        var kinsokuPanel = jpCompositionRow.add("panel", undefined, getLocalizedText(LABELS.field.kinsoku));
        setupPanel(kinsokuPanel, 6);
        kinsokuPanel.alignChildren = ["left", "top"];
        kinsokuPanel.helpTip = getLocalizedText(LABELS.tip.kinsoku);

        var kinsokuRadios = [];
        for (var kinsokuChoiceIndex = 0; kinsokuChoiceIndex < KINSOKU_CHOICES.length; kinsokuChoiceIndex++) {
            var kinsokuRadio = kinsokuPanel.add("radiobutton", undefined, KINSOKU_CHOICES[kinsokuChoiceIndex].label);
            kinsokuRadio.index = kinsokuChoiceIndex;
            kinsokuRadio.helpTip = kinsokuPanel.helpTip;
            kinsokuRadios.push(kinsokuRadio);
        }
        kinsokuRadios[0].value = true;

        // 文字組みアキ量設定（ラジオボタン）/ Mojikumi spacing set (radio buttons)
        var mojikumiPanel = jpCompositionRow.add("panel", undefined, getLocalizedText(LABELS.field.mojikumi));
        setupPanel(mojikumiPanel, 6);
        mojikumiPanel.alignChildren = ["left", "top"];
        mojikumiPanel.helpTip = getLocalizedText(LABELS.tip.mojikumi);

        var mojikumiRadios = [];
        for (var mojikumiChoiceIndex = 0; mojikumiChoiceIndex < MOJIKUMI_CHOICES.length; mojikumiChoiceIndex++) {
            var mojikumiRadio = mojikumiPanel.add("radiobutton", undefined, MOJIKUMI_CHOICES[mojikumiChoiceIndex].label);
            mojikumiRadio.index = mojikumiChoiceIndex;
            mojikumiRadio.helpTip = mojikumiPanel.helpTip;
            mojikumiRadios.push(mojikumiRadio);
        }
        mojikumiRadios[0].value = true;

        // フッター（左=制御文字／中央=スペーサー／右=再読み込み・リセット）/ Footer: hidden-chars, spacer, reload, reset
        var footerGroup = palette.add("group");
        footerGroup.orientation = "row";
        footerGroup.alignment = "fill";
        var hiddenCharButton = footerGroup.add("button", undefined, getLocalizedText(LABELS.button.hiddenChar));
        hiddenCharButton.alignment = ["left", "center"];
        hiddenCharButton.helpTip = getLocalizedText(LABELS.tip.hiddenChar);
        var footerSpacer = footerGroup.add("statictext", undefined, "");
        footerSpacer.alignment = ["fill", "center"];
        var reloadButton = footerGroup.add("button", undefined, getLocalizedText(LABELS.button.reload));
        reloadButton.alignment = ["right", "center"];
        reloadButton.helpTip = getLocalizedText(LABELS.tip.reload);
        var resetButton = footerGroup.add("button", undefined, getLocalizedText(LABELS.button.reset));
        resetButton.alignment = ["right", "center"];
        resetButton.helpTip = getLocalizedText(LABELS.tip.reset);

        return {
            palette: palette,
            infoText: infoText,
            infoTextSub: infoTextSub,
            fontList: fontList,
            presetList: presetList,
            addPresetButton: addPresetButton,
            overwritePresetButton: overwritePresetButton,
            deletePresetButton: deletePresetButton,
            fontSizeInput: fontSizeInput,
            scaleInput: scaleInput,
            fsScaleLabel: fsScaleLabel,
            apparentValueLabel: apparentValueLabel,
            fsApparentLabel: fsApparentLabel,
            fsApparentUnit: fsApparentUnit,
            roleBodyRadio: roleBodyRadio,
            roleHeadingRadio: roleHeadingRadio,
            roleJapaneseRadio: roleJapaneseRadio,
            roleRomanRadio: roleRomanRadio,
            roleMixedRadio: roleMixedRadio,
            alignRadios: alignRadios,
            alignRomanRadio: alignRomanRadio,
            alignCenterRadio: alignCenterRadio,
            alignOtherRadio: alignOtherRadio,
            alignOtherDropdown: alignOtherDropdown,
            alignOtherOptions: alignOtherOptions,
            justifyButtons: justifyButtons,
            justifyState: justifyState,
            alignOptions: alignOptions,
            justifyOptions: justifyOptions,
            kernRadios: kernRadios,
            tsumeSlider: tsumeSlider,
            tsumeInput: tsumeInput,
            trackingSlider: trackingSlider,
            trackingInput: trackingInput,
            resetButton: resetButton,
            reloadButton: reloadButton,
            hiddenCharButton: hiddenCharButton,
            textUnit: textUnit,
            leadingBasisChoices: leadingBasisChoices,
            leadingRadios: leadingRadios,
            leadingInput: leadingInput,
            autoRadio: autoRadio,
            changeAutoButton: changeAutoButton,
            modePerLine: modePerLine,
            modeCommon: modeCommon,
            leadingBasisRadios: leadingBasisRadios,
            mojikumiRadios: mojikumiRadios,
            mojikumiChoices: MOJIKUMI_CHOICES,
            kinsokuRadios: kinsokuRadios,
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

    /* カーニング method id を表示ラベルへ / Kerning method id to a display label */
    function kernLabel(kernId) {
        if (kernId === "mono") return getLocalizedText(LABELS.autoKern.mono);
        if (kernId === "zero") return getLocalizedText(LABELS.autoKern.zero);
        if (kernId === "metrics") return getLocalizedText(LABELS.autoKern.metrics);
        if (kernId === "optical") return getLocalizedText(LABELS.autoKern.optical);
        return "—";
    }

    /* 行揃え id を表示ラベルへ / Justification id to a display label */
    function justifyLabel(justifyId) {
        return (justifyId && LABELS.justify[justifyId]) ? getLocalizedText(LABELS.justify[justifyId]) : "—";
    }

    /* 禁則 id を表示ラベルへ / Kinsoku id to a display label */
    function kinsokuLabel(kinsokuId) {
        for (var i = 0; i < KINSOKU_CHOICES.length; i++) {
            if (KINSOKU_CHOICES[i].id === kinsokuId) return KINSOKU_CHOICES[i].label;
        }
        return "—";
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
        var lastAutoAmount = 175; // 直近に読み取った自動行送り量（適用時はこの値を維持）/ Last-read auto-leading amount (kept on apply)
        var lastLeadingPercent = NaN; // 直近に読み取った行送りの割合（%）／「その他」初期値に使う / Last-read leading percentage (used as the default for "Other")
        var suppressUiEvents = false; // 反映中のラジオ onClick を抑止 / Suppress radio onClick while reflecting state

        /* 重要処理の失敗をインフォバーに表示（握りつぶさず、どのアクションで何が起きたか見えるように）
           Surface an important-op failure in the info bar (don't swallow it; show which action failed and why) */
        function showWorkerError(actionId, payload) {
            var detail = payload ? (": " + String(payload)) : "";
            try { ui.infoText.text = "⚠ " + getLocalizedText(LABELS.msg.applyError) + " [" + actionId + "]" + detail; ui.infoTextSub.text = ""; } catch (e) { }
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

        // ---- ドキュメントフォント / Document fonts（フォントのみ適用）----
        ui.fontList.onChange = function () {
            if (this.selection && this.selection.psName) {
                runApply("applyFont", { psName: this.selection.psName });
            }
        };

        // ---- プリセット / Presets（フォント＋カーニング・ツメ・トラッキングをまとめて適用）----
        ui.presetList.onChange = function () {
            var selectedPreset = this.selection;
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
        };

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
            fillPresetList(ui.presetList, presets);
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
            var index = ui.presetList.selection ? ui.presetList.selection.index : -1;
            if (index < 0) return; // 未選択なら何もしない / Do nothing if nothing is selected
            buildPresetFromSelection(function (newPreset) {
                if (!newPreset) return;
                var presets = loadPresets();
                if (index >= presets.length) return;
                presets[index] = newPreset;
                commitPresets(presets);
            });
        };

        // ---- プリセットを削除（リストで選択中の項目を削除）/ Delete the selected list item ----
        ui.deletePresetButton.onClick = function () {
            var index = ui.presetList.selection ? ui.presetList.selection.index : -1;
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
        // 行送りラジオを倍率（null=自動）で選択 / Select the leading radio by ratio (null = auto)
        function reflectLeadingRatio(ratio) {
            for (var i = 0; i < LEADING_CHOICES.length; i++) {
                var choice = LEADING_CHOICES[i];
                if (ratio === null) {
                    if (choice.auto) { selectExclusiveRadio(ui.leadingRadios, i); return; }
                } else if (!choice.auto && !choice.other && choice.ratio === ratio) {
                    selectExclusiveRadio(ui.leadingRadios, i); return;
                }
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
        // 文字組みアキ量設定ラジオを mojikumiSet インデックスで選択 / Select the mojikumi radio by mojikumiSet index
        function reflectMojikumiByIndex(mojikumiIndex) {
            for (var i = 0; i < ui.mojikumiChoices.length; i++) {
                if (ui.mojikumiChoices[i].index === mojikumiIndex) { selectExclusiveRadio(ui.mojikumiRadios, i); return; }
            }
        }
        // 禁則ラジオを id（"None"/"Hard"/"Soft"/"Soft_v2"）で選択 / Select the kinsoku radio by id
        function reflectKinsokuById(kinsokuId) {
            for (var i = 0; i < ui.kinsokuChoices.length; i++) {
                if (ui.kinsokuChoices[i].id === kinsokuId) {
                    suppressUiEvents = true;
                    selectExclusiveRadio(ui.kinsokuRadios, i);
                    suppressUiEvents = false;
                    return;
                }
            }
        }

        // 本文：和文等幅／行送り 150%／均等配置（最終行左）／文字組みベタ組み / Body preset (mojikumi: Solid)
        ui.roleBodyRadio.onClick = function () {
            runApply("applyProfile", { kern: "mono", ratio: 1.5, justify: "justifyLeft", mojikumiIndex: 6, kinsoku: "Soft_v2" });
            selectRadioById(ui.kernRadios, autoKernOptions, "mono");
            reflectLeadingRatio(1.5);
            setActiveJustify("justifyLeft");
            reflectMojikumiByIndex(6); // ベタ組み / Solid
            reflectKinsokuById("Soft_v2"); // 弱い禁則 v2 / Loose v2
        };

        // 見出し：メトリクス／行送り 110%／左揃え／文字組みツメ組み／弱い禁則 v2 / Heading preset (mojikumi: Tight, kinsoku: Soft_v2)
        ui.roleHeadingRadio.onClick = function () {
            runApply("applyProfile", { kern: "metrics", ratio: 1.1, justify: "left", mojikumiIndex: 5, kinsoku: "Soft_v2" });
            selectRadioById(ui.kernRadios, autoKernOptions, "metrics");
            reflectLeadingRatio(1.1);
            setActiveJustify("left");
            reflectMojikumiByIndex(5); // ツメ組み / Tight
            reflectKinsokuById("Soft_v2"); // 弱い禁則 v2 / Loose v2
        };

        // 和文：文字揃え＝中央／行送りの基準＝仮想ボディの上 / Japanese: center alignment / top-to-top leading
        ui.roleJapaneseRadio.onClick = function () {
            runApply("applyProfile", { align: "center", leadingType: "top" });
            selectAlignRadio("center");
            selectRadioById(ui.leadingBasisRadios, ui.leadingBasisChoices, "top");
        };

        // 欧文：文字揃え・行送りの基準ともに欧文ベースライン / Roman: Roman baseline for both alignment and leading
        ui.roleRomanRadio.onClick = function () {
            runApply("applyProfile", { align: "roman", leadingType: "baseline" });
            selectAlignRadio("roman");
            selectRadioById(ui.leadingBasisRadios, ui.leadingBasisChoices, "baseline");
        };

        // 混植：文字揃え＝欧文ベースライン／行送りの基準＝仮想ボディの上 / Mixed: Roman-baseline alignment / top-to-top leading
        ui.roleMixedRadio.onClick = function () {
            runApply("applyProfile", { align: "roman", leadingType: "top" });
            selectAlignRadio("roman");
            selectRadioById(ui.leadingBasisRadios, ui.leadingBasisChoices, "top");
        };

        // 制御文字の表示／非表示 / Toggle hidden characters
        ui.hiddenCharButton.onClick = function () {
            runApply("toggleHiddenChar", null);
        };

        // リセット：すべて既定値へ / Reset everything to defaults
        ui.resetButton.onClick = function () {
            runApply("applyProfile", {
                kern: "metrics", tsume: 0, tracking: 0, align: "roman", justify: "left",
                ratio: null, leadingType: "top",
                sizePt: 12, scale: 100, mojikumiIndex: 5, kinsoku: "Soft_v2", clearAki: true
            });
            selectRadioById(ui.kernRadios, autoKernOptions, "metrics");
            reflectTsume(0);
            reflectTracking(0);
            selectAlignRadio("roman");
            setActiveJustify("left");
            reflectLeadingRatio(null);
            selectRadioById(ui.leadingBasisRadios, ui.leadingBasisChoices, "top");
            // 文字組みアキ量設定を「ツメ組み」（mojikumiSet の index 5）へ / Mojikumi to Tight (index 5)
            suppressUiEvents = true;
            for (var mojikumiChoiceIndex = 0; mojikumiChoiceIndex < ui.mojikumiChoices.length; mojikumiChoiceIndex++) {
                if (ui.mojikumiChoices[mojikumiChoiceIndex].index === 5) { selectExclusiveRadio(ui.mojikumiRadios, mojikumiChoiceIndex); break; }
            }
            // 禁則を「弱い禁則 v2」（Soft_v2）へ / Kinsoku to Loose v2 (Soft_v2)
            for (var kinsokuChoiceIndex = 0; kinsokuChoiceIndex < ui.kinsokuChoices.length; kinsokuChoiceIndex++) {
                if (ui.kinsokuChoices[kinsokuChoiceIndex].id === "Soft_v2") { selectExclusiveRadio(ui.kinsokuRadios, kinsokuChoiceIndex); break; }
            }
            suppressUiEvents = false;
            // 文字サイズ 12pt・比率 100% / Font size 12pt, scale 100%
            apparentToggleState = null;
            ui.fontSizeInput.text = String(Math.round((12 / textUnit.factor) * 10) / 10);
            ui.scaleInput.text = "100";
            updateApparentDisplay();
            ui.roleBodyRadio.value = false;
            ui.roleHeadingRadio.value = false;
        };

        // ---- カーニング・文字ツメ / Kerning and Tsume ----
        for (var i = 0; i < ui.kernRadios.length; i++) {
            ui.kernRadios[i].onClick = function () { runApply("apply", { method: autoKernOptions[this.index].id }); };
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
        var apparentToggleState = null; // 焼き込み前の状態（順方向で保存→逆方向で復元）/ Pre-bake state
        function calcApparentSize(size, scale) { return Math.round(size * scale) / 100; }
        function updateApparentDisplay() {
            var size = parseFloat(ui.fontSizeInput.text);
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
        function applyFontSizeFromInput() {
            apparentToggleState = null;
            var inputValue = parseFloat(ui.fontSizeInput.text);
            if (!isNaN(inputValue)) runApply("applyFontSize", { sizePt: inputValue * textUnit.factor });
            updateApparentDisplay();
        }
        function applyScaleFromInput() {
            apparentToggleState = null;
            var inputValue = parseFloat(ui.scaleInput.text);
            if (!isNaN(inputValue)) runApply("applyScale", { scale: inputValue });
            updateApparentDisplay();
        }
        ui.fontSizeInput.onChange = applyFontSizeFromInput;
        ui.fontSizeInput.onChanging = function () { updateApparentDisplay(); };
        changeValueByArrowKey(ui.fontSizeInput, false, applyFontSizeFromInput, 1);
        ui.scaleInput.onChange = applyScaleFromInput;
        ui.scaleInput.onChanging = function () { updateApparentDisplay(); };
        changeValueByArrowKey(ui.scaleInput, false, applyScaleFromInput, 1);

        // ---- 行送り / Leading ----

        /* UI の現在状態から行送りパラメータを組み立てる（その他で値が空なら null）
           Build leading params from the current UI (null if "Other" has no value) */
        function currentLeadingParams() {
            var selectedIndex = selectedRadioIndex(ui.leadingRadios);
            if (selectedIndex < 0) return null;
            var choice = LEADING_CHOICES[selectedIndex];
            var directMode = false, ratio = null, directLeadingPt = NaN;
            if (choice.other) {
                // その他：% 入力をそのまま倍率に / Other: the percentage is used directly as a ratio
                var percentValue = parseFloat(ui.leadingInput.text);
                if (isNaN(percentValue)) return null;
                ratio = percentValue / 100;
            } else if (choice.auto) {
                ratio = null;
            } else {
                ratio = choice.ratio;
            }
            var basisIndex = selectedRadioIndex(ui.leadingBasisRadios);
            if (basisIndex < 0) basisIndex = 0;
            var autoAmount = isNaN(lastAutoAmount) ? 175 : lastAutoAmount;
            return {
                ratio: ratio,
                leadingType: ui.leadingBasisChoices[basisIndex].id,
                autoAmount: autoAmount,
                common: ui.modeCommon.value,
                directMode: directMode,
                directLeadingPt: directLeadingPt
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

        for (var k = 0; k < ui.leadingRadios.length; k++) {
            (function (index) {
                ui.leadingRadios[index].onClick = function () {
                    selectExclusiveRadio(ui.leadingRadios, index);
                    // 「その他」を選んだら現在の行送り（%）を入力欄に補完 / Fill in the current leading (%) when "Other" is chosen
                    if (LEADING_CHOICES[index].other && !isNaN(lastLeadingPercent)) {
                        ui.leadingInput.text = String(lastLeadingPercent);
                    }
                    applyLeading();
                };
            })(k);
        }
        for (var basisRadioIndex = 0; basisRadioIndex < ui.leadingBasisRadios.length; basisRadioIndex++) {
            (function (index) {
                ui.leadingBasisRadios[index].onClick = function () {
                    selectExclusiveRadio(ui.leadingBasisRadios, index);
                    applyLeading();
                };
            })(basisRadioIndex);
        }

        // 個別／共通の切り替え / Individual vs common toggle
        ui.modePerLine.onClick = applyLeading;
        ui.modeCommon.onClick = applyLeading;

        // その他＝行送りを % で直接入力 / Editing the leading percentage selects "Other"
        var otherIndex = -1, autoIndex = -1;
        for (var choiceIndex = 0; choiceIndex < LEADING_CHOICES.length; choiceIndex++) {
            if (LEADING_CHOICES[choiceIndex].other) otherIndex = choiceIndex;
            if (LEADING_CHOICES[choiceIndex].auto) autoIndex = choiceIndex;
        }
        ui.leadingInput.onChange = function () {
            if (otherIndex >= 0) selectExclusiveRadio(ui.leadingRadios, otherIndex);
            applyLeading();
        };

        changeValueByArrowKey(ui.leadingInput, false, function () {
            if (otherIndex >= 0) selectExclusiveRadio(ui.leadingRadios, otherIndex);
            applyLeading();
        }, 1);

        // 自動の値を変更：割合（%）を入力して自動行送り量を更新 / Change the auto-leading amount via a prompt
        // モーダル prompt がパレットのクリックを再入で再配信し二重に開くのを防ぐ
        // Guard against the modal prompt re-delivering the palette click (re-entrant double open)
        var changeAutoBusy = false;
        ui.changeAutoButton.onClick = function () {
            if (changeAutoBusy) return;
            changeAutoBusy = true;
            var input = null;
            try {
                var currentAmount = isNaN(lastAutoAmount) ? 175 : lastAutoAmount;
                input = prompt(getLocalizedText(LABELS.tip.changeAuto), String(currentAmount), getLocalizedText(LABELS.button.changeAuto));
            } finally {
                changeAutoBusy = false;
            }
            if (input === null) return;
            var inputValue = parseFloat(input);
            if (isNaN(inputValue)) return;
            inputValue = Math.round(inputValue);
            if (inputValue < 0) inputValue = 0;
            lastAutoAmount = inputValue;
            // ラジオボタン「自動」の表示も新しい割合に更新 / Update the "Auto" radio label to the new percentage
            if (ui.autoRadio) ui.autoRadio.text = formatAutoLabel(lastAutoAmount);
            // 自動を選択して適用 / Select Auto and apply
            if (autoIndex >= 0) selectExclusiveRadio(ui.leadingRadios, autoIndex);
            applyLeading();
        };

        // 文字組みアキ量設定（ラジオボタン）/ Mojikumi spacing set (radio buttons)
        for (var mojikumiRadioIndex = 0; mojikumiRadioIndex < ui.mojikumiRadios.length; mojikumiRadioIndex++) {
            (function (index) {
                ui.mojikumiRadios[index].onClick = function () {
                    selectExclusiveRadio(ui.mojikumiRadios, index);
                    if (suppressUiEvents) return;
                    runApply("applyMojikumi", { mojikumiIndex: ui.mojikumiChoices[index].index });
                };
            })(mojikumiRadioIndex);
        }

        // 禁則（ラジオボタン）/ Kinsoku (radio buttons)
        for (var kinsokuRadioIndex = 0; kinsokuRadioIndex < ui.kinsokuRadios.length; kinsokuRadioIndex++) {
            (function (index) {
                ui.kinsokuRadios[index].onClick = function () {
                    selectExclusiveRadio(ui.kinsokuRadios, index);
                    if (suppressUiEvents) return;
                    runApply("applyKinsoku", { kinsoku: ui.kinsokuChoices[index].id });
                };
            })(kinsokuRadioIndex);
        }

        /* 選択状態を読み取って UI に反映 / Read the selection state and reflect it into the UI */
        function refreshState() {
            runWorker("getState", null, function (status, payload) {
                if (status !== "ok") return;
                var state = parseLeadingState(payload);
                // 上部インフォバー：テキストが1つだけのとき、値のみを1行表示（タイトルなし）
                // Top info bar: a single line of values only (no titles), shown when exactly one text is selected
                if (state.count === 1) {
                    var sizeText = isNaN(state.fontSizePt) ? "—" : (Math.round((state.fontSizePt / textUnit.factor) * 10) / 10) + textUnit.label;
                    // 行送り：自動は「自動（175%）」、固定は単位付き値のみ / Leading: "Auto (175%)" or a unit value only
                    var leadingText;
                    if (state.isAuto) {
                        leadingText = formatAutoLabel(state.autoAmount);
                    } else if (isNaN(state.leadingPt)) {
                        leadingText = "—";
                    } else {
                        leadingText = (Math.round((state.leadingPt / textUnit.factor) * 10) / 10) + textUnit.label;
                    }
                    // 1行目：フォント（ファミリー＋スタイル）・フォントサイズ・行送り
                    // Line 1: font (family + style), font size, leading
                    var infoLine1 = [
                        state.fontFamily || "—",
                        state.fontStyle || "—",
                        sizeText,
                        leadingText
                    ].join(" / ");
                    // 2行目：その他（行揃え・カーニング・禁則）/ Line 2: the rest (justify, kerning, kinsoku)
                    var infoLine2 = [
                        justifyLabel(state.justifyId),
                        kernLabel(state.kernId),
                        kinsokuLabel(state.kinsokuId)
                    ].join(" / ");
                    ui.infoText.text = infoLine1;
                    ui.infoTextSub.text = infoLine2;
                } else {
                    ui.infoText.text = "";
                    ui.infoTextSub.text = "";
                }
                if (state.count <= 0) return;
                // フォントサイズ・比率・見かけ / Font size, scale, apparent
                apparentToggleState = null;
                ui.fontSizeInput.text = isNaN(state.fontSizePt) ? "" : String(Math.round((state.fontSizePt / textUnit.factor) * 10) / 10);
                ui.scaleInput.text = isNaN(state.hScale) ? "100" : String(Math.round(state.hScale * 10) / 10);
                updateApparentDisplay();
                // 行送りラジオ / Leading radio
                var matchedLeadingIndex = findMatchingLeadingChoiceIndex(state.isAuto, state.leadingPt, state.baseFontSizePt);
                selectExclusiveRadio(ui.leadingRadios, matchedLeadingIndex);
                // 現在の行送り（%）を控えておく（「その他」クリック時の初期値に使う）/ Remember the current leading (%) for the "Other" default
                if (!isNaN(state.leadingPt) && !isNaN(state.baseFontSizePt) && state.baseFontSizePt > 0) {
                    lastLeadingPercent = Math.round((state.leadingPt / state.baseFontSizePt) * 100);
                } else {
                    lastLeadingPercent = NaN;
                }
                // 「その他」が一致したときだけ、現在の行送りを % で表示 / Show the ratio as % only when "Other" matches
                if (LEADING_CHOICES[matchedLeadingIndex] && LEADING_CHOICES[matchedLeadingIndex].other && !isNaN(lastLeadingPercent)) {
                    ui.leadingInput.text = String(lastLeadingPercent);
                } else {
                    ui.leadingInput.text = "";
                }
                if (!isNaN(state.autoAmount)) {
                    lastAutoAmount = Math.round(state.autoAmount);
                    if (ui.autoRadio) ui.autoRadio.text = formatAutoLabel(lastAutoAmount);
                }
                // 共通／個別（段落が1つだけなら共通）/ Common vs individual (common when there is a single paragraph)
                ui.modeCommon.value = (state.paragraphCount === 1);
                ui.modePerLine.value = !ui.modeCommon.value;
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
                    selectExclusiveRadio(ui.mojikumiRadios, mojikumiSelectionIndex);
                    suppressUiEvents = false;
                }
                // 禁則 / Kinsoku
                var kinsokuSelectionIndex = 0;
                for (var kinsokuChoiceIndex = 0; kinsokuChoiceIndex < ui.kinsokuChoices.length; kinsokuChoiceIndex++) {
                    if (ui.kinsokuChoices[kinsokuChoiceIndex].id === state.kinsokuId) { kinsokuSelectionIndex = kinsokuChoiceIndex; break; }
                }
                suppressUiEvents = true;
                selectExclusiveRadio(ui.kinsokuRadios, kinsokuSelectionIndex);
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

        /* ドキュメントフォント一覧を読み直して listbox に反映 / Reload the document fonts into the listbox
           選択の復元はしない（プログラムからの再選択が onChange を発火し再適用するのを避ける）
           Selection is not restored, to avoid a programmatic re-select firing onChange and re-applying */
        function refreshFonts() {
            runWorker("getFonts", null, function (status, payload) {
                if (status !== "ok") return;
                var fontData = decodeURIComponent(payload);
                ui.fontList.removeAll();
                if (fontData === "") return;
                var rows = fontData.split(String.fromCharCode(10));
                for (var i = 0; i < rows.length; i++) {
                    var tabIndex = rows[i].indexOf(String.fromCharCode(9));
                    if (tabIndex < 0) continue;
                    var listItem = ui.fontList.add("item", rows[i].substring(tabIndex + 1));
                    listItem.psName = rows[i].substring(0, tabIndex);
                }
            });
        }

        /* プリセットの表示名を和文フォント名に差し替え（1回だけ）/ Replace preset labels with Japanese font names (once) */
        var presetLabelsResolved = false;
        function refreshPresetLabels() {
            if (presetLabelsResolved) return;
            var psList = "";
            for (var i = 0; i < ui.presetList.items.length; i++) {
                psList += (i ? String.fromCharCode(10) : "") + ui.presetList.items[i].psName;
            }
            runWorker("resolveFontNames", { psList: psList }, function (status, payload) {
                if (status !== "ok") return;
                var resolvedData = decodeURIComponent(payload);
                if (resolvedData === "") return;
                var rows = resolvedData.split(String.fromCharCode(10));
                for (var rowIndex = 0; rowIndex < rows.length && rowIndex < ui.presetList.items.length; rowIndex++) {
                    var tabIndex = rows[rowIndex].indexOf(String.fromCharCode(9));
                    if (tabIndex >= 0) ui.presetList.items[rowIndex].text = rows[rowIndex].substring(tabIndex + 1);
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
        ui.palette.onShow = function () { refreshJustifyTheme(); refreshState(); refreshFonts(); refreshPresetLabels(); };
        ui.palette.onActivate = function () { refreshJustifyTheme(); refreshState(); refreshFonts(); refreshPresetLabels(); };

        // 再読み込み：選択中のテキストの現在値を読み直して UI に反映 / Reload: re-read the selection's current values into the UI
        ui.reloadButton.onClick = function () { refreshState(); refreshFonts(); refreshPresetLabels(); };

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
    }

    main();

})();
