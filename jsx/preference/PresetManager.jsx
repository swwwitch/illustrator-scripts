#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

(function () {

    /*
    
    ### スクリプト名：
    
    Illustrator環境設定ユーティリティ
    PresetManager.jsx
    
    ### GitHub：
    
    https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/preference/PresetManager.jsx
    
    ### 概要：

    - Illustrator の主要な環境設定を、カテゴリ別の GUI からまとめて確認・変更します。
    - 現在の設定の読み込みに加え、［デフォルト］／［プリセット1］をワンクリックで UI に反映できます。
    - 一般、選択範囲・アンカー表示、アートボード、テキスト、ガイド、スマートガイド、UI、パフォーマンス、ファイル管理、クリップボード、選択範囲の制限をカバーします。
    - アートボード名表示、ハイライトカラー／幅、ガイド、保存先なども一括調整できます。
    
    ### 主な機能：
    
    - 詳細なツールヒント／ホーム画面／レガシー新規ドキュメント UI の切替
    - アートボード移動時のロック/非表示オブジェクト追従、アートボード名表示、アートボード境界線のカラー・幅の設定
    - 選択範囲へズーム、アンカー強調表示、オブジェクト選択・テキスト選択の「パスに制限」
    - テキスト：エリア内文字の自動サイズ、最近使用フォントの表示数、異体字・見つからない字形の保護
    - ガイド：カラー（シアン／ライトブルー）とスタイル（ライン／点線）の切替
    - ユーザーインターフェイス：カンバスカラー（UIに合わせる／ホワイト）
    - パフォーマンス：アニメーションズーム、ヒストリー数、リアルタイム描画
    - スマートガイド：オブジェクトのハイライト表示のオン／オフ
    - ファイル管理：編集の既定アプリ、Adobe Fonts 自動アクティベート、ファイルの保存先（コンピューター／クラウド）
    - クリップボード：SVG コードを含めるのオン／オフ
    
    ### 処理の流れ：
    
    1) ダイアログ生成 → 2) 現在の環境設定を読み込み UI に反映 → 3) プリセット選択で UI を上書き → 4) ［OK］で各プリファレンスキーへ保存。
    
    ### 紹介記事（note）
    
    https://note.com/dtp_tranist/n/n3b33862538f6
    
    ### 更新日：
    
    - 2026-04-22（v1.7.0）
    
    ### 更新履歴：
    
    - v1.0 (20250807) : 初期リリース
    - v1.6（2026-03-23）：UI配置調整、ローカライズ改善、不要機能削除、キー修正（useSysDefEdit）
    - v1.7.0（2026-04-22）：ローカライズ、命名、コメント、構造を整理
    
    ---
    
    ### Script Name:
    
    Illustrator Preferences Utility
    
    ### GitHub:
    
    https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/Preference/PresetManager.jsx
    
    ### Overview:

    - Review and change major Illustrator preferences from a categorized GUI in one place.
    - Load current settings or apply [Default] / [Preset 1] to the UI with a single selection.
    - Covers General, Selection & Anchor Display, Artboard, Type, Guides, Smart Guides, UI, Performance, File Management, Clipboard, and selection limiting.
    - Also supports unified adjustment of artboard name visibility, highlight color/width, guides, and save location.
    
    ### Key Features:
    
    - Toggle Rich Tool Tips, Home Screen, and Legacy New Document UI
    - Configure moving locked/hidden artwork with artboards, artboard name visibility, and artboard boundary color/width
    - Control Zoom to Selection, anchor highlighting, and path-only selection for objects and type
    - Type: Auto Size Area Type, recent fonts count, alternate glyphs, and missing glyph protection
    - Guides: switch color (Cyan / Light Blue) and style (Lines / Dots)
    - User Interface: Canvas Color (Match Brightness / White)
    - Performance: Animated Zoom, History States, Real-Time Drawing and Editing
    - Smart Guides: Object Highlighting on/off
    - File Management: system default app for Edit Original, auto-activate Adobe Fonts, and save location (Computer / Cloud)
    - Clipboard: include SVG code on/off
    
    ### Flow:
    
    1) Build dialog → 2) Load current preferences into UI → 3) Apply preset to UI → 4) Save preferences when pressing OK.
    
    ### Updated:
    
    - 2026-04-22 (v1.7.0)
    
    ### Changelog:
    
    - v1.0 (2025-08-07): Initial release
    - v1.6 (2026-03-23): UI layout adjustments, localization improvements, removed incomplete features, fixed preference key (useSysDefEdit)
    - v1.7.0 (2026-04-22): Refined localization, naming, comments, and structure
    
    */
    var SCRIPT_VERSION = "v1.7.0";

    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var lang = getCurrentLang();

    // Label helper function for localization
    function L(key) {
        var item = LABELS[key];
        if (!item) return key; // 保険（キー名をそのまま返す）

        if (item[lang]) return item[lang];
        if (item.en) return item.en;

        return key;
    }

    /* 日英ラベル定義 / Japanese-English label definitions */

    var LABELS = {
        dialogTitle: {
            ja: "環境設定をまとめて変更",
            en: "Illustrator Preferences Utility"
        },
        OK: {
            ja: "OK",
            en: "OK"
        },
        Cancel: {
            ja: "キャンセル",
            en: "Cancel"
        },
        presetLabel: {
            ja: "プリセット：",
            en: "Preset:"
        },
        presetCurrent: {
            ja: "現在の設定",
            en: "Current Settings"
        },
        presetDefault: {
            ja: "デフォルト",
            en: "Default"
        },
        preset1: {
            ja: "プリセット1",
            en: "Preset 1"
        },
        // ［一般］
        cbToolTips: {
            ja: "詳細なツールヒントを表示",
            en: "Show Rich Tool Tips"
        },
        cbHomeScreen: {
            ja: "ドキュメントを開いていないときに「ホーム画面」を表示",
            en: "Show the Home Screen When No Documents Are Open"
        },
        cbLegacyNewDoc: {
            ja: "以前の「新規ドキュメント」インターフェイスを使用",
            en: "Use Legacy \"File > New\" Interface"
        },
        cbBleedAI: {
            ja: "「裁ち落としを印刷」生成AIボタンを表示",
            en: "Show 'Print Bleed' generative AI buttons on Bleed"
        },

        // ［選択範囲・アンカー表示］
        cbMoveLocked: {
            ja: "ロックまたは非表示オブジェクトを一緒に移動",
            en: "Move Locked and Hidden Artwork"
        },
        cbHitShape: {
            ja: "オブジェクトの選択範囲をパスに制限",
            en: "Object Selection by Path Only"
        },
        cbZoomToSel: {
            ja: "選択範囲へズーム",
            en: "Zoom to Selection"
        },
        // スマートガイド
        smartHighlight: {
            ja: "オブジェクトのハイライト表示",
            en: "Object Highlighting"
        },
        // テキスト
        cbHitTypeShape: {
            ja: "テキストオブジェクトの選択範囲をパスに制限",
            en: "Type Object Selection by Path Only"
        },
        cbAutoSizing: {
            ja: "新規エリア内文字の自動サイズ調整",
            en: "Auto Size New Area Type"
        },
        cbRecentFonts: {
            ja: "最近使用したフォントの表示数",
            en: "Number of Recent Fonts"
        },
        cbFontLocking: {
            ja: "見つからない字形の保護を有効にする",
            en: "Enable Missing Glyph Protection"
        },
        cbAlternateGlyph: {
            ja: "選択された文字の異体字を表示",
            en: "Show Character Alternates"
        },

        // ユーザーインターフェイス
        canvasColor: {
            ja: "カンバスカラー",
            en: "Canvas Color"
        },
        canvasMatch: {
            ja: "UIに合わせる",
            en: "Match Brightness"
        },
        canvasWhite: {
            ja: "ホワイト",
            en: "White"
        },

        // パフォーマンス
        cbAnimZoom: {
            ja: "アニメーションズーム",
            en: "Animated Zoom"
        },
        cbLiveEdit: {
            ja: "リアルタイムの描画と編集",
            en: "Real-Time Drawing and Editing"
        },
        cbHistoryLabel: {
            ja: "ヒストリー数",
            en: "History States"
        },

        // ファイル管理
        cbEditOriginal: {
            ja: "「オリジナルの編集」にシステムデフォルトを使用",
            en: "Use System Defaults for ‘Edit Original’"
        },
        cbFontsAuto: {
            ja: "Adobe Fonts を自動アクティベート",
            en: "Auto-activate Adobe Fonts"
        },
        saveDest: {
            ja: "ファイルの保存先",
            en: "Save Location"
        },
        saveLocal: {
            ja: "コンピューター",
            en: "Computer"
        },
        saveCloud: {
            ja: "クラウド",
            en: "Cloud"
        },
        cbIncludeSVG: {
            ja: "SVGコードを含める",
            en: "Include SVG Code"
        },
        updateLinks: {
            ja: "リンクを更新",
            en: "Update Links"
        },
        updateLinksAuto: {
            ja: "自動",
            en: "Automatic"
        },
        updateLinksManual: {
            ja: "手動",
            en: "Manual"
        },
        updateLinksAsk: {
            ja: "確認",
            en: "Ask When Modified"
        },

        // パネルタイトル / Panel titles
        panelGeneralTitle: {
            ja: "［一般］カテゴリ",
            en: "[General] Category"
        },
        panelSelectAnchorTitle: {
            ja: "［選択範囲・アンカー表示］カテゴリ",
            en: "[Selection & Anchor Display] Category"
        },
        panelTextTitle: {
            ja: "［テキスト］カテゴリ",
            en: "[Text] Category"
        },
        panelGuidesTitle: {
            ja: "ガイド",
            en: "Guides"
        },
        panelUITitle: {
            ja: "［ユーザーインターフェイス］カテゴリ",
            en: "[User Interface] Category"
        },
        panelPerfTitle: {
            ja: "［パフォーマンス］カテゴリ",
            en: "[Performance] Category"
        },
        panelFileTitle: {
            ja: "［ファイル管理］カテゴリ",
            en: "[File Management] Category"
        },
        panelClipboardTitle: {
            ja: "クリップボードの処理",
            en: "Clipboard Handling"
        },
        panelLimitPathTitle: {
            ja: "パスに制限",
            en: "Limit to Path"
        },
        panelSmartGuideTitle: {
            ja: "スマートガイド",
            en: "Smart Guides"
        },
        // ガイドパネル内ラベル / Guides panel inner labels
        guideColorLabel: {
            ja: "カラー：",
            en: "Color:"
        },
        guideColorCyan: {
            ja: "シアン",
            en: "Cyan"
        },
        guideColorLightBlue: {
            ja: "ライトブルー",
            en: "Light Blue"
        },
        guideStyleLabel: {
            ja: "スタイル：",
            en: "Style:"
        },
        guideStyleLine: {
            ja: "ライン",
            en: "Lines"
        },
        guideStyleDots: {
            ja: "点線",
            en: "Dots"
        },
        // アートボード / Artboard
        panelArtboardTitle: {
            ja: "アートボード",
            en: "Artboard"
        },
        cbShowArtboardName: {
            ja: "アートボード名を表示",
            en: "Show Artboard Name"
        },
        artboardStrokeColor: {
            ja: "ハイライトのカラー",
            en: "Highlight Color"
        },
        artboardStrokeWidth: {
            ja: "ストロークの幅",
            en: "Stroke Width"
        },
        artboardColorBlack: {
            ja: "ブラック",
            en: "Black"
        },
        artboardColorLightBlue: {
            ja: "ライトブルー",
            en: "Light Blue"
        },
        artboardColorRed: {
            ja: "サーモンピンク",
            en: "Light Red"
        },
        artboardColorGreen: {
            ja: "グリーン",
            en: "Green"
        },
        artboardColorBlue: {
            ja: "ミディアムブルー",
            en: "Medium Blue"
        },
        artboardColorCyan: {
            ja: "シアン",
            en: "Cyan"
        },
        artboardColorMagenta: {
            ja: "マゼンタ",
            en: "Magenta"
        },
        artboardColorYellow: {
            ja: "イエロー",
            en: "Yellow"
        },
        artboardColorWhite: {
            ja: "ライトグレー",
            en: "Light Gray"
        },
        historyHelpTip: {
            ja: "ヒストリー数を設定",
            en: "Set history states"
        },
        alertSavePrefsFailed: {
            ja: "環境設定の保存に失敗しました: ",
            en: "Failed to save preferences: "
        }
    };


    function main() {
        var PANEL_MARGINS = [15, 20, 15, 10];

        /*
      ダイアログ生成 / Build main dialog
    */
        var dlg = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
        /*
      メイングループ / Main vertical container
    */
        var mainGroup = dlg.add("group");
        mainGroup.orientation = "column";
        mainGroup.alignChildren = "left";

        /*
      プリセット選択（Current / Default / Preset 1） / Preset selector (Current / Default / Preset 1)
    */
        var groupPreset = mainGroup.add("group");
        groupPreset.name = "groupPreset";
        groupPreset.alignment = "center";
        groupPreset.orientation = "row";
        groupPreset.margins = [0, 10, 20, 20];
        var groupPresetInner = groupPreset.add("group");
        groupPresetInner.orientation = "row";
        groupPresetInner.alignChildren = "center";
        groupPresetInner.add("statictext", undefined, L("presetLabel"));
        var PRESET_IDS = {
            current: "current",
            defaultPreset: "default",
            preset1: "preset1"
        };
        var presetItems = [
            L("presetCurrent"),
            L("presetDefault"),
            L("preset1")
        ];
        var ddPreset = groupPresetInner.add("dropdownlist", undefined, presetItems);
        ddPreset.selection = 0; // 初期選択は「現在の設定」 / Default selection is "Current Settings"

        /*
        Apply preset settings to the UI / プリセット適用
        - current は現在の環境設定を読んでUIへ反映
        - default / preset1 は定義オブジェクトから反映
      */
        function applyPresetToUIFromId(presetId) {
            if (presetId === PRESET_IDS.current) {
                /*
                  現在の環境設定を UI に反映する
                */
                loadPrefsFromSystem();

                /*
                  カンバスカラー
                */
                var currentCanvasPref = app.preferences.getIntegerPreference("uiCanvasIsWhite");
                rbCanvasWhite.value = (currentCanvasPref === 1);
                rbCanvasMatch.value = (currentCanvasPref !== 1);
                // ガイド色の現在値でラジオ表示だけ合わせる（書き込みなし）
                var r = safeGetReal("Guide/Color/red", 0.0);
                var g = safeGetReal("Guide/Color/green", 1.0);
                var b = safeGetReal("Guide/Color/blue", 1.0);

                var isLightBlue = approx(r, 0.29) && approx(g, 0.52) && approx(b, 1.0);
                if (typeof rbGuideLightBlue !== 'undefined' && typeof rbGuideCyan !== 'undefined') {
                    rbGuideLightBlue.value = !!isLightBlue;
                    rbGuideCyan.value = !isLightBlue;
                }
                /*
                  ファイルの保存先（現在値をUIに反映）
                */
                var curCloud = safeGetBool("AdobeSaveAsCloudDocumentPreference", false);
                rbSaveCloud.value = curCloud;
                rbSaveLocal.value = !curCloud;
                /*
                  リンクを更新（現在値をUIに反映）
                */
                var curLinks = safeGetInt("plugin/FileClipboard/linkoptions", 2);
                rbUpdateLinksAuto.value = (curLinks === 0);
                rbUpdateLinksManual.value = (curLinks === 1);
                rbUpdateLinksAsk.value = (curLinks !== 0 && curLinks !== 1);

            } else if (presetId === PRESET_IDS.preset1) {
                applyPresetToUI(PRESET1_UI_STATE);
            } else if (presetId === PRESET_IDS.defaultPreset) {
                applyPresetToUI(DEFAULT_UI_STATE);
            }
        }

        function getSelectedPresetId(selectionIndex) {
            if (selectionIndex === 0) return PRESET_IDS.current;
            if (selectionIndex === 1) return PRESET_IDS.defaultPreset;
            if (selectionIndex === 2) return PRESET_IDS.preset1;
            return PRESET_IDS.current;
        }

        ddPreset.onChange = function () {
            if (ddPreset.selection) {
                applyPresetToUIFromId(getSelectedPresetId(ddPreset.selection.index));
                app.redraw();
            }
        };


        /*
      2カラム / Two-column layout
    */
        var columnsGroup = mainGroup.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = "top";

        var leftColumnGroup = columnsGroup.add("group");
        leftColumnGroup.orientation = "column";
        leftColumnGroup.alignChildren = ["fill", "top"];

        var rightColumnGroup = columnsGroup.add("group");
        rightColumnGroup.orientation = "column";
        rightColumnGroup.alignChildren = ["fill", "top"];

        /*
      ［一般］パネル（左） / General panel (left)
    */
        var panelGeneral = leftColumnGroup.add("panel", undefined, L("panelGeneralTitle"));
        panelGeneral.orientation = "column";
        panelGeneral.alignChildren = ["fill", "top"];
        panelGeneral.alignment = ["fill", "top"];
        panelGeneral.margins = PANEL_MARGINS;

        var cbToolTips = panelGeneral.add("checkbox", undefined, L("cbToolTips"));
        cbToolTips.helpTip = L("cbToolTips");

        var cbHomeScreen = panelGeneral.add("checkbox", undefined, L("cbHomeScreen"));
        cbHomeScreen.helpTip = L("cbHomeScreen");

        var cbLegacyNewDoc = panelGeneral.add("checkbox", undefined, L("cbLegacyNewDoc"));
        cbLegacyNewDoc.helpTip = L("cbLegacyNewDoc");

        var cbBleedAI = panelGeneral.add("checkbox", undefined, L("cbBleedAI"));
        cbBleedAI.helpTip = L("cbBleedAI");

        /*
      ［選択範囲・アンカー表示］パネル（左） / Selection & Anchor panel (left)
    */
        var panelSelectAnchor = leftColumnGroup.add("panel", undefined, L("panelSelectAnchorTitle"));
        panelSelectAnchor.orientation = "column";
        panelSelectAnchor.alignChildren = ["fill", "top"];
        panelSelectAnchor.alignment = ["fill", "top"];
        panelSelectAnchor.margins = PANEL_MARGINS;

        var cbZoomToSel = panelSelectAnchor.add("checkbox", undefined, L("cbZoomToSel"));
        cbZoomToSel.helpTip = L("cbZoomToSel");

        /*
      ［アートボード］パネル（左） / Artboard panel (left)
    */
        var panelArtboard = leftColumnGroup.add("panel", undefined, L("panelArtboardTitle"));
        panelArtboard.orientation = "column";
        panelArtboard.alignChildren = ["fill", "top"];
        panelArtboard.alignment = ["fill", "top"];
        panelArtboard.margins = PANEL_MARGINS;

        var cbMoveLocked = panelArtboard.add("checkbox", undefined, L("cbMoveLocked"));
        cbMoveLocked.helpTip = L("cbMoveLocked");

        var cbShowArtboardName = panelArtboard.add("checkbox", undefined, L("cbShowArtboardName"));
        cbShowArtboardName.helpTip = L("cbShowArtboardName");

        // ストロークのカラー（ドロップダウン） / Stroke color (dropdown)
        var strokeColorRow = panelArtboard.add("group");
        strokeColorRow.orientation = "row";
        strokeColorRow.alignChildren = ["left", "center"];
        strokeColorRow.add("statictext", undefined, L("artboardStrokeColor") + "：");

        /* アートボードのハイライトカラープリセット（RGB 0..1） / Artboard stroke color presets (RGB 0..1) */
        var STROKE_COLOR_PRESETS = [
            { label: L("artboardColorLightBlue"), r: 0.29, g: 0.52, b: 1.0 },
            { label: L("artboardColorRed"), r: 1.0, g: 0.29, b: 0.29 },
            { label: L("artboardColorGreen"), r: 0.0, g: 0.65, b: 0.31 },
            { label: L("artboardColorBlue"), r: 0.0, g: 0.45, b: 0.78 },
            { label: L("artboardColorMagenta"), r: 1.0, g: 0.0, b: 1.0 },
            { label: L("artboardColorCyan"), r: 0.0, g: 1.0, b: 1.0 },
            { label: L("artboardColorWhite"), r: 1.0, g: 1.0, b: 1.0 },
            { label: L("artboardColorBlack"), r: 0.0, g: 0.0, b: 0.0 },
            { label: L("artboardColorYellow"), r: 1.0, g: 1.0, b: 0.0 }
        ];
        var STROKE_COLOR_NAMES = [];
        for (var sc = 0; sc < STROKE_COLOR_PRESETS.length; sc++) {
            STROKE_COLOR_NAMES.push(STROKE_COLOR_PRESETS[sc].label);
        }
        var ddStrokeColor = strokeColorRow.add("dropdownlist", undefined, STROKE_COLOR_NAMES);

        // 現在のRGB値から最も近いプリセットを選択 / Select the nearest preset from current RGB
        function findClosestStrokeColor(r, g, b) {
            var bestIdx = 0;
            var bestDist = Infinity;
            for (var i = 0; i < STROKE_COLOR_PRESETS.length; i++) {
                var p = STROKE_COLOR_PRESETS[i];
                var dist = Math.abs(p.r - r) + Math.abs(p.g - g) + Math.abs(p.b - b);
                if (dist < bestDist) {
                    bestDist = dist;
                    bestIdx = i;
                }
            }
            return bestIdx;
        }

        var curSCR = safeGetReal("ArtboardBBColorRed", 0.0);
        var curSCG = safeGetReal("ArtboardBBColorGreen", 0.0);
        var curSCB = safeGetReal("ArtboardBBColorBlue", 0.0);
        ddStrokeColor.selection = findClosestStrokeColor(curSCR, curSCG, curSCB);

        // ストロークの幅（1〜4、ラジオボタン） / Stroke width (1–4, radio buttons)
        var strokeWidthRow = panelArtboard.add("group");
        strokeWidthRow.orientation = "row";
        strokeWidthRow.alignChildren = ["left", "center"];
        strokeWidthRow.add("statictext", undefined, L("artboardStrokeWidth") + "：");
        var rbStrokeWidth1 = strokeWidthRow.add("radiobutton", undefined, "1");
        var rbStrokeWidth2 = strokeWidthRow.add("radiobutton", undefined, "2");
        var rbStrokeWidth3 = strokeWidthRow.add("radiobutton", undefined, "3");
        var rbStrokeWidth4 = strokeWidthRow.add("radiobutton", undefined, "4");
        var rbStrokeWidths = [rbStrokeWidth1, rbStrokeWidth2, rbStrokeWidth3, rbStrokeWidth4];

        var curStrokeWidth = Math.round(safeGetReal("ArtboardBBWidth", 1.0));
        var swIdx = clamp(curStrokeWidth, 1, 4) - 1;
        rbStrokeWidths[swIdx].value = true;

        function getSelectedStrokeWidth() {
            for (var i = 0; i < rbStrokeWidths.length; i++) {
                if (rbStrokeWidths[i].value) return i + 1;
            }
            return 1;
        }


        /*
      ［テキスト］パネル（左） / Text panel (left)
    */
        var panelText = leftColumnGroup.add("panel", undefined, L("panelTextTitle"));
        panelText.orientation = "column";
        panelText.alignChildren = ["fill", "top"];
        panelText.alignment = ["fill", "top"];
        panelText.margins = PANEL_MARGINS;

        var cbAutoSizing = panelText.add("checkbox", undefined, L("cbAutoSizing"));
        cbAutoSizing.helpTip = L("cbAutoSizing");


        var currentRecentCount = app.preferences.getIntegerPreference("text/recentFontMenu/showNEntries");
        var recentFontsGroup = panelText.add("group");
        recentFontsGroup.orientation = "row";

        var cbRecentFonts = recentFontsGroup.add("checkbox", undefined, L("cbRecentFonts"));
        cbRecentFonts.helpTip = L("cbRecentFonts");
        cbRecentFonts.value = (currentRecentCount > 0);

        var etRecentFonts = recentFontsGroup.add("edittext", undefined, currentRecentCount.toString());
        etRecentFonts.characters = 3;
        etRecentFonts.enabled = cbRecentFonts.value;

        cbRecentFonts.onClick = function () {
            // UI only: enable/disable and default value. Persist on OK.
            if (cbRecentFonts.value) {
                etRecentFonts.enabled = true;
                validateEditTextInt(etRecentFonts, VALIDATION.recentFonts);
            } else {
                etRecentFonts.enabled = false;
                etRecentFonts.text = "0"; // OFF時は 0 固定（保存はOK時）
            }
        };

        etRecentFonts.onChange = function () {
            // UI only validation. Persist on OK.
            validateEditTextInt(etRecentFonts, VALIDATION.recentFonts);
        };

        var cbFontLocking = panelText.add("checkbox", undefined, L("cbFontLocking"));
        cbFontLocking.helpTip = L("cbFontLocking");

        var cbAlternateGlyph = panelText.add("checkbox", undefined, L("cbAlternateGlyph"));
        cbAlternateGlyph.helpTip = L("cbAlternateGlyph");


        /*
      ［ガイド］パネル（右） / Guides panel (right)
    */
        var panelGuides = rightColumnGroup.add("panel", undefined, L("panelGuidesTitle"));
        panelGuides.orientation = "column";
        panelGuides.alignChildren = ["fill", "top"];
        panelGuides.alignment = ["fill", "top"];
        panelGuides.margins = PANEL_MARGINS;

        /* ガイド内ラベル共通幅 / Unified label width inside Guides */
        var GUIDE_LABEL_WIDTH = 80;

        // ガイド：カラー設定（1行レイアウト） / Guide color settings (single-row layout)
        var guideColorRow = panelGuides.add("group", undefined, "");
        guideColorRow.orientation = "row";
        guideColorRow.alignChildren = ["left", "center"];
        guideColorRow.spacing = 10;

        var lblGuideColor = guideColorRow.add("statictext", undefined, L("guideColorLabel"));
        lblGuideColor.preferredSize = [GUIDE_LABEL_WIDTH, lblGuideColor.preferredSize ? lblGuideColor.preferredSize[1] : 20];
        lblGuideColor.justify = "right";
        var rbGuideCyan = guideColorRow.add("radiobutton", undefined, L("guideColorCyan"));
        var rbGuideLightBlue = guideColorRow.add("radiobutton", undefined, L("guideColorLightBlue"));

        /* 既知の色プリセット（0..1） / Known color presets (0..1) */
        var GUIDE_CYAN = {
            r: 0.0,
            g: 1.0,
            b: 1.0
        };
        var GUIDE_LIGHTBLUE = {
            r: 0.29,
            g: 0.52,
            b: 1.0
        };

        function setGuideColor(r01, g01, b01) {
            safeSetReal("Guide/Color/red", r01);
            safeSetReal("Guide/Color/green", g01);
            safeSetReal("Guide/Color/blue", b01);
        }


        /* ガイド色の取得元はここ（safeGetReal に統一） / Read current guide color once (source of truth) */
        var curR = safeGetReal("Guide/Color/red", GUIDE_CYAN.r);
        var curG = safeGetReal("Guide/Color/green", GUIDE_CYAN.g);
        var curB = safeGetReal("Guide/Color/blue", GUIDE_CYAN.b);

        if (approx(curR, GUIDE_LIGHTBLUE.r) && approx(curG, GUIDE_LIGHTBLUE.g) && approx(curB, GUIDE_LIGHTBLUE.b)) {
            rbGuideLightBlue.value = true;
        } else {
            rbGuideCyan.value = true; // 既定（シアン）
        }


        rbGuideCyan.onClick = function () {
            // UI only; persist on OK
        };
        rbGuideLightBlue.onClick = function () {
            // UI only; persist on OK
        };

        // ガイド：スタイル設定（1行レイアウト） / Guide style settings (single-row layout)
        var guideStyleRow = panelGuides.add("group", undefined, "");
        guideStyleRow.orientation = "row";
        guideStyleRow.alignChildren = ["left", "center"];
        guideStyleRow.spacing = 10;

        var lblGuideStyle = guideStyleRow.add("statictext", undefined, L("guideStyleLabel"));
        lblGuideStyle.preferredSize = [GUIDE_LABEL_WIDTH, lblGuideStyle.preferredSize ? lblGuideStyle.preferredSize[1] : 20];
        lblGuideStyle.justify = "right";
        var rbGuideStyleLine = guideStyleRow.add("radiobutton", undefined, L("guideStyleLine"));
        var rbGuideStyleDots = guideStyleRow.add("radiobutton", undefined, L("guideStyleDots"));

        // Guide/Style: 0 = Lines, 1 = Dots
        var curStyle = 0;
        try {
            curStyle = app.preferences.getIntegerPreference("Guide/Style");
        } catch (e) {
            curStyle = 0;
        }
        rbGuideStyleLine.value = (curStyle !== 1);
        rbGuideStyleDots.value = (curStyle === 1);

        rbGuideStyleLine.onClick = function () {
            // UI only; persist on OK
        };
        rbGuideStyleDots.onClick = function () {
            // UI only; persist on OK
        };


        /*
      ［スマートガイド］パネル（右） / Smart Guides panel (right)
    */
        var panelSmartGuide = rightColumnGroup.add("panel", undefined, L("panelSmartGuideTitle"));
        panelSmartGuide.orientation = "column";
        panelSmartGuide.alignChildren = ["fill", "top"];
        panelSmartGuide.alignment = ["fill", "top"];
        panelSmartGuide.margins = PANEL_MARGINS;

        var cbSmartHighlight = panelSmartGuide.add("checkbox", undefined, L("smartHighlight"));
        cbSmartHighlight.helpTip = L("smartHighlight");

        /*
      ［ユーザーインターフェイス］パネル（左） / UI panel (left)
    */
        var panelUI = leftColumnGroup.add("panel", undefined, L("panelUITitle"));
        panelUI.orientation = "column";
        panelUI.alignChildren = ["fill", "top"];
        panelUI.alignment = ["fill", "top"];
        panelUI.margins = PANEL_MARGINS;

        /* カンバスカラー（UIパネル） / Canvas color (UI panel) */
        // NOTE: All controls update UI only; actual writes occur on OK.
        // カンバスカラー（1行レイアウト）
        var canvasRow = panelUI.add("group", undefined, "");
        canvasRow.orientation = "row";
        canvasRow.alignChildren = ["left", "center"];
        canvasRow.spacing = 10;

        canvasRow.add("statictext", undefined, L("canvasColor") + "：");
        var rbCanvasMatch = canvasRow.add("radiobutton", undefined, L("canvasMatch"));
        var rbCanvasWhite = canvasRow.add("radiobutton", undefined, L("canvasWhite"));

        /*
          初期状態を反映
        */
        var currentCanvasValue = safeGetInt("uiCanvasIsWhite", 0);
        rbCanvasWhite.value = (currentCanvasValue === 1);
        rbCanvasMatch.value = (currentCanvasValue !== 1);
        /*
          ラジオボタン変更時の設定
        */
        rbCanvasMatch.onClick = function () {
            // UI only; persist on OK
        };
        rbCanvasWhite.onClick = function () {
            // UI only; persist on OK
        };


        /*
      ［パフォーマンス］パネル（右） / Performance panel (right)
    */
        var panelPerf = rightColumnGroup.add("panel", undefined, L("panelPerfTitle"));
        panelPerf.orientation = "column";
        panelPerf.alignChildren = ["fill", "top"];
        panelPerf.alignment = ["fill", "top"];
        panelPerf.margins = PANEL_MARGINS;

        var cbAnimZoom = panelPerf.add("checkbox", undefined, L("cbAnimZoom"));
        cbAnimZoom.helpTip = L("cbAnimZoom");

        var historyGroup = panelPerf.add("group");
        historyGroup.add("statictext", undefined, L("cbHistoryLabel") + "：");
        var etHistory = historyGroup.add("edittext", undefined, "50");

        etHistory.onChange = function () {
            validateEditTextInt(etHistory, VALIDATION.history);
        };

        etHistory.characters = 4;
        etHistory.helpTip = L("historyHelpTip");

        var cbLiveEdit = panelPerf.add("checkbox", undefined, L("cbLiveEdit"));
        cbLiveEdit.helpTip = L("cbLiveEdit");


        /*
      ［ファイル管理］パネル（右） / File Management panel (right)
    */
        var panelFile = rightColumnGroup.add("panel", undefined, L("panelFileTitle"));
        panelFile.orientation = "column";
        panelFile.alignChildren = ["fill", "top"];
        panelFile.alignment = ["fill", "top"];
        panelFile.margins = PANEL_MARGINS;

        var cbEditOriginal = panelFile.add("checkbox", undefined, L("cbEditOriginal"));
        cbEditOriginal.helpTip = L("cbEditOriginal");

        var cbFontsAuto = panelFile.add("checkbox", undefined, L("cbFontsAuto"));
        cbFontsAuto.helpTip = L("cbFontsAuto");

        /* ［ファイル管理］内の保存先 / Save location inside [File Management] */
        var saveLocationRow = panelFile.add("group", undefined, "");
        saveLocationRow.orientation = "row";
        saveLocationRow.alignChildren = ["left", "center"];
        saveLocationRow.spacing = 10;

        saveLocationRow.add("statictext", undefined, L("saveDest") + "：");
        var rbSaveLocal = saveLocationRow.add("radiobutton", undefined, L("saveLocal"));
        var rbSaveCloud = saveLocationRow.add("radiobutton", undefined, L("saveCloud"));

        // 初期状態（true=クラウド、false=コンピューター）/ Initial state
        var prefCloud = false;
        try {
            prefCloud = app.preferences.getBooleanPreference("AdobeSaveAsCloudDocumentPreference");
        } catch (e) {
            prefCloud = false;
        }
        rbSaveCloud.value = prefCloud;
        rbSaveLocal.value = !prefCloud;

        // UI only; persisted on OK / ここではUIのみ更新し、保存はOK時に実行
        rbSaveLocal.onClick = function () {
            // UI only; persist on OK
        };
        rbSaveCloud.onClick = function () {
            // UI only; persist on OK
        };

        /* ［ファイル管理］内の「リンクを更新」 / Update Links inside [File Management] */
        var updateLinksRow = panelFile.add("group", undefined, "");
        updateLinksRow.orientation = "row";
        updateLinksRow.alignChildren = ["left", "center"];
        updateLinksRow.spacing = 10;

        updateLinksRow.add("statictext", undefined, L("updateLinks") + "：");
        var rbUpdateLinksAuto = updateLinksRow.add("radiobutton", undefined, L("updateLinksAuto"));
        var rbUpdateLinksManual = updateLinksRow.add("radiobutton", undefined, L("updateLinksManual"));
        var rbUpdateLinksAsk = updateLinksRow.add("radiobutton", undefined, L("updateLinksAsk"));

        // 初期状態（0=自動、1=手動、2=確認）/ Initial state
        var curUpdateLinks = safeGetInt("plugin/FileClipboard/linkoptions", 2);
        rbUpdateLinksAuto.value = (curUpdateLinks === 0);
        rbUpdateLinksManual.value = (curUpdateLinks === 1);
        rbUpdateLinksAsk.value = (curUpdateLinks !== 0 && curUpdateLinks !== 1);

        /*
          ［クリップボードの処理］パネル（右） / Clipboard Handling panel (right)
        */
        var panelClipboard = rightColumnGroup.add("panel", undefined, L("panelClipboardTitle"));
        panelClipboard.orientation = "column";
        panelClipboard.alignChildren = ["fill", "top"];
        panelClipboard.alignment = ["fill", "top"];
        panelClipboard.margins = PANEL_MARGINS;

        var cbIncludeSVG = panelClipboard.add("checkbox", undefined, L("cbIncludeSVG"));
        cbIncludeSVG.helpTip = L("cbIncludeSVG");

        /*
      「パスに制限」パネル（右） / Limit to Path panel (right)
    */
        var panelLimitPath = rightColumnGroup.add("panel", undefined, L("panelLimitPathTitle"));
        panelLimitPath.orientation = "column";
        panelLimitPath.alignChildren = ["fill", "top"];
        panelLimitPath.alignment = ["fill", "top"];
        panelLimitPath.margins = PANEL_MARGINS;

        var cbHitShape = panelLimitPath.add("checkbox", undefined, L("cbHitShape"));
        cbHitShape.helpTip = L("cbHitShape");
        /*
          Load current values into the UI / 現在の設定値をUIへ反映
        */
        var hitShapePrefValue = safeGetInt("hitShapeOnPreview", 1);
        cbHitShape.value = (hitShapePrefValue === 0); // 0がON（true）、1がOFF（false）

        var cbHitTypeShape = panelLimitPath.add("checkbox", undefined, L("cbHitTypeShape"));
        cbHitTypeShape.helpTip = L("cbHitTypeShape");
        // hitTypeShapeOnPreview: 0がON（true）、1がOFF（false）
        var hitTypeShapePrefValue = safeGetInt("hitTypeShapeOnPreview", 1);
        cbHitTypeShape.value = (hitTypeShapePrefValue === 0); // 0がON（true）、1がOFF（false）

        // ===== Simple Preference-to-UI Bindings =====
        // 単純なプリファレンスとUIの対応定義 / Simple preference-to-UI bindings
        var PREFERENCE_UI_BINDINGS = [{
            key: "showRichToolTips",
            type: "bool",
            ui: cbToolTips
        },
        {
            key: "Hello/ShowHomeScreenWS",
            type: "bool",
            ui: cbHomeScreen
        },
        {
            key: "Hello/NewDoc",
            type: "bool",
            ui: cbLegacyNewDoc
        },
        {
            key: "enablePrintBleedWidget",
            type: "bool",
            ui: cbBleedAI
        },
        {
            key: "moveLockedAndHiddenArt",
            type: "bool",
            ui: cbMoveLocked
        },
        {
            key: "hitShapeOnPreview",
            type: "int01inv",
            ui: cbHitShape
        }, // 0=>ON(true),1=>OFF(false)
        {
            key: "zoomToSelection",
            type: "bool",
            ui: cbZoomToSel
        },
        {
            key: "hitTypeShapeOnPreview",
            type: "int01inv",
            ui: cbHitTypeShape
        },
        {
            key: "smartGuides/showObjectHighlighting",
            type: "bool",
            ui: cbSmartHighlight
        },
        {
            key: "text/autoSizing",
            type: "bool",
            ui: cbAutoSizing
        },
        {
            key: "text/doFontLocking",
            type: "bool",
            ui: cbFontLocking
        },
        {
            key: "text/enableAlternateGlyph",
            type: "bool",
            ui: cbAlternateGlyph
        },
        {
            key: "Performance/AnimZoom",
            type: "bool",
            ui: cbAnimZoom
        },
        {
            key: "LiveEdit_State_Machine",
            type: "bool",
            ui: cbLiveEdit
        },
        {
            key: "useSysDefEdit",
            type: "bool",
            ui: cbEditOriginal
        },
        {
            key: "AutoActivateMissingFont",
            type: "bool",
            ui: cbFontsAuto
        },
        {
            key: "plugin/FileClipboard/copySVGCode",
            type: "bool",
            ui: cbIncludeSVG
        },
        {
            key: "showArtboardLabelOnCanvas",
            type: "bool",
            ui: cbShowArtboardName
        },
            // (Removed composite/special-case bindings)
        ];

        // システムの現在値 → UIへ反映
        function loadPrefsFromSystem() {
            for (var i = 0; i < PREFERENCE_UI_BINDINGS.length; i++) {
                var p = PREFERENCE_UI_BINDINGS[i];

                if (p.type === "bool") {
                    var b = safeGetBool(p.key, false);
                    if (p.ui) p.ui.value = !!b;

                } else if (p.type === "int01inv") {
                    var v01 = safeGetInt(p.key, 1);
                    if (p.ui) p.ui.value = (v01 === 0); // 0=>ON(true)
                }
            }
            // 複合UI項目は個別に反映 / Load composite UI fields separately
            etHistory.text = String(validateInt(safeGetInt("maximumUndoDepth", 0), VALIDATION.history));

            var recentFontsValue = validateInt(safeGetInt("text/recentFontMenu/showNEntries", 0), VALIDATION.recentFonts);
            etRecentFonts.text = String(recentFontsValue);
            cbRecentFonts.value = (recentFontsValue > 0);
            etRecentFonts.enabled = (recentFontsValue > 0);

            var canvasPreferenceValue = safeGetInt("uiCanvasIsWhite", 0);
            rbCanvasWhite.value = (canvasPreferenceValue === 1);
            rbCanvasMatch.value = (canvasPreferenceValue !== 1);
            // ストロークカラー・幅は Real 値なので bindings 外で個別読み込み
            var scR = safeGetReal("ArtboardBBColorRed", 0.0);
            var scG = safeGetReal("ArtboardBBColorGreen", 0.0);
            var scB = safeGetReal("ArtboardBBColorBlue", 0.0);
            ddStrokeColor.selection = findClosestStrokeColor(scR, scG, scB);
            var swVal = safeGetReal("ArtboardBBWidth", 1.0);
            var swLoadIdx = clamp(Math.round(swVal), 1, 4) - 1;
            for (var swi = 0; swi < rbStrokeWidths.length; swi++) rbStrokeWidths[swi].value = (swi === swLoadIdx);
        }
        // ===== End Simple Preference-to-UI Bindings =====

        // ===== Validation Rules =====
        var VALIDATION = {
            recentFonts: {
                min: 0,
                max: 30,
                def: 15
            }, // 0 hides the list / 0で非表示
            history: {
                min: 1,
                max: 1000,
                def: 100
            } // Expected range: 1-1000 / 想定範囲: 1〜1000
        };

        function toInt(val) {
            var n = parseInt(val, 10);
            return isNaN(n) ? null : n;
        }

        function clamp(n, min, max) {
            return Math.max(min, Math.min(max, n));
        }

        function validateInt(val, rule) {
            var n = toInt(val);
            if (n === null) return rule.def;
            return clamp(n, rule.min, rule.max);
        }
        // Validate EditText input; out-of-range or non-numeric values fall back to default / EditText入力を検証し、範囲外や非数はデフォルトへ補正
        function validateEditTextInt(editText, rule) {
            var v = validateInt(editText.text, rule);
            editText.text = String(v);
            return v;
        }
        // ===== End Validation Rules =====

        // ===== Safe Preference Accessors =====
        function safeSetInt(key, val) {
            try {
                app.preferences.setIntegerPreference(key, val);
            } catch (e) {
                $.writeln("[PresetManager] safeSetInt failed: " + key + " = " + val + " / " + e);
            }
        }

        function safeSetBool(key, val) {
            try {
                app.preferences.setBooleanPreference(key, val);
            } catch (e) {
                $.writeln("[PresetManager] safeSetBool failed: " + key + " = " + val + " / " + e);
            }
        }

        function safeSetReal(key, val) {
            try {
                app.preferences.setRealPreference(key, val);
            } catch (e) {
                $.writeln("[PresetManager] safeSetReal failed: " + key + " = " + val + " / " + e);
            }
        }

        function safeGetInt(key, fb) {
            try {
                return app.preferences.getIntegerPreference(key);
            } catch (e) {
                $.writeln("[PresetManager] safeGetInt fallback: " + key + " / " + e);
                return fb;
            }
        }

        function safeGetBool(key, fb) {
            try {
                return app.preferences.getBooleanPreference(key);
            } catch (e) {
                $.writeln("[PresetManager] safeGetBool fallback: " + key + " / " + e);
                return fb;
            }
        }

        function safeGetReal(key, fb) {
            try {
                return app.preferences.getRealPreference(key);
            } catch (e) {
                $.writeln("[PresetManager] safeGetReal fallback: " + key + " / " + e);
                return fb;
            }
        }
        // ===== End Safe Preference Accessors =====

        // ===== Math Utilities =====
        /**
         * Compare with tolerance / 許容誤差つき比較
         * @param {Number} a
         * @param {Number} b
         * @param {Number} [eps=0.02]
         * @returns {Boolean}
         */
        function approx(a, b, eps) {
            if (typeof eps !== 'number') eps = 0.02;
            return Math.abs(a - b) < eps;
        }
        // ===== End Math Utilities =====

        // ===== Preset Applier =====
        // Shared helper to apply preset values to the UI / プリセット値をUIへ反映する共通ヘルパー
        function applyPresetToUI(uiState) {
            // 単純なチェックボックス・数値はそのまま反映
            if (typeof uiState.cbToolTips !== 'undefined') cbToolTips.value = !!uiState.cbToolTips;
            if (typeof uiState.cbHomeScreen !== 'undefined') cbHomeScreen.value = !!uiState.cbHomeScreen;
            if (typeof uiState.cbLegacyNewDoc !== 'undefined') cbLegacyNewDoc.value = !!uiState.cbLegacyNewDoc;
            if (typeof uiState.cbBleedAI !== 'undefined') cbBleedAI.value = !!uiState.cbBleedAI;
            if (typeof uiState.cbMoveLocked !== 'undefined') cbMoveLocked.value = !!uiState.cbMoveLocked;
            if (typeof uiState.cbHitShape !== 'undefined') cbHitShape.value = !!uiState.cbHitShape; // true=ON(path only)
            if (typeof uiState.cbZoomToSel !== 'undefined') cbZoomToSel.value = !!uiState.cbZoomToSel;
            if (typeof uiState.cbHitTypeShape !== 'undefined') cbHitTypeShape.value = !!uiState.cbHitTypeShape; // true=ON(path only)
            if (typeof uiState.cbAutoSizing !== 'undefined') cbAutoSizing.value = !!uiState.cbAutoSizing;
            if (typeof uiState.cbRecentFonts !== 'undefined') cbRecentFonts.value = !!uiState.cbRecentFonts;
            if (typeof uiState.etRecentFonts !== 'undefined') {
                etRecentFonts.text = String(uiState.etRecentFonts);
                etRecentFonts.enabled = !!uiState.cbRecentFonts;
            }
            if (typeof uiState.cbFontLocking !== 'undefined') cbFontLocking.value = !!uiState.cbFontLocking;
            if (typeof uiState.cbAlternateGlyph !== 'undefined') cbAlternateGlyph.value = !!uiState.cbAlternateGlyph;
            if (typeof uiState.cbSmartHighlight !== 'undefined') cbSmartHighlight.value = !!uiState.cbSmartHighlight;
            if (typeof uiState.cbAnimZoom !== 'undefined') cbAnimZoom.value = !!uiState.cbAnimZoom;
            if (typeof uiState.etHistory !== 'undefined') etHistory.text = String(uiState.etHistory);
            if (typeof uiState.cbLiveEdit !== 'undefined') cbLiveEdit.value = !!uiState.cbLiveEdit;
            if (typeof uiState.cbEditOriginal !== 'undefined') cbEditOriginal.value = !!uiState.cbEditOriginal;
            if (typeof uiState.cbFontsAuto !== 'undefined') cbFontsAuto.value = !!uiState.cbFontsAuto;
            if (typeof uiState.cbIncludeSVG !== 'undefined') cbIncludeSVG.value = !!uiState.cbIncludeSVG;
            if (typeof uiState.cbShowArtboardName !== 'undefined') cbShowArtboardName.value = !!uiState.cbShowArtboardName;
            if (typeof uiState.artboardStrokeColor !== 'undefined') ddStrokeColor.selection = uiState.artboardStrokeColor;
            if (typeof uiState.artboardStrokeWidth !== 'undefined') {
                var swpIdx = clamp(uiState.artboardStrokeWidth, 1, 4) - 1;
                for (var swj = 0; swj < rbStrokeWidths.length; swj++) rbStrokeWidths[swj].value = (swj === swpIdx);
            }

            // カンバス（ラジオ）
            if (typeof uiState.canvasWhite !== 'undefined') {
                rbCanvasWhite.value = !!uiState.canvasWhite;
                rbCanvasMatch.value = !rbCanvasWhite.value;
            }

            // ガイド色（UIのみ）
            if (uiState.guideColor === 'lightblue') {
                if (typeof rbGuideLightBlue !== 'undefined' && typeof rbGuideCyan !== 'undefined') {
                    rbGuideLightBlue.value = true;
                    rbGuideCyan.value = false;
                }
            } else if (uiState.guideColor === 'cyan') {
                if (typeof rbGuideLightBlue !== 'undefined' && typeof rbGuideCyan !== 'undefined') {
                    rbGuideCyan.value = true;
                    rbGuideLightBlue.value = false;
                }
            }

            // 保存先（クラウド/ローカル）（UIのみ）
            if (typeof uiState.saveCloud !== 'undefined') {
                rbSaveCloud.value = !!uiState.saveCloud;
                rbSaveLocal.value = !rbSaveCloud.value;
            }

            // リンクを更新（UIのみ）
            if (typeof uiState.updateLinks !== 'undefined') {
                rbUpdateLinksAuto.value = (uiState.updateLinks === 0);
                rbUpdateLinksManual.value = (uiState.updateLinks === 1);
                rbUpdateLinksAsk.value = (uiState.updateLinks !== 0 && uiState.updateLinks !== 1);
            }
        }

        // Preset definitions / プリセット定義
        var PRESET1_UI_STATE = {
            cbToolTips: false,
            cbHomeScreen: false,
            cbLegacyNewDoc: true,
            cbBleedAI: false,
            cbMoveLocked: true,
            cbHitShape: false,
            cbZoomToSel: false,
            cbHitTypeShape: false,
            cbAutoSizing: true,
            cbRecentFonts: true,
            etRecentFonts: 15,
            cbFontLocking: false,
            cbAlternateGlyph: false,
            cbAnimZoom: false,
            etHistory: 50,
            cbLiveEdit: false,
            cbEditOriginal: true,
            cbFontsAuto: true,
            canvasWhite: true,
            guideColor: 'lightblue',
            saveCloud: false,
            cbSmartHighlight: false,
            cbIncludeSVG: true,
            cbShowArtboardName: false,
            artboardStrokeColor: 7,
            artboardStrokeWidth: 2,
            updateLinks: 0
        };

        var DEFAULT_UI_STATE = {
            cbToolTips: true,
            cbHomeScreen: true,
            cbLegacyNewDoc: false,
            cbBleedAI: true,
            cbMoveLocked: false,
            cbHitShape: false,
            cbZoomToSel: true,
            cbHitTypeShape: false,
            cbAutoSizing: false,
            cbRecentFonts: true,
            etRecentFonts: 10,
            cbFontLocking: true,
            cbAlternateGlyph: true,
            cbAnimZoom: true,
            etHistory: 100,
            cbLiveEdit: true,
            cbEditOriginal: false,
            cbFontsAuto: false,
            canvasWhite: false,
            guideColor: 'cyan',
            saveCloud: true,
            cbSmartHighlight: true,
            cbIncludeSVG: false,
            cbShowArtboardName: true,
            artboardStrokeColor: 0,
            artboardStrokeWidth: 1,
            updateLinks: 2
        };
        // ===== End Preset Applier =====

        /* 下部ボタン行（Cancel / OK） / Bottom button row (Cancel / OK) */
        var bottomButtonRow = mainGroup.add("group"); // 親groupに追加
        bottomButtonRow.orientation = "row";
        bottomButtonRow.alignChildren = ["fill", "center"];
        bottomButtonRow.alignment = ["fill", "bottom"];


        // 伸縮するスペーサー / Flexible spacer
        var flexibleSpacer = bottomButtonRow.add("group");
        flexibleSpacer.alignment = ["fill", "fill"];

        // 右側ボタングループ（Cancel / OK） / Right button group (Cancel / OK)
        var buttonGroup = bottomButtonRow.add("group");
        buttonGroup.orientation = "row";
        buttonGroup.alignChildren = ["right", "center"];
        buttonGroup.spacing = 10;
        var btnCancel = buttonGroup.add("button", undefined, L("Cancel"), {
            name: "cancel"
        });
        var btnOK = buttonGroup.add("button", undefined, L("OK"), {
            name: "ok"
        });

        btnOK.onClick = function () {
            try {
                /*
                  Save "Rich Tool Tips" setting / 「詳細なツールヒント」設定を保存
                */
                safeSetBool("showRichToolTips", cbToolTips.value);
                /*
                  Save "Home Screen" setting / 「ホーム画面」設定を保存
                */
                safeSetBool("Hello/ShowHomeScreenWS", cbHomeScreen.value);
                /*
                  Save "Legacy New Document" setting / 「以前の新規ドキュメント」設定を保存
                */
                safeSetBool("Hello/NewDoc", cbLegacyNewDoc.value);
                /*
                  Save "Print Bleed" generative AI button setting / 「裁ち落としを印刷」生成AIボタン設定を保存
                */
                safeSetBool("enablePrintBleedWidget", cbBleedAI.value);
                /*
                  Save "Move with Artboard" setting / 「アートボードと一緒に移動」設定を保存
                */
                safeSetBool("moveLockedAndHiddenArt", cbMoveLocked.value);
                /*
                  Save "Object Selection by Path Only" setting / 「オブジェクトの選択範囲をパスに制限」設定を保存
                */
                safeSetInt("hitShapeOnPreview", cbHitShape.value ? 0 : 1);
                /*
                  Save "Zoom to Selection" setting / 「選択範囲へズーム」設定を保存
                */
                safeSetBool("zoomToSelection", cbZoomToSel.value);
                /*
                  Save "Type Object Selection by Path Only" setting / 「テキストをパスに制限」設定を保存
                */
                safeSetInt("hitTypeShapeOnPreview", cbHitTypeShape.value ? 0 : 1);
                /*
                  Save "Auto Size New Area Type" setting / 「新規エリア内文字の自動サイズ調整」設定を保存
                */
                safeSetBool("text/autoSizing", cbAutoSizing.value);
                /*
                  Save "Recent Fonts" setting / 「最近使用したフォント」設定を保存
                */
                if (cbRecentFonts.value) {
                    safeSetInt("text/recentFontMenu/showNEntries", validateEditTextInt(etRecentFonts, VALIDATION.recentFonts));
                } else {
                    safeSetInt("text/recentFontMenu/showNEntries", 0);
                }
                /*
                  Save "Missing Glyph Protection" setting / 「見つからない字形の保護」設定を保存
                */
                safeSetBool("text/doFontLocking", cbFontLocking.value);
                /*
                  Save "Character Alternates" setting / 「選択された文字の異体字」設定を保存
                */
                safeSetBool("text/enableAlternateGlyph", cbAlternateGlyph.value);
                /*
                  Save canvas color setting from radio selection / ラジオ選択に基づいてカンバスカラー設定を保存
                */
                safeSetInt("uiCanvasIsWhite", rbCanvasWhite.value ? 1 : 0);
                /*
                  Save guide color from radio selection / ラジオ選択に基づいてガイドカラーを保存
                */
                if (rbGuideLightBlue && rbGuideLightBlue.value) {
                    safeSetReal("Guide/Color/red", 0.29);
                    safeSetReal("Guide/Color/green", 0.52);
                    safeSetReal("Guide/Color/blue", 1.0);
                } else {
                    // Default (Cyan) / 既定（シアン）
                    safeSetReal("Guide/Color/red", 0.0);
                    safeSetReal("Guide/Color/green", 1.0);
                    safeSetReal("Guide/Color/blue", 1.0);
                }
                /*
                  Save guide style (0: Lines / 1: Dots) / ガイドスタイルを保存（0:ライン / 1:点線）
                */
                if (rbGuideStyleLine && rbGuideStyleLine.value) {
                    safeSetInt("Guide/Style", 0);
                } else if (rbGuideStyleDots && rbGuideStyleDots.value) {
                    safeSetInt("Guide/Style", 1);
                }
                /*
                  Save "Object Highlighting" smart guide setting / 「オブジェクトのハイライト表示（スマートガイド）」設定を保存
                */
                safeSetBool("smartGuides/showObjectHighlighting", cbSmartHighlight.value);
                /*
                  Save "Animated Zoom" setting / 「アニメーションズーム」設定を保存
                */
                safeSetBool("Performance/AnimZoom", cbAnimZoom.value);
                /*
                  Save "History States" setting / 「ヒストリー数」設定を保存
                */
                safeSetInt("maximumUndoDepth", validateEditTextInt(etHistory, VALIDATION.history));
                /*
                  Save "Real-Time Drawing and Editing" setting / 「リアルタイムの描画と編集」設定を保存
                */
                safeSetBool("LiveEdit_State_Machine", cbLiveEdit.value);
                /*
                  Save "Use System Defaults for Edit Original" setting / 「オリジナルの編集にシステムデフォルトを使用」設定を保存
                */
                safeSetBool("useSysDefEdit", cbEditOriginal.value);
                /*
                  Save "Auto-activate Adobe Fonts" setting / 「Adobe Fontsを自動アクティベート」設定を保存
                */
                safeSetBool("AutoActivateMissingFont", cbFontsAuto.value);
                /*
                  Save "Include SVG Code" setting / 「SVGコードを含める」設定を保存
                */
                safeSetBool("plugin/FileClipboard/copySVGCode", cbIncludeSVG.value);
                /*
                  Save "Show Artboard Name" setting / 「アートボード名を表示」設定を保存
                */
                safeSetBool("showArtboardLabelOnCanvas", cbShowArtboardName.value);
                /*
                  Save artboard stroke color / アートボードのストロークカラーを保存
                */
                var scIdx = ddStrokeColor.selection ? ddStrokeColor.selection.index : 0;
                var scPreset = STROKE_COLOR_PRESETS[scIdx];
                safeSetReal("ArtboardBBColorRed", scPreset.r);
                safeSetReal("ArtboardBBColorGreen", scPreset.g);
                safeSetReal("ArtboardBBColorBlue", scPreset.b);
                /*
                  Save artboard stroke width / アートボードのストローク幅を保存
                */
                safeSetReal("ArtboardBBWidth", getSelectedStrokeWidth());
                /*
                  Save file save location explicitly on OK / OK押下時にファイル保存先を明示保存
                */
                safeSetBool("AdobeSaveAsCloudDocumentPreference", !!rbSaveCloud.value);
                /*
                  Save "Update Links" setting / 「リンクを更新」設定を保存
                */
                var updateLinksValue = 2;
                if (rbUpdateLinksAuto.value) {
                    updateLinksValue = 0;
                } else if (rbUpdateLinksManual.value) {
                    updateLinksValue = 1;
                }
                safeSetInt("plugin/FileClipboard/linkoptions", updateLinksValue);

                // Force screen refresh workaround: redraw() alone does not reliably update UI after preference changes
                // 画面更新のためのハック：redraw() だけでは環境設定変更後にUIが確実に更新されないため、ズーム操作で強制再描画
                app.executeMenuCommand('zoomout');
                app.executeMenuCommand('zoomin');
                dlg.close();
            } catch (e) {
                alert(L("alertSavePrefsFailed") + e);
            }
        };
        /*
          Load current preferences as the initial UI state / 初期状態として現在の環境設定を読み込む
        */
        applyPresetToUIFromId(PRESET_IDS.current);
        /* Adjust dialog opacity & position / ダイアログの不透明度と位置 */
        var offsetX = 300;
        var dialogOpacity = 0.98;

        function shiftDialogPosition(dlg, offsetX, offsetY) {
            dlg.onShow = function () {
                var currentX = dlg.location[0];
                var currentY = dlg.location[1];
                dlg.location = [currentX + offsetX, currentY + offsetY];
            };
        }

        function setDialogOpacity(dlg, opacityValue) {
            dlg.opacity = opacityValue;
        }

        setDialogOpacity(dlg, dialogOpacity);
        shiftDialogPosition(dlg, offsetX, 0);
        /* Show dialog / ダイアログ表示 */
        dlg.show();
    }

    main();
})();