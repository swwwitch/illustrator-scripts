#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### スクリプト名：

Illustrator環境設定ユーティリティ

### GitHub：

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/Preference/PresetManager.jsx

### 概要：

- Illustrator の各種環境設定を GUI から一括で確認・変更します。
- プリセット（無指定／デフォルト／プリセット1）を選ぶだけで推奨セットを適用できます。
- カンバスカラー、選択範囲・アンカー表示、テキスト、パフォーマンス、ファイル管理（保存先を含む）をカバーします。

### 主な機能：

- 詳細なツールヒント／ホーム画面／レガシー新規ドキュメント UI の切替
- アートボード移動時のロック/非表示オブジェクト追従、ズーム先、アンカー強調表示
- テキスト：エリア内文字の自動サイズ、最近使用フォントの件数、異体字・見つからない字形の保護
- ユーザーインターフェイス：カンバスカラー（UIに合わせる／ホワイト）
- パフォーマンス：アニメーションズーム、ヒストリー数、リアルタイム描画
- ファイル管理：編集の既定アプリ、Adobe Fonts 自動アクティベート、ファイルの保存先（コンピューター／クラウド）

### 処理の流れ：

1) ダイアログ生成 → 2) 現在の環境設定を読み込み UI に反映 → 3) プリセット選択で UI を上書き → 4) ［OK］で各プリファレンスキーへ保存。

### 更新履歴：

- v1.0 (20250807) : 初期リリース
- v1.1 (20250915) : ファイルの保存先を追加
- v1.2（20250910）：ガイドのカラー、スタイルを追加
- v1.3（20250911）：微調整
- v1.4（20250912）：クリップボード（SVGコードを含める）UIとロジックを削除

---

### Script Name:

Illustrator Preferences Utility

### GitHub:

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/Preference/PresetManager.jsx

### Overview:

- View and change multiple Illustrator preferences at once via a single GUI.
- Apply recommended presets (None/Default/Preset 1) with one selection.
- Covers Canvas Color, Selection & Anchor Display, Type, Performance, and File Management (including Save Location).

### Key Features:

- Rich Tool Tips / Home Screen / Legacy New Document toggle
- Move Locked/Hidden with Artboard, Zoom to Selection, Highlight Anchors on Hover
- Type: Auto Size Area Type, Recent Fonts count, Alternate Glyphs & Missing Glyph Protection
- UI: Canvas Color (Match UI / White)
- Performance: Animated Zoom, History States, Real-Time Drawing
- File Management: Edit Original default app, Auto-activate Adobe Fonts, Save Location (Computer/Cloud)

### Flow:

1) Build dialog → 2) Load current preferences into UI → 3) Apply preset to UI → 4) Save preferences when pressing OK.

### Changelog:

- v1.0 (2025-08-07): Initial release
- v1.1 (2025-09-15): Added Save Location
- v1.2（20250910）：Added Guide Color
- v1.3 (20250911) : Minor Changes
- v1.4 (2025-09-12): Removed Clipboard "Include SVG code" UI and logic

*/

var SCRIPT_VERSION = "v1.4";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */

var LABELS = {
    dialogTitle: {
        ja: "環境設定をまとめて変更 " + SCRIPT_VERSION,
        en: "Batch Preferences Dialog " + SCRIPT_VERSION
    },
    OK: {
        ja: "OK",
        en: "OK"
    },
    Cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },

    // ［一般］
    cbToolTips: {
        ja: "詳細なツールヒントを表示",
        en: "Show Rich Tool Tips"
    },
    cbHomeScreen: {
        ja: "ドキュメントを開いていないときに「ホーム画面」を表示",
        en: "Show The Home Screen When No Documents Are Open"
    },
    cbLegacyNewDoc: {
        ja: "以前の「新規ドキュメント」インターフェイスを使用",
        en: "Use legacy “File New” interface"
    },
    cbBleedAI: {
        ja: "「裁ち落としを印刷」生成AIボタンを表示",
        en: "Show 'Print Bleed' generative AI buttons on Bleed"
    },

    // ［選択範囲・アンカー表示］
    cbMoveLocked: {
        ja: "ロックまたは非表示オブジェクトをアートボードと一緒に移動",
        en: "Move Locked and Hidden Artwork with Artboard"
    },
    cbHitShape: {
        ja: "オブジェクトの選択範囲をパスに制限",
        en: "Object Selection by Path Only"
    },
    cbZoomToSel: {
        ja: "選択範囲へズーム",
        en: "Zoom to Selection"
    },
    cbHighlightAnchor: {
        ja: "カーソルを合わせたときにアンカーを強調表示",
        en: "Highlight anchors on mouse over"
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
    }
    ,
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
    panelLimitPathTitle: {
        ja: "パスに制限",
        en: "Limit to Path"
    }
    ,
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
    }
};

function main() {

    /*
  Build main dialog / ダイアログ生成
*/
    var dlg = new Window("dialog", LABELS.dialogTitle[lang]);
    /*
  Main vertical container / メイングループ
*/
    var mainGroup = dlg.add("group");
    mainGroup.orientation = "column";
    mainGroup.alignChildren = "left";

    /*
  Preset selector (Current/Default/Preset1) / プリセット選択
*/
    var groupPreset = mainGroup.add("group");
    groupPreset.name = "groupPreset";
    groupPreset.alignment = "center";
    groupPreset.orientation = "row";
    groupPreset.margins = [0, 10, 20, 20];
    var groupPresetInner = groupPreset.add("group");
    groupPresetInner.orientation = "row";
    groupPresetInner.alignChildren = "center";
    groupPresetInner.add("statictext", undefined, "プリセット：");
    var presetItems = ["現在の設定", "デフォルト", "プリセット1"];
    var ddPreset = groupPresetInner.add("dropdownlist", undefined, presetItems);
    ddPreset.selection = 0; // 初期選択を「無指定」に

    /*
    Apply preset settings to the UI / プリセット適用
    - "無指定" は現在の環境設定を読んでUIへ反映
    - "デフォルト" / "プリセット1" は定義オブジェクトから反映
  */
    function applyPresetSettings(presetName) {
        if (presetName === "現在の設定") {
            /*
              現在の環境設定を UI に反映する
            */
            loadPrefsFromSystem();

            /*
              カンバスカラー
            */
            var currentCanvas = app.preferences.getIntegerPreference("uiCanvasIsWhite");
            rbCanvasWhite.value = (currentCanvas === 1);
            rbCanvasMatch.value = (currentCanvas !== 1);
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

        } else if (presetName === "プリセット1") {
            applyPresetToUI(PRESET1_DEF);


        } else if (presetName === "デフォルト") {
            applyPresetToUI(DEFAULT_DEF);
        }
    }

    ddPreset.onChange = function() {
        if (ddPreset.selection && ddPreset.selection.text) {
            applyPresetSettings(ddPreset.selection.text);
            app.redraw();
        }
    };


    /*
  Two-column layout / 2カラム
*/
    var colGroup = mainGroup.add("group");
    colGroup.orientation = "row";
    colGroup.alignChildren = "top";

    var colLeft = colGroup.add("group");
    colLeft.orientation = "column";
    colLeft.alignChildren = ["fill", "top"];

    var colRight = colGroup.add("group");
    colRight.orientation = "column";
    colRight.alignChildren = ["fill", "top"];

    /*
  General panel (left) / ［一般］
*/
    var panelGeneral = colLeft.add("panel", undefined, LABELS.panelGeneralTitle[lang]);
    panelGeneral.orientation = "column";
    panelGeneral.alignChildren = ["fill", "top"];
    panelGeneral.alignment = ["fill", "top"];
    panelGeneral.margins = [15, 25, 15, 10]; // 上、右、下、左のマージン

    var cbToolTips = panelGeneral.add("checkbox", undefined, LABELS.cbToolTips[lang]);
    cbToolTips.helpTip = LABELS.cbToolTips.ja + " / " + LABELS.cbToolTips.en;

    var cbHomeScreen = panelGeneral.add("checkbox", undefined, LABELS.cbHomeScreen[lang]);
    cbHomeScreen.helpTip = LABELS.cbHomeScreen.ja + " / " + LABELS.cbHomeScreen.en;

    var cbLegacyNewDoc = panelGeneral.add("checkbox", undefined, LABELS.cbLegacyNewDoc[lang]);
    cbLegacyNewDoc.helpTip = LABELS.cbLegacyNewDoc.ja + " / " + LABELS.cbLegacyNewDoc.en;

    var cbBleedAI = panelGeneral.add("checkbox", undefined, LABELS.cbBleedAI[lang]);
    cbBleedAI.helpTip = LABELS.cbBleedAI.ja + " / " + LABELS.cbBleedAI.en;

    /*
  Selection & Anchor panel (left) / ［選択範囲・アンカー表示］
*/
    var panelSelectAnchor = colLeft.add("panel", undefined, LABELS.panelSelectAnchorTitle[lang]);
    panelSelectAnchor.orientation = "column";
    panelSelectAnchor.alignChildren = ["fill", "top"];
    panelSelectAnchor.alignment = ["fill", "top"];
    panelSelectAnchor.margins = [15, 25, 15, 10];

    var cbMoveLocked = panelSelectAnchor.add("checkbox", undefined, LABELS.cbMoveLocked[lang]);
    cbMoveLocked.helpTip = LABELS.cbMoveLocked.ja + " / " + LABELS.cbMoveLocked.en;

    var cbZoomToSel = panelSelectAnchor.add("checkbox", undefined, LABELS.cbZoomToSel[lang]);
    cbZoomToSel.helpTip = LABELS.cbZoomToSel.ja + " / " + LABELS.cbZoomToSel.en;

    var cbHighlightAnchor = panelSelectAnchor.add("checkbox", undefined, LABELS.cbHighlightAnchor[lang]);
    cbHighlightAnchor.helpTip = LABELS.cbHighlightAnchor.ja + " / " + LABELS.cbHighlightAnchor.en;

    /*
  Text panel (left) / ［テキスト］
*/
    var panelTextRight = colLeft.add("panel", undefined, LABELS.panelTextTitle[lang]);
    panelTextRight.orientation = "column";
    panelTextRight.alignChildren = ["fill", "top"];
    panelTextRight.alignment = ["fill", "top"];
    panelTextRight.margins = [15, 25, 15, 10];

    var cbAutoSizing = panelTextRight.add("checkbox", undefined, LABELS.cbAutoSizing[lang]);
    cbAutoSizing.helpTip = LABELS.cbAutoSizing.ja + " / " + LABELS.cbAutoSizing.en;


    var currentRecentCount = app.preferences.getIntegerPreference("text/recentFontMenu/showNEntries");
    var groupRecentFonts = panelTextRight.add("group");
    groupRecentFonts.orientation = "row";

    var cbRecentFonts = groupRecentFonts.add("checkbox", undefined, LABELS.cbRecentFonts[lang]);
    cbRecentFonts.helpTip = LABELS.cbRecentFonts.ja + " / " + LABELS.cbRecentFonts.en;
    cbRecentFonts.value = (currentRecentCount > 0);

    var etRecentFonts = groupRecentFonts.add("edittext", undefined, currentRecentCount.toString());
    etRecentFonts.characters = 3;
    etRecentFonts.enabled = cbRecentFonts.value;

    cbRecentFonts.onClick = function() {
        // UI only: enable/disable and default value. Persist on OK.
        if (cbRecentFonts.value) {
            etRecentFonts.enabled = true;
            validateEditTextInt(etRecentFonts, VALIDATION.recentFonts);
        } else {
            etRecentFonts.enabled = false;
            etRecentFonts.text = "0"; // OFF時は 0 固定（保存はOK時）
        }
    };

    etRecentFonts.onChange = function() {
        // UI only validation. Persist on OK.
        validateEditTextInt(etRecentFonts, VALIDATION.recentFonts);
    };

    var cbFontLocking = panelTextRight.add("checkbox", undefined, LABELS.cbFontLocking[lang]);
    cbFontLocking.helpTip = LABELS.cbFontLocking.ja + " / " + LABELS.cbFontLocking.en;

    var cbAlternateGlyph = panelTextRight.add("checkbox", undefined, LABELS.cbAlternateGlyph[lang]);
    cbAlternateGlyph.helpTip = LABELS.cbAlternateGlyph.ja + " / " + LABELS.cbAlternateGlyph.en;


    /*
  Guides panel (left) / ［ガイド］
*/
    var panelGuides = colLeft.add("panel", undefined, LABELS.panelGuidesTitle[lang]);
    panelGuides.orientation = "column";
    panelGuides.alignChildren = ["fill", "top"];
    panelGuides.alignment = ["fill", "top"];
    panelGuides.margins = [15, 20, 15, 10];

    /* Unified label width inside Guides / ガイド内ラベル共通幅 */
    var GUIDE_LABEL_WIDTH = 80;

    // ガイド：カラー設定（1行レイアウト）
    var guideColorRow = panelGuides.add("group", undefined, "");
    guideColorRow.orientation = "row";
    guideColorRow.alignChildren = ["left", "center"];
    guideColorRow.spacing = 10;

    var lblGuideColor = guideColorRow.add("statictext", undefined, LABELS.guideColorLabel[lang]);
    lblGuideColor.preferredSize = [GUIDE_LABEL_WIDTH, lblGuideColor.preferredSize ? lblGuideColor.preferredSize[1] : 20];
    lblGuideColor.justify = "right";
    var rbGuideCyan = guideColorRow.add("radiobutton", undefined, LABELS.guideColorCyan[lang]);
    var rbGuideLightBlue = guideColorRow.add("radiobutton", undefined, LABELS.guideColorLightBlue[lang]);

    /* Known color presets (0..1) / 既知の色プリセット */
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


    /* Read current guide color once (source of truth) / ガイド色の取得元はここ（safeGetReal に統一） */
    var curR = safeGetReal("Guide/Color/red", GUIDE_CYAN.r);
    var curG = safeGetReal("Guide/Color/green", GUIDE_CYAN.g);
    var curB = safeGetReal("Guide/Color/blue", GUIDE_CYAN.b);

    if (approx(curR, GUIDE_LIGHTBLUE.r) && approx(curG, GUIDE_LIGHTBLUE.g) && approx(curB, GUIDE_LIGHTBLUE.b)) {
        rbGuideLightBlue.value = true;
    } else {
        rbGuideCyan.value = true; // 既定（シアン）
    }


    rbGuideCyan.onClick = function() {
        // UI only; persist on OK
    };
    rbGuideLightBlue.onClick = function() {
        // UI only; persist on OK
    };

    // ガイド：スタイル設定（1行レイアウト）
    var guideStyleRow = panelGuides.add("group", undefined, "");
    guideStyleRow.orientation = "row";
    guideStyleRow.alignChildren = ["left", "center"];
    guideStyleRow.spacing = 10;

    var lblGuideStyle = guideStyleRow.add("statictext", undefined, LABELS.guideStyleLabel[lang]);
    lblGuideStyle.preferredSize = [GUIDE_LABEL_WIDTH, lblGuideStyle.preferredSize ? lblGuideStyle.preferredSize[1] : 20];
    lblGuideStyle.justify = "right";
    var rbGuideStyleLine = guideStyleRow.add("radiobutton", undefined, LABELS.guideStyleLine[lang]);
    var rbGuideStyleDots = guideStyleRow.add("radiobutton", undefined, LABELS.guideStyleDots[lang]);

    // Guide/Style: 0 = Lines, 1 = Dots
    var curStyle = 0;
    try {
        curStyle = app.preferences.getIntegerPreference("Guide/Style");
    } catch (e) {
        curStyle = 0;
    }
    rbGuideStyleLine.value = (curStyle !== 1);
    rbGuideStyleDots.value = (curStyle === 1);

    rbGuideStyleLine.onClick = function() {
        // UI only; persist on OK
    };
    rbGuideStyleDots.onClick = function() {
        // UI only; persist on OK
    };


    /*
  UI panel (right) / ［ユーザーインターフェイス］
*/
    var panelUI = colRight.add("panel", undefined, LABELS.panelUITitle[lang]);
    panelUI.orientation = "column";
    panelUI.alignChildren = ["fill", "top"];
    panelUI.alignment = ["fill", "top"];
    panelUI.margins = [15, 25, 15, 10]; // 上、右、下、左のマージン

    /* Canvas color (UI panel) / カンバスカラー（UIパネル） */
    // NOTE: All controls update UI only; actual writes occur on OK.
    // カンバスカラー（1行レイアウト）
    var canvasRow = panelUI.add("group", undefined, "");
    canvasRow.orientation = "row";
    canvasRow.alignChildren = ["left", "center"];
    canvasRow.spacing = 10;

    canvasRow.add("statictext", undefined, LABELS.canvasColor[lang] + "：");
    var rbCanvasMatch = canvasRow.add("radiobutton", undefined, LABELS.canvasMatch[lang]);
    var rbCanvasWhite = canvasRow.add("radiobutton", undefined, LABELS.canvasWhite[lang]);

    /*
      初期状態を反映
    */
    var currentCanvas = safeGetInt("uiCanvasIsWhite", 0);
    rbCanvasWhite.value = (currentCanvas === 1);
    rbCanvasMatch.value = (currentCanvas !== 1);
    /*
      ラジオボタン変更時の設定
    */
    rbCanvasMatch.onClick = function() {
        // UI only; persist on OK
    };
    rbCanvasWhite.onClick = function() {
        // UI only; persist on OK
    };


    /*
  Performance panel (right) / ［パフォーマンス］
*/
    var panelPerf = colRight.add("panel", undefined, LABELS.panelPerfTitle[lang]);
    panelPerf.orientation = "column";
    panelPerf.alignChildren = ["fill", "top"];
    panelPerf.alignment = ["fill", "top"];
    panelPerf.margins = [15, 25, 15, 10];

    var cbAnimZoom = panelPerf.add("checkbox", undefined, LABELS.cbAnimZoom[lang]);
    cbAnimZoom.helpTip = LABELS.cbAnimZoom.ja + " / " + LABELS.cbAnimZoom.en;

    var groupHistory = panelPerf.add("group");
    groupHistory.add("statictext", undefined, LABELS.cbHistoryLabel[lang] + "：");
    var etHistory = groupHistory.add("edittext", undefined, "50");

    etHistory.onChange = function() {
        validateEditTextInt(etHistory, VALIDATION.history);
    };

    etHistory.characters = 4;
    etHistory.helpTip = "ヒストリー数を設定 / Set history steps";

    var cbLiveEdit = panelPerf.add("checkbox", undefined, LABELS.cbLiveEdit[lang]);
    cbLiveEdit.helpTip = LABELS.cbLiveEdit.ja + " / " + LABELS.cbLiveEdit.en;


    /*
  File Management panel (right) / ［ファイル管理］
*/
    var panelFile = colRight.add("panel", undefined, LABELS.panelFileTitle[lang]);
    panelFile.orientation = "column";
    panelFile.alignChildren = ["fill", "top"];
    panelFile.alignment = ["fill", "top"];
    panelFile.margins = [15, 25, 15, 10];

    var cbEditOriginal = panelFile.add("checkbox", undefined, LABELS.cbEditOriginal[lang]);
    cbEditOriginal.helpTip = LABELS.cbEditOriginal.ja + " / " + LABELS.cbEditOriginal.en;

    var cbFontsAuto = panelFile.add("checkbox", undefined, LABELS.cbFontsAuto[lang]);
    cbFontsAuto.helpTip = LABELS.cbFontsAuto.ja + " / " + LABELS.cbFontsAuto.en;

    /* Save location inside [File Management] / ［ファイル管理］内の保存先 */
    var saveRow2 = panelFile.add("group", undefined, "");
    saveRow2.orientation = "row";
    saveRow2.alignChildren = ["left", "center"];
    saveRow2.spacing = 10;

    saveRow2.add("statictext", undefined, LABELS.saveDest[lang] + "：");
    var rbSaveLocal = saveRow2.add("radiobutton", undefined, LABELS.saveLocal[lang]);
    var rbSaveCloud = saveRow2.add("radiobutton", undefined, LABELS.saveCloud[lang]);

    // 初期状態（true=クラウド、false=コンピューター）/ Initial state
    var prefCloud = false;
    try {
        prefCloud = app.preferences.getBooleanPreference("AdobeSaveAsCloudDocumentPreference");
    } catch (e) {
        prefCloud = false;
    }
    rbSaveCloud.value = prefCloud;
    rbSaveLocal.value = !prefCloud;

    // 変更時に保存 / Apply on change
    // NOTE: Persisted on OK.
    rbSaveLocal.onClick = function() {
        // UI only; persist on OK
    };
    rbSaveCloud.onClick = function() {
        // UI only; persist on OK
    };


    /*
  Limit to Path panel (right) / 「パスに制限」
*/
    var panelLimitPath = colRight.add("panel", undefined, LABELS.panelLimitPathTitle[lang]);
    panelLimitPath.orientation = "column";
    panelLimitPath.alignChildren = ["fill", "top"];
    panelLimitPath.alignment = ["fill", "top"];
    panelLimitPath.margins = [15, 25, 15, 10];

    var cbHitShape = panelLimitPath.add("checkbox", undefined, LABELS.cbHitShape[lang]);
    cbHitShape.helpTip = LABELS.cbHitShape.ja + " / " + LABELS.cbHitShape.en;
    /*
      設定読み込み（初期値代入）
      ※最適化案: UI構築と初期値代入を関数化して整理可能
    */
    var val = app.preferences.getIntegerPreference("hitShapeOnPreview");
    cbHitShape.value = (val === 0); // 0がON（true）、1がOFF（false）

    var cbHitTypeShape = panelLimitPath.add("checkbox", undefined, LABELS.cbHitTypeShape[lang]);
    cbHitTypeShape.helpTip = LABELS.cbHitTypeShape.ja + " / " + LABELS.cbHitTypeShape.en;
    // hitTypeShapeOnPreview: 0がON（true）、1がOFF（false）
    var val = app.preferences.getIntegerPreference("hitTypeShapeOnPreview");
    cbHitTypeShape.value = (val === 0); // 0がON（true）、1がOFF（false）

    // ===== Preferences Map (Proposal #1) =====
    // 各プリファレンスキー ↔ UI部品 ↔ 型をテーブル化
    var PREF_MAP = [{
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
        // cbHighlightAnchor はキー未確定のため保留
        {
            key: "hitTypeShapeOnPreview",
            type: "int01inv",
            ui: cbHitTypeShape
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
            key: "maximumUndoDepth",
            type: "int",
            ui: null
        }, // etHistory に反映
        {
            key: "LiveEdit_State_Machine",
            type: "bool",
            ui: cbLiveEdit
        },
        {
            key: "enableBackgroundExport",
            type: "bool",
            ui: cbEditOriginal
        },
        {
            key: "AutoActivateMissingFont",
            type: "bool",
            ui: cbFontsAuto
        },
        {
            key: "text/recentFontMenu/showNEntries",
            type: "int",
            ui: null
        }, // etRecentFonts, cbRecentFonts に反映
        {
            key: "uiCanvasIsWhite",
            type: "int",
            ui: null
        }
    ];

    // システムの現在値 → UIへ反映
    function loadPrefsFromSystem() {
        for (var i = 0; i < PREF_MAP.length; i++) {
            var p = PREF_MAP[i];
            try {
                if (p.type === "bool") {
                    var b = safeGetBool(p.key, false);
                    if (p.ui) p.ui.value = !!b;
                } else if (p.type === "int01inv") {
                    var v01 = app.preferences.getIntegerPreference(p.key);
                    if (p.ui) p.ui.value = (v01 === 0); // 0=>ON(true)
                } else if (p.type === "int") {
                    var iv = app.preferences.getIntegerPreference(p.key);
                    if (p.key === "maximumUndoDepth") {
                        etHistory.text = String(validateInt(iv, VALIDATION.history));
                    } else if (p.key === "text/recentFontMenu/showNEntries") {
                        var rv = validateInt(iv, VALIDATION.recentFonts);
                        etRecentFonts.text = String(rv);
                        cbRecentFonts.value = (rv > 0);
                        etRecentFonts.enabled = (rv > 0);
                    } else if (p.key === "uiCanvasIsWhite") {
                        rbCanvasWhite.value = (iv === 1);
                        rbCanvasMatch.value = (iv !== 1);
                    }
                }
            } catch (e) {
                // 読み込み失敗時は無視（既定UIのまま）
            }
        }
    }
    // ===== End Preferences Map =====

    // ===== Validation Rules (Proposal #8) =====
    var VALIDATION = {
        recentFonts: {
            min: 0,
            max: 30,
            def: 15
        }, // 0=非表示、上限は30程度が目安
        history: {
            min: 1,
            max: 1000,
            def: 100
        } // 1〜1000の想定
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
    // EditText用: 入力を検証し、範囲外や非数ならデフォルトへ
    function validateEditTextInt(editText, rule) {
        var v = validateInt(editText.text, rule);
        editText.text = String(v);
        return v;
    }
    // ===== End Validation Rules =====

    // ===== Safe Preference Accessors (Proposal #5) =====
    function safeSetInt(key, val) {
        try {
            app.preferences.setIntegerPreference(key, val);
        } catch (e) {}
    }

    function safeSetBool(key, val) {
        try {
            app.preferences.setBooleanPreference(key, val);
        } catch (e) {}
    }

    function safeSetReal(key, val) {
        try {
            app.preferences.setRealPreference(key, val);
        } catch (e) {}
    }

    function safeGetInt(key, fb) {
        try {
            return app.preferences.getIntegerPreference(key);
        } catch (e) {
            return fb;
        }
    }

    function safeGetBool(key, fb) {
        try {
            return app.preferences.getBooleanPreference(key);
        } catch (e) {
            return fb;
        }
    }

    function safeGetReal(key, fb) {
        try {
            return app.preferences.getRealPreference(key);
        } catch (e) {
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

    // ===== Preset Applier (Proposal #2) =====
    // データ駆動でUIへ反映する共通ヘルパー
    function applyPresetToUI(p) {
        // 単純なチェックボックス・数値はそのまま反映
        if (typeof p.cbToolTips !== 'undefined') cbToolTips.value = !!p.cbToolTips;
        if (typeof p.cbHomeScreen !== 'undefined') cbHomeScreen.value = !!p.cbHomeScreen;
        if (typeof p.cbLegacyNewDoc !== 'undefined') cbLegacyNewDoc.value = !!p.cbLegacyNewDoc;
        if (typeof p.cbBleedAI !== 'undefined') cbBleedAI.value = !!p.cbBleedAI;
        if (typeof p.cbMoveLocked !== 'undefined') cbMoveLocked.value = !!p.cbMoveLocked;
        if (typeof p.cbHitShape !== 'undefined') cbHitShape.value = !!p.cbHitShape; // true=ON(path only)
        if (typeof p.cbZoomToSel !== 'undefined') cbZoomToSel.value = !!p.cbZoomToSel;
        if (typeof p.cbHighlightAnchor !== 'undefined') cbHighlightAnchor.value = !!p.cbHighlightAnchor;
        if (typeof p.cbHitTypeShape !== 'undefined') cbHitTypeShape.value = !!p.cbHitTypeShape; // true=ON(path only)
        if (typeof p.cbAutoSizing !== 'undefined') cbAutoSizing.value = !!p.cbAutoSizing;
        if (typeof p.cbRecentFonts !== 'undefined') cbRecentFonts.value = !!p.cbRecentFonts;
        if (typeof p.etRecentFonts !== 'undefined') {
            etRecentFonts.text = String(p.etRecentFonts);
            etRecentFonts.enabled = !!p.cbRecentFonts;
        }
        if (typeof p.cbFontLocking !== 'undefined') cbFontLocking.value = !!p.cbFontLocking;
        if (typeof p.cbAlternateGlyph !== 'undefined') cbAlternateGlyph.value = !!p.cbAlternateGlyph;
        if (typeof p.cbAnimZoom !== 'undefined') cbAnimZoom.value = !!p.cbAnimZoom;
        if (typeof p.etHistory !== 'undefined') etHistory.text = String(p.etHistory);
        if (typeof p.cbLiveEdit !== 'undefined') cbLiveEdit.value = !!p.cbLiveEdit;
        if (typeof p.cbEditOriginal !== 'undefined') cbEditOriginal.value = !!p.cbEditOriginal;
        if (typeof p.cbFontsAuto !== 'undefined') cbFontsAuto.value = !!p.cbFontsAuto;
        // cbIncludeSVG removed

        // カンバス（ラジオ）
        if (typeof p.canvasWhite !== 'undefined') {
            rbCanvasWhite.value = !!p.canvasWhite;
            rbCanvasMatch.value = !rbCanvasWhite.value;
        }

        // ガイド色（UIのみ）
        if (p.guideColor === 'lightblue') {
            if (typeof rbGuideLightBlue !== 'undefined' && typeof rbGuideCyan !== 'undefined') {
                rbGuideLightBlue.value = true;
                rbGuideCyan.value = false;
            }
        } else if (p.guideColor === 'cyan') {
            if (typeof rbGuideLightBlue !== 'undefined' && typeof rbGuideCyan !== 'undefined') {
                rbGuideCyan.value = true;
                rbGuideLightBlue.value = false;
            }
        }

        // 保存先（クラウド/ローカル）（UIのみ）
        if (typeof p.saveCloud !== 'undefined') {
            rbSaveCloud.value = !!p.saveCloud;
            rbSaveLocal.value = !rbSaveCloud.value;
        }
    }

    // 既存プリセットをデータで表現
    var PRESET1_DEF = {
        cbToolTips: false,
        cbHomeScreen: false,
        cbLegacyNewDoc: true,
        cbBleedAI: false,
        cbMoveLocked: true,
        cbHitShape: false,
        cbZoomToSel: false,
        cbHighlightAnchor: true,
        cbHitTypeShape: false,
        cbAutoSizing: true,
        cbRecentFonts: false,
        etRecentFonts: 0,
        cbFontLocking: false,
        cbAlternateGlyph: false,
        cbAnimZoom: false,
        etHistory: 50,
        cbLiveEdit: true,
        cbEditOriginal: true,
        cbFontsAuto: true,
        canvasWhite: true,
        guideColor: 'lightblue',
        saveCloud: false
    };

    var DEFAULT_DEF = {
        cbToolTips: true,
        cbHomeScreen: true,
        cbLegacyNewDoc: false,
        cbBleedAI: true,
        cbMoveLocked: false,
        cbHitShape: false,
        cbZoomToSel: true,
        cbHighlightAnchor: true,
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
        saveCloud: true
    };
    // ===== End Preset Applier =====

    /* Bottom button row (Cancel/OK) / 右側にキャンセル・OK */
    var outerGroup = mainGroup.add("group"); // 親groupに追加
    outerGroup.orientation = "row";
    outerGroup.alignChildren = ["fill", "center"];
    outerGroup.alignment = ["fill", "bottom"];


    // 真ん中スペーサー（伸縮する空白）
    var spacer = outerGroup.add("group");
    spacer.alignment = ["fill", "fill"];

    // 右側グループ（キャンセル・OK）
    var rightGroup = outerGroup.add("group");
    rightGroup.orientation = "row";
    rightGroup.alignChildren = ["right", "center"];
    rightGroup.spacing = 10;
    var btnCancel = rightGroup.add("button", undefined, LABELS.Cancel[lang], {
        name: "cancel"
    });
    var btnOK = rightGroup.add("button", undefined, LABELS.OK[lang], {
        name: "ok"
    });

    btnOK.onClick = function() {
        try {
            /*
              「詳細なツールヒント」設定を保存
            */
            safeSetBool("showRichToolTips", cbToolTips.value);
            /*
              「ホーム画面」設定を保存
            */
            safeSetBool("Hello/ShowHomeScreenWS", cbHomeScreen.value);
            /*
              「以前の新規ドキュメント」設定を保存
            */
            safeSetBool("Hello/NewDoc", cbLegacyNewDoc.value);
            /*
              「裁ち落としを印刷」生成AIボタン設定を保存
            */
            safeSetBool("enablePrintBleedWidget", cbBleedAI.value);
            /*
              「アートボードと一緒に移動」設定を保存
            */
            safeSetBool("moveLockedAndHiddenArt", cbMoveLocked.value);
            /*
              「オブジェクトの選択範囲をパスに制限」設定を保存
            */
            safeSetInt("hitShapeOnPreview", cbHitShape.value ? 0 : 1);
            /*
              「選択範囲へズーム」設定を保存
            */
            safeSetBool("zoomToSelection", cbZoomToSel.value);
            /*
              「アンカーを強調表示」設定の保存は未実装（保存用のキーが不明なため）
            */
            /*
              「テキストをパスに制限」設定を保存
            */
            safeSetInt("hitTypeShapeOnPreview", cbHitTypeShape.value ? 0 : 1);
            /*
              「新規エリア内文字の自動サイズ調整」設定を保存
            */
            safeSetBool("text/autoSizing", cbAutoSizing.value);
            /*
              「最近使用したフォント」設定を保存
            */
            if (cbRecentFonts.value) {
                safeSetInt("text/recentFontMenu/showNEntries", validateEditTextInt(etRecentFonts, VALIDATION.recentFonts));
            } else {
                safeSetInt("text/recentFontMenu/showNEntries", 0);
            }
            /*
              「見つからない字形の保護」設定を保存
            */
            safeSetBool("text/doFontLocking", cbFontLocking.value);
            /*
              「選択された文字の異体字」設定を保存
            */
            safeSetBool("text/enableAlternateGlyph", cbAlternateGlyph.value);
            /*
              カンバス白設定を保存（ラジオ選択に基づく）
            */
            safeSetInt("uiCanvasIsWhite", rbCanvasWhite.value ? 1 : 0);
            /*
              ガイド色を保存（ラジオ選択に基づく）
            */
            if (rbGuideLightBlue && rbGuideLightBlue.value) {
                safeSetReal("Guide/Color/red", 0.29);
                safeSetReal("Guide/Color/green", 0.52);
                safeSetReal("Guide/Color/blue", 1.0);
            } else {
                // 既定（シアン）
                safeSetReal("Guide/Color/red", 0.0);
                safeSetReal("Guide/Color/green", 1.0);
                safeSetReal("Guide/Color/blue", 1.0);
            }
            /*
              ガイドスタイルを保存（0:ライン / 1:点線）
            */
            if (rbGuideStyleLine && rbGuideStyleLine.value) {
                safeSetInt("Guide/Style", 0);
            } else if (rbGuideStyleDots && rbGuideStyleDots.value) {
                safeSetInt("Guide/Style", 1);
            }
            /*
              「アニメーションズーム」設定を保存
            */
            safeSetBool("Performance/AnimZoom", cbAnimZoom.value);
            /*
              「ヒストリー数」設定を保存
            */
            safeSetInt("maximumUndoDepth", validateEditTextInt(etHistory, VALIDATION.history));
            /*
              「リアルタイムの描画と編集」設定を保存
            */
            safeSetBool("LiveEdit_State_Machine", cbLiveEdit.value);
            /*
              「オリジナルの編集にシステムデフォルトを使用」設定を保存
            */
            safeSetBool("enableBackgroundExport", cbEditOriginal.value);
            /*
              「Adobe Fontsを自動アクティベート」設定を保存
            */
            safeSetBool("AutoActivateMissingFont", cbFontsAuto.value);
            /*
              ファイルの保存先を保存（OK押下時にも明示保存）
            */
            safeSetBool("AdobeSaveAsCloudDocumentPreference", !!rbSaveCloud.value);
        } catch (e) {
            alert("環境設定の保存に失敗しました: " + e + "\nFailed to save preferences: " + e);
        }
        dlg.close();
    };
    /*
      初期状態として現状の環境設定を読み込み
    */
    applyPresetSettings("現在の設定");
    /* Adjust dialog opacity & position / ダイアログの不透明度と位置 */
    var offsetX = 300;
    var dialogOpacity = 0.97;

    function shiftDialogPosition(dlg, offsetX, offsetY) {
        dlg.onShow = function() {
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