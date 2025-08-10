// 既に #target illustrator, ShowExternalJSXWarning の記述があるか確認し、なければ冒頭に挿入
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
  スクリプト名：Illustrator環境設定ユーティリティ
  概要：Illustratorの各種環境設定をGUIからまとめて変更できます。
       ユニット、キャンバスカラー、ヒットテスト条件などをプリセットや現在値から制御可能です。

  更新履歴：
  - v1.0 (20250807): 初期リリース
  - v1.1 (20250915): ファイルの保存先を追加
*/

var SCRIPT_VERSION = "v1.1";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */

var LABELS = {
    dialogTitle: {
        ja: "環境設定をまとめて変更",
        en: "Batch Preferences Dialog"
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
        ja: "ドキュメントを開いていないときにホーム画面を表示",
        en: "Show The Home Screen When No Documents Are Open"
    },
    cbLegacyNewDoc: {
        ja: "以前の「新規ドキュメント」インターフェイスを使用",
        en: "Use legacy “File New” interface"
    },
    cbBleedAI: {
        ja: "裁ち落とし部分に「裁ち落としを印刷」生成AIボタンを表示",
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
    }
    , saveDest: {
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
};

function main() {

    /*
      Illustrator 環境設定ダイアログを作成
    */
    var dlg = new Window("dialog", LABELS.dialogTitle[lang]);
    /*
      メイングループ（縦方向）を作成
    */
    var mainGroup = dlg.add("group");
    mainGroup.orientation = "column";
    mainGroup.alignChildren = "left";

    /*
      プリセット選択グループ
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
    var presetItems = ["無指定", "デフォルト", "プリセット1"];
    var ddPreset = groupPresetInner.add("dropdownlist", undefined, presetItems);
    ddPreset.selection = 0; // 初期選択を「無指定」に

    /*
      プリセット設定適用関数
      ※最適化案: 各プリセットの重複処理を共通関数にまとめることが可能
    */
    function applyPresetSettings(presetName) {
        if (presetName === "無指定") {
            /*
              現在の環境設定を UI に反映する
            */
            cbToolTips.value = app.preferences.getBooleanPreference("showRichToolTips");
            cbHomeScreen.value = app.preferences.getBooleanPreference("Hello/ShowHomeScreenWS");
            cbLegacyNewDoc.value = app.preferences.getBooleanPreference("Hello/NewDoc");
            cbBleedAI.value = app.preferences.getBooleanPreference("enablePrintBleedWidget");
            cbMoveLocked.value = app.preferences.getBooleanPreference("moveLockedAndHiddenArt");
            cbHitShape.value = app.preferences.getBooleanPreference("hitShapeOnPreview");
            cbZoomToSel.value = app.preferences.getBooleanPreference("zoomToSelection");
            cbHighlightAnchor.value = app.preferences.getBooleanPreference("cbHighlightAnchor");
            cbHitTypeShape.value = app.preferences.getBooleanPreference("hitTypeShapeOnPreview");
            cbAutoSizing.value = app.preferences.getBooleanPreference("text/autoSizing");
            var recentCount = app.preferences.getIntegerPreference("text/recentFontMenu/showNEntries");
            cbRecentFonts.value = (recentCount > 0);
            etRecentFonts.text = String(recentCount);
            etRecentFonts.enabled = (recentCount > 0);
            cbFontLocking.value = app.preferences.getBooleanPreference("text/doFontLocking");
            cbAlternateGlyph.value = app.preferences.getBooleanPreference("text/enableAlternateGlyph");
            cbAnimZoom.value = app.preferences.getBooleanPreference("Performance/AnimZoom");
            etHistory.text = String(app.preferences.getIntegerPreference("maximumUndoDepth"));
            cbLiveEdit.value = app.preferences.getBooleanPreference("LiveEdit_State_Machine");
            cbEditOriginal.value = app.preferences.getBooleanPreference("enableBackgroundExport");
            cbFontsAuto.value = app.preferences.getBooleanPreference("AutoActivateMissingFont");
            /*
              カンバスカラー
            */
            var currentCanvas = app.preferences.getIntegerPreference("uiCanvasIsWhite");
            rbCanvasWhite.value = (currentCanvas === 1);
            rbCanvasMatch.value = (currentCanvas !== 1);
            /*
              ファイルの保存先（現在値をUIに反映）
            */
            var curCloud = false;
            try {
                curCloud = app.preferences.getBooleanPreference("AdobeSaveAsCloudDocumentPreference");
            } catch (e) {
                curCloud = false;
            }
            rbSaveCloud.value = curCloud;
            rbSaveLocal.value = !curCloud;
        } else if (presetName === "プリセット1") {
            cbToolTips.value = false;
            cbHomeScreen.value = false;
            cbLegacyNewDoc.value = true;
            cbBleedAI.value = false;
            cbMoveLocked.value = true;
            cbHitShape.value = false;
            cbZoomToSel.value = false;
            cbHighlightAnchor.value = true;
            cbHitTypeShape.value = false;
            cbAutoSizing.value = true;
            cbRecentFonts.value = false;
            etRecentFonts.text = "0";
            etRecentFonts.enabled = false;
            cbFontLocking.value = false;
            cbAlternateGlyph.value = false;
            cbAnimZoom.value = false;
            etHistory.text = "50";
            cbLiveEdit.value = true;
            cbEditOriginal.value = true;
            cbFontsAuto.value = true;
            rbCanvasWhite.value = true;
            rbCanvasMatch.value = false;
            app.preferences.setIntegerPreference("uiCanvasIsWhite", 1);
            // ファイルの保存先：コンピューター（false）
            rbSaveLocal.value = true;
            rbSaveCloud.value = false;
            try { app.preferences.setBooleanPreference("AdobeSaveAsCloudDocumentPreference", false); } catch (e) {}
        } else if (presetName === "デフォルト") {
            cbToolTips.value = true;
            cbHomeScreen.value = true;
            cbLegacyNewDoc.value = false;
            cbBleedAI.value = true;
            cbMoveLocked.value = false;
            cbHitShape.value = true;
            cbZoomToSel.value = true;
            cbHighlightAnchor.value = true;
            cbHitTypeShape.value = true;
            cbAutoSizing.value = false;
            cbRecentFonts.value = true;
            etRecentFonts.text = "10";
            etRecentFonts.enabled = true;
            cbFontLocking.value = true;
            cbAlternateGlyph.value = true;
            cbAnimZoom.value = true;
            etHistory.text = "100";
            cbLiveEdit.value = true;
            cbEditOriginal.value = false;
            cbFontsAuto.value = false;
            rbCanvasWhite.value = false;
            rbCanvasMatch.value = true;
            app.preferences.setIntegerPreference("uiCanvasIsWhite", 0);
            // ファイルの保存先：クラウド（true）
            rbSaveLocal.value = false;
            rbSaveCloud.value = true;
            try { app.preferences.setBooleanPreference("AdobeSaveAsCloudDocumentPreference", true); } catch (e) {}
        }
    }



    ddPreset.onChange = function() {
        if (ddPreset.selection && ddPreset.selection.text) {
            applyPresetSettings(ddPreset.selection.text);
            app.redraw();
        }
    };


    /*
      デフォルトプリセットに現状の環境設定を取得
      ※最適化案: プリセット値の保存・読込を関数化し、プリセット管理ロジックを整理可能
    */
    var defaultPreset = {
        // 現在の「詳細なツールヒント」設定を取得 / Get current setting of rich tooltips
        showRichToolTips: app.preferences.getBooleanPreference("showRichToolTips"),
        // 現在の「ホーム画面」設定を取得 / Get current setting of Home Screen
        homeScreen: app.preferences.getBooleanPreference("Hello/ShowHomeScreenWS"),
        // 現在の「以前の新規ドキュメント」設定を取得 / Get current setting of legacy New Document
        legacyNewDoc: app.preferences.getBooleanPreference("Hello/NewDoc"),
        // 現在の「裁ち落としを印刷」生成AIボタン設定を取得 / Get current Bleed AI button setting
        bleedAI: app.preferences.getBooleanPreference("enablePrintBleedWidget"),
        // 現在の「アートボードと一緒に移動」設定を取得 / Get current move locked objects with artboard
        moveLocked: app.preferences.getBooleanPreference("moveLockedAndHiddenArt"),
        // 現在の「オブジェクトの選択範囲をパスに制限」設定を取得 / Get current limit selection area to path
        hitShape: app.preferences.getBooleanPreference("hitShapeOnPreview"),
        // 現在の「選択範囲へズーム」設定を取得 / Get current zoom to selection setting
        zoomToSel: app.preferences.getBooleanPreference("zoomToSelection"),
        // 現在の「テキストをパスに制限」設定を取得 / Get current limit selection area of text objects to path
        hitTypeShape: app.preferences.getBooleanPreference("hitTypeShapeOnPreview"),
        // 現在の「新規エリア内文字の自動サイズ調整」設定を取得 / Get current auto size new area text
        autoSizing: app.preferences.getBooleanPreference("text/autoSizing"),
        // 現在の「最近使用したフォント」数を取得 / Get current recent fonts count
        recentFonts: app.preferences.getIntegerPreference("text/recentFontMenu/showNEntries"),
        // 現在の「見つからない字形の保護」設定を取得 / Get current font locking setting
        fontLocking: app.preferences.getBooleanPreference("text/doFontLocking"),
        // 現在の「選択された文字の異体字」設定を取得 / Get current alternate glyphs setting
        alternateGlyph: app.preferences.getBooleanPreference("text/enableAlternateGlyph"),
        // 現在のUI明るさ設定を取得 / Get current UI brightness
        uiBrightness: app.preferences.getRealPreference("uiBrightness"),
        // 現在のカンバス白設定を取得 / Get current canvas white setting
        uiCanvasIsWhite: app.preferences.getIntegerPreference("uiCanvasIsWhite"),
        // 現在の「アニメーションズーム」設定を取得 / Get current animated zoom setting
        animZoom: app.preferences.getBooleanPreference("Performance/AnimZoom"),
        // 現在の「ヒストリー数」を取得 / Get current history steps
        undoDepth: app.preferences.getIntegerPreference("maximumUndoDepth"),
        // 現在の「リアルタイムの描画と編集」設定を取得 / Get current real-time editing setting
        liveEdit: app.preferences.getBooleanPreference("LiveEdit_State_Machine"),
        // 現在の「enableBackgroundExport」設定を取得 / Get current background export setting
        bgExport: app.preferences.getBooleanPreference("enableBackgroundExport"),
        // 現在の「cloudAIEnableAutoSave」設定を取得 / Get current cloud auto save setting
        cloudAutoSave: app.preferences.getBooleanPreference("cloudAIEnableAutoSave"),
        // 現在の「enableBackgroundSave」設定を取得 / Get current background save setting
        bgSave: app.preferences.getBooleanPreference("enableBackgroundSave"),
        // 現在の「enableBackgroundExport」設定を取得 / Get current edit original setting
        editOriginal: app.preferences.getBooleanPreference("enableBackgroundExport"),
        // 現在の「AutoActivateMissingFont」設定を取得 / Get current auto-activate missing font setting
        fontsAuto: app.preferences.getBooleanPreference("AutoActivateMissingFont")
    };

    /*
      2カラム用グループを作成
    */
    var colGroup = mainGroup.add("group");
    colGroup.orientation = "row";
    colGroup.alignChildren = "top";

    var colLeft = colGroup.add("group");
    colLeft.orientation = "column";
    colLeft.alignChildren = "left";

    var colRight = colGroup.add("group");
    colRight.orientation = "column";
    colRight.alignChildren = "left";

    /*
      ［一般］パネル（左カラム）
    */
    var panelGeneral = colLeft.add("panel", undefined, "［一般］カテゴリ");
    panelGeneral.orientation = "column";
    panelGeneral.alignChildren = "left";
    panelGeneral.margins = [15, 25, 15, 10]; // 上、右、下、左のマージン

    /*
      「詳細なツールヒント」チェックボックスを作成
    */
    var cbToolTips = panelGeneral.add("checkbox", undefined, LABELS.cbToolTips[lang]);
    cbToolTips.helpTip = LABELS.cbToolTips.ja + " / " + LABELS.cbToolTips.en;

    /*
      「ホーム画面」チェックボックスを作成
    */
    var cbHomeScreen = panelGeneral.add("checkbox", undefined, LABELS.cbHomeScreen[lang]);
    cbHomeScreen.helpTip = LABELS.cbHomeScreen.ja + " / " + LABELS.cbHomeScreen.en;

    /*
      「以前の新規ドキュメント」チェックボックスを作成
    */
    var cbLegacyNewDoc = panelGeneral.add("checkbox", undefined, LABELS.cbLegacyNewDoc[lang]);
    cbLegacyNewDoc.helpTip = LABELS.cbLegacyNewDoc.ja + " / " + LABELS.cbLegacyNewDoc.en;

    /*
      「裁ち落としを印刷」生成AIボタンチェックボックスを作成
    */
    var cbBleedAI = panelGeneral.add("checkbox", undefined, LABELS.cbBleedAI[lang]);
    cbBleedAI.helpTip = LABELS.cbBleedAI.ja + " / " + LABELS.cbBleedAI.en;

    /*
      ［選択範囲・アンカー表示］パネル（左カラム）
    */
    var panelSelectAnchor = colLeft.add("panel", undefined, "［選択範囲・アンカー表示］カテゴリ");
    panelSelectAnchor.orientation = "column";
    panelSelectAnchor.alignChildren = "left";
    panelSelectAnchor.margins = [15, 25, 15, 10];

    /*
      「アートボードと一緒に移動」チェックボックスを作成
    */
    var cbMoveLocked = panelSelectAnchor.add("checkbox", undefined, LABELS.cbMoveLocked[lang]);
    cbMoveLocked.helpTip = LABELS.cbMoveLocked.ja + " / " + LABELS.cbMoveLocked.en;

    /*
      「選択範囲へズーム」チェックボックスを作成
    */
    var cbZoomToSel = panelSelectAnchor.add("checkbox", undefined, LABELS.cbZoomToSel[lang]);
    cbZoomToSel.helpTip = LABELS.cbZoomToSel.ja + " / " + LABELS.cbZoomToSel.en;

    /*
      「アンカーを強調表示」チェックボックスを作成
    */
    var cbHighlightAnchor = panelSelectAnchor.add("checkbox", undefined, LABELS.cbHighlightAnchor[lang]);
    cbHighlightAnchor.helpTip = LABELS.cbHighlightAnchor.ja + " / " + LABELS.cbHighlightAnchor.en;

    /*
      テキスト関連UIを左カラムの下部にパネルとして追加
    */
    var panelTextRight = colLeft.add("panel", undefined, "［テキスト］カテゴリ");
    panelTextRight.orientation = "column";
    panelTextRight.alignChildren = "left";
    panelTextRight.margins = [15, 25, 15, 10];

    /*
      「新規エリア内文字の自動サイズ調整」チェックボックスを作成
    */
    var cbAutoSizing = panelTextRight.add("checkbox", undefined, LABELS.cbAutoSizing[lang]);
    cbAutoSizing.helpTip = LABELS.cbAutoSizing.ja + " / " + LABELS.cbAutoSizing.en;


    /*
      最近使用したフォント数のUIを作成
    */
    var currentRecentCount = app.preferences.getIntegerPreference("text/recentFontMenu/showNEntries");
    var groupRecentFonts = panelTextRight.add("group");
    groupRecentFonts.orientation = "row";

    /*
      「最近使用したフォント」チェックボックスを作成
    */
    var cbRecentFonts = groupRecentFonts.add("checkbox", undefined, LABELS.cbRecentFonts[lang]);
    cbRecentFonts.helpTip = LABELS.cbRecentFonts.ja + " / " + LABELS.cbRecentFonts.en;
    cbRecentFonts.value = (currentRecentCount > 0);

    /*
      「最近使用したフォント数」入力欄を作成
    */
    var etRecentFonts = groupRecentFonts.add("edittext", undefined, currentRecentCount.toString());
    etRecentFonts.characters = 3;
    etRecentFonts.enabled = cbRecentFonts.value;

    cbRecentFonts.onClick = function() {
        /*
          「最近使用したフォント」チェックボックスの状態に応じて入力欄を有効化
        */
        if (cbRecentFonts.value) {
            if (parseInt(etRecentFonts.text, 10) === 0 || isNaN(parseInt(etRecentFonts.text, 10))) {
                etRecentFonts.text = "15";
            }
            etRecentFonts.enabled = true;
            app.preferences.setIntegerPreference("text/recentFontMenu/showNEntries", parseInt(etRecentFonts.text, 10));
        } else {
            etRecentFonts.enabled = false;
            etRecentFonts.text = "0";
            app.preferences.setIntegerPreference("text/recentFontMenu/showNEntries", 0);
        }
    };

    /*
      「見つからない字形の保護」チェックボックスを作成
    */
    var cbFontLocking = panelTextRight.add("checkbox", undefined, LABELS.cbFontLocking[lang]);
    cbFontLocking.helpTip = LABELS.cbFontLocking.ja + " / " + LABELS.cbFontLocking.en;

    /*
      「選択された文字の異体字」チェックボックスを作成
    */
    var cbAlternateGlyph = panelTextRight.add("checkbox", undefined, LABELS.cbAlternateGlyph[lang]);
    cbAlternateGlyph.helpTip = LABELS.cbAlternateGlyph.ja + " / " + LABELS.cbAlternateGlyph.en;


    /*
      ユーザーインターフェイスパネル
    */
    var panelUI = colRight.add("panel", undefined, "［ユーザーインターフェイス］カテゴリ");
    panelUI.orientation = "column";
    panelUI.alignChildren = "left";
    panelUI.margins = [15, 25, 15, 10]; // 上、右、下、左のマージン

    /*
      カンバスカラー設定
      （1行レイアウトでラベルとラジオボタンを横並びにする）
    */
    // カンバスカラー（1行レイアウト）
    var canvasRow = panelUI.add("group", undefined, "");
    canvasRow.orientation = "row";
    canvasRow.alignChildren = ["left", "center"];
    canvasRow.spacing = 10;

    canvasRow.add("statictext", undefined, LABELS.canvasColor[lang] + "：");
    var rbCanvasMatch = canvasRow.add("radiobutton", undefined, "UIに合わせる");
    var rbCanvasWhite = canvasRow.add("radiobutton", undefined, "ホワイト");

    /*
      初期状態を反映
    */
    var currentCanvas = app.preferences.getIntegerPreference("uiCanvasIsWhite");
    rbCanvasWhite.value = (currentCanvas === 1);
    rbCanvasMatch.value = (currentCanvas !== 1);
    /*
      ラジオボタン変更時の設定
    */
    rbCanvasMatch.onClick = function() {
        app.preferences.setIntegerPreference("uiCanvasIsWhite", 0);
    };
    rbCanvasWhite.onClick = function() {
        app.preferences.setIntegerPreference("uiCanvasIsWhite", 1);
    };


    /*
      ［パフォーマンス］パネル（右カラム）
    */
    var panelPerf = colRight.add("panel", undefined, "［パフォーマンス］カテゴリ");
    panelPerf.orientation = "column";
    panelPerf.alignChildren = "left";
    panelPerf.margins = [15, 25, 15, 10];

    /*
      「アニメーションズーム」チェックボックスを作成
    */
    var cbAnimZoom = panelPerf.add("checkbox", undefined, LABELS.cbAnimZoom[lang]);
    cbAnimZoom.helpTip = LABELS.cbAnimZoom.ja + " / " + LABELS.cbAnimZoom.en;

    /*
      「ヒストリー数」入力欄を作成
    */
    var groupHistory = panelPerf.add("group");
    groupHistory.add("statictext", undefined, LABELS.cbHistoryLabel[lang] + "：");
    var etHistory = groupHistory.add("edittext", undefined, "50");
    etHistory.characters = 4;
    etHistory.helpTip = "ヒストリー数を設定 / Set history steps";

    /*
      「リアルタイムの描画と編集」チェックボックスを作成
    */
    var cbLiveEdit = panelPerf.add("checkbox", undefined, LABELS.cbLiveEdit[lang]);
    cbLiveEdit.helpTip = LABELS.cbLiveEdit.ja + " / " + LABELS.cbLiveEdit.en;


    /*
      ［ファイル管理］パネル（右カラム）
    */
    var panelFile = colRight.add("panel", undefined, "［ファイル管理］カテゴリ");
    panelFile.orientation = "column";
    panelFile.alignChildren = "left";
    panelFile.margins = [15, 25, 15, 10];

    /*
      「オリジナルの編集にシステムデフォルトを使用」チェックボックスを作成
    */
    var cbEditOriginal = panelFile.add("checkbox", undefined, LABELS.cbEditOriginal[lang]);
    cbEditOriginal.helpTip = LABELS.cbEditOriginal.ja + " / " + LABELS.cbEditOriginal.en;

    /*
      「Adobe Fontsを自動アクティベート」チェックボックスを作成
    */
    var cbFontsAuto = panelFile.add("checkbox", undefined, LABELS.cbFontsAuto[lang]);
    cbFontsAuto.helpTip = LABELS.cbFontsAuto.ja + " / " + LABELS.cbFontsAuto.en;

    // ファイルの保存先（テキスト＋ラジオを［ファイル管理］に統合）/ Save location inside [File Management]
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
    rbSaveLocal.onClick = function () {
        try { app.preferences.setBooleanPreference("AdobeSaveAsCloudDocumentPreference", false); } catch (e) {}
    };
    rbSaveCloud.onClick = function () {
        try { app.preferences.setBooleanPreference("AdobeSaveAsCloudDocumentPreference", true); } catch (e) {}
    };

    /*
      右カラムの一番下に「パスに制限」パネルを追加
    */
    var panelLimitPath = colRight.add("panel", undefined, "パスに制限");
    panelLimitPath.orientation = "column";
    panelLimitPath.alignChildren = "left";
    panelLimitPath.margins = [15, 25, 15, 10];

    /*
      「オブジェクトの選択範囲をパスに制限」チェックボックスを panelLimitPath に追加
    */
    var cbHitShape = panelLimitPath.add("checkbox", undefined, LABELS.cbHitShape[lang]);
    cbHitShape.helpTip = LABELS.cbHitShape.ja + " / " + LABELS.cbHitShape.en;
    /*
      設定読み込み（初期値代入）
      ※最適化案: UI構築と初期値代入を関数化して整理可能
    */
    var val = app.preferences.getIntegerPreference("hitShapeOnPreview");
    cbHitShape.value = (val === 0); // 0がON（true）、1がOFF（false）

    /*
      「テキストをパスに制限」チェックボックスを panelLimitPath に追加
    */
    var cbHitTypeShape = panelLimitPath.add("checkbox", undefined, LABELS.cbHitTypeShape[lang]);
    cbHitTypeShape.helpTip = LABELS.cbHitTypeShape.ja + " / " + LABELS.cbHitTypeShape.en;
    // hitTypeShapeOnPreview: 0がON（true）、1がOFF（false）
    var val = app.preferences.getIntegerPreference("hitTypeShapeOnPreview");
    cbHitTypeShape.value = (val === 0); // 0がON（true）、1がOFF（false）

    /*
      ボタン行を mainGroup の下部に追加（テンプレート形式：左：書き出し、中央スペーサー、右：キャンセル・OK）
    */
    var outerGroup = mainGroup.add("group"); // 親groupに追加
    outerGroup.orientation = "row";
outerGroup.alignChildren = ["fill", "center"];
outerGroup.alignment = ["fill", "bottom"];

    // 左側グループ（書き出し）
    var leftGroup = outerGroup.add("group");
    leftGroup.orientation = "row";
    leftGroup.alignChildren = "left";
    var btnExport = leftGroup.add("button", undefined, "書き出し", { name: "export" });

    // 真ん中スペーサー（伸縮する空白）
    var spacer = outerGroup.add("group");
    spacer.alignment = ["fill", "fill"];

    // 右側グループ（キャンセル・OK）
    var rightGroup = outerGroup.add("group");
    rightGroup.orientation = "row";
    rightGroup.alignChildren = ["right", "center"];
    rightGroup.spacing = 10;
    var btnCancel = rightGroup.add("button", undefined, LABELS.Cancel[lang], { name: "cancel" });
    var btnOK = rightGroup.add("button", undefined, LABELS.OK[lang], { name: "ok" });

    btnOK.onClick = function() {
        try {
            /*
              「詳細なツールヒント」設定を保存
            */
            app.preferences.setBooleanPreference("showRichToolTips", cbToolTips.value);
            /*
              「ホーム画面」設定を保存
            */
            app.preferences.setBooleanPreference("Hello/ShowHomeScreenWS", cbHomeScreen.value);
            /*
              「以前の新規ドキュメント」設定を保存
            */
            app.preferences.setBooleanPreference("Hello/NewDoc", cbLegacyNewDoc.value);
            /*
              「裁ち落としを印刷」生成AIボタン設定を保存
            */
            app.preferences.setBooleanPreference("enablePrintBleedWidget", cbBleedAI.value);
            /*
              「アートボードと一緒に移動」設定を保存
            */
            app.preferences.setBooleanPreference("moveLockedAndHiddenArt", cbMoveLocked.value);
            /*
              「オブジェクトの選択範囲をパスに制限」設定を保存
            */
            app.preferences.setIntegerPreference("hitShapeOnPreview", cbHitShape.value ? 0 : 1);
            /*
              「選択範囲へズーム」設定を保存
            */
            app.preferences.setBooleanPreference("zoomToSelection", cbZoomToSel.value);
            /*
              「アンカーを強調表示」設定の保存は未実装（保存用のキーが不明なため）
            */
            /*
              「テキストをパスに制限」設定を保存
            */
            app.preferences.setIntegerPreference("hitTypeShapeOnPreview", cbHitTypeShape.value ? 0 : 1);
            /*
              「新規エリア内文字の自動サイズ調整」設定を保存
            */
            app.preferences.setBooleanPreference("text/autoSizing", cbAutoSizing.value);
            /*
              「最近使用したフォント」設定を保存
            */
            if (cbRecentFonts.value) {
                app.preferences.setIntegerPreference("text/recentFontMenu/showNEntries", parseInt(etRecentFonts.text, 10));
            } else {
                app.preferences.setIntegerPreference("text/recentFontMenu/showNEntries", 0);
            }
            /*
              「見つからない字形の保護」設定を保存
            */
            app.preferences.setBooleanPreference("text/doFontLocking", cbFontLocking.value);
            /*
              「選択された文字の異体字」設定を保存
            */
            app.preferences.setBooleanPreference("text/enableAlternateGlyph", cbAlternateGlyph.value);
            /*
              カンバス白設定を保存（ラジオ選択に基づく）
            */
            app.preferences.setIntegerPreference("uiCanvasIsWhite", rbCanvasWhite.value ? 1 : 0);
            /*
              「アニメーションズーム」設定を保存
            */
            app.preferences.setBooleanPreference("Performance/AnimZoom", cbAnimZoom.value);
            /*
              「ヒストリー数」設定を保存
            */
            app.preferences.setIntegerPreference("maximumUndoDepth", parseInt(etHistory.text, 10));
            /*
              「リアルタイムの描画と編集」設定を保存
            */
            app.preferences.setBooleanPreference("LiveEdit_State_Machine", cbLiveEdit.value);
            /*
              「オリジナルの編集にシステムデフォルトを使用」設定を保存
            */
            app.preferences.setBooleanPreference("enableBackgroundExport", cbEditOriginal.value);
            /*
              「Adobe Fontsを自動アクティベート」設定を保存
            */
            app.preferences.setBooleanPreference("AutoActivateMissingFont", cbFontsAuto.value);
            /*
              ファイルの保存先を保存（OK押下時にも明示保存）
            */
            try {
                app.preferences.setBooleanPreference("AdobeSaveAsCloudDocumentPreference", rbSaveCloud.value ? true : false);
            } catch (e) {}
        } catch (e) {
            alert("環境設定の保存に失敗しました: " + e + "\nFailed to save preferences: " + e);
        }
        dlg.close();
    };
    /*
      環境設定値をUIに反映
    */
    if (cbHitShape && app.preferences.getIntegerPreference) {
        cbHitShape.value = (app.preferences.getIntegerPreference("hitShapeOnPreview") === 0);
    }
    if (cbHitTypeShape && app.preferences.getIntegerPreference) {
        cbHitTypeShape.value = (app.preferences.getIntegerPreference("hitTypeShapeOnPreview") === 0);
    }
    /*
      初期状態として現状の環境設定を読み込み
    */
    applyPresetSettings("無指定");
    /*
      ダイアログの透明度と位置を調整 / Adjust dialog opacity and position
    */
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
    /*
      最後に表示
    */
    dlg.show();
}

main();