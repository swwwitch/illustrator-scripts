#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
  Illustrator 環境設定プリセット切替スクリプト
  このスクリプトは、Illustrator の環境設定を保存・切り替え可能なプリセットとして管理します。
  ユーザーは「デフォルト」と「プリセット1」などを選択し、ワンクリックで設定を切り替えることができます。
  
  対象設定：
  - 一般
  - 選択範囲・アンカー表示
  - テキスト
  - ユーザーインターフェイス（カンバスカラー）
  - パフォーマンス
  - ファイル管理

  更新履歴：
  - v1.0 (20250807) : 初期バージョン
*/


function main() {
    try {
        // Illustrator 環境設定ダイアログを作成 / Create Illustrator preferences dialog
        var dlg = new Window("dialog", "Illustrator 環境設定");
        // メイングループ（縦方向）を作成 / Create main vertical group
        var mainGroup = dlg.add("group");
        mainGroup.orientation = "column";
        mainGroup.alignChildren = "left";

        // プリセット選択グループ / Preset selection group
        var groupPreset = mainGroup.add("group");
        groupPreset.name = "groupPreset";
        groupPreset.alignment = "center";
        groupPreset.orientation = "row";
        groupPreset.margins = [0, 10, 20, 20]; // 上、右、下、左のマージン
        // wrap label and dropdown in a sub-group for alignment (if needed)
        var groupPresetInner = groupPreset.add("group");
        groupPresetInner.orientation = "row";
        groupPresetInner.alignChildren = "center";
        groupPresetInner.add("statictext", undefined, "プリセット："); // ラベル / Label
        var presetItems = ["デフォルト", "プリセット1"];
        var ddPreset = groupPresetInner.add("dropdownlist", undefined, presetItems);
        ddPreset.selection = 0; // デフォルトを選択

        // プリセット1の設定を適用 / Apply settings for preset1
        function applyPreset1Settings() {
            // 「詳細なツールヒント」設定を適用 / Apply rich tooltips setting
            cbToolTips.value = false;
            // 「ホーム画面」設定を適用 / Apply Home Screen setting
            cbHomeScreen.value = false;
            // 「以前の新規ドキュメント」設定を適用 / Apply legacy new document setting
            cbLegacyNewDoc.value = true;
            // 「裁ち落としを印刷」生成AIボタン設定を適用 / Apply Bleed AI button setting
            cbBleedAI.value = false;

            // 「アートボードと一緒に移動」設定を適用 / Apply move locked objects with artboard
            cbMoveLocked.value = true;
            // 「オブジェクトの選択範囲をパスに制限」設定を適用 / Apply limit selection area to path
            cbHitShape.value = false;
            // 「選択範囲へズーム」設定を適用 / Apply zoom to selection
            cbZoomToSel.value = false;
            // 「アンカーを強調表示」設定を適用 / Apply highlight anchor
            cbHighlightAnchor.value = true;

            // 「テキストをパスに制限」設定を適用 / Apply limit selection area of text objects to path
            cbHitTypeShape.value = false;
            // 「新規エリア内文字の自動サイズ調整」設定を適用 / Apply auto size new area text
            cbAutoSizing.value = true;
            // 「最近使用したフォント」設定を適用 / Apply recent fonts setting
            cbRecentFonts.value = false;
            etRecentFonts.text = "0";
            etRecentFonts.enabled = false;
            // 「見つからない字形の保護」設定を適用 / Apply font locking setting
            cbFontLocking.value = false;
            // 「選択された文字の異体字」設定を適用 / Apply alternate glyphs setting
            cbAlternateGlyph.value = false;

            // カンバスカラー等のUI設定を適用 / Apply UI settings
            app.preferences.setIntegerPreference("uiCanvasIsWhite", 1); // カンバス白

            // 「アニメーションズーム」設定を適用 / Apply animated zoom setting
            cbAnimZoom.value = false;
            // 「ヒストリー数」設定を適用 / Apply history steps setting
            etHistory.text = "50";
            // 「リアルタイムの描画と編集」設定を適用 / Apply real-time editing setting
            cbLiveEdit.value = true;

            // 「オリジナルの編集にシステムデフォルトを使用」設定を適用 / Apply edit original setting
            cbEditOriginal.value = true;
            // 「Adobe Fontsを自動アクティベート」設定を適用 / Apply auto-activate Adobe Fonts setting
            cbFontsAuto.value = true;
        }

        ddPreset.onChange = function() {
            if (ddPreset.selection && ddPreset.selection.text) {
                if (ddPreset.selection.text === "プリセット1") {
                    applyPreset1Settings();
                } else if (ddPreset.selection.text === "デフォルト") {
                    applyDefaultSettings();
                    app.redraw();
                }
            }
        };

        dropdownPreset.onChange = function() {
            if (dropdownPreset.selection.text === "デフォルト") {
                applyDefaultSettings();
            } else if (dropdownPreset.selection.text === "プリセット1") {
                applyPreset1Settings();
            }
            updateUIFromPreferences(); // ここでUI更新を実行
        };
    } catch (e) {
        alert("環境設定ダイアログの作成に失敗しました: " + e + "\nFailed to create preferences dialog: " + e);
        return;
    } finally {
        // メイングループの下にプリセット選択グループを追加 / Add preset selection group below main group
        mainGroup.add(groupPreset);
        // プリセット選択グループの下にプリセット選択ラベルを追加 / Add preset selection label below preset selection group
        groupPreset.add("statictext", undefined, "プリセットを選択してください / Please select a preset");
        // メイングループの下にプリセット選択ドロップダウンを追加 / Add preset selection dropdown below preset selection group
        groupPreset.add(ddPreset);
    }

    // デフォルトプリセットに現状の環境設定を取得 / Get current preferences for default preset
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

    // 2カラム用グループを作成 / Create group for two columns
    var colGroup = mainGroup.add("group");
    colGroup.orientation = "row";
    colGroup.alignChildren = "top";

    var colLeft = colGroup.add("group");
    colLeft.orientation = "column";
    colLeft.alignChildren = "left";

    var colRight = colGroup.add("group");
    colRight.orientation = "column";
    colRight.alignChildren = "left";

    // ［一般］パネル（左カラム）/ [General] panel (left column)
    var panelGeneral = colLeft.add("panel", undefined, "［一般］");
    panelGeneral.orientation = "column";
    panelGeneral.alignChildren = "left";
    panelGeneral.margins = [15, 25, 15, 10]; // 上、右、下、左のマージン

    // 「詳細なツールヒント」チェックボックスを作成 / Create rich tooltips checkbox
    var cbToolTips = panelGeneral.add("checkbox", undefined, "詳細なツールヒント");
    cbToolTips.helpTip = "詳細なツールヒントを表示 / Show detailed tooltips";
    cbToolTips.value = false; // デフォルト OFF / Default: OFF

    // 「ホーム画面」チェックボックスを作成 / Create Home Screen checkbox
    var cbHomeScreen = panelGeneral.add("checkbox", undefined, "ホーム画面");
    cbHomeScreen.helpTip = "ドキュメントを開いていないときにホーム画面を表示 / Show Home Screen when no document is open";
    cbHomeScreen.value = true; // デフォルト ON / Default: ON

    // 「以前の新規ドキュメント」チェックボックスを作成 / Create legacy new document checkbox
    var cbLegacyNewDoc = panelGeneral.add("checkbox", undefined, "以前の「新規ドキュメント」");
    cbLegacyNewDoc.helpTip = "以前の「新規ドキュメント」インターフェイスを使用 / Use legacy New Document interface";
    cbLegacyNewDoc.value = true; // デフォルト ON / Default: ON

    // 「裁ち落としを印刷」生成AIボタンチェックボックスを作成 / Create Bleed AI button checkbox
    var cbBleedAI = panelGeneral.add("checkbox", undefined, "「裁ち落としを印刷」生成AIボタン");
    cbBleedAI.helpTip = "裁ち落とし部分に「裁ち落としを印刷」生成AIボタンを表示 / Show AI button for bleed printing";
    cbBleedAI.value = false; // デフォルト OFF / Default: OFF

    // ［選択範囲・アンカー表示］パネル（左カラム）/ [Selection & Anchor Display] panel (left column)
    var panelSelectAnchor = colLeft.add("panel", undefined, "［選択範囲・アンカー表示］");
    panelSelectAnchor.orientation = "column";
    panelSelectAnchor.alignChildren = "left";
    panelSelectAnchor.margins = [15, 25, 15, 10];

    // 「アートボードと一緒に移動」チェックボックスを作成 / Create move locked objects with artboard checkbox
    var cbMoveLocked = panelSelectAnchor.add("checkbox", undefined, "アートボードと一緒に移動");
    cbMoveLocked.helpTip = "ロックまたは非表示オブジェクトをアートボードと一緒に移動 / Move locked or hidden objects with artboard";
    cbMoveLocked.value = true; // デフォルト ON / Default: ON

    // 「オブジェクトの選択範囲をパスに制限」チェックボックスを作成 / Create limit selection area to path checkbox
    var cbHitShape = panelSelectAnchor.add("checkbox", undefined, "オブジェクトの選択範囲をパスに制限");
    cbHitShape.helpTip = "オブジェクトの選択範囲をパスに制限 / Limit selection area to path";
    cbHitShape.value = false; // デフォルト OFF / Default: OFF

    // 「選択範囲へズーム」チェックボックスを作成 / Create zoom to selection checkbox
    var cbZoomToSel = panelSelectAnchor.add("checkbox", undefined, "選択範囲へズーム");
    cbZoomToSel.helpTip = "選択範囲へズーム / Zoom to selection";
    cbZoomToSel.value = false; // デフォルト OFF / Default: OFF

    // 「アンカーを強調表示」チェックボックスを作成 / Create highlight anchor checkbox
    var cbHighlightAnchor = panelSelectAnchor.add("checkbox", undefined, "アンカーを強調表示");
    cbHighlightAnchor.helpTip = "カーソルを合わせたときにアンカーを強調表示 / Highlight anchor when cursor is over";
    cbHighlightAnchor.value = true; // デフォルト ON / Default: ON



    // テキスト関連UIを右カラムの最上部にパネルとして追加 / Add text panel to top of right column
    var panelTextRight = colLeft.add("panel", undefined, "［テキスト］");
    panelTextRight.orientation = "column";
    panelTextRight.alignChildren = "left";
    panelTextRight.margins = [15, 25, 15, 10];

    // 「テキストをパスに制限」チェックボックスを作成 / Create limit selection area of text objects to path checkbox
    var cbHitTypeShape = panelTextRight.add("checkbox", undefined, "テキストをパスに制限");
    cbHitTypeShape.helpTip = "テキストオブジェクトの選択範囲をパスに制限 / Limit selection area of text objects to path";
    cbHitTypeShape.value = false; // デフォルト OFF / Default: OFF

    // 「新規エリア内文字の自動サイズ調整」チェックボックスを作成 / Create auto size new area text checkbox
    var cbAutoSizing = panelTextRight.add("checkbox", undefined, "新規エリア内文字の自動サイズ調整");
    cbAutoSizing.helpTip = "新規エリア内文字の自動サイズ調整 / Auto size new area text";
    cbAutoSizing.value = true; // デフォルト ON / Default: ON

    // 最近使用したフォント数のUIを作成 / Create UI for recent fonts count
    var currentRecentCount = app.preferences.getIntegerPreference("text/recentFontMenu/showNEntries");
    var groupRecentFonts = panelTextRight.add("group");
    groupRecentFonts.orientation = "row";

    // 「最近使用したフォント」チェックボックスを作成 / Create recent fonts checkbox
    var cbRecentFonts = groupRecentFonts.add("checkbox", undefined, "最近使用したフォント");
    cbRecentFonts.helpTip = "最近使用したフォントの表示数 / Number of recent fonts to show";
    cbRecentFonts.value = (currentRecentCount > 0);

    // 「最近使用したフォント数」入力欄を作成 / Create recent fonts count edittext
    var etRecentFonts = groupRecentFonts.add("edittext", undefined, currentRecentCount.toString());
    etRecentFonts.characters = 3;
    etRecentFonts.enabled = cbRecentFonts.value;

    cbRecentFonts.onClick = function() {
        // 「最近使用したフォント」チェックボックスの状態に応じて入力欄を有効化 / Enable or disable edittext based on checkbox
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

    // 「見つからない字形の保護」チェックボックスを作成 / Create font locking checkbox
    var cbFontLocking = panelTextRight.add("checkbox", undefined, "見つからない字形の保護");
    cbFontLocking.helpTip = "見つからない字形の保護を有効にする / Enable protection for missing glyphs";
    cbFontLocking.value = false; // デフォルト OFF / Default: OFF

    // 「選択された文字の異体字」チェックボックスを作成 / Create alternate glyphs checkbox
    var cbAlternateGlyph = panelTextRight.add("checkbox", undefined, "選択された文字の異体字");
    cbAlternateGlyph.helpTip = "選択された文字の異体字を表示 / Show alternate glyphs for selected character";
    cbAlternateGlyph.value = false; // デフォルト OFF / Default: OFF


    // 例: ダイアログ内に UI パネルを作成する
    var panelUI = colRight.add("panel", undefined, "ユーザーインターフェイス");
    panelUI.orientation = "column";
    panelUI.alignChildren = "left";
    panelUI.margins = [15, 25, 15, 10]; // 上、右、下、左のマージン

    // カンバスカラー設定 / Canvas Color
    panelUI.add("statictext", undefined, "カンバスカラー");

    var rbCanvasMatch = panelUI.add("radiobutton", undefined, "UIに合わせる");
    var rbCanvasWhite = panelUI.add("radiobutton", undefined, "ホワイト");

    // 初期状態を反映 / Reflect current setting
    var currentCanvas = app.preferences.getIntegerPreference("uiCanvasIsWhite");
    rbCanvasWhite.value = (currentCanvas === 1);
    rbCanvasMatch.value = (currentCanvas !== 1);

    // ラジオボタン変更時の設定 / Apply on selection
    rbCanvasMatch.onClick = function() {
        app.preferences.setIntegerPreference("uiCanvasIsWhite", 0);
    };
    rbCanvasWhite.onClick = function() {
        app.preferences.setIntegerPreference("uiCanvasIsWhite", 1);
    };


    // ［パフォーマンス］パネル（右カラム）/ [Performance] panel (right column)
    var panelPerf = colRight.add("panel", undefined, "［パフォーマンス］");
    panelPerf.orientation = "column";
    panelPerf.alignChildren = "left";
    panelPerf.margins = [15, 25, 15, 10];

    // 「アニメーションズーム」チェックボックスを作成 / Create animated zoom checkbox
    var cbAnimZoom = panelPerf.add("checkbox", undefined, "アニメーションズーム");
    cbAnimZoom.helpTip = "アニメーションズーム / Animated zoom";
    cbAnimZoom.value = false; // デフォルト OFF / Default: OFF

    // 「ヒストリー数」入力欄を作成 / Create history steps edittext
    var groupHistory = panelPerf.add("group");
    groupHistory.add("statictext", undefined, "ヒストリー数：");
    var etHistory = groupHistory.add("edittext", undefined, "50");
    etHistory.characters = 4;
    etHistory.helpTip = "ヒストリー数を設定 / Set history steps";

    // 「リアルタイムの描画と編集」チェックボックスを作成 / Create real-time drawing and editing checkbox
    var cbLiveEdit = panelPerf.add("checkbox", undefined, "リアルタイムの描画と編集");
    cbLiveEdit.helpTip = "リアルタイムの描画と編集 / Real-time drawing and editing";
    cbLiveEdit.value = true; // デフォルト ON / Default: ON

    // ［ファイル管理］パネル（右カラム）/ [File Management] panel (right column)
    var panelFile = colRight.add("panel", undefined, "［ファイル管理］");
    panelFile.orientation = "column";
    panelFile.alignChildren = "left";
    panelFile.margins = [15, 25, 15, 10];

    // 「オリジナルの編集にシステムデフォルトを使用」チェックボックスを作成 / Create edit original checkbox
    var cbEditOriginal = panelFile.add("checkbox", undefined, "システムデフォルトで編集");
    cbEditOriginal.helpTip = "「オリジナルの編集」にシステムデフォルトを使用 / Use system default for 'Edit Original'";
    cbEditOriginal.value = false;

    // 「Adobe Fontsを自動アクティベート」チェックボックスを作成 / Create auto-activate Adobe Fonts checkbox
    var cbFontsAuto = panelFile.add("checkbox", undefined, "Adobe Fontsを自動アクティベート");
    cbFontsAuto.helpTip = "Adobe Fontsを自動アクティベート / Auto-activate Adobe Fonts";
    cbFontsAuto.value = false;



    // applyDefaultSettings 内
    app.preferences.setIntegerPreference("uiCanvasIsWhite", 0); // Match UI

    // applyPreset1Settings 内
    app.preferences.setIntegerPreference("uiCanvasIsWhite", 1); // White

    // ボタン行を mainGroup の下部に追加 / Add button row at bottom of mainGroup
    var groupBtns = mainGroup.add("group");
    groupBtns.alignment = "center";
    var btnCancel = groupBtns.add("button", undefined, "キャンセル", {
        name: "cancel"
    });
    var btnOK = groupBtns.add("button", undefined, "OK", {
        name: "ok"
    });

    btnOK.onClick = function() {
        try {
            // 「詳細なツールヒント」設定を保存 / Save rich tooltips setting
            app.preferences.setBooleanPreference("showRichToolTips", cbToolTips.value);
            // 「ホーム画面」設定を保存 / Save Home Screen setting
            app.preferences.setBooleanPreference("Hello/ShowHomeScreenWS", cbHomeScreen.value);
            // 「以前の新規ドキュメント」設定を保存 / Save legacy new document setting
            app.preferences.setBooleanPreference("Hello/NewDoc", cbLegacyNewDoc.value);
            // 「裁ち落としを印刷」生成AIボタン設定を保存 / Save Bleed AI button setting
            app.preferences.setBooleanPreference("enablePrintBleedWidget", cbBleedAI.value);
            // 「アートボードと一緒に移動」設定を保存 / Save move locked objects with artboard setting
            app.preferences.setBooleanPreference("moveLockedAndHiddenArt", cbMoveLocked.value);
            // 「オブジェクトの選択範囲をパスに制限」設定を保存 / Save limit selection area to path setting
            app.preferences.setBooleanPreference("hitShapeOnPreview", cbHitShape.value);
            // 「選択範囲へズーム」設定を保存 / Save zoom to selection setting
            app.preferences.setBooleanPreference("zoomToSelection", cbZoomToSel.value);
            // 「アンカーを強調表示」設定を保存 / Save highlight anchor setting
            // app.preferences.setBooleanPreference("highlightAnchorsOnMouseOver", cbHighlightAnchor.value); // キーが不明な場合はコメント化
            // 「テキストをパスに制限」設定を保存 / Save limit selection area of text objects to path setting
            app.preferences.setBooleanPreference("hitTypeShapeOnPreview", cbHitTypeShape.value);
            // 「新規エリア内文字の自動サイズ調整」設定を保存 / Save auto size new area text setting
            app.preferences.setBooleanPreference("text/autoSizing", cbAutoSizing.value);
            // 「最近使用したフォント」設定を保存 / Save recent fonts setting
            if (cbRecentFonts.value) {
                app.preferences.setIntegerPreference("text/recentFontMenu/showNEntries", parseInt(etRecentFonts.text, 10));
            } else {
                app.preferences.setIntegerPreference("text/recentFontMenu/showNEntries", 0);
            }
            // 「見つからない字形の保護」設定を保存 / Save font locking setting
            app.preferences.setBooleanPreference("text/doFontLocking", cbFontLocking.value);
            // 「選択された文字の異体字」設定を保存 / Save alternate glyphs setting
            app.preferences.setBooleanPreference("text/enableAlternateGlyph", cbAlternateGlyph.value);
            // カンバス白設定を保存 / Save canvas white setting
            app.preferences.setIntegerPreference("uiCanvasIsWhite", 1);
            // 「アニメーションズーム」設定を保存 / Save animated zoom setting
            app.preferences.setBooleanPreference("Performance/AnimZoom", cbAnimZoom.value);
            // 「ヒストリー数」設定を保存 / Save history steps setting
            app.preferences.setIntegerPreference("maximumUndoDepth", parseInt(etHistory.text, 10));
            // 「リアルタイムの描画と編集」設定を保存 / Save real-time editing setting
            app.preferences.setBooleanPreference("LiveEdit_State_Machine", cbLiveEdit.value);
            // 「オリジナルの編集にシステムデフォルトを使用」設定を保存 / Save edit original setting
            app.preferences.setBooleanPreference("enableBackgroundExport", cbEditOriginal.value);
            // 「Adobe Fontsを自動アクティベート」設定を保存 / Save auto-activate Adobe Fonts setting
            app.preferences.setBooleanPreference("AutoActivateMissingFont", cbFontsAuto.value);
        } catch (e) {
            alert("環境設定の保存に失敗しました: " + e + "\nFailed to save preferences: " + e);
        }
        dlg.close();
    };
    dlg.show();
}

main();
// デフォルト設定を適用 / Apply default settings
function applyDefaultSettings() {
    app.preferences.setBooleanPreference("showRichToolTips", true);
    app.preferences.setBooleanPreference("Hello/ShowHomeScreenWS", true);
    app.preferences.setBooleanPreference("Hello/NewDoc", false);
    app.preferences.setBooleanPreference("enablePrintBleedWidget", true);
    app.preferences.setBooleanPreference("moveLockedAndHiddenArtWithArtboard", false);
    app.preferences.setBooleanPreference("selectionLimitsToBounds", false);
    app.preferences.setBooleanPreference("zoomToSelection", true);
    app.preferences.setBooleanPreference("highlightAnchorsOnMouseOver", true);

    app.preferences.setBooleanPreference("text/selectPath", false);
    app.preferences.setBooleanPreference("text/autoSizing", false);
    app.preferences.setIntegerPreference("text/recentFontsCount", 10);
    app.preferences.setBooleanPreference("text/protectMissingGlyphs", true);
    app.preferences.setBooleanPreference("text/showAlternateGlyphs", true);

    app.preferences.setIntegerPreference("uiCanvasIsWhite", 0);
    app.preferences.setBooleanPreference("gpuPerformance/animatedZoom", true);
    app.preferences.setIntegerPreference("undo/undoMaximum", 100);
    app.preferences.setBooleanPreference("realTimeDrawing", true);
    app.preferences.setBooleanPreference("useSystemDefaultEditor", false);
    app.preferences.setBooleanPreference("enableTypekitAutoActivate", false);
}