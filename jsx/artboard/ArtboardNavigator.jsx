#target illustrator
#targetengine "artboardNavigatorPalette"

// 外部 JSX 実行時の警告ダイアログを抑制 / Suppress the external-JSX warning dialog
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
  ArtboardNavigator.jsx
  アートボード間をなめらかにズーム移動する Illustrator 用パレット。

  ### ナビゲーションボタン（左から）

  ・|<  … 最初のアートボードへ（先頭にいるときは無効・グレーアウト）
  ・<   … 前のアートボードへ（端で循環、Ctrl+Shift+Opt+←）
  ・■■ … 全アートボードを画面に収めて表示（Ctrl+Shift+Opt+↑）
  ・>   … 次のアートボードへ（端で循環、Ctrl+Shift+Opt+→）
  ・>|  … 最後のアートボードへ（最終にいるときは無効・グレーアウト）
  ・ボタンはクリック中だけ明るさが少し変化（押下フィードバック）

  ### オプションパネル

  ・アニメーション … OFF で補間せず瞬時に切り替え（以降の項目はディム表示）
  ・スピード       … 移動アニメーションの速さ（右ほど速い＝ステップ数を削減）
  ・イーズアウト   … ON で easeOut の加減速、OFF で等速移動
  ・Preziライクモード … 前後移動の中盤で一度ズームを引いてから寄せる
  ・俯瞰の強さ        … Prezi ライクモードで引くズーム量（0〜1、中央 0.5）
  ・アートボード名を表示 … ON のときだけ移動先にラベルを描画（OFF は描画もレイヤー作成もしない）

  ### アートボード一覧

  ・チェックボックスで一覧の表示／非表示を切り替え（OFF はパレットを縮めて領域も確保しない）
  ・一覧（番号｜名前）… 行をクリックでそのアートボードへ移動
  ・パレットはリサイズ可能（一覧が追従して広がる）

  ### アートボードラベル（「アートボード名を表示」ON のとき）

  ・個別のアートボードへ移動した直後、その左上に「番号：アートボード名」を表示
  ・テキストは白／HiraginoSans-W6、背景は黒の長方形
  ・文字サイズ・背景／文字の不透明度はスクリプト先頭の変数で調整可能
  ・専用レイヤー "ArtboardNavigator"（ロックON・プリントOFF）に描画し、切替ごとに作り直す
  ・「アートボード名を表示」OFF・全体表示・パレットを閉じるときは、このレイヤーごと削除
  ・一覧の選択はクリック時にパレット側で確定（BridgeTalk の onResult に依存しない）

  目標ズームは view.bounds から数式で算出し、開始前のタイムラグを抑える。
  実際のビュー操作・アニメーションは BridgeTalk でメインエンジンへ送って実行する。

  ### 謝辞
　
  古島佑起さん
  BridgeTalk のワーカー登録と呼び出しの仕組み、アートボードラベルの描画方法など、多くのアイデアとコードを提供していただきました。
  https://note.com/yukifurushima/n/n9f2078dc156f

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ArtboardNavigator";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.5";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {
    // =========================================
    // ローカライズ / Localization
    // =========================================

    var currentLanguage = ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";

    // =========================================
    // ラベル定義 / Labels
    // =========================================

    var LABELS = {
        dialog: {
            title: { ja: "アートボードナビゲーター", en: "Artboard Navigator" }
        },
        alert: {
            needIllustrator: { ja: "Adobe Illustratorで実行してください。", en: "Please run this in Adobe Illustrator." },
            needArtboards: { ja: "アートボードが2つ以上あるドキュメントで実行してください。", en: "Please run this on a document with two or more artboards." }
        },
        tooltip: {
            first: { ja: "最初のアートボードへ", en: "First artboard" },
            prev: { ja: "前のアートボードへ (Ctrl+Shift+Opt+←)", en: "Previous artboard (Ctrl+Shift+Opt+←)" },
            list: { ja: "全アートボードを一覧表示 (Ctrl+Shift+Opt+↑)", en: "Fit all artboards (Ctrl+Shift+Opt+↑)" },
            next: { ja: "次のアートボードへ (Ctrl+Shift+Opt+→)", en: "Next artboard (Ctrl+Shift+Opt+→)" },
            last: { ja: "最後のアートボードへ", en: "Last artboard" }
        },
        panel: {
            options: { ja: "オプション", en: "Options" }
        },
        checkbox: {
            animation: { ja: "アニメーション", en: "Animation" },
            ease: { ja: "イーズアウト", en: "Ease Out" },
            prezi: { ja: "Preziライクモード", en: "Prezi-like mode" },
            showLabel: { ja: "アートボード名を表示", en: "Show artboard name" },
            listVisible: { ja: "アートボード一覧", en: "Artboard list" }
        },
        slider: {
            speed: { ja: "スピード", en: "Speed" },
            preziDip: { ja: "俯瞰の強さ", en: "Zoom-out amount" }
        },
        column: {
            number: { ja: "番号", en: "No." },
            name: { ja: "アートボード名", en: "Artboard name" }
        }
    };

    /* ドット区切りのキー（例 "checkbox.showLabel"）でカテゴリ分けされた LABELS を引く / Look up a categorized label by dot-separated key */
    // 末尾の {slash} はスラッシュ "/" に置換する（ソース内に裸の "/" を書かないため）
    function L(keyPath) {
        var node = LABELS;
        var parts = keyPath.split(".");
        for (var i = 0; i < parts.length; i++) {
            if (!node || node[parts[i]] === undefined) {
                return keyPath;
            }
            node = node[parts[i]];
        }
        var text = (node && node[currentLanguage] !== undefined) ? node[currentLanguage] : keyPath;
        return String(text).replace(/\{slash\}/g, "/");
    }

    if (app.name !== "Adobe Illustrator") {
        alert(L("alert.needIllustrator"));
        return;
    }

    // 単一アートボード（または2つ未満）のときはパレットを出さずに終了
    // Exit without showing the palette unless there are 2+ artboards
    if (app.documents.length === 0 || app.activeDocument.artboards.length < 2) {
        alert(L("alert.needArtboards"));
        return;
    }

    var SCRIPT_NAME = L("dialog.title") + " " + SCRIPT_VERSION;
    var PREF_FILE = new File(Folder.userData + "/ArtboardNavigator/palette-position.txt");
    var SETTINGS_FILE = new File(Folder.userData + "/ArtboardNavigator/settings.txt");

    // 前回終了時の設定（無ければ既定値で埋める）
    var savedSettings = loadSettings();

    // アニメーション設定
    // スピードは 1（遅い）〜10（速い）。速いほどステップ数を減らして素早く移動する。
    var MIN_SPEED = 1;
    var MAX_SPEED = 10;
    var DEFAULT_SPEED = 6;
    var SLOW_STEP_COUNT = 40; // 最も遅いときのステップ数
    var FAST_STEP_COUNT = 6;  // 最も速いときのステップ数

    var initialSpeed = settingNumber(savedSettings, "speed", DEFAULT_SPEED, MIN_SPEED, MAX_SPEED);
    var animationStepCount = speedToStepCount(initialSpeed);
    var frameDelayMs = 6;

    // Prezi モードで途中に下げるズーム比率（0=引かない 〜 1=最大、中央 0.5）
    var preziDipRatio = settingNumber(savedSettings, "preziDip", 0.2, 0, 1);

    // アートボードラベルの見た目（ここで自由に調整できる） / Tunable label appearance
    var LABEL_BACKGROUND_OPACITY = 80; // 背景の長方形の不透明度（%）
    var LABEL_TEXT_OPACITY = 100;      // アートボード名（文字）の不透明度（%）
    var LABEL_FONT_RATIO = 0.03;       // 文字サイズ＝アートボード幅 × この比率

    // ボタンの配色は環境設定「ユーザーインターフェイスの明るさ」に追従
    // 明るいUI = 白バック・枠線・黒記号 / 暗いUI = 黒バック・白記号（従来）
    var useLightButtons = isLightUI();

    /* スピード値（1〜10）をステップ数に変換 / Convert a speed value (1-10) to a step count */
    function speedToStepCount(speed) {
        var ratio = (speed - MIN_SPEED) / (MAX_SPEED - MIN_SPEED); // 0〜1
        var steps = Math.round(SLOW_STEP_COUNT - ratio * (SLOW_STEP_COUNT - FAST_STEP_COUNT));
        if (steps < FAST_STEP_COUNT) { steps = FAST_STEP_COUNT; }
        if (steps > SLOW_STEP_COUNT) { steps = SLOW_STEP_COUNT; }
        return steps;
    }

    // 画面いっぱいに対する余白率（92% に収める）
    var fitMarginRatio = 0.92;

    // BridgeTalk: ナビゲーション本体をメインエンジンへ登録済みか
    var workerInstalled = false;

    // すでに開いているパレットがあれば閉じてから作り直す（多重表示・再実行対策）
    try {
        if ($.global.artboardNavigatorWindow) {
            $.global.artboardNavigatorWindow.close();
            $.global.artboardNavigatorWindow = null;
        }
    } catch (e) {
    }

    if (!$.global.artboardNavigatorJobs) {
        $.global.artboardNavigatorJobs = [];
    }


    // ラベルは worker（メインエンジン）が専用レイヤー（ロックON・プリントOFF）へ描画し、
    // 切替ごとに作り直す。OFF・全体表示・閉じる時はレイヤーごと削除する。

    var navButtonDefinitions = [
        { tip: L("tooltip.first"), icon: "first", command: "first" },
        { tip: L("tooltip.prev"), icon: "prev", command: "prev" },
        { tip: L("tooltip.list"), icon: "list", command: "list" },
        { tip: L("tooltip.next"), icon: "next", command: "next" },
        { tip: L("tooltip.last"), icon: "last", command: "last" }
    ];

    // コマンド名 → ボタンの対応（先頭／最終ボタンの有効・無効切替に使う）
    var navButtonsByCommand = {};

    // =========================================================
    // パレット
    // =========================================================

    var paletteWindow = new Window("palette", SCRIPT_NAME, undefined, { resizeable: true });
    $.global.artboardNavigatorWindow = paletteWindow;

    paletteWindow.orientation = "column";
    paletteWindow.alignChildren = ["fill", "top"];
    paletteWindow.margins = 12;
    paletteWindow.spacing = 8;
    paletteWindow.preferredSize = [210, 200];

    var navButtonRow = paletteWindow.add("group");
    navButtonRow.orientation = "row";
    navButtonRow.alignChildren = ["center", "center"];
    navButtonRow.alignment = ["center", "top"];
    navButtonRow.spacing = 7;
    navButtonRow.margins = [5, 5, 5, 10]; // 下だけ +5

    for (var i = 0; i < navButtonDefinitions.length; i++) {
        addNavButton(navButtonRow, navButtonDefinitions[i]);
    }

    // オプション設定パネル / Options panel
    var optionsPanel = paletteWindow.add("panel", undefined, L("panel.options"));
    optionsPanel.orientation = "column";
    optionsPanel.alignChildren = ["left", "top"];
    optionsPanel.alignment = ["fill", "top"];
    optionsPanel.margins = [16, 20, 16, 12];

    // アニメーションの ON/OFF（OFF で以降のオプションをディム表示）
    var animationCheckbox = optionsPanel.add("checkbox", undefined, L("checkbox.animation"));
    animationCheckbox.value = settingBool(savedSettings, "animation", true);

    // 前後のアートボードへ移動するスピード（右ほど速い） / Speed of moving to the prev/next artboard (faster toward the right)
    var speedRow = optionsPanel.add("group");
    speedRow.orientation = "row";
    speedRow.alignChildren = ["left", "center"];
    speedRow.alignment = ["fill", "top"];

    var speedLabel = speedRow.add("statictext", undefined, L("slider.speed"));
    var speedSlider = speedRow.add("slider", undefined, initialSpeed, MIN_SPEED, MAX_SPEED);
    speedSlider.alignment = ["fill", "center"];
    speedSlider.preferredSize = [120, -1];

    // ドラッグ中はステップ数へ即反映、確定時（離したとき）に設定を保存
    speedSlider.onChanging = function () {
        animationStepCount = speedToStepCount(this.value);
    };
    speedSlider.onChange = function () {
        animationStepCount = speedToStepCount(this.value);
        saveSettings();
    };

    // イーズの ON/OFF（OFF で線形＝イーズなし）
    var easeCheckbox = optionsPanel.add("checkbox", undefined, L("checkbox.ease"));
    easeCheckbox.value = settingBool(savedSettings, "ease", true);
    easeCheckbox.onClick = function () {
        saveSettings();
    };

    // Prezi モード：前後移動の途中で一度ズームを下げてから寄せる（Prezi 風）
    var preziCheckbox = optionsPanel.add("checkbox", undefined, L("checkbox.prezi"));
    preziCheckbox.value = settingBool(savedSettings, "prezi", true);

    // Prezi モードで下げるズーム量（中央 0.5、左ほど弱く右ほど強い） / Prezi zoom-out amount
    var preziDipRow = optionsPanel.add("group");
    preziDipRow.orientation = "row";
    preziDipRow.alignChildren = ["left", "center"];
    preziDipRow.alignment = ["fill", "top"];

    var preziDipLabel = preziDipRow.add("statictext", undefined, L("slider.preziDip"));
    var preziDipSlider = preziDipRow.add("slider", undefined, preziDipRatio, 0, 1);
    preziDipSlider.alignment = ["fill", "center"];
    preziDipSlider.preferredSize = [120, -1];

    // ドラッグ中は引き量へ即反映、確定時（離したとき）に設定を保存
    preziDipSlider.onChanging = function () {
        preziDipRatio = this.value;
    };
    preziDipSlider.onChange = function () {
        preziDipRatio = this.value;
        saveSettings();
    };

    // 引きスライダーはアニメーション ON かつ Prezi モード ON のときだけ有効
    function applyPreziDipEnabled() {
        preziDipRow.enabled = animationCheckbox.value && preziCheckbox.value;
    }

    preziCheckbox.onClick = function () {
        applyPreziDipEnabled();
        saveSettings();
    };

    // アニメーション OFF のときは以降のオプションをすべてディム表示
    function applyAnimationEnabled(enabled) {
        speedRow.enabled = enabled;
        easeCheckbox.enabled = enabled;
        preziCheckbox.enabled = enabled;
        applyPreziDipEnabled();
    }

    animationCheckbox.onClick = function () {
        applyAnimationEnabled(this.value);
        saveSettings();
    };

    applyAnimationEnabled(animationCheckbox.value);

    // アートボード名ラベルの表示 ON/OFF
    // ON のときだけ、移動先の左上に「番号：アートボード名」を専用レイヤー
    //（ロックON・プリントOFF）へ描画する。OFF のときは描画もレイヤー作成もしない。
    var showArtboardNameCheckbox = optionsPanel.add("checkbox", undefined, L("checkbox.showLabel"));
    showArtboardNameCheckbox.value = settingBool(savedSettings, "showLabel", true);
    showArtboardNameCheckbox.onClick = function () {
        if (!this.value) {
            removeLabelLayerDirect();
            requestLabelRemoval();
        }
        saveSettings();
    };

    // アートボード一覧の表示 ON/OFF（OFF で一覧を隠し、領域も確保しない）
    var listVisibilityRow = paletteWindow.add("group");
    listVisibilityRow.orientation = "row";
    listVisibilityRow.alignChildren = ["center", "center"];
    listVisibilityRow.alignment = ["fill", "top"];
    listVisibilityRow.margins = 5;

    var listVisibilityCheckbox = listVisibilityRow.add("checkbox", undefined, L("checkbox.listVisible"));
    listVisibilityCheckbox.value = settingBool(savedSettings, "listVisible", true);

    // アートボード一覧（番号｜アートボード名）。行を選ぶとそのアートボードへ移動
    var artboardListBox = paletteWindow.add("listbox", undefined, [], {
        numberOfColumns: 2,
        showHeaders: true,
        columnTitles: [L("column.number"), L("column.name")],
        columnWidths: [44, 156]
    });
    // リサイズ時に一覧が縦横とも追従して広がるようにする
    artboardListBox.alignment = ["fill", "fill"];
    artboardListBox.preferredSize = [210, 150];

    // ウィンドウのリサイズに追従してレイアウトを再配置
    paletteWindow.onResizing = paletteWindow.onResize = function () {
        this.layout.resize();
    };

    // 一覧の表示／非表示を切り替え、OFF のときは領域も確保しない
    listVisibilityCheckbox.onClick = function () {
        applyArtboardListVisibility(this.value);
        saveSettings();
    };

    var isPopulatingList = false;

    artboardListBox.onChange = function () {
        if (isPopulatingList) {
            return;
        }
        if (this.selection) {
            runNavigation(String(this.selection.index));
        }
    };

    restoreWindowPosition(paletteWindow);

    paletteWindow.onMove = function () {
        saveWindowPosition(paletteWindow);
    };

    paletteWindow.onClose = function () {
        saveWindowPosition(paletteWindow);
        saveSettings();
        // 表示中のラベルをメインエンジン側で消してから閉じる
        requestLabelRemoval();
        $.global.artboardNavigatorWindow = null;
    };

    paletteWindow.onActivate = function () {
        // 環境設定のUI明るさが変わっていたらボタンの配色を切り替えて描き直す
        var nowLight = isLightUI();
        if (nowLight !== useLightButtons) {
            useLightButtons = nowLight;
            try {
                for (var bi = 0; bi < navButtonRow.children.length; bi++) {
                    navButtonRow.children[bi].notify("onDraw");
                }
            } catch (redrawError) {
            }
        }
        refreshArtboardList();
    };

    // キーボードショートカット（パレットにフォーカスがあるとき）
    // Control + Shift + Option + ← … 前のアートボードへ
    // Control + Shift + Option + → … 次のアートボードへ
    // Control + Shift + Option + ↑ … 全アートボードを一覧表示
    // ※ Option は ScriptUI では altKey
    try {
        paletteWindow.addEventListener("keydown", function (kbEvent) {
            if (!kbEvent.ctrlKey || !kbEvent.shiftKey || !kbEvent.altKey) {
                return;
            }
            if (kbEvent.keyName === "Left") {
                runNavigation("prev");
                kbEvent.preventDefault();
            } else if (kbEvent.keyName === "Right") {
                runNavigation("next");
                kbEvent.preventDefault();
            } else if (kbEvent.keyName === "Up") {
                runNavigation("list");
                kbEvent.preventDefault();
            }
        });
    } catch (e) {
    }

    paletteWindow.show();
    refreshArtboardList();

    // 前回 OFF だった場合は一覧を畳んだ状態で開く（show 後に高さを詰める）
    if (!listVisibilityCheckbox.value) {
        applyArtboardListVisibility(false);
    }

    // 起動直後にワーカーを登録しておく（最初のクリックも短い呼び出しだけで済む）
    if (typeof BridgeTalk !== "undefined") {
        installNavigationWorker();
    }

    // =========================================================
    // ボタン生成・アイコン描画
    // =========================================================

    /* ナビゲーションボタンを1つ生成して配置 / Create and place one navigation button */
    function addNavButton(parentGroup, buttonDefinition) {
        var button = parentGroup.add("button", undefined, "");
        button.helpTip = buttonDefinition.tip;
        button.preferredSize = [26, 26];
        button.minimumSize = [26, 26];
        button.maximumSize = [26, 26];
        button.iconType = buttonDefinition.icon;
        button.navCommand = buttonDefinition.command;
        button.pressed = false;
        navButtonsByCommand[buttonDefinition.command] = button;
        button.onDraw = function () {
            drawNavButton(this);
        };
        button.onClick = function () {
            runNavigation(this.navCommand);
        };
        // クリック中だけ背景の明るさを少し変える（押下フィードバック）
        button.addEventListener("mousedown", function () {
            this.pressed = true;
            this.notify("onDraw");
        });
        button.addEventListener("mouseup", function () {
            this.pressed = false;
            this.notify("onDraw");
        });
        // ボタン外でマウスを離したときも押下状態を戻す
        button.addEventListener("mouseout", function () {
            if (this.pressed) {
                this.pressed = false;
                this.notify("onDraw");
            }
        });
    }

    /* 環境設定のUI明るさが明るい側かどうか / Whether the UI brightness preference is on the light side */
    // uiBrightness は 0（最暗）〜1（最明）の 4 段階。
    // 0=暗 / 0.5=やや暗め → 暗いUI、0.51=やや明るめ / 1=明るい → 明るいUI。
    // 「やや暗め(0.5)」を暗い側に含めるため 0.5 より大きいかで判定する。
    function isLightUI() {
        try {
            return app.preferences.getRealPreference("uiBrightness") > 0.5;
        } catch (e) {
            return false;
        }
    }

    /* ボタンの背景とアイコンを描画 / Draw the button background and icon */
    // 明るいUI = 白バック＋枠線＋黒記号 / 暗いUI = 黒バック＋白記号
    function drawNavButton(button) {
        var graphics = button.graphics;
        var width = button.size[0];
        var height = button.size[1];

        var backgroundColor = useLightButtons ? [1, 1, 1, 1] : [0.30, 0.30, 0.30, 1];
        var glyphColor = useLightButtons ? [0.15, 0.15, 0.15, 1] : [0.92, 0.92, 0.92, 1];

        // クリック中は明るさを少しだけ変える（明るいUIは少し暗く、暗いUIは少し明るく）
        if (button.pressed) {
            backgroundColor = useLightButtons ? [0.86, 0.86, 0.86, 1] : [0.46, 0.46, 0.46, 1];
        }

        // 無効（端にいる先頭／最終ボタン）は記号を淡くしてグレーアウト表示
        if (button.enabled === false) {
            glyphColor = useLightButtons ? [0.75, 0.75, 0.75, 1] : [0.50, 0.50, 0.50, 1];
        }

        try {
            graphics.rectPath(0, 0, width, height);
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, backgroundColor));
            // 明るいUIでは白バックが背景に溶けるので枠線を足す
            if (useLightButtons) {
                graphics.rectPath(0, 0, width, height);
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, [0.62, 0.62, 0.62, 1], 1));
            }
        } catch (e1) {
            try {
                graphics.drawOSControl();
            } catch (e2) {
            }
        }

        drawNavGlyph(graphics, button.iconType, width, height, glyphColor);
    }

    /* < / > / グリッド / |< / >| のアイコン形状を描く / Draw the glyph */
    function drawNavGlyph(graphics, iconType, width, height, glyphColor) {

        if (iconType === "first") {
            // |< 形状（左端の縦棒＋左向き山）
            var firstPen = graphics.newPen(graphics.PenType.SOLID_COLOR, glyphColor, 1.6);
            graphics.newPath();
            graphics.moveTo(9, 7);
            graphics.lineTo(9, 19);
            graphics.strokePath(firstPen);
            graphics.newPath();
            graphics.moveTo(18, 7);
            graphics.lineTo(13, 13);
            graphics.lineTo(18, 19);
            graphics.strokePath(firstPen);
            return;
        }

        if (iconType === "last") {
            // >| 形状（右向き山＋右端の縦棒）
            var lastPen = graphics.newPen(graphics.PenType.SOLID_COLOR, glyphColor, 1.6);
            graphics.newPath();
            graphics.moveTo(8, 7);
            graphics.lineTo(13, 13);
            graphics.lineTo(8, 19);
            graphics.strokePath(lastPen);
            graphics.newPath();
            graphics.moveTo(17, 7);
            graphics.lineTo(17, 19);
            graphics.strokePath(lastPen);
            return;
        }

        if (iconType === "prev") {
            // < 形状
            var prevPen = graphics.newPen(graphics.PenType.SOLID_COLOR, glyphColor, 1.6);
            graphics.newPath();
            graphics.moveTo(16, 7);
            graphics.lineTo(10, 13);
            graphics.lineTo(16, 19);
            graphics.strokePath(prevPen);
            return;
        }

        if (iconType === "next") {
            // > 形状
            var nextPen = graphics.newPen(graphics.PenType.SOLID_COLOR, glyphColor, 1.6);
            graphics.newPath();
            graphics.moveTo(10, 7);
            graphics.lineTo(16, 13);
            graphics.lineTo(10, 19);
            graphics.strokePath(nextPen);
            return;
        }

        if (iconType === "list") {
            // 2×2 のサムネイル風グリッド
            var listPen = graphics.newPen(graphics.PenType.SOLID_COLOR, glyphColor, 1.0);
            var cellSize = 6;
            var cellOrigins = [[5, 5], [15, 5], [5, 15], [15, 15]];
            for (var i = 0; i < cellOrigins.length; i++) {
                // セルごとに newPath() で前のパス（背景・枠の矩形を含む）をクリアしてから描く。
                // これを省くとボタン外枠の矩形が残ってグリフ色で stroke され、枠が濃く見える。
                graphics.newPath();
                graphics.rectPath(cellOrigins[i][0], cellOrigins[i][1], cellSize, cellSize);
                graphics.strokePath(listPen);
            }
        }
    }

    // =========================================================
    // BridgeTalk でメインエンジンへ処理を送る
    // =========================================================

    /* ナビゲーションを実行（必要ならワーカーを登録） / Run navigation (install worker if needed) */
    // listbox 選択はクリック時にパレット側で確定（onResult に依存しない）。
    // ラベル描画と削除は worker 側で行う。OFF にしたときはパレット側でも直接削除する。
    function runNavigation(navCommand) {
        if (typeof BridgeTalk === "undefined") {
            return;
        }
        if (!workerInstalled) {
            installNavigationWorker();
        }

        if (navCommand === "list") {
            sendNavigation("list", true);
            return;
        }

        // パレット側で移動先の絶対インデックスを確定（listbox の選択を現在地とみなす）
        var targetIndex = resolveTargetIndex(navCommand);
        if (targetIndex >= 0) {
            // 一覧の選択を即反映し、ワーカーには絶対インデックスを渡して両者を一致させる
            selectArtboardInList(targetIndex);
            sendNavigation(String(targetIndex), true);
        } else {
            // 解決できなければ元のコマンドのままワーカーへ委ねる
            sendNavigation(navCommand, true);
        }
    }

    /* ナビゲーションコマンドを移動先の絶対インデックスへ解決（list は -1） / Resolve a command to an absolute artboard index (-1 for list) */
    function resolveTargetIndex(navCommand) {
        var count = artboardListBox.items.length;
        if (count === 0) {
            try {
                count = app.activeDocument.artboards.length;
            } catch (e) {
                return -1;
            }
        }
        if (navCommand === "list") {
            return -1;
        }
        var current = currentListIndex();
        if (navCommand === "prev") {
            return (current - 1 + count) % count;
        }
        if (navCommand === "next") {
            return (current + 1) % count;
        }
        if (navCommand === "first") {
            return 0;
        }
        if (navCommand === "last") {
            return count - 1;
        }
        var explicitIndex = parseInt(navCommand, 10);
        if (!isNaN(explicitIndex) && explicitIndex >= 0 && explicitIndex < count) {
            return explicitIndex;
        }
        return -1;
    }

    /* 現在地とみなすインデックス（listbox の選択 → 無ければドキュメント） / Current index (listbox selection, fallback to document) */
    function currentListIndex() {
        if (artboardListBox.selection) {
            return artboardListBox.selection.index;
        }
        try {
            return app.activeDocument.artboards.getActiveArtboardIndex();
        } catch (e) {
            return 0;
        }
    }

    /* onChange を発火させずに一覧の選択を更新 / Select a list row without triggering onChange */
    function selectArtboardInList(index) {
        if (index < 0 || index >= artboardListBox.items.length) {
            return;
        }
        isPopulatingList = true;
        try {
            artboardListBox.selection = index;
        } catch (e) {
        }
        isPopulatingList = false;
        updateNavButtonStates(index);
    }

    /* 先頭にいるとき先頭ボタン、最終にいるとき最終ボタンを無効化 / Disable first/last button at the respective ends */
    // prev/next は循環移動なので常に有効のまま。
    function updateNavButtonStates(currentIndex) {
        var count = artboardListBox.items.length;
        if (count === 0) {
            try {
                count = app.activeDocument.artboards.length;
            } catch (e) {
                return;
            }
        }
        if (typeof currentIndex !== "number" || currentIndex < 0) {
            currentIndex = currentListIndex();
        }
        setButtonEnabled(navButtonsByCommand.first, currentIndex > 0);
        setButtonEnabled(navButtonsByCommand.last, currentIndex < count - 1);
    }

    /* ボタンの有効状態を変え、必要なら再描画 / Set a button's enabled state and repaint if it changed */
    function setButtonEnabled(button, enabled) {
        if (!button || button.enabled === enabled) {
            return;
        }
        button.enabled = enabled;
        try {
            button.notify("onDraw");
        } catch (e) {
        }
    }

    /* ナビゲーション本体をメインエンジンへ一度だけ登録 / Install the worker into the main engine once */
    // 以降のクリックは短い呼び出しだけ送るので、毎回の送信ペイロードが最小になる。
    function installNavigationWorker() {
        var bridgeTalk = newIllustratorBridgeTalk();
        bridgeTalk.body = "$.global.artboardNavigatorWorker = (" + navigationScriptMain.toString() + ");";
        bridgeTalk.onResult = function () {
            removeNavigationJob(bridgeTalk);
        };
        bridgeTalk.onError = function () {
            removeNavigationJob(bridgeTalk);
        };
        $.global.artboardNavigatorJobs.push(bridgeTalk);
        trimNavigationJobs();
        bridgeTalk.send();
        workerInstalled = true;
    }

    /* 登録済みワーカーを短い呼び出しだけで実行 / Invoke the installed worker with a short call */
    function sendNavigation(navCommand, allowReinstall) {
        var bridgeTalk = newIllustratorBridgeTalk();
        bridgeTalk.body = buildNavigationCall(navCommand);
        bridgeTalk.onResult = function () {
            removeNavigationJob(bridgeTalk);
        };
        bridgeTalk.onError = function () {
            removeNavigationJob(bridgeTalk);
            // ワーカー未登録などで失敗したら、登録し直して1回だけ再試行
            if (allowReinstall) {
                workerInstalled = false;
                installNavigationWorker();
                sendNavigation(navCommand, false);
            }
        };
        $.global.artboardNavigatorJobs.push(bridgeTalk);
        trimNavigationJobs();
        bridgeTalk.send();
    }

    /* メインエンジンへラベル即時消去を依頼（パレットを閉じるときなど） / Ask the main engine to remove labels now (e.g. on close) */
    function requestLabelRemoval() {
        if (typeof BridgeTalk === "undefined") {
            return;
        }
        try {
            var bridgeTalk = newIllustratorBridgeTalk();
            bridgeTalk.body = "if($.global.artboardNavigatorRemoveLabels)$.global.artboardNavigatorRemoveLabels();";
            // 送信完了前に GC されないよう、他の送信と同じジョブ配列で保持する
            bridgeTalk.onResult = function () {
                removeNavigationJob(bridgeTalk);
            };
            bridgeTalk.onError = function () {
                removeNavigationJob(bridgeTalk);
            };
            $.global.artboardNavigatorJobs.push(bridgeTalk);
            trimNavigationJobs();
            bridgeTalk.send();
        } catch (e) {
        }
    }

    /* パレット側から "ArtboardNavigator" レイヤーを直接削除 / Remove the "ArtboardNavigator" layer directly from the palette engine */
    // requestLabelRemoval（メインエンジン経由）は worker 関数が未定義だと無効なため、
    // 「アートボード名を表示」OFF の瞬間に確実に消すための、BridgeTalk に依存しない直接削除。
    // 前回セッションから書類に残った専用レイヤーも、未ナビゲートのまま OFF にすれば消える。
    function removeLabelLayerDirect() {
        if (app.documents.length === 0) {
            return;
        }
        try {
            var doc = app.activeDocument;
            for (var i = doc.layers.length - 1; i >= 0; i--) {
                if (doc.layers[i].name === "ArtboardNavigator") {
                    try {
                        doc.layers[i].locked = false;
                        doc.layers[i].remove();
                    } catch (layerError) {
                    }
                }
            }
            app.redraw();
        } catch (e) {
        }
    }

    /* Illustrator 宛ての BridgeTalk を生成 / Create a BridgeTalk addressed to Illustrator */
    function newIllustratorBridgeTalk() {
        var bridgeTalk = new BridgeTalk();
        try {
            bridgeTalk.target = BridgeTalk.getSpecifier("illustrator");
        } catch (e) {
            bridgeTalk.target = "illustrator";
        }
        return bridgeTalk;
    }

    /* アートボード一覧を再構築し現在のものを選択 / Rebuild the artboard list and select the active one */
    // 先に値を取得し、取得できたときだけ作り直す（失敗時は現状維持で空白化を防ぐ）
    function refreshArtboardList() {
        if (app.documents.length === 0) {
            return;
        }

        var activeIndex;
        var artboardNames = [];

        try {
            var doc = app.activeDocument;
            activeIndex = doc.artboards.getActiveArtboardIndex();
            for (var i = 0; i < doc.artboards.length; i++) {
                artboardNames.push(doc.artboards[i].name);
            }
        } catch (e) {
            // ドキュメントが一時的に参照できない場合は現状維持
            return;
        }

        if (artboardNames.length === 0) {
            return;
        }

        isPopulatingList = true;
        try {
            artboardListBox.removeAll();

            for (var j = 0; j < artboardNames.length; j++) {
                var listItem = artboardListBox.add("item", String(j + 1));
                listItem.subItems[0].text = artboardNames[j];
            }

            if (activeIndex >= 0 && activeIndex < artboardListBox.items.length) {
                artboardListBox.selection = activeIndex;
            }
        } catch (e2) {
        } finally {
            isPopulatingList = false;
        }

        // 先頭／最終ボタンの有効・無効を現在地に合わせて更新
        updateNavButtonStates(activeIndex);
    }

    // 一覧表示時のウィンドウ高さ（再表示でこの値へ正確に戻す）
    var shownWindowHeight = null;

    /* 一覧の表示／非表示を切り替え（OFF は領域も確保せずパレットを縮める） / Toggle the list and shrink/grow the palette by its height */
    // OFF 時は layout(true) で残りの要素にぴったり収まる高さへ詰め、ON 時は記録した高さへ正確に戻す
    function applyArtboardListVisibility(showList) {
        var keptWidth = paletteWindow.size[0];

        if (showList) {
            artboardListBox.visible = true;
            // 固定 preferredSize を解除してから再計算（解除しないと縮まらない）
            paletteWindow.preferredSize = [-1, -1];
            paletteWindow.layout.layout(true);
            // 幅は維持し、高さは OFF 前の値へ戻す（手動リサイズも保持）
            paletteWindow.size = [keptWidth, (shownWindowHeight !== null) ? shownWindowHeight : paletteWindow.size[1]];
            shownWindowHeight = null;
            paletteWindow.layout.resize();
        } else {
            // 戻すべき高さを記録してから、一覧を隠した状態の実高さへ詰める
            shownWindowHeight = paletteWindow.size[1];
            artboardListBox.visible = false;

            paletteWindow.preferredSize = [-1, -1];
            paletteWindow.layout.layout(true);

            var collapsedHeight = paletteWindow.bounds.height;
            // alert(collapsedHeight);
            paletteWindow.size = [keptWidth, collapsedHeight];
            paletteWindow.layout.resize();
        }
    }

    /* 登録済みワーカーの呼び出し文を組み立てる / Build the call string for the installed worker */
    function buildNavigationCall(navCommand) {
        return "$.global.artboardNavigatorWorker(" +
            quoteForScript(navCommand) + "," +
            animationStepCount + "," +
            frameDelayMs + "," +
            fitMarginRatio + "," +
            (easeCheckbox.value ? "true" : "false") + "," +
            (preziCheckbox.value ? "true" : "false") + "," +
            (animationCheckbox.value ? "true" : "false") + "," +
            preziDipRatio + "," +
            LABEL_FONT_RATIO + "," +
            LABEL_BACKGROUND_OPACITY + "," +
            LABEL_TEXT_OPACITY + "," +
            (showArtboardNameCheckbox.value ? "true" : "false") +
            ");";
    }

    /* 文字列を安全に引用符で囲む / Quote a string safely for embedding in script */
    function quoteForScript(value) {
        return '"' + String(value).replace(/\\/g, "\\\\").replace(/"/g, '\\"') + '"';
    }

    /* 保留中の BridgeTalk ジョブを一定数に保つ / Keep pending BridgeTalk jobs under a limit */
    function trimNavigationJobs() {
        var pendingJobs = $.global.artboardNavigatorJobs;
        while (pendingJobs.length > 12) {
            pendingJobs.shift();
        }
    }

    /* 完了したジョブを配列から取り除く / Remove a finished job from the array */
    function removeNavigationJob(bridgeTalk) {
        var pendingJobs = $.global.artboardNavigatorJobs;
        for (var i = pendingJobs.length - 1; i >= 0; i--) {
            if (pendingJobs[i] === bridgeTalk) {
                pendingJobs.splice(i, 1);
            }
        }
    }

    // =========================================================
    // メインエンジンで実行される本体（BridgeTalk で送信）
    // ※ この関数は文字列化して送るため、外側の変数を参照せず
    //   引数とアプリ DOM だけで完結させる
    // =========================================================

    /* メインエンジンで移動を実行する本体 / Worker that performs the move in the main engine */
    function navigationScriptMain(navCommand, stepCount, delayMs, marginRatio, useEasing, preziMode, animate, preziDipRatio, labelFontRatio, labelBackgroundOpacity, labelTextOpacity, showLabel) {
        try {
            if (!app.documents || app.documents.length < 1) {
                return "ドキュメントが開かれていません。";
            }

            var doc = app.activeDocument;
            if (doc.artboards.length < 2) {
                return "アートボードが2つ以上必要です。";
            }

            var view = doc.views[0];
            var targetView;

            // Prezi 風のズームダウンは個別アートボードへの移動時のみ（全体表示では使わない）
            var usePreziDip = false;

            // 移動先のラベルを出すアートボード番号（-1 = ラベルを出さない＝全体表示）
            var labelIndex = -1;

            // ラベルは専用レイヤー "ArtboardNavigator"（ロックON・プリントOFF）に描画する。
            var LABEL_LAYER_NAME = "ArtboardNavigator";
            var LABEL_ITEM_NAME = "__ArtboardNavigatorLabel";

            // パレットからの即時消去依頼（チェックOFF・パレットを閉じる時など）に応える削除関数。
            // 専用レイヤーごと削除し、レイヤーも残さない。文字列から呼べるよう $.global に公開。
            if (typeof $.global.artboardNavigatorRemoveLabels !== "function") {
                $.global.artboardNavigatorRemoveLabels = function () {
                    try {
                        if (app.documents.length === 0) {
                            return;
                        }
                        var labelDoc = app.activeDocument;
                        for (var k = labelDoc.layers.length - 1; k >= 0; k--) {
                            if (labelDoc.layers[k].name === "ArtboardNavigator") {
                                try {
                                    labelDoc.layers[k].locked = false;
                                    labelDoc.layers[k].remove();
                                } catch (layerError) {
                                }
                            }
                        }
                        app.redraw();
                    } catch (removeError) {
                    }
                };
            }

            if (navCommand === "list") {
                targetView = measureAllArtboardsView();
            } else {
                var artboardCount = doc.artboards.length;
                var currentIndex = doc.artboards.getActiveArtboardIndex();
                var targetIndex;

                if (navCommand === "prev") {
                    targetIndex = (currentIndex - 1 + artboardCount) % artboardCount;
                } else if (navCommand === "next") {
                    targetIndex = (currentIndex + 1) % artboardCount;
                } else if (navCommand === "first") {
                    targetIndex = 0;
                } else if (navCommand === "last") {
                    targetIndex = artboardCount - 1;
                } else {
                    // 一覧から選ばれた絶対インデックス（数値文字列）
                    targetIndex = parseInt(navCommand, 10);
                    if (isNaN(targetIndex) || targetIndex < 0 || targetIndex >= artboardCount) {
                        return "範囲外のアートボード: " + navCommand;
                    }
                }

                targetView = measureArtboardView(targetIndex);
                doc.artboards.setActiveArtboardIndex(targetIndex);
                usePreziDip = preziMode;
                labelIndex = targetIndex;
            }

            animateViewTo(targetView.centerX, targetView.centerY, targetView.zoom, usePreziDip);

            // 移動後に左上ラベルを更新。
            // ・「アートボード名を表示」OFF、または全体表示のときは描画せず専用レイヤーごと削除
            // ・ラベル描画が失敗しても移動自体は成立させる
            try {
                if (showLabel && labelIndex >= 0) {
                    showArtboardLabel(labelIndex);
                } else {
                    removeLabelLayer();
                    app.redraw();
                }
            } catch (labelError) {
            }

            return "ok";

            // ----- 以下、メインエンジン内のヘルパー（巻き上げで先に定義扱い） -----
            // 役割ごとに ①アニメーション ②ビュー計測 ③ラベルレイヤー操作 ④ラベル処理 に分かれる。

            // ===== ① アニメーション / Animation =====

            /* 現在のビューから目標へなめらかに移動 / Animate the view from current to target */
            function animateViewTo(targetCenterX, targetCenterY, targetZoom, preziDip) {
                // アニメーション OFF のときは補間せず一気に目標へジャンプ
                if (!animate) {
                    view.centerPoint = [targetCenterX, targetCenterY];
                    view.zoom = targetZoom;
                    app.redraw();
                    return;
                }

                var startCenterPoint = view.centerPoint;
                var startCenterX = startCenterPoint[0];
                var startCenterY = startCenterPoint[1];
                var startZoom = view.zoom;

                // 移動量に応じてステップ数を調整（小さい移動は少ないステップで素早く）
                var effectiveSteps = resolveStepCount(
                    startCenterX, startCenterY, startZoom,
                    targetCenterX, targetCenterY, targetZoom
                );

                for (var i = 1; i <= effectiveSteps; i++) {
                    var progress = i / effectiveSteps;

                    // useEasing が OFF のときは線形（イーズなし）
                    // ON のときは easeOutQuad（出だしは速く、終わりにゆっくり減速）
                    // ＝押した瞬間に動き出すので開始のタイムラグを感じにくい
                    var easedProgress = useEasing
                        ? 1 - (1 - progress) * (1 - progress)
                        : progress;

                    var frameCenterX = startCenterX + (targetCenterX - startCenterX) * easedProgress;
                    var frameCenterY = startCenterY + (targetCenterY - startCenterY) * easedProgress;
                    var frameZoom = startZoom + (targetZoom - startZoom) * easedProgress;

                    // Prezi モード：移動の中盤でズームを最大 preziDipRatio ぶん下げ、
                    // 出入りは sin カーブで滑らかに（始点・終点では 1.0 に戻る）
                    if (preziDip) {
                        var dipMultiplier = 1 - preziDipRatio * Math.sin(progress * Math.PI);
                        if (dipMultiplier < 0.05) { dipMultiplier = 0.05; } // ズーム 0 を回避
                        frameZoom = frameZoom * dipMultiplier;
                    }

                    view.centerPoint = [frameCenterX, frameCenterY];
                    view.zoom = frameZoom;

                    app.redraw();
                    sleep(delayMs);
                }

                // 最後に正確に合わせる
                view.centerPoint = [targetCenterX, targetCenterY];
                view.zoom = targetZoom;
                app.redraw();
            }

            /* 移動量に応じてステップ数を決める / Decide step count from the move magnitude */
            // 画面に対するパン量とズーム変化率の大きい方を minSteps〜stepCount に対応づける
            function resolveStepCount(startX, startY, startZoom, targetX, targetY, targetZoom) {
                var magnitude;

                try {
                    var bounds = view.bounds;
                    // bounds = [left, top, right, bottom]（現在ズームでのドキュメント座標）
                    var visibleExtent = Math.max(
                        Math.abs(bounds[2] - bounds[0]),
                        Math.abs(bounds[1] - bounds[3])
                    );

                    var moveDistance = Math.sqrt((targetX - startX) * (targetX - startX) + (targetY - startY) * (targetY - startY));
                    var moveFraction = (visibleExtent > 0) ? (moveDistance / visibleExtent) : 1;

                    var maxZoom = Math.max(startZoom, targetZoom);
                    var zoomFraction = (maxZoom > 0) ? (Math.abs(targetZoom - startZoom) / maxZoom) : 0;

                    magnitude = Math.max(moveFraction, zoomFraction);
                } catch (e) {
                    // 計測できなければフルステップ
                    return stepCount;
                }

                if (magnitude > 1) {
                    magnitude = 1;
                }

                var minSteps = Math.max(2, Math.round(stepCount * 0.2));
                var resolvedSteps = Math.round(stepCount * magnitude);

                if (resolvedSteps < minSteps) {
                    resolvedSteps = minSteps;
                }
                if (resolvedSteps > stepCount) {
                    resolvedSteps = stepCount;
                }

                return resolvedSteps;
            }

            // ===== ② ビュー計測（中心・ズームの算出）/ View metrics =====

            /* 指定アートボードの中心とズームを求める / Compute center and zoom for one artboard */
            function measureArtboardView(artboardIndex) {
                var rect = doc.artboards[artboardIndex].artboardRect;
                // rect = [left, top, right, bottom]（Illustrator は y 上向きで top > bottom）

                return {
                    centerX: (rect[0] + rect[2]) / 2,
                    centerY: (rect[1] + rect[3]) / 2,
                    zoom: computeFitZoom(rect[2] - rect[0], rect[1] - rect[3], artboardIndex)
                };
            }

            /* 全アートボードを収める中心とズームを求める / Compute center and zoom to fit all artboards */
            function measureAllArtboardsView() {
                var left = null, top = null, right = null, bottom = null;

                for (var i = 0; i < doc.artboards.length; i++) {
                    var rect = doc.artboards[i].artboardRect;
                    if (left === null || rect[0] < left) { left = rect[0]; }
                    if (top === null || rect[1] > top) { top = rect[1]; }
                    if (right === null || rect[2] > right) { right = rect[2]; }
                    if (bottom === null || rect[3] < bottom) { bottom = rect[3]; }
                }

                return {
                    centerX: (left + right) / 2,
                    centerY: (top + bottom) / 2,
                    zoom: computeFitZoom(right - left, top - bottom, "all")
                };
            }

            /* 対象を収めるズームを view.bounds から算出 / Compute fit zoom from view.bounds */
            // メニューコマンドもビュー変更も使わないため、移動開始前のタイムラグが出ない。
            function computeFitZoom(targetWidth, targetHeight, fallbackTarget) {
                try {
                    var bounds = view.bounds;
                    // bounds = [left, top, right, bottom]（ドキュメント座標）
                    var visibleWidth = Math.abs(bounds[2] - bounds[0]);
                    var visibleHeight = Math.abs(bounds[1] - bounds[3]);

                    if (visibleWidth > 0 && visibleHeight > 0 && targetWidth > 0 && targetHeight > 0) {
                        var currentZoom = view.zoom;
                        var fitZoom = Math.min(
                            currentZoom * visibleWidth / targetWidth,
                            currentZoom * visibleHeight / targetHeight
                        );
                        return fitZoom * marginRatio;
                    }
                } catch (e) {
                }

                // フォールバック: 実際に fitin / fitall して計測（環境差対策）
                return measureFitZoomByMenu(fallbackTarget);
            }

            /* fitin/fitall で実測するフォールバック / Fallback that measures via fitin/fitall */
            function measureFitZoomByMenu(fallbackTarget) {
                var beforeCenterPoint = view.centerPoint;
                var beforeZoom = view.zoom;

                if (fallbackTarget === "all") {
                    app.executeMenuCommand("fitall");
                } else {
                    doc.artboards.setActiveArtboardIndex(fallbackTarget);
                    app.executeMenuCommand("fitin");
                }

                var fitZoom = view.zoom;

                view.centerPoint = beforeCenterPoint;
                view.zoom = beforeZoom;

                return fitZoom * marginRatio;
            }

            /* 指定ミリ秒だけ待機（同期、アニメーションのフレーム間隔用） / Busy-wait for the given milliseconds (animation frame delay) */
            function sleep(durationMs) {
                var sleepStartTime = new Date().getTime();
                while (new Date().getTime() - sleepStartTime < durationMs) { }
            }

            // ===== ③ ラベルレイヤー操作 / Label layer =====

            /* ラベル専用レイヤーを探す（無ければ null） / Find the dedicated label layer (null if absent) */
            function findLabelLayer() {
                for (var i = 0; i < doc.layers.length; i++) {
                    if (doc.layers[i].name === LABEL_LAYER_NAME) {
                        return doc.layers[i];
                    }
                }
                return null;
            }


            /* ラベル専用レイヤーごと削除（レイヤーも残さない） / Remove the dedicated label layer entirely */
            function removeLabelLayer() {
                var layer = findLabelLayer();
                if (!layer) {
                    return;
                }
                try {
                    layer.locked = false;
                    layer.remove();
                } catch (removeError) {
                }
            }

            // ===== ④ ラベル処理（描画）/ Label drawing =====

            /* ドキュメントのカラースペースに合わせて白／黒を作る / Build white or black for the document color space */
            function makeLabelColor(isWhite) {
                if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
                    var cmyk = new CMYKColor();
                    cmyk.cyan = 0;
                    cmyk.magenta = 0;
                    cmyk.yellow = 0;
                    cmyk.black = isWhite ? 0 : 100;
                    return cmyk;
                }
                var rgb = new RGBColor();
                var value = isWhite ? 255 : 0;
                rgb.red = value;
                rgb.green = value;
                rgb.blue = value;
                return rgb;
            }

            /* アートボード左上に「番号：名前」ラベルを専用レイヤー（ロックON・プリントOFF）へ描画 / Draw the "index: name" label on a dedicated locked, non-printing layer */
            function showArtboardLabel(artboardIndex) {
                var artboard = doc.artboards[artboardIndex];
                var artboardRect = artboard.artboardRect; // [left, top, right, bottom]
                var artboardLeft = artboardRect[0];
                var artboardTop = artboardRect[1];
                var artboardWidth = artboardRect[2] - artboardRect[0];

                // 文字サイズはアートボード幅に比例（画面フィット後の見た目をほぼ一定に）
                var fontSize = artboardWidth * labelFontRatio;
                if (fontSize < 1) {
                    fontSize = 1;
                }

                var paddingX = fontSize * 0.45;
                var paddingY = fontSize * 0.30;
                var edgeMargin = fontSize * 0.6;

                var labelText = (artboardIndex + 1) + "：" + artboard.name;

                // 専用レイヤーを作り直す（位置・内容をリセットし、常に1枚に保つ）。
                // 作業レイヤーを巻き込まないよう、描画後にアクティブレイヤーを元へ戻す。
                var prevActiveLayer = doc.activeLayer;
                removeLabelLayer();
                var labelLayer = doc.layers.add();
                labelLayer.name = LABEL_LAYER_NAME;

                var labelGroup = labelLayer.groupItems.add();
                labelGroup.name = LABEL_ITEM_NAME;

                // テキスト（白・HiraginoSans-W6）
                var textFrame = labelLayer.textFrames.pointText([0, 0]);
                textFrame.contents = labelText;
                var attributes = textFrame.textRange.characterAttributes;
                try {
                    attributes.textFont = app.textFonts.getByName("HiraginoSans-W6");
                } catch (fontError) {
                }
                attributes.size = fontSize;
                attributes.fillColor = makeLabelColor(true);
                textFrame.opacity = labelTextOpacity;

                var textBounds = textFrame.geometricBounds; // [left, top, right, bottom]
                var textLeft = textBounds[0];
                var textTop = textBounds[1];
                var textWidth = textBounds[2] - textBounds[0];
                var textHeight = textBounds[1] - textBounds[3];

                // 背景の長方形（黒・不透明度は呼び出し側の設定値）
                var backgroundRect = labelLayer.pathItems.rectangle(
                    textTop + paddingY,
                    textLeft - paddingX,
                    textWidth + paddingX * 2,
                    textHeight + paddingY * 2
                );
                backgroundRect.stroked = false;
                backgroundRect.filled = true;
                backgroundRect.fillColor = makeLabelColor(false);
                backgroundRect.opacity = labelBackgroundOpacity;

                // 重ね順を明示：背景の長方形を背面（PLACEATEND）、テキストを前面（PLACEATBEGINNING）へ。
                // 挿入順に依存せず、必ずテキストが前面・長方形が背面になる。
                backgroundRect.move(labelGroup, ElementPlacement.PLACEATEND);
                textFrame.move(labelGroup, ElementPlacement.PLACEATBEGINNING);

                // アートボードの左上へ少し内側に配置
                labelGroup.position = [artboardLeft + edgeMargin, artboardTop - edgeMargin];

                // レイヤーをプリントOFF・ロックON にし、作業レイヤーを元に戻す
                labelLayer.printable = false;
                labelLayer.locked = true;
                try {
                    doc.activeLayer = prevActiveLayer;
                } catch (restoreError) {
                }

                app.redraw();
            }
        } catch (e) {
            return "エラー: " + e.message;
        }
    }

    // =========================================================
    // ウィンドウ位置の保存・復元
    // =========================================================

    /* 保存したウィンドウ位置を復元 / Restore the saved window position */
    function restoreWindowPosition(windowRef) {
        if (!PREF_FILE.exists) {
            return;
        }

        try {
            if (!PREF_FILE.open("r")) {
                return;
            }
            var savedPositionText = PREF_FILE.read();
            PREF_FILE.close();

            var positionParts = savedPositionText.split(",");
            if (positionParts.length !== 2) {
                return;
            }

            var x = parseInt(positionParts[0], 10);
            var y = parseInt(positionParts[1], 10);
            if (isNaN(x) || isNaN(y)) {
                return;
            }

            // モニタ構成の変更などで画面外（不可視）になる位置は無視し、既定位置で開く
            if (!isLocationVisible(x, y)) {
                return;
            }

            windowRef.location = [x, y];
        } catch (e) {
        }
    }

    /* 指定座標（ウィンドウ左上）がいずれかのスクリーン内で、タイトルバーを掴める範囲か / Whether the top-left point sits on a screen with the title bar reachable */
    function isLocationVisible(x, y) {
        var screens;
        try {
            screens = $.screens;
        } catch (screenError) {
            return true; // スクリーン情報を取得できない環境では従来どおり許可
        }
        if (!screens || !screens.length) {
            return true;
        }
        for (var i = 0; i < screens.length; i++) {
            var s = screens[i];
            // 左上がスクリーン内に収まり、右端・下端には掴むための余白（幅60/高さ40）を残す
            if (x >= s.left && x <= s.right - 60 && y >= s.top && y <= s.bottom - 40) {
                return true;
            }
        }
        return false;
    }

    /* 現在のウィンドウ位置を保存 / Save the current window position */
    function saveWindowPosition(windowRef) {
        try {
            var prefFolder = PREF_FILE.parent;
            if (!prefFolder.exists) {
                prefFolder.create();
            }

            if (!PREF_FILE.open("w")) {
                return;
            }
            PREF_FILE.write(windowRef.location[0] + "," + windowRef.location[1]);
            PREF_FILE.close();
        } catch (e) {
        }
    }

    // =========================================================
    // 各設定の保存・復元（key=value 形式）
    // =========================================================

    /* 設定ファイルを読み込み key=value を連想配列に / Load settings.txt into a key=value map */
    function loadSettings() {
        var result = {};
        if (!SETTINGS_FILE.exists) {
            return result;
        }

        try {
            if (!SETTINGS_FILE.open("r")) {
                return result;
            }
            var content = SETTINGS_FILE.read();
            SETTINGS_FILE.close();

            var lines = content.split(/\r\n|\r|\n/);
            for (var i = 0; i < lines.length; i++) {
                var separatorIndex = lines[i].indexOf("=");
                if (separatorIndex < 1) {
                    continue;
                }
                var key = lines[i].substring(0, separatorIndex);
                var value = lines[i].substring(separatorIndex + 1);
                result[key] = value;
            }
        } catch (e) {
        }
        return result;
    }

    /* 現在の各設定を key=value 形式で保存 / Save current settings in key=value form */
    function saveSettings() {
        try {
            var settingsFolder = SETTINGS_FILE.parent;
            if (!settingsFolder.exists) {
                settingsFolder.create();
            }

            var lines = [
                "animation=" + (animationCheckbox.value ? "1" : "0"),
                "speed=" + speedSlider.value,
                "ease=" + (easeCheckbox.value ? "1" : "0"),
                "prezi=" + (preziCheckbox.value ? "1" : "0"),
                "preziDip=" + preziDipSlider.value,
                "listVisible=" + (listVisibilityCheckbox.value ? "1" : "0"),
                "showLabel=" + (showArtboardNameCheckbox.value ? "1" : "0")
            ];

            if (!SETTINGS_FILE.open("w")) {
                return;
            }
            SETTINGS_FILE.write(lines.join("\n"));
            SETTINGS_FILE.close();
        } catch (e) {
        }
    }

    /* 設定値を真偽として取得（未保存なら既定値） / Read a setting as boolean */
    function settingBool(source, key, fallback) {
        var raw = source[key];
        if (raw === undefined) {
            return fallback;
        }
        return raw === "1" || raw === "true";
    }

    /* 設定値を数値として取得し範囲内に収める（未保存なら既定値） / Read a setting as a clamped number */
    function settingNumber(source, key, fallback, minValue, maxValue) {
        var raw = source[key];
        if (raw === undefined) {
            return fallback;
        }
        var parsed = parseFloat(raw);
        if (isNaN(parsed)) {
            return fallback;
        }
        if (typeof minValue === "number" && parsed < minValue) {
            parsed = minValue;
        }
        if (typeof maxValue === "number" && parsed > maxValue) {
            parsed = maxValue;
        }
        return parsed;
    }

})();