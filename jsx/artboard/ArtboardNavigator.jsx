#target illustrator
#targetengine "artboardNavigatorPalette"

// 外部 JSX 実行時の警告ダイアログを抑制 / Suppress the external-JSX warning dialog
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
  ArtboardNavigator.jsx
  アートボード間をなめらかにズーム移動する Illustrator 用パレット。

  ■ ナビゲーションボタン（左から）
  ・|<  … 最初のアートボードへ
  ・<   … 前のアートボードへ（端で循環、Ctrl+Shift+Opt+←）
  ・■■ … 全アートボードを画面に収めて表示（Ctrl+Shift+Opt+↑）
  ・>   … 次のアートボードへ（端で循環、Ctrl+Shift+Opt+→）
  ・>|  … 最後のアートボードへ

  ■ オプションパネル
  ・アニメーション … OFF で補間せず瞬時に切り替え（以降の項目はディム表示）
  ・スピード       … 移動アニメーションの速さ（右ほど速い＝ステップ数を削減）
  ・イーズアウト   … ON で easeOut の加減速、OFF で等速移動
  ・Preziモード    … 前後移動の中盤で一度ズームを引いてから寄せる
  ・俯瞰の強さ     … Prezi モードで引くズーム量（0〜1、中央 0.5）

  ■ アートボード一覧
  ・チェックボックスで一覧の表示／非表示を切り替え（OFF はパレットを縮めて領域も確保しない）
  ・一覧（番号｜名前）… 行をクリックでそのアートボードへ移動
  ・パレットはリサイズ可能（一覧が追従して広がる）

  目標ズームは view.bounds から数式で算出し、開始前のタイムラグを抑える。
  実際のビュー操作・アニメーションは BridgeTalk でメインエンジンへ送って実行する。
*/

// =========================================
// バージョン / Version
// =========================================

var SCRIPT_VERSION = "v1.1.0";

(function () {
    if (app.name !== "Adobe Illustrator") {
        alert("Adobe Illustratorで実行してください。");
        return;
    }

    // 単一アートボード（または2つ未満）のときはパレットを出さずに終了
    // Exit without showing the palette unless there are 2+ artboards
    if (app.documents.length === 0 || app.activeDocument.artboards.length < 2) {
        alert("アートボードが2つ以上あるドキュメントで実行してください。");
        return;
    }

    var SCRIPT_NAME = "アートボードナビゲーター " + SCRIPT_VERSION;
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

    /* スピード値（1〜10）をステップ数に変換 / Convert a speed value (1-10) to a step count */
    function speedToStepCount(speed) {
        var ratio = (speed - MIN_SPEED) / (MAX_SPEED - MIN_SPEED); // 0〜1
        var steps = Math.round(SLOW_STEP_COUNT - ratio * (SLOW_STEP_COUNT - FAST_STEP_COUNT));
        if (steps < FAST_STEP_COUNT) { steps = FAST_STEP_COUNT; }
        if (steps > SLOW_STEP_COUNT) { steps = SLOW_STEP_COUNT; }
        return steps;
    }

    // 画面いっぱいに対する余白率（90% に収める）
    var fitMarginRatio = 0.9;

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

    var navButtonDefinitions = [
        { tip: "最初のアートボードへ", icon: "first", command: "first" },
        { tip: "前のアートボードへ (Ctrl+Shift+Opt+←)", icon: "prev", command: "prev" },
        { tip: "全アートボードを一覧表示 (Ctrl+Shift+Opt+↑)", icon: "list", command: "list" },
        { tip: "次のアートボードへ (Ctrl+Shift+Opt+→)", icon: "next", command: "next" },
        { tip: "最後のアートボードへ", icon: "last", command: "last" }
    ];

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
    var optionsPanel = paletteWindow.add("panel", undefined, "オプション");
    optionsPanel.orientation = "column";
    optionsPanel.alignChildren = ["left", "top"];
    optionsPanel.alignment = ["fill", "top"];
    optionsPanel.margins = [16, 20, 16, 12];

    // アニメーションの ON/OFF（OFF で以降のオプションをディム表示）
    var animationCheckbox = optionsPanel.add("checkbox", undefined, "アニメーション");
    animationCheckbox.value = settingBool(savedSettings, "animation", true);

    // 前後のアートボードへ移動するスピード（右ほど速い） / Speed of moving to the prev/next artboard (faster toward the right)
    var speedRow = optionsPanel.add("group");
    speedRow.orientation = "row";
    speedRow.alignChildren = ["left", "center"];
    speedRow.alignment = ["fill", "top"];

    var speedLabel = speedRow.add("statictext", undefined, "スピード");
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
    var easeCheckbox = optionsPanel.add("checkbox", undefined, "イーズアウト");
    easeCheckbox.value = settingBool(savedSettings, "ease", true);
    easeCheckbox.onClick = function () {
        saveSettings();
    };

    // Prezi モード：前後移動の途中で一度ズームを下げてから寄せる（Prezi 風）
    var preziCheckbox = optionsPanel.add("checkbox", undefined, "Preziモード");
    preziCheckbox.value = settingBool(savedSettings, "prezi", true);

    // Prezi モードで下げるズーム量（中央 0.5、左ほど弱く右ほど強い） / Prezi zoom-out amount
    var preziDipRow = optionsPanel.add("group");
    preziDipRow.orientation = "row";
    preziDipRow.alignChildren = ["left", "center"];
    preziDipRow.alignment = ["fill", "top"];

    var preziDipLabel = preziDipRow.add("statictext", undefined, "俯瞰の強さ");
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

    // アートボード一覧の表示 ON/OFF（OFF で一覧を隠し、領域も確保しない）
    var listVisibilityRow = paletteWindow.add("group");
    listVisibilityRow.orientation = "row";
    listVisibilityRow.alignChildren = ["center", "center"];
    listVisibilityRow.alignment = ["fill", "top"];
    listVisibilityRow.margins = 5;

    var listVisibilityCheckbox = listVisibilityRow.add("checkbox", undefined, "アートボード一覧");
    listVisibilityCheckbox.value = settingBool(savedSettings, "listVisible", true);

    // アートボード一覧（番号｜アートボード名）。行を選ぶとそのアートボードへ移動
    var artboardListBox = paletteWindow.add("listbox", undefined, [], {
        numberOfColumns: 2,
        showHeaders: true,
        columnTitles: ["番号", "アートボード名"],
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
        $.global.artboardNavigatorWindow = null;
    };

    paletteWindow.onActivate = function () {
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
        button.onDraw = function () {
            drawNavButton(this);
        };
        button.onClick = function () {
            runNavigation(this.navCommand);
        };
    }

    /* ボタンの背景とアイコンを描画 / Draw the button background and icon */
    function drawNavButton(button) {
        var graphics = button.graphics;
        var width = button.size[0];
        var height = button.size[1];

        try {
            graphics.rectPath(0, 0, width, height);
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, [0.30, 0.30, 0.30, 1]));
        } catch (e1) {
            try {
                graphics.drawOSControl();
            } catch (e2) {
            }
        }

        drawNavGlyph(graphics, button.iconType, width, height);
    }

    /* < / > / グリッド / |< / >| のアイコン形状を描く / Draw the glyph */
    function drawNavGlyph(graphics, iconType, width, height) {
        var glyphColor = [0.92, 0.92, 0.92, 1];

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
                graphics.rectPath(cellOrigins[i][0], cellOrigins[i][1], cellSize, cellSize);
                graphics.strokePath(listPen);
            }
        }
    }

    // =========================================================
    // BridgeTalk でメインエンジンへ処理を送る
    // =========================================================

    /* ナビゲーションを実行（必要ならワーカーを登録） / Run navigation (install worker if needed) */
    function runNavigation(navCommand) {
        if (typeof BridgeTalk === "undefined") {
            return;
        }
        if (!workerInstalled) {
            installNavigationWorker();
        }
        sendNavigation(navCommand, true);
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
            // 移動後のアクティブアートボードを一覧の選択に反映
            refreshArtboardList();
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
            // 戻すべき高さを記録してから、固定高さへ詰める
            shownWindowHeight = paletteWindow.size[1];
            artboardListBox.visible = false;
            paletteWindow.size = [keptWidth, 290];
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
            preziDipRatio +
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
    function navigationScriptMain(navCommand, stepCount, delayMs, marginRatio, useEasing, preziMode, animate, preziDipRatio) {
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
            }

            animateViewTo(targetView.centerX, targetView.centerY, targetView.zoom, usePreziDip);
            return "ok";

            // ----- 以下、メインエンジン内のヘルパー（巻き上げで先に定義扱い） -----

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

            /* 指定ミリ秒だけ待機（同期） / Busy-wait for the given milliseconds */
            function sleep(durationMs) {
                var sleepStartTime = new Date().getTime();
                while (new Date().getTime() - sleepStartTime < durationMs) {}
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
            PREF_FILE.open("r");
            var savedPositionText = PREF_FILE.read();
            PREF_FILE.close();

            var positionParts = savedPositionText.split(",");
            if (positionParts.length !== 2) {
                return;
            }

            var x = parseInt(positionParts[0], 10);
            var y = parseInt(positionParts[1], 10);
            if (!isNaN(x) && !isNaN(y)) {
                windowRef.location = [x, y];
            }
        } catch (e) {
        }
    }

    /* 現在のウィンドウ位置を保存 / Save the current window position */
    function saveWindowPosition(windowRef) {
        try {
            var prefFolder = PREF_FILE.parent;
            if (!prefFolder.exists) {
                prefFolder.create();
            }

            PREF_FILE.open("w");
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
            SETTINGS_FILE.open("r");
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
                "listVisible=" + (listVisibilityCheckbox.value ? "1" : "0")
            ];

            SETTINGS_FILE.open("w");
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