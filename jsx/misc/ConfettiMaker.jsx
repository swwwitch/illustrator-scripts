#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/* =========================================
 * 紙吹雪を生成 / Generate Confetti
 *
 * 概要 / Overview
 * - 選択オブジェクトの領域を基準に、紙吹雪（円/長方形/三角形/スター/キラキラ）を生成します。
 * - ダイアログ上でプレビューを表示し、OKで確定（Confetti レイヤーに出力）します。
 * - UIは日本語/英語に対応し、タイトルバーにバージョンを表示します。
 *
 * 使い方 / Usage
 * 1) マスクにしたいオブジェクトを1つ選択
 * 2) スクリプトを実行
 * 3) 設定を調整し、OK
 *
 * メモ / Notes
 * - マージンは生成エリアのみ拡張します（マスク領域は拡張しません）。
 * - TextFrame（テキスト）選択時はマスク処理を無効化します。
 * - プレビュー用レイヤー（__ConfettiPreview__）は終了時に削除されます。
 * ========================================= */

// コンフェティ（紙吹雪）作成スクリプト

(function () {
    var SCRIPT_VERSION = "v1.0";

    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var lang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialogTitle: { ja: "紙吹雪を生成", en: "Generate Confetti" },

        pnlShape: { ja: "形状", en: "Shapes" },
        circle: { ja: "円", en: "Circle" },
        rect: { ja: "長方形", en: "Rectangle" },
        triangle: { ja: "三角形", en: "Triangle" },
        star: { ja: "スター", en: "Star" },
        sparkle: { ja: "キラキラ", en: "Sparkle" },

        pnlDist: { ja: "配置分布", en: "Distribution" },
        distEven: { ja: "全体に均等", en: "Uniform" },
        distGrad: { ja: "グラデーション（上濃→下薄）", en: "Vertical gradient (top→bottom)" },
        distHollow: { ja: "中心を空ける", en: "Hollow center" },

        pnlCount: { ja: "生成数", en: "Count" },

        pnlOption: { ja: "オプション", en: "Options" },
        mask: { ja: "マスク処理", en: "Mask" },
        margin: { ja: "マージン", en: "Margin" },
        random: { ja: "ランダム", en: "Random" },
        zoom: { ja: "ズーム", en: "Zoom" },
        opacity: { ja: "不透明度", en: "Opacity" },

        btnCancel: { ja: "キャンセル", en: "Cancel" },
        btnOK: { ja: "OK", en: "OK" }
    };

    function L(key) {
        try {
            if (LABELS[key] && LABELS[key][lang]) return LABELS[key][lang];
            if (LABELS[key] && LABELS[key].ja) return LABELS[key].ja;
        } catch (_) { }
        return key;
    }
    var doc = app.activeDocument;

    // 選択オブジェクト取得（なければアートボードを対象）
    var selectedObj = null;
    var __useArtboard = false;

    if (doc.selection.length > 0) {
        selectedObj = doc.selection[0];
    } else {
        __useArtboard = true;
    }

// テキスト選択時フラグ / Flag for text selection
var __isTextSelection = false;
try { __isTextSelection = (selectedObj && selectedObj.typename === "TextFrame"); } catch (_) { __isTextSelection = false; }
    
    // ビュー初期値（ズーム復元用） / Initial view state
    var __initView = null;
    var __initZoom = null;
    var __initCenter = null;
    try {
        __initView = doc.activeView;
        __initZoom = __initView.zoom;
        __initCenter = __initView.centerPoint;
    } catch (_) { }

    // ダイアログ作成
    var dlg = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
    dlg.orientation = "column";
    dlg.alignChildren = "left";

    // ダイアログの位置・透明度 / Dialog position & opacity
    var offsetX = 300;
    var offsetY = 0;
    var dialogOpacity = 0.98;

    function shiftDialogPosition(dlg, dx, dy) {
        try {
            var currentX = dlg.location[0];
            var currentY = dlg.location[1];
            dlg.location = [currentX + dx, currentY + dy];
        } catch (_) { }
    }

    function setDialogOpacity(dlg, opacityValue) {
        try { dlg.opacity = opacityValue; } catch (_) { }
    }

    setDialogOpacity(dlg, dialogOpacity);

    /* 上段のみ2カラム / Two-column top area */
    var gTop = dlg.add("group");
    gTop.orientation = "row";
    gTop.alignChildren = ["left", "top"];

    /* 形状 / Shapes */
    var pnlShape = gTop.add("panel", undefined, L("pnlShape"));
    pnlShape.orientation = "column";
    pnlShape.alignChildren = "left";
    pnlShape.margins = [15, 20, 15, 10];

    var chkCircle = pnlShape.add("checkbox", undefined, L("circle"));
    var chkRect = pnlShape.add("checkbox", undefined, L("rect"));
    var chkTriangle = pnlShape.add("checkbox", undefined, L("triangle"));
    var chkStar = pnlShape.add("checkbox", undefined, L("star"));
    var chkStar4 = pnlShape.add("checkbox", undefined, L("sparkle"));

    // デフォルトはすべてON
    chkCircle.value = true;
    chkRect.value = true;
    chkTriangle.value = true;
    chkStar.value = false;
    chkStar4.value = false;

    /* Option(Alt)+クリックで単独選択 / Option(Alt)+click to solo */
    function soloShape(activeChk) {
        try {
            chkCircle.value = (activeChk === chkCircle);
            chkRect.value = (activeChk === chkRect);
            chkTriangle.value = (activeChk === chkTriangle);
            chkStar.value = (activeChk === chkStar);
            chkStar4.value = (activeChk === chkStar4);
            activeChk.value = true;
        } catch (_) { }
    }

    function isAltDown() {
        try {
            return !!(ScriptUI.environment && ScriptUI.environment.keyboardState && ScriptUI.environment.keyboardState.altKey);
        } catch (_) { }
        return false;
    }

    chkCircle.onClick = function () {
        if (isAltDown()) soloShape(chkCircle);
        drawPreview();
    };
    chkRect.onClick = function () {
        if (isAltDown()) soloShape(chkRect);
        drawPreview();
    };
    chkTriangle.onClick = function () {
        if (isAltDown()) soloShape(chkTriangle);
        drawPreview();
    };
    chkStar.onClick = function () {
        if (isAltDown()) soloShape(chkStar);
        drawPreview();
    };
    chkStar4.onClick = function () {
        if (isAltDown()) soloShape(chkStar4);
        drawPreview();
    };

    /* 右カラム / Right column */
    var gRight = gTop.add("group");
    gRight.orientation = "column";
    gRight.alignChildren = "left";




    /* オプション / Options */
    var pnlOption = dlg.add("panel", undefined, L("pnlOption"));
    pnlOption.orientation = "column";
    pnlOption.alignChildren = "left";
    pnlOption.margins = [15, 20, 15, 10];

    // 不透明度 + マスク処理 / Opacity + Mask
    var gOptTop = pnlOption.add("group");
    gOptTop.orientation = "row";
    gOptTop.alignChildren = ["left", "center"];

    var chkOpacity = gOptTop.add("checkbox", undefined, L("opacity"));
    chkOpacity.value = true;

    var chkMask = gOptTop.add("checkbox", undefined, L("mask"));
    chkMask.value = true;
    // テキスト選択時・アートボード使用時はマスク不可
    try {
        if (__isTextSelection || __useArtboard) {
            chkMask.value = false;
            chkMask.enabled = false;
        }
    } catch (_) { }

    // マージン（生成エリア用）
    var gMargin = pnlOption.add("group");
    gMargin.orientation = "row";
    gMargin.alignChildren = ["left", "center"];

    var chkMargin = gMargin.add("checkbox", undefined, L("margin"));
    chkMargin.value = false;

    var sldMargin = gMargin.add("slider", undefined, 0, 0, 50);
    sldMargin.preferredSize.width = 200;
    sldMargin.enabled = chkMargin.value;

    var maskMarginPt = 0;

    // サイズランダム（□ランダム <=====>）
    var gRandSize = pnlOption.add("group");
    gRandSize.orientation = "row";
    gRandSize.alignChildren = ["left", "center"];

    var chkRandSize = gRandSize.add("checkbox", undefined, L("random"));
    chkRandSize.value = true; // デフォルトON（現状の挙動を維持）

    var sldRandSize = gRandSize.add("slider", undefined, 100, 100, 300);
    sldRandSize.preferredSize.width = 200;
    sldRandSize.enabled = chkRandSize.value;

    var randSizeStrength = 100;

    /* 配置分布 / Distribution */
    var pnlOpt = gRight.add("panel", undefined, L("pnlDist"));
    pnlOpt.orientation = "column";
    pnlOpt.alignChildren = "left";
    pnlOpt.margins = [15, 20, 15, 10];

    var rbEven = pnlOpt.add("radiobutton", undefined, L("distEven"));
    var rbGradY = pnlOpt.add("radiobutton", undefined, L("distGrad"));
    var rbHollow = pnlOpt.add("radiobutton", undefined, L("distHollow"));

    // 中心の空き具合（分布の強度）
    // 1.0 = ほぼ均等寄り / 6.0 = 強く外側へ
    var gHollow = pnlOpt.add("group");
    gHollow.orientation = "row";
    gHollow.alignChildren = ["left", "center"];

    var sldHollow = gHollow.add("slider", undefined, 2.0, 1.0, 6.0);
    sldHollow.preferredSize.width = 200;


    // 初期値
    var hollowStrength = 2.0;

    // デフォルト（今）：均等に
    rbEven.value = true;
    rbHollow.value = false;
    rbGradY.value = false;

    // 「中心をだんだんと空ける」以外では操作できない
    gHollow.enabled = rbHollow.value;


    /* 生成数 / Count */
    var pnlCount = gRight.add("panel", undefined, L("pnlCount"));
    pnlCount.orientation = "column";
    pnlCount.alignChildren = "fill";
    pnlCount.margins = [15, 20, 15, 10];

    var gCountRow = pnlCount.add("group");
    gCountRow.orientation = "row";
    gCountRow.alignChildren = ["left", "center"];

    var sldCount = gCountRow.add("slider", undefined, 150, 10, 500);
    sldCount.preferredSize.width = 200;

    var confettiCount = 150;

    sldCount.onChanging = function () {
        confettiCount = Math.round(sldCount.value);
        drawPreview();
    };

    /* ズーム / Zoom */
    var gZoom = dlg.add("group");
    gZoom.orientation = "row";
    gZoom.alignChildren = ["center", "center"];
    gZoom.alignment = "center";

    var stZoom = gZoom.add("statictext", undefined, L("zoom"));
    var sldZoom = gZoom.add("slider", undefined, (__initZoom != null ? __initZoom : 1), 0.1, 16);
    sldZoom.preferredSize.width = 260;

    function applyZoom(z) {
        try {
            if (!__initView) __initView = doc.activeView;
            if (!__initView) return;
            __initView.zoom = z;
        } catch (_) { }
    }

    sldZoom.onChanging = function () {
        applyZoom(Number(sldZoom.value));
        try { app.redraw(); } catch (_) { }
    };

    /* OK / Cancel */
    var gBtns = dlg.add("group");
    gBtns.alignment = "right";
    gBtns.alignChildren = ["right", "center"];

    var btnCancel = gBtns.add("button", undefined, L("btnCancel"), { name: "cancel" });
    var btnOK = gBtns.add("button", undefined, L("btnOK"), { name: "ok" });
    dlg.defaultElement = btnOK;

    // 有効なバウンディングボックスを取得する関数（マージン考慮）
    function getEffectiveBounds() {
        var b = null;

        try {
            if (__useArtboard) {
                var abIndex = doc.artboards.getActiveArtboardIndex();
                b = doc.artboards[abIndex].artboardRect; // [L, T, R, B]
            } else if (selectedObj) {
                b = selectedObj.geometricBounds; // [L, T, R, B]
            }
        } catch (e) {
            b = null;
        }

        if (!b) return null;

        var L = b[0], T = b[1], R = b[2], B = b[3];

        try {
            if (chkMargin && chkMargin.value && typeof maskMarginPt === "number" && maskMarginPt !== 0) {
                L -= maskMarginPt;
                T += maskMarginPt;
                R += maskMarginPt;
                B -= maskMarginPt;
            }
        } catch (_) { }

        return [L, T, R, B];
    }

    // 設定
    var minSize = 3; // 最小サイズ（mm）
    var maxSize = 8; // 最大サイズ（mm）

    // カラフルな色の配列
    var colors = [
        [255, 77, 77],    // 赤
        [255, 153, 51],   // オレンジ
        [255, 255, 51],   // 黄色
        [102, 255, 102],  // 緑
        [51, 153, 255],   // 青
        [153, 102, 255],  // 紫
        [255, 102, 178],  // ピンク
        [255, 204, 51]    // 金色
    ];

    // プレビュー用レイヤー（ダイアログ中のみ）
    var previewLayer = doc.layers.add();
    previewLayer.name = "__ConfettiPreview__";

    var previewItems = [];

    function clearPreview() {
        try {
            // 既存プレビューを完全に消す（グループ/マスク含む）
            if (previewLayer) {
                while (previewLayer.pageItems.length > 0) {
                    try { previewLayer.pageItems[0].remove(); } catch (e1) { break; }
                }
            }
        } catch (e2) { }
        previewItems = [];
        try { app.redraw(); } catch (_) { }
    }

    function drawPreview() {
        clearPreview();

        var mg = buildMaskGroup(previewLayer);
        var container = mg.container;

        var __b = getEffectiveBounds();
        if (!__b) return;

        for (var i = 0; i < confettiCount; i++) {
            var pt = pickPoint(__b);
            var x = pt.x;
            var y = pt.y;
            var size = getConfettiSize();
            var color = colors[Math.floor(Math.random() * colors.length)];
            var item = createConfetti(container, x, y, size, color);
            if (item) previewItems.push(item);
        }
        if (mg.group && !__useArtboard) {
            applyMaskToGroup(mg.group, selectedObj);
        }
        try { app.redraw(); } catch (_) { }
    }


    rbEven.onClick = function () {
        gHollow.enabled = false;
        drawPreview();
    };
    rbHollow.onClick = function () {
        gHollow.enabled = true;
        drawPreview();
    };
    rbGradY.onClick = function () {
        gHollow.enabled = false;
        drawPreview();
    };
    chkMask.onClick = function () { drawPreview(); };

    chkOpacity.onClick = function () { drawPreview(); };

    chkRandSize.onClick = function () {
        sldRandSize.enabled = chkRandSize.value;
        drawPreview();
    };

    sldRandSize.onChanging = function () {
        randSizeStrength = Math.round(sldRandSize.value);
        if (chkRandSize.value) drawPreview();
    };

    chkMargin.onClick = function () {
        sldMargin.enabled = chkMargin.value;
        drawPreview();
    };

    sldMargin.onChanging = function () {
        maskMarginPt = Math.round(sldMargin.value);
        drawPreview();
    };

    sldHollow.onChanging = function () {
        hollowStrength = Math.round(sldHollow.value * 10) / 10; // 0.1刻み表示
        if (rbHollow.value) drawPreview();
    };

    dlg.onShow = function () {
        shiftDialogPosition(dlg, offsetX, offsetY);
        try {
            if (__initView && __initZoom != null) {
                sldZoom.value = __initView.zoom;
            }
        } catch (_) { }
        drawPreview();
        try { app.redraw(); } catch (_) { }
    };

    // 位置サンプリング（分布）
    function pickPoint(bounds) {
        // bounds: [L, T, R, B]
        var L = bounds[0], T = bounds[1], R = bounds[2], B = bounds[3];
        var cx = (L + R) * 0.5;
        var cy = (T + B) * 0.5;

        var pt;

        // 全体に均等
        if (rbEven && rbEven.value) {
            pt = {
                x: random(L, R),
                y: random(B, T)
            };
            return pt;
        }

        // グラデーション（上濃→下薄）
        if (rbGradY && rbGradY.value) {
            var h = (T - B);
            var u = Math.random(); // 0..1
            // top-bias: u を上側に寄せる
            var gy = 2.2; // 固定強度（必要ならUI化）
            var t2 = 1 - Math.pow(1 - u, gy); // 0(bottom) .. 1(top)
            pt = {
                x: random(L, R),
                y: B + h * t2
            };
            return pt;
        }

        // 中心を空ける：中心から離れるほど出やすい（滑らかに外側へバイアス）
        var ang = random(0, Math.PI * 2);
        var c = Math.cos(ang);
        var s = Math.sin(ang);

        // その角度で矩形内に収まる最大半径を計算
        var maxRx = (c === 0) ? 1e12 : (c > 0 ? (R - cx) / c : (L - cx) / c);
        var maxRy = (s === 0) ? 1e12 : (s > 0 ? (T - cy) / s : (B - cy) / s);
        var maxR = Math.min(Math.abs(maxRx), Math.abs(maxRy));

        // 外側寄り＋中心を空ける（リング状）
        // hollowStrength: 1..6
        var k = (typeof hollowStrength === "number" && hollowStrength > 0) ? hollowStrength : 2.0;

        // 内側半径（中心の空き具合）: 1→0 / 6→0.65 くらい
        var hs = k;
        if (hs < 1) hs = 1;
        if (hs > 6) hs = 6;
        var innerRatio = (hs - 1) / 5; // 0..1
        innerRatio = innerRatio * 0.65;
        var rMin = maxR * innerRatio;

        // rMin..maxR の範囲で外側寄りにサンプル
        var uu = Math.random(); // 0..1
        var r01 = 1 - Math.pow(1 - uu, k); // bias to 1
        var r = rMin + (maxR - rMin) * r01;

        pt = {
            x: cx + c * r,
            y: cy + s * r
        };

        return pt;
    }

    // ランダム関数
    function random(min, max) {
        return Math.random() * (max - min) + min;
    }

    function getConfettiSize() {
        var base = (minSize + maxSize) * 0.5;

        try {
            if (!(chkRandSize && chkRandSize.value)) {
                return base;
            }
        } catch (_) { }

        var strength = (typeof randSizeStrength === "number") ? randSizeStrength : 0;
        if (strength < 100) strength = 100;
        if (strength > 300) strength = 300;

        // 100 → ほぼ固定
        // 300 → 最大約±150%揺れ
        var ratio = (strength - 100) / 200; // 0..1

        // 揺れ幅を base 基準で拡張
        var maxRange = base * 1.5 * ratio; // 最大±150%
        var v = base + random(-maxRange, maxRange);

        if (v < 0.1) v = 0.1;
        return v;
    }

    function setClippingFlag(item) {
        if (!item) return false;
        try {
            if (item.typename === "PathItem") {
                item.clipping = true;
                return true;
            }
            if (item.typename === "CompoundPathItem") {
                if (item.pathItems && item.pathItems.length > 0) {
                    item.pathItems[0].clipping = true;
                    return true;
                }
            }
            if (item.typename === "GroupItem") {
                // Pathfinder結果がGroupになる場合があるため、内側の最初のPathをクリッピングに
                try {
                    if (item.compoundPathItems && item.compoundPathItems.length > 0) {
                        var cp = item.compoundPathItems[0];
                        if (cp.pathItems && cp.pathItems.length > 0) {
                            cp.pathItems[0].clipping = true;
                            return true;
                        }
                    }
                } catch (_) { }
                try {
                    if (item.pathItems && item.pathItems.length > 0) {
                        item.pathItems[0].clipping = true;
                        return true;
                    }
                } catch (_) { }
            }
        } catch (e) { }
        return false;
    }

    function buildMaskGroup(layerOrGroup) {
        // 戻り値: { container: (GroupItem or Layer), group: GroupItem|null }
        if (!chkMask || !chkMask.value) {
            return { container: layerOrGroup, group: null };
        }
        var g = layerOrGroup.groupItems.add();
        g.name = "__ConfettiMasked__";
        return { container: g, group: g };
    }

    function createMaskShapeFromSelection(container, sourceItem) {
        // 選択オブジェクトに合わせたマスク図形を作成
        // - PathItem: 選択パスを複製してマスクに
        // - クリップグループ: クリッピングパス（Path/Compound）を複製
        // - それ以外: バウンディング矩形でフォールバック
        if (!container || !sourceItem) return null;

        // Path / CompoundPath はそのまま複製
        try {
            if (sourceItem.typename === "PathItem" || sourceItem.typename === "CompoundPathItem") {
                return sourceItem.duplicate(container, ElementPlacement.PLACEATEND);
            }
        } catch (_) { }

        // テキスト: 複製してマスクとして利用（アウトライン化しない）
        try {
            if (sourceItem.typename === "TextFrame") {
                return sourceItem.duplicate(container, ElementPlacement.PLACEATEND);
            }
        } catch (_) { }

        // クリップグループ: クリッピングパスを探して複製
        try {
            if (sourceItem.typename === "GroupItem") {
                var gi = sourceItem;
                var clipCandidate = null;

                // GroupItem.clipped が true の場合はほぼクリップグループ
                // クリッピング指定の付いた Path / Compound を優先して探す
                try {
                    // CompoundPath の方を先に探す（穴あき形状などを想定）
                    if (gi.compoundPathItems && gi.compoundPathItems.length > 0) {
                        for (var cpi = 0; cpi < gi.compoundPathItems.length; cpi++) {
                            var c = gi.compoundPathItems[cpi];
                            try {
                                if (c.pathItems && c.pathItems.length > 0 && c.pathItems[0].clipping) {
                                    clipCandidate = c;
                                    break;
                                }
                            } catch (_) { }
                        }
                    }

                    if (!clipCandidate && gi.pathItems && gi.pathItems.length > 0) {
                        for (var pi = 0; pi < gi.pathItems.length; pi++) {
                            var p = gi.pathItems[pi];
                            try {
                                if (p.clipping) {
                                    clipCandidate = p;
                                    break;
                                }
                            } catch (_) { }
                        }
                    }

                    // クリッピング指定が見つからない場合: 最初のCompound/Pathを使う（最後の手段）
                    if (!clipCandidate) {
                        if (gi.compoundPathItems && gi.compoundPathItems.length > 0) clipCandidate = gi.compoundPathItems[0];
                        else if (gi.pathItems && gi.pathItems.length > 0) clipCandidate = gi.pathItems[0];
                    }

                    if (clipCandidate) {
                        return clipCandidate.duplicate(container, ElementPlacement.PLACEATEND);
                    }

                    // クリップ候補が取れない（通常グループなど）の場合：合体して単一パス化
                    var united = uniteGroupToSinglePath(container, gi);
                    if (united) {
                        return united;
                    }
                } catch (_) { }
            }
        } catch (_) { }

        // フォールバック: バウンディング矩形
        var b;
        try {
            b = sourceItem.geometricBounds; // [L, T, R, B]
        } catch (e) {
            return null;
        }
        var L = b[0], T = b[1], R = b[2], B = b[3];

        var w = (R - L);
        var h = (T - B);
        if (!w || !h) return null;

        var m = container.pathItems.rectangle(T, L, w, h);
        // クリッピングパスは「存在」している必要があるため、塗りはON（None色）にして不可視化する
        try { m.stroked = false; } catch (_) { }
        try { m.filled = true; m.fillColor = new NoColor(); } catch (_) { }
        return m;
    }

    function uniteGroupToSinglePath(container, groupItem) {
        if (!container || !groupItem) return null;
        var dup = null;
        var oldSel = null;
        try {
            oldSel = doc.selection;
        } catch (_) { oldSel = null; }

        try {
            // 複製して container に入れる
            dup = groupItem.duplicate(container, ElementPlacement.PLACEATEND);
        } catch (e) {
            return null;
        }

        try {
            // 選択を複製グループに切り替え、合体→展開
            doc.selection = null;
            dup.selected = true;
            app.executeMenuCommand('Live Pathfinder Add');
            app.executeMenuCommand('expandStyle');

            // expandStyle後は selection が合体結果（Path/Compound/Group）になる想定
            if (doc.selection && doc.selection.length > 0) {
                return doc.selection[0];
            }
        } catch (e2) {
            // 失敗したら複製を残さない
            try { dup.remove(); } catch (_) { }
            return null;
        } finally {
            // 選択復元
            try {
                doc.selection = null;
                if (oldSel && oldSel.length) {
                    for (var i = 0; i < oldSel.length; i++) {
                        try { oldSel[i].selected = true; } catch (_) { }
                    }
                } else {
                    // もともと単一選択前提
                    try { selectedObj.selected = true; } catch (_) { }
                }
            } catch (_) { }
        }

        return null;
    }

    function applyMaskToGroup(group, sourceItem) {
        if (!group || !sourceItem) return;

        // 選択オブジェクトに合わせたマスク図形を作成（複製しない）
        var maskItem = createMaskShapeFromSelection(group, sourceItem);
        if (!maskItem) return;

        // クリッピングパスは最前面に
        try { maskItem.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (_) { }

        // グループの最前面へ（クリッピングパスは最前面が原則）
        try { maskItem.move(group, ElementPlacement.PLACEATEND); } catch (_) { }
        // move後にも最前面を保証
        try { maskItem.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (_) { }

        // TextFrame は .clipping を立てられないため、makeMask を使う
        if (maskItem && maskItem.typename === "TextFrame") {
            try {
                // グループ内の全アイテムを選択してクリッピングマスクを作成
                doc.selection = null;
                try { group.selected = true; } catch (_) { }
                app.executeMenuCommand('makeMask');
            } catch (_) { }
            return;
        }

        // クリッピングフラグ
        setClippingFlag(maskItem);

        try { group.clipped = true; } catch (_) { }
    }

    // 五芒星を生成
    function createStar(layer, left, top, size) {
        // left, top は他形状同様、バウンディングの左上基準
        // size は外接正方形の一辺
        var outerR = size * 0.5;
        var innerR = outerR * 0.5; // 五芒星の内側半径（見た目優先）
        var cx = left + outerR;
        var cy = top - outerR;

        var pts = [];
        // 上向きスタート
        var startDeg = -90;
        for (var i = 0; i < 10; i++) {
            var r = (i % 2 === 0) ? outerR : innerR;
            var ang = (startDeg + i * 36) * Math.PI / 180; // 360/10=36
            var px = cx + Math.cos(ang) * r;
            var py = cy + Math.sin(ang) * r;
            pts.push([px, py]);
        }

        var p = layer.pathItems.add();
        p.setEntirePath(pts);
        p.closed = true;
        return p;
    }

    // スター（4点）を生成
    function createStar4(layer, left, top, size) {
        // left, top は他形状同様、バウンディングの左上基準
        // size は外接正方形の一辺
        var outerR = size * 0.5;
        var innerR = outerR * 0.25; // 4点スターの内側半径（尖り強め）
        var cx = left + outerR;
        var cy = top - outerR;

        var pts = [];
        // 上向きスタート（外側→内側→…）
        var startDeg = -90;
        for (var i = 0; i < 8; i++) {
            var r = (i % 2 === 0) ? outerR : innerR;
            var ang = (startDeg + i * 45) * Math.PI / 180; // 360/8=45
            var px = cx + Math.cos(ang) * r;
            var py = cy + Math.sin(ang) * r;
            pts.push([px, py]);
        }

        var p = layer.pathItems.add();
        p.setEntirePath(pts);
        p.closed = true;
        return p;
    }

    // コンフェティの形状をランダムに作成
    function createConfetti(layer, x, y, size, color) {
        var enabledShapes = [];
        if (chkRect.value) enabledShapes.push(0);
        if (chkCircle.value) enabledShapes.push(1);
        if (chkTriangle.value) enabledShapes.push(2);
        if (chkStar.value) enabledShapes.push(3);
        if (chkStar4.value) enabledShapes.push(4);

        if (enabledShapes.length === 0) return null;

        var shape = enabledShapes[Math.floor(Math.random() * enabledShapes.length)];
        var confetti;

        switch (shape) {
            case 0: // 四角形
                confetti = layer.pathItems.rectangle(y, x, size, size * random(0.3, 0.6));
                break;
            case 1: // 円
                confetti = layer.pathItems.ellipse(y, x, size, size);
                break;
            case 2: // 三角形
                confetti = layer.pathItems.add();
                var h = size * 1.2;
                confetti.setEntirePath([
                    [x, y],
                    [x + size, y],
                    [x + size * 0.5, y - h]
                ]);
                confetti.closed = true;
                break;
            case 3: // スター（五芒星）
                confetti = createStar(layer, x, y, size);
                break;
            case 4: // スター（4点）
                confetti = createStar4(layer, x, y, size);
                break;
        }

        // 色を設定
        confetti.filled = true;
        confetti.stroked = false;
        var rgbColor = new RGBColor();
        rgbColor.red = color[0];
        rgbColor.green = color[1];
        rgbColor.blue = color[2];
        confetti.fillColor = rgbColor;

        // ランダムに回転
        confetti.rotate(random(0, 360));

        // 不透明度 / Opacity
        try {
            if (chkOpacity && chkOpacity.value) {
                // 透明度をランダムに設定
                confetti.opacity = random(70, 100);
            } else {
                // OFF: 設定しない（=100%）
                confetti.opacity = 100;
            }
        } catch (_) { }

        return confetti;
    }

    var result = dlg.show();

    if (result !== 1) {
        clearPreview();
        try { previewLayer.remove(); } catch (e) { }
        try {
            if (__initView && __initZoom != null) {
                __initView.zoom = __initZoom;
                if (__initCenter) __initView.centerPoint = __initCenter;
            }
        } catch (_) { }
        return;
    }

    clearPreview();
    try { previewLayer.remove(); } catch (e) { }

    // 本番出力レイヤー
    var confettiLayer = doc.layers.add();
    confettiLayer.name = "Confetti";

    var mgFinal = buildMaskGroup(confettiLayer);
    var finalContainer = mgFinal.container;

    // コンフェティを生成
    var __bFinal = getEffectiveBounds();
    if (!__bFinal) {
        return;
    }

    for (var i = 0; i < confettiCount; i++) {
        var pt = pickPoint(__bFinal);
        var x = pt.x;
        var y = pt.y;
        var size = getConfettiSize();
        var color = colors[Math.floor(Math.random() * colors.length)];

        var item = createConfetti(finalContainer, x, y, size, color);
        if (!item) continue;
    }
    if (mgFinal.group && !__useArtboard) {
        applyMaskToGroup(mgFinal.group, selectedObj);
    }

})();
