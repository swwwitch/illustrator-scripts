#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

var SCRIPT_VERSION = "v1.0";

function getCurrentLang() {
    return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

// 日英ラベル定義 / Japanese-English label definitions
var LABELS = {
    dialogTitle: { ja: "罫線の欠け処理", en: "Line Gap Cut" },
    offset: { ja: "オフセット", en: "Offset" },
    strokeWidth: { ja: "線幅", en: "Stroke" },
    strokePanel: { ja: "線", en: "Stroke" }, // Stroke settings panel label
    capPanel: { ja: "線端", en: "Cap" },
    capButt: { ja: "なし", en: "Butt" },
    capRound: { ja: "丸型線端", en: "Round" },
    joinPanel: { ja: "角の形状", en: "Join" },
    joinMiter: { ja: "マイター", en: "Miter" },
    joinRound: { ja: "ラウンド", en: "Round" },
    joinBevel: { ja: "ベベル", en: "Bevel" },
    ok: { ja: "OK", en: "OK" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    alertNoDoc: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
    alertSelectTwo: { ja: "2つのオブジェクト（円と長方形）を選択してください。", en: "Select two objects (circle and rectangle)." }
};

function L(key) {
    var obj = LABELS[key];
    if (!obj) return key;
    return obj[lang] || obj.ja || key;
}

(function () {
    // ドキュメントが開かれているか確認
    if (app.documents.length === 0) {
        alert(L('alertNoDoc'));
        return;
    }

    var doc = app.activeDocument;

    // ↑↓ / Shift+↑↓ / Option(Alt)+↑↓ で数値を増減（ScriptUI edittext 用）

    var sel = doc.selection;

    // 2つのオブジェクトが選択されているか確認
    if (!sel || sel.length !== 2) {
        alert(L('alertSelectTwo'));
        return;
    }

    // selection[0] が最前面（上） = 長方形 (B)
    // selection[1] が最背面（下） = 円 (A)
    var objectB = sel[0];
    var objectA = sel[1];

    // 最後に必ずBを再表示するための保険
    var __restoreB = objectB;

    function safeDo(fn) { try { fn(); } catch (_) { } }
    function safeRemove(it) { safeDo(function () { if (it) it.remove(); }); }
    function safeSelect(it) {
        safeDo(function () { doc.selection = null; });
        safeDo(function () { if (it) it.selected = true; });
    }
    function safeMenu(cmd) { safeDo(function () { app.executeMenuCommand(cmd); }); }
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                // Shiftキー押下時は10の倍数にスナップ
                if (event.keyName == "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value = Math.floor((value - 1) / delta) * delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                // Option(Alt)キー押下時は0.1単位で増減
                if (event.keyName == "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            } else {
                delta = 1;
                if (event.keyName == "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            }

            if (keyboard.altKey) {
                // 小数第1位までに丸め
                value = Math.round(value * 10) / 10;
            } else {
                // 整数に丸め
                value = Math.round(value);
            }

            editText.text = value;
        });
    }

    /**
     * Illustrator 単位ユーティリティ（表示ラベル＋pt換算）
     * - rulerType: 定規
     * - strokeUnits: 線
     * - text/units: 文字
     * - text/asianunits: 東アジア言語（Q/H 表示切替）
     */
    var unitMap = {
        0: "in",
        1: "mm",
        2: "pt",
        3: "pica",
        4: "cm",
        6: "px",
        7: "ft/in",
        8: "m",
        9: "yd",
        10: "ft"
    };

    function getUnitCode(prefKey) {
        var v = 2; // pt
        safeDo(function () { v = app.preferences.getIntegerPreference(prefKey); });
        return v;
    }

    // 単位コードと設定キーから適切な単位ラベルを返す（Q/H分岐含む）
    function getUnitLabel(code, prefKey) {
        if (code === 5) {
            var hKeys = {
                "text/asianunits": true,
                "rulerType": true,
                "strokeUnits": true
            };
            return hKeys[prefKey] ? "H" : "Q";
        }
        return unitMap[code] || "pt";
    }

    // 単位コードから「1単位あたり何ptか」を返す
    function getPtFactorFromUnitCode(code) {
        switch (code) {
            case 0: return 72.0;                        // in
            case 1: return 72.0 / 25.4;                 // mm
            case 2: return 1.0;                         // pt
            case 3: return 12.0;                        // pica
            case 4: return 72.0 / 2.54;                 // cm
            case 5: return 72.0 / 25.4 * 0.25;          // Q or H (0.25mm)
            case 6: return 1.0;                         // px（内部では概ねpt扱い）
            case 7: return 72.0 * 12.0;                 // ft
            case 8: return 72.0 / 25.4 * 1000.0;        // m
            case 9: return 72.0 * 36.0;                 // yd
            case 10: return 72.0 * 12.0;                // ft
            default: return 1.0;
        }
    }

    function unitToPt(valueInUnit, prefKey) {
        var code = getUnitCode(prefKey);
        var f = getPtFactorFromUnitCode(code);
        return valueInUnit * f;
    }

    function ptToUnit(valuePt, prefKey) {
        var code = getUnitCode(prefKey);
        var f = getPtFactorFromUnitCode(code);
        if (!f) return valuePt;
        return valuePt / f;
    }

    function getCurrentUnitLabel(prefKey) {
        var code = getUnitCode(prefKey);
        return getUnitLabel(code, prefKey);
    }

    function round1(v) {
        return Math.round(v * 10) / 10;
    }
    // デフォルトのオフセット値：長方形化（C）の高さの1/3（pt）
    function getCHeightPtFromB(bItem) {
        if (!bItem) return 0;

        // TextFrame：複製→アウトライン→外接bounds（後始末あり）
        if (bItem.typename === 'TextFrame') {
            var tmpDup = null;
            var tmpOutlined = [];
            var region = null;

            try {
                tmpDup = bItem.duplicate();
                safeSelect(tmpDup);
                safeMenu('outline');

                var s = doc.selection;
                if (s && s.length) {
                    for (var i = 0; i < s.length; i++) tmpOutlined.push(s[i]);
                }

                if (tmpOutlined.length > 1) {
                    var g = doc.groupItems.add();
                    for (var j = tmpOutlined.length - 1; j >= 0; j--) {
                        safeDo(function () { tmpOutlined[j].move(g, ElementPlacement.PLACEATBEGINNING); });
                    }
                    region = g.geometricBounds;
                    tmpOutlined = [g];
                } else if (tmpOutlined.length === 1) {
                    region = tmpOutlined[0].geometricBounds;
                }
            } catch (_) {
                region = null;
            } finally {
                for (var k = 0; k < tmpOutlined.length; k++) safeRemove(tmpOutlined[k]);
                safeRemove(tmpDup);
                safeDo(function () { doc.selection = null; });
            }

            if (!region) return 0;
            var h = region[1] - region[3];
            return (h > 0) ? h : 0;
        }

        // 通常：Bのbounds高さを採用（CはB複製）
        var b = bItem.geometricBounds; // [L,T,R,B]
        var h2 = b[1] - b[3];
        return (h2 > 0) ? h2 : 0;
    }

    var __cHeightPtForDefault = getCHeightPtFromB(objectB);
    // 高さが取れない場合のフォールバックは従来の 10pt
    var __offsetDefaultPt = (__cHeightPtForDefault > 0) ? (__cHeightPtForDefault / 3) : 10;

    // オフセット値ダイアログ
    var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
    dlg.orientation = 'column';
    dlg.alignChildren = ['fill', 'top'];

    var grp1 = dlg.add('group');
    var lblOffset = grp1.add('statictext', undefined, L('offset'));
    // lblOffset.justification = 'right';
    var __offsetPrefKey = "rulerType";
    // __offsetDefaultPt は「C高さ/3（pt）」で事前計算済み
    var offsetInput = grp1.add('edittext', undefined, String(round1(ptToUnit(__offsetDefaultPt, __offsetPrefKey))));
    offsetInput.characters = 3;
    changeValueByArrowKey(offsetInput);
    grp1.add('statictext', undefined, getCurrentUnitLabel(__offsetPrefKey));

    // 線設定パネル
    var panelStroke = dlg.add('panel', undefined, L('strokePanel'));
    panelStroke.orientation = 'column';
    panelStroke.alignChildren = ['fill', 'top'];
    panelStroke.margins = [15, 20, 15, 10];

    var grp2 = panelStroke.add('group');
    var lblStroke = grp2.add('statictext', undefined, L('strokeWidth'));

    var __strokePrefKey = "strokeUnits";
    var __strokeDefaultPt = 2;
    var strokeInput = grp2.add('edittext', undefined, String(round1(ptToUnit(__strokeDefaultPt, __strokePrefKey))));
    strokeInput.characters = 3;
    changeValueByArrowKey(strokeInput);
    grp2.add('statictext', undefined, getCurrentUnitLabel(__strokePrefKey));

    // 線端設定
    var panelCap = panelStroke.add('panel', undefined, L('capPanel'));
    panelCap.orientation = 'column';
    panelCap.alignChildren = 'left';
    panelCap.margins = [15, 20, 15, 10];

    var rbCapButt = panelCap.add('radiobutton', undefined, L('capButt'));
    var rbCapRound = panelCap.add('radiobutton', undefined, L('capRound'));

    rbCapButt.value = true; // デフォルト

    // 丸型線端が選択されたら、角の形状をラウンドに自動切替
    rbCapRound.onClick = function () {
        if (rbCapRound.value) {
            rbJoinRound.value = true;
        }
    };

    // 角の形状（線のジョイン）
    var panelJoin = panelStroke.add('panel', undefined, L('joinPanel'));
    panelJoin.orientation = 'column';
    panelJoin.alignChildren = 'left';
    panelJoin.margins = [15, 20, 15, 10];

    var rbJoinMiter = panelJoin.add('radiobutton', undefined, L('joinMiter'));
    var rbJoinRound = panelJoin.add('radiobutton', undefined, L('joinRound'));
    var rbJoinBevel = panelJoin.add('radiobutton', undefined, L('joinBevel'));

    rbJoinMiter.value = true; // デフォルト

    var btns = dlg.add('group');
    btns.orientation = 'row';
    btns.alignChildren = ['fill', 'center'];
    btns.alignment = 'fill';

    var btnLeft = btns.add('group');
    btnLeft.alignment = 'left';
    var cancelBtn = btnLeft.add('button', undefined, L('cancel'), { name: 'cancel' });

    var btnRight = btns.add('group');
    btnRight.alignment = 'right';
    var okBtn = btnRight.add('button', undefined, L('ok'), { name: 'ok' });

    if (dlg.show() !== 1) return;

    var offsetValue = parseFloat(offsetInput.text);
    if (isNaN(offsetValue)) offsetValue = ptToUnit(10, __offsetPrefKey);
    var offsetValuePt = unitToPt(offsetValue, __offsetPrefKey);

    var strokeWidthValue = parseFloat(strokeInput.text);
    if (isNaN(strokeWidthValue)) strokeWidthValue = ptToUnit(2, __strokePrefKey);
    var strokeWidthValuePt = unitToPt(strokeWidthValue, __strokePrefKey);

    var strokeCapValue = rbCapRound.value ? StrokeCap.ROUNDENDCAP : StrokeCap.BUTTENDCAP;
    var strokeJoinValue = rbJoinRound.value ? StrokeJoin.ROUNDENDJOIN : (rbJoinBevel.value ? StrokeJoin.BEVELENDJOIN : StrokeJoin.MITERENDJOIN);



    try {

        //念のためロックされていたら解除
        objectA.locked = false;
        objectB.locked = false;

        // 0. selection[1]（=A）を複製し、ガイド化
        //    ガイドはロックしておく（誤操作防止）
        (function makeGuideFromA() {
            var dup = null;
            try {
                dup = objectA.duplicate();
                safeDo(function () { dup.name = '__GUIDE_A__'; });

                // ガイド化（メニュー優先、失敗時はプロパティでフォールバック）
                safeSelect(dup);
                safeMenu('makeGuide');
                // makeGuide が失敗した場合のみ guides=true を試す
                safeDo(function () {
                    if (dup && dup.guides !== true) dup.guides = true;
                });

                safeDo(function () { dup.locked = true; });
                safeDo(function () { doc.selection = null; });
            } catch (_) {
                safeRemove(dup);
                safeDo(function () { doc.selection = null; });
            }
        })();

        // 1. Cを作成
        //  - 通常：Bを複製してC
        //  - Bがテキストの場合：複製→アウトライン→バウンディングから長方形化（C）
        function makeRectFromBounds(bounds) {
            // bounds: [L, T, R, B]
            var L = bounds[0], T = bounds[1], R = bounds[2], B = bounds[3];
            var w = R - L;
            var h = T - B;
            // rectangle(top, left, width, height)
            var rect = doc.pathItems.rectangle(T, L, w, h);
            return rect;
        }

        function createCFromTextFrame(tf) {
            // tf: TextFrame
            var tmpDup = null;
            var tmpOutlined = [];
            try {
                // 1) 複製（A）
                tmpDup = tf.duplicate();

                // 2) アウトライン化（B）
                safeSelect(tmpDup);
                safeMenu('outline');

                // outline後の選択を回収
                var s = doc.selection;
                if (!s || s.length === 0) return null;

                // 選択されたものを一旦配列に退避
                for (var i = 0; i < s.length; i++) tmpOutlined.push(s[i]);

                // 3) 長方形化（C）: アウトラインの外接バウンディングから矩形を作る
                var region = null;

                // 可能ならグループ化して bounds を安定させる
                if (tmpOutlined.length > 1) {
                    var g = doc.groupItems.add();
                    for (var j = tmpOutlined.length - 1; j >= 0; j--) {
                        try { tmpOutlined[j].move(g, ElementPlacement.PLACEATBEGINNING); } catch (_) { }
                    }
                    region = g.geometricBounds;
                    // グループを後で消せるように差し替え
                    tmpOutlined = [g];
                } else {
                    region = tmpOutlined[0].geometricBounds;
                }


                if (!region) return null;

                var rectC = makeRectFromBounds(region);

                // アウトライン化した一時オブジェクトは削除
                for (var k = 0; k < tmpOutlined.length; k++) {
                    safeRemove(tmpOutlined[k]);
                }
                safeRemove(tmpDup);
                safeDo(function () { doc.selection = null; });
                return rectC;
            } catch (e) {
                // 後始末
                safeRemove(tmpDup);
                for (var x = 0; x < tmpOutlined.length; x++) {
                    safeRemove(tmpOutlined[x]);
                }
                safeDo(function () { doc.selection = null; });
                return null;
            }
        }

        var objectC = null;
        if (objectB && objectB.typename === 'TextFrame') {
            objectC = createCFromTextFrame(objectB);
        }
        if (!objectC) {
            // 通常ケース：Bを複製してC
            objectC = objectB.duplicate();
        }

        // Cを上下左右にオフセット（外側へ拡張）
        safeDo(function () {
            objectC.left -= offsetValuePt;
            objectC.top += offsetValuePt;
            objectC.width += offsetValuePt * 2;
            objectC.height += offsetValuePt * 2;
        });

        // Cの領域（後で、この領域内にあるパスを削除する）
        var __cBounds = null; // [L, T, R, B]
        safeDo(function () { __cBounds = objectC.geometricBounds; });

        // 2. Bは隠す
        objectB.hidden = true;

        // 3. AとCに対して設定：塗り=K30、線=なし
        function applyFillK30NoStroke(item) {
            if (!item) return;

            // K30（CMYK）
            var k30 = new CMYKColor();
            k30.cyan = 0;
            k30.magenta = 0;
            k30.yellow = 0;
            k30.black = 30;

            function applyToPath(p) {
                safeDo(function () {
                    p.filled = true;
                    p.fillColor = k30;
                    p.stroked = false;
                });
            }

            try {
                var t = item.typename;
                if (t === 'PathItem') {
                    applyToPath(item);
                    return;
                }
                if (t === 'CompoundPathItem') {
                    for (var i = 0; i < item.pathItems.length; i++) {
                        applyToPath(item.pathItems[i]);
                    }
                    return;
                }
                if (t === 'GroupItem') {
                    // 直下のパス
                    for (var j = 0; j < item.pathItems.length; j++) {
                        applyToPath(item.pathItems[j]);
                    }
                    // 直下の複合パス
                    for (var k = 0; k < item.compoundPathItems.length; k++) {
                        applyFillK30NoStroke(item.compoundPathItems[k]);
                    }
                    // ネストしたグループ
                    for (var g = 0; g < item.groupItems.length; g++) {
                        applyFillK30NoStroke(item.groupItems[g]);
                    }
                    return;
                }
                if (t === 'TextFrame') {
                    // 念のため（今回の用途では想定外だが）
                    safeDo(function () {
                        item.textRange.characterAttributes.fillColor = k30;
                        item.textRange.characterAttributes.strokeColor = k30;
                        item.textRange.characterAttributes.strokeWeight = 0;
                    });
                    return;
                }
            } catch (_) { }
        }

        applyFillK30NoStroke(objectA);
        applyFillK30NoStroke(objectC);


        // 3. AとCをグループ化 (D)
        // 新しいグループアイテムを作成
        var groupD = doc.groupItems.add();

        // AとCをグループ内に移動
        // move関数は move(target, elementPlacement)
        objectA.move(groupD, ElementPlacement.PLACEATEND);
        objectC.move(groupD, ElementPlacement.PLACEATBEGINNING);

        // グループDを選択状態にする（executeMenuCommandは選択オブジェクトに対して実行されるため）
        doc.selection = null; // 一旦選択解除
        groupD.selected = true;

        // 4. Dに次を実行し、パスを分割
        // Live Pathfinder Divide (パスファインダー：分割) を適用
        // 注意: 'Live Pathfinder Divide' は効果としてのパスファインダーです。
        // 通常のパネル操作の「分割」挙動に近いのは 'Live Pathfinder Divide' 後に 'expandStyle' するか、
        // あるいは直接 'group' に対してパスファインダーパネルの分割コマンドを送ることです。

        // 方法1: ライブ効果を適用して拡張する（ご要望の手順）
        app.executeMenuCommand('Live Pathfinder Outline');
        app.executeMenuCommand('expandStyle');

        // 5. 生成結果（選択中）をグループ解除し、線幅1px・黒を適用
        function hasSelectedGroup() {
            var s = doc.selection;
            if (!s || s.length === 0) return false;
            for (var i = 0; i < s.length; i++) {
                if (s[i] && s[i].typename === 'GroupItem') return true;
            }
            return false;
        }

        // 可能な限りグループ解除（ネストも含めて）
        (function ungroupFully() {
            var safety = 0;
            while (hasSelectedGroup() && safety < 50) {
                app.executeMenuCommand('ungroup');
                safety++;
            }
        })();

        // 6. Cの領域にあるパスを削除（Cの内側に完全に収まるものを対象）
        function boundsInside(b, region, tol) {
            if (!b || !region) return false;
            tol = tol || 0;
            // b, region: [L, T, R, B]
            return (
                b[0] >= region[0] - tol &&
                b[1] <= region[1] + tol &&
                b[2] <= region[2] + tol &&
                b[3] >= region[3] - tol
            );
        }

        (function deleteInsideC() {
            if (!__cBounds) return;

            // 実行直後の選択（分割結果）を走査
            var s = doc.selection;
            if (!s || s.length === 0) return;

            // 選択を壊さないようにコピーしてから削除
            var items = [];
            for (var i = 0; i < s.length; i++) items.push(s[i]);

            // 少しだけ許容（境界線上の誤差吸収）
            var tol = 0.01;

            for (var j = 0; j < items.length; j++) {
                var it = items[j];
                if (!it) continue;
                safeDo(function () {
                    // ほとんどの生成物は PathItem / CompoundPathItem だが、念のためtypenameは問わず bounds で判断
                    var b = it.geometricBounds;
                    if (boundsInside(b, __cBounds, tol)) {
                        it.remove();
                    }
                });
            }

            // selectionを更新（削除されたものを除外）
            safeDo(function () {
                doc.selection = null;
                // 残っているものだけ再選択
                for (var k = 0; k < items.length; k++) {
                    safeDo(function () {
                        if (items[k] && items[k].parent) items[k].selected = true;
                    });
                }
            });
        })();

        // 黒の作成（ドキュメントのカラースペースに合わせる）
        function makeBlackColor() {
            var isCMYK = false;
            safeDo(function () { isCMYK = (doc.documentColorSpace === DocumentColorSpace.CMYK); });
            if (isCMYK) {
                var c = new CMYKColor();
                c.cyan = 0; c.magenta = 0; c.yellow = 0; c.black = 100;
                return c;
            }
            var r = new RGBColor();
            r.red = 0; r.green = 0; r.blue = 0;
            return r;
        }

        var __black = makeBlackColor();

        function applyStrokeBlack1px(item) {
            if (!item) return;

            function applyToPath(p) {
                safeDo(function () {
                    p.stroked = true;
                    p.strokeColor = __black;
                    // Illustratorはpt基準。strokeWidthValuePt pt を採用
                    p.strokeWidth = strokeWidthValuePt;
                    p.strokeCap = strokeCapValue;
                    p.strokeJoin = strokeJoinValue;
                });
            }

            var t = item.typename;
            if (t === 'PathItem') {
                applyToPath(item);
                return;
            }
            if (t === 'CompoundPathItem') {
                for (var i = 0; i < item.pathItems.length; i++) applyToPath(item.pathItems[i]);
                return;
            }
            if (t === 'GroupItem') {
                for (var j = 0; j < item.pathItems.length; j++) applyToPath(item.pathItems[j]);
                for (var k = 0; k < item.compoundPathItems.length; k++) applyStrokeBlack1px(item.compoundPathItems[k]);
                for (var g = 0; g < item.groupItems.length; g++) applyStrokeBlack1px(item.groupItems[g]);
                return;
            }
        }

        // 選択中の生成物に適用
        var __outSel = doc.selection;
        if (__outSel && __outSel.length) {
            for (var __i = 0; __i < __outSel.length; __i++) {
                applyStrokeBlack1px(__outSel[__i]);
            }
        }

    } finally {
        // すべて終了したら、Bを再表示
        safeDo(function () {
            if (__restoreB) __restoreB.hidden = false;
        });
    }

})();