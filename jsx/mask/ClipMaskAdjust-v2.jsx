#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
  ClipMaskAdjust-v2 (Auto-Preview)
  更新日: 2026-01-03

  クリップグループ（クリッピングマスク）のマスクパスと内容を調整するスクリプトです。
  ダイアログ操作はオートプレビューで即時反映されます（Undo/ヒストリー処理は現在無効）。

  - 基準点（9点）で内容を揃える
  - フィットとスケール（Cover / Contain / サイズ保持 / 手動スケール）
  - マスクパス（そのまま / 内容に合わせる / 正方形に）
  - 微調整（X/Y）
  - 角丸：クリップグループ全体に「角を丸くする」効果を適用
  - 正円：ON時は角丸をONにし、半径を幅/高さの半分に設定（可能なら既存角丸の値のみ更新）
  - 実行時、クリップグループのアピアランスを消去

  ショートカット:
  - 基準点: qwe / asd / zxc
  - 数値増減: ↑↓(±1), Shift+↑↓(±10), Opt+↑↓(±0.1)
*/

(function () {
    var doc = app.activeDocument;
    var sel = doc.selection;

    // 初期選択（参照）を保持：最終確定時に選択を戻すため
    var originalSelection = [];
    for (var si = 0; si < sel.length; si++) originalSelection.push(sel[si]);

    if (sel.length === 0) {
        alert("クリッピングマスクを選択してください。");
        return;
    }



    // --- Localization (JA / EN) ---
    var LANG = (app.locale && app.locale.indexOf('ja') === 0) ? 'ja' : 'en';
    var LABELS = {
        ja: {
            dialogTitle: 'クリップグループの調整',
            anchor: '基準点',
            fitScale: 'フィットとスケール',
            maskPath: 'マスクパス',
            cover: '縦横比を保持して切り取り',
            contain: '縦横比を保持して縮小',
            none: 'サイズ保持',
            manual: 'スケールを指定',
            maskNone: 'そのまま',
            fitFrame: '内容に合わせる',
            square: '正方形に',
            round: '角丸',
            roundPanel: '角丸',
            circle: '正円',
            cancel: 'キャンセル',
            ok: 'OK',
            tweak: '微調整',
            x: 'X',
            y: 'Y'
        },
        en: {
            dialogTitle: 'Clip Group Adjust',
            anchor: 'Anchor',
            fitScale: 'Fit & Scale',
            maskPath: 'Mask Path',
            cover: 'Cover (Keep Ratio)',
            contain: 'Contain (Keep Ratio)',
            none: 'Keep Size',
            manual: 'Set Scale',
            maskNone: 'None',
            fitFrame: 'Fit to Content',
            square: 'Square',
            round: 'Round Corners',
            roundPanel: 'Round Corners',
            circle: 'Circle',
            cancel: 'Cancel',
            ok: 'OK',
            tweak: 'Nudge',
            x: 'X',
            y: 'Y'
        }
    };
    var L = LABELS[LANG];

    // --- Units (follow rulerType) ---
    // 単位コードとラベルのマップ
    var unitLabelMap = {
        0: "in",
        1: "mm",
        2: "pt",
        3: "pica",
        4: "cm",
        5: "Q/H",
        6: "px",
        7: "ft/in",
        8: "m",
        9: "yd",
        10: "ft"
    };

    // 現在の単位ラベルを取得
    function getCurrentUnitLabel() {
        var unitCode = app.preferences.getIntegerPreference("rulerType");
        return unitLabelMap[unitCode] || "pt";
    }

    // UnitValue に渡す単位（対応範囲のみ）
    function getUnitValueUnit() {
        var unitCode = app.preferences.getIntegerPreference("rulerType");
        switch (unitCode) {
            case 0: return "in";
            case 1: return "mm";
            case 2: return "pt";
            case 3: return "pica";
            case 4: return "cm";
            case 5: return "Q";   // 表示は Q/H だが UnitValue は Q
            case 6: return "px";
            default: return "pt";
        }
    }

    function valueInCurrentUnitToPt(n) {
        var u = getUnitValueUnit();
        try {
            return new UnitValue(n, u).as('pt');
        } catch (e) {
            return n; // フォールバック（pt想定）
        }
    }

    function ptToValueInCurrentUnit(pt) {
        var u = getUnitValueUnit();
        try {
            return new UnitValue(pt, 'pt').as(u);
        } catch (e) {
            return pt;
        }
    }

    var CURRENT_UNIT_LABEL = getCurrentUnitLabel();

    // --- アピアランス消去（クリップグループ対象） ---
    // executeMenuCommand("clearAppearance") では環境差が出ることがあるため、アクションを一時生成して実行
    var CLEAR_APPEARANCE_ACTION_CODE = '''/version 3
/name [ 20
    5f5f7374746b335f617070656172616e63655f5f
]
/isOpen 1
/actionCount 1
/action-1 {
    /name [ 5
        636c656172
    ]
    /keyIndex 0
    /colorIndex 0
    /isOpen 0
    /eventCount 1
    /event-1 {
        /useRulersIn1stQuadrant 0
        /internalName (ai_plugin_appearance)
        /localizedName [ 18
            e382a2e38394e382a2e383a9e383b3e382b9
        ]
        /isOpen 0
        /isOn 1
        /hasDialog 0
        /parameterCount 1
        /parameter-1 {
            /key 1835363957
            /showInPalette 4294967295
            /type (enumerated)
            /name [ 27
                e382a2e38394e382a2e383a9e383b3e382b9e38292e6b688e58ebb
            ]
            /value 6
        }
    }
}''';

    function tempAction(actionCode, func) {
        var hexToString = function (hex) {
            return decodeURIComponent(hex.replace(/(.{2})/g, '%$1'));
        };

        var ActionItem = function ActionItem(index, name, parent) {
            this.index = index;
            this.name = name; // actionName
            this.parent = parent; // setName
        };
        ActionItem.prototype.exec = function (showDialog) {
            doScript(this.name, this.parent, showDialog);
        };

        var ActionItems = function ActionItems() { this.length = 0; };
        ActionItems.prototype.getByName = function (nameStr) {
            for (var i = 0, len = this.length; i < len; i++) {
                if (this[i].name == nameStr) return this[i];
            }
        };
        ActionItems.prototype.index = function (keyNumber) {
            if (keyNumber >= 0) return this[keyNumber];
            return this[this.length + keyNumber];
        };

        var regExpSetName = /^\/name\s+\[\s+\d+\s+([^\]]+?)\s+\]/m;
        var setNameMatch = actionCode.match(regExpSetName);
        // 何らかの理由でパースできない場合のフォールバック
        var setName = setNameMatch ? hexToString(setNameMatch[1].replace(/\s+/g, '')) : '__sttk3_appearance__';

        var regExpActionNames = /^\/action-\d+\s+\{\s+\/name\s+\[\s+\d+\s+([^\]]+?)\s+\]/mg;
        var actionItemsObj = new ActionItems();
        var i = 0;
        var matchObj;
        while ((matchObj = regExpActionNames.exec(actionCode))) {
            var actionName = hexToString(matchObj[1].replace(/\s+/g, ''));
            var actionObj = new ActionItem(i, actionName, setName);
            actionItemsObj[actionName] = actionObj;
            actionItemsObj[i] = actionObj;
            i++;
            if (i > 1000) break;
        }
        actionItemsObj.length = i;

        var failed = false;
        var aiaFileObj = new File(Folder.temp + '/tempActionSet.aia');
        try {
            aiaFileObj.open('w');
            aiaFileObj.write(actionCode);
        } catch (e) {
            failed = true;
            alert(e);
            return;
        } finally {
            aiaFileObj.close();
            if (failed) { try { aiaFileObj.remove(); } catch (e2) { } }
        }

        try { app.unloadAction(setName, ''); } catch (e3) { }

        var actionLoaded = false;
        var executed = false;
        try {
            app.loadAction(aiaFileObj);
            actionLoaded = true;
            func.call(func, actionItemsObj);
            executed = true;
        } catch (e4) {
            // ここではアラートを出さず、失敗してもスクリプト自体は継続
        } finally {
            if (actionLoaded) { try { app.unloadAction(setName, ''); } catch (e5) { } }
            try { aiaFileObj.remove(); } catch (e6) { }
        }
        return executed;
    }

    function setSelection(items) {
        try { app.executeMenuCommand('deselectall'); } catch (e) { }
        for (var i = 0; i < items.length; i++) {
            try { items[i].selected = true; } catch (e2) { }
        }
    }

    function clearAppearanceForClipGroups(selection) {
        // 現在の選択を保持（参照でOK）
        var original = [];
        for (var i = 0; i < selection.length; i++) original.push(selection[i]);

        var anyCleared = false;

        for (var j = 0; j < selection.length; j++) {
            var it = selection[j];
            if (it && it.typename === 'GroupItem' && it.clipped) {
                setSelection([it]);
                var ok = tempAction(CLEAR_APPEARANCE_ACTION_CODE, function (actionItems) {
                    actionItems[0].exec(false);
                });
                if (ok) anyCleared = true;
            }
        }

        // 選択状態を元に戻す
        setSelection(original);

        return anyCleared;
    }

    // --- 補助関数：角丸デフォルト値（pt）を計算 ---
    // （マスクパスの幅＋マスクパスの高さ）÷25（端数は繰り上げ）
    function getDefaultRoundRadiusFromMask(groupItem) {
        try {
            if (!groupItem || groupItem.typename !== 'GroupItem' || !groupItem.clipped) return null;
            var clipPath = null;
            for (var i = 0; i < groupItem.pageItems.length; i++) {
                var p = groupItem.pageItems[i];
                if (p && p.clipping) { clipPath = p; break; }
            }
            if (!clipPath) return null;
            var b = clipPath.geometricBounds; // [L, T, R, B]
            var w = b[2] - b[0];
            var h = b[1] - b[3];
            if (!isFinite(w) || !isFinite(h)) return null;
            var vPt = (w + h) / 25;
            if (!isFinite(vPt)) return null;
            // 現在の単位に変換して表示用の値を作る（端数は繰り上げ）
            var v = ptToValueInCurrentUnit(vPt);
            return Math.ceil(v);
        } catch (e) {
            return null;
        }
    }

    // --- 補助関数：現在のスケール（％）を取得 ---
    function getCurrentScale(groupItem) {
        var contents = [];
        for (var i = 0; i < groupItem.pageItems.length; i++) {
            if (!groupItem.pageItems[i].clipping) contents.push(groupItem.pageItems[i]);
        }
        if (contents.length === 0) return 100;
        var target = contents[0];
        var m = target.matrix;
        var scaleX = Math.sqrt(m.mValueA * m.mValueA + m.mValueB * m.mValueB);
        return (scaleX * 100).toFixed(2);
    }

    // --- 角丸用XML生成 ---
    function createRoundCornersEffectXML(radius) {
        var xml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius #value# "/></LiveEffect>';
        return xml.replace('#value#', radius);
    }

    // --- Round Corners: avoid stacking duplicate LiveEffects ---
    function stripRoundCornersEffect(effectStr) {
        if (!effectStr) return '';
        // Match any LiveEffect whose name contains "Round Corners" (handles serialization/name variants)
        return effectStr.replace(/<LiveEffect\b[^>]*\bname=['\"][^'\"]*Round Corners[^'\"]*['\"][^>]*>[\s\S]*?<\/LiveEffect>/g, '');
    }

    function removeRoundCornersEffect(targetItem) {
        if (!targetItem) return;
        try {
            var current = targetItem.appliedEffect || '';
            targetItem.appliedEffect = stripRoundCornersEffect(current);
        } catch (e) {
            // ignore
        }
    }

    function applyRoundCornersEffect(targetItem, radiusPt) {
        if (!targetItem) return;
        try {
            // まず既存の角丸だけ除去（他の効果は維持）
            var current = targetItem.appliedEffect || '';
            current = stripRoundCornersEffect(current);
            // appliedEffect の直接連結は環境によって無視されることがあるため、
            // 一度 appliedEffect を更新してから applyEffect で追加する
            targetItem.appliedEffect = current;
            targetItem.applyEffect(createRoundCornersEffectXML(radiusPt));
        } catch (e) {
            // Fallback: 何もしない（角丸は必須ではない）
            try { targetItem.applyEffect(createRoundCornersEffectXML(radiusPt)); } catch (e2) { }
        }
    }

    // Try to update the radius value inside an existing Round Corners LiveEffect without re-applying.
    // Returns true if updated, false if the effect block was not found or could not be rewritten.
    function updateRoundCornersRadiusOnly(targetItem, radiusPt) {
        if (!targetItem) return false;
        try {
            var effects = targetItem.appliedEffect || '';
            if (!effects) return false;

            // Round Corners を含む LiveEffect ブロックを抜き出す（名前の揺れ対策）
            var reEffect = /<LiveEffect\b[^>]*\bname=['\"][^'\"]*Round Corners[^'\"]*['\"][^>]*>[\s\S]*?<\/LiveEffect>/;
            var m = effects.match(reEffect);
            if (!m || !m[0]) return false;

            var block = m[0];

            // ブロック内の "R radius <number>" の最初の1箇所だけ差し替え
            var reRadius = /(R\s+radius\s+)(-?\d+(?:\.\d+)?)/;
            if (!reRadius.test(block)) return false;

            var newBlock = block.replace(reRadius, function (_, p1) {
                return p1 + String(radiusPt);
            });

            // 差し替えたブロックを戻す
            var newEffects = effects.replace(block, newBlock);
            targetItem.appliedEffect = newEffects;
            return true;
        } catch (e) {
            return false;
        }
    }

    // --- 矢印キーでの数値増減機能 ---
    function changeValueByArrowKey(editText, onUpdate, allowNegative) {
        editText.addEventListener("keydown", function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                if (event.keyName == "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value = Math.floor((value - 1) / delta) * delta;
                    if (!allowNegative && value < 0) value = 0;
                    event.preventDefault();
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                if (event.keyName == "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= delta;
                    event.preventDefault();
                }
            } else {
                delta = 1;
                if (event.keyName == "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= delta;
                    if (!allowNegative && value < 0) value = 0;
                    event.preventDefault();
                }
            }

            if (keyboard.altKey) {
                value = Math.round(value * 10) / 10;
            } else {
                value = Math.round(value);
            }

            editText.text = value;
            if (onUpdate) onUpdate();
        });
    }

    var initialScale = getCurrentScale(sel[0]);

    // UI作成
    var win = new Window("dialog", L.dialogTitle);
    win.orientation = "column";
    win.alignChildren = ["center", "top"];
    win.spacing = 15;

    var mainGroup = win.add("group");
    mainGroup.orientation = "row";
    mainGroup.alignChildren = ["fill", "top"];
    mainGroup.spacing = 10;
    var commonMargins = [15, 20, 15, 10];

    // --- 1. 基準点 ---
    var leftCol = mainGroup.add("group");
    leftCol.orientation = "column";
    leftCol.alignChildren = ["fill", "top"];
    leftCol.spacing = 10;

    var anchorPanel = leftCol.add("panel", undefined, L.anchor);
    anchorPanel.margins = commonMargins;
    anchorPanel.alignChildren = ["center", "center"];
    var anchorRadios = [];
    var matrixGroup = anchorPanel.add("group");
    matrixGroup.orientation = "column";
    matrixGroup.spacing = 5;
    for (var i = 0; i < 3; i++) {
        var row = matrixGroup.add("group");
        for (var j = 0; j < 3; j++) {
            var r = row.add("radiobutton", undefined, "");
            r.size = [15, 15];
            anchorRadios.push(r);
        }
    }
    anchorRadios[4].value = true;

    // --- 1.1 微調整 ---
    var tweakPanel = leftCol.add("panel", undefined, L.tweak);
    tweakPanel.margins = commonMargins;
    tweakPanel.alignChildren = ["left", "top"];
    tweakPanel.spacing = 6;

    var tweakXGroup = tweakPanel.add("group");
    tweakXGroup.orientation = "row";
    tweakXGroup.spacing = 5;
    tweakXGroup.add("statictext", undefined, L.x);
    var tweakXInput = tweakXGroup.add("edittext", undefined, "0");
    tweakXInput.characters = 4;
    tweakXGroup.add("statictext", undefined, CURRENT_UNIT_LABEL);

    var tweakYGroup = tweakPanel.add("group");
    tweakYGroup.orientation = "row";
    tweakYGroup.spacing = 5;
    tweakYGroup.add("statictext", undefined, L.y);
    var tweakYInput = tweakYGroup.add("edittext", undefined, "0");
    tweakYInput.characters = 4;
    tweakYGroup.add("statictext", undefined, CURRENT_UNIT_LABEL);

    changeValueByArrowKey(tweakXInput, updatePreview, true);
    changeValueByArrowKey(tweakYInput, updatePreview, true);
    tweakXInput.onChanging = updatePreview;
    tweakYInput.onChanging = updatePreview;

    // --- 2. スケール ---
    var fitPanel = mainGroup.add("panel", undefined, L.fitScale);
    fitPanel.margins = commonMargins;
    fitPanel.alignChildren = ["left", "top"];
    fitPanel.spacing = 8;

    var radioCover = fitPanel.add("radiobutton", undefined, L.cover);
    var radioContain = fitPanel.add("radiobutton", undefined, L.contain);
    var radioNone = fitPanel.add("radiobutton", undefined, L.none);

    var manualScaleGroup = fitPanel.add("group");
    manualScaleGroup.orientation = "row";
    manualScaleGroup.spacing = 5;
    var radioManual = manualScaleGroup.add("radiobutton", undefined, L.manual);
    var scaleInput = manualScaleGroup.add("edittext", undefined, initialScale);
    scaleInput.characters = 6;
    manualScaleGroup.add("statictext", undefined, "%");

    radioNone.value = true;
    changeValueByArrowKey(scaleInput, function () { setScaleRadio(radioManual); updatePreview(); });

    // --- 3. マスクパス ---
    var rightCol = mainGroup.add("group");
    rightCol.orientation = "column";
    rightCol.alignChildren = ["fill", "top"];
    rightCol.spacing = 10;

    var maskPanel = rightCol.add("panel", undefined, L.maskPath);
    maskPanel.margins = commonMargins;
    maskPanel.orientation = "column";
    maskPanel.alignChildren = ["left", "top"];
    maskPanel.spacing = 8;

    var maskNone = maskPanel.add("radiobutton", undefined, L.maskNone);
    var radioFitFrame = maskPanel.add("radiobutton", undefined, L.fitFrame);
    var radioSquare = maskPanel.add("radiobutton", undefined, L.square);


    // --- 3.1 角丸 ---
    var roundPanel = rightCol.add("panel", undefined, L.roundPanel);
    roundPanel.margins = commonMargins;
    roundPanel.orientation = "column";
    roundPanel.alignChildren = ["left", "top"];
    roundPanel.spacing = 8;

    var roundGroup = roundPanel.add("group");
    roundGroup.orientation = "row";
    roundGroup.spacing = 5;
    var checkRound = roundGroup.add("checkbox", undefined, "");
    var defaultRound = getDefaultRoundRadiusFromMask(sel[0]);
    var roundInput = roundGroup.add("edittext", undefined, (defaultRound != null ? String(defaultRound) : "10"));
    roundInput.characters = 4;
    roundGroup.add("statictext", undefined, CURRENT_UNIT_LABEL);

    // ［正円］（ロジックは後で追加）
    var checkCircle = roundPanel.add("checkbox", undefined, L.circle);
    checkCircle.value = false;
    updateCircleAvailability();

    maskNone.value = true;
    checkRound.value = false;
    changeValueByArrowKey(roundInput, function () { checkRound.value = true; clearAppearanceForClipGroups(sel); updatePreview(); });
    // ボタン
    var btnGroup = win.add("group");
    var cancelBtn = btnGroup.add("button", undefined, L.cancel, { name: "cancel" });
    var okBtn = btnGroup.add("button", undefined, L.ok, { name: "ok" });

    // ---------------------------------------------------------
    // ロジック
    // ---------------------------------------------------------

    // 基準点ショートカット
    win.addEventListener("keydown", function (event) {
        if (win.activeControl === scaleInput || win.activeControl === roundInput) return;
        var key = event.keyName.toLowerCase();
        var keyMap = { 'q': 0, 'w': 1, 'e': 2, 'a': 3, 's': 4, 'd': 5, 'z': 6, 'x': 7, 'c': 8 };
        if (keyMap.hasOwnProperty(key)) {
            setAnchorRadio(keyMap[key]);
            resetTweakXY();
            updatePreview();
        }
    });

    var scaleRadios = [radioCover, radioContain, radioNone, radioManual];
    function setScaleRadio(selectedRadio) {
        for (var i = 0; i < scaleRadios.length; i++) scaleRadios[i].value = (scaleRadios[i] === selectedRadio);
    }

    function setAnchorRadio(index) {
        for (var i = 0; i < anchorRadios.length; i++) {
            anchorRadios[i].value = (i === index);
        }
    }

    function resetTweakXY() {
        tweakXInput.text = '0';
        tweakYInput.text = '0';
    }

    function updateCircleAvailability() {
        // 「正方形に」のときのみ「正円」を使える
        checkCircle.enabled = !!radioSquare.value;
    }

    function setRoundRadiusToHalfOfSquareMask() {
        try {
            var g = sel[0];
            if (!g || g.typename !== 'GroupItem' || !g.clipped) return;

            // クリッピングパスを取得
            var clipPath = null;
            for (var i = 0; i < g.pageItems.length; i++) {
                var p = g.pageItems[i];
                if (p && p.clipping) { clipPath = p; break; }
            }
            if (!clipPath) return;

            // 見た目の寸法を基準にする（stroke等の影響を含める）
            var b = clipPath.visibleBounds; // [L,T,R,B]
            var w = b[2] - b[0];
            var h = b[1] - b[3];
            if (!isFinite(w) || !isFinite(h)) return;

            // 半径＝幅または高さの半分（正方形なら同値）→半径＝短辺の半分
            var radiusPt = Math.min(w, h) / 2;

            // 表示用：現在の rulerType 単位へ変換
            var v = ptToValueInCurrentUnit(radiusPt);
            v = Math.round(v * 100) / 100; // 小数2桁で整形
            roundInput.text = String(v);
        } catch (e) {
            // 失敗しても無視
        }
    }

    function getInputState() {
        var anchorIndex = 4;
        for (var i = 0; i < anchorRadios.length; i++) {
            if (anchorRadios[i].value) { anchorIndex = i; break; }
        }

        var fitMode = "cover";
        if (radioContain.value) fitMode = "contain";
        else if (radioNone.value) fitMode = "none";
        else if (radioManual.value) fitMode = "manual";

        var maskMode = "none";
        if (radioFitFrame.value) maskMode = "fitFrame";
        else if (radioSquare.value) maskMode = "makeSquare";

        var applyRound = !!checkRound.value;
        var manualScaleVal = parseFloat(scaleInput.text) || 100;
        var roundVal = valueInCurrentUnitToPt(parseFloat(roundInput.text) || 0);
        var tweakX = valueInCurrentUnitToPt(parseFloat(tweakXInput.text) || 0);
        var tweakY = valueInCurrentUnitToPt(parseFloat(tweakYInput.text) || 0);

        return {
            anchorIndex: anchorIndex,
            fitMode: fitMode,
            maskMode: maskMode,
            applyRound: applyRound,
            manualScaleVal: manualScaleVal,
            roundVal: roundVal,
            tweakX: tweakX,
            tweakY: tweakY
        };
    }

    function updatePreview() {
        // Undo を使わずに直接反映（角丸の重複だけは毎回除去してから再適用）
        for (var i = 0; i < sel.length; i++) {
            var it = sel[i];
            if (it && it.typename === 'GroupItem' && it.clipped) {
                removeRoundCornersEffect(it);
            }
        }

        var state = getInputState();
        processSelection(sel, state.fitMode, state.maskMode, state.anchorIndex, state.manualScaleVal, state.roundVal, state.applyRound, state.tweakX, state.tweakY);
        if (state.fitMode !== "manual") scaleInput.text = getCurrentScale(sel[0]);

        app.redraw();
    }

    radioCover.onClick = function () { setScaleRadio(this); updatePreview(); };
    radioContain.onClick = function () { setScaleRadio(this); updatePreview(); };
    radioNone.onClick = function () { setScaleRadio(this); updatePreview(); };
    radioManual.onClick = function () { setScaleRadio(this); updatePreview(); };

    radioSquare.onClick = function () {
        // 「正方形に」を選択したときは中央基準＆カバーに固定
        setScaleRadio(radioCover);
        for (var j = 0; j < anchorRadios.length; j++) anchorRadios[j].value = (j === 4);

        // 「正円」は「正方形に」のときのみ
        updateCircleAvailability();

        // 「正方形に」選択だけでは角丸値は変更しない（角丸値の自動設定は［正円］ON時のみ）
        updatePreview();
    };

    radioFitFrame.onClick = function () { updateCircleAvailability(); updatePreview(); };
    maskNone.onClick = function () { updateCircleAvailability(); updatePreview(); };
    checkRound.onClick = function () {
        // ［角丸］をOFFにしたら、クリップグループのアピアランスを消去して重複を防ぐ
        if (!checkRound.value) {
            clearAppearanceForClipGroups(sel);
        }
        updatePreview();
    };
    checkCircle.onClick = function () {
        // ONにしたら「角丸」を自動ONし、半径を幅/高さの半分に設定
        if (checkCircle.value) {
            checkRound.value = true;
            setRoundRadiusToHalfOfSquareMask();

            // 角丸の値（UI表示）→ pt に変換
            var radiusPt = valueInCurrentUnitToPt(parseFloat(roundInput.text) || 0);

            // 既存の LiveEffect があれば「半径だけ」更新（新規適用しない）
            var ok = updateRoundCornersRadiusOnly(sel[0], radiusPt);

            // 難しい場合（見つからない／形式が違う）は、アピアランス消去→再適用
            if (!ok) {
                clearAppearanceForClipGroups(sel);
                applyRoundCornersEffect(sel[0], radiusPt);
            }

            app.redraw();
            return;
        }

        // OFF にしただけでは処理を走らせない（新規に［角を丸くする］が適用されるのを防ぐ）
        app.redraw();
        return;
    };

    scaleInput.onChanging = function () {
        setScaleRadio(radioManual);
        updatePreview();
    };
    roundInput.onChanging = function () {
        checkRound.value = true;
        clearAppearanceForClipGroups(sel);
        updatePreview();
    };
    for (var i = 0; i < anchorRadios.length; i++) {
        (function (idx) {
            anchorRadios[idx].onClick = function () {
                setAnchorRadio(idx);
                resetTweakXY();
                updatePreview();
            };
        })(i);
    }

    okBtn.onClick = function () {
        var state = getInputState();

        // 選択を戻してから処理（複数選択にも対応）
        setSelection(originalSelection);

        // 実行時に1回：クリップグループのアピアランスを消去
        clearAppearanceForClipGroups(originalSelection);

        // 入力状態で最終適用
        processSelection(originalSelection, state.fitMode, state.maskMode, state.anchorIndex, state.manualScaleVal, state.roundVal, state.applyRound, state.tweakX, state.tweakY);

        app.redraw();
        win.close();
    };
    cancelBtn.onClick = function () {
        // Undo を使わないため、現状の反映状態のまま閉じる
        win.close();
    };

    // 実行時に1回だけ：クリップグループのアピアランスを消去
    clearAppearanceForClipGroups(sel);

    updatePreview();
    win.show();

    function processSelection(selection, fitMode, maskMode, anchorIndex, manualScale, roundVal, applyRound, tweakX, tweakY) {
        for (var i = 0; i < selection.length; i++) {
            var item = selection[i];
            if (item.typename === "GroupItem" && item.clipped) {
                fitContent(item, fitMode, maskMode, anchorIndex, manualScale, roundVal, applyRound, tweakX, tweakY);
            }
        }
    }

    function fitContent(groupItem, fitMode, maskMode, anchorIndex, manualScale, roundVal, applyRound, tweakX, tweakY) {
        var clipPath = null;
        var contents = [];
        for (var i = 0; i < groupItem.pageItems.length; i++) {
            var p = groupItem.pageItems[i];
            if (p.clipping) clipPath = p; else contents.push(p);
        }
        if (!clipPath || contents.length === 0) return;

        // 形状変更
        if (maskMode === "makeSquare") {
            var b = clipPath.geometricBounds;
            var w = b[2] - b[0], h = b[1] - b[3];
            var side = Math.min(w, h);
            var cx = b[0] + w / 2, cy = b[1] - h / 2;
            clipPath.width = side;
            clipPath.height = side;
            clipPath.position = [cx - side / 2, cy + side / 2];
        } else if (maskMode === "fitFrame") {
            var cb = getCombinedBounds(contents);
            if (cb) {
                clipPath.position = [cb[0], cb[1]];
                clipPath.width = cb[2] - cb[0];
                clipPath.height = cb[1] - cb[3];
                // ここでは return しない：後続の「フィットとスケール」処理を有効にする
            }
        }

        // 角丸エフェクト適用（対象をクリップグループに変更）
        if (applyRound) {
            applyRoundCornersEffect(groupItem, roundVal);
        }

        // 位置合わせは visibleBounds を基準に（上揃えがズレるケース対策）
        var frameBounds = clipPath.visibleBounds;
        var fW = frameBounds[2] - frameBounds[0];
        var fH = frameBounds[1] - frameBounds[3];

        for (var k = 0; k < contents.length; k++) {
            var content = contents[k];
            var cBounds = content.visibleBounds;
            var cW = cBounds[2] - cBounds[0];
            var cH = cBounds[1] - cBounds[3];

            var ratio = 1.0;
            if (fitMode === "cover") ratio = Math.max(fW / cW, fH / cH);
            else if (fitMode === "contain") ratio = Math.min(fW / cW, fH / cH);
            else if (fitMode === "manual") {
                var currentS = parseFloat(getCurrentScale(groupItem)) / 100;
                var targetS = manualScale / 100;
                ratio = (currentS === 0) ? 0 : targetS / currentS;
            }

            if (fitMode !== "none") {
                content.resize(ratio * 100, ratio * 100, true, true, true, true, ratio * 100);
                cBounds = content.visibleBounds;
                cW = cBounds[2] - cBounds[0]; cH = cBounds[1] - cBounds[3];
            }

            var col = anchorIndex % 3, row = Math.floor(anchorIndex / 3);
            var fRefX = (col === 0) ? frameBounds[0] : (col === 1) ? frameBounds[0] + fW / 2 : frameBounds[2];
            var fRefY = (row === 0) ? frameBounds[1] : (row === 1) ? frameBounds[1] - fH / 2 : frameBounds[3];
            var cRefX = (col === 0) ? cBounds[0] : (col === 1) ? cBounds[0] + cW / 2 : cBounds[2];
            var cRefY = (row === 0) ? cBounds[1] : (row === 1) ? cBounds[1] - cH / 2 : cBounds[3];
            content.translate((fRefX - cRefX) + tweakX, (fRefY - cRefY) + tweakY);
        }
    }

    function getCombinedBounds(items) {
        if (items.length === 0) return null;
        var b = items[0].geometricBounds;
        var l = b[0], t = b[1], r = b[2], bot = b[3];
        for (var i = 1; i < items.length; i++) {
            var nb = items[i].geometricBounds;
            if (nb[0] < l) l = nb[0]; if (nb[1] > t) t = nb[1];
            if (nb[2] > r) r = nb[2]; if (nb[3] < bot) bot = nb[3];
        }
        return [l, t, r, bot];
    }
})();
