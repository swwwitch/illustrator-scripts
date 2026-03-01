/*
  BrightnessAdjuster_History.jsx
  - Localized: Japanese and English
  - Slider: -50 to +50 / Text: -100 to +100
  - R Key: Reset / Arrow Keys: Up/Down
  - History: Single Undo Step (Prevents history pollution)
*/

(function() {
    var lang = ($.locale && $.locale.indexOf("ja") > -1) ? "ja" : "en";
    var uiText = {
        ja: {
            title: "輝度調整",
            noSelection: "オブジェクトを選択してください。",
            reset: "リセット (R)",
            cancel: "キャンセル",
            ok: "OK"
        },
        en: {
            title: "Brightness Adjustment",
            noSelection: "Please select an object.",
            reset: "Reset (R)",
            cancel: "Cancel",
            ok: "OK"
        }
    }[lang];

    if (app.documents.length === 0) return;
    var doc = app.activeDocument;
    var sel = doc.selection;

    if (sel.length === 0) {
        alert(uiText.noSelection);
        return;
    }

    // --- 1. 元の色情報を収集 ---
    var editItems = [];
    function collectColors(items) {
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (item.typename === "GroupItem") {
                collectColors(item.pageItems);
            } else if (item.typename === "CompoundPathItem") {
                collectColors(item.pathItems);
            } else if (item.typename === "PathItem") {
                var data = { obj: item, fill: null, stroke: null };
                if (item.filled) data.fill = cloneColor(item.fillColor);
                if (item.stroked) data.stroke = cloneColor(item.strokeColor);
                if (data.fill || data.stroke) editItems.push(data);
            } else if (item.typename === "TextFrame") {
                try {
                    var tData = { obj: item, fill: null, stroke: null };
                    tData.fill = cloneColor(item.textRange.characterAttributes.fillColor);
                    tData.stroke = cloneColor(item.textRange.characterAttributes.strokeColor);
                    editItems.push(tData);
                } catch(e) {}
            }
        }
    }

    function cloneColor(color) {
        if (!color || color.typename === "GradientColor" || color.typename === "PatternColor") return null;
        if (color.typename === "CMYKColor") {
            var c = new CMYKColor(); c.cyan=color.cyan; c.magenta=color.magenta; c.yellow=color.yellow; c.black=color.black; return c;
        } else if (color.typename === "RGBColor") {
            var c = new RGBColor(); c.red=color.red; c.green=color.green; c.blue=color.blue; return c;
        } else if (color.typename === "GrayColor") {
            var c = new GrayColor(); c.gray=color.gray; return c;
        } else if (color.typename === "SpotColor") {
            var c = new SpotColor(); c.spot=color.spot; c.tint=color.tint; return c;
        }
        return null;
    }

    collectColors(sel);
    if (editItems.length === 0) return;

    // --- 2. UIの作成 ---
    var win = new Window("dialog", uiText.title);
    win.alignChildren = "fill";

    var sliderGroup = win.add("group");
    sliderGroup.add("statictext", undefined, "-50");
    var slider = sliderGroup.add("slider", undefined, 0, -50, 50);
    slider.preferredSize.width = 300;
    sliderGroup.add("statictext", undefined, "+50");

    var valGroup = win.add("group");
    valGroup.alignment = "center";
    var etValue = valGroup.add("edittext", undefined, "0");
    etValue.characters = 5;
    valGroup.add("statictext", undefined, "%");

    var btnFooter = win.add("group");
    btnFooter.alignment = "fill";
    var leftBtns = btnFooter.add("group");
    var btnReset = leftBtns.add("button", undefined, uiText.reset);
    var rightBtns = btnFooter.add("group");
    rightBtns.alignment = ["right", "center"];
    var btnCancel = rightBtns.add("button", undefined, uiText.cancel, {name: "cancel"});
    var btnOk = rightBtns.add("button", undefined, uiText.ok, {name: "ok"});

    // --- 3. 更新・プレビュー管理ロジック ---
    var isFirstChange = true;

    function updateUI(val) {
        if (val > 100) val = 100;
        if (val < -100) val = -100;
        val = Math.round(val);

        etValue.text = val;
        var sliderVal = (val > 50) ? 50 : (val < -50) ? -50 : val;
        slider.value = sliderVal;

        // 【重要】ヒストリーの汚染を防ぐ処理
        // 2回目以降の更新では、直前の変更をUndoしてから新しい値を適用する
        if (!isFirstChange) {
            app.undo();
        }
        
        updateColors(val);
        isFirstChange = false;
        app.redraw();
    }

    function doReset() {
        updateUI(0);
    }

    // --- 4. イベントハンドラ ---
    win.addEventListener("keydown", function(event) {
        if (event.keyName === "R") { doReset(); event.preventDefault(); }
    });

    etValue.addEventListener("keydown", function(event) {
        var value = Number(etValue.text);
        if (isNaN(value)) return;
        var delta = (ScriptUI.environment.keyboardState.shiftKey) ? 10 : 1;
        if (event.keyName == "Up") { event.preventDefault(); updateUI(value + delta); }
        else if (event.keyName == "Down") { event.preventDefault(); updateUI(value - delta); }
    });

    slider.onChanging = function() { updateUI(slider.value); };
    etValue.onChange = function() {
        var val = Number(this.text);
        if (isNaN(val)) val = 0;
        updateUI(val);
    };

    btnReset.onClick = doReset;

    btnCancel.onClick = function() {
        if (!isFirstChange) app.undo(); // プレビュー中の変更をすべて破棄
        win.close();
    };

    btnOk.onClick = function() {
        win.close(); // OK時は現在のプレビュー状態を確定（Undo履歴には1件だけ残る）
    };

    etValue.active = true;

    // --- 5. 色計算ロジック ---
    function updateColors(percent) {
        var ratio = percent / 100;
        for (var i = 0; i < editItems.length; i++) {
            var itemData = editItems[i];
            if (itemData.fill) applyColor(itemData.obj, "fillColor", calcBrightness(itemData.fill, ratio));
            if (itemData.stroke) applyColor(itemData.obj, "strokeColor", calcBrightness(itemData.stroke, ratio));
        }
    }

    function applyColor(item, prop, color) {
        if (item.typename === "TextFrame") item.textRange.characterAttributes[prop] = color;
        else item[prop] = color;
    }

    function calcBrightness(baseColor, ratio) {
        var absRatio = Math.abs(ratio);
        if (baseColor.typename === "CMYKColor") {
            var c = new CMYKColor();
            var keys = ["cyan","magenta","yellow","black"];
            for(var i=0; i<4; i++) {
                var k = keys[i];
                if (ratio > 0) c[k] = baseColor[k] * (1 - absRatio);
                else c[k] = baseColor[k] + (100 - baseColor[k]) * absRatio;
            }
            return c;
        } else if (baseColor.typename === "RGBColor") {
            var c = new RGBColor();
            var keys = ["red","green","blue"];
            for(var i=0; i<3; i++) {
                var k = keys[i];
                if (ratio > 0) c[k] = baseColor[k] + (255 - baseColor[k]) * absRatio;
                else c[k] = baseColor[k] * (1 - absRatio);
            }
            return c;
        } else if (baseColor.typename === "GrayColor") {
            var c = new GrayColor();
            if (ratio > 0) c.gray = baseColor.gray * (1 - absRatio);
            else c.gray = baseColor.gray + (100 - baseColor.gray) * absRatio;
            return c;
        } else if (baseColor.typename === "SpotColor") {
            var c = new SpotColor(); c.spot = baseColor.spot;
            if (ratio > 0) c.tint = baseColor.tint * (1 - absRatio);
            else c.tint = baseColor.tint + (100 - baseColor.tint) * absRatio;
            return c;
        }
        return baseColor;
    }

    win.show();
})();
