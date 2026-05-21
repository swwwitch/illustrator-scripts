#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

ShuffleObjectColors.jsx

### 概要

- 選択オブジェクト（パス、グループ、複合シェイプ）の塗り・線カラーを再配色します。
- RGB／CMYK／グレースケール／特色／グラデーション／パターンに対応。
- 黒・白の除外、カラー比率（バランス）保持、ランダム順／元の順序の切替が可能です。
- 適用ボタンでダイアログを閉じずにプレビュー、キャンセルで元の状態に戻せます。
- 日本語／英語UIに対応。

### 主な機能

- 塗り・線の適用切替（各チェックボックスで個別に ON/OFF）
- 黒・白の除外（特色は tint 100% のときのみ判定）
- レジストレーション（見当合わせ色）は自動的に除外
- カラー比率（バランス）保持オプション
- ランダム順／元の順序の切替
- プレビュー適用（ダイアログを閉じずに結果を確認）

### 処理の流れ

1. 対象オブジェクトを収集
2. 使用カラーを抽出
3. プールを生成・並べ替え
4. カラーを再適用

### 更新履歴

- v1.0.0 (20240624) : 初期バージョン
- v1.0.1 (20240624) : バグ修正
- v1.0.2 (20240624) : ランダム／順番適用切替追加
- v1.0.3 (20240625) : ローカライズ調整
- v1.1 (20250708) : テキスト機能削除、構造整理

---

### Script Name:

ShuffleObjectColors.jsx

### Overview

- Reapplies fill and stroke colors to selected objects (paths, groups, compound shapes).
- Supports RGB, CMYK, Grayscale, Spot, Gradient, and Pattern colors.
- Options for excluding black/white, preserving color balance, and toggling random/sequential order.
- Apply button previews the result without closing the dialog; Cancel restores the original state.
- Supports Japanese and English UI.

### Main Features

- Toggle fill and stroke application (independent checkboxes)
- Exclude black and white (spot colors only when tint is 100%)
- Registration color is automatically excluded
- Preserve color balance option
- Random or sequential color application
- Live preview without closing the dialog

### Workflow

1. Collect target objects
2. Extract used colors
3. Create and shuffle color pool
4. Reapply colors

### Changelog

- v1.0.0 (20240624): Initial release
- v1.0.1 (20240624): Bug fixes
- v1.0.2 (20240624): Added random/sequential application
- v1.0.3 (20240625): Localization updates
- v1.1 (20250708): Removed text feature, refactored structure
*/

(function () {

    // =========================================
    // バージョンとローカライズ
    // =========================================

    var SCRIPT_VERSION = "v1.1";

    function getCurrentLang() {
        return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
    }

    var lang = getCurrentLang();

    /* ラベル定義 / Labels */
    var LABELS = {
        /* --- ダイアログ / Dialog --- */
        dialogTitle: { ja: "選択オブジェクトのカラーをシャッフル", en: "Shuffle Selected Object Colors" },

        /* --- パネル / Panels --- */
        panelTarget: { ja: "対象", en: "Target" },
        panelExclude: { ja: "除外", en: "Exclude" },

        /* --- 対象オプション / Target options --- */
        fill: { ja: "塗り", en: "Fill" },
        stroke: { ja: "線", en: "Stroke" },

        /* --- 除外オプション / Exclude options --- */
        black: { ja: "黒", en: "Black" },
        white: { ja: "白", en: "White" },

        /* --- 適用オプション / Apply options --- */
        random: { ja: "ランダム順", en: "Random order" },
        balance: { ja: "バランスを保持", en: "Preserve balance" },

        /* --- ボタン / Buttons --- */
        apply: { ja: "適用", en: "Apply" },
        cancel: { ja: "キャンセル", en: "Cancel" },

        /* --- ツールチップ / Tooltips --- */
        tipFill: { ja: "塗りカラーをシャッフルの対象にする", en: "Include fill colors in the shuffle" },
        tipStroke: { ja: "線カラーをシャッフルの対象にする", en: "Include stroke colors in the shuffle" },
        tipBlack: { ja: "黒（および近いカラー）は変更しない", en: "Leave black (and near-black) colors untouched" },
        tipWhite: { ja: "白は変更しない", en: "Leave white untouched" },
        tipRandom: { ja: "ON：ランダム順／OFF：元の順序で適用", en: "ON: random order / OFF: keep original order" },
        tipBalance: { ja: "同じカラーの出現回数を保持して配色", en: "Keep the original frequency of each color" },
        tipApply: { ja: "ダイアログを閉じずにプレビュー", en: "Preview without closing the dialog" },

        /* --- アラート / Alerts --- */
        alertSelect: { ja: "オブジェクトを選択してください。", en: "Please select some objects." },
        alertEmpty: { ja: "使用可能なカラーが見つかりません（白と黒以外）。", en: "No usable colors found (excluding white and black)." },
        alertChoice: { ja: "塗りまたは線のいずれかを選択してください。", en: "Please select either fill or stroke." }
    };

    function L(key) {
        return LABELS[key][lang];
    }

    // =========================================
    // UI 状態 / UI state
    // =========================================

    var fillCheckbox, strokeCheckbox;
    var blackCheckbox, whiteCheckbox;
    var randomCheckbox, balanceCheckbox;
    var previewApplied = false;

    main();

    // =========================================
    // メイン処理 / Main
    // =========================================

    function main() {
        if (app.documents.length === 0 || !app.activeDocument.selection || app.activeDocument.selection.length === 0) {
            alert(L("alertSelect"));
            return;
        }
        showDialog();
    }

    function showDialog() {
        var dialog = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
        dialog.orientation = "row";
        dialog.alignChildren = "top";

        var leftGroup = dialog.add("group");
        leftGroup.orientation = "column";
        leftGroup.alignChildren = "left";

        buildTargetPanel(leftGroup);
        buildExcludePanel(leftGroup);

        randomCheckbox = leftGroup.add("checkbox", undefined, L("random"));
        randomCheckbox.value = true;
        randomCheckbox.helpTip = L("tipRandom");

        balanceCheckbox = leftGroup.add("checkbox", undefined, L("balance"));
        balanceCheckbox.value = true;
        balanceCheckbox.helpTip = L("tipBalance");

        buildButtons(dialog);

        dialog.show();
    }

    function buildTargetPanel(parent) {
        var panel = parent.add("panel", undefined, L("panelTarget"));
        panel.preferredSize.width = 120;
        panel.orientation = "column";
        panel.alignChildren = "left";
        panel.margins = [10, 25, 10, 10];

        var row = panel.add("group");
        row.orientation = "row";
        row.alignChildren = "left";

        fillCheckbox = row.add("checkbox", undefined, L("fill"));
        fillCheckbox.value = true;
        fillCheckbox.helpTip = L("tipFill");

        strokeCheckbox = row.add("checkbox", undefined, L("stroke"));
        strokeCheckbox.value = true;
        strokeCheckbox.helpTip = L("tipStroke");
    }

    function buildExcludePanel(parent) {
        var panel = parent.add("panel", undefined, L("panelExclude"));
        panel.preferredSize.width = 120;
        panel.orientation = "column";
        panel.alignChildren = "left";
        panel.margins = [10, 25, 10, 10];

        var row = panel.add("group");
        row.orientation = "row";
        row.alignChildren = "left";

        blackCheckbox = row.add("checkbox", undefined, L("black"));
        blackCheckbox.value = true;
        blackCheckbox.helpTip = L("tipBlack");

        whiteCheckbox = row.add("checkbox", undefined, L("white"));
        whiteCheckbox.value = true;
        whiteCheckbox.helpTip = L("tipWhite");
    }

    function buildButtons(dialog) {
        var rightGroup = dialog.add("group");
        rightGroup.orientation = "column";
        rightGroup.alignChildren = "right";

        /* Mac 規約: 縦並びは OK が上 / Vertical: OK on top */
        var okBtn = rightGroup.add("button", undefined, "OK");
        okBtn.preferredSize.width = 80;

        var cancelBtn = rightGroup.add("button", undefined, L("cancel"));
        cancelBtn.preferredSize.width = 80;

        // スペーサー（縦に伸びる）
        var verticalSpacer = rightGroup.add("statictext", undefined, "");
        verticalSpacer.alignment = ["fill", "fill"];

        // 適用ボタン用グループ（右下）
        var applyGroup = rightGroup.add("group");
        applyGroup.orientation = "column";
        applyGroup.alignChildren = ["fill", "bottom"];

        var applyBtn = applyGroup.add("button", undefined, L("apply"));
        applyBtn.preferredSize.width = 80;
        applyBtn.helpTip = L("tipApply");

        applyBtn.onClick = function () {
            applyBtn.enabled = false;
            applyWithPreview();
            applyBtn.enabled = true;
        };

        okBtn.onClick = function () {
            applyWithPreview();
            dialog.close(1);
        };

        cancelBtn.onClick = function () {
            if (previewApplied) {
                app.undo();
                previewApplied = false;
            }
            dialog.close(0);
        };
    }

    function applyWithPreview() {
        if (previewApplied) {
            app.undo();
            previewApplied = false;
        }
        if (applyColors()) {
            previewApplied = true;
            app.redraw();
        }
    }

    // =========================================
    // カラー収集 / Color collection
    // =========================================

    function applyColors() {
        var changeFill = fillCheckbox.value;
        var changeStroke = strokeCheckbox.value;
        var excludeBlack = blackCheckbox.value;
        var excludeWhite = whiteCheckbox.value;
        var balancePreserve = balanceCheckbox.value;
        var randomize = randomCheckbox.value;

        if (!changeFill && !changeStroke) {
            alert(L("alertChoice"));
            return false;
        }

        var pathItems = [];
        collectPathItems(app.activeDocument.selection, pathItems);

        var colorPool = buildColorPool(pathItems, changeFill, changeStroke, excludeBlack, excludeWhite, balancePreserve);
        if (colorPool.length === 0) {
            alert(L("alertEmpty"));
            return false;
        }

        var shuffled = randomize ? shuffleArray(colorPool) : colorPool;
        reapplyColors(pathItems, shuffled, changeFill, changeStroke, excludeBlack, excludeWhite);
        return true;
    }

    function buildColorPool(pathItems, changeFill, changeStroke, excludeBlack, excludeWhite, balancePreserve) {
        var colorCountMap = {};
        var uniqueColors = {};

        for (var i = 0; i < pathItems.length; i++) {
            var pathItem = pathItems[i];
            if (changeFill && pathItem.filled) {
                addColorToPool(pathItem.fillColor, colorCountMap, uniqueColors, excludeBlack, excludeWhite, balancePreserve);
            }
            if (changeStroke && pathItem.stroked && pathItem.strokeWidth > 0) {
                addColorToPool(pathItem.strokeColor, colorCountMap, uniqueColors, excludeBlack, excludeWhite, balancePreserve);
            }
        }

        var pool = [];
        if (balancePreserve) {
            for (var countKey in colorCountMap) {
                var entry = colorCountMap[countKey];
                for (var r = 0; r < entry.count; r++) {
                    pool.push(cloneColor(entry.color));
                }
            }
        } else {
            for (var uniqueKey in uniqueColors) {
                pool.push(cloneColor(uniqueColors[uniqueKey]));
            }
        }
        return pool;
    }

    function addColorToPool(rawColor, colorCountMap, uniqueColors, excludeBlack, excludeWhite, balancePreserve) {
        var color = cloneColor(rawColor);
        if (!color) return;
        if (excludeBlack && isNearBlack(color)) return;
        if (excludeWhite && isPureWhite(color)) return;

        var colorKey = getColorKey(color);
        if (!colorKey) return;

        if (!uniqueColors[colorKey]) {
            uniqueColors[colorKey] = cloneColor(color);
        }
        if (balancePreserve) {
            if (!colorCountMap[colorKey]) {
                colorCountMap[colorKey] = { color: cloneColor(color), count: 0 };
            }
            colorCountMap[colorKey].count++;
        }
    }

    // 対象パス収集（再帰） / Collect path items recursively
    function collectPathItems(items, results) {
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (item.typename === "GroupItem") {
                collectPathItems(item.pageItems, results);
            } else if (item.typename === "CompoundPathItem") {
                collectPathItems(item.pathItems, results);
            } else if (item.typename === "PathItem") {
                results.push(item);
            }
        }
    }

    // =========================================
    // カラー適用 / Color application
    // =========================================

    function reapplyColors(pathItems, colorPool, changeFill, changeStroke, excludeBlack, excludeWhite) {
        var colorIndex = 0;
        for (var i = 0; i < pathItems.length; i++) {
            var pathItem = pathItems[i];
            colorIndex = applyToSlot(pathItem, true, changeFill, colorPool, colorIndex, excludeBlack, excludeWhite);
            colorIndex = applyToSlot(pathItem, false, changeStroke, colorPool, colorIndex, excludeBlack, excludeWhite);
        }
    }

    function applyToSlot(pathItem, isFill, changeFlag, colorPool, colorIndex, excludeBlack, excludeWhite) {
        if (!changeFlag) return colorIndex;

        var currentColor = cloneColor(isFill ? pathItem.fillColor : pathItem.strokeColor);
        if (!currentColor) return colorIndex;
        if (excludeBlack && isNearBlack(currentColor)) return colorIndex;
        if (excludeWhite && isPureWhite(currentColor)) return colorIndex;

        var nextColor = cloneColor(colorPool[colorIndex % colorPool.length]);
        if (isFill) {
            pathItem.filled = true;
            pathItem.fillColor = nextColor;
        } else {
            pathItem.stroked = true;
            pathItem.strokeColor = nextColor;
        }
        return colorIndex + 1;
    }

    // =========================================
    // ユーティリティ / Utilities
    // =========================================

    // Fisher–Yates シャッフル / Fisher–Yates shuffle
    function shuffleArray(arr) {
        var shuffled = arr.slice();
        for (var i = shuffled.length - 1; i > 0; i--) {
            var j = Math.floor(Math.random() * (i + 1));
            var temp = shuffled[i];
            shuffled[i] = shuffled[j];
            shuffled[j] = temp;
        }
        return shuffled;
    }

    function getColorKey(color) {
        if (color.typename === "RGBColor") {
            return "rgb:" + color.red + "," + color.green + "," + color.blue;
        }
        if (color.typename === "CMYKColor") {
            return "cmyk:" + color.cyan + "," + color.magenta + "," + color.yellow + "," + color.black;
        }
        if (color.typename === "GrayColor") {
            return "gray:" + color.gray;
        }
        if (color.typename === "SpotColor") {
            return "spot:" + color.spot.name + "@" + color.tint;
        }
        if (color.typename === "GradientColor") {
            return "grad:" + color.gradient.name;
        }
        if (color.typename === "PatternColor") {
            return "pat:" + color.pattern.name;
        }
        return null;
    }

    // カラー複製（RGB/CMYK/Gray/Spot/Gradient/Pattern 対応） / Clone color (RGB/CMYK/Gray/Spot/Gradient/Pattern)
    function cloneColor(color) {
        if (!color || !color.typename) return null;
        if (color.typename === "RGBColor") {
            var rgb = new RGBColor();
            rgb.red = color.red;
            rgb.green = color.green;
            rgb.blue = color.blue;
            return rgb;
        }
        if (color.typename === "CMYKColor") {
            var cmyk = new CMYKColor();
            cmyk.cyan = color.cyan;
            cmyk.magenta = color.magenta;
            cmyk.yellow = color.yellow;
            cmyk.black = color.black;
            return cmyk;
        }
        if (color.typename === "GrayColor") {
            var gr = new GrayColor();
            gr.gray = color.gray;
            return gr;
        }
        if (color.typename === "SpotColor") {
            /* 見当合わせ色（レジストレーション）はシャッフル対象外 / Skip registration */
            if (color.spot.colorType === ColorModel.REGISTRATION) return null;
            var sc = new SpotColor();
            sc.spot = color.spot;
            sc.tint = color.tint;
            return sc;
        }
        if (color.typename === "GradientColor") {
            var gc = new GradientColor();
            gc.gradient = color.gradient;
            try { gc.angle = color.angle; } catch (e) { }
            try { gc.length = color.length; } catch (e) { }
            try { gc.origin = color.origin; } catch (e) { }
            try { gc.hiliteAngle = color.hiliteAngle; } catch (e) { }
            try { gc.hiliteLength = color.hiliteLength; } catch (e) { }
            try { gc.matrix = color.matrix; } catch (e) { }
            return gc;
        }
        if (color.typename === "PatternColor") {
            var pc = new PatternColor();
            pc.pattern = color.pattern;
            try { pc.matrix = color.matrix; } catch (e) { }
            try { pc.shiftAngle = color.shiftAngle; } catch (e) { }
            try { pc.shiftDistance = color.shiftDistance; } catch (e) { }
            try { pc.reflect = color.reflect; } catch (e) { }
            try { pc.reflectAngle = color.reflectAngle; } catch (e) { }
            try { pc.rotation = color.rotation; } catch (e) { }
            try { pc.scaleFactor = color.scaleFactor; } catch (e) { }
            try { pc.shearAngle = color.shearAngle; } catch (e) { }
            try { pc.shearAxis = color.shearAxis; } catch (e) { }
            return pc;
        }
        return null;
    }

    // 黒近似判定 / Near-black detection
    function isNearBlack(color) {
        if (color.typename === "RGBColor") {
            return color.red <= 51 && color.green <= 51 && color.blue <= 51;
        }
        if (color.typename === "CMYKColor") {
            return color.black >= 0.8 || (color.cyan <= 0.2 && color.magenta <= 0.2 && color.yellow <= 0.2 && color.black >= 0.7);
        }
        if (color.typename === "GrayColor") {
            return color.gray >= 80;
        }
        if (color.typename === "SpotColor") {
            /* tint 100% のときだけ基底色で判定 / Only when tint is 100% */
            return color.tint >= 99.999 && isNearBlack(color.spot.color);
        }
        return false;
    }

    // 純白判定 / Pure-white detection
    function isPureWhite(color) {
        if (color.typename === "RGBColor") {
            return color.red === 255 && color.green === 255 && color.blue === 255;
        }
        if (color.typename === "CMYKColor") {
            return color.cyan === 0 && color.magenta === 0 && color.yellow === 0 && color.black === 0;
        }
        if (color.typename === "GrayColor") {
            return color.gray === 0;
        }
        if (color.typename === "SpotColor") {
            /* tint 0% は紙色（白） / Tint 0% = paper (white) */
            if (color.tint <= 0.001) return true;
            return color.tint >= 99.999 && isPureWhite(color.spot.color);
        }
        return false;
    }

})();