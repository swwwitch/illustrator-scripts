#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

FitShapeToContent.jsx

 * スクリプトの概要:
 * 選択したテキストまたはグループと図形をもとに、図形をコンテンツ中央へ揃えながらサイズ調整するスクリプトです。
 * - テキストは複製してアウトライン化し、グループは複製して境界を計測
 * - パディング（幅・高さの加算値）をダイアログでプレビュー調整
 * - 幅と高さは「連動」で同値入力に対応
 * - 角丸はプレビュー専用オブジェクトで差し替え表示し、効果の累積を防止
 * - キャンセル時は元の図形をそのまま復元します
 *
 * Overview:
 * Fit and center a shape to selected text or group content.
 * - Duplicate text and outline it for measurement; duplicate groups and measure their bounds
 * - Preview padding values for width and height in a dialog
 * - Width and height can be linked and edited together
 * - Rounded corners are previewed with a dedicated preview object to avoid effect accumulation
 * - Cancelling restores the original shape as-is
 *
 * 作成日: 2026-03-25
 * 更新日: 2026-03-25
 * Created: 2026-03-25
 * Updated: 2026-03-25 (Keep the original shape and generate preview from a cleared preview base shape)
 */

(function () {

    // =========================================
    // バージョンとローカライズ / Version and Localization
    // =========================================

    var SCRIPT_VERSION = "v1.0";

    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var lang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */

    // =========================================
    // 単位 / Units
    // =========================================

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


    var LABELS = {
        dialogTitle: { ja: "コンテンツに合わせて図形を調整", en: "Fit Shape to Content" },
        panelPadding: { ja: "パディング", en: "Padding" },
        panelCorner: { ja: "角丸", en: "Rounded Corners" },
        labelWidth: { ja: "幅:", en: "Width:" },
        labelHeight: { ja: "高さ:", en: "Height:" },
        labelRadius: { ja: "半径:", en: "Radius:" },
        labelLink: { ja: "連動", en: "Link" },
        btnCancel: { ja: "キャンセル", en: "Cancel" },
        btnOK: { ja: "OK", en: "OK" },
        alertNoDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        alertSelectTwo: { ja: "テキストまたはグループと図形を1つずつ（計2つ）選択して実行してください。", en: "Select one text or group item and one shape item (2 items total)." },
        alertSelectError: { ja: "選択エラー", en: "Selection Error" },
        alertOnlyOneContent: { ja: "テキストまたはグループは1つだけ選択してください。", en: "Select only one text or group item." },
        alertOnlyOneShape: { ja: "図形は1つだけ選択してください。", en: "Select only one shape item." },
        alertUnsupportedSelection: { ja: "対応していない選択オブジェクトが含まれています。テキストまたはグループと、パスまたは複合パスを選択してください。", en: "Unsupported selection included. Select a text or group item, and a path or compound path." },
        alertNeedContentAndShape: { ja: "テキストまたはグループ1つと、パスまたは複合パス1つを選択してください。", en: "Select one text or group item and one path or compound path." },
        alertClippingGroupNotSupported: { ja: "クリッピンググループの計測には未対応です。クリッピングを解除するか、計測対象を単純なグループにしてください。", en: "Clipping groups are not supported for measurement. Release the clipping mask or use a simple group as the content item." },
        alertMeasureFailed: { ja: "コンテンツの計測に失敗しました。選択内容を確認してください。", en: "Failed to measure the content item. Check the selected objects and try again." },
        alertInvalidNumber: { ja: "数値を入力してください。", en: "Enter a numeric value." },
        alertInvalidRadius: { ja: "角丸の半径は0以上の数値を入力してください。", en: "Enter a radius value of 0 or greater." }
    };

    function L(key) {
        var entry = LABELS[key];
        if (!entry) return key;
        return entry[lang] || entry.ja || key;
    }


    /* 単位ラベルを取得 / Get unit label */
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

    /* ↑↓キーで数値を変更 / Change numeric value with arrow keys */
    function changeValueByArrowKey(editText, onChanged) {
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
                    event.preventDefault();
                } else {
                    return;
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                if (event.keyName == "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= delta;
                    event.preventDefault();
                } else {
                    return;
                }
            } else {
                if (event.keyName == "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= delta;
                    event.preventDefault();
                } else {
                    return;
                }
            }

            if (keyboard.altKey) {
                value = Math.round(value * 10) / 10;
            } else {
                value = Math.round(value);
            }

            editText.text = value;
            if (typeof onChanged === "function") onChanged();
        });
    }

    /* 数値を解析し、無効時は既定値を返す / Parse number or return default */
    function parseNumberOrDefault(value, defaultValue) {
        var num = parseFloat(value);
        return isNaN(num) ? defaultValue : num;
    }

    /* 入力値を検証 / Validate numeric input */
    function validateNumericField(editText, allowNegative, titleKey, messageKey) {
        var value = parseFloat(editText.text);
        if (isNaN(value) || (!allowNegative && value < 0)) {
            alert(L(messageKey), L(titleKey));
            editText.active = true;
            editText.selection = [0, editText.text.length];
            return null;
        }
        return value;
    }

    /* オブジェクトの境界情報を取得 / Get bounds information from item */
    function getBoundsFromItem(item) {
        var vb = item.visibleBounds;
        return {
            left: vb[0],
            top: vb[1],
            right: vb[2],
            bottom: vb[3],
            width: vb[2] - vb[0],
            height: vb[1] - vb[3],
            centerX: vb[0] + ((vb[2] - vb[0]) / 2),
            centerY: vb[1] - ((vb[1] - vb[3]) / 2)
        };
    }

    /* コンテンツ対象か判定 / Check whether item is valid content */
    function isContentItem(item) {
        return item && (item.typename === "TextFrame" || item.typename === "GroupItem");
    }

    /* 図形対象か判定 / Check whether item is a valid shape */
    function isShapeItem(item) {
        return item && (
            item.typename === "PathItem" ||
            item.typename === "CompoundPathItem"
        );
    }

    /* カラー情報を複製 / Clone color value */
    function cloneColorValue(color) {
        if (!color) return null;

        var cloned;
        switch (color.typename) {
            case "RGBColor":
                cloned = new RGBColor();
                cloned.red = color.red;
                cloned.green = color.green;
                cloned.blue = color.blue;
                return cloned;
            case "CMYKColor":
                cloned = new CMYKColor();
                cloned.cyan = color.cyan;
                cloned.magenta = color.magenta;
                cloned.yellow = color.yellow;
                cloned.black = color.black;
                return cloned;
            case "GrayColor":
                cloned = new GrayColor();
                cloned.gray = color.gray;
                return cloned;
            case "SpotColor":
                cloned = new SpotColor();
                cloned.spot = color.spot;
                cloned.tint = color.tint;
                return cloned;
            case "PatternColor":
                cloned = new PatternColor();
                cloned.pattern = color.pattern;
                return cloned;
            case "GradientColor":
                cloned = new GradientColor();
                cloned.gradient = color.gradient;
                cloned.angle = color.angle;
                cloned.length = color.length;
                cloned.matrix = color.matrix;
                cloned.origin = color.origin;
                cloned.hiliteAngle = color.hiliteAngle;
                cloned.hiliteLength = color.hiliteLength;
                return cloned;
            case "NoColor":
                return new NoColor();
            default:
                return color;
        }
    }

    /* 図形スタイルを退避 / Capture shape style */
    function captureShapeStyle(shapeItem) {
        return {
            filled: !!shapeItem.filled,
            fillColor: shapeItem.filled ? cloneColorValue(shapeItem.fillColor) : null,
            stroked: !!shapeItem.stroked,
            strokeColor: shapeItem.stroked ? cloneColorValue(shapeItem.strokeColor) : null,
            strokeWidth: shapeItem.strokeWidth,
            opacity: shapeItem.opacity
        };
    }

    /* 図形スタイルを復元 / Restore shape style */
    function restoreShapeStyle(shapeItem, styleInfo) {
        if (!shapeItem || !styleInfo) return;

        shapeItem.opacity = styleInfo.opacity;

        shapeItem.filled = styleInfo.filled;
        if (styleInfo.filled && styleInfo.fillColor) {
            shapeItem.fillColor = cloneColorValue(styleInfo.fillColor);
        }

        shapeItem.stroked = styleInfo.stroked;
        if (styleInfo.stroked && styleInfo.strokeColor) {
            shapeItem.strokeColor = cloneColorValue(styleInfo.strokeColor);
            shapeItem.strokeWidth = styleInfo.strokeWidth;
        }
    }

    /* 選択を配列として複製 / Copy current selection into an array */
    function copySelectionItems(doc) {
        var copied = [];
        if (!doc || !doc.selection) return copied;
        for (var i = 0; i < doc.selection.length; i++) {
            copied.push(doc.selection[i]);
        }
        return copied;
    }

    /* クリッピンググループか判定 / Check whether group is a clipping group */
    function isClippingGroupItem(item) {
        return !!(item && item.typename === "GroupItem" && item.clipped);
    }

    /* アクションでアピアランスをクリア / Clear appearance by embedded action */
    function clearAppearanceByAction(targetItem) {
        if (!targetItem) return;

        var str = '/version 3' + '/name [ 10' + ' 417070656172616e6365' + ']' + '/isOpen 1' + '/actionCount 1' + '/action-1 {' + ' /name [ 5' + ' 636c656172' + ' ]' + ' /keyIndex 0' + ' /colorIndex 0' + ' /isOpen 1' + ' /eventCount 1' + ' /event-1 {' + ' /useRulersIn1stQuadrant 0' + ' /internalName (ai_plugin_appearance)' + ' /localizedName [ 18' + ' e382a2e38394e382a2e383a9e383b3e382b9' + ' ]' + ' /isOpen 1' + ' /isOn 1' + ' /hasDialog 0' + ' /parameterCount 1' + ' /parameter-1 {' + ' /key 1835363957' + ' /showInPalette 4294967295' + ' /type (enumerated)' + ' /name [ 27' + ' e382a2e38394e382a2e383a9e383b3e382b9e38292e6b688e58ebb' + ' ]' + ' /value 6' + ' }' + ' }' + '}';

        var prevSelection = copySelectionItems(app.activeDocument);
        var tempActionPath = Folder.temp.fsName + '/Appearance_clear_' + (new Date().getTime()) + '_' + Math.floor(Math.random() * 100000) + '.aia';
        var f = new File(tempActionPath);
        var originalStyle = captureShapeStyle(targetItem);

        try {
            app.activeDocument.selection = null;
            targetItem.selected = true;

            f.open('w');
            f.write(str);
            f.close();

            app.loadAction(f);
            f.remove();
            app.doScript('clear', 'Appearance', false);
            restoreShapeStyle(targetItem, originalStyle);
        } finally {
            try {
                app.unloadAction('Appearance', '');
            } catch (e) { }
            try {
                if (f && f.exists) {
                    f.remove();
                }
            } catch (e) { }
            try {
                app.activeDocument.selection = null;
                for (var i = 0; i < prevSelection.length; i++) {
                    prevSelection[i].selected = true;
                }
            } catch (e) { }
        }
    }

    /* ダイアログを表示 / Show dialog */
    function showDialog(outWidth, outHeight, previewBaseShapeItem, contentCenterX, contentCenterY) {
        var rulerUnit = getUnitLabel(app.preferences.getIntegerPreference("rulerType"), "rulerType");
        var radiusUnit = "pt";

        var currentPreviewItem = null;

        function removePreviewItem() {
            if (currentPreviewItem) {
                try {
                    currentPreviewItem.remove();
                } catch (e) { }
                currentPreviewItem = null;
            }
        }

        function buildPreview(addW, addH, radius) {
            if (isNaN(addW)) addW = 0;
            if (isNaN(addH)) addH = 0;
            if (isNaN(radius) || radius < 0) radius = 0;

            removePreviewItem();

            currentPreviewItem = previewBaseShapeItem.duplicate();
            previewBaseShapeItem.hidden = true;
            currentPreviewItem.hidden = false;

            var newWidth = outWidth + addW;
            var newHeight = outHeight + addH;
            currentPreviewItem.width = newWidth;
            currentPreviewItem.height = newHeight;
            currentPreviewItem.left = contentCenterX - (newWidth / 2);
            currentPreviewItem.top = contentCenterY + (newHeight / 2);

            if (radius > 0) {
                var xml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + radius + ' "/></LiveEffect>';
                currentPreviewItem.applyEffect(xml);
            }

            app.redraw();
        }

        var win = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
        win.orientation = "column";
        win.alignChildren = ["left", "top"];

        var labelWidth = 40;

        var panel = win.add("panel", undefined, L("panelPadding"));
        panel.orientation = "row";
        panel.alignChildren = ["left", "center"];
        panel.margins = [15, 20, 15, 10];

        // 左カラム：幅・高さ入力 / Left column: width and height inputs
        var colLeft = panel.add("group");
        colLeft.orientation = "column";
        colLeft.alignChildren = ["left", "top"];

        var groupW = colLeft.add("group");
        var labelW = groupW.add("statictext", undefined, L("labelWidth"));
        labelW.preferredSize = [labelWidth, -1];
        labelW.justify = "right";
        var inputW = groupW.add("edittext", undefined, "20");
        inputW.characters = 5;
        groupW.add("statictext", undefined, rulerUnit);

        var groupH = colLeft.add("group");
        var labelH = groupH.add("statictext", undefined, L("labelHeight"));
        labelH.preferredSize = [labelWidth, -1];
        labelH.justify = "right";
        var inputH = groupH.add("edittext", undefined, "20");
        inputH.characters = 5;
        groupH.add("statictext", undefined, rulerUnit);

        // 右カラム：連動チェックボックス / Right column: link checkbox
        var colRight = panel.add("group");
        colRight.orientation = "column";
        colRight.alignChildren = ["left", "center"];

        var linkCheck = colRight.add("checkbox", undefined, L("labelLink"));
        linkCheck.value = true;

        // 初期状態：連動ONなので高さをディム表示 / Initial state: height field disabled when linked
        inputH.enabled = false;
        inputH.text = inputW.text;

        // プレビュー更新関数 / Update preview
        function updatePreview() {
            var addW = parseNumberOrDefault(inputW.text, 0);
            var addH = parseNumberOrDefault(inputH.text, 0);
            var radius = parseNumberOrDefault(inputR.text, 0);
            if (radius < 0) radius = 0;

            buildPreview(addW, addH, radius);
        }

        // 連動チェック変更時 / When link checkbox changes
        linkCheck.onClick = function () {
            if (linkCheck.value) {
                inputH.text = inputW.text;
                inputH.enabled = false;
                updatePreview();
            } else {
                inputH.enabled = true;
            }
        };

        // 連動同期＋プレビュー更新 / Sync linked values and update preview
        function syncAndPreviewW() {
            if (linkCheck.value) {
                inputH.text = inputW.text;
            }
            updatePreview();
        }
        function syncAndPreviewH() {
            if (linkCheck.value) {
                inputW.text = inputH.text;
            }
            updatePreview();
        }

        // 角丸同期＋プレビュー更新 / Sync corner radius and update preview
        function syncAndPreviewR() {
            updatePreview();
        }

        // 入力値変更時にプレビューを更新（連動対応） / Update preview while editing inputs
        inputW.onChanging = syncAndPreviewW;
        inputH.onChanging = syncAndPreviewH;

        // ↑↓キーでの値の増減 / Increment or decrement with arrow keys
        changeValueByArrowKey(inputW, syncAndPreviewW);
        changeValueByArrowKey(inputH, syncAndPreviewH);

        // 角丸パネル / Rounded corner panel
        var panelR = win.add("panel", undefined, L("panelCorner"));
        panelR.orientation = "row";
        panelR.alignChildren = ["left", "center"];
        panelR.margins = [15, 20, 15, 10];

        var groupR = panelR.add("group");
        var labelR = groupR.add("statictext", undefined, L("labelRadius"));
        labelR.preferredSize = [labelWidth, -1];
        labelR.justify = "right";
        var inputR = groupR.add("edittext", undefined, "0");
        inputR.characters = 5;
        groupR.add("statictext", undefined, radiusUnit);

        inputR.onChanging = syncAndPreviewR;
        changeValueByArrowKey(inputR, syncAndPreviewR);

        var btnGroup = win.add("group");
        btnGroup.alignment = "center";
        var btnCancel = btnGroup.add("button", undefined, L("btnCancel"), { name: "cancel" });
        var btnOK = btnGroup.add("button", undefined, L("btnOK"), { name: "ok" });

        btnOK.onClick = function () {
            var addW = validateNumericField(inputW, true, "dialogTitle", "alertInvalidNumber");
            if (addW === null) return;

            var addH;
            if (linkCheck.value) {
                addH = addW;
            } else {
                addH = validateNumericField(inputH, true, "dialogTitle", "alertInvalidNumber");
                if (addH === null) return;
            }

            var radius = validateNumericField(inputR, false, "dialogTitle", "alertInvalidRadius");
            if (radius === null) return;

            inputW.text = addW;
            inputH.text = addH;
            inputR.text = radius;
            app.executeMenuCommand('edge');

            win.close(1);
        };

        win.onShow = function () {
            inputW.active = true;
            inputW.selection = [0, inputW.text.length];
        };

        buildPreview(20, 20, 0);

        // ダイアログ表示 / Show dialog
        if (win.show() === 1) {
            var addW = parseNumberOrDefault(inputW.text, 0);
            var addH = parseNumberOrDefault(inputH.text, 0);
            var radius = parseNumberOrDefault(inputR.text, 0);
            if (radius < 0) radius = 0;
            return { addW: addW, addH: addH, radius: radius, previewItem: currentPreviewItem };
        } else {
            removePreviewItem();
            return null;
        }
    }

    /* メイン処理 / Main process */
    function main() {
        if (app.documents.length === 0) {
            alert(L("alertNoDocument"));
            return;
        }

        var doc = app.activeDocument;
        var sel = doc.selection;

        app.executeMenuCommand('edge');


        // 選択オブジェクトが2つではない場合のエラー処理 / Error when selection count is not two
        if (sel.length !== 2) {
            alert(L("alertSelectTwo"), L("alertSelectError"));
            return;
        }

        var contentItem = null;
        var shapeItem = null;

        // コンテンツ候補と図形候補を厳密に判別 / Strictly classify content and shape candidates
        for (var i = 0; i < sel.length; i++) {
            var item = sel[i];

            if (isContentItem(item)) {
                if (contentItem) {
                    alert(L("alertOnlyOneContent"), L("alertSelectError"));
                    return;
                }
                contentItem = item;
                continue;
            }

            if (isShapeItem(item)) {
                if (shapeItem) {
                    alert(L("alertOnlyOneShape"), L("alertSelectError"));
                    return;
                }
                shapeItem = item;
                continue;
            }

            alert(L("alertUnsupportedSelection"), L("alertSelectError"));
            return;
        }

        if (!contentItem || !shapeItem) {
            alert(L("alertNeedContentAndShape"), L("alertSelectError"));
            return;
        }

        if (isClippingGroupItem(contentItem)) {
            alert(L("alertClippingGroupNotSupported"), L("alertSelectError"));
            return;
        }



        // 1. コンテンツを複製して計測用オブジェクトを作成 / Duplicate content and create measurement item
        var measureItem = null;
        var dupText = null;
        var boundsInfo = null;

        try {
            if (contentItem.typename === "TextFrame") {
                dupText = contentItem.duplicate();
                measureItem = dupText.createOutline();
                boundsInfo = getBoundsFromItem(measureItem);
            } else {
                measureItem = contentItem.duplicate();
                boundsInfo = getBoundsFromItem(measureItem);
            }
        } finally {
            if (measureItem) {
                try {
                    measureItem.remove();
                } catch (e) { }
            }
            if (dupText) {
                try {
                    dupText.remove();
                } catch (e) { }
            }
        }

        if (!boundsInfo) {
            shapeItem.hidden = false;
            alert(L("alertMeasureFailed"), L("alertSelectError"));
            return;
        }
        // 2. 幅と高さ、およびコンテンツの中心座標を計測 / Measure width, height, and content center
        var outWidth = boundsInfo.width;
        var outHeight = boundsInfo.height;
        var contentCenterX = boundsInfo.centerX;
        var contentCenterY = boundsInfo.centerY;

        // 4. 元図形は保持し、プレビュー専用のベース図形を1回だけ作成してアピアランスをクリア / Keep original shape untouched and prepare one cleared preview base shape
        var previewBaseShapeItem = shapeItem.duplicate();
        previewBaseShapeItem.hidden = false;
        clearAppearanceByAction(previewBaseShapeItem);
        previewBaseShapeItem.hidden = true;

        // 4. 元の図形は保持し、ダイアログ中はプレビュー専用オブジェクトを使う / Keep original shape and use preview-only object in dialog
        shapeItem.hidden = true;
        app.redraw();

        // 5. ダイアログ表示 / Show dialog
        var result = showDialog(outWidth, outHeight, previewBaseShapeItem, contentCenterX, contentCenterY);

        if (result) {
            try {
                if (previewBaseShapeItem && previewBaseShapeItem !== result.previewItem) {
                    previewBaseShapeItem.remove();
                }
            } catch (e) { }
            // OKの場合、元の図形を削除してプレビュー専用オブジェクトを確定版として残す / On OK, remove original and keep preview object as final
            try {
                shapeItem.remove();
            } catch (e) {
                shapeItem.hidden = false;
            }

            if (result.previewItem) {
                result.previewItem.hidden = false;
                try {
                    if (previewBaseShapeItem && previewBaseShapeItem !== result.previewItem) {
                        previewBaseShapeItem.hidden = true;
                    }
                    doc.selection = null;
                    contentItem.selected = true;
                    result.previewItem.selected = true;
                    app.executeMenuCommand('Horizontal Align Center');
                } catch (e) { }
            }

            app.redraw();
        } else {
            try {
                if (previewBaseShapeItem) {
                    previewBaseShapeItem.remove();
                }
            } catch (e) { }
            // キャンセル時はプレビューを破棄し、元の図形をそのまま戻す / On cancel, discard preview and restore original shape
            shapeItem.hidden = false;
            app.redraw();
        }
    }

    main();

})();
