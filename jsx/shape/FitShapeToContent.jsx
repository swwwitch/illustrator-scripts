#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);
#targetengine "FitShapeToContentSession"

    /*
    FitShapeToContent.jsx
    座布団メーカー
    
     * スクリプトの概要:
     * テキストやグループに合わせて、背面の座布団形状を手早く作成・調整するスクリプトです。
     * - 【1つ選択】テキストまたはグループのみを選ぶと、背面に座布団用の長方形を自動作成します
     * - 【2つ選択】テキストまたはグループと図形を選ぶと、既存の図形を座布団として使えます
     * - パディングの幅・高さをプレビューしながら調整できます
     * - 「連動」ON時は幅と高さを同じ値で扱えます
     * - 座布団の調整をOFFにすると、サイズ変更や角丸調整は行わず、中央合わせのみで使えます
     * - 角丸は有効/無効を切り替えできます
     * - 「ピル形状」ON時は高さに合わせたカプセル形状にできます
     * - セッション中は前回の入力値を保持します
     * - キャンセル時はプレビューを破棄し、1つ選択時に自動作成した長方形も削除します
     *
     * 紹介記事（note）
     * https://note.com/dtp_tranist/n/n6e4a6a2b175f
     *
     * 作成日: 2026-03-25
     * 更新日: 2026-03-27（概要コメントをユーザー視点に整理）
     */

    (function () {
        // =========================================
        // バージョンとローカライズ / Version and Localization
        // =========================================

        var SCRIPT_VERSION = "v2.0.0";
        var AUTO_CREATED_SHAPE_OPACITY = 20;

        var __FS2C_SESSION_KEY = "__FitShapeToContentSession__";

        if (!$.global[__FS2C_SESSION_KEY]) {
            $.global[__FS2C_SESSION_KEY] = {
                addW: "20",
                addH: "20",
                radius: "0",
                radiusEnabled: true,
                link: true,
                pill: false
            };
        }

        function getSessionState() {
            return $.global[__FS2C_SESSION_KEY];
        }

        function saveSessionState(state) {
            $.global[__FS2C_SESSION_KEY] = state;
        }


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
            dialogTitle: { ja: "座布団メーカー", en: "Fit Shape to Content" },
            labelAdjustEnabled: { ja: "座布団の調整", en: "Adjust Shape" },
            panelPadding: { ja: "パディング", en: "Padding" },
            panelCorner: { ja: "角丸", en: "Rounded Corners" },
            labelWidth: { ja: "幅:", en: "Width:" },
            labelHeight: { ja: "高さ:", en: "Height:" },
            labelRadius: { ja: "半径:", en: "Radius:" },
            labelLink: { ja: "連動", en: "Link" },
            labelPill: { ja: "ピル形状", en: "Pill Shape" },
            btnCancel: { ja: "キャンセル", en: "Cancel" },
            btnOK: { ja: "OK", en: "OK" },
            alertNoDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            alertSelectOne: { ja: "テキストまたはグループを1つ、もしくはテキスト/グループと図形を計2つ選択して実行してください。", en: "Select one text/group item, or one text/group and one shape (2 items total)." },
            alertSelectError: { ja: "選択エラー", en: "Selection Error" },
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
        function changeValueByArrowKey(editText, onChanged, minValue) {
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

                if (typeof minValue === "number" && value < minValue) {
                    value = minValue;
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


        /* クリッピンググループか判定 / Check whether group is a clipping group */
        function isClippingGroupItem(item) {
            return !!(item && item.typename === "GroupItem" && item.clipped);
        }

        /* アクションでアピアランスをクリア / Clear appearance by embedded action */
        function clearAppearanceByAction(targetItem) {
            if (!targetItem) return;

            var str = '/version 3' + '/name [ 10' + ' 417070656172616e6365' + ']' + '/isOpen 1' + '/actionCount 1' + '/action-1 {' + ' /name [ 5' + ' 636c656172' + ' ]' + ' /keyIndex 0' + ' /colorIndex 0' + ' /isOpen 1' + ' /eventCount 1' + ' /event-1 {' + ' /useRulersIn1stQuadrant 0' + ' /internalName (ai_plugin_appearance)' + ' /localizedName [ 18' + ' e382a2e38394e382a2e383a9e383b3e382b9' + ' ]' + ' /isOpen 1' + ' /isOn 1' + ' /hasDialog 0' + ' /parameterCount 1' + ' /parameter-1 {' + ' /key 1835363957' + ' /showInPalette 4294967295' + ' /type (enumerated)' + ' /name [ 27' + ' e382a2e38394e382a2e383a9e383b3e382b9e38292e6b688e58ebb' + ' ]' + ' /value 6' + ' }' + ' }' + '}';

            var doc = app.activeDocument;
            var tempActionPath = Folder.temp.fsName + '/Appearance_clear_' + (new Date().getTime()) + '_' + Math.floor(Math.random() * 100000) + '.aia';
            var f = new File(tempActionPath);
            var originalStyle = captureShapeStyle(targetItem);

            try {
                doc.selection = null;
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
                    doc.selection = null;
                } catch (e) { }
            }
        }

        /* ダイアログを表示 / Show dialog */
        function showDialog(outWidth, outHeight, previewSourceShapeItem, previewBaseShapeItem, contentCenterX, contentCenterY, shapeIsAutoCreated) {
            var rulerUnit = getUnitLabel(app.preferences.getIntegerPreference("rulerType"), "rulerType");
            var radiusUnit = "pt";
            var sessionState = getSessionState();

            var currentPreviewItem = null;
            var confirmedDialogValues = null;

            function reflectAdjustEnabledUI() {
                var isAdjustEnabled = chkAdjustEnabled.value;
                var isRadiusEnabled = chkRadiusEnabled.value;
                var isPill = chkPill.value;
                var isLinked = linkCheck.value;
                var widthEnabled;
                var heightEnabled;
                var radiusEnabled;
                var pillEnabled;
                var linkEnabled;

                if (isPill && isLinked) {
                    linkCheck.value = false;
                    isLinked = false;
                }

                widthEnabled = isAdjustEnabled && !isPill;
                heightEnabled = isAdjustEnabled && (isPill || !isLinked);
                radiusEnabled = isAdjustEnabled && isRadiusEnabled && !isPill;
                pillEnabled = isAdjustEnabled && isRadiusEnabled;
                linkEnabled = isAdjustEnabled && !isPill;

                inputW.enabled = widthEnabled;
                inputH.enabled = heightEnabled;
                inputR.enabled = radiusEnabled;
                chkPill.enabled = pillEnabled;
                linkCheck.enabled = linkEnabled;
                chkAdjustEnabled.enabled = true;
                chkRadiusEnabled.enabled = isAdjustEnabled;
            }
            var MIN_PREVIEW_SIZE = 0.1;

            function removePreviewItem() {
                if (currentPreviewItem) {
                    try {
                        currentPreviewItem.remove();
                    } catch (e) { }
                    currentPreviewItem = null;
                }
            }

            function getMaxRadiusFromAddW(addW) {
                var effectiveWidth = Math.max(MIN_PREVIEW_SIZE, outWidth + addW);
                return effectiveWidth / 2;
            }

            function buildPreview(addW, addH, radius) {
                if (isNaN(addW)) addW = 0;
                if (isNaN(addH)) addH = 0;
                if (isNaN(radius) || radius < 0) radius = 0;

                removePreviewItem();

                var useAdjustedPreview = chkAdjustEnabled.value;
                var previewTemplate = (useAdjustedPreview || !previewSourceShapeItem) ? previewBaseShapeItem : previewSourceShapeItem;
                if (previewSourceShapeItem) previewSourceShapeItem.hidden = true;
                previewBaseShapeItem.hidden = true;

                currentPreviewItem = previewTemplate.duplicate();
                currentPreviewItem.hidden = false;

                if (useAdjustedPreview || !previewSourceShapeItem) {
                    var newWidth = Math.max(MIN_PREVIEW_SIZE, outWidth + addW);
                    var newHeight = Math.max(MIN_PREVIEW_SIZE, outHeight + addH);
                    currentPreviewItem.width = newWidth;
                    currentPreviewItem.height = newHeight;
                    currentPreviewItem.left = contentCenterX - (newWidth / 2);
                    currentPreviewItem.top = contentCenterY + (newHeight / 2);

                    if (radius > 0) {
                        var xml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + radius + ' "/></LiveEffect>';
                        currentPreviewItem.applyEffect(xml);
                    }
                } else {
                    var bounds = getBoundsFromItem(currentPreviewItem);
                    currentPreviewItem.left = contentCenterX - (bounds.width / 2);
                    currentPreviewItem.top = contentCenterY + (bounds.height / 2);
                }

                app.redraw();
            }

            var win = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
            win.orientation = "column";
            win.alignChildren = ["fill", "top"];
            win.margins = [15, 20, 15, 15];

            var uiLabelWidth = 50;

            var adjustRow = win.add("group");
            adjustRow.orientation = "row";
            adjustRow.alignChildren = ["center", "center"];
            adjustRow.alignment = ["center", "top"];

            var chkAdjustEnabled = adjustRow.add("checkbox", undefined, L("labelAdjustEnabled"));
            chkAdjustEnabled.value = !!shapeIsAutoCreated;

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
            labelW.preferredSize = [uiLabelWidth, -1];
            labelW.justify = "right";
            var inputW = groupW.add("edittext", undefined, sessionState.addW);
            inputW.characters = 5;
            groupW.add("statictext", undefined, rulerUnit);

            var groupH = colLeft.add("group");
            var labelH = groupH.add("statictext", undefined, L("labelHeight"));
            labelH.preferredSize = [uiLabelWidth, -1];
            labelH.justify = "right";
            var inputH = groupH.add("edittext", undefined, sessionState.addH);
            inputH.characters = 5;
            groupH.add("statictext", undefined, rulerUnit);

            // 右カラム：連動チェックボックス / Right column: link checkbox
            var colRight = panel.add("group");
            colRight.orientation = "column";
            colRight.alignChildren = ["left", "center"];

            var linkCheck = colRight.add("checkbox", undefined, L("labelLink"));
            linkCheck.value = sessionState.link;

            // 初期状態：連動ON時は高さをディム表示 / Initial state: height field disabled when linked
            inputH.enabled = !linkCheck.value;
            if (linkCheck.value) {
                inputH.text = inputW.text;
            }

            // 現在のUI入力値を読む / Read current UI input values
            function readUIValues() {
                return {
                    adjustEnabled: chkAdjustEnabled.value,
                    addW: parseNumberOrDefault(inputW.text, 0),
                    addH: parseNumberOrDefault(inputH.text, 0),
                    radius: parseNumberOrDefault(inputR.text, 0),
                    radiusEnabled: chkRadiusEnabled.value,
                    pill: chkPill.value,
                    link: linkCheck.value,
                    widthText: inputW.text,
                    heightText: inputH.text,
                    radiusText: inputR.text
                };
            }

            // UI入力値からプレビュー用の派生値を計算 / Compute preview values from UI inputs
            function computePreviewValues(uiValues) {
                var addW = uiValues.addW;
                var addH = uiValues.addH;
                var radius = uiValues.radius;
                var previewHeight;
                var newWidth;
                var maxRadius;

                if (isNaN(addW)) addW = 0;
                if (isNaN(addH)) addH = 0;
                if (isNaN(radius) || radius < 0) radius = 0;

                if (!uiValues.adjustEnabled) {
                    return {
                        addW: 0,
                        addH: 0,
                        radius: 0,
                        widthText: uiValues.widthText,
                        heightText: uiValues.heightText,
                        radiusText: uiValues.radiusText
                    };
                }

                if (!uiValues.radiusEnabled) {
                    radius = 0;
                } else if (uiValues.pill) {
                    previewHeight = Math.max(MIN_PREVIEW_SIZE, outHeight + addH);
                    radius = previewHeight / 2;
                    newWidth = Math.max(MIN_PREVIEW_SIZE, radius * 2);
                    addW = newWidth;
                } else {
                    maxRadius = getMaxRadiusFromAddW(addW);
                    if (radius > maxRadius) radius = maxRadius;
                }

                return {
                    addW: addW,
                    addH: addH,
                    radius: radius,
                    widthText: String(addW),
                    heightText: String(addH),
                    radiusText: String(radius)
                };
            }

            // 派生UI値を反映 / Apply derived UI values
            function applyDerivedUIValues(previewValues) {
                if (!chkAdjustEnabled.value) {
                    return;
                }

                if (!chkRadiusEnabled.value) {
                    inputR.text = "0";
                    return;
                }

                inputR.text = previewValues.radiusText;

                if (chkPill.value) {
                    inputW.text = previewValues.widthText;
                    inputH.text = previewValues.heightText;
                }
            }

            // プレビュー更新関数（描画のみ担当） / Update preview (drawing only)
            function updatePreview() {
                var uiValues = readUIValues();
                var previewValues = computePreviewValues(uiValues);
                buildPreview(previewValues.addW, previewValues.addH, previewValues.radius);
            }

            function refreshPreviewFromUI() {
                var uiValues = readUIValues();
                var previewValues = computePreviewValues(uiValues);
                applyDerivedUIValues(previewValues);
                buildPreview(previewValues.addW, previewValues.addH, previewValues.radius);
            }

            function getFinalDialogValues() {
                var addW = 0;
                var addH = 0;
                var radius = 0;
                var uiValues;
                var previewValues;

                if (!chkAdjustEnabled.value) {
                    return {
                        addW: 0,
                        addH: 0,
                        radius: 0,
                        radiusEnabled: false,
                        shouldRunPathfinder: false,
                        adjustEnabled: false
                    };
                }

                addW = validateNumericField(inputW, false, "dialogTitle", "alertInvalidNumber");
                if (addW === null) return null;

                if (linkCheck.value) {
                    addH = addW;
                    inputH.text = String(addH);
                } else {
                    addH = validateNumericField(inputH, false, "dialogTitle", "alertInvalidNumber");
                    if (addH === null) return null;
                }

                if (!chkRadiusEnabled.value) {
                    inputR.text = "0";
                } else {
                    if (chkPill.value) {
                        inputH.text = String(addH);
                    } else {
                        radius = validateNumericField(inputR, false, "dialogTitle", "alertInvalidRadius");
                        if (radius === null) return null;
                        inputR.text = String(radius);
                    }
                }

                uiValues = readUIValues();
                previewValues = computePreviewValues(uiValues);
                applyDerivedUIValues(previewValues);

                return {
                    addW: previewValues.addW,
                    addH: previewValues.addH,
                    radius: previewValues.radius,
                    radiusEnabled: chkRadiusEnabled.value,
                    shouldRunPathfinder: chkRadiusEnabled.value && chkPill.value,
                    adjustEnabled: true
                };
            }

            chkAdjustEnabled.onClick = function () {
                reflectAdjustEnabledUI();
                refreshPreviewFromUI();
            };

            // 連動チェック変更時 / When link checkbox changes
            linkCheck.onClick = function () {
                if (linkCheck.value) {
                    inputH.text = inputW.text;
                }
                reflectAdjustEnabledUI();
                refreshPreviewFromUI();
            };

            // 連動同期＋プレビュー更新 / Sync linked values and update preview
            function syncAndPreviewW() {
                if (linkCheck.value) {
                    inputH.text = inputW.text;
                }
                refreshPreviewFromUI();
            }
            function syncAndPreviewH() {
                if (linkCheck.value) {
                    inputW.text = inputH.text;
                }
                refreshPreviewFromUI();
            }

            // 角丸入力変更時のプレビュー更新 / Update preview when the corner radius input changes
            function syncAndPreviewR() {
                refreshPreviewFromUI();
            }

            // 入力値変更時にプレビューを更新（連動対応） / Update preview while editing inputs
            inputW.onChanging = function () {
                var widthValue = parseFloat(inputW.text);
                if (!isNaN(widthValue) && widthValue < 0) {
                    inputW.text = "0";
                }
                syncAndPreviewW();
            };
            inputH.onChanging = function () {
                var heightValue = parseFloat(inputH.text);
                if (!isNaN(heightValue) && heightValue < 0) {
                    inputH.text = "0";
                }
                syncAndPreviewH();
            };

            // ↑↓キーでの値の増減 / Increment or decrement with arrow keys
            changeValueByArrowKey(inputW, syncAndPreviewW, 0);
            changeValueByArrowKey(inputH, syncAndPreviewH, 0);

            // 角丸パネル（有効切り替え・半径・ピル形状） / Rounded corner panel (enable toggle, radius, and pill shape)
            var panelR = win.add("panel", undefined, L("panelCorner"));
            panelR.orientation = "row";
            panelR.alignChildren = ["left", "center"];
            panelR.alignment = ["fill", "top"];
            panelR.margins = [15, 20, 15, 10];

            var colR = panelR.add("group");
            colR.orientation = "column";
            colR.alignChildren = ["left", "top"];

            var groupR = colR.add("group");
            var chkRadiusEnabled = groupR.add("checkbox", undefined, L("labelRadius"));
            chkRadiusEnabled.value = (sessionState.radiusEnabled !== false);
            var inputR = groupR.add("edittext", undefined, sessionState.radius);
            inputR.characters = 5;
            groupR.add("statictext", undefined, radiusUnit);

            var groupPill = colR.add("group");
            groupPill.alignment = ["left", "top"];
            var chkPill = groupPill.add("checkbox", undefined, L("labelPill"));
            chkPill.value = !!sessionState.pill;
            chkPill.onClick = function () {
                if (!chkPill.value && linkCheck.value) {
                    inputH.text = inputW.text;
                }
                reflectAdjustEnabledUI();
                refreshPreviewFromUI();
            };

            chkRadiusEnabled.onClick = function () {
                if (!chkRadiusEnabled.value) {
                    inputR.text = "0";
                    chkPill.value = false;
                }
                if (linkCheck.value) {
                    inputH.text = inputW.text;
                }
                reflectAdjustEnabledUI();
                refreshPreviewFromUI();
            };

            if (chkPill.value) {
                linkCheck.value = false;
            }
            reflectAdjustEnabledUI();
            refreshPreviewFromUI();

            inputR.onChanging = function () {
                var radiusValue = parseFloat(inputR.text);
                if (!isNaN(radiusValue) && radiusValue < 0) {
                    inputR.text = "0";
                }
                syncAndPreviewR();
            };
            changeValueByArrowKey(inputR, syncAndPreviewR, 0);


            var btnGroup = win.add("group");
            btnGroup.alignment = "center";
            var btnCancel = btnGroup.add("button", undefined, L("btnCancel"), { name: "cancel" });
            var btnOK = btnGroup.add("button", undefined, L("btnOK"), { name: "ok" });

            btnOK.onClick = function () {
                confirmedDialogValues = getFinalDialogValues();
                if (!confirmedDialogValues) return;

                saveSessionState({
                    addW: inputW.text,
                    addH: inputH.text,
                    radius: inputR.text,
                    radiusEnabled: chkRadiusEnabled.value,
                    link: linkCheck.value,
                    pill: chkPill.value
                });

                win.close(1);
            };

            btnCancel.onClick = function () {
                saveSessionState({
                    addW: inputW.text,
                    addH: inputH.text,
                    radius: inputR.text,
                    radiusEnabled: chkRadiusEnabled.value,
                    link: linkCheck.value,
                    pill: chkPill.value
                });
                win.close(0);
            };

            win.onShow = function () {
                inputW.active = true;
                inputW.selection = [0, inputW.text.length];
                if (chkPill.value) {
                    linkCheck.value = false;
                }
                reflectAdjustEnabledUI();
                refreshPreviewFromUI();
            };

            updatePreview();

            // ダイアログ表示 / Show dialog
            if (win.show() === 1) {
                var finalValues = confirmedDialogValues || getFinalDialogValues();
                if (!finalValues) {
                    removePreviewItem();
                    return null;
                }
                return {
                    addW: finalValues.addW,
                    addH: finalValues.addH,
                    radius: finalValues.radius,
                    radiusEnabled: finalValues.radiusEnabled,
                    shouldRunPathfinder: finalValues.shouldRunPathfinder,
                    adjustEnabled: finalValues.adjustEnabled,
                    previewItem: currentPreviewItem
                };
            } else {
                removePreviewItem();
                return null;
            }
        }

        /* 条件限定でグループを分解 / Extract content only from a very limited text/group + shape group */
        function tryExtractContentFromGroup(groupItem) {
            // 対応条件はかなり限定的です:
            // - GroupItem であること
            // - clipping group ではないこと
            // - 子要素がちょうど2つであること
            // - 構成が content 1つ + shape 1つ であること
            // それ以外は分解せず null を返します。
            if (!groupItem || groupItem.typename !== "GroupItem") return null;
            if (groupItem.clipped) return null;

            var contentChild = null;
            var shapeChild = null;
            var totalItems = groupItem.pageItems.length;

            if (totalItems !== 2) return null;

            for (var i = 0; i < totalItems; i++) {
                var child = groupItem.pageItems[i];
                if (isContentItem(child)) {
                    if (contentChild) return null;
                    contentChild = child;
                } else if (isShapeItem(child)) {
                    if (shapeChild) return null;
                    shapeChild = child;
                } else {
                    return null;
                }
            }

            if (!contentChild || !shapeChild) return null;

            // グループを解除して個別アイテムにする / Ungroup to get individual items
            var insertBefore = groupItem;

            // グループ内のアイテムをグループの前に移動 / Move items out of the group
            contentChild.move(insertBefore, ElementPlacement.PLACEBEFORE);
            shapeChild.move(contentChild, ElementPlacement.PLACEAFTER);

            // 空になったグループを削除 / Remove the now-empty group
            try { groupItem.remove(); } catch (e) { }

            // 旧図形を削除 / Remove the old shape
            try { shapeChild.remove(); } catch (e) { }

            return contentChild;
        }

        /* メイン処理 / Main process */
        function main() {
            if (app.documents.length === 0) {
                alert(L("alertNoDocument"));
                return;
            }

            var doc = app.activeDocument;
            var sel = doc.selection;

            if (sel.length < 1 || sel.length > 2) {
                alert(L("alertSelectOne"), L("alertSelectError"));
                return;
            }

            var contentItem = null;
            var shapeItem = null;
            var shapeIsAutoCreated = false;

            if (sel.length === 2) {
                // 2つ選択：テキスト/グループ＋図形（従来動作） / Two items: text/group + shape (original behavior)
                for (var i = 0; i < sel.length; i++) {
                    var item = sel[i];
                    if (isContentItem(item)) {
                        if (contentItem) {
                            alert(L("alertSelectOne"), L("alertSelectError"));
                            return;
                        }
                        contentItem = item;
                    } else if (isShapeItem(item)) {
                        if (shapeItem) {
                            alert(L("alertSelectOne"), L("alertSelectError"));
                            return;
                        }
                        shapeItem = item;
                    } else {
                        alert(L("alertSelectOne"), L("alertSelectError"));
                        return;
                    }
                }
                if (!contentItem || !shapeItem) {
                    alert(L("alertSelectOne"), L("alertSelectError"));
                    return;
                }
            } else {
                // 1つ選択：テキスト/グループのみ / One item: text/group only
                var selectedItem = sel[0];

                // グループ内がテキスト＋図形の場合は分解して続行 / If group contains text+shape, extract and continue
                if (selectedItem.typename === "GroupItem" && !isClippingGroupItem(selectedItem)) {
                    contentItem = tryExtractContentFromGroup(selectedItem);
                }

                if (!contentItem) {
                    if (!isContentItem(selectedItem)) {
                        alert(L("alertSelectOne"), L("alertSelectError"));
                        return;
                    }
                    contentItem = selectedItem;
                }

                shapeIsAutoCreated = true;
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
                alert(L("alertMeasureFailed"), L("alertSelectError"));
                return;
            }

            // 2. 幅と高さ、およびコンテンツの中心座標を計測 / Measure width, height, and content center
            var outWidth = boundsInfo.width;
            var outHeight = boundsInfo.height;
            var contentCenterX = boundsInfo.centerX;
            var contentCenterY = boundsInfo.centerY;

            // 3. 図形の準備 / Prepare shape item
            if (shapeIsAutoCreated) {
                // 1つ選択時：コンテンツ背面に長方形を自動作成（初期不透明度は定数で管理） / Auto-create rectangle behind content (initial opacity is managed by a constant)
                shapeItem = doc.pathItems.rectangle(
                    boundsInfo.top,
                    boundsInfo.left,
                    boundsInfo.width,
                    boundsInfo.height
                );
                shapeItem.opacity = AUTO_CREATED_SHAPE_OPACITY;
                shapeItem.move(contentItem, ElementPlacement.PLACEAFTER);
            }

            // 4. プレビュー用の複製を作成 / Create preview duplicates
            var previewSourceShapeItem = null;
            var previewBaseShapeItem = null;

            if (!shapeIsAutoCreated) {
                // 2つ選択時：整列専用と調整専用の2つを作成 / Two-item mode: create align-only and adjust-only copies
                previewSourceShapeItem = shapeItem.duplicate();
                previewSourceShapeItem.hidden = true;

                previewBaseShapeItem = shapeItem.duplicate();
                previewBaseShapeItem.hidden = false;
                clearAppearanceByAction(previewBaseShapeItem);
                previewBaseShapeItem.hidden = true;
            } else {
                // 1つ選択時：ベース図形のみ / One-item mode: base shape only
                previewBaseShapeItem = shapeItem.duplicate();
                previewBaseShapeItem.hidden = true;
            }

            // 5. 元の図形は非表示にし、ダイアログ中はプレビュー専用オブジェクトを使う / Hide original shape and use preview-only object in dialog
            shapeItem.hidden = true;
            app.redraw();

            // 6. ダイアログ表示 / Show dialog
            var result = showDialog(outWidth, outHeight, previewSourceShapeItem, previewBaseShapeItem, contentCenterX, contentCenterY, shapeIsAutoCreated);

            if (result) {
                try {
                    if (previewSourceShapeItem && previewSourceShapeItem !== result.previewItem) {
                        previewSourceShapeItem.remove();
                        previewSourceShapeItem = null;
                    }
                } catch (e) {
                    previewSourceShapeItem = null;
                }
                try {
                    if (previewBaseShapeItem && previewBaseShapeItem !== result.previewItem) {
                        previewBaseShapeItem.remove();
                        previewBaseShapeItem = null;
                    }
                } catch (e) {
                    previewBaseShapeItem = null;
                }
                // OKの場合、元の図形を削除してプレビュー専用オブジェクトを確定版として残す / On OK, remove the original shape and keep the preview object as the final shape

                if (result.previewItem) {
                    var finalPreviewItem = result.previewItem;
                    finalPreviewItem.hidden = false;

                    var baseShapeRef = result.adjustEnabled ? previewBaseShapeItem : previewSourceShapeItem;
                    if (baseShapeRef && baseShapeRef !== finalPreviewItem) {
                        try {
                            baseShapeRef.hidden = true;
                        } catch (e) { }
                    }

                    if (result.shouldRunPathfinder) {
                        doc.selection = null;
                        try {
                            finalPreviewItem.selected = true;
                        } catch (e) {
                            throw new Error('Preview item became invalid before Live Pathfinder Add.');
                        }

                        app.executeMenuCommand('Live Pathfinder Add');

                        if (!doc.selection || doc.selection.length !== 1) {
                            throw new Error('Live Pathfinder Add did not return exactly one selected item.');
                        }

                        finalPreviewItem = doc.selection[0];
                        try {
                            finalPreviewItem.hidden = false;
                        } catch (e) {
                            throw new Error('Live Pathfinder Add returned an invalid item.');
                        }
                    }

                    // OKの場合、元の図形を削除 / On OK, remove the original shape
                    try {
                        shapeItem.remove();
                    } catch (removeError) {
                        if (!shapeIsAutoCreated) {
                            shapeItem.hidden = false;
                        }
                        throw removeError;
                    }

                    result.previewItem = finalPreviewItem;

                    doc.selection = null;
                    try {
                        contentItem.selected = true;
                    } catch (e) { }
                    try {
                        finalPreviewItem.selected = true;
                    } catch (e) { }
                }

                app.redraw();
            } else {
                try {
                    if (previewSourceShapeItem) {
                        previewSourceShapeItem.remove();
                    }
                } catch (e) { }
                try {
                    if (previewBaseShapeItem) {
                        previewBaseShapeItem.remove();
                    }
                } catch (e) { }
                if (shapeIsAutoCreated) {
                    // キャンセル時：自動作成した長方形を削除 / On cancel: remove auto-created rectangle
                    try {
                        shapeItem.remove();
                    } catch (e) { }
                } else {
                    // キャンセル時：元の図形を復元 / On cancel: restore original shape
                    shapeItem.hidden = false;
                }
                app.redraw();
            }
        }

        main();

    })();