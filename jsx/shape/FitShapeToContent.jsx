#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);
#targetengine "FitShapeToContentSession"

    /*
    FitShapeToContent.jsx
    座布団メーカー

    概要：
    テキストやグループに合わせて、背面の座布団形状を手早く作成・調整するスクリプトです。
    - 【1つ選択】テキストまたはグループのみを選ぶと、背面に座布団用の長方形を自動作成します
    - 【2つ選択】テキストまたはグループと図形を選ぶと、既存の図形を座布団として使えます
    - パディングの幅・高さをプレビューしながら調整できます
    - 「連動」ON時は幅と高さを同じ値で扱えます
    - 座布団の調整をOFFにすると、サイズ変更や角丸調整は行わず、中央合わせのみで使えます
    - 角丸は有効/無効を切り替えできます
    - 「ピル形状」ON時は高さに合わせたカプセル形状にできます
    - セッション中は前回の入力値を保持します
    - キャンセル時はプレビューを破棄し、1つ選択時に自動作成した長方形も削除します
    - OK確定後の Undo は 1 ステップで全変更を巻き戻せます

    紹介記事（note）
    https://note.com/dtp_tranist/n/n6e4a6a2b175f

    作成日: 2026-03-25
    更新日: 2026-05-25（プレビューを runPreview/undoPreview パターンに刷新し、main を 5 関数 + プレビューコントローラに分割、safeRemove ヘルパー導入）
    */

    (function () {
        // =========================================
        // バージョンとローカライズ / Version and Localization
        // =========================================

        var SCRIPT_VERSION = "v2.0.2";
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
            /* === ダイアログ / Dialog === */
            dialogTitle: { ja: "座布団メーカー", en: "Fit Shape to Content" },
            panelPadding: { ja: "パディング", en: "Padding" },
            panelCorner: { ja: "角丸", en: "Rounded Corners" },

            /* === ラベル / Labels === */
            labelAdjustEnabled: { ja: "座布団の調整", en: "Adjust Shape" },
            labelWidth: { ja: "幅:", en: "Width:" },
            labelHeight: { ja: "高さ:", en: "Height:" },
            labelRadius: { ja: "半径:", en: "Radius:" },
            labelLink: { ja: "連動", en: "Link" },
            labelPill: { ja: "ピル形状", en: "Pill Shape" },

            /* === ボタン / Buttons === */
            btnCancel: { ja: "キャンセル", en: "Cancel" },
            btnOK: { ja: "OK", en: "OK" },

            /* === アラート / Alerts === */
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

        /* 例外を握り潰す共通ラッパ / Swallow-exception wrappers */
        function safeDo(fn) { try { fn(); } catch (e) { } }
        function safeRemove(item) { safeDo(function () { if (item) item.remove(); }); }

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
                try { app.unloadAction('Appearance', ''); } catch (e) { }
                if (f && f.exists) f.remove();
                doc.selection = null;
            }
        }

        /* =========================================
         * プレビュー undo ヘルパー / Preview undo helpers
         *
         * 「設定変更のたびに 前回 undo → 再生成」を回す共通パターン。
         * process() は ScriptUI 値から毎回再構築する純粋関数として書く。
         * - runPreview: UI 変更ごと
         * - undoPreview: OK 直前。前回プレビューを巻き戻して process() を本番として再実行する
         * - cleanupPreview: Cancel / ダイアログクローズ時
         * ========================================= */

        function runPreview(state, processFn, isEnabled) {
            try {
                if (isEnabled) {
                    if (state.isUndo) app.undo();
                    else state.isUndo = true;
                    processFn();
                    app.redraw();
                } else if (state.isUndo) {
                    app.undo();
                    app.redraw();
                    state.isUndo = false;
                }
            } catch (err) { }
        }

        function undoPreview(state) {
            try {
                if (state.isUndo) app.undo();
            } catch (err) { }
            state.isUndo = false;
        }

        function cleanupPreview(state) {
            try {
                if (state.isUndo) app.undo();
                state.isUndo = false;
            } catch (err) { }
        }

        /* =========================================
         * プレビューコントローラ / Preview controller
         *
         * widget / geometry / template / previewState を束ねて preview ロジックを一括管理。
         * 公開 API: reflectEnabled, refresh, getFinalValues, commitFinal。
         * widgets   = { chkAdjustEnabled, chkRadiusEnabled, chkPill, linkCheck, inputW, inputH, inputR }
         * geometry  = { outWidth, outHeight, contentCenterX, contentCenterY }
         * templates = { previewSourceShapeItem, previewBaseShapeItem }
         * previewState = { isUndo: boolean }
         * ========================================= */
        function createPreviewController(widgets, geometry, templates, previewState) {
            var MIN_PREVIEW_SIZE = 0.1;

            function getMaxRadiusFromAddW(addW) {
                var effectiveWidth = Math.max(MIN_PREVIEW_SIZE, geometry.outWidth + addW);
                return effectiveWidth / 2;
            }

            // UI の enabled 状態を現在のチェック状態から計算して反映 / Reflect enabled state
            function reflectEnabled() {
                var isAdjustEnabled = widgets.chkAdjustEnabled.value;
                var isRadiusEnabled = widgets.chkRadiusEnabled.value;
                var isPill = widgets.chkPill.value;
                var isLinked = widgets.linkCheck.value;

                if (isPill && isLinked) {
                    widgets.linkCheck.value = false;
                    isLinked = false;
                }

                widgets.inputW.enabled = isAdjustEnabled && !isPill;
                widgets.inputH.enabled = isAdjustEnabled && (isPill || !isLinked);
                widgets.inputR.enabled = isAdjustEnabled && isRadiusEnabled && !isPill;
                widgets.chkPill.enabled = isAdjustEnabled && isRadiusEnabled;
                widgets.linkCheck.enabled = isAdjustEnabled && !isPill;
                widgets.chkAdjustEnabled.enabled = true;
                widgets.chkRadiusEnabled.enabled = isAdjustEnabled;
            }

            function readUIValues() {
                return {
                    adjustEnabled: widgets.chkAdjustEnabled.value,
                    addW: parseNumberOrDefault(widgets.inputW.text, 0),
                    addH: parseNumberOrDefault(widgets.inputH.text, 0),
                    radius: parseNumberOrDefault(widgets.inputR.text, 0),
                    radiusEnabled: widgets.chkRadiusEnabled.value,
                    pill: widgets.chkPill.value,
                    link: widgets.linkCheck.value,
                    widthText: widgets.inputW.text,
                    heightText: widgets.inputH.text,
                    radiusText: widgets.inputR.text
                };
            }

            function computePreviewValues(uiValues) {
                var addW = uiValues.addW;
                var addH = uiValues.addH;
                var radius = uiValues.radius;
                if (isNaN(addW)) addW = 0;
                if (isNaN(addH)) addH = 0;
                if (isNaN(radius) || radius < 0) radius = 0;

                if (!uiValues.adjustEnabled) {
                    return {
                        addW: 0, addH: 0, radius: 0,
                        widthText: uiValues.widthText,
                        heightText: uiValues.heightText,
                        radiusText: uiValues.radiusText
                    };
                }

                if (!uiValues.radiusEnabled) {
                    radius = 0;
                } else if (uiValues.pill) {
                    var previewHeight = Math.max(MIN_PREVIEW_SIZE, geometry.outHeight + addH);
                    radius = previewHeight / 2;
                    addW = Math.max(MIN_PREVIEW_SIZE, radius * 2);
                } else {
                    var maxRadius = getMaxRadiusFromAddW(addW);
                    if (radius > maxRadius) radius = maxRadius;
                }

                return {
                    addW: addW, addH: addH, radius: radius,
                    widthText: String(addW),
                    heightText: String(addH),
                    radiusText: String(radius)
                };
            }

            function applyDerivedUIValues(previewValues) {
                if (!widgets.chkAdjustEnabled.value) return;
                if (!widgets.chkRadiusEnabled.value) {
                    widgets.inputR.text = "0";
                    return;
                }
                widgets.inputR.text = previewValues.radiusText;
                if (widgets.chkPill.value) {
                    widgets.inputW.text = previewValues.widthText;
                    widgets.inputH.text = previewValues.heightText;
                }
            }

            /*
             * 現在の UI 値からプレビュー shape を毎回ゼロから再構築する純粋関数。
             * runPreview / undoPreview / cleanupPreview ヘルパーが undo タイミングを管理し、
             * ユーザー履歴には「最後の process() 1 回分」だけが残る。
             * 確定後の特定用に末尾で previewItem を選択する。
             */
            function process() {
                var uiValues = readUIValues();
                var previewValues = computePreviewValues(uiValues);
                applyDerivedUIValues(previewValues);

                var addW = previewValues.addW;
                var addH = previewValues.addH;
                var radius = previewValues.radius;

                var doc = app.activeDocument;
                var useAdjustedPreview = widgets.chkAdjustEnabled.value;
                var previewTemplate = (useAdjustedPreview || !templates.previewSourceShapeItem)
                    ? templates.previewBaseShapeItem : templates.previewSourceShapeItem;

                var previewItem = previewTemplate.duplicate();
                previewItem.hidden = false;

                if (useAdjustedPreview || !templates.previewSourceShapeItem) {
                    var newWidth = Math.max(MIN_PREVIEW_SIZE, geometry.outWidth + addW);
                    var newHeight = Math.max(MIN_PREVIEW_SIZE, geometry.outHeight + addH);
                    previewItem.width = newWidth;
                    previewItem.height = newHeight;
                    previewItem.left = geometry.contentCenterX - (newWidth / 2);
                    previewItem.top = geometry.contentCenterY + (newHeight / 2);

                    if (radius > 0) {
                        var xml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + radius + ' "/></LiveEffect>';
                        previewItem.applyEffect(xml);
                    }
                } else {
                    var bounds = getBoundsFromItem(previewItem);
                    previewItem.left = geometry.contentCenterX - (bounds.width / 2);
                    previewItem.top = geometry.contentCenterY + (bounds.height / 2);
                }

                // 後で finalPreviewItem として特定するため選択 / Select for post-OK identification
                doc.selection = null;
                previewItem.selected = true;
            }

            function refresh() {
                runPreview(previewState, process, true);
            }

            // バリデーション専用：validateNumericField 呼び出しと link/pill 連動の text 書き換え。
            // 成功 → true、失敗 → null（呼び出し側は OK を中断する。alert は validateNumericField が出す）
            // Validation only. Returns true on success, null on failure (alert already shown).
            function validateInputs() {
                if (!widgets.chkAdjustEnabled.value) return true;

                var addW = validateNumericField(widgets.inputW, false, "dialogTitle", "alertInvalidNumber");
                if (addW === null) return null;

                var addH;
                if (widgets.linkCheck.value) {
                    addH = addW;
                    widgets.inputH.text = String(addH);
                } else {
                    addH = validateNumericField(widgets.inputH, false, "dialogTitle", "alertInvalidNumber");
                    if (addH === null) return null;
                }

                if (!widgets.chkRadiusEnabled.value) {
                    widgets.inputR.text = "0";
                } else if (widgets.chkPill.value) {
                    widgets.inputH.text = String(addH);
                } else {
                    var radius = validateNumericField(widgets.inputR, false, "dialogTitle", "alertInvalidRadius");
                    if (radius === null) return null;
                    widgets.inputR.text = String(radius);
                }
                return true;
            }

            // バリデーション済み前提で UI 値から確定オブジェクトを派生
            // Assumes inputs are already validated. Returns the final values object.
            function deriveFinalValues() {
                if (!widgets.chkAdjustEnabled.value) {
                    return {
                        addW: 0, addH: 0, radius: 0,
                        radiusEnabled: false,
                        shouldRunPathfinder: false,
                        adjustEnabled: false
                    };
                }

                var uiValues = readUIValues();
                var previewValues = computePreviewValues(uiValues);
                applyDerivedUIValues(previewValues);

                return {
                    addW: previewValues.addW,
                    addH: previewValues.addH,
                    radius: previewValues.radius,
                    radiusEnabled: widgets.chkRadiusEnabled.value,
                    shouldRunPathfinder: widgets.chkRadiusEnabled.value && widgets.chkPill.value,
                    adjustEnabled: true
                };
            }

            function getFinalValues() {
                if (validateInputs() === null) return null;
                return deriveFinalValues();
            }

            // OK 確定：前回プレビューを undo して process を本番として 1 回だけ実行 /
            // Undo last preview then run process once as the final commit
            function commitFinal() {
                undoPreview(previewState);
                process();
            }

            return {
                reflectEnabled: reflectEnabled,
                refresh: refresh,
                getFinalValues: getFinalValues,
                commitFinal: commitFinal
            };
        }

        /* ダイアログを表示 / Show dialog */
        function showDialog(outWidth, outHeight, previewSourceShapeItem, previewBaseShapeItem, contentCenterX, contentCenterY, shapeIsAutoCreated) {
            var rulerUnit = getUnitLabel(app.preferences.getIntegerPreference("rulerType"), "rulerType");
            var radiusUnit = "pt";
            var sessionState = getSessionState();

            var finalPreviewItem = null;
            var confirmedDialogValues = null;
            var previewState = { isUndo: false };

            // コントローラはダイアログ UI 全 widget 生成後に作成する（後段の var で参照可能）
            // Controller is instantiated after all widgets exist; ctrl is hoisted via var
            var ctrl;

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

            chkAdjustEnabled.onClick = function () {
                ctrl.reflectEnabled();
                ctrl.refresh();
            };

            // 連動チェック変更時 / When link checkbox changes
            linkCheck.onClick = function () {
                if (linkCheck.value) {
                    inputH.text = inputW.text;
                }
                ctrl.reflectEnabled();
                ctrl.refresh();
            };

            // 負値を 0 にクランプ / Clamp negative input text to 0
            function clampNonNegativeText(editText) {
                var v = parseFloat(editText.text);
                if (!isNaN(v) && v < 0) editText.text = "0";
            }

            // 連動同期＋プレビュー更新 / Sync linked values and update preview
            function syncAndPreviewW() {
                if (linkCheck.value) inputH.text = inputW.text;
                ctrl.refresh();
            }
            function syncAndPreviewH() {
                if (linkCheck.value) inputW.text = inputH.text;
                ctrl.refresh();
            }

            // 入力値変更時にプレビューを更新（連動対応） / Update preview while editing inputs
            inputW.onChanging = function () {
                clampNonNegativeText(inputW);
                syncAndPreviewW();
            };
            inputH.onChanging = function () {
                clampNonNegativeText(inputH);
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

            // 全 widget が揃ったのでコントローラを生成 / All widgets exist; instantiate controller
            ctrl = createPreviewController(
                {
                    chkAdjustEnabled: chkAdjustEnabled,
                    chkRadiusEnabled: chkRadiusEnabled,
                    chkPill: chkPill,
                    linkCheck: linkCheck,
                    inputW: inputW,
                    inputH: inputH,
                    inputR: inputR
                },
                {
                    outWidth: outWidth,
                    outHeight: outHeight,
                    contentCenterX: contentCenterX,
                    contentCenterY: contentCenterY
                },
                {
                    previewSourceShapeItem: previewSourceShapeItem,
                    previewBaseShapeItem: previewBaseShapeItem
                },
                previewState
            );

            chkPill.onClick = function () {
                if (!chkPill.value && linkCheck.value) {
                    inputH.text = inputW.text;
                }
                ctrl.reflectEnabled();
                ctrl.refresh();
            };

            chkRadiusEnabled.onClick = function () {
                if (!chkRadiusEnabled.value) {
                    inputR.text = "0";
                    chkPill.value = false;
                }
                if (linkCheck.value) {
                    inputH.text = inputW.text;
                }
                ctrl.reflectEnabled();
                ctrl.refresh();
            };

            if (chkPill.value) {
                linkCheck.value = false;
            }
            ctrl.reflectEnabled();
            ctrl.refresh();

            inputR.onChanging = function () {
                clampNonNegativeText(inputR);
                ctrl.refresh();
            };
            changeValueByArrowKey(inputR, ctrl.refresh, 0);


            var btnGroup = win.add("group");
            btnGroup.alignment = "center";
            var btnCancel = btnGroup.add("button", undefined, L("btnCancel"), { name: "cancel" });
            var btnOK = btnGroup.add("button", undefined, L("btnOK"), { name: "ok" });

            // 現在の UI 状態からセッション保存ペイロードを作る / Build session payload from current UI
            function buildSessionPayload() {
                return {
                    addW: inputW.text,
                    addH: inputH.text,
                    radius: inputR.text,
                    radiusEnabled: chkRadiusEnabled.value,
                    link: linkCheck.value,
                    pill: chkPill.value
                };
            }

            btnOK.onClick = function () {
                confirmedDialogValues = ctrl.getFinalValues();
                if (!confirmedDialogValues) return;

                saveSessionState(buildSessionPayload());

                // 前回プレビューを巻き戻して、本番として 1 回だけ確定 /
                // Undo last preview then commit as the final operation
                ctrl.commitFinal();

                // commitFinal の末尾で previewItem を選択しているので、そこから取得 /
                // commitFinal leaves the previewItem selected
                var sel = app.activeDocument.selection;
                if (sel && sel.length > 0) {
                    finalPreviewItem = sel[0];
                }

                win.close(1);
            };

            btnCancel.onClick = function () {
                saveSessionState(buildSessionPayload());
                // 残ったプレビューは win.onClose の cleanupPreview で巻き戻す / onClose handles undo
                win.close(0);
            };

            // ダイアログがどう閉じられても（OK / Cancel / Esc / X ボタン）残プレビューを片付ける /
            // Catch-all: revert any leftover preview when the dialog closes
            win.onClose = function () {
                cleanupPreview(previewState);
            };

            win.onShow = function () {
                inputW.active = true;
                inputW.selection = [0, inputW.text.length];
                if (chkPill.value) {
                    linkCheck.value = false;
                }
                ctrl.reflectEnabled();
                ctrl.refresh();
            };

            // 初回プレビューは win.onShow（コールバック）から起動する。main() 同期側で呼ぶと
            // 後続コールバックの app.undo() が main の setup（テンプレート作成等）まで巻き戻すリスクがある。
            // Initial preview is fired from win.onShow (callback) — calling it from main()'s sync
            // context risks subsequent app.undo() rolling back template setup.

            // ダイアログ表示 / Show dialog
            if (win.show() === 1) {
                var finalValues = confirmedDialogValues || ctrl.getFinalValues();
                if (!finalValues) {
                    if (finalPreviewItem) {
                        safeRemove(finalPreviewItem);
                        finalPreviewItem = null;
                    }
                    return null;
                }
                return {
                    addW: finalValues.addW,
                    addH: finalValues.addH,
                    radius: finalValues.radius,
                    radiusEnabled: finalValues.radiusEnabled,
                    shouldRunPathfinder: finalValues.shouldRunPathfinder,
                    adjustEnabled: finalValues.adjustEnabled,
                    previewItem: finalPreviewItem
                };
            } else {
                // キャンセル時は最後の app.undo() で内部は未適用に戻り済み / On cancel, last app.undo() already reverted
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

            // 空になったグループと旧図形を削除 / Remove the empty group and old shape
            safeRemove(groupItem);
            safeRemove(shapeChild);

            return contentChild;
        }

        /* 選択を検証して { contentItem, shapeItem, shapeIsAutoCreated } を返す。不正なら alert を出して null。
           Validate selection. Returns parsed object or null (with alert) on invalid input. */
        function parseSelection(sel) {
            if (sel.length < 1 || sel.length > 2) {
                alert(L("alertSelectOne"), L("alertSelectError"));
                return null;
            }

            var contentItem = null;
            var shapeItem = null;
            var shapeIsAutoCreated = false;

            if (sel.length === 2) {
                // 2つ選択：テキスト/グループ＋図形 / Two items: text/group + shape
                for (var i = 0; i < sel.length; i++) {
                    var item = sel[i];
                    if (isContentItem(item)) {
                        if (contentItem) {
                            alert(L("alertSelectOne"), L("alertSelectError"));
                            return null;
                        }
                        contentItem = item;
                    } else if (isShapeItem(item)) {
                        if (shapeItem) {
                            alert(L("alertSelectOne"), L("alertSelectError"));
                            return null;
                        }
                        shapeItem = item;
                    } else {
                        alert(L("alertSelectOne"), L("alertSelectError"));
                        return null;
                    }
                }
                if (!contentItem || !shapeItem) {
                    alert(L("alertSelectOne"), L("alertSelectError"));
                    return null;
                }
            } else {
                // 1つ選択：テキスト/グループのみ。グループはテキスト＋図形なら分解して採用 /
                // One item: text/group; if a group contains text+shape, extract and adopt
                var selectedItem = sel[0];
                if (selectedItem.typename === "GroupItem" && !isClippingGroupItem(selectedItem)) {
                    contentItem = tryExtractContentFromGroup(selectedItem);
                }
                if (!contentItem) {
                    if (!isContentItem(selectedItem)) {
                        alert(L("alertSelectOne"), L("alertSelectError"));
                        return null;
                    }
                    contentItem = selectedItem;
                }
                shapeIsAutoCreated = true;
            }

            if (isClippingGroupItem(contentItem)) {
                alert(L("alertClippingGroupNotSupported"), L("alertSelectError"));
                return null;
            }

            return {
                contentItem: contentItem,
                shapeItem: shapeItem,
                shapeIsAutoCreated: shapeIsAutoCreated
            };
        }

        /* コンテンツを複製して bounds を計測。失敗時 alert + null。
           Duplicate content (outline if text) and read bounds. Returns boundsInfo or null. */
        function measureContent(contentItem) {
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
                safeRemove(measureItem);
                safeRemove(dupText);
            }

            if (!boundsInfo) {
                alert(L("alertMeasureFailed"), L("alertSelectError"));
                return null;
            }
            return boundsInfo;
        }

        /* auto モードなら長方形を作成。プレビュー用テンプレートを用意し、元図形を非表示にする。
           Auto-create rectangle if needed, prepare preview templates, hide original. */
        function prepareShapeAndPreviews(doc, contentItem, shapeItem, shapeIsAutoCreated, boundsInfo) {
            if (shapeIsAutoCreated) {
                // 初期不透明度は定数管理 / Initial opacity controlled by constant
                shapeItem = doc.pathItems.rectangle(
                    boundsInfo.top, boundsInfo.left, boundsInfo.width, boundsInfo.height
                );
                shapeItem.opacity = AUTO_CREATED_SHAPE_OPACITY;
                shapeItem.move(contentItem, ElementPlacement.PLACEAFTER);
            }

            var previewSourceShapeItem = null;
            var previewBaseShapeItem = null;

            if (!shapeIsAutoCreated) {
                // 2つ選択：整列専用と調整専用の2つ / Two-item: align-only + adjust-only
                previewSourceShapeItem = shapeItem.duplicate();
                previewSourceShapeItem.hidden = true;

                previewBaseShapeItem = shapeItem.duplicate();
                previewBaseShapeItem.hidden = false;
                clearAppearanceByAction(previewBaseShapeItem);
                previewBaseShapeItem.hidden = true;
            } else {
                // 1つ選択：ベース図形のみ / One-item: base only
                previewBaseShapeItem = shapeItem.duplicate();
                previewBaseShapeItem.hidden = true;
            }

            shapeItem.hidden = true;
            app.redraw();

            return {
                shapeItem: shapeItem,
                previewSourceShapeItem: previewSourceShapeItem,
                previewBaseShapeItem: previewBaseShapeItem
            };
        }

        /* OK 確定：使わないテンプレートを削除、必要なら Live Pathfinder、元図形削除、最終選択を設定。
           On OK: drop unused templates, run pathfinder if needed, remove original, set final selection. */
        function commitDialogResult(doc, result, prepared, contentItem, shapeIsAutoCreated) {
            var previewSourceShapeItem = prepared.previewSourceShapeItem;
            var previewBaseShapeItem = prepared.previewBaseShapeItem;
            var shapeItem = prepared.shapeItem;

            if (previewSourceShapeItem && previewSourceShapeItem !== result.previewItem) {
                safeRemove(previewSourceShapeItem);
                previewSourceShapeItem = null;
            }
            if (previewBaseShapeItem && previewBaseShapeItem !== result.previewItem) {
                safeRemove(previewBaseShapeItem);
                previewBaseShapeItem = null;
            }

            if (result.previewItem) {
                var finalPreviewItem = result.previewItem;
                finalPreviewItem.hidden = false;

                var baseShapeRef = result.adjustEnabled ? previewBaseShapeItem : previewSourceShapeItem;
                if (baseShapeRef && baseShapeRef !== finalPreviewItem) {
                    try { baseShapeRef.hidden = true; } catch (e) { }
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

                try {
                    shapeItem.remove();
                } catch (removeError) {
                    if (!shapeIsAutoCreated) shapeItem.hidden = false;
                    throw removeError;
                }

                doc.selection = null;
                contentItem.selected = true;
                finalPreviewItem.selected = true;
            }

            app.redraw();
        }

        /* キャンセル：テンプレート削除、auto なら長方形も削除、それ以外は元図形を再表示。
           On cancel: drop templates; drop or restore the original shape. */
        function cancelDialogResult(prepared, shapeIsAutoCreated) {
            safeRemove(prepared.previewSourceShapeItem);
            safeRemove(prepared.previewBaseShapeItem);
            if (shapeIsAutoCreated) {
                safeRemove(prepared.shapeItem);
            } else {
                prepared.shapeItem.hidden = false;
            }
            app.redraw();
        }

        /* メイン処理 / Main process */
        function main() {
            if (app.documents.length === 0) {
                alert(L("alertNoDocument"));
                return;
            }

            var doc = app.activeDocument;
            var parsed = parseSelection(doc.selection);
            if (!parsed) return;

            var boundsInfo = measureContent(parsed.contentItem);
            if (!boundsInfo) return;

            var prepared = prepareShapeAndPreviews(
                doc, parsed.contentItem, parsed.shapeItem, parsed.shapeIsAutoCreated, boundsInfo
            );

            var result = showDialog(
                boundsInfo.width, boundsInfo.height,
                prepared.previewSourceShapeItem, prepared.previewBaseShapeItem,
                boundsInfo.centerX, boundsInfo.centerY,
                parsed.shapeIsAutoCreated
            );

            if (result) {
                commitDialogResult(doc, result, prepared, parsed.contentItem, parsed.shapeIsAutoCreated);
            } else {
                cancelDialogResult(prepared, parsed.shapeIsAutoCreated);
            }
        }

        main();

    })();