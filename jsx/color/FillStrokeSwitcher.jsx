#target illustrator

    /*
     * スクリプトの概要 / Script overview
     * 選択したオブジェクトの塗りと線を操作するスクリプトです。
     * Tool for editing fill and stroke on selected objects.
     *
     * 対応モード / Supported modes
     * ・塗り↔線 / Fill ↔ Stroke
     * ・塗り→線 / Fill → Stroke
     * ・線→塗り / Stroke → Fill
     * ・塗りを削除 / Remove Fill
     * ・線を削除 / Remove Stroke
     * ・塗りと線を削除 / Remove Fill and Stroke
     *
     * 主な機能 / Main features
     * ・複数オブジェクト対応 / Supports multiple objects
     * ・グラデーション対応 / Supports gradients
     * ・複合パス、グループ対応 / Supports compound paths and groups
     * ・テキスト対応（ポイント文字、エリア文字、パス上文字） / Supports point text, area text, and text on a path
     * ・文字単位の処理で部分スタイルを維持 / Preserves partial text styling by processing text character by character
     * ・必要に応じて線幅を補完 / Adds a stroke width when needed
     * ・ダイアログで処理モードを選択 / Choose the processing mode in a dialog
     *
     * 参考：オリジナル / Original
     * しぶやみゃむさんのスクリプトをベースに、機能追加とリファクタリングを行いました。
     * Based on a script by しぶやみゃむさん, with added features and refactoring.
     * https://note.com/shibumi/n/n5229b4357dd3
     */


    (function () {
        app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

        // =========================================
        // バージョンとローカライズ
        // Version and localization
        // =========================================

        var SCRIPT_VERSION = "v1.0";

        function getCurrentLang() {
            return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
        }
        var lang = getCurrentLang();

        /* 日英ラベル定義 / Japanese-English label definitions */
        var LABELS = {
            dialogTitle: {
                ja: "塗り・線の変換",
                en: "Fill/Stroke Converter"
            },
            convertPanelTitle: {
                ja: "変換",
                en: "Convert"
            },
            removePanelTitle: {
                ja: "削除",
                en: "Remove"
            },
            modeSwap: {
                ja: "塗り↔線",
                en: "Fill ↔ Stroke"
            },
            modeFillToStroke: {
                ja: "塗り→線",
                en: "Fill → Stroke"
            },
            modeStrokeToFill: {
                ja: "線→塗り",
                en: "Stroke → Fill"
            },
            modeFillNone: {
                ja: "塗りを削除",
                en: "Remove Fill"
            },
            modeStrokeNone: {
                ja: "線を削除",
                en: "Remove Stroke"
            },
            modeFillStrokeNone: {
                ja: "塗りと線を削除",
                en: "Remove Fill and Stroke"
            },
            cancel: {
                ja: "キャンセル",
                en: "Cancel"
            },
            ok: {
                ja: "OK",
                en: "OK"
            },
            noSelection: {
                ja: "オブジェクトを選択してください",
                en: "Please select at least one object."
            },
            noDocument: {
                ja: "ドキュメントが開かれていません",
                en: "No document is open."
            },
            pathFailures: {
                ja: "パス処理の失敗",
                en: "Path processing failures"
            },
            textFailures: {
                ja: "テキスト処理の失敗",
                en: "Text processing failures"
            },
            actionFailures: {
                ja: "アクション実行の失敗",
                en: "Action execution failures"
            },
            selectionRestoreFailures: {
                ja: "選択の復元失敗",
                en: "Selection restoration failures"
            },
            details: {
                ja: "詳細:",
                en: "Details:"
            }
        };

        function L(key) {
            return LABELS[key][lang];
        }

        var MODE_SWAP = "swap";
        var MODE_FILL_TO_STROKE = "fillToStroke";
        var MODE_STROKE_TO_FILL = "strokeToFill";
        var MODE_FILL_NONE = "fillNone";
        var MODE_STROKE_NONE = "strokeNone";
        var MODE_FILL_AND_STROKE_NONE = "fillAndStrokeNone";

        function createProcessStats() {
            return {
                pathFailureCount: 0,
                textFailureCount: 0,
                selectionRestoreFailureCount: 0,
                actionFailureCount: 0,
                failureDetails: []
            };
        }

        function addFailureDetail(stats, category, item, error) {
            if (!stats || !stats.failureDetails) {
                return;
            }
            if (stats.failureDetails.length >= 8) {
                return;
            }

            var typename = 'Unknown';
            try {
                if (item && item.typename) {
                    typename = item.typename;
                }
            } catch (e1) { }

            var name = '';
            try {
                if (item && item.name) {
                    name = String(item.name);
                }
            } catch (e2) { }

            var message = 'Unknown error';
            try {
                if (error && error.message) {
                    message = String(error.message);
                } else if (error) {
                    message = String(error);
                }
            } catch (e3) { }

            var detail = category + ': ' + typename;
            if (name !== '') {
                detail += ' [' + name + ']';
            }
            detail += ' - ' + message;

            stats.failureDetails.push(detail);
        }

        // =========================================
        // 色処理
        // Color handling
        // =========================================

        /* 色を安全にコピー / Safely clone color */
        function cloneColor(color) {
            if (!color) return null;

            switch (color.typename) {

                case "RGBColor":
                    var c = new RGBColor();
                    c.red = color.red;
                    c.green = color.green;
                    c.blue = color.blue;
                    return c;

                case "CMYKColor":
                    var c2 = new CMYKColor();
                    c2.cyan = color.cyan;
                    c2.magenta = color.magenta;
                    c2.yellow = color.yellow;
                    c2.black = color.black;
                    return c2;

                case "GrayColor":
                    var c3 = new GrayColor();
                    c3.gray = color.gray;
                    return c3;

                case "SpotColor":
                    var c4 = new SpotColor();
                    c4.spot = color.spot;
                    c4.tint = color.tint;
                    return c4;

                case "GradientColor":
                    var g = new GradientColor();
                    g.gradient = color.gradient;
                    g.angle = color.angle;
                    g.length = color.length;
                    g.origin = color.origin;
                    g.matrix = color.matrix;
                    return g;

                case "NoColor":
                    return new NoColor();

                default:
                    return null;
            }
        }

        function createNoColor() {
            return new NoColor();
        }

        function isUsableTextColor(color) {
            return color && color.typename && color.typename !== "NoColor";
        }

        function getTextRangeFillColor(textRange) {
            try {
                return textRange.characterAttributes.fillColor;
            } catch (e) {
                return null;
            }
        }

        function getTextRangeStrokeColor(textRange) {
            try {
                return textRange.characterAttributes.strokeColor;
            } catch (e) {
                return null;
            }
        }

        function hasTextRangeFill(textRange) {
            return isUsableTextColor(getTextRangeFillColor(textRange));
        }

        function hasTextRangeStroke(textRange) {
            return isUsableTextColor(getTextRangeStrokeColor(textRange));
        }

        function applyTextFillAndStrokeToRange(textRange, mode) {
            var attrs = textRange.characterAttributes;
            var hasFill = hasTextRangeFill(textRange);
            var hasStroke = hasTextRangeStroke(textRange);

            var fill = hasFill ? cloneColor(getTextRangeFillColor(textRange)) : null;
            var stroke = hasStroke ? cloneColor(getTextRangeStrokeColor(textRange)) : null;

            switch (mode) {
                case MODE_FILL_TO_STROKE:
                    if (hasFill && fill) {
                        attrs.strokeColor = fill;
                        if (!hasStroke || !attrs.strokeWeight || attrs.strokeWeight <= 0) {
                            attrs.strokeWeight = 1;
                        }
                    }
                    break;

                case MODE_STROKE_TO_FILL:
                    if (hasStroke && stroke) {
                        attrs.fillColor = stroke;
                    }
                    break;

                case MODE_FILL_NONE:
                    attrs.fillColor = createNoColor();
                    break;

                case MODE_STROKE_NONE:
                    attrs.strokeColor = createNoColor();
                    break;

                case MODE_FILL_AND_STROKE_NONE:
                    attrs.fillColor = createNoColor();
                    attrs.strokeColor = createNoColor();
                    break;

                case MODE_SWAP:
                default:
                    if (!hasFill && !hasStroke) {
                        return;
                    }

                    if (hasStroke && stroke) {
                        attrs.fillColor = stroke;
                    } else {
                        attrs.fillColor = createNoColor();
                    }

                    if (hasFill && fill) {
                        attrs.strokeColor = fill;
                        if (!hasStroke || !attrs.strokeWeight || attrs.strokeWeight <= 0) {
                            attrs.strokeWeight = 1;
                        }
                    } else {
                        attrs.strokeColor = createNoColor();
                    }
                    break;
            }
        }
        /* 対象オブジェクトのみを選択 / Select only the target item */
        function selectOnlyItem(targetItem) {
            var doc = app.activeDocument;
            doc.selection = null;
            targetItem.selected = true;
        }

        function setExclusiveRadio(selectedRadio, radioButtons) {
            for (var i = 0; i < radioButtons.length; i++) {
                radioButtons[i].value = (radioButtons[i] === selectedRadio);
            }
        }

        // =========================================
        // メイン処理
        // Main processing
        // =========================================

        /* 塗りと線の処理を実行 / Apply fill and stroke processing */
        function swapFillAndStroke(item, mode, stats) {
            try {
                var hasFill = item.filled;
                var hasStroke = item.stroked;

                var fill = hasFill ? cloneColor(item.fillColor) : null;
                var stroke = hasStroke ? cloneColor(item.strokeColor) : null;

                switch (mode) {
                    case MODE_FILL_TO_STROKE:
                        if (hasFill && fill) {
                            item.stroked = true;
                            item.strokeColor = fill;
                            if (!hasStroke) {
                                item.strokeWidth = 1;
                            }
                        }
                        break;

                    case MODE_STROKE_TO_FILL:
                        if (hasStroke && stroke) {
                            item.filled = true;
                            item.fillColor = stroke;
                        }
                        break;

                    case MODE_FILL_NONE:
                        item.filled = false;
                        item.fillColor = createNoColor();
                        break;

                    case MODE_STROKE_NONE:
                        item.stroked = false;
                        item.strokeColor = createNoColor();
                        break;

                    case MODE_FILL_AND_STROKE_NONE:
                        item.filled = false;
                        item.stroked = false;
                        item.fillColor = createNoColor();
                        item.strokeColor = createNoColor();
                        break;

                    case MODE_SWAP:
                    default:
                        if (!hasStroke && hasFill) {
                            item.strokeWidth = 1;
                        }
                        selectOnlyItem(item);
                        act_Swap(stats);
                        break;
                }
            } catch (e) {
                if (stats) {
                    stats.pathFailureCount++;
                }
                addFailureDetail(stats, 'Path', item, e);
            }
        }

        /* テキストの塗りと線の処理を実行 / Apply fill and stroke processing to text */
        function swapTextFillAndStroke(textFrame, mode, stats) {
            try {
                var textRange = textFrame.textRange;
                var characters = null;
                var count = 0;
                var processed = false;

                try {
                    characters = textRange.characters;
                    count = characters.length;
                } catch (e1) {
                    characters = null;
                    count = 0;
                }

                if (characters && count > 0) {
                    for (var i = 0; i < count; i++) {
                        applyTextFillAndStrokeToRange(characters[i], mode);
                    }
                    processed = true;
                }

                if (!processed) {
                    applyTextFillAndStrokeToRange(textRange, mode);
                }
            } catch (e) {
                if (stats) {
                    stats.textFailureCount++;
                }
                addFailureDetail(stats, 'Text', textFrame, e);
            }
        }

        /* 選択オブジェクトを再帰処理 / Process selected items recursively */
        function processItems(items, mode, stats) {
            for (var i = 0; i < items.length; i++) {
                var item = items[i];

                switch (item.typename) {

                    case "GroupItem":
                        processItems(item.pageItems, mode, stats);
                        break;

                    case "PathItem":
                        swapFillAndStroke(item, mode, stats);
                        break;

                    case "CompoundPathItem":
                        processItems(item.pathItems, mode, stats);
                        break;

                    case "TextFrame":
                        swapTextFillAndStroke(item, mode, stats);
                        break;
                }
            }
        }

        // =========================================
        // ダイアログ
        // Dialog
        // =========================================

        /* 処理モード選択ダイアログ / Show the processing mode dialog */
        function showModeDialog() {
            var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
            dlg.orientation = 'column';
            dlg.alignChildren = ['fill', 'top'];

            var panelsGroup = dlg.add('group');
            panelsGroup.orientation = 'row';
            panelsGroup.alignChildren = ['fill', 'fill'];

            var convertPanel = panelsGroup.add('panel', undefined, L('convertPanelTitle'));
            convertPanel.orientation = 'column';
            convertPanel.alignChildren = ['left', 'top'];
            convertPanel.margins = [15, 20, 15, 10];

            var removePanel = panelsGroup.add('panel', undefined, L('removePanelTitle'));
            removePanel.orientation = 'column';
            removePanel.alignChildren = ['left', 'top'];
            removePanel.margins = [15, 20, 15, 10];

            var rbSwap = convertPanel.add('radiobutton', undefined, L('modeSwap'));
            var rbFillToStroke = convertPanel.add('radiobutton', undefined, L('modeFillToStroke'));
            var rbStrokeToFill = convertPanel.add('radiobutton', undefined, L('modeStrokeToFill'));

            var rbFillNone = removePanel.add('radiobutton', undefined, L('modeFillNone'));
            var rbStrokeNone = removePanel.add('radiobutton', undefined, L('modeStrokeNone'));
            var rbFillStrokeNone = removePanel.add('radiobutton', undefined, L('modeFillStrokeNone'));

            var allRadios = [
                rbSwap,
                rbFillToStroke,
                rbStrokeToFill,
                rbFillNone,
                rbStrokeNone,
                rbFillStrokeNone
            ];

            function bindExclusiveRadio(radio) {
                radio.onClick = function () {
                    setExclusiveRadio(radio, allRadios);
                };
            }

            for (var i = 0; i < allRadios.length; i++) {
                bindExclusiveRadio(allRadios[i]);
            }

            setExclusiveRadio(rbSwap, allRadios);

            var buttonGroup = dlg.add('group');
            buttonGroup.orientation = 'row';
            buttonGroup.alignment = ['center', 'center'];
            var cancelBtn = buttonGroup.add('button', undefined, L('cancel'), { name: 'cancel' });
            var okBtn = buttonGroup.add('button', undefined, L('ok'), { name: 'ok' });

            if (dlg.show() !== 1) {
                return null;
            }

            if (rbFillToStroke.value) return MODE_FILL_TO_STROKE;
            if (rbStrokeToFill.value) return MODE_STROKE_TO_FILL;
            if (rbFillNone.value) return MODE_FILL_NONE;
            if (rbStrokeNone.value) return MODE_STROKE_NONE;
            if (rbFillStrokeNone.value) return MODE_FILL_AND_STROKE_NONE;
            return MODE_SWAP;
        }

        // =========================================
        // Illustratorアクション
        // Illustrator action
        // =========================================

        /* Swap アクションを実行 / Run the Swap action */
        function act_Swap(stats) {
            var actionSetName = 'Color';
            var actionName = 'Swap';
            var tempFileName = 'ScriptAction_' + new Date().getTime() + '_' + Math.floor(Math.random() * 100000) + '.aia';
            var f = new File(Folder.temp.fsName + '/' + tempFileName);
            var str = '/version 3' +
                '/name [ 5' +
                ' 436f6c6f72' +
                ']' +
                '/isOpen 1' +
                '/actionCount 1' +
                '/action-1 {' +
                ' /name [ 4' +
                ' 53776170' +
                ' ]' +
                ' /keyIndex 0' +
                ' /colorIndex 0' +
                ' /isOpen 1' +
                ' /eventCount 1' +
                ' /event-1 {' +
                ' /useRulersIn1stQuadrant 0' +
                ' /internalName (ai_plugin_setColor)' +
                ' /localizedName [ 18' +
                ' e382abe383a9e383bce38292e8a8ade5ae9a' +
                ' ]' +
                ' /isOpen 1' +
                ' /isOn 1' +
                ' /hasDialog 0' +
                ' /parameterCount 1' +
                ' /parameter-1 {' +
                ' /key 1836349808' +
                ' /showInPalette 4294967295' +
                ' /type (enumerated)' +
                ' /name [ 27' +
                ' e5a197e3828ae381a8e7b79ae38292e585a5e3828ce69bbfe38188' +
                ' ]' +
                ' /value 7' +
                ' }' +
                ' }' +
                '}';

            try {
                f.open('w');
                f.write(str);
                f.close();
                app.loadAction(f);
                app.doScript(actionName, actionSetName, false);
            } catch (e) {
                if (stats) {
                    stats.actionFailureCount++;
                }
                addFailureDetail(stats, 'Action', null, e);
                throw e;
            } finally {
                try {
                    app.unloadAction(actionSetName, '');
                } catch (e2) {
                    if (stats) {
                        stats.actionFailureCount++;
                    }
                    addFailureDetail(stats, 'Action unload', null, e2);
                }
                try {
                    if (f.exists) {
                        f.remove();
                    }
                } catch (e3) { }
            }
        }
        /* 選択状態を復元 / Restore the original selection */
        function restoreSelection(items, stats) {
            if (!items || items.length === 0) {
                return;
            }

            app.activeDocument.selection = null;
            for (var i = 0; i < items.length; i++) {
                try {
                    if (items[i]) {
                        items[i].selected = true;
                    }
                } catch (e) {
                    if (stats) {
                        stats.selectionRestoreFailureCount++;
                    }
                    addFailureDetail(stats, 'Selection restore', items[i], e);
                }
            }
        }

        // =========================================
        // 実行
        // Run
        // =========================================
        function main() {
            if (app.documents.length === 0) {
                alert(L('noDocument'));
                return;
            }

            if (app.selection.length === 0) {
                alert(L('noSelection'));
                return;
            }

            var stats = createProcessStats();
            var originalSelection = [];
            for (var i = 0; i < app.selection.length; i++) {
                originalSelection.push(app.selection[i]);
            }

            var mode = showModeDialog();
            if (mode === null) {
                return;
            }

            try {
                processItems(originalSelection, mode, stats);
            } finally {
                restoreSelection(originalSelection, stats);
            }

            var messages = [];
            if (stats.pathFailureCount > 0) {
                messages.push(L('pathFailures') + ': ' + stats.pathFailureCount);
            }
            if (stats.textFailureCount > 0) {
                messages.push(L('textFailures') + ': ' + stats.textFailureCount);
            }
            if (stats.actionFailureCount > 0) {
                messages.push(L('actionFailures') + ': ' + stats.actionFailureCount);
            }
            if (stats.selectionRestoreFailureCount > 0) {
                messages.push(L('selectionRestoreFailures') + ': ' + stats.selectionRestoreFailureCount);
            }
            if (stats.failureDetails.length > 0) {
                messages.push('');
                messages.push(L('details'));
                for (var j = 0; j < stats.failureDetails.length; j++) {
                    messages.push('- ' + stats.failureDetails[j]);
                }
            }
            if (messages.length > 0) {
                alert(messages.join('\n'));
            }
        }

        main();
    })();