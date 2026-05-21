#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

スクリプトの概要 / Script overview
選択したオブジェクトの塗りと線を調整するスクリプトです。
Tool for adjusting fill and stroke on selected objects.

対応モード / Supported modes
・塗り↔線 / Fill ↔ Stroke
・塗り→線 / Fill → Stroke
・線→塗り / Stroke → Fill
・2つのオブジェクト間で交換 / Swap Between 2 Objects
・塗りを消去 / Erase Fill
・線を消去 / Erase Stroke
・塗りと線を消去 / Erase Fill and Stroke

主な機能 / Main features
・選択は最大 2 オブジェクトまで / Up to two objects can be selected
・グラデーション対応 / Supports gradients
・複合パス、グループ対応 / Supports compound paths and groups
・テキスト対応（ポイント文字、エリア文字、パス上文字） / Supports point text, area text, and text on a path
・文字単位の処理で部分スタイルを維持 / Preserves partial text styling by processing text character by character
・必要に応じて線幅を補完 / Adds a stroke width when needed
・ダイアログで処理モードを選択 / Choose the processing mode in a dialog
・プレビュー時はキャンセルで元の見た目に復元 / Preview restores the original appearance when canceled
・OK 後に元の選択を復元 / Restores the original selection after OK

参考：オリジナル / Original
しぶやみゃむさんのスクリプトをベースに、機能追加とリファクタリングを行いました。
Based on a script by しぶやみゃむさん, with added features and refactoring.
https://note.com/shibumi/n/n5229b4357dd3

*/


(function () {

    // =========================================
    // バージョンとローカライズ
    // Version and localization
    // =========================================

    var SCRIPT_VERSION = "v1.1.0";

    /* 現在の言語コードを取得 / Detect current UI language */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {

        // UI
        dialogTitle: {
            ja: "塗りと線の調整",
            en: "Fill and Stroke Adjustments"
        },
        convertPanelTitle: {
            ja: "変換",
            en: "Convert"
        },
        removePanelTitle: {
            ja: "消去",
            en: "Erase"
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
        modeSwapBetween: {
            ja: "2つのオブジェクト間で交換",
            en: "Swap Between 2 Objects"
        },
        modeFillNone: {
            ja: "塗りを消去",
            en: "Erase Fill"
        },
        modeStrokeNone: {
            ja: "線を消去",
            en: "Erase Stroke"
        },
        modeFillStrokeNone: {
            ja: "塗りと線を消去",
            en: "Erase Fill and Stroke"
        },
        preview: {
            ja: "プレビュー",
            en: "Preview"
        },
        cancel: {
            ja: "キャンセル",
            en: "Cancel"
        },
        ok: {
            ja: "OK",
            en: "OK"
        },

        // Help tips
        modeSwapTip: {
            ja: "選択オブジェクトごとに、塗りと線を入れ替えます。",
            en: "Swaps fill and stroke within each selected object."
        },
        modeFillToStrokeTip: {
            ja: "塗りの色を線に適用します。必要に応じて線幅を補完します。",
            en: "Applies the fill color to the stroke. Adds a stroke width when needed."
        },
        modeStrokeToFillTip: {
            ja: "線の色を塗りに適用します。",
            en: "Applies the stroke color to the fill."
        },
        modeSwapBetweenTip: {
            ja: "選択した2つのオブジェクトの塗りと線の見た目を交換します。",
            en: "Swaps the fill and stroke appearance between the two selected objects."
        },
        modeFillNoneTip: {
            ja: "選択オブジェクトの塗りをなしにします。",
            en: "Removes the fill from selected objects."
        },
        modeStrokeNoneTip: {
            ja: "選択オブジェクトの線をなしにします。",
            en: "Removes the stroke from selected objects."
        },
        modeFillStrokeNoneTip: {
            ja: "選択オブジェクトの塗りと線をどちらもなしにします。",
            en: "Removes both fill and stroke from selected objects."
        },
        previewTip: {
            ja: "結果を一時的に表示します。キャンセル時は元に戻します。",
            en: "Temporarily shows the result. Restores the original appearance when canceled."
        },

        // Alerts
        noSelection: {
            ja: "オブジェクトを選択してください",
            en: "Please select at least one object."
        },
        tooManyObjects: {
            ja: "オブジェクトは 2 つまでにしてください",
            en: "Please select up to two objects."
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
        selectionRestoreFailures: {
            ja: "選択の復元失敗",
            en: "Selection restoration failures"
        },
        details: {
            ja: "詳細:",
            en: "Details:"
        }
    };

    /* キーから現在言語のラベルを取得 / Look up a localized label */
    function L(key) {
        if (!LABELS[key]) {
            return key;
        }
        return LABELS[key][currentLanguage] || LABELS[key].en || key;
    }

    var MODE_SWAP = "swap";
    var MODE_FILL_TO_STROKE = "fillToStroke";
    var MODE_STROKE_TO_FILL = "strokeToFill";
    var MODE_SWAP_BETWEEN = "swapBetween";
    var MODE_FILL_NONE = "fillNone";
    var MODE_STROKE_NONE = "strokeNone";
    var MODE_FILL_AND_STROKE_NONE = "fillAndStrokeNone";

    var PANEL_MARGINS = [15, 20, 15, 10];
    var PANEL_SPACING = 8;

    /* パネルに共通のレイアウトを適用 / Apply shared layout settings to a panel */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ['fill', 'top'];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* 処理結果の集計オブジェクトを生成 / Create a fresh processing-stats object */
    function createProcessStats() {
        return {
            pathFailureCount: 0,
            textFailureCount: 0,
            selectionRestoreFailureCount: 0,
            failureDetails: []
        };
    }

    /* 失敗の詳細を集計に追加 / Append a failure detail entry to stats */
    function addFailureDetail(stats, category, item, error) {
        if (!stats || !stats.failureDetails) {
            return;
        }
        if (stats.failureDetails.length >= 8) {
            return;
        }

        var typename = 'Unknown';
        var name = '';
        var message = 'Unknown error';
        try {
            if (item && item.typename) typename = item.typename;
            if (item && item.name) name = String(item.name);
            if (error && error.message) {
                message = String(error.message);
            } else if (error) {
                message = String(error);
            }
        } catch (e) { }

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
                var rgb = new RGBColor();
                rgb.red = color.red;
                rgb.green = color.green;
                rgb.blue = color.blue;
                return rgb;

            case "CMYKColor":
                var cmyk = new CMYKColor();
                cmyk.cyan = color.cyan;
                cmyk.magenta = color.magenta;
                cmyk.yellow = color.yellow;
                cmyk.black = color.black;
                return cmyk;

            case "GrayColor":
                var gray = new GrayColor();
                gray.gray = color.gray;
                return gray;

            case "SpotColor":
                var spot = new SpotColor();
                spot.spot = color.spot;
                spot.tint = color.tint;
                return spot;

            case "GradientColor":
                var gradient = new GradientColor();
                gradient.gradient = color.gradient;
                gradient.angle = color.angle;
                gradient.length = color.length;
                gradient.origin = color.origin;
                gradient.matrix = color.matrix;
                return gradient;

            case "NoColor":
                return new NoColor();

            default:
                return null;
        }
    }

    /* NoColor を生成 / Create a NoColor instance */
    function createNoColor() {
        return new NoColor();
    }

    /* テキスト用に有効な色か判定 / Check if color is usable for text */
    function isUsableTextColor(color) {
        return color && color.typename && color.typename !== "NoColor";
    }

    /* TextRange の塗り色を安全取得 / Safely read fillColor from a text range */
    function getTextRangeFillColor(textRange) {
        try {
            return textRange.characterAttributes.fillColor;
        } catch (e) {
            return null;
        }
    }

    /* TextRange の線色を安全取得 / Safely read strokeColor from a text range */
    function getTextRangeStrokeColor(textRange) {
        try {
            return textRange.characterAttributes.strokeColor;
        } catch (e) {
            return null;
        }
    }

    /* TextRange が有効な塗りを持つか / Does the text range have a usable fill */
    function hasTextRangeFill(textRange) {
        return isUsableTextColor(getTextRangeFillColor(textRange));
    }

    /* TextRange が有効な線を持つか / Does the text range have a usable stroke */
    function hasTextRangeStroke(textRange) {
        return isUsableTextColor(getTextRangeStrokeColor(textRange));
    }

    /* TextRange にモードに応じた塗り/線処理を適用 / Apply mode to a text range */
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

    /* ラジオボタンの排他選択を実現 / Make one radio exclusively selected in a group */
    function setExclusiveRadio(selectedRadio, radioButtons) {
        for (var i = 0; i < radioButtons.length; i++) {
            radioButtons[i].value = (radioButtons[i] === selectedRadio);
        }
    }

    /* 選択中のラジオボタンから処理モードを取得 / Get the mode from selected radio buttons */
    function getModeFromRadios(radios) {
        if (radios.fillToStrokeRadio.value) return MODE_FILL_TO_STROKE;
        if (radios.strokeToFillRadio.value) return MODE_STROKE_TO_FILL;
        if (radios.swapBetweenObjectsRadio.value) return MODE_SWAP_BETWEEN;
        if (radios.eraseFillRadio.value) return MODE_FILL_NONE;
        if (radios.eraseStrokeRadio.value) return MODE_STROKE_NONE;
        if (radios.eraseFillAndStrokeRadio.value) return MODE_FILL_AND_STROKE_NONE;
        return MODE_SWAP;
    }

    /* ラジオボタンの選択とプレビュー更新を接続 / Bind radio selection and preview refresh */
    function bindModeRadioEvents(radioButtons, refreshPreview) {
        function bindExclusiveRadio(radio) {
            radio.onClick = function () {
                setExclusiveRadio(radio, radioButtons);
                refreshPreview();
            };
        }

        for (var i = 0; i < radioButtons.length; i++) {
            bindExclusiveRadio(radioButtons[i]);
        }
    }

    // =========================================
    // メイン処理
    // Main processing
    // =========================================

    /* パスに処理モードを適用 / Apply the processing mode to a path */
    function applyPathFillStrokeMode(item, mode, stats) {
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
                    if (!hasFill && !hasStroke) break;
                    item.filled = hasStroke;
                    item.fillColor = (hasStroke && stroke) ? stroke : createNoColor();
                    item.stroked = hasFill;
                    if (hasFill && fill) {
                        item.strokeColor = fill;
                        if (!hasStroke || !item.strokeWidth || item.strokeWidth <= 0) {
                            item.strokeWidth = 1;
                        }
                    } else {
                        item.strokeColor = createNoColor();
                    }
                    break;
            }
        } catch (e) {
            if (stats) {
                stats.pathFailureCount++;
            }
            addFailureDetail(stats, 'Path', item, e);
        }
    }

    /* テキストに処理モードを適用 / Apply the processing mode to text */
    function applyTextFillStrokeMode(textFrame, mode, stats) {
        try {
            var textRange = textFrame.textRange;
            var characters = null;
            try { characters = textRange.characters; } catch (eC) { characters = null; }

            if (characters && characters.length > 0) {
                for (var i = 0; i < characters.length; i++) {
                    applyTextFillAndStrokeToRange(characters[i], mode);
                }
            } else {
                applyTextFillAndStrokeToRange(textRange, mode);
            }
        } catch (e) {
            if (stats) {
                stats.textFailureCount++;
            }
            addFailureDetail(stats, 'Text', textFrame, e);
        }
    }

    // =========================================
    // 状態の取得と適用
    // State capture and apply
    // =========================================

    /* PathItem の塗り/線スタイルを取得 / Capture path appearance */
    function capturePathAppearance(path) {
        return {
            filled: path.filled,
            stroked: path.stroked,
            fill: path.filled ? cloneColor(path.fillColor) : null,
            stroke: path.stroked ? cloneColor(path.strokeColor) : null,
            strokeWidth: path.strokeWidth
        };
    }

    /* TextRange の塗り/線スタイルを取得 / Capture text range appearance */
    function captureTextRangeAppearance(textRange) {
        var attrs = textRange.characterAttributes;
        var hasFill = hasTextRangeFill(textRange);
        var hasStroke = hasTextRangeStroke(textRange);
        var weight = attrs.strokeWeight;
        return {
            filled: hasFill,
            stroked: hasStroke,
            fill: hasFill ? cloneColor(getTextRangeFillColor(textRange)) : null,
            stroke: hasStroke ? cloneColor(getTextRangeStrokeColor(textRange)) : null,
            strokeWidth: weight
        };
    }

    /* PathItem に塗り/線スタイルを適用 / Apply appearance to path */
    function applyPathAppearance(path, appearance) {
        path.filled = appearance.filled;
        path.fillColor = (appearance.filled && appearance.fill) ? appearance.fill : createNoColor();
        path.stroked = appearance.stroked;
        if (appearance.stroked && appearance.stroke) {
            path.strokeColor = appearance.stroke;
            if (appearance.strokeWidth && appearance.strokeWidth > 0) {
                path.strokeWidth = appearance.strokeWidth;
            }
        } else {
            path.strokeColor = createNoColor();
        }
    }

    /* TextRange に塗り/線スタイルを適用 / Apply appearance to text range */
    function applyTextRangeAppearance(textRange, appearance) {
        var attrs = textRange.characterAttributes;
        attrs.fillColor = (appearance.filled && appearance.fill) ? appearance.fill : createNoColor();
        if (appearance.stroked && appearance.stroke) {
            attrs.strokeColor = appearance.stroke;
            if (appearance.strokeWidth && appearance.strokeWidth > 0) {
                attrs.strokeWeight = appearance.strokeWidth;
            }
        } else {
            attrs.strokeColor = createNoColor();
        }
    }

    /* 単一の見た目をスナップショット / Snapshot a single overall appearance */
    function snapshotAppearance(item) {
        var typename = item.typename;
        if (typename === "PathItem") return capturePathAppearance(item);
        if (typename === "CompoundPathItem") {
            if (!item.pathItems || item.pathItems.length === 0) {
                throw new Error("CompoundPathItem has no pathItems");
            }
            return capturePathAppearance(item.pathItems[0]);
        }
        if (typename === "TextFrame") return captureTextRangeAppearance(item.textRange);
        throw new Error("Unsupported type for swap: " + typename);
    }

    /* 単一の見た目を要素全体に均一適用 / Apply a single appearance uniformly to an item */
    function applyAppearance(item, appearance) {
        var typename = item.typename;
        if (typename === "PathItem") {
            applyPathAppearance(item, appearance);
            return;
        }
        if (typename === "CompoundPathItem") {
            for (var i = 0; i < item.pathItems.length; i++) {
                applyPathAppearance(item.pathItems[i], appearance);
            }
            return;
        }
        if (typename === "TextFrame") {
            var range = item.textRange;
            var chars = null;
            try { chars = range.characters; } catch (eC) { chars = null; }
            if (chars && chars.length > 0) {
                for (var j = 0; j < chars.length; j++) {
                    applyTextRangeAppearance(chars[j], appearance);
                }
            } else {
                applyTextRangeAppearance(range, appearance);
            }
            return;
        }
        throw new Error("Unsupported type for swap: " + typename);
    }

    /* 2つのオブジェクト間で塗りと線をスワップ / Swap fill and stroke between two items */
    function swapAppearanceBetween(itemA, itemB, stats) {
        try {
            var snapA = snapshotAppearance(itemA);
            var snapB = snapshotAppearance(itemB);
            applyAppearance(itemA, snapB);
            applyAppearance(itemB, snapA);
        } catch (e) {
            if (stats) {
                stats.pathFailureCount++;
            }
            addFailureDetail(stats, 'SwapBetween', itemA, e);
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
                    applyPathFillStrokeMode(item, mode, stats);
                    break;

                case "CompoundPathItem":
                    processItems(item.pathItems, mode, stats);
                    break;

                case "TextFrame":
                    applyTextFillStrokeMode(item, mode, stats);
                    break;
            }
        }
    }

    // =========================================
    // プレビュー用スナップショット
    // Preview snapshot
    // =========================================

    /* プレビュー用に1要素の状態をスナップショット / Snapshot a leaf item for preview restore */
    function snapshotLeafForPreview(item) {
        var typename = item.typename;
        if (typename === "PathItem") {
            var snap = capturePathAppearance(item);
            snap.kind = "path";
            snap.item = item;
            return snap;
        }
        if (typename === "TextFrame") {
            var textSnap = { kind: "text", item: item, characters: [] };
            try {
                var chars = item.textRange.characters;
                for (var i = 0; i < chars.length; i++) {
                    textSnap.characters.push(captureTextRangeAppearance(chars[i]));
                }
            } catch (e) { }
            return textSnap;
        }
        return null;
    }

    /* 選択中のリーフ要素を再帰的にスナップショット / Recursively snapshot selection leaves */
    function captureLeafSnapshots(items, snapshots) {
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            switch (item.typename) {
                case "GroupItem":
                    captureLeafSnapshots(item.pageItems, snapshots);
                    break;
                case "PathItem":
                case "TextFrame":
                    var snap = snapshotLeafForPreview(item);
                    if (snap) snapshots.push(snap);
                    break;
                case "CompoundPathItem":
                    captureLeafSnapshots(item.pathItems, snapshots);
                    break;
            }
        }
    }

    /* 選択全体のスナップショットを作成 / Capture full selection state */
    function captureSelectionState(items) {
        var snapshots = [];
        captureLeafSnapshots(items, snapshots);
        return snapshots;
    }

    /* 1要素をスナップショットから復元 / Restore a single leaf from its snapshot */
    function restoreLeafSnapshot(snap) {
        if (snap.kind === "path") {
            applyPathAppearance(snap.item, snap);
            return;
        }
        if (snap.kind === "text") {
            var chars = null;
            try { chars = snap.item.textRange.characters; } catch (eC) { return; }
            if (!chars) return;
            var pairCount = chars.length < snap.characters.length ? chars.length : snap.characters.length;
            for (var i = 0; i < pairCount; i++) {
                try { applyTextRangeAppearance(chars[i], snap.characters[i]); } catch (eR) { }
            }
        }
    }

    /* スナップショットから選択を一括復元 / Restore selection from snapshots */
    function restoreSelectionState(snapshots) {
        for (var i = 0; i < snapshots.length; i++) {
            var snap = snapshots[i];
            if (!snap) continue;
            try { restoreLeafSnapshot(snap); } catch (e) { }
        }
    }

    /* 選択オブジェクトを復元 / Restore selected objects */
    function restoreSelectedItems(items, stats) {
        var restoredItems = [];
        for (var i = 0; i < items.length; i++) {
            try {
                if (items[i]) {
                    restoredItems.push(items[i]);
                }
            } catch (e) {
                if (stats) {
                    stats.selectionRestoreFailureCount++;
                }
                addFailureDetail(stats, 'Selection', items[i], e);
            }
        }

        try {
            app.selection = restoredItems;
        } catch (eR) {
            if (stats) {
                stats.selectionRestoreFailureCount++;
            }
            addFailureDetail(stats, 'Selection', null, eR);
        }
    }

    /* プレビュー：選択中の要素に処理モードを適用 / Apply mode for live preview */
    function applyPreview(items, mode) {
        if (mode === MODE_SWAP_BETWEEN) {
            if (items.length === 2) {
                swapAppearanceBetween(items[0], items[1], null);
            }
            return;
        }
        processItems(items, mode, null);
    }

    // =========================================
    // ダイアログ
    // Dialog
    // =========================================

    /* 処理モード選択ダイアログ / Show the processing mode dialog */
    function showModeDialog(originalSelection) {
        var snapshots = captureSelectionState(originalSelection);

        var dialog = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
        dialog.orientation = 'column';
        dialog.alignChildren = ['fill', 'top'];

        var panelsGroup = dialog.add('group');
        panelsGroup.orientation = 'row';
        panelsGroup.alignChildren = ['fill', 'fill'];

        var convertPanel = panelsGroup.add('panel', undefined, L('convertPanelTitle'));
        setupPanel(convertPanel);

        var removePanel = panelsGroup.add('panel', undefined, L('removePanelTitle'));
        setupPanel(removePanel);

        var swapRadio = convertPanel.add('radiobutton', undefined, L('modeSwap'));
        swapRadio.helpTip = L('modeSwapTip');
        var fillToStrokeRadio = convertPanel.add('radiobutton', undefined, L('modeFillToStroke'));
        fillToStrokeRadio.helpTip = L('modeFillToStrokeTip');
        var strokeToFillRadio = convertPanel.add('radiobutton', undefined, L('modeStrokeToFill'));
        strokeToFillRadio.helpTip = L('modeStrokeToFillTip');
        var swapBetweenObjectsRadio = convertPanel.add('radiobutton', undefined, L('modeSwapBetween'));
        swapBetweenObjectsRadio.helpTip = L('modeSwapBetweenTip');
        swapBetweenObjectsRadio.enabled = (originalSelection && originalSelection.length === 2);

        var eraseFillRadio = removePanel.add('radiobutton', undefined, L('modeFillNone'));
        eraseFillRadio.helpTip = L('modeFillNoneTip');
        var eraseStrokeRadio = removePanel.add('radiobutton', undefined, L('modeStrokeNone'));
        eraseStrokeRadio.helpTip = L('modeStrokeNoneTip');
        var eraseFillAndStrokeRadio = removePanel.add('radiobutton', undefined, L('modeFillStrokeNone'));
        eraseFillAndStrokeRadio.helpTip = L('modeFillStrokeNoneTip');

        var allRadios = [
            swapRadio,
            fillToStrokeRadio,
            strokeToFillRadio,
            swapBetweenObjectsRadio,
            eraseFillRadio,
            eraseStrokeRadio,
            eraseFillAndStrokeRadio
        ];

        var modeRadios = {
            swapRadio: swapRadio,
            fillToStrokeRadio: fillToStrokeRadio,
            strokeToFillRadio: strokeToFillRadio,
            swapBetweenObjectsRadio: swapBetweenObjectsRadio,
            eraseFillRadio: eraseFillRadio,
            eraseStrokeRadio: eraseStrokeRadio,
            eraseFillAndStrokeRadio: eraseFillAndStrokeRadio
        };

        function refreshPreview() {
            restoreSelectionState(snapshots);
            if (previewCheckbox.value) {
                applyPreview(originalSelection, getModeFromRadios(modeRadios));
            }
            app.redraw();
        }

        bindModeRadioEvents(allRadios, refreshPreview);

        var defaultRadio = (originalSelection && originalSelection.length === 2) ? swapBetweenObjectsRadio : swapRadio;
        setExclusiveRadio(defaultRadio, allRadios);

        var buttonGroup = dialog.add('group');
        buttonGroup.orientation = 'row';
        buttonGroup.alignment = ['fill', 'center'];
        buttonGroup.alignChildren = ['fill', 'center'];

        var leftCol = buttonGroup.add('group');
        leftCol.orientation = 'row';
        leftCol.alignment = ['left', 'center'];
        var previewCheckbox = leftCol.add('checkbox', undefined, L('preview'));
        previewCheckbox.helpTip = L('previewTip');
        previewCheckbox.onClick = refreshPreview;

        var spacerCol = buttonGroup.add('group');
        spacerCol.alignment = ['fill', 'center'];

        var rightCol = buttonGroup.add('group');
        rightCol.orientation = 'row';
        rightCol.alignment = ['right', 'center'];
        rightCol.add('button', undefined, L('cancel'), { name: 'cancel' });
        rightCol.add('button', undefined, L('ok'), { name: 'ok' });

        if (dialog.show() !== 1) {
            restoreSelectionState(snapshots);
            app.redraw();
            return null;
        }

        return {
            mode: getModeFromRadios(modeRadios),
            previewApplied: previewCheckbox.value
        };
    }

    // =========================================
    // 実行
    // Run
    // =========================================

    /* エントリポイント / Entry point */
    function main() {
        if (app.documents.length === 0) {
            alert(L('noDocument'));
            return;
        }

        if (app.selection.length === 0) {
            alert(L('noSelection'));
            return;
        }

        if (app.selection.length > 2) {
            alert(L('tooManyObjects'));
            return;
        }

        var stats = createProcessStats();
        var originalSelection = [];
        for (var i = 0; i < app.selection.length; i++) {
            originalSelection.push(app.selection[i]);
        }

        var dialogResult = showModeDialog(originalSelection);
        if (dialogResult === null) {
            return;
        }

        var mode = dialogResult.mode;
        var alreadyApplied = dialogResult.previewApplied;

        try {
            if (!alreadyApplied) {
                if (mode === MODE_SWAP_BETWEEN) {
                    swapAppearanceBetween(originalSelection[0], originalSelection[1], stats);
                } else {
                    processItems(originalSelection, mode, stats);
                }
            }
        } finally {
            restoreSelectedItems(originalSelection, stats);
        }

        var messages = [];
        if (stats.pathFailureCount > 0) {
            messages.push(L('pathFailures') + ': ' + stats.pathFailureCount);
        }
        if (stats.textFailureCount > 0) {
            messages.push(L('textFailures') + ': ' + stats.textFailureCount);
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