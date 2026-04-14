#target illustrator

/* =========================================
概要 / Overview

選択したオブジェクトに対して「アピアランスの消去」を実行します。
実行前にダイアログボックスを表示し、復元するかどうかと、選択オブジェクトに応じた復元方法を選択できます。

パスオブジェクトは、元の塗り・線・線幅を再適用します。
「カラーのみ（塗りと線）」または「カラーと線の設定」を選択できます。
後者を選んだ場合は、線端・角の形状・破線・破線オフセット・角の比率も戻します。
「不透明度」「描画モード」「オーバープリント」が ON の場合は、それぞれ元の値を戻します。

テキストは、次のいずれかを選択できます。
- 復元しない
- 1文字目の塗りをテキスト全体に戻す
- 文字単位の塗りを戻す
テキストの線は戻しません。

パス、複合パス（親単位でアピアランスの消去のみ実行）、グループ内オブジェクト、テキストに対応します。
クリップグループは再帰処理せず、グループに対してアピアランスの消去のみ実行します。
処理後は選択状態を復元します。
失敗があった場合は件数と詳細を通知します。

---

Clears appearance from selected objects.
Before execution, a dialog lets you choose whether to restore attributes and which restore options are available for the current selection.

For path objects, the original fill, stroke, and stroke width are reapplied.
You can choose either “Color Only (Fill and Stroke)” or “Color + Stroke Settings”.
When the latter is selected, stroke cap, join, dash pattern, dash offset, and miter limit are also restored.
When “Opacity”, “Blending Mode”, or “Overprint” is enabled, the original value is restored for each option.

For text, you can choose one of the following:
- do not restore text fill
- restore the first character's fill to the whole text
- restore fill per character
Text stroke is not restored.

Supports paths, compound paths (cleared at parent level without restoring attributes), grouped items, and text.
Clipped groups are not processed recursively; the script only clears appearance on the group itself.
Restores the original selection after processing.
If any failures occur, the script reports counts and details.
Updated: 2026-04-14  |  Version: v1.0
========================================= */


    (function () {
        app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

        // =========================================
        // バージョンとローカライズ
        // Version and localization
        // =========================================

        function getCurrentLang() {
            return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
        }
        var lang = getCurrentLang();
        var SCRIPT_VERSION = "v1.0";

        /* 日英ラベル定義 / Japanese-English label definitions */
        var LABELS = {
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
            },
            detailCategoryPath: {
                ja: "パス",
                en: "Path"
            },
            detailCategoryText: {
                ja: "テキスト",
                en: "Text"
            },
            detailCategoryAction: {
                ja: "アクション",
                en: "Action"
            },
            detailCategoryActionUnload: {
                ja: "アクション解除",
                en: "Action unload"
            },
            detailCategorySelectionRestore: {
                ja: "選択復元",
                en: "Selection restore"
            },
            dialogTitle: {
                ja: "アピアランスの消去",
                en: "Clear Appearance"
            },
            dialogRestore: {
                ja: "復元する",
                en: "Restore"
            },
            dialogNoRestore: {
                ja: "復元しない",
                en: "Do not restore"
            },
            dialogFillStrokePanel: {
                ja: "パスオブジェクト",
                en: "Path Objects"
            },
            dialogTextPanel: {
                ja: "テキスト",
                en: "Text"
            },
            dialogOptionPanel: {
                ja: "オプション",
                en: "Options"
            },
            dialogPathColorOnly: {
                ja: "カラーのみ（塗りと線）",
                en: "Color Only (Fill and Stroke)"
            },
            dialogPathColorAndStrokeSettings: {
                ja: "カラーと線の設定",
                en: "Color + Stroke Settings"
            },
            dialogTextFillFirst: {
                ja: "テキストの塗り（1文字目）",
                en: "Text Fill (First Character)"
            },
            dialogTextFillPerChar: {
                ja: "テキストの塗り（文字単位）",
                en: "Text Fill (Per Character)"
            },
            dialogOpacity: {
                ja: "不透明度",
                en: "Opacity"
            },
            dialogBlendingMode: {
                ja: "描画モード",
                en: "Blending Mode"
            },
            dialogOverprint: {
                ja: "オーバープリント",
                en: "Overprint"
            },
            dialogTextNoRestore: {
                ja: "復元しない",
                en: "Do not restore"
            },
            dialogOk: {
                ja: "OK",
                en: "OK"
            },
            dialogCancel: {
                ja: "キャンセル",
                en: "Cancel"
            }
        };

        function L(key) {
            return LABELS[key][lang];
        }

        function detailCategoryLabel(categoryKey) {
            if (LABELS[categoryKey] && LABELS[categoryKey][lang]) {
                return LABELS[categoryKey][lang];
            }
            return String(categoryKey);
        }

        function createProcessStats() {
            return {
                pathFailureCount: 0,
                textFailureCount: 0,
                selectionRestoreFailureCount: 0,
                actionFailureCount: 0,
                failureDetails: []
            };
        }

        function addFailureDetail(stats, categoryKey, item, error) {
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

            var detail = detailCategoryLabel(categoryKey) + ': ' + typename;
            if (name !== '') {
                detail += ' [' + name + ']';
            }
            detail += ' - ' + message;

            stats.failureDetails.push(detail);
        }

        // =========================================
        // ダイアログ
        // Dialog
        // =========================================

        function collectSelectionAvailability(items, result) {
            if (!items || !result) {
                return;
            }

            for (var i = 0; i < items.length; i++) {
                var item = items[i];
                if (!item) {
                    continue;
                }

                switch (item.typename) {
                    case "GroupItem":
                        if (!item.clipped) {
                            collectSelectionAvailability(item.pageItems, result);
                        }
                        break;

                    case "PathItem":
                    case "CompoundPathItem":
                        result.hasPathObjects = true;
                        break;

                    case "TextFrame":
                        result.hasTextFrames = true;
                        break;
                }
            }
        }

        function getSelectionAvailability(items) {
            var result = {
                hasPathObjects: false,
                hasTextFrames: false
            };
            collectSelectionAvailability(items, result);
            return result;
        }

        /* 復元オプションダイアログを表示 / Show restore option dialog */
        function showRestoreDialog() {
            var dlg = new Window('dialog', L('dialogTitle') + '  ' + SCRIPT_VERSION);
            dlg.orientation = 'column';
            dlg.alignChildren = ['fill', 'top'];
            dlg.margins = [15, 20, 15, 15];
            dlg.spacing = 10;

            var availability = getSelectionAvailability(app.selection);

            /* 復元する／復元しない / Restore or not */
            var restoreGroup = dlg.add('group');
            restoreGroup.orientation = 'row';
            restoreGroup.alignment = ['center', 'top'];
            restoreGroup.alignChildren = ['center', 'center'];
            var rbRestore = restoreGroup.add('radiobutton', undefined, L('dialogRestore'));
            var rbNoRestore = restoreGroup.add('radiobutton', undefined, L('dialogNoRestore'));
            rbRestore.value = true;

            /* パスオブジェクト / Path Objects */
            var pathPanel = dlg.add('panel', undefined, L('dialogFillStrokePanel'));
            pathPanel.orientation = 'column';
            pathPanel.alignment = ['fill', 'top'];
            pathPanel.alignChildren = ['left', 'top'];
            pathPanel.margins = [15, 20, 15, 10];
            pathPanel.spacing = 8;

            var rbPathColorOnly = pathPanel.add('radiobutton', undefined, L('dialogPathColorOnly'));
            var rbPathColorAndStrokeSettings = pathPanel.add('radiobutton', undefined, L('dialogPathColorAndStrokeSettings'));
            rbPathColorAndStrokeSettings.value = true;

            /* テキスト / Text */
            var textPanel = dlg.add('panel', undefined, L('dialogTextPanel'));
            textPanel.orientation = 'column';
            textPanel.alignment = ['fill', 'top'];
            textPanel.alignChildren = ['left', 'top'];
            textPanel.margins = [15, 20, 15, 10];
            textPanel.spacing = 8;

            var rbTextNoRestore = textPanel.add('radiobutton', undefined, L('dialogTextNoRestore'));
            var rbTextFillFirst = textPanel.add('radiobutton', undefined, L('dialogTextFillFirst'));
            var rbTextFillPerChar = textPanel.add('radiobutton', undefined, L('dialogTextFillPerChar'));
            rbTextFillPerChar.value = true;

            /* オプション / Options */
            var checkPanel = dlg.add('panel', undefined, L('dialogOptionPanel'));
            checkPanel.orientation = 'column';
            checkPanel.alignment = ['fill', 'top'];
            checkPanel.alignChildren = ['left', 'top'];
            checkPanel.margins = [15, 20, 15, 10];
            checkPanel.spacing = 8;

            var cbOpacity = checkPanel.add('checkbox', undefined, L('dialogOpacity'));
            var cbBlendingMode = checkPanel.add('checkbox', undefined, L('dialogBlendingMode'));
            var cbOverprint = checkPanel.add('checkbox', undefined, L('dialogOverprint'));
            cbOpacity.value = true;
            cbBlendingMode.value = true;
            cbOverprint.value = true;

            function updateEnabled() {
                var restoreEnabled = rbRestore.value;

                pathPanel.enabled = restoreEnabled && availability.hasPathObjects;
                rbPathColorOnly.enabled = restoreEnabled && availability.hasPathObjects;
                rbPathColorAndStrokeSettings.enabled = restoreEnabled && availability.hasPathObjects;

                textPanel.enabled = restoreEnabled && availability.hasTextFrames;
                rbTextNoRestore.enabled = restoreEnabled && availability.hasTextFrames;
                rbTextFillFirst.enabled = restoreEnabled && availability.hasTextFrames;
                rbTextFillPerChar.enabled = restoreEnabled && availability.hasTextFrames;

                checkPanel.enabled = restoreEnabled;
                cbOpacity.enabled = restoreEnabled;
                cbBlendingMode.enabled = restoreEnabled;
                cbOverprint.enabled = restoreEnabled;
            }

            rbRestore.onClick = updateEnabled;
            rbNoRestore.onClick = updateEnabled;
            updateEnabled();

            /* ボタン / Buttons */
            var btnGroup = dlg.add('group');
            btnGroup.alignment = ['center', 'center'];
            btnGroup.alignChildren = ['center', 'center'];
            btnGroup.add('button', undefined, L('dialogCancel'), { name: 'cancel' });
            btnGroup.add('button', undefined, L('dialogOk'), { name: 'ok' });

            if (dlg.show() !== 1) return null;

            if (!rbRestore.value) {
                return {
                    fillStroke: false,
                    textFillFirst: false,
                    textFillPerChar: false,
                    strokeSettings: false,
                    opacity: false,
                    blendingMode: false,
                    overprint: false
                };
            }

            return {
                fillStroke: availability.hasPathObjects,
                textFillFirst: availability.hasTextFrames && rbTextFillFirst.value,
                textFillPerChar: availability.hasTextFrames && rbTextFillPerChar.value,
                strokeSettings: availability.hasPathObjects && rbPathColorAndStrokeSettings.value,
                opacity: cbOpacity.value,
                blendingMode: cbBlendingMode.value,
                overprint: cbOverprint.value
            };
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

                case "PatternColor":
                    var p = new PatternColor();
                    p.pattern = color.pattern;
                    try {
                        p.matrix = color.matrix;
                    } catch (e4) { }
                    return p;

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

        function getTextRangeFillColorClone(textRange) {
            var fill = getTextRangeFillColor(textRange);
            return isUsableTextColor(fill) ? cloneColor(fill) : null;
        }

        function restoreTextFillOnlyToRange(textRange, fill) {
            var attrs = textRange.characterAttributes;

            if (fill) {
                attrs.fillColor = cloneColor(fill);
            } else {
                attrs.fillColor = createNoColor();
            }

            attrs.strokeColor = createNoColor();
        }
        /* 対象オブジェクトのみを選択 / Select only the target item */
        function selectOnlyItem(targetItem) {
            var doc = app.activeDocument;
            doc.selection = null;
            targetItem.selected = true;
        }

        /* クリップグループ・複合パス・非復元時はアピアランスの消去のみ実行し、必要に応じてオブジェクト属性だけ戻す / For clipped groups, compound paths, or no-restore cases, clear appearance only and optionally restore object-level attributes */
        function clearAppearanceOnly(item, stats, opts) {
            try {
                var opacity = null;
                var blendingMode = null;

                if (opts && opts.opacity) {
                    try { opacity = item.opacity; } catch (e1) { }
                }
                if (opts && opts.blendingMode) {
                    try { blendingMode = item.blendingMode; } catch (e2) { }
                }

                selectOnlyItem(item);
                act_Clear(stats);

                try {
                    if (opacity !== null) {
                        item.opacity = opacity;
                    }
                } catch (e3) { }

                try {
                    if (blendingMode !== null) {
                        item.blendingMode = blendingMode;
                    }
                } catch (e4) { }
            } catch (e) {
                if (stats) {
                    stats.pathFailureCount++;
                }
                var categoryKey = 'detailCategoryPath';
                try {
                    if (item && item.typename === 'TextFrame') {
                        categoryKey = 'detailCategoryText';
                    }
                } catch (e5) { }
                addFailureDetail(stats, categoryKey, item, e);
            }
        }

        // =========================================
        // メイン処理
        // Main processing
        // =========================================

        /* パスオブジェクトのアピアランスを消去し、塗り・線・線幅を復元 / Clear path appearance and restore fill, stroke, and stroke width */
        // destructive clear → restore visual attributes for path objects
        function clearAppearancePreserveFillStroke(item, stats, opts) {
            try {
                /* 元の塗り・線・線幅を記憶 / Remember original fill, stroke, and stroke width */
                var hasFill = item.filled;
                var hasStroke = item.stroked;
                var fill = hasFill ? cloneColor(item.fillColor) : null;
                var stroke = hasStroke ? cloneColor(item.strokeColor) : null;
                var strokeWidth = hasStroke ? item.strokeWidth : 1;

                /* 線属性を記憶 / Remember stroke attributes */
                var strokeCap = null;
                var strokeJoin = null;
                var strokeDashes = null;
                var strokeDashOffset = null;
                var strokeMiterLimit = null;
                var strokeOverprint = null;
                var fillOverprint = null;

                if (hasFill && opts.overprint) {
                    try { fillOverprint = item.fillOverprint; } catch (e0) { }
                }

                if (hasStroke && opts.strokeSettings) {
                    try { strokeCap = item.strokeCap; } catch (e1) { }
                    try { strokeJoin = item.strokeJoin; } catch (e2) { }
                    try { strokeDashes = item.strokeDashes ? item.strokeDashes.slice(0) : null; } catch (e3) { }
                    try { strokeDashOffset = item.strokeDashOffset; } catch (e4) { }
                    try { strokeMiterLimit = item.strokeMiterLimit; } catch (e5) { }
                }

                if (hasStroke && opts.overprint) {
                    try { strokeOverprint = item.strokeOverprint; } catch (e6) { }
                }

                /* 不透明度と描画モードを記憶 / Remember opacity and blending mode */
                var opacity = null;
                var blendingMode = null;
                if (opts.opacity) {
                    try { opacity = item.opacity; } catch (e13) { }
                }
                if (opts.blendingMode) {
                    try { blendingMode = item.blendingMode; } catch (e14) { }
                }

                /* アピアランスを消去 / Clear appearance */
                selectOnlyItem(item);
                act_Clear(stats);

                /* 記憶した塗り・線をそのまま再適用 / Reapply original fill and stroke */
                if (hasFill && fill) {
                    item.filled = true;
                    item.fillColor = fill;
                    if (opts.overprint) {
                        try { if (fillOverprint !== null) item.fillOverprint = fillOverprint; } catch (e6a) { }
                    }
                } else {
                    item.filled = false;
                    item.fillColor = createNoColor();
                }

                if (hasStroke && stroke) {
                    item.stroked = true;
                    item.strokeColor = stroke;
                    item.strokeWidth = strokeWidth;

                    /* 線属性を復元 / Restore stroke attributes */
                    if (opts.strokeSettings) {
                        try { if (strokeCap !== null) item.strokeCap = strokeCap; } catch (e7) { }
                        try { if (strokeJoin !== null) item.strokeJoin = strokeJoin; } catch (e8) { }
                        try { if (strokeDashes !== null) item.strokeDashes = strokeDashes; } catch (e9) { }
                        try { if (strokeDashOffset !== null) item.strokeDashOffset = strokeDashOffset; } catch (e10) { }
                        try { if (strokeMiterLimit !== null) item.strokeMiterLimit = strokeMiterLimit; } catch (e11) { }
                    }
                    if (opts.overprint) {
                        try { if (strokeOverprint !== null) item.strokeOverprint = strokeOverprint; } catch (e12) { }
                    }
                } else {
                    item.stroked = false;
                    item.strokeColor = createNoColor();
                }

                /* 不透明度と描画モードを復元 / Restore opacity and blending mode */
                try {
                    if (opacity !== null) {
                        item.opacity = opacity;
                    }
                } catch (e15) { }

                try {
                    if (blendingMode !== null) {
                        item.blendingMode = blendingMode;
                    }
                } catch (e16) { }
            } catch (e) {
                if (stats) {
                    stats.pathFailureCount++;
                }
                addFailureDetail(stats, 'detailCategoryPath', item, e);
            }
        }

        /* テキストのアピアランスを消去し、設定に応じて塗りのみ復元 / Clear text appearance and restore fill only based on the selected mode */
        // destructive clear → restore text fill only
        function clearTextAppearancePreserveFill(textFrame, stats, opts) {
            try {
                var textRange = textFrame.textRange;
                var characters = null;
                var count = 0;
                var characterFills = [];
                var hasCharacters = false;

                try {
                    characters = textRange.characters;
                    count = characters.length;
                    hasCharacters = (count > 0);
                } catch (e1) {
                    characters = null;
                    count = 0;
                    hasCharacters = false;
                }

                /* 文字単位の塗りを記憶 / Remember per-character fills */
                if (opts.textFillPerChar && hasCharacters) {
                    for (var i = 0; i < count; i++) {
                        characterFills.push(getTextRangeFillColorClone(characters[i]));
                    }
                }

                /* 1文字目の塗りを記憶 / Remember first character fill */
                var firstCharFill = null;
                if (opts.textFillFirst && hasCharacters) {
                    firstCharFill = getTextRangeFillColorClone(characters[0]);
                }

                var rangeFill = getTextRangeFillColorClone(textRange);

                /* テキストフレームの不透明度と描画モードを記憶 / Remember text-frame opacity and blending mode */
                var opacity = null;
                var blendingMode = null;
                if (opts && opts.opacity) {
                    try { opacity = textFrame.opacity; } catch (e2) { }
                }
                if (opts && opts.blendingMode) {
                    try { blendingMode = textFrame.blendingMode; } catch (e3) { }
                }


                selectOnlyItem(textFrame);
                act_Clear(stats);

                if (opts.textFillPerChar && hasCharacters) {
                    for (var j = 0; j < count; j++) {
                        restoreTextFillOnlyToRange(characters[j], characterFills[j]);
                    }
                } else if (opts.textFillFirst) {
                    var fillToApply = firstCharFill || rangeFill;
                    restoreTextFillOnlyToRange(textRange, fillToApply);
                }

                /* テキストフレームの不透明度と描画モードを復元 / Restore text-frame opacity and blending mode */
                try {
                    if (opacity !== null) {
                        textFrame.opacity = opacity;
                    }
                } catch (e4) { }

                try {
                    if (blendingMode !== null) {
                        textFrame.blendingMode = blendingMode;
                    }
                } catch (e5) { }

            } catch (e) {
                if (stats) {
                    stats.textFailureCount++;
                }
                addFailureDetail(stats, 'detailCategoryText', textFrame, e);
            }
        }

        /* 選択オブジェクトを再帰処理 / Process selected items recursively */
        function processItems(items, stats, opts) {
            for (var i = 0; i < items.length; i++) {
                var item = items[i];

                switch (item.typename) {
                    case "GroupItem":
                        if (item.clipped) {
                            clearAppearanceOnly(item, stats, opts);
                        } else {
                            processItems(item.pageItems, stats, opts);
                        }
                        break;

                    case "PathItem":
                        if (opts.fillStroke) {
                            clearAppearancePreserveFillStroke(item, stats, opts);
                        } else {
                            clearAppearanceOnly(item, stats, opts);
                        }
                        break;

                    case "CompoundPathItem":
                        clearAppearanceOnly(item, stats, opts);
                        break;

                    case "TextFrame":
                        if (opts.textFillFirst || opts.textFillPerChar) {
                            clearTextAppearancePreserveFill(item, stats, opts);
                        } else {
                            clearAppearanceOnly(item, stats, opts);
                        }
                        break;
                }
            }
        }

        // =========================================
        // Illustratorアクション
        // Illustrator action
        // =========================================

        /* Clear Appearance アクションを実行 / Run the Clear Appearance action */
        function act_Clear(stats) {
            var actionSetName = 'Appearance';
            var actionName = 'Clear';
            var tempFileName = 'ScriptAction_' + new Date().getTime() + '_' + Math.floor(Math.random() * 100000) + '.aia';
            var f = new File(Folder.temp.fsName + '/' + tempFileName);
            var str = '/version 3/name [ 10 417070656172616e6365]/isOpen 1/actionCount 1/action-1 { /name [ 5 436c656172 ] /keyIndex 0 /colorIndex 0 /isOpen 1 /eventCount 1 /event-1 { /useRulersIn1stQuadrant 0 /internalName (ai_plugin_appearance) /localizedName [ 18 e382a2e38394e382a2e383a9e383b3e382b9 ] /isOpen 1 /isOn 1 /hasDialog 0 /parameterCount 1 /parameter-1 { /key 1835363957 /showInPalette 4294967295 /type (enumerated) /name [ 27 e382a2e38394e382a2e383a9e383b3e382b9e38292e6b688e58ebb ] /value 6 } } }';

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
                addFailureDetail(stats, 'detailCategoryAction', null, e);
                throw e;
            } finally {
                try {
                    app.unloadAction(actionSetName, '');
                } catch (e2) {
                    if (stats) {
                        stats.actionFailureCount++;
                    }
                    addFailureDetail(stats, 'detailCategoryActionUnload', null, e2);
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
                    addFailureDetail(stats, 'detailCategorySelectionRestore', items[i], e);
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

            var opts = showRestoreDialog();
            if (opts === null) return;

            var stats = createProcessStats();
            var originalSelection = [];
            for (var i = 0; i < app.selection.length; i++) {
                originalSelection.push(app.selection[i]);
            }

            try {
                processItems(originalSelection, stats, opts);
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