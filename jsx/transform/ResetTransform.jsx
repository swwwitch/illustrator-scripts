#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);
#targetengine "DialogEngine"

/*

### 概要

配置画像・テキスト・長方形・クリップグループ・直線パスに対して、回転／シアー／スケール／縦横比を安全にリセットします。
バウンディングボックスをリセットしたあと元の中心位置へ戻すため、見た目の位置は保たれます。

詳細は README を参照してください。

### Overview

Safely resets rotation, shear, scale and aspect ratio on placed images, text, rectangles, clipping groups and straight paths.
The bounding box is reset and the item is moved back to its original center, so its apparent position is preserved.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ResetTransform";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.6.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-05";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-11";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ResetTransform.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ResetTransform.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n52f6b645bc70"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 軸スナップの許容範囲（度）/ Angle range that counts as "near an axis" */
    var AXIS_SNAP_MIN_DEG = 0.5;    /* これ未満はすでに正立とみなす / below this the path is treated as upright */
    var AXIS_SNAP_MAX_DEG = 44;     /* これを超えると意図的な傾きとみなす / above this the tilt is treated as intentional */

    /* スケール入力の下限（%）/ Minimum scale percent allowed in the input */
    var SCALE_MIN_PERCENT = 20;

    /* 判定用の微小値 / Numerical tolerances */
    var MATRIX_EPSILON = 1e-8;         /* 行列演算のゼロ判定 / zero threshold for matrix math */
    var SCALE_EPSILON = 1e-6;          /* スケール・シアーの残差判定 / residual threshold for scale and shear */
    var ROTATION_EPSILON_DEG = 0.0001; /* 回転の残差判定（度）/ residual threshold for rotation (deg) */

    // =========================================
    // レイアウト設定 / Layout settings
    // =========================================

    var WINDOW_MARGINS = 16;                /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING = 8;                  /* パネル内の要素間隔 / panel spacing */
    var PANEL_SPACING_COMPACT = 6;          /* チェックボックスを並べるときの間隔 / spacing for checkbox stacks */
    var COLUMN_SPACING = 12;                /* 2カラムの間隔 / gap between columns */

    var DIALOG_OFFSET_X = 300;              /* 初回表示時の横シフト（px）/ first-run horizontal shift (px) */
    var DIALOG_OFFSET_Y = 0;                /* 初回表示時の縦シフト（px）/ first-run vertical shift (px) */
    var DIALOG_OPACITY = 0.98;              /* ダイアログの不透明度 / dialog opacity (0.0-1.0) */

    /* ダイアログ位置の記憶キー（targetengine 内で共有）/ Storage key for the dialog location */
    var DIALOG_LOCATION_KEY = "__ResetTransform_OptionsDialog";

    // =========================================
    // 単位 / Units
    // =========================================
    /* ルーラー単位の換算は使用しません（スケールは % 指定）/ No ruler-unit conversion (scale is percent-based) */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * UI言語を判定する。
     * @returns {string} 日本語環境なら "ja"、それ以外は "en"
     */
    function getCurrentLanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLanguage();

    /**
     * ラベル定義から現在の言語の文字列を取り出す。
     * @param {object} labelEntry - { ja: string, en: string } 形式のラベル定義
     * @returns {string} 現在の言語のラベル文字列
     */
    function L(labelEntry) {
        if (!labelEntry) return "";
        return String((labelEntry[currentLanguage] != null) ? labelEntry[currentLanguage] : labelEntry.en);
    }

    /* UIラベル（カテゴリ別）/ UI labels grouped by category */
    var LABELS = {
        dialog: {
            title: { ja: "リセット（回転・比率）", en: "Reset (Rotate / Scale)" }
        },
        panel: {
            placedImage: { ja: "配置画像", en: "Placed Images" },
            clippedGroup: { ja: "クリップグループ", en: "Clip Group" },
            textFrame: { ja: "テキスト", en: "Text" },
            rectanglePath: { ja: "長方形（パス）", en: "Rectangle (Path)" },
            straightLine: { ja: "パス（直線）", en: "Path (Line)" }
        },
        checkbox: {
            rotate: { ja: "回転", en: "Rotate" },
            shear: { ja: "シアー", en: "Shear" },
            aspectRatio: { ja: "縦横比", en: "Aspect Ratio" },
            flip: { ja: "反転", en: "Flip" },
            scale: { ja: "スケール", en: "Scale" },
            textScaleRatio: { ja: "垂直比率／水平比率", en: "Horizontal & Vertical Scale" }
        },
        button: {
            reset: { ja: "リセット", en: "Reset" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectFirst: { ja: "オブジェクトを選択してください。", en: "Please select an object." },
            noTarget: { ja: "リセットできる対象が選択されていません。", en: "No resettable objects are selected." }
        }
    };

    // =========================================
    // ダイアログ位置の記憶 / Dialog location memory
    // =========================================

    /**
     * 記憶しておいたダイアログ位置を取り出す。
     * @returns {array} [x, y] の座標配列。未保存なら null
     */
    function getSavedDialogLocation() {
        var savedLocation = $.global[DIALOG_LOCATION_KEY];
        return (savedLocation && savedLocation.length === 2) ? savedLocation : null;
    }

    /**
     * ダイアログ位置を記憶する（targetengine が生きている間だけ保持）。
     * @param {object} dialogWindow - 対象のダイアログ
     * @returns {void}
     */
    function saveDialogLocation(dialogWindow) {
        $.global[DIALOG_LOCATION_KEY] = [dialogWindow.location[0], dialogWindow.location[1]];
    }

    /**
     * ダイアログ位置を画面内に収める。
     * @param {array} location - [x, y] の座標配列
     * @returns {array} 画面内に収めた [x, y]
     */
    function clampLocationToScreen(location) {
        var visibleBounds = ($.screens && $.screens.length) ? $.screens[0].visibleBounds : [0, 0, 1920, 1080];
        return [
            Math.max(visibleBounds[0] + 10, Math.min(location[0], visibleBounds[2] - 10)),
            Math.max(visibleBounds[1] + 10, Math.min(location[1], visibleBounds[3] - 10))
        ];
    }

    /**
     * 前回位置の復元と、移動時の位置記憶をダイアログに組み込む。
     * @param {object} dialogWindow - 対象のダイアログ
     * @returns {void}
     */
    function bindDialogLocationMemory(dialogWindow) {
        var savedLocation = getSavedDialogLocation();
        dialogWindow.onShow = function () {
            dialogWindow.location = savedLocation ?
                clampLocationToScreen(savedLocation) :
                [dialogWindow.location[0] + DIALOG_OFFSET_X, dialogWindow.location[1] + DIALOG_OFFSET_Y];
        };
        dialogWindow.onMove = function () {
            saveDialogLocation(dialogWindow);
        };
    }

    // =========================================
    // UIヘルパー / UI helpers
    // =========================================

    /**
     * ウィンドウに共通のレイアウトを適用する。
     * @param {object} dialogWindow - 対象のウィンドウ
     * @returns {void}
     */
    function applyWindowLayout(dialogWindow) {
        dialogWindow.orientation = "column";
        dialogWindow.alignChildren = "fill";
        dialogWindow.margins = WINDOW_MARGINS;
        dialogWindow.spacing = WINDOW_SPACING;
    }

    /**
     * パネルに共通のレイアウトを適用する。
     * @param {object} targetPanel - 対象のパネル
     * @param {number} spacing - 要素間隔。省略時は PANEL_SPACING
     * @returns {void}
     */
    function applyPanelLayout(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * カラム用のグループを作成する。
     * @param {object} parentGroup - 親グループ
     * @returns {object} 作成した縦並びグループ
     */
    function addColumnGroup(parentGroup) {
        var columnGroup = parentGroup.add('group');
        columnGroup.orientation = 'column';
        columnGroup.alignChildren = 'left';
        columnGroup.alignment = 'fill';
        return columnGroup;
    }

    /**
     * パネルにチェックボックスを追加する。
     * @param {object} targetPanel - 追加先のパネル
     * @param {object} labelEntry - ラベル定義
     * @param {boolean} initialValue - 初期のオン／オフ
     * @returns {object} 追加したチェックボックス
     */
    function addCheckbox(targetPanel, labelEntry, initialValue) {
        var checkbox = targetPanel.add('checkbox', undefined, L(labelEntry));
        checkbox.value = initialValue;
        return checkbox;
    }

    /**
     * 入力欄にフォーカスして内容を全選択する。
     * @param {object} editText - 対象の入力欄
     * @returns {void}
     */
    function focusAndSelectAll(editText) {
        editText.active = true;
        editText.textselection = editText.text;
    }

    /**
     * 入力欄の値を↑↓キーで増減できるようにする。Shift併用で10の倍数へスナップする。
     * @param {object} editText - 対象の入力欄
     * @param {number} minimumValue - 下限値
     * @returns {void}
     */
    function changeValueByArrowKey(editText, minimumValue) {
        editText.addEventListener("keydown", function (event) {
            var inputValue = Number(editText.text);
            if (isNaN(inputValue)) return;

            var keyboardState = ScriptUI.environment.keyboardState;
            var stepAmount = keyboardState.shiftKey ? 10 : 1;

            if (keyboardState.shiftKey) {
                /* Shiftキー押下時は10の倍数にスナップ / snap to multiples of 10 when Shift is held */
                if (event.keyName === "Up") {
                    inputValue = Math.ceil((inputValue + 1) / stepAmount) * stepAmount;
                    event.preventDefault();
                } else if (event.keyName === "Down") {
                    inputValue = Math.floor((inputValue - 1) / stepAmount) * stepAmount;
                    event.preventDefault();
                }
            } else {
                if (event.keyName === "Up") {
                    inputValue += stepAmount;
                    event.preventDefault();
                } else if (event.keyName === "Down") {
                    inputValue -= stepAmount;
                    event.preventDefault();
                }
            }

            /* 整数に丸めて下限でクランプ / round to an integer and clamp to the minimum */
            inputValue = Math.round(inputValue);
            if (inputValue < minimumValue) inputValue = minimumValue;
            editText.text = inputValue;
        });
    }

    /**
     * 1文字のホットキーでチェックボックスを切り替えられるようにする。
     * @param {object} dialogWindow - 対象のダイアログ
     * @param {string} hotkeyChar - 割り当てる文字
     * @param {object} targetCheckbox - 切り替えるチェックボックス
     * @param {function} onToggle - 切り替え後に呼ぶ処理（省略可）
     * @returns {void}
     */
    function addHotkeyToggle(dialogWindow, hotkeyChar, targetCheckbox, onToggle) {
        dialogWindow.addEventListener('keydown', function (event) {
            var pressedKey = (event.keyName || '').toUpperCase();
            if (pressedKey !== String(hotkeyChar).toUpperCase()) return;
            event.preventDefault();
            if (!targetCheckbox.enabled) return;

            targetCheckbox.value = !targetCheckbox.value;
            if (typeof onToggle === 'function') onToggle(targetCheckbox.value);
        });
    }

    // =========================================
    // 選択状態の判定 / Selection analysis
    // =========================================

    /**
     * 4点の閉じたパス（長方形とみなす）かどうかを判定する。
     * @param {object} pathItem - 判定するパス
     * @returns {boolean} 長方形とみなせるなら true
     */
    function isRectanglePath(pathItem) {
        return !!(pathItem.closed && pathItem.pathPoints && pathItem.pathPoints.length === 4);
    }

    /**
     * 2点の開いたパス（直線）かどうかを判定する。
     * @param {object} pathItem - 判定するパス
     * @returns {boolean} 直線とみなせるなら true
     */
    function isStraightLinePath(pathItem) {
        return !!(!pathItem.closed && pathItem.pathPoints && pathItem.pathPoints.length === 2);
    }

    /**
     * 選択内容から、どの種別のリセットが使えるかを調べる。
     * @param {array} selectedItems - 選択中のページアイテム
     * @returns {object} 種別ごとの可否フラグ
     */
    function getSelectionCapabilities(selectedItems) {
        var capabilities = {
            hasPlacedOrRaster: false,
            hasClippedGroup: false,
            hasTextFrame: false,
            hasRectanglePath: false,
            hasStraightLine: false
        };
        if (!selectedItems || !selectedItems.length) return capabilities;

        for (var i = 0; i < selectedItems.length; i++) {
            var selectedItem = selectedItems[i];
            if (!selectedItem || !selectedItem.typename) continue;
            var typeName = selectedItem.typename;

            if (typeName === 'PlacedItem' || typeName === 'RasterItem') {
                capabilities.hasPlacedOrRaster = true;
            } else if (typeName === 'GroupItem' && selectedItem.clipped === true) {
                capabilities.hasClippedGroup = true;
            } else if (typeName === 'TextFrame') {
                capabilities.hasTextFrame = true;
            } else if (typeName === 'PathItem') {
                if (isRectanglePath(selectedItem)) capabilities.hasRectanglePath = true;
                if (isStraightLinePath(selectedItem)) capabilities.hasStraightLine = true;
            }
        }
        return capabilities;
    }

    // =========================================
    // ダイアログの各パネル / Dialog panels
    // =========================================

    /**
     * 配置画像パネルを作成する。
     * @param {object} parentGroup - 追加先のカラムグループ
     * @param {boolean} isEnabled - 選択内容に配置画像が含まれるか
     * @returns {object} パネル内のコントロール一式
     */
    function buildPlacedImagePanel(parentGroup, isEnabled) {
        var pnlPlacedImage = parentGroup.add('panel', undefined, L(LABELS.panel.placedImage));
        applyPanelLayout(pnlPlacedImage, PANEL_SPACING_COMPACT);

        var cbRotate = addCheckbox(pnlPlacedImage, LABELS.checkbox.rotate, true);
        var cbShear = addCheckbox(pnlPlacedImage, LABELS.checkbox.shear, true);
        var cbAspectRatio = addCheckbox(pnlPlacedImage, LABELS.checkbox.aspectRatio, true);
        var cbFlip = addCheckbox(pnlPlacedImage, LABELS.checkbox.flip, true);
        var cbScale = addCheckbox(pnlPlacedImage, LABELS.checkbox.scale, false);

        var scaleInputGroup = pnlPlacedImage.add('group');
        scaleInputGroup.orientation = 'row';
        scaleInputGroup.alignChildren = 'center';
        scaleInputGroup.alignment = 'left'; /* 入力欄はパネル幅いっぱいに広げない / keep the scale input compact */

        var etScalePercent = scaleInputGroup.add('edittext', undefined, '100');
        etScalePercent.characters = 5;
        var stPercentUnit = scaleInputGroup.add('statictext', undefined, '%');
        changeValueByArrowKey(etScalePercent, SCALE_MIN_PERCENT);

        /**
         * スケール入力欄の有効・無効をチェックボックスに連動させる。
         * @param {boolean} isScaleOn - スケールがオンかどうか
         * @returns {void}
         */
        function syncScaleInput(isScaleOn) {
            etScalePercent.enabled = isScaleOn;
            stPercentUnit.enabled = isScaleOn;
            if (isScaleOn) focusAndSelectAll(etScalePercent);
        }

        etScalePercent.enabled = cbScale.value;
        stPercentUnit.enabled = cbScale.value;
        cbScale.onClick = function () {
            syncScaleInput(cbScale.value);
        };

        /* パネルを無効化すれば子コントロールもまとめて無効になる / disabling the panel disables its children */
        pnlPlacedImage.enabled = isEnabled;

        return {
            rotate: cbRotate,
            shear: cbShear,
            aspectRatio: cbAspectRatio,
            flip: cbFlip,
            scale: cbScale,
            scalePercent: etScalePercent,
            syncScaleInput: syncScaleInput
        };
    }

    /**
     * クリップグループパネルを作成する。
     * @param {object} parentGroup - 追加先のカラムグループ
     * @param {boolean} isEnabled - 選択内容にクリップグループが含まれるか
     * @returns {object} パネル内のコントロール一式
     */
    function buildClippedGroupPanel(parentGroup, isEnabled) {
        var pnlClippedGroup = parentGroup.add('panel', undefined, L(LABELS.panel.clippedGroup));
        applyPanelLayout(pnlClippedGroup, PANEL_SPACING_COMPACT);

        var controls = {
            rotate: addCheckbox(pnlClippedGroup, LABELS.checkbox.rotate, true),
            aspectRatio: addCheckbox(pnlClippedGroup, LABELS.checkbox.aspectRatio, true),
            flip: addCheckbox(pnlClippedGroup, LABELS.checkbox.flip, true)
        };
        pnlClippedGroup.enabled = isEnabled;
        return controls;
    }

    /**
     * テキストパネルを作成する。
     * @param {object} parentGroup - 追加先のカラムグループ
     * @param {boolean} isEnabled - 選択内容にテキストが含まれるか
     * @returns {object} パネル内のコントロール一式
     */
    function buildTextFramePanel(parentGroup, isEnabled) {
        var pnlTextFrame = parentGroup.add('panel', undefined, L(LABELS.panel.textFrame));
        applyPanelLayout(pnlTextFrame, PANEL_SPACING_COMPACT);

        var controls = {
            rotate: addCheckbox(pnlTextFrame, LABELS.checkbox.rotate, true),
            shear: addCheckbox(pnlTextFrame, LABELS.checkbox.shear, true),
            scaleRatio: addCheckbox(pnlTextFrame, LABELS.checkbox.textScaleRatio, true)
        };
        pnlTextFrame.enabled = isEnabled;
        return controls;
    }

    /**
     * 回転チェックボックスだけを持つパネルを作成する（長方形・直線用）。
     * @param {object} parentGroup - 追加先のカラムグループ
     * @param {object} titleEntry - パネル見出しのラベル定義
     * @param {boolean} isEnabled - 選択内容に対象が含まれるか
     * @returns {object} 追加した回転チェックボックス
     */
    function buildRotateOnlyPanel(parentGroup, titleEntry, isEnabled) {
        var rotateOnlyPanel = parentGroup.add('panel', undefined, L(titleEntry));
        applyPanelLayout(rotateOnlyPanel, PANEL_SPACING_COMPACT);

        var cbRotate = addCheckbox(rotateOnlyPanel, LABELS.checkbox.rotate, true);
        rotateOnlyPanel.enabled = isEnabled;
        return cbRotate;
    }

    /**
     * ボタンエリア（左右分割）を作成する。
     * @param {object} dialogWindow - 対象のダイアログ
     * @returns {void}
     */
    function buildButtonRow(dialogWindow) {
        var btnRowGroup = dialogWindow.add('group');
        btnRowGroup.orientation = 'row';
        btnRowGroup.alignChildren = ['fill', 'center'];
        btnRowGroup.alignment = ['fill', 'top'];

        var spacer = btnRowGroup.add('group');
        spacer.alignment = ['fill', 'fill'];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add('group');
        btnRightGroup.orientation = 'row';
        btnRightGroup.alignChildren = ['right', 'center'];
        btnRightGroup.spacing = PANEL_SPACING;

        var btnCancel = btnRightGroup.add('button', undefined, L(LABELS.button.cancel), { name: 'cancel' });
        var btnReset = btnRightGroup.add('button', undefined, L(LABELS.button.reset), { name: 'ok' });

        /* 閉じる直前の位置を記憶 / remember the location just before closing */
        btnReset.onClick = function () {
            saveDialogLocation(dialogWindow);
            dialogWindow.close(1);
        };
        btnCancel.onClick = function () {
            saveDialogLocation(dialogWindow);
            dialogWindow.close(0);
        };
    }

    /**
     * 入力された倍率を整数％に整えて下限でクランプする。
     * @param {string} scaleText - 入力欄の文字列
     * @returns {number} 適用する倍率（%）
     */
    function parseScalePercent(scaleText) {
        var scalePercent = parseFloat(scaleText);
        if (isNaN(scalePercent)) scalePercent = 100;
        scalePercent = Math.round(scalePercent); /* 整数％に統一（例：16.3 → 16）/ keep it an integer percent */
        return (scalePercent < SCALE_MIN_PERCENT) ? SCALE_MIN_PERCENT : scalePercent;
    }

    /**
     * リセットオプションのダイアログを表示する。
     * @param {array} selectedItems - 選択中のページアイテム
     * @returns {object} 選択されたオプション。キャンセル時は null
     */
    function showResetOptionsDialog(selectedItems) {
        var capabilities = getSelectionCapabilities(selectedItems);

        var mainDialog = new Window('dialog', L(LABELS.dialog.title) + ' ' + SCRIPT_VERSION);
        applyWindowLayout(mainDialog);
        mainDialog.opacity = DIALOG_OPACITY;
        bindDialogLocationMemory(mainDialog);

        /* 2カラムレイアウト / two-column layout */
        var columnsGroup = mainDialog.add('group');
        columnsGroup.orientation = 'row';
        columnsGroup.alignment = 'fill';
        columnsGroup.alignChildren = 'top';
        columnsGroup.spacing = COLUMN_SPACING;

        var leftColumnGroup = addColumnGroup(columnsGroup);
        var rightColumnGroup = addColumnGroup(columnsGroup);

        var placedImageControls = buildPlacedImagePanel(leftColumnGroup, capabilities.hasPlacedOrRaster);
        var clippedGroupControls = buildClippedGroupPanel(leftColumnGroup, capabilities.hasClippedGroup);
        var textFrameControls = buildTextFramePanel(rightColumnGroup, capabilities.hasTextFrame);
        var cbRectangleRotate = buildRotateOnlyPanel(rightColumnGroup, LABELS.panel.rectanglePath, capabilities.hasRectanglePath);
        var cbStraightLineRotate = buildRotateOnlyPanel(rightColumnGroup, LABELS.panel.straightLine, capabilities.hasStraightLine);

        /* Sキーでスケール、Fキーで反転を切り替え / 'S' toggles Scale, 'F' toggles Flip */
        addHotkeyToggle(mainDialog, 'S', placedImageControls.scale, placedImageControls.syncScaleInput);
        addHotkeyToggle(mainDialog, 'F', placedImageControls.flip);

        buildButtonRow(mainDialog);

        if (mainDialog.show() !== 1) return null; /* キャンセル / cancelled */

        return {
            placedRotate: placedImageControls.rotate.value,
            placedShear: placedImageControls.shear.value,
            placedAspectRatio: placedImageControls.aspectRatio.value,
            placedFlip: placedImageControls.flip.value,
            placedScale: placedImageControls.scale.value,
            placedScalePercent: parseScalePercent(placedImageControls.scalePercent.text),
            clipRotate: clippedGroupControls.rotate.value,
            clipAspectRatio: clippedGroupControls.aspectRatio.value,
            clipFlip: clippedGroupControls.flip.value,
            textRotate: textFrameControls.rotate.value,
            textShear: textFrameControls.shear.value,
            textScaleRatio: textFrameControls.scaleRatio.value,
            rectangleRotate: cbRectangleRotate.value,
            straightLineRotate: cbStraightLineRotate.value
        };
    }

    // =========================================
    // 行列ユーティリティ / Matrix utilities
    // =========================================

    /**
     * 行列を持つオブジェクトかどうかを判定する。
     * PathItem や GroupItem は matrix を持たないため、ここで弾く。
     * @param {object} pageItem - 判定するページアイテム
     * @returns {boolean} matrix を参照できるなら true
     */
    function hasMatrix(pageItem) {
        try {
            return !!(pageItem && pageItem.matrix && typeof pageItem.matrix.mValueA !== 'undefined');
        } catch (e) {
            return false;
        }
    }

    /**
     * 2×2行列を掛け合わせる。
     * @param {object} leftComponents - 左側の成分 { a, b, c, d }
     * @param {object} rightComponents - 右側の成分 { a, b, c, d }
     * @returns {object} 積の成分 { a, b, c, d }
     */
    function multiply2x2(leftComponents, rightComponents) {
        return {
            a: leftComponents.a * rightComponents.a + leftComponents.c * rightComponents.b,
            b: leftComponents.b * rightComponents.a + leftComponents.d * rightComponents.b,
            c: leftComponents.a * rightComponents.c + leftComponents.c * rightComponents.d,
            d: leftComponents.b * rightComponents.c + leftComponents.d * rightComponents.d
        };
    }

    /**
     * 2×2行列の逆行列を求める（行列式が0に近い場合は微小値で代用）。
     * @param {object} components - 成分 { a, b, c, d }
     * @returns {object} 逆行列の成分 { a, b, c, d }
     */
    function invert2x2(components) {
        var determinant = components.a * components.d - components.b * components.c;
        if (Math.abs(determinant) < MATRIX_EPSILON) {
            determinant = (determinant < 0 ? -1 : 1) * MATRIX_EPSILON;
        }
        var inverseDeterminant = 1.0 / determinant;
        return {
            a: components.d * inverseDeterminant,
            b: -components.b * inverseDeterminant,
            c: -components.c * inverseDeterminant,
            d: components.a * inverseDeterminant
        };
    }

    /**
     * 成分から平行移動なしの Matrix を作る。
     * @param {object} components - 成分 { a, b, c, d }
     * @returns {object} Illustrator の Matrix
     */
    function buildMatrixFromComponents(components) {
        var matrix = new Matrix();
        matrix.mValueA = components.a;
        matrix.mValueB = components.b;
        matrix.mValueC = components.c;
        matrix.mValueD = components.d;
        matrix.mValueTX = 0;
        matrix.mValueTY = 0;
        return matrix;
    }

    /**
     * 行列をQR分解し、向き・スケール・シアーに分ける。
     * @param {object} itemMatrix - 対象の Matrix
     * @returns {object} { scaleX, scaleY, shear, axis1X, axis1Y, axis2X, axis2Y }
     */
    function decomposeMatrix(itemMatrix) {
        var matrixA = itemMatrix.mValueA;
        var matrixB = itemMatrix.mValueB;
        var matrixC = itemMatrix.mValueC;
        var matrixD = itemMatrix.mValueD;

        var scaleX = Math.sqrt(matrixA * matrixA + matrixB * matrixB);
        if (scaleX === 0) scaleX = MATRIX_EPSILON;

        var axis1X = matrixA / scaleX;
        var axis1Y = matrixB / scaleX;

        var projection = axis1X * matrixC + axis1Y * matrixD;
        var residualX = matrixC - projection * axis1X;
        var residualY = matrixD - projection * axis1Y;

        var scaleY = Math.sqrt(residualX * residualX + residualY * residualY);
        if (scaleY === 0) {
            scaleY = MATRIX_EPSILON;
            residualX = -axis1Y;
            residualY = axis1X;
        }

        return {
            scaleX: scaleX,
            scaleY: scaleY,
            shear: projection / scaleX,
            axis1X: axis1X,
            axis1Y: axis1Y,
            axis2X: residualX / scaleY,
            axis2Y: residualY / scaleY
        };
    }

    /**
     * 分解結果の一部を差し替えて、目標となる2×2成分を組み立てる。
     * @param {object} decomposed - decomposeMatrix の戻り値
     * @param {number} scaleX - 目標の水平スケール
     * @param {number} scaleY - 目標の垂直スケール
     * @param {number} shear - 目標のシアー量
     * @returns {object} 成分 { a, b, c, d }
     */
    function buildTargetComponents(decomposed, scaleX, scaleY, shear) {
        var shearTerm = shear * scaleX;
        return {
            a: decomposed.axis1X * scaleX,
            b: decomposed.axis1Y * scaleX,
            c: decomposed.axis1X * shearTerm + decomposed.axis2X * scaleY,
            d: decomposed.axis1Y * shearTerm + decomposed.axis2Y * scaleY
        };
    }

    /**
     * 現在の行列を目標成分にするための差分行列を作る。
     * @param {object} itemMatrix - 現在の Matrix
     * @param {object} targetComponents - 目標の成分 { a, b, c, d }
     * @returns {object} 差分の Matrix
     */
    function buildDeltaMatrix(itemMatrix, targetComponents) {
        var currentComponents = {
            a: itemMatrix.mValueA,
            b: itemMatrix.mValueB,
            c: itemMatrix.mValueC,
            d: itemMatrix.mValueD
        };
        return buildMatrixFromComponents(multiply2x2(invert2x2(currentComponents), targetComponents));
    }

    /**
     * 行列を分解し、目標成分との差分だけを適用する。
     * @param {object} pageItem - 対象のページアイテム
     * @param {function} buildTargetFn - 分解結果を受け取り目標成分を返す関数
     * @returns {void}
     */
    function applyDecomposedTransform(pageItem, buildTargetFn) {
        if (!hasMatrix(pageItem)) return;
        var itemMatrix = pageItem.matrix;
        pageItem.transform(buildDeltaMatrix(itemMatrix, buildTargetFn(decomposeMatrix(itemMatrix))));
    }

    /**
     * 縦横のスケールを大きいほうに揃えた目標成分を求める（整数％にスナップ）。
     * @param {object} decomposed - decomposeMatrix の戻り値
     * @returns {object} 成分 { a, b, c, d }
     */
    function getUniformScaleTarget(decomposed) {
        var uniformScale = Math.max(decomposed.scaleX, decomposed.scaleY);
        /* 整数％にスナップ（例：16.321% → 16%）/ snap to an integer percent */
        var uniformPercent = Math.round(uniformScale * 100);
        /* 0% に丸めるとオブジェクトが潰れるため最小1%を保証 / never collapse the item to 0% */
        if (uniformPercent < 1) uniformPercent = 1;
        uniformScale = uniformPercent / 100;
        return buildTargetComponents(decomposed, uniformScale, uniformScale, decomposed.shear);
    }

    // =========================================
    // 変形ヘルパー / Transform helpers
    // =========================================

    /**
     * バウンディングボックスをリセットする（位置は変えない）。
     * @param {object} pageItem - 対象のページアイテム
     * @returns {void}
     */
    function resetBoundingBox(pageItem) {
        try {
            app.selection = null;
            app.selection = [pageItem];
            app.executeMenuCommand("AI Reset Bounding Box");
        } catch (e) { }
    }

    /**
     * 変形前の中心位置に戻す。
     * @param {object} pageItem - 対象のページアイテム
     * @param {array} originalPosition - 変形前の position（左上座標）
     * @param {number} widthBefore - 変形前の幅
     * @param {number} heightBefore - 変形前の高さ
     * @returns {void}
     */
    function recenterToOriginalCenter(pageItem, originalPosition, widthBefore, heightBefore) {
        pageItem.position = [
            originalPosition[0] + widthBefore / 2 - pageItem.width / 2,
            originalPosition[1] - heightBefore / 2 + pageItem.height / 2
        ];
    }

    /**
     * 変形 → バウンディングボックスのリセット → 元の中心へ再配置、をまとめて行う。
     * @param {object} pageItem - 対象のページアイテム
     * @param {function} transformFn - 実際の変形処理
     * @returns {void}
     */
    function withBoundsResetAndRecenter(pageItem, transformFn) {
        var originalPosition = pageItem.position;
        var widthBefore = pageItem.width;
        var heightBefore = pageItem.height;

        transformFn();
        resetBoundingBox(pageItem);
        recenterToOriginalCenter(pageItem, originalPosition, widthBefore, heightBefore);
    }

    /**
     * ロック・非表示を一時解除して処理を実行し、元の状態に戻す。
     * @param {object} pageItem - 対象のページアイテム
     * @param {function} actionFn - 実行する処理
     * @returns {void}
     */
    function withUnlockedVisible(pageItem, actionFn) {
        var wasLocked = pageItem.locked;
        var wasHidden = pageItem.hidden;
        pageItem.locked = false;
        pageItem.hidden = false;
        try {
            actionFn();
        } finally {
            pageItem.locked = wasLocked;
            pageItem.hidden = wasHidden;
        }
    }

    /**
     * ロック・非表示を解除したうえで、中心基準・全オプション有効で変形を適用する。
     * @param {object} pageItem - 対象のページアイテム
     * @param {object} transformMatrix - 適用する Matrix
     * @returns {void}
     */
    function transformItemUnlocked(pageItem, transformMatrix) {
        withUnlockedVisible(pageItem, function () {
            pageItem.transform(transformMatrix, true, true, true, true, true, Transformation.CENTER);
        });
    }

    /**
     * 行列の成分から回転角を求める。
     * @param {number} matrixA - 行列の a 成分
     * @param {number} matrixB - 行列の b 成分
     * @param {number} angleSign - 符号（RasterItem は -1、それ以外は 1）
     * @returns {number} 回転角（度）
     */
    function getRotationAngleDeg(matrixA, matrixB, angleSign) {
        var angleDeg = Math.atan2(matrixB, matrixA) * 180 / Math.PI;
        return (angleSign < 0) ? -angleDeg : angleDeg;
    }

    /**
     * アイテムを指定角度だけ回転する。
     * @param {object} pageItem - 対象のページアイテム
     * @param {number} degrees - 回転角（度）
     * @returns {void}
     */
    function rotateItemBy(pageItem, degrees) {
        pageItem.transform(app.getRotationMatrix(degrees));
    }

    /**
     * テキストフレームを中心基準で回転する。
     * @param {object} textFrame - 対象のテキストフレーム
     * @param {number} degrees - 回転角（度）
     * @returns {void}
     */
    function rotateTextFrameBy(textFrame, degrees) {
        textFrame.rotate(degrees, true, true, true, true, Transformation.CENTER);
    }

    /**
     * 行列の符号規則に合わせて回転を打ち消す（配置画像・ラスター用）。
     * @param {object} pageItem - 対象のページアイテム
     * @param {number} angleSign - 符号（RasterItem は -1、それ以外は 1）
     * @returns {void}
     */
    function cancelRotationBySign(pageItem, angleSign) {
        if (!hasMatrix(pageItem)) return;
        var itemMatrix = pageItem.matrix;
        rotateItemBy(pageItem, getRotationAngleDeg(itemMatrix.mValueA, itemMatrix.mValueB, angleSign));
    }

    /**
     * 現在の角度の逆回転を掛けて 0° に戻す（テキスト用）。
     * @param {object} pageItem - 対象のページアイテム
     * @returns {void}
     */
    function cancelRotationToZero(pageItem) {
        if (!hasMatrix(pageItem)) return;
        var itemMatrix = pageItem.matrix;
        var angleDeg = Math.atan2(itemMatrix.mValueB, itemMatrix.mValueA) * 180 / Math.PI;
        if (pageItem.typename === 'TextFrame') {
            rotateTextFrameBy(pageItem, -angleDeg);
        } else {
            rotateItemBy(pageItem, -angleDeg);
        }
    }

    /**
     * シアーだけを取り除く。
     * @param {object} pageItem - 対象のページアイテム
     * @returns {void}
     */
    function removeShear(pageItem) {
        applyDecomposedTransform(pageItem, function (decomposed) {
            return buildTargetComponents(decomposed, decomposed.scaleX, decomposed.scaleY, 0);
        });
    }

    /**
     * シアーを取り除き、丸め誤差で残った微小シアーをもう一度取り除く。
     * @param {object} pageItem - 対象のページアイテム
     * @returns {void}
     */
    function removeShearWithRetry(pageItem) {
        removeShear(pageItem);
        if (!hasMatrix(pageItem)) return;
        if (Math.abs(decomposeMatrix(pageItem.matrix).shear) > SCALE_EPSILON) removeShear(pageItem);
    }

    /**
     * スケールを100%に正規化する（向きとシアーは保つ）。
     * @param {object} pageItem - 対象のページアイテム
     * @returns {void}
     */
    function normalizeScaleTo100(pageItem) {
        applyDecomposedTransform(pageItem, function (decomposed) {
            return buildTargetComponents(decomposed, 1, 1, decomposed.shear);
        });
    }

    /**
     * 縦横のスケールを大きいほうに揃える。
     * @param {object} pageItem - 対象のページアイテム
     * @returns {void}
     */
    function equalizeScaleToLarger(pageItem) {
        applyDecomposedTransform(pageItem, getUniformScaleTarget);
    }

    /**
     * 指定した倍率で等比拡大・縮小する。
     * @param {object} pageItem - 対象のページアイテム
     * @param {number} scalePercent - 倍率（100 = 100%）
     * @returns {void}
     */
    function applyUniformScalePercent(pageItem, scalePercent) {
        if (!(scalePercent > 0)) return;
        pageItem.resize(scalePercent, scalePercent, true, true, true, true, true, Transformation.CENTER);
    }

    // =========================================
    // 反転の解除 / Flip handling
    // =========================================

    /**
     * 左右反転しているかを判定する（回転・シアーなしが前提）。
     * @param {object} itemMatrix - 対象の Matrix
     * @returns {boolean} 左右反転していれば true
     */
    function isFlippedHorizontal(itemMatrix) {
        return itemMatrix.mValueA < 0;
    }

    /**
     * 上下反転しているかを判定する（回転・シアーなしが前提）。
     * 配置画像・ラスターは既定で mValueD が負のため、正のときを反転とみなす。
     * @param {object} itemMatrix - 対象の Matrix
     * @returns {boolean} 上下反転していれば true
     */
    function isFlippedVertical(itemMatrix) {
        return itemMatrix.mValueD > 0;
    }

    /**
     * 反転フラグに応じて反転を打ち消す。
     * @param {object} pageItem - 対象のページアイテム
     * @param {boolean} hasHorizontalFlip - 左右反転しているか
     * @param {boolean} hasVerticalFlip - 上下反転しているか
     * @returns {void}
     */
    function undoFlipByFlags(pageItem, hasHorizontalFlip, hasVerticalFlip) {
        if (!pageItem || (!hasHorizontalFlip && !hasVerticalFlip)) return;
        pageItem.transform(
            app.getScaleMatrix(hasHorizontalFlip ? -100 : 100, hasVerticalFlip ? -100 : 100),
            true, /* changePositions */
            true, /* changeFillPatterns */
            true, /* changeFillGradients */
            true, /* changeStrokePattern */
            true, /* changeLineWidths */
            Transformation.CENTER
        );
    }

    /**
     * 自身の行列から反転を判定して打ち消す。
     * @param {object} pageItem - 対象のページアイテム
     * @returns {void}
     */
    function undoItemFlip(pageItem) {
        if (!hasMatrix(pageItem)) return;
        var itemMatrix = pageItem.matrix;
        undoFlipByFlags(pageItem, isFlippedHorizontal(itemMatrix), isFlippedVertical(itemMatrix));
    }

    // =========================================
    // 軸へのスナップ / Axis snapping
    // =========================================

    /**
     * 角度を -180〜180 の範囲に正規化する。
     * @param {number} angleDeg - 角度（度）
     * @returns {number} 正規化した角度（度）
     */
    function normalizeTo180(angleDeg) {
        while (angleDeg > 180) angleDeg -= 360;
        while (angleDeg < -180) angleDeg += 360;
        return angleDeg;
    }

    /**
     * 角度を -90〜90 の範囲に畳み込む。
     * @param {number} angleDeg - 角度（度）
     * @returns {number} 畳み込んだ角度（度）
     */
    function clampTo90(angleDeg) {
        angleDeg = normalizeTo180(angleDeg);
        if (angleDeg > 90) angleDeg -= 180;
        if (angleDeg < -90) angleDeg += 180;
        return angleDeg;
    }

    /**
     * 最も近い軸（0°／90°）へ向かう最小の回転量を求める。
     * @param {number} angleDeg - 現在の角度（度）
     * @returns {number} 最小の回転量（度）
     */
    function getRotationToNearestAxis(angleDeg) {
        var toHorizontal = clampTo90(-angleDeg);
        var toVertical = clampTo90(90 - angleDeg);
        return (Math.abs(toHorizontal) <= Math.abs(toVertical)) ? toHorizontal : toVertical;
    }

    /**
     * パスの最初の2アンカーが作る辺の角度を求める。
     * @param {object} pathItem - 対象のパス
     * @returns {number} 水平からの角度（度）。2点が同一なら null
     */
    function getFirstEdgeAngleDeg(pathItem) {
        var anchorA = pathItem.pathPoints[0].anchor;
        var anchorB = pathItem.pathPoints[1].anchor;
        var dx = anchorB[0] - anchorA[0];
        var dy = anchorB[1] - anchorA[1];
        if (dx === 0 && dy === 0) return null;
        return Math.atan2(dy, dx) * 180 / Math.PI;
    }

    /**
     * わずかに傾いたパスを最も近い軸へスナップする（長方形・直線で共用）。
     * 許容範囲を外れた傾きは意図的とみなして何もしない。
     * @param {object} pathItem - 対象のパス
     * @returns {boolean} スナップしたら true
     */
    function snapPathToNearestAxis(pathItem) {
        var edgeAngleDeg = getFirstEdgeAngleDeg(pathItem);
        if (edgeAngleDeg === null) return false;

        var rotationDeg = getRotationToNearestAxis(edgeAngleDeg);
        var distanceToAxis = Math.abs(rotationDeg);
        if (distanceToAxis < AXIS_SNAP_MIN_DEG || distanceToAxis > AXIS_SNAP_MAX_DEG) return false;

        withBoundsResetAndRecenter(pathItem, function () {
            rotateItemBy(pathItem, rotationDeg);
        });
        return true;
    }

    // =========================================
    // 配置画像・ラスター / Placed & raster items
    // =========================================

    /**
     * 配置画像・ラスターの変形をリセットする（回転→反転→縦横比→スケール→シアーの順）。
     * @param {object} pageItem - 対象のページアイテム
     * @param {string} itemTypeName - typename（"PlacedItem" または "RasterItem"）
     * @param {object} resetOptions - ダイアログで選択したオプション
     * @returns {void}
     */
    function resetPlacedOrRasterTransforms(pageItem, itemTypeName, resetOptions) {
        var rotationSign = (itemTypeName === "RasterItem") ? -1 : 1;

        withBoundsResetAndRecenter(pageItem, function () {
            /* 1) 回転（向きを安定させるため最初に）/ rotation first */
            if (resetOptions.placedRotate) cancelRotationBySign(pageItem, rotationSign);

            /* 2) 反転（回転補正後に判定）/ flip, after the rotation is cancelled */
            if (resetOptions.placedFlip) undoItemFlip(pageItem);

            /* 3) 縦横比の等比化（絶対スケールの前に）/ equalize the aspect ratio before absolute scaling */
            if (resetOptions.placedAspectRatio) equalizeScaleToLarger(pageItem);

            /* 4) 絶対スケール（100%へ正規化してから指定%）/ normalize to 100%, then apply the requested percent */
            if (resetOptions.placedScale) {
                normalizeScaleTo100(pageItem);
                applyUniformScalePercent(pageItem, resetOptions.placedScalePercent);
            }

            /* 5) シアー除去は最後（上記で混入した微小シアーも取り除く）/ shear removal last */
            if (resetOptions.placedShear) removeShearWithRetry(pageItem);
        });
    }

    // =========================================
    // テキスト / Text frames
    // =========================================

    /**
     * テキストの水平比率・垂直比率を100%に戻す。
     * @param {object} textFrame - 対象のテキストフレーム
     * @returns {void}
     */
    function resetTextFrameScaleRatio(textFrame) {
        if (!textFrame.textRange) return;
        var textRange = textFrame.textRange;
        textRange.scaling = [1, 1];

        /* 丸め誤差が残った場合だけもう一度適用 / re-apply only when a rounding residual remains */
        var currentScaling = textRange.scaling;
        if (Math.abs(currentScaling[0] - 1) > SCALE_EPSILON || Math.abs(currentScaling[1] - 1) > SCALE_EPSILON) {
            textRange.scaling = [1, 1];
        }
    }

    /**
     * テキストの回転・シアー・比率をまとめてリセットする。
     * @param {object} textFrame - 対象のテキストフレーム
     * @param {boolean} doRotate - 回転をリセットするか
     * @param {boolean} doShear - シアーを除去するか
     * @param {boolean} doScaleRatio - 水平比率・垂直比率を戻すか
     * @returns {void}
     */
    function resetTextFrameTransforms(textFrame, doRotate, doShear, doScaleRatio) {
        withBoundsResetAndRecenter(textFrame, function () {
            /* 1) まず回転を0°へ（ポイント文字・エリア内文字の双方に有効）/ rotation first */
            if (doRotate) cancelRotationToZero(textFrame);
            /* 2) シアー除去 / shear removal */
            if (doShear) removeShearWithRetry(textFrame);
            /* 3) 比率は最後 / ratio last */
            if (doScaleRatio) resetTextFrameScaleRatio(textFrame);
        });
    }

    // =========================================
    // クリップグループ / Clipped groups
    // =========================================

    /**
     * バウンディングボックスから面積を求める。
     * @param {object} pageItem - 対象のページアイテム
     * @returns {number} 面積
     */
    function getBoundsArea(pageItem) {
        return Math.abs(pageItem.width * pageItem.height);
    }

    /**
     * 複合パスがマスクとして使われているかを判定する。
     * @param {object} compoundPath - 対象の CompoundPathItem
     * @returns {boolean} 子パスのいずれかがマスクなら true
     */
    function isClippingCompoundPath(compoundPath) {
        var childPaths = compoundPath.pathItems || [];
        for (var i = 0; i < childPaths.length; i++) {
            if (childPaths[i] && childPaths[i].clipping) return true;
        }
        return false;
    }

    /**
     * クリップグループから、代表となる配置画像とマスクパスを再帰的に探す。
     * どちらも面積が最大のものを採用する。
     * @param {object} container - 探索するグループ
     * @returns {object} { image, clipPath }（見つからなければ null）
     */
    function findLargestImageAndClipPath(container) {
        var largestImage = null;
        var largestImageArea = -1;
        var largestClipPath = null;
        var largestClipPathArea = -1;

        /**
         * 候補のほうが面積が大きければ採用する。
         * @param {object} candidate - 候補のページアイテム
         * @param {boolean} isClipPath - マスクパスとして扱うか
         * @returns {void}
         */
        function keepIfLarger(candidate, isClipPath) {
            if (!candidate) return;
            var candidateArea = getBoundsArea(candidate);
            if (isClipPath) {
                if (candidateArea > largestClipPathArea) {
                    largestClipPath = candidate;
                    largestClipPathArea = candidateArea;
                }
            } else if (candidateArea > largestImageArea) {
                largestImage = candidate;
                largestImageArea = candidateArea;
            }
        }

        var childItems = container.pageItems || [];
        for (var i = 0; i < childItems.length; i++) {
            var childItem = childItems[i];
            if (!childItem) continue;
            var typeName = childItem.typename;

            if (typeName === 'PlacedItem' || typeName === 'RasterItem') {
                keepIfLarger(childItem, false);
            } else if (typeName === 'PathItem' && childItem.clipping) {
                keepIfLarger(childItem, true);
            } else if (typeName === 'CompoundPathItem' && isClippingCompoundPath(childItem)) {
                /* 複合パス自身の面積を代用値として使う / use the compound's own area as a proxy */
                keepIfLarger(childItem, true);
            } else if (typeName === 'GroupItem') {
                var nestedResult = findLargestImageAndClipPath(childItem);
                keepIfLarger(nestedResult.image, false);
                keepIfLarger(nestedResult.clipPath, true);
            }
        }

        return { image: largestImage, clipPath: largestClipPath };
    }

    /**
     * マスクパスのうち、実際に変形を掛ける対象を決める。
     * 複合パスなら、マスク指定されている子パスを優先する。
     * @param {object} clipPathCandidate - 見つかったマスクパス
     * @returns {object} 変形対象。マスクパスがなければ null
     */
    function resolveClippingTransformTarget(clipPathCandidate) {
        if (!clipPathCandidate) return null;
        if (clipPathCandidate.typename !== 'CompoundPathItem') return clipPathCandidate;

        var childPaths = clipPathCandidate.pathItems || [];
        for (var i = 0; i < childPaths.length; i++) {
            if (childPaths[i] && childPaths[i].clipping) return childPaths[i];
        }
        /* マスク指定の子が見つからなければ複合パス自体を変形 / fall back to the compound itself */
        return clipPathCandidate;
    }

    /**
     * クリップグループの構成要素を一度だけ集める。
     * @param {object} groupItem - 対象のクリップグループ
     * @returns {object} { image, clipPath, clipTarget }
     */
    function collectClippedGroupParts(groupItem) {
        var clipParts = findLargestImageAndClipPath(groupItem);
        clipParts.clipTarget = resolveClippingTransformTarget(clipParts.clipPath);
        return clipParts;
    }

    /**
     * クリップグループの回転をリセットする。
     * 子同士の位置関係を保つため、グループごと回転する。
     * @param {object} groupItem - 対象のクリップグループ
     * @param {object} clipParts - collectClippedGroupParts の戻り値
     * @returns {boolean} 回転したら true
     */
    function resetClippedGroupRotation(groupItem, clipParts) {
        var representativeImage = clipParts.image;
        if (!representativeImage || !hasMatrix(representativeImage)) return false;

        var imageMatrix = representativeImage.matrix;
        var rotationSign = (representativeImage.typename === 'RasterItem') ? -1 : 1;
        var rotationDeg = getRotationAngleDeg(imageMatrix.mValueA, imageMatrix.mValueB, rotationSign);
        if (Math.abs(rotationDeg) <= ROTATION_EPSILON_DEG) return false;

        withBoundsResetAndRecenter(groupItem, function () {
            rotateItemBy(groupItem, rotationDeg);
        });
        return true;
    }

    /**
     * クリップグループの反転を、配置画像とマスクパスに同じだけ適用して打ち消す。
     * @param {object} clipParts - collectClippedGroupParts の戻り値
     * @returns {boolean} 反転を打ち消したら true
     */
    function undoClippedGroupFlip(clipParts) {
        var representativeImage = clipParts.image;
        if (!representativeImage || !hasMatrix(representativeImage)) return false;

        var imageMatrix = representativeImage.matrix;
        var hasHorizontalFlip = isFlippedHorizontal(imageMatrix);
        var hasVerticalFlip = isFlippedVertical(imageMatrix);
        if (!hasHorizontalFlip && !hasVerticalFlip) return false;

        withUnlockedVisible(representativeImage, function () {
            undoFlipByFlags(representativeImage, hasHorizontalFlip, hasVerticalFlip);
        });
        if (clipParts.clipTarget) {
            withUnlockedVisible(clipParts.clipTarget, function () {
                undoFlipByFlags(clipParts.clipTarget, hasHorizontalFlip, hasVerticalFlip);
            });
        }
        return true;
    }

    /**
     * クリップグループの縦横比を等比に戻す。
     * 配置画像から求めた差分を、マスクパスにも同じだけ適用する。
     * @param {object} clipParts - collectClippedGroupParts の戻り値
     * @returns {boolean} 等比化したら true
     */
    function resetClippedGroupAspectRatio(clipParts) {
        var representativeImage = clipParts.image;
        if (!representativeImage || !hasMatrix(representativeImage)) return false;

        var imageMatrix = representativeImage.matrix;
        var deltaMatrix = buildDeltaMatrix(imageMatrix, getUniformScaleTarget(decomposeMatrix(imageMatrix)));

        transformItemUnlocked(representativeImage, deltaMatrix);
        if (clipParts.clipTarget) transformItemUnlocked(clipParts.clipTarget, deltaMatrix);

        /* 丸め誤差で等比になりきらなかった場合の再調整 / re-equalize if a rounding residual remains */
        var decomposedAfter = decomposeMatrix(representativeImage.matrix);
        if (Math.abs(decomposedAfter.scaleX - decomposedAfter.scaleY) > SCALE_EPSILON) {
            withUnlockedVisible(representativeImage, function () {
                equalizeScaleToLarger(representativeImage);
            });
        }
        return true;
    }

    // =========================================
    // 種別ごとの振り分け / Dispatch by typename
    // =========================================

    /**
     * typename をキーにしたハンドラー一覧を作る。
     * 各ハンドラーは「リセット対象として処理したか」を返す。
     * @param {object} resetOptions - ダイアログで選択したオプション
     * @returns {object} typename → ハンドラー関数の対応表
     */
    function makeItemHandlers(resetOptions) {

        /**
         * テキストを処理する。
         * @param {object} textFrame - 対象のテキストフレーム
         * @returns {boolean} 処理したら true
         */
        function handleTextFrame(textFrame) {
            if (!resetOptions.textRotate && !resetOptions.textShear && !resetOptions.textScaleRatio) return false;
            resetTextFrameTransforms(textFrame, resetOptions.textRotate, resetOptions.textShear, resetOptions.textScaleRatio);
            return true;
        }

        /**
         * パス（長方形・直線）を処理する。
         * すでに正立している場合も「対象として処理済み」として扱う。
         * @param {object} pathItem - 対象のパス
         * @returns {boolean} 処理したら true
         */
        function handlePathItem(pathItem) {
            if (resetOptions.straightLineRotate && isStraightLinePath(pathItem)) {
                snapPathToNearestAxis(pathItem);
                return true;
            }
            if (resetOptions.rectangleRotate && isRectanglePath(pathItem)) {
                snapPathToNearestAxis(pathItem);
                return true;
            }
            return false;
        }

        /**
         * クリップグループを処理する。
         * @param {object} groupItem - 対象のグループ
         * @returns {boolean} 処理したら true
         */
        function handleClippedGroup(groupItem) {
            if (groupItem.clipped !== true) return false;

            /* 行列を読む前にバウンディングボックスを更新 / refresh the bounds before reading matrices */
            resetBoundingBox(groupItem);
            var clipParts = collectClippedGroupParts(groupItem);

            var didRotate = resetOptions.clipRotate && resetClippedGroupRotation(groupItem, clipParts);
            /* 反転判定の前に、回転後の状態を反映させる / let flip detection see the rotated state */
            if (didRotate && resetOptions.clipFlip) resetBoundingBox(groupItem);

            var didFlip = resetOptions.clipFlip && undoClippedGroupFlip(clipParts);
            var didAspectRatio = resetOptions.clipAspectRatio && resetClippedGroupAspectRatio(clipParts);

            /* 配置画像があれば無変更でも対象とみなす（誤アラート防止）/ an image means it was a valid target */
            return !!(didRotate || didFlip || didAspectRatio || clipParts.image);
        }

        /**
         * 配置画像・ラスターを処理する。
         * @param {object} pageItem - 対象のページアイテム
         * @param {string} itemTypeName - typename
         * @returns {boolean} 処理したら true
         */
        function handlePlacedOrRaster(pageItem, itemTypeName) {
            if (!resetOptions.placedRotate && !resetOptions.placedShear && !resetOptions.placedScale &&
                !resetOptions.placedAspectRatio && !resetOptions.placedFlip) return false;
            resetPlacedOrRasterTransforms(pageItem, itemTypeName, resetOptions);
            return true;
        }

        return {
            TextFrame: handleTextFrame,
            PathItem: handlePathItem,
            GroupItem: handleClippedGroup,
            PlacedItem: function (pageItem) {
                return handlePlacedOrRaster(pageItem, 'PlacedItem');
            },
            RasterItem: function (pageItem) {
                return handlePlacedOrRaster(pageItem, 'RasterItem');
            }
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * エントリーポイント。
     * @returns {void}
     */
    function main() {
        if (!app.documents.length) {
            alert(L(LABELS.alert.noDocument));
            return;
        }

        var currentDocument = app.activeDocument;
        var currentSelection = currentDocument.selection;
        if (!currentSelection || currentSelection.length === 0) {
            alert(L(LABELS.alert.selectFirst));
            return;
        }

        /* 処理中に app.selection を張り替えるため、開始時の選択を控える / snapshot the selection */
        var originalSelection = [];
        for (var i = 0; i < currentSelection.length; i++) originalSelection.push(currentSelection[i]);

        var resetOptions = showResetOptionsDialog(originalSelection);
        if (!resetOptions) return;

        var itemHandlers = makeItemHandlers(resetOptions);
        var processedCount = 0;

        for (var j = 0; j < originalSelection.length; j++) {
            var selectedItem = originalSelection[j];
            if (!selectedItem || !selectedItem.typename) continue;
            var itemHandler = itemHandlers[selectedItem.typename];
            if (itemHandler && itemHandler(selectedItem)) processedCount++;
        }

        /* 元の選択に戻す / restore the original selection */
        currentDocument.selection = originalSelection;

        if (processedCount === 0) alert(L(LABELS.alert.noTarget));
    }

    main();

})();
