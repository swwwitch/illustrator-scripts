#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

既存のアートボードの並び（行・列）を解析し、その規則を維持したまま新しいアートボードを挿入します。
追加方法・追加位置・追加数・間隔をダイアログで指定できます。

詳細は README を参照してください。

### Overview

Analyzes how the existing artboards are arranged in rows and columns and inserts new ones that follow the same pattern.
The method, position, count and spacing are all set in a dialog.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AddArtboardPlus";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Takeshi Umeda (noellabo)";     /* 作者 / author */
var SCRIPT_MODIFIED = "Masahiro Takano (@swwwitch)";  /* 改変 / modified by */
var SCRIPT_RELEASED = "2026-04-15";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddArtboardPlus.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddArtboardPlus.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/naf239a44b8ff"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/**
 * @author Takeshi Umeda (noellabo)
 * @discussion https://dtp-discourse.jp/t/illustrator/99
 */

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 追加したアートボード名に付ける接尾辞 / Suffix appended to the new artboard's name */
    var NAME_SUFFIX_BLANK = '_blank';
    var NAME_SUFFIX_DUPLICATE = '_copy';

    /* 「追加方法」の初期選択 / Initial selection of the Method panel */
    /* 'blank' = 空のアートボード / 'duplicate' = 現在のアートボードを複製 */
    /* 'blank' = blank artboard, 'duplicate' = duplicate the current artboard */
    var DEFAULT_ADD_METHOD = 'blank';

    // =========================================
    // ローカライズ / Localization
    // =========================================

    function getCurrentLang() {
        return ($.locale.indexOf('ja') === 0) ? 'ja' : 'en';
    }
    var uiLang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: 'アートボードを追加', en: 'Add Artboards' }
        },
        panel: {
            addMethod: { ja: '追加方法', en: 'Method' },
            insertPosition: { ja: '追加位置', en: 'Insert Position' },
            spacing: { ja: '間隔', en: 'Spacing' }
        },
        radio: {
            blank: { ja: '空のアートボード', en: 'Blank Artboard' },
            duplicate: { ja: '現在のアートボードを複製', en: 'Duplicate Current Artboard' },
            insertAfterCurrent: { ja: '現在のアートボードの次', en: 'After Current Artboard' },
            insertAtEnd: { ja: '末尾', en: 'After Last Artboard' },
            scopeAddedOnly: { ja: '追加分にのみ適用', en: 'Apply to Added Only' },
            scopeAll: { ja: 'すべてに適用', en: 'Apply to All' },
            directionRight: { ja: '右', en: 'Right' },
            directionDown: { ja: '下', en: 'Down' }
        },
        fieldLabel: {
            addCount: { ja: '追加数', en: 'Count' },
            direction: { ja: '方向', en: 'Direction' }
        },
        button: {
            cancel: { ja: 'キャンセル', en: 'Cancel' },
            ok: { ja: 'OK', en: 'OK' }
        },
        tooltip: {
            blank: {
                ja: '挿入位置の直前のアートボードと同じサイズで追加します（B）',
                en: 'Adds artboards the same size as the one before the insert position (B)'
            },
            duplicate: {
                ja: 'アクティブなアートボードをアートワークごと複製します。追加数は1枚です（D）',
                en: 'Duplicates the active artboard with its artwork. Always adds one (D)'
            },
            addCount: { ja: '↑↓で±1、Shift＋↑↓で±10', en: 'Up/Down: ±1, Shift+Up/Down: ±10' },
            insertAfterCurrent: { ja: 'アクティブなアートボードの次に挿入します（N）', en: 'Inserts after the active artboard (N)' },
            insertAtEnd: { ja: '最後のアートボードの次に追加します（E）', en: 'Adds after the last artboard (E)' },
            directionRight: {
                ja: '既存の並びが横並びなら、行の折り返しを引き継ぎます。縦並びのときは直前のアートボードの右に置きます',
                en: 'Follows the row wrapping when the existing artboards run horizontally; otherwise places them to the right of the previous artboard'
            },
            directionDown: {
                ja: '既存の並びが縦並びなら、列の折り返しを引き継ぎます。横並びのときは直前のアートボードの下に置きます',
                en: 'Follows the column wrapping when the existing artboards run vertically; otherwise places them below the previous artboard'
            },
            scopeAddedOnly: {
                ja: '既存のアートボードの間隔は変えず、追加したアートボードの前後だけを指定の間隔にします',
                en: 'Keeps the existing gaps and uses this spacing only around the added artboards'
            },
            scopeAll: {
                ja: '既存のアートボードもアートワークごと指定の間隔で並べ直します（方向が既存の並びと同じときのみ）',
                en: 'Also rearranges the existing artboards and their artwork with this spacing (only when the direction matches the existing layout)'
            },
            spacing: {
                ja: '初期値は既存の並びから推定した値です（1枚のときは［アートボードを再配置］の設定値）。\n↑↓で±1、Shift＋↑↓で±10、Option＋↑↓で±0.1',
                en: 'Initially estimated from the existing layout (the Rearrange Artboards setting when there is only one artboard).\nUp/Down: ±1, Shift+Up/Down: ±10, Option+Up/Down: ±0.1'
            }
        },
        alert: {
            noDocument: {
                ja: 'ドキュメントがないため実行できません。',
                en: 'Cannot run because no document is open.'
            },
            exception: {
                ja: 'エラーが発生したため、処理を実行できませんでした。\nエラー内容：',
                en: 'Processing could not be completed because an error occurred.\nError:'
            },
            artboardLimit: {
                ja: '追加すると、アートボードの上限（{max}）を超えます。',
                en: 'Adding these would exceed the artboard limit ({max}).'
            },
            noSpace: {
                ja: 'カンバスに、アートボードを追加する十分なスペースがありません。',
                en: 'There is not enough room on the canvas to add the artboards.'
            }
        }
    };

    /* ラベル（ja/en のリーフ）を現在の言語に解決 / Resolve a label leaf (ja/en) to the current language */
    function getLabel(labelSet) {
        return (labelSet && labelSet[uiLang]) || '';
    }

    /* 項目名にコロンを付ける（日本語は全角、英語は半角）/ Append a colon to a field label (full-width in Japanese) */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === 'ja' ? '：' : ': ');
    }

    // =========================================
    // 単位と座標 / Units & coordinates
    // =========================================

    // 座標が同じとみなす許容値（pt）。吸着させた辺でも 1e-12 程度ずれるため生比較しない
    // Tolerance (pt) for treating coordinates as equal; snapped edges still differ by ~1e-12
    var COORDINATE_TOLERANCE_PT = 0.001;

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コードに対応する表示ラベルと、1単位あたりのポイント数
       Unit code -> display label and points per unit */
    var UNITS = [
        { label: "in",    pointsPerUnit: 72 },                /* 0 */
        { label: "mm",    pointsPerUnit: 72 / 25.4 },         /* 1 */
        { label: "pt",    pointsPerUnit: 1 },                 /* 2 */
        { label: "pica",  pointsPerUnit: 12 },                /* 3 */
        { label: "cm",    pointsPerUnit: 72 / 2.54 },         /* 4 */
        { label: "Q",     pointsPerUnit: 72 / 25.4 * 0.25 },  /* 5 */
        { label: "px",    pointsPerUnit: 1 },                 /* 6 */
        { label: "ft/in", pointsPerUnit: 72 * 12 },           /* 7 */
        { label: "m",     pointsPerUnit: 72 / 25.4 * 1000 },  /* 8 */
        { label: "yd",    pointsPerUnit: 72 * 36 },           /* 9 */
        { label: "ft",    pointsPerUnit: 72 * 12 }            /* 10 */
    ];

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitCode = app.preferences.getIntegerPreference(prefKey || "rulerType");
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        return { code: unitCode, label: unit.label, pointsPerUnit: unit.pointsPerUnit };
    }

    // 表示用に小数第2位までに丸める / Round to 2 decimals for display
    function formatSpacingForDisplay(valueInUnit) {
        return String(Math.round(valueInUnit * 100) / 100);
    }

    function isSameCoordinate(coordA, coordB) {
        return Math.abs(coordA - coordB) <= COORDINATE_TOLERANCE_PT;
    }

    // 軸方向のアートボードサイズ（0=幅 / 1=高さ）/ Artboard size along an axis (0=width, 1=height)
    function getArtboardSizeOnAxis(artboardRect, axisIndex) {
        return (axisIndex === 0)
            ? (artboardRect[2] - artboardRect[0])           // 幅 / width
            : Math.abs(artboardRect[3] - artboardRect[1]);  // 高さ / height
    }

    // 隣り合う2枚から並びの軸を判定（左端が同じ＝縦並び→1 / 違う＝横並び→0）
    // Detect the layout axis from two neighbors (same left edge ⇒ vertical → 1, otherwise horizontal → 0)
    function detectLayoutAxisIndex(firstRect, secondRect) {
        return isSameCoordinate(firstRect[0], secondRect[0]) ? 1 : 0;
    }

    // 既存アートボードの並びから現在の間隔（pt）を推定
    // 2枚未満の場合は fallbackSpacingPt（環境設定値）を返す
    // Estimate the current spacing (pt) from the existing artboard arrangement
    // Returns fallbackSpacingPt (the preference value) when there are fewer than 2 artboards
    function estimateArtboardSpacingPt(fallbackSpacingPt) {
        var artboards = doc.artboards;
        if (artboards.length < 2) return fallbackSpacingPt;

        var firstRect = artboards[0].artboardRect;
        var secondRect = artboards[1].artboardRect;
        var layoutAxisIndex = detectLayoutAxisIndex(firstRect, secondRect);
        var artboardPitch = Math.abs(secondRect[layoutAxisIndex] - firstRect[layoutAxisIndex]);
        return Math.max(0, artboardPitch - getArtboardSizeOnAxis(firstRect, layoutAxisIndex));
    }

    // =========================================
    // UIレイアウトの共通設定 / Shared UI layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 8;                  /* パネル内の要素間隔 / panel spacing */

    /* ウィンドウの共通設定 / Apply shared window layout */
    function setupWindow(win, spacing) {
        win.orientation = 'column';
        win.alignChildren = 'fill';
        win.margins = WINDOW_MARGINS;
        win.spacing = (typeof spacing === 'number') ? spacing : WINDOW_SPACING;
    }

    /* パネルの共通設定 / Apply shared panel layout */
    function setupPanel(panel, spacing) {
        panel.orientation = 'column';
        panel.alignChildren = ['fill', 'top'];
        panel.alignment = 'fill';
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === 'number') ? spacing : PANEL_SPACING;
    }

    /* 行グループの共通設定（ボタン列など）/ Apply a horizontal row group */
    /* パネル幅いっぱいに広げず、既定では左寄せ / Doesn't stretch to the panel width; left-aligned by default */
    function setupRow(group, alignment, spacing) {
        group.orientation = 'row';
        group.alignment = alignment || 'left';
        group.spacing = (typeof spacing === 'number') ? spacing : PANEL_SPACING;
    }

    /* ↑↓キーで値を増減（Shift=±10で10の倍数にスナップ / Option=±0.1 / 通常=±1）。負値は0でクランプ */
    /* Arrow keys adjust the value (Shift = ±10 snapped to multiples of 10, Option = ±0.1, otherwise ±1); clamps at 0 */
    function changeValueByArrowKey(editText) {
        editText.addEventListener('keydown', function (event) {
            // ↑↓以外のキーは素通し（入力中の値を丸め直さない）/ Ignore keys other than Up/Down so typing isn't rounded
            if (event.keyName != 'Up' && event.keyName != 'Down') return;
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                // Shiftキー押下時は10の倍数にスナップ / Snap to multiples of 10 with Shift
                if (event.keyName == 'Up') {
                    value = Math.ceil((value + 1) / delta) * delta;
                    event.preventDefault();
                } else if (event.keyName == 'Down') {
                    value = Math.floor((value - 1) / delta) * delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                // Optionキー押下時は0.1単位で増減 / Step by 0.1 with Option
                if (event.keyName == 'Up') {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == 'Down') {
                    value -= delta;
                    event.preventDefault();
                }
            } else {
                delta = 1;
                if (event.keyName == 'Up') {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == 'Down') {
                    value -= delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            }

            if (keyboard.altKey) {
                value = Math.round(value * 10) / 10; // 小数第1位までに丸め / Round to 1 decimal
            } else {
                value = Math.round(value); // 整数に丸め / Round to integer
            }

            editText.text = value;
        });
    }

    // =========================================
    // メイン処理 / Main process
    // =========================================

    if (app.documents.length == 0) {
        alert(getLabel(LABELS.alert.noDocument));
        return;
    }
    var doc = app.activeDocument;

    try {
        var insertSettings = showInsertDialog();
        if (insertSettings) insertArtboardsWithShift(insertSettings);
    } catch (e) {
        alert(getLabel(LABELS.alert.exception) + e);
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    // ダイアログを表示し、OK なら挿入の設定を返す（キャンセルは null）
    // Show the dialog and return the insert settings on OK (null on cancel)
    function showInsertDialog() {
        var dialog = new Window('dialog', getLabel(LABELS.dialog.title) + ' ' + SCRIPT_VERSION);
        setupWindow(dialog);

        // 追加方法パネル / Add method panel
        var addMethodPanel = dialog.add('panel', undefined, getLabel(LABELS.panel.addMethod));
        setupPanel(addMethodPanel, 6);
        var blankArtboardRadio = addMethodPanel.add('radiobutton', undefined, getLabel(LABELS.radio.blank));
        blankArtboardRadio.helpTip = getLabel(LABELS.tooltip.blank);
        var duplicateArtboardRadio = addMethodPanel.add('radiobutton', undefined, getLabel(LABELS.radio.duplicate));
        duplicateArtboardRadio.helpTip = getLabel(LABELS.tooltip.duplicate);
        // 初期選択はユーザー設定に従う（'duplicate' 以外は「空のアートボード」）
        // The initial selection follows the user setting (anything but 'duplicate' means blank)
        duplicateArtboardRadio.value = (DEFAULT_ADD_METHOD === 'duplicate');
        blankArtboardRadio.value = !duplicateArtboardRadio.value;

        // 追加数 / Count
        var addCountGroup = addMethodPanel.add('group');
        setupRow(addCountGroup);
        var addCountLabel = addCountGroup.add('statictext', undefined, labelText(LABELS.fieldLabel.addCount));
        var addCountInput = addCountGroup.add('edittext', undefined, '1');
        addCountInput.characters = 4;
        addCountInput.helpTip = getLabel(LABELS.tooltip.addCount);
        changeValueByArrowKey(addCountInput);

        // 追加位置パネル / Insert position panel
        var insertPositionPanel = dialog.add('panel', undefined, getLabel(LABELS.panel.insertPosition));
        setupPanel(insertPositionPanel, 6);
        var insertAfterCurrentRadio = insertPositionPanel.add('radiobutton', undefined, getLabel(LABELS.radio.insertAfterCurrent));
        insertAfterCurrentRadio.helpTip = getLabel(LABELS.tooltip.insertAfterCurrent);
        var insertAtEndRadio = insertPositionPanel.add('radiobutton', undefined, getLabel(LABELS.radio.insertAtEnd));
        insertAtEndRadio.helpTip = getLabel(LABELS.tooltip.insertAtEnd);

        // 追加方向（右＝横並び / 下＝縦並び）/ Add direction (right = horizontal, down = vertical)
        var directionGroup = insertPositionPanel.add('group');
        setupRow(directionGroup);
        directionGroup.add('statictext', undefined, labelText(LABELS.fieldLabel.direction));
        var directionRightRadio = directionGroup.add('radiobutton', undefined, getLabel(LABELS.radio.directionRight));
        directionRightRadio.helpTip = getLabel(LABELS.tooltip.directionRight);
        var directionDownRadio = directionGroup.add('radiobutton', undefined, getLabel(LABELS.radio.directionDown));
        directionDownRadio.helpTip = getLabel(LABELS.tooltip.directionDown);
        directionRightRadio.value = true;

        // 追加方法に応じて UI を更新 / Update the UI according to the add method:
        //  - 追加数は「空のアートボード」のときのみ有効 / Count is enabled only for blank artboards
        //  - 追加位置の既定は 複製→現在の次 / 空→末尾 / Default position: duplicate → next, blank → end
        function syncUIWithAddMethod() {
            var isBlank = blankArtboardRadio.value;
            addCountLabel.enabled = isBlank;
            addCountInput.enabled = isBlank;
            insertAfterCurrentRadio.value = !isBlank;
            insertAtEndRadio.value = isBlank;
        }
        duplicateArtboardRadio.onClick = syncUIWithAddMethod;
        blankArtboardRadio.onClick = syncUIWithAddMethod;
        syncUIWithAddMethod();

        // 間隔パネル / Spacing panel
        // 適用範囲（追加分のみ / すべて）と、間隔の数値を指定
        // Choose the scope (added artboards only / all) and the spacing value
        var spacingPanel = dialog.add('panel', undefined, getLabel(LABELS.panel.spacing));
        setupPanel(spacingPanel, 6);
        var spacingScopeAddedOnlyRadio = spacingPanel.add('radiobutton', undefined, getLabel(LABELS.radio.scopeAddedOnly));
        spacingScopeAddedOnlyRadio.helpTip = getLabel(LABELS.tooltip.scopeAddedOnly);
        var spacingScopeAllRadio = spacingPanel.add('radiobutton', undefined, getLabel(LABELS.radio.scopeAll));
        spacingScopeAllRadio.helpTip = getLabel(LABELS.tooltip.scopeAll);
        spacingScopeAddedOnlyRadio.value = true;

        // 間隔入力 / Spacing input
        // 2枚以上は既存の並びから推定した間隔、1枚以下は「アートボードを再配置」の間隔（pt）を
        // 定規単位に変換して初期表示し、編集可能にする
        // For 2+ artboards use the spacing inferred from the existing layout, otherwise the
        // Rearrange Artboards spacing (pt); shown in the ruler unit and editable
        var rulerUnitInfo = getUnitInfo();
        var spacingPreferencePt = app.preferences.getRealPreference('plugin/ArtboardRearrange/ArtboardSpacing');
        var estimatedSpacingPt = estimateArtboardSpacingPt(spacingPreferencePt);
        var initialSpacingText = formatSpacingForDisplay(estimatedSpacingPt / rulerUnitInfo.pointsPerUnit);

        // 項目名はパネル名「間隔」と重なるので付けない / No field label; the panel title already says Spacing
        var spacingGroup = spacingPanel.add('group');
        setupRow(spacingGroup);
        var spacingInput = spacingGroup.add('edittext', undefined, initialSpacingText);
        spacingInput.characters = 4;
        spacingInput.helpTip = getLabel(LABELS.tooltip.spacing);
        changeValueByArrowKey(spacingInput);
        spacingGroup.add('statictext', undefined, rulerUnitInfo.label);

        addRadioShortcutKeyHandler(dialog, blankArtboardRadio, duplicateArtboardRadio, insertAfterCurrentRadio, insertAtEndRadio, syncUIWithAddMethod);

        // ボタン列はダイアログ幅いっぱいに広げず中央に置く / Center the button row instead of stretching it
        var btnRowGroup = dialog.add('group');
        setupRow(btnRowGroup, 'center');
        var btnCancel = btnRowGroup.add('button', undefined, getLabel(LABELS.button.cancel), { name: 'cancel' });
        var btnOK = btnRowGroup.add('button', undefined, getLabel(LABELS.button.ok), { name: 'ok' });

        if (dialog.show() != 1) return null;

        // 手動で初期値から変更されたかどうか（空欄・不正値は未変更扱い）。変更時のみ入力値を pt に戻して使い、
        // 未変更なら丸めた表示値を経由せず推定値（pt）をそのまま使う。負値は 0 にする
        // Whether the value was manually changed (blank/invalid counts as unchanged). Only a changed value
        // is converted back to pt; otherwise use the estimated pt value, not the rounded display. Negatives become 0
        var spacingInputValue = parseFloat(spacingInput.text);
        var useManualSpacing = !isNaN(spacingInputValue) && (spacingInput.text !== initialSpacingText);

        // 追加数（1以上の整数）。複製モードでは常に1枚 / Number to add (integer ≥ 1); duplicate mode always adds 1
        var isDuplicateMode = duplicateArtboardRadio.value;
        var addCount = parseInt(addCountInput.text, 10);
        if (isDuplicateMode || isNaN(addCount) || addCount < 1) addCount = 1;

        return {
            isDuplicateMode: isDuplicateMode,
            insertAfterCurrent: insertAfterCurrentRadio.value,
            addCount: addCount,
            // 追加方向：0=右（横並び）/ 1=下（縦並び）/ Add direction: 0 = right (horizontal), 1 = down (vertical)
            directionAxisIndex: directionDownRadio.value ? 1 : 0,
            spacingPt: useManualSpacing ? Math.max(0, spacingInputValue * rulerUnitInfo.pointsPerUnit) : estimatedSpacingPt,
            estimatedSpacingPt: estimatedSpacingPt,
            useManualSpacing: useManualSpacing,
            // 間隔の適用範囲：true=すべてのアートボード / false=追加分のみ
            // Spacing scope: true = all artboards, false = added artboards only
            applySpacingToAll: spacingScopeAllRadio.value
        };
    }

    // キー入力でラジオボタンを切り替え（B=空 / D=複製 / N=現在の次 / E=末尾）
    // Switch radio buttons with keys (B = blank, D = duplicate, N = after current, E = at end)
    function addRadioShortcutKeyHandler(targetDialog, blankRadio, duplicateRadio, afterCurrentRadio, atEndRadio, onAddMethodChange) {
        targetDialog.addEventListener('keydown', function (event) {
            var keyName = String(event.keyName || '').toUpperCase();

            if (keyName === 'B' || keyName === 'D') {
                duplicateRadio.value = (keyName === 'D');
                blankRadio.value = !duplicateRadio.value;
                onAddMethodChange();
            } else if (keyName === 'N' || keyName === 'E') {
                atEndRadio.value = (keyName === 'E');
                afterCurrentRadio.value = !atEndRadio.value;
            } else {
                return;
            }
            event.preventDefault();
        });
    }

    // =========================================
    // グリッドの解析 / Grid analysis
    // =========================================

    /*
    最大キャンバス範囲を取得 / Get the largest canvas bounds
    Original idea by OMOTI
    https://forums.adobe.com/thread/2459293
    */
    function getLargestCanvasBounds() {
        var MAX_CANVAS_SIZE = 16383;
        var tempLayer = doc.layers.add();
        var tempText = tempLayer.textFrames.add();
        var canvasLeft = tempText.matrix.mValueTX;
        var canvasTop = tempText.matrix.mValueTY;
        tempLayer.remove();

        return new Rect(
            canvasLeft,
            canvasTop,
            canvasLeft + MAX_CANVAS_SIZE,
            canvasTop - MAX_CANVAS_SIZE
        );
    }

    // 先頭のアートボードを原点に、グリッド1セルあたりの移動量・列数・行数を求める。
    // 既存の並びが指定方向と一致するときだけ、ピッチと列数を引き継ぐ（canInheritLayout）。
    // 一致しない・1枚のみのときはグリッドを使わず、挿入位置の直前を基準に並べる → getAnchoredPosition
    // Build the grid anchored on the first artboard: step per cell, column count and row count.
    // The pitch and column count are inherited only when the existing layout matches the chosen
    // direction (canInheritLayout); otherwise new artboards are placed relative to the artboard
    // just before the insert point → getAnchoredPosition
    function analyzeArtboardGrid(artboards, insertSettings) {
        var originRect = artboards[0].artboardRect;
        var primaryAxis = insertSettings.directionAxisIndex;
        var secondaryAxis = 1 - primaryAxis;
        var spacingPt = insertSettings.spacingPt;

        // 既定の移動量：幅＋間隔、高さは上→下で Y が減るので負方向
        // Default step: width + spacing; the height step is negative because Y decreases downward
        var gridStep = [
            getArtboardSizeOnAxis(originRect, 0) + spacingPt,
            -(getArtboardSizeOnAxis(originRect, 1) + spacingPt)
        ];
        var columnCount = 0;
        var canInheritLayout = (artboards.length >= 2) &&
            (detectLayoutAxisIndex(originRect, artboards[1].artboardRect) === primaryAxis);

        if (canInheritLayout) {
            gridStep[primaryAxis] = artboards[1].artboardRect[primaryAxis] - originRect[primaryAxis];

            // 副軸の座標が先頭と変わる最初のアートボードが次の行の先頭。そのインデックスが列数
            // The first artboard whose secondary coordinate differs starts the next row; its index is the column count
            for (var i = 2; i < artboards.length; i++) {
                var scannedRect = artboards[i].artboardRect;
                if (!isSameCoordinate(originRect[secondaryAxis], scannedRect[secondaryAxis])) {
                    gridStep[secondaryAxis] = scannedRect[secondaryAxis] - originRect[secondaryAxis];
                    columnCount = i;
                    break;
                }
            }

            // 「すべてのアートボード」かつ手動間隔のときだけ、向きを保ったまま移動量を入力値で作り直す
            // （追加分のみのときは既存のピッチを保持）
            // Only when scope = all AND the spacing was changed manually, rebuild the steps from the input
            // while keeping their direction (scope = added-only keeps the existing pitch)
            if (insertSettings.useManualSpacing && insertSettings.applySpacingToAll) {
                for (var axisIndex = 0; axisIndex < 2; axisIndex++) {
                    var stepSign = (gridStep[axisIndex] < 0) ? -1 : 1;
                    gridStep[axisIndex] = stepSign * (getArtboardSizeOnAxis(originRect, axisIndex) + spacingPt);
                }
            }
        }

        var canvasRect = getLargestCanvasBounds();
        var gridCellRect = [
            originRect[0] + Math.abs(gridStep[0]),
            originRect[1] - Math.abs(gridStep[1]),
            originRect[0],
            originRect[1]
        ];

        // 進む向きの側のキャンバス端までに入るセル数 / Number of cells that fit up to the canvas edge in the step direction
        function countCellsToCanvasEdge(targetAxis) {
            var edgeIndex = (targetAxis ^ +(gridStep[targetAxis] < 0)) ? targetAxis : targetAxis + 2;
            return Math.abs(Math.floor((canvasRect[edgeIndex] - gridCellRect[edgeIndex]) / gridStep[targetAxis]));
        }

        return {
            originRect: originRect,
            primaryAxis: primaryAxis,
            secondaryAxis: secondaryAxis,
            primarySign: (gridStep[primaryAxis] < 0) ? -1 : 1,
            gridStep: gridStep,
            columnCount: columnCount || countCellsToCanvasEdge(primaryAxis),
            rowCount: countCellsToCanvasEdge(secondaryAxis),
            canInheritLayout: canInheritLayout
        };
    }

    // グリッド上のインデックスの位置を計算 / Calculate the position of a grid index
    function getGridCellPosition(artboardGrid, gridIndex) {
        var cellOffset = [];
        cellOffset[artboardGrid.primaryAxis] = (gridIndex % artboardGrid.columnCount) * artboardGrid.gridStep[artboardGrid.primaryAxis];
        cellOffset[artboardGrid.secondaryAxis] = Math.floor(gridIndex / artboardGrid.columnCount) * artboardGrid.gridStep[artboardGrid.secondaryAxis];
        return [artboardGrid.originRect[0] + cellOffset[0], artboardGrid.originRect[1] + cellOffset[1]];
    }

    // グリッドを引き継げないとき（既存の並びが指定方向と違う／1枚のみ）の配置。
    // 先頭ではなく挿入位置の直前のアートボードを基準に、指定方向へ (addedOrder+1) 枚目を置く。
    // グリッドのインデックスをそのまま歩数に使うと、既存の並びと軸が違う場合に
    // 実際の枚数分だけ離れた位置へ飛んでしまうため、こちらで実位置から積み上げる。
    // Placement used when the grid can't be inherited (existing layout uses the other axis,
    // or there is only one artboard). Anchors on the artboard just before the insert point
    // instead of the first one, and steps along the chosen direction. Reusing the grid index
    // as a step count would fling the new artboard away by the whole artboard count when the
    // existing layout runs along the other axis, so build up from the real position instead.
    function getAnchoredPosition(insertPlan, referenceRect, addedOrder) {
        var primaryAxis = insertPlan.artboardGrid.primaryAxis;
        var spacingPt = insertPlan.spacingPt;
        var anchorRect = insertPlan.artboards[insertPlan.insertIndex - 1].artboardRect;
        // 1歩目はアンカーのサイズ、2枚目以降は新規アートボードのサイズで進む
        // The first step uses the anchor's size; later ones use the new artboards' size
        var primaryAdvance = getArtboardSizeOnAxis(anchorRect, primaryAxis) + spacingPt
            + addedOrder * (getArtboardSizeOnAxis(referenceRect, primaryAxis) + spacingPt);
        var anchoredPosition = [anchorRect[0], anchorRect[1]];
        anchoredPosition[primaryAxis] += insertPlan.artboardGrid.primarySign * primaryAdvance;
        return anchoredPosition;
    }

    // =========================================
    // 挿入処理 / Insertion
    // =========================================

    function insertArtboardsWithShift(insertSettings) {
        var insertPlan = buildInsertPlan(insertSettings);
        if (!insertPlan) return;

        // 既存アートボードを最終位置へ移動して挿入スペースを空ける。
        // グリッドを引き継げないときは既存の並びを動かさない（別軸のグリッドへ流し込むと
        // 先頭は横並び・後続は縦並びのように崩れるため）
        // Move existing artboards to their final positions to free the insert space.
        // When the grid can't be inherited, leave the existing artboards alone — re-flowing them
        // onto a grid that runs along the other axis would break the arrangement apart.
        if (insertPlan.artboardGrid.canInheritLayout) {
            relayoutExistingArtboards(insertPlan);
        }
        addNewArtboards(insertPlan);
        reorderAppendedArtboards(insertPlan);

        // 最初に追加したアートボードを選択・アクティブにする
        // Select and activate the first newly added artboard
        doc.selection = null;
        insertPlan.artboards.setActiveArtboardIndex(insertPlan.insertIndex);
        app.redraw();
    }

    // 挿入位置・基準アートボード・グリッド・移動量をまとめる。作成できないときは警告して null
    // Gather the insert index, reference artboard, grid and offsets; alerts and returns null when it can't proceed
    function buildInsertPlan(insertSettings) {
        var artboards = doc.artboards;
        var originalCount = artboards.length;
        var addCount = insertSettings.addCount;

        var maxArtboardCount = (parseFloat(app.version) >= 22) ? 1000 : 100;
        if (originalCount + addCount > maxArtboardCount) {
            alert(getLabel(LABELS.alert.artboardLimit).replace('{max}', maxArtboardCount));
            return null;
        }

        var artboardGrid = analyzeArtboardGrid(artboards, insertSettings);
        if (originalCount + addCount > artboardGrid.columnCount * artboardGrid.rowCount) {
            alert(getLabel(LABELS.alert.noSpace));
            return null;
        }

        var activeArtboardIndex = artboards.getActiveArtboardIndex();
        var insertIndex = insertSettings.insertAfterCurrent ? activeArtboardIndex + 1 : originalCount;

        // 新規アートボードのサイズと名前の基準。複製は「現在のアートボード」、
        // 空のアートボードは挿入位置の直前を基準にする（サイズがまちまちのドキュメントで
        // 末尾に追加したのに1枚目のサイズや名前になる、といったズレを避けるため）
        // Reference artboard for the new artboards' size and name: the current artboard for
        // duplicates, otherwise the artboard just before the insert point. Keeps appending at the
        // end from producing an artboard sized or named like the first one in a mixed-size document.
        var referenceIndex = insertSettings.isDuplicateMode ? activeArtboardIndex : (insertIndex - 1);

        // 追加分のみ：新規を (指定間隔 − 既存の間隔) ぶん主軸方向へずらす。既存の間隔はダイアログの推定値と同じ。
        // 後続は新規の前後どちらの隙間も指定間隔になるよう (追加数 + 1) 回分ずらす。
        // すべて：手動間隔のとき先頭も含めて新グリッドへ再配置（reflowFromHead）
        // Added-only: nudge the new artboards by (spacing − existing gap), where the existing gap is the
        // dialog's estimate; the tail moves (addCount + 1) times that so both gaps around the new ones match.
        // All: with manual spacing, re-flow everything including the head onto the new grid (reflowFromHead)
        var addedOnlyStep = (insertSettings.useManualSpacing && !insertSettings.applySpacingToAll)
            ? artboardGrid.primarySign * (insertSettings.spacingPt - insertSettings.estimatedSpacingPt)
            : 0;

        return {
            artboards: artboards,
            originalCount: originalCount,
            addCount: addCount,
            insertIndex: insertIndex,
            referenceIndex: referenceIndex,
            isDuplicateMode: insertSettings.isDuplicateMode,
            spacingPt: insertSettings.spacingPt,
            artboardGrid: artboardGrid,
            reflowFromHead: insertSettings.useManualSpacing && insertSettings.applySpacingToAll && artboardGrid.canInheritLayout,
            addedOnlyStep: addedOnlyStep,
            tailPrimaryOffset: (addCount + 1) * addedOnlyStep
        };
    }

    // 既存アートボードを最終位置へ移動し、アートワークも一緒に運ぶ。
    // 挿入位置以降は addCount セル分だけ後ろへ（追加分のみモードでは tailPrimaryOffset を加算して
    // 新規の間隔ぶんさらにずらす）。reflowFromHead=true（すべてモード）のときは先頭も新グリッドへ再配置。
    // アイテムの帰属は移動前の位置でまとめて取得（スナップショット）してから動かすので、
    // 移動順による取り違えが起きない。
    // Move existing artboards to their final positions, carrying their artwork.
    // Artboards at/after the insert point move back by `addCount` cells (plus tailPrimaryOffset
    // in added-only mode, to make room for the new spacing). When reflowFromHead is true (scope =
    // all) the head is also re-flowed onto the new grid. Artwork assignment is snapshotted
    // from the pre-move rects, so move order can't misassign items.
    function relayoutExistingArtboards(insertPlan) {
        var artboards = insertPlan.artboards;
        var plannedMoves = [];
        for (var i = 0; i < insertPlan.originalCount; i++) {
            var isTail = (i >= insertPlan.insertIndex);
            if (!isTail && !insertPlan.reflowFromHead) continue;
            var currentRect = artboards[i].artboardRect;
            var targetPosition = getGridCellPosition(insertPlan.artboardGrid, isTail ? i + insertPlan.addCount : i);
            if (isTail) targetPosition[insertPlan.artboardGrid.primaryAxis] += insertPlan.tailPrimaryOffset;
            plannedMoves.push({
                artboardIndex: i,
                dx: targetPosition[0] - currentRect[0],
                dy: targetPosition[1] - currentRect[1],
                assignedItems: getItemsAssignedToArtboard(currentRect)
            });
        }

        for (var j = 0; j < plannedMoves.length; j++) {
            var plannedMove = plannedMoves[j];
            for (var k = 0; k < plannedMove.assignedItems.length; k++) {
                var itemToMove = plannedMove.assignedItems[k];
                runWithItemUnlocked(itemToMove, function () {
                    itemToMove.translate(plannedMove.dx, plannedMove.dy);
                });
            }
            var movedArtboard = artboards[plannedMove.artboardIndex];
            var rectBeforeMove = movedArtboard.artboardRect;
            movedArtboard.artboardRect = [
                rectBeforeMove[0] + plannedMove.dx, rectBeforeMove[1] + plannedMove.dy,
                rectBeforeMove[2] + plannedMove.dx, rectBeforeMove[3] + plannedMove.dy
            ];
        }
    }

    // 指定数の新規アートボードを末尾に追加して名前を付け、複製モードなら内容もコピーする
    // Append the requested number of new artboards, name them, and copy the contents in duplicate mode
    function addNewArtboards(insertPlan) {
        var artboardGrid = insertPlan.artboardGrid;
        var addCount = insertPlan.addCount;
        var referenceArtboard = insertPlan.artboards[insertPlan.referenceIndex];
        var referenceRect = referenceArtboard.artboardRect;
        var referenceName = referenceArtboard.name;
        var newArtboardWidth = referenceRect[2] - referenceRect[0];
        var newArtboardHeight = referenceRect[3] - referenceRect[1];
        var nameSuffix = insertPlan.isDuplicateMode ? NAME_SUFFIX_DUPLICATE : NAME_SUFFIX_BLANK;

        for (var i = 0; i < addCount; i++) {
            var newPosition;
            if (artboardGrid.canInheritLayout) {
                // 追加分のみモードでは、各新規を指定間隔ぶん主軸方向へずらす
                // In added-only mode each one is nudged along the primary axis by the new spacing
                newPosition = getGridCellPosition(artboardGrid, insertPlan.insertIndex + i);
                newPosition[artboardGrid.primaryAxis] += (i + 1) * insertPlan.addedOnlyStep;
            } else {
                newPosition = getAnchoredPosition(insertPlan, referenceRect, i);
            }
            var newArtboard = insertPlan.artboards.add([
                newPosition[0], newPosition[1],
                newPosition[0] + newArtboardWidth, newPosition[1] + newArtboardHeight
            ]);
            // 複数追加時は連番を付与 / Append a sequence number when adding multiple
            newArtboard.name = referenceName + nameSuffix + (addCount > 1 ? ' ' + (i + 1) : '');
            if (insertPlan.isDuplicateMode) {
                duplicateArtboardContents(referenceRect, newArtboard.artboardRect);
            }
        }
    }

    // 末尾に追加された addCount 枚を、挿入位置 insertIndex..insertIndex+addCount-1 へ並べ替え。
    // 既存の後続（insertIndex..originalCount-1）は addCount 枚分だけ後ろへ送る。
    // Move the `addCount` just-appended artboards into slots insertIndex..insertIndex+addCount-1,
    // pushing the existing trailing artboards (insertIndex..originalCount-1) back by `addCount`.
    function reorderAppendedArtboards(insertPlan) {
        var artboards = insertPlan.artboards;
        var addCount = insertPlan.addCount;
        var insertIndex = insertPlan.insertIndex;
        var originalCount = insertPlan.originalCount;
        if (insertIndex >= originalCount) return; // 末尾への追加は並べ替え不要 / Appending needs no reorder

        // 追加分の枠と名前を退避（後続シフトで上書きされる前に）/ Save the appended rects and names first
        var appendedStates = [];
        for (var i = 0; i < addCount; i++) {
            var appendedArtboard = artboards[originalCount + i];
            appendedStates.push({ artboardRect: appendedArtboard.artboardRect, name: appendedArtboard.name });
        }
        // 後続を addCount 枚分だけ後ろへ（高位から処理して上書き衝突を回避）
        // Shift trailing artboards back by `addCount` (high → low to avoid clobbering)
        for (var j = originalCount - 1; j >= insertIndex; j--) {
            copyRectAndName(artboards[j], artboards[j + addCount]);
        }
        // 追加分を挿入位置へ配置 / Place the appended artboards at the insert position
        for (var k = 0; k < addCount; k++) {
            copyRectAndName(appendedStates[k], artboards[insertIndex + k]);
        }
    }

    // 枠と名前を写す（source は Artboard でも退避したオブジェクトでもよい）
    // Copy the rect and name (source can be an Artboard or a saved state object)
    function copyRectAndName(source, targetArtboard) {
        targetArtboard.artboardRect = source.artboardRect;
        targetArtboard.name = source.name;
    }

    // =========================================
    // アートワークの収集・移動・複製 / Artwork collection, moving and duplication
    // =========================================

    // アートボード上のアイテムを収集 / Collect items on an artboard
    // doc.pageItems はグループ・複合パスの子まで再帰的に含むため、
    // 最上位（親がレイヤー）のアイテムのみを対象にする。
    // 子は親と一緒に移動・複製されるので、ここで拾うと二重に処理されてしまう。
    // doc.pageItems also returns children of groups/compound paths, so only
    // collect top-level items (whose parent is a Layer). Children move and
    // duplicate together with their parent, so including them here would
    // process them twice.
    function getItemsAssignedToArtboard(artboardRect) {
        var assignedItems = [];
        var pageItems = doc.pageItems;
        for (var i = 0; i < pageItems.length; i++) {
            var pageItem = pageItems[i];
            if (pageItem.parent.typename !== 'Layer') continue;
            var itemBounds = pageItem.geometricBounds;
            var centerX = (itemBounds[0] + itemBounds[2]) / 2;
            var centerY = (itemBounds[1] + itemBounds[3]) / 2;
            if (centerX >= artboardRect[0] && centerX <= artboardRect[2] &&
                centerY <= artboardRect[1] && centerY >= artboardRect[3]) {
                assignedItems.push(pageItem);
            }
        }
        return assignedItems;
    }

    // ロック／非表示のアイテムは translate() や duplicate() が例外になる。
    // 途中で例外になると、そこまで動かしたアートボードだけが残って崩れるため、
    // 対象アイテムと祖先レイヤーのロック／表示状態を一時解除して action を実行し、終わったら元へ戻す。
    // translate() and duplicate() throw on locked/hidden items. An exception partway through
    // would leave the already-moved artboards in a broken state, so temporarily clear the
    // lock/visibility of the item and its ancestor layers, run the action, then restore them.
    function runWithItemUnlocked(pageItem, action) {
        var savedStates = [];

        // 値が違うときだけ元の値を控えてから書き換える / Record and overwrite only when the value differs
        function overrideState(target, propertyName, temporaryValue) {
            if (target[propertyName] === temporaryValue) return;
            savedStates.push({ target: target, propertyName: propertyName, originalValue: target[propertyName] });
            target[propertyName] = temporaryValue;
        }

        for (var ancestorLayer = pageItem.parent; ancestorLayer && ancestorLayer.typename === 'Layer'; ancestorLayer = ancestorLayer.parent) {
            overrideState(ancestorLayer, 'locked', false);
            overrideState(ancestorLayer, 'visible', true);
        }
        overrideState(pageItem, 'locked', false);
        overrideState(pageItem, 'hidden', false);

        try {
            action();
        } finally {
            // 書き換えと逆順に戻す / Restore in reverse order
            for (var i = savedStates.length - 1; i >= 0; i--) {
                savedStates[i].target[savedStates[i].propertyName] = savedStates[i].originalValue;
            }
        }
    }

    // sourceRect のアートボード上のアイテムを targetRect の位置へ複製（ロック／表示状態も引き継ぐ）
    // Duplicate the items on the sourceRect artboard to targetRect, carrying their lock/visibility
    function duplicateArtboardContents(sourceRect, targetRect) {
        var dx = targetRect[0] - sourceRect[0];
        var dy = targetRect[1] - sourceRect[1];
        var itemsToCopy = getItemsAssignedToArtboard(sourceRect);
        for (var i = 0; i < itemsToCopy.length; i++) {
            var sourceItem = itemsToCopy[i];
            var wasLocked = sourceItem.locked;
            var wasHidden = sourceItem.hidden;
            runWithItemUnlocked(sourceItem, function () {
                var copiedItem = sourceItem.duplicate();
                copiedItem.translate(dx, dy);
                // ロックした後だと表示状態を変えられないことがあるので hidden を先に
                // Set hidden first; a locked item may reject visibility changes
                copiedItem.hidden = wasHidden;
                copiedItem.locked = wasLocked;
            });
        }
    }

}());
