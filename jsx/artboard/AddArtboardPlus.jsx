#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
 
概要 / Overview
既存のアートボードの並び（行・列）を解析し、その規則を維持したまま新しいアートボードを挿入します。

・追加方法：「空のアートボード」または「現在のアートボードを複製」。初期選択は DEFAULT_ADD_METHOD で変更できます。
・追加位置：「現在のアートボードの次」または「末尾」。方向は「右」（横並び）／「下」（縦並び）から選べます。
・追加数：空のアートボードのときのみ有効で、指定した枚数をまとめて追加します。
・間隔：既存の並びから自動推定した値を定規単位で初期表示し、ダイアログ上で変更できます。
　適用範囲は「追加するアートボードのみ」または「すべてのアートボード」（並び全体を再配置）から選べます。

既存の並びが指定方向と一致する場合は、そのピッチと列数を引き継いだうえで後続のアートボードと
アートワーク（グループ・複合パスを含む）をシフトし、既存レイアウトを崩さずに挿入します。
一致しない場合や1枚のみの場合は既存の並びを動かさず、挿入位置の直前のアートボードを基準に指定方向へ配置します。
移動・複製の対象がロックまたは非表示でも、一時的に解除してから処理し、完了後に元の状態へ戻します。

追加されたアートボード名は、基準となるアートボード名に「_blank」（空）または「_copy」（複製）を付けたものになります。
実行後は最初に追加されたアートボードをアクティブにします。

キーボードショートカット（B / D / N / E）による操作にも対応しています。

Analyzes the current artboard layout (rows/columns) and inserts new artboards while preserving that structure.

- Add method: a blank artboard, or a duplicate of the current one. The initial choice is set by DEFAULT_ADD_METHOD.
- Insert position: after the current artboard or at the end, running either Right (horizontal) or Down (vertical).
- Count: enabled for blank artboards only, adding the requested number in one go.
- Spacing: auto-estimated from the existing layout, shown in the current ruler unit and editable in the dialog.
  It can be applied to the added artboards only, or to all artboards (re-flowing the whole arrangement).

When the existing layout runs along the chosen direction, its pitch and column count are inherited, and the
trailing artboards and their artwork (including groups and compound paths) are shifted so the layout stays intact.
Otherwise the existing artboards are left untouched and the new ones are placed relative to the artboard just
before the insert point. Locked or hidden artwork is temporarily unlocked/shown and restored afterwards.

The added artboard is named after the reference artboard plus _blank (blank) or _copy (duplicate).
After execution, the first newly added artboard becomes the active artboard.
Keyboard shortcuts (B / D / N / E) are also supported for quick operation.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AddArtboardPlus";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-15";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-07-22";                   /* 更新日 / last updated */

var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/naf239a44b8ff"; /* 紹介記事 / article URL */

// オリジナル / Original
// Copyright (c) 2018 Takeshi Umeda (noellabo)
// https://dtp-discourse.jp/t/illustrator/99
// Modified by Masahiro Takano in 2026

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 追加したアートボード名に付ける接尾辞 / Suffix appended to the new artboard's name */
    var SUFFIX_BLANK = '_blank';
    var SUFFIX_COPY = '_copy';

    /* 「追加方法」の初期選択 / Initial selection of the Add Method panel */
    /* 'blank' = 空のアートボード / 'duplicate' = 現在のアートボードを複製 */
    /* 'blank' = blank artboard, 'duplicate' = duplicate the current artboard */
    var DEFAULT_ADD_METHOD = 'blank';

    // =========================================
    // ローカライズ / Localization
    // =========================================

    function getCurrentLang() {
        return ($.locale.indexOf('ja') === 0) ? 'ja' : 'en';
    }
    var lang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: 'アートボードを追加', en: 'Add Artboard' }
        },
        panel: {
            addMethod: { ja: '追加方法', en: 'Add Method' },
            insertPosition: { ja: '追加位置', en: 'Insert Position' },
            spacing: { ja: '間隔', en: 'Spacing' }
        },
        radio: {
            blank: { ja: '空のアートボード', en: 'Blank Artboard' },
            duplicate: { ja: '現在のアートボードを複製', en: 'Duplicate Current Artboard' },
            insertAfterCurrent: { ja: '現在のアートボードの次', en: 'After Current Artboard' },
            insertAtEnd: { ja: '末尾', en: 'At End' },
            scopeAddedOnly: { ja: '追加するアートボードのみ', en: 'Added artboards only' },
            scopeAll: { ja: 'すべてのアートボード', en: 'All artboards' },
            directionRight: { ja: '右', en: 'Right' },
            directionDown: { ja: '下', en: 'Down' }
        },
        label: {
            spacing: { ja: '間隔', en: 'Spacing' },
            addCount: { ja: '追加数', en: 'Count' },
            direction: { ja: '方向', en: 'Direction' }
        },
        button: {
            cancel: { ja: 'キャンセル', en: 'Cancel' },
            ok: { ja: 'OK', en: 'OK' }
        },
        error: {
            noDocument: {
                ja: 'ドキュメントがないため実行できません。',
                en: 'Cannot run because no document is open.'
            },
            exception: {
                ja: 'エラーが発生したため、処理を実行できませんでした。\nエラー内容：',
                en: 'Processing could not be completed because an error occurred.\nError:'
            },
            artboardLimit: {
                ja: 'アートボードの最大作成可能数を超えています。',
                en: 'The maximum number of artboards would be exceeded.'
            },
            noSpace: {
                ja: 'アートボードを作成する十分なスペースがありません。',
                en: 'There is not enough space to create the artboard.'
            }
        }
    };

    /* ラベル（ja/en のリーフ）を現在の言語に解決。{slash} は / に置換 */
    /* Resolve a label leaf (ja/en) to the current language; {slash} becomes / */
    function L(entry) {
        var text = (entry && entry[lang]) || '';
        return text.replace(/\{slash\}/g, '/');
    }

    // =========================================
    // 単位 / Unit
    // =========================================

    // 現在の定規単位の pt 換算係数とラベルを取得
    // Get the pt conversion factor and label for the current ruler unit
    function getRulerUnitInfo() {
        var rulerType = app.preferences.getIntegerPreference('rulerType');
        switch (rulerType) {
            case 0: return { factor: 72.0, label: 'inch' };          // インチ / inch
            case 1: return { factor: 72.0 / 25.4, label: 'mm' };     // ミリ / mm
            case 2: return { factor: 1.0, label: 'pt' };             // ポイント / point
            case 3: return { factor: 12.0, label: 'pica' };          // パイカ / pica
            case 4: return { factor: 72.0 / 2.54, label: 'cm' };     // センチ / cm
            case 5: return { factor: 72.0 / 25.4 * 0.25, label: 'Q' }; // 歯 / Q
            case 6: return { factor: 1.0, label: 'px' };             // ピクセル / pixel
            default: return { factor: 1.0, label: 'pt' };
        }
    }

    // 表示用に小数を整える / Trim a number for display
    function formatSpacingValue(value) {
        return String(Math.round(value * 100) / 100);
    }

    // 軸方向のアートボードサイズ（0=幅 / 1=高さ）/ Artboard size along an axis (0=width, 1=height)
    function getAxisSize(rect, axisIndex) {
        return (axisIndex === 0)
            ? (rect[2] - rect[0])           // 幅 / width
            : Math.abs(rect[3] - rect[1]);  // 高さ / height
    }

    // 隣り合う2枚から主軸を判定（左端が同じ＝縦並び→1 / 違う＝横並び→0）
    // Determine the primary axis from two neighbors (same left edge ⇒ vertical → 1, otherwise horizontal → 0)
    function getPrimaryAxisIndex(rectA, rectB) {
        return +(rectA[0] == rectB[0]);
    }

    // 既存アートボードの並びから現在の間隔（pt）を推定
    // 2枚未満の場合は fallbackPt（環境設定値）を返す
    // Estimate the current spacing (pt) from the existing artboard arrangement
    // Returns fallbackPt (the preference value) when there are fewer than 2 artboards
    function computeAutoSpacingPt(fallbackPt) {
        var artboardList = doc.artboards;
        if (artboardList.length < 2) return fallbackPt;

        var baseRect = artboardList[0].artboardRect;
        var adjacentRect = artboardList[1].artboardRect;
        var primaryAxisIndex = getPrimaryAxisIndex(baseRect, adjacentRect);
        var pitch = Math.abs(adjacentRect[primaryAxisIndex] - baseRect[primaryAxisIndex]);
        var spacing = pitch - getAxisSize(baseRect, primaryAxisIndex);
        return (spacing < 0) ? 0 : spacing;
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
        alert(L(LABELS.error.noDocument));
        return;
    }

    try {
        var doc = app.activeDocument;
        var appPreferences = app.preferences;

        // ダイアログボックスを表示 / Show dialog
        var dialog = new Window('dialog', L(LABELS.dialog.title) + ' ' + SCRIPT_VERSION);
        setupWindow(dialog);

        // 追加方法パネル / Add method panel
        var addMethodPanel = dialog.add('panel', undefined, L(LABELS.panel.addMethod));
        setupPanel(addMethodPanel, 6);
        var blankArtboardRadio = addMethodPanel.add('radiobutton', undefined, L(LABELS.radio.blank));
        var duplicateArtboardRadio = addMethodPanel.add('radiobutton', undefined, L(LABELS.radio.duplicate));
        // 初期選択はユーザー設定に従う（'duplicate' 以外は「空のアートボード」）
        // The initial selection follows the user setting (anything but 'duplicate' means blank)
        duplicateArtboardRadio.value = (DEFAULT_ADD_METHOD === 'duplicate');
        blankArtboardRadio.value = !duplicateArtboardRadio.value;

        // 追加数 / Count
        var addCountRow = addMethodPanel.add('group');
        setupRow(addCountRow);
        var addCountLabel = addCountRow.add('statictext', undefined, L(LABELS.label.addCount) + '：');
        var addCountInput = addCountRow.add('edittext', undefined, '1');
        addCountInput.characters = 4;
        changeValueByArrowKey(addCountInput);

        // 追加位置パネル / Insert position panel
        var insertPositionPanel = dialog.add('panel', undefined, L(LABELS.panel.insertPosition));
        setupPanel(insertPositionPanel, 6);
        var insertAfterCurrentRadio = insertPositionPanel.add('radiobutton', undefined, L(LABELS.radio.insertAfterCurrent));
        var insertAtEndRadio = insertPositionPanel.add('radiobutton', undefined, L(LABELS.radio.insertAtEnd));

        // 追加方向（右＝横並び / 下＝縦並び）/ Add direction (right = horizontal, down = vertical)
        var directionRow = insertPositionPanel.add('group');
        setupRow(directionRow);
        directionRow.add('statictext', undefined, L(LABELS.label.direction) + '：');
        var directionRightRadio = directionRow.add('radiobutton', undefined, L(LABELS.radio.directionRight));
        var directionDownRadio = directionRow.add('radiobutton', undefined, L(LABELS.radio.directionDown));
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
        var spacingPanel = dialog.add('panel', undefined, L(LABELS.panel.spacing));
        setupPanel(spacingPanel, 6);
        var spacingScopeAddedRadio = spacingPanel.add('radiobutton', undefined, L(LABELS.radio.scopeAddedOnly));
        var spacingScopeAllRadio = spacingPanel.add('radiobutton', undefined, L(LABELS.radio.scopeAll));
        spacingScopeAddedRadio.value = true;

        // 間隔入力 / Spacing input
        // 2枚以上は既存の並びから推定した間隔、1枚以下は環境設定の値（pt）を
        // 定規単位に変換して初期表示し、編集可能にする
        // For 2+ artboards use the spacing inferred from the existing layout,
        // otherwise the preference value (pt); shown in the ruler unit and editable
        var rulerUnit = getRulerUnitInfo();
        var spacingPreferencePt = appPreferences.getRealPreference('plugin/ArtboardRearrange/ArtboardSpacing'); // rearrange (pt)
        var autoSpacingPt = computeAutoSpacingPt(spacingPreferencePt);
        var initialSpacingInUnit = autoSpacingPt / rulerUnit.factor;
        var initialSpacingText = formatSpacingValue(initialSpacingInUnit);

        var spacingRow = spacingPanel.add('group');
        setupRow(spacingRow);
        spacingRow.add('statictext', undefined, L(LABELS.label.spacing) + '：');
        var spacingInputField = spacingRow.add('edittext', undefined, initialSpacingText);
        spacingInputField.characters = 4;
        changeValueByArrowKey(spacingInputField);
        spacingRow.add('statictext', undefined, rulerUnit.label);

        addShortcutKeyHandler(dialog, blankArtboardRadio, duplicateArtboardRadio, insertAfterCurrentRadio, insertAtEndRadio, syncUIWithAddMethod);

        // ボタン列はダイアログ幅いっぱいに広げず中央に置く / Center the button row instead of stretching it
        var buttonRow = dialog.add('group');
        setupRow(buttonRow, 'center');
        buttonRow.add('button', undefined, L(LABELS.button.cancel), { name: 'cancel' });
        buttonRow.add('button', undefined, L(LABELS.button.ok), { name: 'ok' });

        if (dialog.show() != 1) return;

        var isDuplicateMode = duplicateArtboardRadio.value;
        var insertAfterCurrent = insertAfterCurrentRadio.value;

        // 入力値を pt に戻す。空欄・不正値は初期表示値にフォールバック
        // Convert the input back to pt; fall back to the initial value for blank/invalid input
        var spacingInputValue = parseFloat(spacingInputField.text);
        if (isNaN(spacingInputValue)) spacingInputValue = initialSpacingInUnit;
        var spacing = spacingInputValue * rulerUnit.factor;

        // 手動で初期値から変更されたかどうか。変更時のみ入力値を優先
        // Whether the value was manually changed from the initial; the input wins only when changed
        var useManualSpacing = (spacingInputField.text !== initialSpacingText);

        // 追加数（1以上の整数）。複製モードでは常に1枚 / Number to add (integer ≥ 1); duplicate mode always adds 1
        var addCount = parseInt(addCountInput.text, 10);
        if (isNaN(addCount) || addCount < 1) addCount = 1;
        if (isDuplicateMode) addCount = 1;

        // 間隔の適用範囲：true=すべてのアートボード / false=追加分のみ
        // Spacing scope: true = all artboards, false = added artboards only
        var applySpacingToAll = spacingScopeAllRadio.value;

        // 追加方向：0=右（横並び）/ 1=下（縦並び）/ Add direction: 0 = right (horizontal), 1 = down (vertical)
        var directionAxisIndex = directionDownRadio.value ? 1 : 0;

        insertArtboardWithShift(isDuplicateMode, insertAfterCurrent, spacing, useManualSpacing, addCount, applySpacingToAll, directionAxisIndex);

    } catch (e) {
        alert(L(LABELS.error.exception) + e);
    }

    /*
    最大キャンバス範囲を取得 / Get the largest canvas bounds
    Original idea by OMOTI
    https://forums.adobe.com/thread/2459293
    */
    function getLargestCanvasBounds() {
        var tempLayer,
            tempText,
            left,
            top,
            LARGEST_SIZE = 16383;

        if (!app.documents.length) {
            return;
        }

        tempLayer = app.activeDocument.layers.add();
        tempText = tempLayer.textFrames.add();
        left = tempText.matrix.mValueTX;
        top = tempText.matrix.mValueTY;
        tempLayer.remove();

        return new Rect(
            left,
            top,
            left + LARGEST_SIZE,
            top - LARGEST_SIZE
        );
    }

    function insertArtboardWithShift(isDuplicateMode, insertAfterCurrent, spacing, useManualSpacing, addCount, applySpacingToAll, directionAxisIndex) {
        var artboards = doc.artboards;
        var artboardCount = artboards.length;
        var artboardLimit = (parseFloat(app.version) >= 22) ? 1000 : 100;

        // 現在選択しているアートボードのインデックスを取得 / Get active artboard index
        var activeArtboardIndex = doc.artboards.getActiveArtboardIndex();

        if (artboardCount + addCount > artboardLimit) {
            alert(L(LABELS.error.artboardLimit));
            return;
        }

        // 挿入位置を決定 / Determine insertion index
        var insertIndex = insertAfterCurrent ? activeArtboardIndex + 1 : artboardCount;

        // 新規アートボードのサイズと名前の基準。複製は「現在のアートボード」、
        // 空のアートボードは挿入位置の直前を基準にする（サイズがまちまちのドキュメントで
        // 末尾に追加したのに1枚目のサイズになる、といったズレを避けるため）
        // Reference artboard for the new artboards' size and name: the current artboard for
        // duplicates, otherwise the artboard just before the insert point. Keeps appending at the
        // end from producing an artboard sized like the first one in a mixed-size document.
        var referenceArtboardIndex = isDuplicateMode ? activeArtboardIndex : (insertIndex - 1);
        var referenceArtboardRect = artboards[referenceArtboardIndex].artboardRect;
        var newArtboardWidth = referenceArtboardRect[2] - referenceArtboardRect[0];
        var newArtboardHeight = referenceArtboardRect[3] - referenceArtboardRect[1];

        // グリッドの計算は先頭のアートボードを原点にする / The grid math is anchored on the first artboard
        var baseArtboardRect = artboards[0].artboardRect;
        var artboardWidth = baseArtboardRect[2] - baseArtboardRect[0];
        var artboardHeight = baseArtboardRect[3] - baseArtboardRect[1];

        // spacing はダイアログで指定された間隔（pt） / spacing is the spacing (pt) given in the dialog
        // グリッド1セルあたりの移動量（幅+間隔、高さは上→下なので負方向に間隔を引く）
        // Grid step per cell (width + spacing, height subtracts spacing because Y axis is top-down)
        var gridStep = [
            artboardWidth + spacing,
            artboardHeight - spacing
        ];
        var columns = 0;
        var rows = 0;
        // 既存の並びのグリッドを引き継げるか / Whether the existing layout's grid can be inherited
        var canInheritLayout = false;

        // 主軸はダイアログで選んだ追加方向（0=右＝横 / 1=下＝縦）
        // Primary axis is the chosen add direction (0 = right/horizontal, 1 = down/vertical)
        var layoutPrimaryAxisIndex = directionAxisIndex;
        var layoutSecondaryAxisIndex = 1 - directionAxisIndex;

        // 既存の並びが指定方向と一致するときだけ、ピッチと列数を引き継ぐ
        // （一致しない・1枚のみのときはグリッドを使わず、挿入位置の直前のアートボードを
        //   基準に指定方向へ並べる → getAnchoredArtboardPosition）
        // Inherit pitch and columns only when the existing layout matches the chosen direction
        // (otherwise skip the grid entirely and place along the chosen direction relative to the
        //  artboard just before the insert point → getAnchoredArtboardPosition)
        if (artboardCount >= 2) {
            var adjacentArtboardRect = artboards[1].artboardRect;
            var detectedPrimaryAxisIndex = getPrimaryAxisIndex(baseArtboardRect, adjacentArtboardRect);

            if (detectedPrimaryAxisIndex === directionAxisIndex) {
                canInheritLayout = true;
                gridStep[layoutPrimaryAxisIndex] = adjacentArtboardRect[layoutPrimaryAxisIndex] - baseArtboardRect[layoutPrimaryAxisIndex];

                for (var i = 2; i < artboardCount; i++) {
                    var scannedArtboardRect = artboards[i].artboardRect;
                    if (baseArtboardRect[layoutSecondaryAxisIndex] != scannedArtboardRect[layoutSecondaryAxisIndex]) {
                        gridStep[layoutSecondaryAxisIndex] = scannedArtboardRect[layoutSecondaryAxisIndex] - baseArtboardRect[layoutSecondaryAxisIndex];
                        columns = i;
                        break;
                    }
                }
            }

            // 「すべてのアートボード」かつ手動間隔のときだけ、並び全体を入力値で再配置するため
            // グリッドの移動量を作り直す（追加分のみのときは既存のピッチを保持）
            // Only when scope = all AND the spacing was changed manually, rebuild the grid step
            // from the input to re-flow the whole arrangement (scope = added-only keeps the existing pitch)
            if (useManualSpacing && applySpacingToAll) {
                var primarySize = getAxisSize(baseArtboardRect, layoutPrimaryAxisIndex);
                var secondarySize = getAxisSize(baseArtboardRect, layoutSecondaryAxisIndex);
                var primarySign = (gridStep[layoutPrimaryAxisIndex] < 0) ? -1 : 1;
                var secondarySign = (gridStep[layoutSecondaryAxisIndex] < 0) ? -1 : 1;
                gridStep[layoutPrimaryAxisIndex] = primarySign * (primarySize + spacing);
                gridStep[layoutSecondaryAxisIndex] = secondarySign * (secondarySize + spacing);
            }
        }

        var canvasRect = getLargestCanvasBounds();

        var gridUnitRect = [
            baseArtboardRect[0] + Math.abs(gridStep[0]),
            baseArtboardRect[1] - Math.abs(gridStep[1]),
            baseArtboardRect[0],
            baseArtboardRect[1]
        ];

        var layoutPrimaryEdgeIndex =
            (layoutPrimaryAxisIndex ^ +(gridStep[layoutPrimaryAxisIndex] < 0)) ? layoutPrimaryAxisIndex : layoutPrimaryAxisIndex + 2;
        var layoutSecondaryEdgeIndex =
            (layoutSecondaryAxisIndex ^ +(gridStep[layoutSecondaryAxisIndex] < 0)) ? layoutSecondaryAxisIndex : layoutSecondaryAxisIndex + 2;

        columns = columns ||
            Math.abs(Math.floor((canvasRect[layoutPrimaryEdgeIndex] - gridUnitRect[layoutPrimaryEdgeIndex]) / gridStep[layoutPrimaryAxisIndex]));
        rows = Math.abs(Math.floor((canvasRect[layoutSecondaryEdgeIndex] - gridUnitRect[layoutSecondaryEdgeIndex]) / gridStep[layoutSecondaryAxisIndex]));
        if (artboardCount + addCount > columns * rows) {
            alert(L(LABELS.error.noSpace));
            return;
        }

        // 間隔の適用範囲ごとの移動量 / Movement amounts depending on the spacing scope
        //  - すべて：手動間隔のとき先頭も含めて新グリッドへ再配置 / all: re-flow everything (head too)
        //  - 追加分のみ：既存はそのまま、新規だけを (指定間隔 − 既存の隙間) ぶん主軸方向へずらす
        //    added-only: keep existing as-is, nudge only the new artboards by (spacing − existing gap)
        var primaryDirectionSign = (gridStep[layoutPrimaryAxisIndex] < 0) ? -1 : 1;
        var primaryArtboardSize = getAxisSize(baseArtboardRect, layoutPrimaryAxisIndex);
        var existingPrimaryGap = Math.abs(gridStep[layoutPrimaryAxisIndex]) - primaryArtboardSize;
        if (existingPrimaryGap < 0) existingPrimaryGap = 0;
        var addedOnlySpacingDelta = (useManualSpacing && !applySpacingToAll) ? (spacing - existingPrimaryGap) : 0;
        var reflowFromHead = (useManualSpacing && applySpacingToAll && canInheritLayout);
        var tailPrimaryOffset = primaryDirectionSign * addCount * addedOnlySpacingDelta;

        // グリッド上のインデックスの位置を計算 / Calculate the position of a grid index
        function getArtboardGridPosition(gridIndex) {
            var offset = [];
            offset[layoutPrimaryAxisIndex] = (gridIndex % columns) * gridStep[layoutPrimaryAxisIndex];
            offset[layoutSecondaryAxisIndex] = Math.floor(gridIndex / columns) * gridStep[layoutSecondaryAxisIndex];
            return [baseArtboardRect[0] + offset[0], baseArtboardRect[1] + offset[1]];
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
        function getAnchoredArtboardPosition(anchorIndex, addedOrder) {
            var anchorRect = artboards[anchorIndex].artboardRect;
            // 1歩目はアンカーのサイズ、2枚目以降は新規アートボードのサイズで進む
            // The first step uses the anchor's size; later ones use the new artboards' size
            var advance = getAxisSize(anchorRect, layoutPrimaryAxisIndex) + spacing
                + addedOrder * (getAxisSize(referenceArtboardRect, layoutPrimaryAxisIndex) + spacing);
            var position = [anchorRect[0], anchorRect[1]];
            position[layoutPrimaryAxisIndex] += primaryDirectionSign * advance;
            return position;
        }

        // アートボード上のアイテムを収集 / Collect items on an artboard
        // doc.pageItems はグループ・複合パスの子まで再帰的に含むため、
        // 最上位（親がレイヤー）のアイテムのみを対象にする。
        // 子は親と一緒に移動・複製されるので、ここで拾うと二重に処理されてしまう。
        // doc.pageItems also returns children of groups/compound paths, so only
        // collect top-level items (whose parent is a Layer). Children move and
        // duplicate together with their parent, so including them here would
        // process them twice.
        function getItemsAssignedToArtboard(artboardRect) {
            var items = [];
            for (var k = 0; k < doc.pageItems.length; k++) {
                var item = doc.pageItems[k];
                if (item.parent.typename !== 'Layer') continue;
                var itemBounds = item.geometricBounds;
                var centerX = (itemBounds[0] + itemBounds[2]) / 2;
                var centerY = (itemBounds[1] + itemBounds[3]) / 2;
                if (centerX >= artboardRect[0] && centerX <= artboardRect[2] &&
                    centerY <= artboardRect[1] && centerY >= artboardRect[3]) {
                    items.push(item);
                }
            }
            return items;
        }

        // ロック／非表示のアイテムは translate() や duplicate() が例外になる。
        // 途中で例外になると、そこまで動かしたアートボードだけが残って崩れるため、
        // 対象アイテムと祖先レイヤーのロック／表示状態を一時解除し、処理後に元へ戻す。
        // translate() and duplicate() throw on locked/hidden items. An exception partway through
        // would leave the already-moved artboards in a broken state, so temporarily clear the
        // lock/visibility of the item and its ancestor layers, then restore them afterwards.
        function unlockItemTemporarily(item) {
            var lockRestoreList = [];
            var ancestorLayer = item.parent;
            while (ancestorLayer && ancestorLayer.typename === 'Layer') {
                if (ancestorLayer.locked) { ancestorLayer.locked = false; lockRestoreList.push({ target: ancestorLayer, property: 'locked', value: true }); }
                if (!ancestorLayer.visible) { ancestorLayer.visible = true; lockRestoreList.push({ target: ancestorLayer, property: 'visible', value: false }); }
                ancestorLayer = ancestorLayer.parent;
            }
            if (item.locked) { item.locked = false; lockRestoreList.push({ target: item, property: 'locked', value: true }); }
            if (item.hidden) { item.hidden = false; lockRestoreList.push({ target: item, property: 'hidden', value: true }); }
            return lockRestoreList;
        }

        /* 解除と逆順に戻す（内側→外側）/ Restore in reverse order (inner → outer) */
        function restoreLockedState(lockRestoreList) {
            for (var i = lockRestoreList.length - 1; i >= 0; i--) {
                lockRestoreList[i].target[lockRestoreList[i].property] = lockRestoreList[i].value;
            }
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
        function relayoutExistingArtboards(fromIndex, reflowFromHead, tailPrimaryOffset) {
            var plannedMoves = [];
            for (var i = 0; i < artboardCount; i++) {
                if (!reflowFromHead && i < fromIndex) continue;
                var targetIndex = (i >= fromIndex) ? i + addCount : i;
                var currentRect = artboards[i].artboardRect;
                var targetPosition = getArtboardGridPosition(targetIndex);
                if (i >= fromIndex) targetPosition[layoutPrimaryAxisIndex] += tailPrimaryOffset;
                plannedMoves.push({
                    index: i,
                    dx: targetPosition[0] - currentRect[0],
                    dy: targetPosition[1] - currentRect[1],
                    items: getItemsAssignedToArtboard(currentRect)
                });
            }

            for (var j = 0; j < plannedMoves.length; j++) {
                var plannedMove = plannedMoves[j];
                for (var k = 0; k < plannedMove.items.length; k++) {
                    var itemToMove = plannedMove.items[k];
                    var lockRestoreList = unlockItemTemporarily(itemToMove);
                    try {
                        itemToMove.translate(plannedMove.dx, plannedMove.dy);
                    } finally {
                        restoreLockedState(lockRestoreList);
                    }
                }
                var movedRect = artboards[plannedMove.index].artboardRect;
                artboards[plannedMove.index].artboardRect = [
                    movedRect[0] + plannedMove.dx, movedRect[1] + plannedMove.dy,
                    movedRect[2] + plannedMove.dx, movedRect[3] + plannedMove.dy
                ];
            }
        }

        // 末尾に追加された addCount 枚を、挿入位置 fromIndex..fromIndex+addCount-1 へ並べ替え。
        // 既存の後続（fromIndex..originalCount-1）は addCount 枚分だけ後ろへ送る。
        // Move the `addCount` just-appended artboards into slots fromIndex..fromIndex+addCount-1,
        // pushing the existing trailing artboards (fromIndex..originalCount-1) back by `addCount`.
        function reorderAppendedArtboards(fromIndex, originalCount) {
            // 追加分の枠と名前を退避（後続シフトで上書きされる前に）/ Save the appended rects and names first
            var appendedArtboards = [];
            for (var i = 0; i < addCount; i++) {
                var appendedArtboard = artboards[originalCount + i];
                appendedArtboards.push({ rect: appendedArtboard.artboardRect, name: appendedArtboard.name });
            }
            // 後続を addCount 枚分だけ後ろへ（高位から処理して上書き衝突を回避）
            // Shift trailing artboards back by `addCount` (high → low to avoid clobbering)
            for (var j = originalCount - 1; j >= fromIndex; j--) {
                artboards[j + addCount].artboardRect = artboards[j].artboardRect;
                artboards[j + addCount].name = artboards[j].name;
            }
            // 追加分を挿入位置へ配置 / Place the appended artboards at the insert position
            for (var k = 0; k < addCount; k++) {
                artboards[fromIndex + k].artboardRect = appendedArtboards[k].rect;
                artboards[fromIndex + k].name = appendedArtboards[k].name;
            }
        }

        // 複製モード：sourceIndex の内容を targetIndex へコピー / Duplicate mode: copy contents from sourceIndex to targetIndex
        function duplicateArtboardContents(sourceIndex, targetIndex) {
            var sourceRect = artboards[sourceIndex].artboardRect;
            var targetRect = artboards[targetIndex].artboardRect;
            var offsetX = targetRect[0] - sourceRect[0];
            var offsetY = targetRect[1] - sourceRect[1];

            var itemsToCopy = getItemsAssignedToArtboard(sourceRect);
            for (var i = 0; i < itemsToCopy.length; i++) {
                var sourceItem = itemsToCopy[i];
                // 複製元のロック／表示状態は複製先にも引き継ぐ / Carry the source's lock/visibility to the copy
                var wasLocked = sourceItem.locked;
                var wasHidden = sourceItem.hidden;
                var lockRestoreList = unlockItemTemporarily(sourceItem);
                try {
                    var copiedItem = sourceItem.duplicate();
                    copiedItem.translate(offsetX, offsetY);
                    copiedItem.locked = wasLocked;
                    copiedItem.hidden = wasHidden;
                } finally {
                    restoreLockedState(lockRestoreList);
                }
            }
        }

        // 既存アートボードを最終位置へ移動して挿入スペースを空ける。
        // グリッドを引き継げないときは既存の並びを動かさない（別軸のグリッドへ流し込むと
        // 先頭は横並び・後続は縦並びのように崩れるため）
        // Move existing artboards to their final positions to free the insert space.
        // When the grid can't be inherited, leave the existing artboards alone — re-flowing them
        // onto a grid that runs along the other axis would break the arrangement apart.
        if (canInheritLayout) {
            relayoutExistingArtboards(insertIndex, reflowFromHead, tailPrimaryOffset);
        }

        // 指定数の新規アートボードを追加（いずれも末尾に追加される）
        // 追加分のみモードでは、各新規を指定間隔ぶん主軸方向へずらす
        // Add the requested number of new artboards (each appended at the end);
        // in added-only mode each one is nudged along the primary axis by the new spacing
        for (var i = 0; i < addCount; i++) {
            var newArtboardPosition = canInheritLayout
                ? getArtboardGridPosition(insertIndex + i)
                : getAnchoredArtboardPosition(insertIndex - 1, i);
            if (canInheritLayout) {
                newArtboardPosition[layoutPrimaryAxisIndex] += primaryDirectionSign * (i + 1) * addedOnlySpacingDelta;
            }
            artboards.add([
                newArtboardPosition[0], newArtboardPosition[1],
                newArtboardPosition[0] + newArtboardWidth, newArtboardPosition[1] + newArtboardHeight
            ]);
        }

        // パネル上の順序を挿入位置へ並べ替え / Reorder in the panel to the insert position
        reorderAppendedArtboards(insertIndex, artboardCount);

        // 新規アートボードの命名と、複製モードなら内容のコピー。
        // 複製は「現在のアートボード」が元なのでアクティブを基準にする。
        // 空のアートボードは挿入位置の直前を基準にする（末尾に追加したのに
        // 先頭の名前が付く、といったズレを避けるため）
        // Name the new artboards and, in duplicate mode, copy the contents.
        // Duplicates are copies of the *current* artboard, so they follow the active one.
        // Blank artboards follow the artboard just before the insert point, so appending at the
        // end doesn't produce a name taken from an artboard elsewhere in the document.
        var nameSourceIndex = isDuplicateMode ? activeArtboardIndex : (insertIndex - 1);
        var sourceArtboardName = artboards[nameSourceIndex].name;
        var nameSuffix = isDuplicateMode ? SUFFIX_COPY : SUFFIX_BLANK;
        for (var j = 0; j < addCount; j++) {
            var newArtboardIndex = insertIndex + j;
            // 複数追加時は連番を付与 / Append a sequence number when adding multiple
            artboards[newArtboardIndex].name = sourceArtboardName + nameSuffix + (addCount > 1 ? ' ' + (j + 1) : '');
            if (isDuplicateMode) {
                duplicateArtboardContents(activeArtboardIndex, newArtboardIndex);
            }
        }

        // 最初に追加したアートボードを選択・アクティブにする
        // Select and activate the first newly added artboard
        doc.selection = null;
        doc.artboards.setActiveArtboardIndex(insertIndex);
        app.redraw();
    }

    // =========================================
    // キー入力で選択切り替え / Change selections with keyboard shortcuts
    // =========================================
    function addShortcutKeyHandler(targetDialog, blankRadio, duplicateRadio, afterCurrentRadio, atEndRadio, onAddMethodChange) {
        targetDialog.addEventListener('keydown', function (event) {
            var keyName = event.keyName;
            if (!keyName) return;

            keyName = String(keyName).toUpperCase();

            if (keyName === 'B') {
                blankRadio.value = true;
                duplicateRadio.value = false;
                if (onAddMethodChange) onAddMethodChange();
                event.preventDefault();
            } else if (keyName === 'D') {
                blankRadio.value = false;
                duplicateRadio.value = true;
                if (onAddMethodChange) onAddMethodChange();
                event.preventDefault();
            } else if (keyName === 'N') {
                afterCurrentRadio.value = true;
                atEndRadio.value = false;
                event.preventDefault();
            } else if (keyName === 'E') {
                afterCurrentRadio.value = false;
                atEndRadio.value = true;
                event.preventDefault();
            }
        });
    }

}());