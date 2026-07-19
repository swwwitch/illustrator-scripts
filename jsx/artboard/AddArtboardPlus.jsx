#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
AddArtboardPlus.jsx
 
概要 / Overview
既存のアートボードの並び（行・列）を解析し、その規則を維持したまま新しいアートボードを挿入します。
「空のアートボード」または「現在のアートボードの内容を複製」を選択でき、
追加位置は「現在の次」または「末尾」から指定できます。
アートボード間の間隔は既存の並びから自動推定して定規単位で初期表示し、ダイアログ上で変更できます。
間隔の適用範囲は「追加するアートボードのみ」または「すべてのアートボード」（並び全体を再配置）から選べます。
「追加数」を指定すると、その枚数だけ新規アートボードをまとめて追加できます。
挿入時は後続のアートボードとアートワーク（グループ・複合パスを含む）を安全にシフトし、既存レイアウトを崩さずに追加します。
追加されたアートボード名は、空のアートボードでは「元のアートボード名_blank」、複製では「元のアートボード名_copy」として設定されます。
実行後は追加されたアートボードをアクティブにします。

キーボードショートカット（B / D / N / E）による操作にも対応しています。

Analyzes the current artboard layout (rows/columns) and inserts a new artboard while preserving that structure.
You can create a blank artboard or duplicate the current artboard, and choose to insert it after the current one or at the end.
The spacing between artboards is auto-estimated from the existing layout, shown in the current ruler unit, and editable in the dialog.
The spacing can be applied to the added artboards only, or to all artboards (re-flowing the whole arrangement).
A Count field lets you add several new artboards at once.
During insertion, subsequent artboards and artwork (including groups and compound paths) are safely shifted so the existing layout remains intact.
The added artboard is named with the original artboard name plus _blank for a blank artboard or _copy for a duplicated one.
After execution, the newly added artboard becomes the active artboard.
Keyboard shortcuts (B / D / N / E) are also supported for quick operation.
 
Copyright (c) 2018 Takeshi Umeda (noellabo)
https://dtp-discourse.jp/t/illustrator/99
 
Modified by Masahiro Takano in 2026
 
Released under the MIT license
http://opensource.org/licenses/mit-license.php
 
Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:
 
The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.
 
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
 
*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AddArtboardPlus";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-06-04";                   /* 更新日 / last updated */

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    var SUFFIX_BLANK = '_blank';
    var SUFFIX_COPY = '_copy';

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
            mode: { ja: '追加方法', en: 'Add Method' },
            position: { ja: '追加位置', en: 'Insert Position' },
            spacing: { ja: '間隔', en: 'Spacing' }
        },
        radio: {
            blank: { ja: '空のアートボード', en: 'Blank Artboard' },
            duplicate: { ja: '現在のアートボードを複製', en: 'Duplicate Current Artboard' },
            next: { ja: '現在のアートボードの次', en: 'After Current Artboard' },
            last: { ja: '末尾', en: 'At End' },
            scopeNew: { ja: '追加するアートボードのみ', en: 'Added artboards only' },
            scopeAll: { ja: 'すべてのアートボード', en: 'All artboards' },
            directionRight: { ja: '右', en: 'Right' },
            directionDown: { ja: '下', en: 'Down' }
        },
        label: {
            spacing: { ja: '間隔', en: 'Spacing' },
            count: { ja: '追加数', en: 'Count' },
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
    function getPrimaryAxis(rectA, rectB) {
        return +(rectA[0] == rectB[0]);
    }

    // 既存アートボードの並びから現在の間隔（pt）を推定
    // 2枚未満の場合は fallbackPt（環境設定値）を返す
    // Estimate the current spacing (pt) from the existing artboard arrangement
    // Returns fallbackPt (the preference value) when there are fewer than 2 artboards
    function computeAutoSpacingPt(fallbackPt) {
        var abs = doc.artboards;
        if (abs.length < 2) return fallbackPt;

        var baseRect = abs[0].artboardRect;
        var adjacentRect = abs[1].artboardRect;
        var primaryAxis = getPrimaryAxis(baseRect, adjacentRect);
        var pitch = Math.abs(adjacentRect[primaryAxis] - baseRect[primaryAxis]);
        var spacing = pitch - getAxisSize(baseRect, primaryAxis);
        return (spacing < 0) ? 0 : spacing;
    }

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* パネルの余白と間隔 / Panel margins and spacing */
    var PANEL_MARGINS = [16, 20, 16, 12];
    var PANEL_SPACING = 8;

    /* パネルの共通設定 / Apply shared panel layout */
    function setupPanel(panel, spacing) {
        panel.orientation = 'column';
        panel.alignChildren = ['fill', 'top'];
        panel.alignment = 'fill';
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === 'number') ? spacing : PANEL_SPACING;
    }

    /* グループの共通設定（orientation は呼び出し側で指定）/ Apply shared group layout (orientation passed in) */
    function setupGroup(group, orientation, spacing) {
        group.orientation = orientation || 'column';
        group.alignChildren = ['fill', 'top'];
        group.alignment = 'fill';
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
        var pref = app.preferences;

        // ダイアログボックスを表示 / Show dialog
        var dlg = new Window('dialog', L(LABELS.dialog.title) + ' ' + SCRIPT_VERSION);
        dlg.alignChildren = ['fill', 'top'];

        // 追加方法パネル / Add method panel
        var modePanel = dlg.add('panel', undefined, L(LABELS.panel.mode));
        setupPanel(modePanel);
        var rbDuplicate = modePanel.add('radiobutton', undefined, L(LABELS.radio.duplicate));
        var rbBlank = modePanel.add('radiobutton', undefined, L(LABELS.radio.blank));
        rbDuplicate.value = true;

        // 追加数 / Count
        var countGroup = modePanel.add('group');
        var countLabel = countGroup.add('statictext', undefined, L(LABELS.label.count) + '：');
        var etCount = countGroup.add('edittext', undefined, '1');
        etCount.characters = 4;
        changeValueByArrowKey(etCount);

        // 追加位置パネル / Insert position panel
        var posPanel = dlg.add('panel', undefined, L(LABELS.panel.position));
        setupPanel(posPanel);
        var rbNext = posPanel.add('radiobutton', undefined, L(LABELS.radio.next));
        var rbLast = posPanel.add('radiobutton', undefined, L(LABELS.radio.last));

        // 追加方向（右＝横並び / 下＝縦並び）/ Add direction (right = horizontal, down = vertical)
        var directionGroup = posPanel.add('group');
        directionGroup.add('statictext', undefined, L(LABELS.label.direction) + '：');
        var rbDirectionRight = directionGroup.add('radiobutton', undefined, L(LABELS.radio.directionRight));
        var rbDirectionDown = directionGroup.add('radiobutton', undefined, L(LABELS.radio.directionDown));
        rbDirectionRight.value = true;

        // 追加方法に応じて UI を更新 / Update the UI according to the add method:
        //  - 追加数は「空のアートボード」のときのみ有効 / Count is enabled only for blank artboards
        //  - 追加位置の既定は 複製→現在の次 / 空→末尾 / Default position: duplicate → next, blank → end
        function syncModeDependent() {
            var isBlank = rbBlank.value;
            countLabel.enabled = isBlank;
            etCount.enabled = isBlank;
            rbNext.value = !isBlank;
            rbLast.value = isBlank;
        }
        rbDuplicate.onClick = syncModeDependent;
        rbBlank.onClick = syncModeDependent;
        syncModeDependent();

        // 間隔パネル / Spacing panel
        // 適用範囲（追加分のみ / すべて）と、間隔の数値を指定
        // Choose the scope (added artboards only / all) and the spacing value
        var spacingPanel = dlg.add('panel', undefined, L(LABELS.panel.spacing));
        setupPanel(spacingPanel);
        var rbScopeNew = spacingPanel.add('radiobutton', undefined, L(LABELS.radio.scopeNew));
        var rbScopeAll = spacingPanel.add('radiobutton', undefined, L(LABELS.radio.scopeAll));
        rbScopeNew.value = true;

        // 間隔入力 / Spacing input
        // 2枚以上は既存の並びから推定した間隔、1枚以下は環境設定の値（pt）を
        // 定規単位に変換して初期表示し、編集可能にする
        // For 2+ artboards use the spacing inferred from the existing layout,
        // otherwise the preference value (pt); shown in the ruler unit and editable
        var unitInfo = getRulerUnitInfo();
        var spacingPref = pref.getRealPreference('plugin/ArtboardRearrange/ArtboardSpacing'); // rearrange (pt)
        var autoSpacingPt = computeAutoSpacingPt(spacingPref);
        var spacingInUnit = autoSpacingPt / unitInfo.factor;
        var initialSpacingText = formatSpacingValue(spacingInUnit);

        var spacingGroup = spacingPanel.add('group');
        spacingGroup.add('statictext', undefined, L(LABELS.label.spacing) + '：');
        var etSpacing = spacingGroup.add('edittext', undefined, initialSpacingText);
        etSpacing.characters = 4;
        changeValueByArrowKey(etSpacing);
        spacingGroup.add('statictext', undefined, unitInfo.label);

        addDialogKeyHandler(dlg, rbBlank, rbDuplicate, rbNext, rbLast, syncModeDependent);

        var btnGroup = dlg.add('group');
        btnGroup.alignment = ['center', 'bottom'];
        btnGroup.add('button', undefined, L(LABELS.button.cancel), { name: 'cancel' });
        btnGroup.add('button', undefined, L(LABELS.button.ok), { name: 'ok' });

        if (dlg.show() != 1) return;

        var duplicateMode = rbDuplicate.value;
        var insertNext = rbNext.value;

        // 入力値を pt に戻す。空欄・不正値は初期表示値にフォールバック
        // Convert the input back to pt; fall back to the initial value for blank/invalid input
        var spacingInput = parseFloat(etSpacing.text);
        if (isNaN(spacingInput)) spacingInput = spacingInUnit;
        var spacing = spacingInput * unitInfo.factor;

        // 手動で初期値から変更されたかどうか。変更時のみ入力値を優先
        // Whether the value was manually changed from the initial; the input wins only when changed
        var useManualSpacing = (etSpacing.text !== initialSpacingText);

        // 追加数（1以上の整数）。複製モードでは常に1枚 / Number to add (integer ≥ 1); duplicate mode always adds 1
        var count = parseInt(etCount.text, 10);
        if (isNaN(count) || count < 1) count = 1;
        if (duplicateMode) count = 1;

        // 間隔の適用範囲：true=すべてのアートボード / false=追加分のみ
        // Spacing scope: true = all artboards, false = added artboards only
        var spacingScopeAll = rbScopeAll.value;

        // 追加方向：0=右（横並び）/ 1=下（縦並び）/ Add direction: 0 = right (horizontal), 1 = down (vertical)
        var directionAxis = rbDirectionDown.value ? 1 : 0;

        insertArtboardWithShift(duplicateMode, insertNext, spacing, useManualSpacing, count, spacingScopeAll, directionAxis);

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

    function insertArtboardWithShift(duplicateMode, insertNext, spacing, useManualSpacing, count, spacingScopeAll, directionAxis) {
        var artboards = doc.artboards;
        var artboardCount = artboards.length;
        var artboardLimit = (parseFloat(app.version) >= 22) ? 1000 : 100;

        // 現在選択しているアートボードのインデックスを取得 / Get active artboard index
        var activeIdx = doc.artboards.getActiveArtboardIndex();

        if (artboardCount + count > artboardLimit) {
            alert(L(LABELS.error.artboardLimit));
            return;
        }

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

        // 主軸はダイアログで選んだ追加方向（0=右＝横 / 1=下＝縦）
        // Primary axis is the chosen add direction (0 = right/horizontal, 1 = down/vertical)
        var layoutPrimaryAxisIndex = directionAxis;
        var layoutSecondaryAxisIndex = 1 - directionAxis;

        // 既存の並びが指定方向と一致するときだけ、ピッチと列数を引き継ぐ
        // （一致しない・1枚のみのときは既定グリッド〔サイズ+間隔〕で指定方向に並べる）
        // Inherit pitch and columns only when the existing layout matches the chosen direction
        // (otherwise place along the chosen direction using the default grid of size + spacing)
        if (artboardCount >= 2) {
            var adjacentArtboardRect = artboards[1].artboardRect;
            var detectedPrimaryAxis = getPrimaryAxis(baseArtboardRect, adjacentArtboardRect);

            if (detectedPrimaryAxis === directionAxis) {
                gridStep[layoutPrimaryAxisIndex] = adjacentArtboardRect[layoutPrimaryAxisIndex] - baseArtboardRect[layoutPrimaryAxisIndex];

                for (var i = 2; i < artboardCount; i++) {
                    var iterArtboardRect = artboards[i].artboardRect;
                    if (baseArtboardRect[layoutSecondaryAxisIndex] != iterArtboardRect[layoutSecondaryAxisIndex]) {
                        gridStep[layoutSecondaryAxisIndex] = iterArtboardRect[layoutSecondaryAxisIndex] - baseArtboardRect[layoutSecondaryAxisIndex];
                        columns = i;
                        break;
                    }
                }
            }

            // 「すべてのアートボード」かつ手動間隔のときだけ、並び全体を入力値で再配置するため
            // グリッドの移動量を作り直す（追加分のみのときは既存のピッチを保持）
            // Only when scope = all AND the spacing was changed manually, rebuild the grid step
            // from the input to re-flow the whole arrangement (scope = added-only keeps the existing pitch)
            if (useManualSpacing && spacingScopeAll) {
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
        if (artboardCount + count > columns * rows) {
            alert(L(LABELS.error.noSpace));
            return;
        }

        // 挿入位置を決定 / Determine insertion index
        var insertIdx = insertNext ? activeIdx + 1 : artboardCount;

        // 間隔の適用範囲ごとの移動量 / Movement amounts depending on the spacing scope
        //  - すべて：手動間隔のとき先頭も含めて新グリッドへ再配置 / all: re-flow everything (head too)
        //  - 追加分のみ：既存はそのまま、新規だけを (指定間隔 − 既存の隙間) ぶん主軸方向へずらす
        //    added-only: keep existing as-is, nudge only the new artboards by (spacing − existing gap)
        var primaryDir = (gridStep[layoutPrimaryAxisIndex] < 0) ? -1 : 1;
        var primaryArtboardSize = getAxisSize(baseArtboardRect, layoutPrimaryAxisIndex);
        var autoGapPrimary = Math.abs(gridStep[layoutPrimaryAxisIndex]) - primaryArtboardSize;
        if (autoGapPrimary < 0) autoGapPrimary = 0;
        var newOnlyDelta = (useManualSpacing && !spacingScopeAll) ? (spacing - autoGapPrimary) : 0;
        var reflowHead = (useManualSpacing && spacingScopeAll);
        var tailPrimaryOffset = primaryDir * count * newOnlyDelta;

        // グリッド上のインデックス n の位置を計算 / Calculate grid position for index n
        function getArtboardGridPosition(n) {
            var p = [];
            p[layoutPrimaryAxisIndex] = (n % columns) * gridStep[layoutPrimaryAxisIndex];
            p[layoutSecondaryAxisIndex] = Math.floor(n / columns) * gridStep[layoutSecondaryAxisIndex];
            return [baseArtboardRect[0] + p[0], baseArtboardRect[1] + p[1]];
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
                var gb = item.geometricBounds;
                var cx = (gb[0] + gb[2]) / 2;
                var cy = (gb[1] + gb[3]) / 2;
                if (cx >= artboardRect[0] && cx <= artboardRect[2] &&
                    cy <= artboardRect[1] && cy >= artboardRect[3]) {
                    items.push(item);
                }
            }
            return items;
        }

        // 既存アートボードを最終位置へ移動し、アートワークも一緒に運ぶ。
        // 挿入位置以降は count セル分だけ後ろへ（追加分のみモードでは tailPrimaryOffset を加算して
        // 新規の間隔ぶんさらにずらす）。reflowHead=true（すべてモード）のときは先頭も新グリッドへ再配置。
        // アイテムの帰属は移動前の位置でまとめて取得（スナップショット）してから動かすので、
        // 移動順による取り違えが起きない。
        // Move existing artboards to their final positions, carrying their artwork.
        // Artboards at/after the insert point move back by `count` cells (plus tailPrimaryOffset
        // in added-only mode, to make room for the new spacing). When reflowHead is true (scope =
        // all) the head is also re-flowed onto the new grid. Artwork assignment is snapshotted
        // from the pre-move rects, so move order can't misassign items.
        function relayoutExistingArtboards(fromIdx, reflowHead, tailPrimaryOffset) {
            var moves = [];
            for (var si = 0; si < artboardCount; si++) {
                if (!reflowHead && si < fromIdx) continue;
                var targetIdx = (si >= fromIdx) ? si + count : si;
                var rect = artboards[si].artboardRect;
                var pos = getArtboardGridPosition(targetIdx);
                if (si >= fromIdx) pos[layoutPrimaryAxisIndex] += tailPrimaryOffset;
                moves.push({
                    index: si,
                    dx: pos[0] - rect[0],
                    dy: pos[1] - rect[1],
                    items: getItemsAssignedToArtboard(rect)
                });
            }

            for (var mv = 0; mv < moves.length; mv++) {
                var move = moves[mv];
                for (var it = 0; it < move.items.length; it++) {
                    move.items[it].translate(move.dx, move.dy);
                }
                var r = artboards[move.index].artboardRect;
                artboards[move.index].artboardRect = [r[0] + move.dx, r[1] + move.dy, r[2] + move.dx, r[3] + move.dy];
            }
        }

        // 末尾に追加された count 枚を、挿入位置 fromIdx..fromIdx+count-1 へ並べ替え。
        // 既存の後続（fromIdx..originalCount-1）は count 枚分だけ後ろへ送る。
        // Move the `count` just-appended artboards into slots fromIdx..fromIdx+count-1,
        // pushing the existing trailing artboards (fromIdx..originalCount-1) back by `count`.
        function reorderAppendedArtboards(fromIdx, originalCount) {
            // 追加分の枠と名前を退避（後続シフトで上書きされる前に）/ Save the appended rects and names first
            var appended = [];
            for (var a = 0; a < count; a++) {
                var src = artboards[originalCount + a];
                appended.push({ rect: src.artboardRect, name: src.name });
            }
            // 後続を count 枚分だけ後ろへ（高位から処理して上書き衝突を回避）
            // Shift trailing artboards back by `count` (high → low to avoid clobbering)
            for (var j = originalCount - 1; j >= fromIdx; j--) {
                artboards[j + count].artboardRect = artboards[j].artboardRect;
                artboards[j + count].name = artboards[j].name;
            }
            // 追加分を挿入位置へ配置 / Place the appended artboards at the insert position
            for (var b = 0; b < count; b++) {
                artboards[fromIdx + b].artboardRect = appended[b].rect;
                artboards[fromIdx + b].name = appended[b].name;
            }
        }

        // 複製モード：srcIdx の内容を dstIdx へコピー / Duplicate mode: copy contents from srcIdx to dstIdx
        function duplicateArtboardContents(srcIdx, dstIdx) {
            var srcRect = artboards[srcIdx].artboardRect;
            var dstRect = artboards[dstIdx].artboardRect;
            var offsetX = dstRect[0] - srcRect[0];
            var offsetY = dstRect[1] - srcRect[1];

            var itemsToCopy = getItemsAssignedToArtboard(srcRect);
            for (var m = 0; m < itemsToCopy.length; m++) {
                var copied = itemsToCopy[m].duplicate();
                copied.translate(offsetX, offsetY);
            }
        }

        // 既存アートボードを最終位置へ移動して挿入スペースを空ける
        // Move existing artboards to their final positions to free the insert space
        relayoutExistingArtboards(insertIdx, reflowHead, tailPrimaryOffset);

        // 指定数の新規アートボードを追加（いずれも末尾に追加される）
        // 追加分のみモードでは、各新規を指定間隔ぶん主軸方向へずらす
        // Add the requested number of new artboards (each appended at the end);
        // in added-only mode each one is nudged along the primary axis by the new spacing
        for (var c = 0; c < count; c++) {
            var pos = getArtboardGridPosition(insertIdx + c);
            pos[layoutPrimaryAxisIndex] += primaryDir * (c + 1) * newOnlyDelta;
            artboards.add([pos[0], pos[1], pos[0] + artboardWidth, pos[1] + artboardHeight]);
        }

        // パネル上の順序を挿入位置へ並べ替え / Reorder in the panel to the insert position
        reorderAppendedArtboards(insertIdx, artboardCount);

        // 新規アートボードの命名と、複製モードなら内容のコピー
        // Name the new artboards and, in duplicate mode, copy the contents
        var sourceArtboardName = artboards[activeIdx].name;
        var suffix = duplicateMode ? SUFFIX_COPY : SUFFIX_BLANK;
        for (var n = 0; n < count; n++) {
            var newIdx = insertIdx + n;
            // 複数追加時は連番を付与 / Append a sequence number when adding multiple
            artboards[newIdx].name = sourceArtboardName + suffix + (count > 1 ? ' ' + (n + 1) : '');
            if (duplicateMode) {
                duplicateArtboardContents(activeIdx, newIdx);
            }
        }

        // 最初に追加したアートボードを選択・アクティブにする
        // Select and activate the first newly added artboard
        doc.selection = null;
        doc.artboards.setActiveArtboardIndex(insertIdx);
        app.redraw();
    }

    // =========================================
    // キー入力で選択切り替え / Change selections with keyboard shortcuts
    // =========================================
    function addDialogKeyHandler(dialog, blankRadio, duplicateRadio, nextRadio, endRadio, onModeChange) {
        dialog.addEventListener('keydown', function (event) {
            var keyName = event.keyName;
            if (!keyName) return;

            keyName = String(keyName).toUpperCase();

            if (keyName === 'B') {
                blankRadio.value = true;
                duplicateRadio.value = false;
                if (onModeChange) onModeChange();
                event.preventDefault();
            } else if (keyName === 'D') {
                blankRadio.value = false;
                duplicateRadio.value = true;
                if (onModeChange) onModeChange();
                event.preventDefault();
            } else if (keyName === 'N') {
                nextRadio.value = true;
                endRadio.value = false;
                event.preventDefault();
            } else if (keyName === 'E') {
                nextRadio.value = false;
                endRadio.value = true;
                event.preventDefault();
            }
        });
    }

}());