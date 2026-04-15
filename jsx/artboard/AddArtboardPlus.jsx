#target "illustrator"

    /*
    AddArtboardPlus.jsx
    
    概要 / Overview
既存のアートボードの並び（行・列）を解析し、その規則を維持したまま新しいアートボードを挿入します。
「空のアートボード」または「現在のアートボードの内容を複製」を選択でき、
追加位置は「現在の次」または「末尾」から指定できます。
挿入時は後続のアートボードとアートワークを安全にシフトし、既存レイアウトを崩さずに追加します。
追加されたアートボード名は、空のアートボードでは「元のアートボード名_blank」、複製では「元のアートボード名_copy」として設定されます。
実行後は追加されたアートボードをアクティブにします。

キーボードショートカット（B / D / N / E）による操作にも対応しています。
Analyzes the current artboard layout (rows/columns) and inserts a new artboard while preserving that structure.
You can create a blank artboard or duplicate the current artboard, and choose to insert it after the current one or at the end.
During insertion, subsequent artboards and artwork are safely shifted so the existing layout remains intact.
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
    
    Updated: 2026-04-15
    */

    (function () {

        app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

        var SCRIPT_VERSION = 'v1.0';
        var SUFFIX_BLANK = '_blank';
        var SUFFIX_COPY = '_copy';

        // =========================================
        // バージョンとローカライズ / Version and localization
        // =========================================

        function getCurrentLang() {
            return ($.locale.indexOf('ja') === 0) ? 'ja' : 'en';
        }
        var lang = getCurrentLang();

        /* 日英ラベル定義 / Japanese-English label definitions */
        var LABELS = {
            dialogTitle: { ja: 'アートボードを追加', en: 'Add Artboard' },
            modePanel: { ja: '追加方法', en: 'Add Method' },
            modeBlank: { ja: '空のアートボード', en: 'Blank Artboard' },
            modeDuplicate: { ja: '現在のアートボードを複製', en: 'Duplicate Current Artboard' },
            positionPanel: { ja: '追加位置', en: 'Insert Position' },
            positionNext: { ja: '現在のアートボードの次', en: 'After Current Artboard' },
            positionLast: { ja: '末尾', en: 'At End' },
            cancel: { ja: 'キャンセル', en: 'Cancel' },
            ok: { ja: 'OK', en: 'OK' },
            errorNoDocument: {
                ja: 'ドキュメントがないため実行できません。',
                en: 'Cannot run because no document is open.'
            },
            errorException: {
                ja: 'エラーが発生したため、処理を実行できませんでした。\nエラー内容：',
                en: 'Processing could not be completed because an error occurred.\nError:'
            },
            errorArtboardLimit: {
                ja: 'アートボードの最大作成可能数を超えています。',
                en: 'The maximum number of artboards would be exceeded.'
            },
            errorNoSpace: {
                ja: 'アートボードを作成する十分なスペースがありません。',
                en: 'There is not enough space to create the artboard.'
            }
        };

        function L(key) {
            return (LABELS[key] && LABELS[key][lang]) || key;
        }

        // =========================================
        // メイン処理 / Main process
        // =========================================

        if (app.documents.length == 0) {
            alert(L('errorNoDocument'));
        } else {
            try {
                var doc = app.activeDocument;
                var pref = app.preferences;

                // ダイアログボックスを表示 / Show dialog
                var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
                dlg.alignChildren = ['fill', 'top'];

                // 追加方法パネル / Add method panel
                var modePanel = dlg.add('panel', undefined, L('modePanel'));
                modePanel.alignment = ['fill', 'top'];
                modePanel.alignChildren = ['left', 'top'];
                modePanel.margins = [15, 20, 15, 10];
                var rbBlank = modePanel.add('radiobutton', undefined, L('modeBlank'));
                var rbDuplicate = modePanel.add('radiobutton', undefined, L('modeDuplicate'));
                rbDuplicate.value = true;

                // 追加位置パネル / Insert position panel
                var posPanel = dlg.add('panel', undefined, L('positionPanel'));
                posPanel.alignment = ['fill', 'top'];
                posPanel.alignChildren = ['left', 'top'];
                posPanel.margins = [15, 20, 15, 10];
                var rbNext = posPanel.add('radiobutton', undefined, L('positionNext'));
                var rbLast = posPanel.add('radiobutton', undefined, L('positionLast'));
                rbNext.value = true;

                addDialogKeyHandler(dlg, rbBlank, rbDuplicate, rbNext, rbLast);

                var btnGroup = dlg.add('group');
                btnGroup.alignment = ['center', 'bottom'];
                btnGroup.add('button', undefined, L('cancel'), { name: 'cancel' });
                btnGroup.add('button', undefined, L('ok'), { name: 'ok' });

                if (dlg.show() != 1) return;

                var duplicateMode = rbDuplicate.value;
                var insertNext = rbNext.value;

                insertArtboardWithShift(duplicateMode, insertNext);

            } catch (e) {
                alert(L('errorException') + e);
            }
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

        function insertArtboardWithShift(duplicateMode, insertNext) {
            var artboards = doc.artboards;
            var artboardCount = artboards.length;
            var artboardLimit = (parseFloat(app.version) >= 22) ? 1000 : 100;

            // 現在選択しているアートボードのインデックスを取得 / Get active artboard index
            var activeIdx = doc.artboards.getActiveArtboardIndex();

            if (artboardCount + 1 > artboardLimit) {
                alert(L('errorArtboardLimit'));
                return;
            }

            var baseArtboardRect = artboards[0].artboardRect;
            var artboardWidth = baseArtboardRect[2] - baseArtboardRect[0];
            var artboardHeight = baseArtboardRect[3] - baseArtboardRect[1];

            // var spacing = pref.getRealPreference('artnewdialog/artboardSpacing'); // new dialog
            var spacing = pref.getRealPreference('plugin/ArtboardRearrange/ArtboardSpacing'); // rearrange
            // グリッド1セルあたりの移動量（幅+間隔、高さは上→下なので負方向に間隔を引く）
            // Grid step per cell (width + spacing, height subtracts spacing because Y axis is top-down)
            var gridStep = [
                artboardWidth + spacing,
                artboardHeight - spacing
            ];
            var columns = 0;
            var rows = 0;
            var layoutPrimaryAxisIndex = 0;
            var layoutSecondaryAxisIndex = 1;

            // アートボードの並び方向から主軸（横 or 縦）を推定
            // Determine the primary layout axis (horizontal or vertical) from the artboard arrangement
            if (artboardCount >= 2) {
                var adjacentArtboardRect = artboards[1].artboardRect;
                layoutPrimaryAxisIndex = +(baseArtboardRect[0] == adjacentArtboardRect[0]);
                layoutSecondaryAxisIndex = 1 - layoutPrimaryAxisIndex;
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
            if (artboardCount + 1 > columns * rows) {
                alert(L('errorNoSpace'));
                return;
            }

            // 挿入位置を決定 / Determine insertion index
            var insertIdx = insertNext ? activeIdx + 1 : artboardCount;

            // グリッド上のインデックス n の位置を計算 / Calculate grid position for index n
            function getArtboardGridPosition(n) {
                var p = [];
                p[layoutPrimaryAxisIndex] = (n % columns) * gridStep[layoutPrimaryAxisIndex];
                p[layoutSecondaryAxisIndex] = Math.floor(n / columns) * gridStep[layoutSecondaryAxisIndex];
                return [baseArtboardRect[0] + p[0], baseArtboardRect[1] + p[1]];
            }

            // アートボード上のアイテムを収集 / Collect items on an artboard
            function getItemsAssignedToArtboard(artboardRect) {
                var items = [];
                for (var k = 0; k < doc.pageItems.length; k++) {
                    var item = doc.pageItems[k];
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

            // =========================================
            // 挿入位置以降のアートボード＋アートワークを物理的にずらす
            // Shift artboards + artwork from insertIdx onwards
            // 末尾から処理して重なりを回避
            // Process from last to first to avoid overlap
            // =========================================
            for (var si = artboardCount - 1; si >= insertIdx; si--) {
                var currentShiftTargetRect = artboards[si].artboardRect;
                var targetGridPos = getArtboardGridPosition(si + 1);
                var dx = targetGridPos[0] - currentShiftTargetRect[0];
                var dy = targetGridPos[1] - currentShiftTargetRect[1];

                // アートワークを移動 / Move artwork
                var items = getItemsAssignedToArtboard(currentShiftTargetRect);
                for (var mi = 0; mi < items.length; mi++) {
                    items[mi].translate(dx, dy);
                }

                // アートボードの枠を移動 / Move artboard rect
                artboards[si].artboardRect = [
                    currentShiftTargetRect[0] + dx,
                    currentShiftTargetRect[1] + dy,
                    currentShiftTargetRect[2] + dx,
                    currentShiftTargetRect[3] + dy
                ];
            }

            // =========================================
            // 空いた位置に新規アートボードを追加
            // Add new artboard at the freed position
            // =========================================
            var insertGridPos = getArtboardGridPosition(insertIdx);
            artboards.add([
                insertGridPos[0],
                insertGridPos[1],
                insertGridPos[0] + artboardWidth,
                insertGridPos[1] + artboardHeight
            ]);

            // artboards.add() は末尾に追加されるので、パネル上の順序を入れ替え
            // artboards.add() appends at end, so reorder in the panel
            var lastIdx = artboards.length - 1;
            if (insertIdx < lastIdx) {
                var appendedArtboardRect = artboards[lastIdx].artboardRect;
                var appendedArtboardName = artboards[lastIdx].name;
                for (var j = lastIdx; j > insertIdx; j--) {
                    artboards[j].artboardRect = artboards[j - 1].artboardRect;
                    artboards[j].name = artboards[j - 1].name;
                }
                artboards[insertIdx].artboardRect = appendedArtboardRect;
                artboards[insertIdx].name = appendedArtboardName;
            }

            // 新規アートボードの名前を設定 / Name the new artboard
            var sourceArtboardName = artboards[activeIdx].name;
            var suffix = duplicateMode ? SUFFIX_COPY : SUFFIX_BLANK;
            artboards[insertIdx].name = sourceArtboardName + suffix;

            // 新規アートボードをアクティブに設定 / Activate the new artboard
            doc.artboards.setActiveArtboardIndex(insertIdx);

            // 複製モード：コンテンツをコピー / Duplicate mode: copy contents
            if (duplicateMode) {
                var sourceArtboardRect = artboards[activeIdx].artboardRect;
                var destinationArtboardRect = artboards[insertIdx].artboardRect;
                var offsetX = destinationArtboardRect[0] - sourceArtboardRect[0];
                var offsetY = destinationArtboardRect[1] - sourceArtboardRect[1];

                var itemsToCopy = getItemsAssignedToArtboard(sourceArtboardRect);
                for (var m = 0; m < itemsToCopy.length; m++) {
                    var copied = itemsToCopy[m].duplicate();
                    copied.translate(offsetX, offsetY);
                }
            }

            // 追加されたアートボードのみを選択状態にする
            // Select only the newly added artboard
            doc.selection = null;
            doc.artboards.setActiveArtboardIndex(insertIdx);
            app.redraw();
        }

        // =========================================
        // キー入力で選択切り替え / Change selections with keyboard shortcuts
        // =========================================
        function addDialogKeyHandler(dialog, blankRadio, duplicateRadio, nextRadio, endRadio) {
            dialog.addEventListener('keydown', function (event) {
                var keyName = event.keyName;
                if (!keyName) return;

                keyName = String(keyName).toUpperCase();

                if (keyName === 'B') {
                    blankRadio.value = true;
                    duplicateRadio.value = false;
                    event.preventDefault();
                } else if (keyName === 'D') {
                    blankRadio.value = false;
                    duplicateRadio.value = true;
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