#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アクティブなアートボードだけを残すか、空のアートボードをまとめて削除します。
アクティブなアートボードだけを残す場合は、その外側にあるオブジェクト（ガイドを含む）も削除できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/KeepActiveArtboardOnly.md

### Overview

Keeps only the active artboard, or removes every empty artboard in one pass.
When the active artboard is kept, the objects outside it — guides included — can be deleted as well.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/KeepActiveArtboardOnly.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "KeepActiveArtboardOnly";       /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-15";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/KeepActiveArtboardOnly.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/KeepActiveArtboardOnly.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* アクティブなアートボード以外を削除するとき、外側のオブジェクトも削除するか / Delete the objects outside the kept artboard */
    var DEFAULT_DELETE_OUTSIDE_OBJECTS = true;

    /* 空のアートボードを探すとき、非表示のレイヤー・オブジェクトを無視するか / Ignore hidden layers and objects while looking for empty artboards */
    var DEFAULT_IGNORE_HIDDEN = true;

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_MARGINS = 16;                    /* ダイアログの余白 */
    var DIALOG_SPACING = 12;                    /* ダイアログ内の要素間隔 */
    var PANEL_MARGINS = [15, 20, 15, 10];       /* パネル余白 [左,上,右,下] */
    var PANEL_SPACING = 6;                      /* パネル内の要素間隔 */
    var OPTION_INDENT_MARGINS = [18, 0, 0, 2];  /* ラジオ直下のオプションのインデント */
    var COUNT_TEXT_WIDTH = 220;                 /* 件数表示の幅（後から伸びないため確保） */
    var BUTTON_WIDTH = 80;                      /* ボタン幅 */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * UI言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "アートボードの削除", en: "Remove Artboards" }
        },
        panel: {
            mode: { ja: "削除するアートボード", en: "Artboards to remove" }
        },
        radio: {
            keepActive: { ja: "アクティブなアートボード以外", en: "All but the active artboard" },
            empty:      { ja: "空のアートボード", en: "Empty artboards" }
        },
        checkbox: {
            deleteOutsideObjects: { ja: "アートボード外のオブジェクトも削除", en: "Also delete the objects outside it" },
            ignoreHidden:         { ja: "非表示レイヤー、非表示オブジェクトは無視", en: "Ignore hidden layers and hidden objects" }
        },
        fieldLabel: {
            removalTarget: { ja: "削除対象アートボード数", en: "Artboards to remove" }
        },
        tooltip: {
            keepActive: {
                ja: "アクティブなアートボードだけを残し、ほかのアートボードを削除します。",
                en: "Keeps just the active artboard and removes the others."
            },
            empty: {
                ja: "オブジェクトが載っていないアートボードをまとめて削除します。",
                en: "Removes every artboard with nothing on it."
            },
            deleteOutsideObjects: {
                ja: "残したアートボードの外側にあるオブジェクトとオブジェクトガイドも削除します。",
                en: "Also deletes the objects and object guides that sit outside the artboard you keep."
            },
            ignoreHidden: {
                ja: "非表示のレイヤーやオブジェクトしか載っていないアートボードも「空」とみなします。オフにすると、見えていなくても中身があれば残します。",
                en: "Treats an artboard holding only hidden layers or objects as empty. Off keeps it as long as anything is on it."
            },
            singleArtboard: {
                ja: "アートボードが1枚しかないため削除できません。",
                en: "Only one artboard exists; cannot remove."
            }
        },
        button: {
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." }
        }
    };

    /**
     * ラベルを取得する
     * @param {string} labelPath - "dialog.title" のようなドット区切りのキー
     * @returns {string} 現在のUI言語のラベル
     */
    function getLabel(labelPath) {
        var pathKeys = String(labelPath).split('.');
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return (labelNode[uiLang] != null) ? labelNode[uiLang] : labelPath;
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelPath - ドット区切りのキー
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === 'ja' ? '：' : ':');
    }

    // =========================================
    // 重なり判定 / Hit testing
    // =========================================

    /* 共線とみなす外積の許容値 / Cross product below this counts as collinear */
    var COLLINEAR_EPSILON = 1e-6;

    /**
     * 点が矩形の内側（境界を含む）にあるか判定する
     * @param {number} x - X座標
     * @param {number} y - Y座標
     * @param {number[]} rect - 矩形 [left, top, right, bottom]
     * @returns {boolean} 内側にあれば true
     */
    function isPointInRect(x, y, rect) {
        return (x >= rect[0] && x <= rect[2] && y <= rect[1] && y >= rect[3]);
    }

    /**
     * 2Dベクトルの外積を返す
     * @param {number} ax - ベクトルAのX成分
     * @param {number} ay - ベクトルAのY成分
     * @param {number} bx - ベクトルBのX成分
     * @param {number} by - ベクトルBのY成分
     * @returns {number} 外積
     */
    function cross2D(ax, ay, bx, by) {
        return ax * by - ay * bx;
    }

    /**
     * 線分同士が交差するか判定する（共線で重なる場合も true）
     * @param {number} p1x - 線分Pの始点X
     * @param {number} p1y - 線分Pの始点Y
     * @param {number} p2x - 線分Pの終点X
     * @param {number} p2y - 線分Pの終点Y
     * @param {number} q1x - 線分Qの始点X
     * @param {number} q1y - 線分Qの始点Y
     * @param {number} q2x - 線分Qの終点X
     * @param {number} q2y - 線分Qの終点Y
     * @returns {boolean} 交差していれば true
     */
    function segmentsIntersect(p1x, p1y, p2x, p2y, q1x, q1y, q2x, q2y) {
        var pDx = p2x - p1x,
            pDy = p2y - p1y;
        var qDx = q2x - q1x,
            qDy = q2y - q1y;
        var denominator = cross2D(pDx, pDy, qDx, qDy);
        var p1ToQ1Dx = q1x - p1x,
            p1ToQ1Dy = q1y - p1y;

        if (denominator === 0) {
            /* 平行。共線でなければ交差しない / Parallel: no intersection unless collinear */
            if (Math.abs(cross2D(p1ToQ1Dx, p1ToQ1Dy, pDx, pDy)) > COLLINEAR_EPSILON) return false;
            /* 共線なら各軸への投影で重なりを見る / Collinear: check the overlap on each axis */
            var overlapX = !(Math.max(p1x, p2x) < Math.min(q1x, q2x) || Math.max(q1x, q2x) < Math.min(p1x, p2x));
            var overlapY = !(Math.max(p1y, p2y) < Math.min(q1y, q2y) || Math.max(q1y, q2y) < Math.min(p1y, p2y));
            return overlapX && overlapY;
        }

        var tParam = cross2D(p1ToQ1Dx, p1ToQ1Dy, qDx, qDy) / denominator;
        var uParam = cross2D(p1ToQ1Dx, p1ToQ1Dy, pDx, pDy) / denominator;
        return (tParam >= 0 && tParam <= 1 && uParam >= 0 && uParam <= 1);
    }

    /**
     * 線分が矩形と交差するか判定する
     * @param {number} x1 - 線分の始点X
     * @param {number} y1 - 線分の始点Y
     * @param {number} x2 - 線分の終点X
     * @param {number} y2 - 線分の終点Y
     * @param {number[]} rect - 矩形 [left, top, right, bottom]
     * @returns {boolean} 交差していれば true
     */
    function segmentIntersectsRect(x1, y1, x2, y2, rect) {
        /* 端点のどちらかが矩形内なら交差 / Either endpoint inside means a hit */
        if (isPointInRect(x1, y1, rect) || isPointInRect(x2, y2, rect)) return true;

        var left = rect[0],
            top = rect[1],
            right = rect[2],
            bottom = rect[3];

        return segmentsIntersect(x1, y1, x2, y2, left, top, right, top) ||
            segmentsIntersect(x1, y1, x2, y2, left, bottom, right, bottom) ||
            segmentsIntersect(x1, y1, x2, y2, left, top, left, bottom) ||
            segmentsIntersect(x1, y1, x2, y2, right, top, right, bottom);
    }

    /**
     * パスがアートボードと重なるか判定する（境界ボックスではなくパスの実体で判定）
     * @param {PathItem} pathItem - 判定するパス
     * @param {number[]} artboardRect - アートボードの矩形 [left, top, right, bottom]
     * @returns {boolean} 重なっていれば true
     */
    function pathIntersectsArtboard(pathItem, artboardRect) {
        try {
            var pathPoints = pathItem.pathPoints;
            if (!pathPoints || pathPoints.length === 0) return false;

            for (var i = 0; i < pathPoints.length; i++) {
                /* 開パスでは最後のアンカーから始点へは結ばない / Open paths do not wrap back to the start */
                if (!pathItem.closed && i === pathPoints.length - 1) break;
                var currentAnchor = pathPoints[i].anchor;
                var nextAnchor = pathPoints[(i + 1) % pathPoints.length].anchor;
                if (segmentIntersectsRect(currentAnchor[0], currentAnchor[1], nextAnchor[0], nextAnchor[1], artboardRect)) return true;
            }
            return false;
        } catch (e) {
            return false;
        }
    }

    /**
     * ページアイテムがアートボードと重なるか判定する
     * パス・複合パスはパスの実体で、それ以外は境界ボックスで判定する
     * @param {PageItem} pageItem - 判定するアイテム
     * @param {number[]} artboardRect - アートボードの矩形 [left, top, right, bottom]
     * @returns {boolean} 重なっていれば true
     */
    function itemIntersectsArtboard(pageItem, artboardRect) {
        if (pageItem.typename === 'CompoundPathItem') {
            var subPaths = pageItem.pathItems;
            for (var i = 0; i < subPaths.length; i++) {
                if (pathIntersectsArtboard(subPaths[i], artboardRect)) return true;
            }
            return false;
        }
        if (pageItem.typename === 'PathItem') {
            return pathIntersectsArtboard(pageItem, artboardRect);
        }
        /* その他は境界ボックスでフォールバック / Fall back to the bounding box */
        return rectsOverlap(artboardRect, pageItem.geometricBounds);
    }

    /**
     * アートボードと境界ボックスが重なるか判定する
     * @param {number[]} artboardRect - アートボードの矩形 [left, top, right, bottom]
     * @param {number[]} itemBounds - アイテムの境界 [left, top, right, bottom]
     * @returns {boolean} 重なっていれば true
     */
    function rectsOverlap(artboardRect, itemBounds) {
        return (
            itemBounds[0] < artboardRect[2] &&
            itemBounds[2] > artboardRect[0] &&
            itemBounds[1] > artboardRect[3] &&
            itemBounds[3] < artboardRect[1]
        );
    }

    // =========================================
    // ロック・表示状態の一時解除 / Relaxing locks and visibility
    // =========================================

    /**
     * レイヤーの状態を控えてロック・非表示・テンプレートを解除する（サブレイヤーとグループも再帰）
     * @param {Layer} layer - 対象レイヤー
     * @param {object[]} stateLog - 状態の控え
     * @returns {void}
     */
    function captureAndUnlockLayer(layer, stateLog) {
        stateLog.push({
            type: 'Layer',
            ref: layer,
            locked: layer.locked,
            visible: layer.visible,
            template: layer.template
        });

        layer.locked = false;
        layer.visible = true;
        if (layer.template === true) layer.template = false;

        for (var i = 0; i < layer.layers.length; i++) {
            captureAndUnlockLayer(layer.layers[i], stateLog);
        }
        for (var j = 0; j < layer.groupItems.length; j++) {
            captureAndUnlockGroup(layer.groupItems[j], stateLog);
        }
        captureAndUnlockItems(layer, stateLog);
    }

    /**
     * グループの状態を控えてロック・非表示を解除する（ネストしたグループも再帰）
     * @param {GroupItem} groupItem - 対象グループ
     * @param {object[]} stateLog - 状態の控え
     * @returns {void}
     */
    function captureAndUnlockGroup(groupItem, stateLog) {
        stateLog.push({
            type: 'Group',
            ref: groupItem,
            locked: groupItem.locked,
            visible: groupItem.visible
        });

        groupItem.locked = false;
        groupItem.visible = true;

        for (var i = 0; i < groupItem.groupItems.length; i++) {
            captureAndUnlockGroup(groupItem.groupItems[i], stateLog);
        }
        captureAndUnlockItems(groupItem, stateLog);
    }

    /**
     * コンテナ配下のページアイテムの状態を控えてロック・非表示を解除する
     * @param {Layer|GroupItem} container - レイヤーまたはグループ
     * @param {object[]} stateLog - 状態の控え
     * @returns {void}
     */
    function captureAndUnlockItems(container, stateLog) {
        try {
            var pageItems = container.pageItems;
            for (var i = 0; i < pageItems.length; i++) {
                var pageItem = pageItems[i];
                stateLog.push({
                    type: 'Item',
                    ref: pageItem,
                    locked: pageItem.locked,
                    hidden: pageItem.hidden
                });
                if (pageItem.locked) pageItem.locked = false;
                if (pageItem.hidden) pageItem.hidden = false;
            }
        } catch (e) { }
    }

    /**
     * ドキュメント全体のロック・表示状態を控えて一時解除する
     * @param {Document} doc - 対象ドキュメント
     * @returns {object[]} 復元用の状態の控え
     */
    function captureAndUnlockStructure(doc) {
        var stateLog = [];
        for (var i = 0; i < doc.layers.length; i++) {
            captureAndUnlockLayer(doc.layers[i], stateLog);
        }
        return stateLog;
    }

    /**
     * 控えた状態を復元する（下位から順に戻す）
     * @param {object[]} savedStates - captureAndUnlockStructure() が返した控え
     * @returns {void}
     */
    function restoreStructure(savedStates) {
        for (var i = savedStates.length - 1; i >= 0; i--) {
            var savedState = savedStates[i];
            /* 削除済みのオブジェクトは参照できない / deleted objects can no longer be restored */
            try {
                savedState.ref.locked = savedState.locked;
                if (savedState.type === 'Item') {
                    savedState.ref.hidden = savedState.hidden;
                } else {
                    savedState.ref.visible = savedState.visible;
                }
                if (savedState.type === 'Layer') savedState.ref.template = savedState.template;
            } catch (e) { }
        }
    }

    // =========================================
    // 空アートボードの判定 / Empty artboard detection
    // =========================================

    /**
     * オブジェクトと祖先がすべて表示状態か判定する
     * @param {PageItem} pageItem - 対象アイテム
     * @returns {boolean} 表示されていれば true
     */
    function isItemVisible(pageItem) {
        try {
            var ancestorNode = pageItem;
            while (ancestorNode && ancestorNode.typename !== "Document") {
                if (ancestorNode.typename === "Layer") {
                    if (ancestorNode.visible === false) return false;
                } else if (ancestorNode.hidden) {
                    return false;
                }
                ancestorNode = ancestorNode.parent;
            }
        } catch (err) {
            return false;
        }
        return true;
    }

    /**
     * アートボード上に占有オブジェクトがあるか判定する
     * @param {Document} doc - 対象ドキュメント
     * @param {number[]} artboardRect - アートボードの矩形 [left, top, right, bottom]
     * @param {boolean} ignoreHidden - 非表示のレイヤー・オブジェクトを無視するか
     * @returns {boolean} 占有オブジェクトがあれば true
     */
    function hasOccupantOnArtboard(doc, artboardRect, ignoreHidden) {
        var pageItems = doc.pageItems;
        for (var i = 0; i < pageItems.length; i++) {
            var pageItem = pageItems[i];
            var visible = isItemVisible(pageItem);
            if (ignoreHidden && !visible) continue;
            try {
                var bounds = visible ? pageItem.visibleBounds : pageItem.geometricBounds;
                if (rectsOverlap(artboardRect, bounds)) return true;
            } catch (err) {
                /* 境界を取得できないアイテムは無視 / Skip items without bounds */
            }
        }
        return false;
    }

    /**
     * 空アートボードのインデックス一覧を返す
     * @param {Document} doc - 対象ドキュメント
     * @param {boolean} ignoreHidden - 非表示のレイヤー・オブジェクトを無視するか
     * @returns {number[]} 空アートボードのインデックス
     */
    function findEmptyArtboardIndices(doc, ignoreHidden) {
        var indices = [];
        for (var i = 0; i < doc.artboards.length; i++) {
            if (!hasOccupantOnArtboard(doc, doc.artboards[i].artboardRect, ignoreHidden)) {
                indices.push(i);
            }
        }
        return indices;
    }

    // =========================================
    // 削除処理 / Removal
    // =========================================

    /**
     * 指定インデックスのアートボードを削除する（最後の1枚は残す）
     * @param {Document} doc - 対象ドキュメント
     * @param {number[]} indices - 削除するアートボードのインデックス
     * @returns {number} 削除した枚数
     */
    function removeArtboardsByIndices(doc, indices) {
        var artboards = doc.artboards;
        var removedCount = 0;
        /* 末尾から削除してインデックスのずれを避ける / Remove from the end so the indices stay valid */
        for (var i = indices.length - 1; i >= 0; i--) {
            if (artboards.length <= 1) break;
            try {
                artboards[indices[i]].remove();
                removedCount++;
            } catch (err) { }
        }
        return removedCount;
    }

    /**
     * アクティブなアートボードと重ならないオブジェクトガイドを削除する（ルーラーガイドは対象外）
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function removeGuidesOutsideArtboard(doc) {
        try {
            var artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;

            var guidesToRemove = [];
            for (var i = doc.pageItems.length - 1; i >= 0; i--) {
                var pageItem = doc.pageItems[i];
                if (pageItem.guides !== true) continue;
                if (!itemIntersectsArtboard(pageItem, artboardRect)) guidesToRemove.push(pageItem);
            }
            for (var j = 0; j < guidesToRemove.length; j++) guidesToRemove[j].remove();
        } catch (e) {
            /* 取得できないアイテムがあっても可能な範囲で続ける / Best effort */
        }
    }

    /**
     * アクティブなアートボードの外側にあるオブジェクトを削除する
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function deleteObjectsOutsideActiveArtboard(doc) {
        /* ロック・非表示・テンプレートのままでは選択できないため一時解除 / Relax locks and visibility first */
        var savedStructureStates = captureAndUnlockStructure(doc);

        removeGuidesOutsideArtboard(doc);

        /* アートボード上を選択→反転→削除 / Select on the artboard, invert, then delete */
        app.executeMenuCommand('selectallinartboard');
        app.executeMenuCommand('Inverse menu item');
        app.executeMenuCommand('clear');

        restoreStructure(savedStructureStates);
    }

    /**
     * アクティブなアートボードだけを残す
     * @param {Document} doc - 対象ドキュメント
     * @param {boolean} deleteOutsideObjects - 外側のオブジェクトも削除するか
     * @returns {void}
     */
    function keepActiveArtboardOnly(doc, deleteOutsideObjects) {
        if (doc.artboards.length > 1) {
            var activeArtboardIndex = doc.artboards.getActiveArtboardIndex();
            /* 末尾から削除してインデックスのずれを避ける / Remove from the end so the indices stay valid */
            for (var i = doc.artboards.length - 1; i >= 0; i--) {
                if (i !== activeArtboardIndex) doc.artboards.remove(i);
            }
        }

        if (deleteOutsideObjects) deleteObjectsOutsideActiveArtboard(doc);
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * インデントしたオプション用のグループを作る
     * @param {Panel} parentPanel - 追加先のパネル
     * @returns {Group} オプションを並べるグループ
     */
    function addOptionGroup(parentPanel) {
        var optionGroup = parentPanel.add('group');
        optionGroup.orientation = 'column';
        optionGroup.alignChildren = 'left';
        optionGroup.margins = OPTION_INDENT_MARGINS;
        optionGroup.spacing = PANEL_SPACING;
        return optionGroup;
    }

    /**
     * ボタン行（右寄せの キャンセル／OK）を作る
     * @param {Window} parentWindow - 追加先のダイアログ
     * @returns {{btnOK: Button, btnCancel: Button}} ボタン
     */
    function addButtonRow(parentWindow) {
        var btnRowGroup = parentWindow.add('group');
        btnRowGroup.orientation = 'row';
        btnRowGroup.alignChildren = ['left', 'center'];
        btnRowGroup.alignment = ['fill', 'center'];

        /* スペーサー（右側のボタンを押し出す）/ Spacer that pushes the buttons to the right */
        var spacer = btnRowGroup.add('group');
        spacer.alignment = ['fill', 'fill'];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add('group');
        btnRightGroup.alignment = ['right', 'center'];
        var btnCancel = btnRightGroup.add('button', undefined, getLabel('button.cancel'), { name: 'cancel' });
        var btnOK = btnRightGroup.add('button', undefined, getLabel('button.ok'), { name: 'ok' });
        btnCancel.preferredSize.width = BUTTON_WIDTH;
        btnOK.preferredSize.width = BUTTON_WIDTH;
        return { btnOK: btnOK, btnCancel: btnCancel };
    }

    /**
     * オプションダイアログを表示する
     * @param {object} artboardStats - { total: number, emptyIgnoreHidden: number, emptyAll: number }
     * @returns {object|null} 選択内容。キャンセル時は null
     */
    function showOptionsDialog(artboardStats) {
        var removeDialog = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
        removeDialog.orientation = 'column';
        removeDialog.alignChildren = 'fill';
        removeDialog.margins = DIALOG_MARGINS;
        removeDialog.spacing = DIALOG_SPACING;

        var modePanel = removeDialog.add('panel', undefined, getLabel('panel.mode'));
        modePanel.orientation = 'column';
        modePanel.alignChildren = 'left';
        modePanel.alignment = 'fill';
        modePanel.margins = PANEL_MARGINS;
        modePanel.spacing = PANEL_SPACING;

        /* ラジオは同じ親の中だけ排他になるため、2つともパネル直下に置く
           Radios are only exclusive within one parent, so both sit directly in the panel */
        var keepActiveRadio = modePanel.add('radiobutton', undefined, getLabel('radio.keepActive'));
        keepActiveRadio.helpTip = getLabel('tooltip.keepActive');
        keepActiveRadio.value = true;

        var keepActiveOptionGroup = addOptionGroup(modePanel);
        var deleteOutsideCheckbox = keepActiveOptionGroup.add('checkbox', undefined, getLabel('checkbox.deleteOutsideObjects'));
        deleteOutsideCheckbox.helpTip = getLabel('tooltip.deleteOutsideObjects');
        deleteOutsideCheckbox.value = DEFAULT_DELETE_OUTSIDE_OBJECTS;

        var emptyRadio = modePanel.add('radiobutton', undefined, getLabel('radio.empty'));
        emptyRadio.helpTip = getLabel('tooltip.empty');

        var emptyOptionGroup = addOptionGroup(modePanel);
        var emptyCountText = emptyOptionGroup.add('statictext', undefined, '');
        emptyCountText.preferredSize.width = COUNT_TEXT_WIDTH;
        var ignoreHiddenCheckbox = emptyOptionGroup.add('checkbox', undefined, getLabel('checkbox.ignoreHidden'));
        ignoreHiddenCheckbox.helpTip = getLabel('tooltip.ignoreHidden');
        ignoreHiddenCheckbox.value = DEFAULT_IGNORE_HIDDEN;

        /* アートボードが1枚のときは「空のアートボード」を選べない / Only one artboard leaves nothing to remove */
        var canRemoveEmpty = (artboardStats.total > 1);
        if (!canRemoveEmpty) {
            emptyRadio.enabled = false;
            emptyRadio.helpTip = getLabel('tooltip.singleArtboard');
        }

        var btnOK = addButtonRow(removeDialog).btnOK;

        /**
         * モードに応じてオプションの活性・件数表示・OKの活性を更新する
         * @returns {void}
         */
        function syncDialogState() {
            deleteOutsideCheckbox.enabled = keepActiveRadio.value;

            var emptyModeSelected = emptyRadio.value;
            emptyCountText.enabled = emptyModeSelected;
            ignoreHiddenCheckbox.enabled = emptyModeSelected;

            var emptyCount = ignoreHiddenCheckbox.value ? artboardStats.emptyIgnoreHidden : artboardStats.emptyAll;
            /* 英語はコロンのあとに空白を入れる / add a space after the colon in English */
            emptyCountText.text = labelText('fieldLabel.removalTarget') + (uiLang === 'ja' ? '' : ' ') + emptyCount;

            /* 空アートボードが0枚のときは実行できない / Nothing to do when no artboard is empty */
            btnOK.enabled = emptyModeSelected ? (emptyCount > 0) : true;
        }

        /* クリックだけでなくキーボード操作でも更新 / Refresh on click and on keyboard change */
        keepActiveRadio.onClick = keepActiveRadio.onChanging = syncDialogState;
        emptyRadio.onClick = emptyRadio.onChanging = syncDialogState;
        ignoreHiddenCheckbox.onClick = syncDialogState;
        syncDialogState();

        if (removeDialog.show() !== 1) return null;

        return {
            mode: keepActiveRadio.value ? 'keepActive' : 'empty',
            deleteOutsideObjects: deleteOutsideCheckbox.value,
            ignoreHidden: ignoreHiddenCheckbox.value
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * エントリポイント：ダイアログ→削除
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel('alert.noDocument'));
            return;
        }

        var doc = app.activeDocument;

        /* 非表示の扱いを切り替えても即座に件数を出せるよう、先に両方を数えておく
           Count both ways up front so the dialog can switch instantly */
        var emptyIndicesIgnoreHidden = findEmptyArtboardIndices(doc, true);
        var emptyIndicesAll = findEmptyArtboardIndices(doc, false);

        var removalOptions = showOptionsDialog({
            total: doc.artboards.length,
            emptyIgnoreHidden: emptyIndicesIgnoreHidden.length,
            emptyAll: emptyIndicesAll.length
        });
        if (!removalOptions) return;

        if (removalOptions.mode === 'keepActive') {
            keepActiveArtboardOnly(doc, removalOptions.deleteOutsideObjects);
        } else {
            removeArtboardsByIndices(doc, removalOptions.ignoreHidden ? emptyIndicesIgnoreHidden : emptyIndicesAll);
        }
    }

    main();

})();
