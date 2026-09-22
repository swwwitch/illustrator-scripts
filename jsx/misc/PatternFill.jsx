#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

2つのオブジェクトを選択し、大きい方を「容器」、小さい方を「タイル」として、容器のバウンディングボックス内にタイルを等間隔で敷き詰めます。
完全に内側に収まるタイルのみを残し、元の2オブジェクトはそのまま残します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PatternFill.md

### Overview

With two objects selected, treats the larger as the container and the smaller as the tile, then fills the container's bounding box with an even grid of tiles.
Only the tiles that fit entirely inside are kept, and both originals are left in place.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PatternFill.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PatternFill";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.4.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-10-26";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PatternFill.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PatternFill.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 間隔の初期値を決める比率（タイルの幅に掛ける）/ Initial spacing as a ratio of the tile width */
    var DEFAULT_GAP_RATIO = 0.2;

    /* プレビューの不透明度（%）/ Opacity of the preview (%) */
    var PREVIEW_OPACITY = 60;

    /* 敷き詰めたタイルを入れるグループ名（プレビューは名前で探して消す）/ Group names; previews are found and removed by name */
    var GRID_GROUP_NAME = 'TiledGrid';
    var PREVIEW_GROUP_NAME = 'TiledGrid_preview';

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_OFFSET_X = 300;    /* ダイアログを右へずらす量（px）/ Horizontal shift of the dialog */
    var DIALOG_OPACITY = 0.98;    /* ダイアログの不透明度 / Dialog opacity */
    var LABEL_WIDTH = 60;         /* 項目名の幅 / Field label width */
    var FIELD_CHARACTERS = 4;     /* 数値入力欄の桁数 / Width of the number fields in characters */

    /**
     * ダイアログを表示するときに、既定の位置からずらす
     * @param {Window} targetDialog - 対象ダイアログ
     * @param {number} offsetX - 横のずらし量
     * @param {number} offsetY - 縦のずらし量
     * @returns {void}
     */
    function shiftDialogPosition(targetDialog, offsetX, offsetY) {
        targetDialog.onShow = function() {
            var currentX = targetDialog.location[0];
            var currentY = targetDialog.location[1];
            targetDialog.location = [currentX + offsetX, currentY + offsetY];
        };
    }

    /**
     * 右揃えの項目名と入力欄を並べる行を作る
     * @param {Window} parentDialog - 親ダイアログ
     * @param {string} labelPath - 項目名のラベルのパス
     * @returns {Group} 項目名を追加済みの行
     */
    function addFieldRow(parentDialog, labelPath) {
        var fieldRow = parentDialog.add('group');
        fieldRow.alignment = ['fill', 'top'];
        fieldRow.alignChildren = ['left', 'center'];
        var fieldLabel = fieldRow.add('statictext', undefined, labelText(labelPath));
        fieldLabel.justify = 'right';
        fieldLabel.preferredSize.width = LABEL_WIDTH;
        return fieldRow;
    }

    /**
     * 行に数値入力欄を追加する
     * @param {Group} fieldRow - 追加先の行
     * @param {string} initialText - 初期値
     * @param {string} tooltipPath - helpTip のラベルのパス
     * @returns {EditText} 追加した入力欄
     */
    function addNumberField(fieldRow, initialText, tooltipPath) {
        var numberField = fieldRow.add('edittext', undefined, initialText);
        numberField.characters = FIELD_CHARACTERS;
        numberField.helpTip = getLabel(tooltipPath);
        return numberField;
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位テーブル（配列の添字が rulerType コードと一致：0=in, 1=mm, 2=pt …）/ Unit table; the array index equals the rulerType code */
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

    /* Q ではなく H と表示する設定キー / Preference keys that display H instead of Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 設定キーごとの単位情報を取得する
     * @param {string} prefKey - 環境設定キー（省略時は "rulerType"）
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        var unit = UNITS[unitCode] || UNITS[2];
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * Illustrator の UI 言語から表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = detectUILanguage();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "敷き詰め設定", en: "Tile Fill Settings" }
        },
        fieldLabel: {
            gridCount: { ja: "グリッド数", en: "Grid count" },
            spacing: { ja: "間隔", en: "Spacing" },
            margin: { ja: "マージン", en: "Margin" }
        },
        unit: {
            columns: { ja: "列", en: "Columns" },
            rows: { ja: "行", en: "Rows" }
        },
        checkbox: {
            brick: { ja: "レンガ状", en: "Brick pattern" },
            symbolize: { ja: "シンボル化", en: "Symbolize" }
        },
        tooltip: {
            columns: {
                ja: "横に並べる数です。0 のときは容器に収まるだけ自動で並べます。",
                en: "How many tiles to place horizontally. 0 fills the container automatically."
            },
            rows: {
                ja: "縦に並べる数です。0 のときは容器に収まるだけ自動で並べます。",
                en: "How many tiles to place vertically. 0 fills the container automatically."
            },
            spacing: {
                ja: "隣り合うタイルのアキです。縦横とも同じ値になります。",
                en: "Space between neighbouring tiles, applied both horizontally and vertically."
            },
            margin: {
                ja: "容器の内側に空ける余白です。負の値も入力できます。",
                en: "Inset kept inside the container. Negative values are allowed."
            },
            brick: {
                ja: "1行おきに半個分ずらして、レンガのように並べます。",
                en: "Offsets every other row by half a tile, like brickwork."
            },
            symbolize: {
                ja: "タイルをシンボルとして複製します。あとからまとめて差し替えられます。",
                en: "Duplicates the tile as a symbol, so every copy can be swapped later at once."
            }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "fieldLabel.spacing" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return labelNode[uiLang] || labelNode.en || labelPath;
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 単位を括弧でくくった表記を返す（日本語は全角括弧）
     * @param {string} unitLabel - 単位の表示名
     * @returns {string} 括弧付きの単位
     */
    function unitSuffix(unitLabel) {
        return (uiLang === "ja") ? ("（" + unitLabel + "）") : (" (" + unitLabel + ")");
    }

    // =========================================
    // 形状と配置の計算 / Geometry
    // =========================================

    /**
     * オブジェクトの表示上の境界と寸法を取得する
     * @param {PageItem} pageItem - 対象オブジェクト
     * @returns {{left: number, top: number, right: number, bottom: number, width: number, height: number, area: number}} 境界情報
     */
    function getBoundsInfo(pageItem) {
        var visibleBounds = pageItem.visibleBounds; /* [left, top, right, bottom] */
        var itemWidth = visibleBounds[2] - visibleBounds[0];
        var itemHeight = visibleBounds[1] - visibleBounds[3];
        return {
            left: visibleBounds[0],
            top: visibleBounds[1],
            right: visibleBounds[2],
            bottom: visibleBounds[3],
            width: itemWidth,
            height: itemHeight,
            area: Math.abs(itemWidth * itemHeight)
        };
    }

    /**
     * 4点の閉じたパスで、すべてのアンカーが指定の種類か判定する
     * @param {PageItem} pageItem - 対象オブジェクト
     * @param {PointType} pointType - アンカーの種類
     * @returns {boolean} 当てはまれば true
     */
    function isFourPointPathOfType(pageItem, pointType) {
        if (pageItem.typename !== 'PathItem' || !pageItem.closed || pageItem.pathPoints.length !== 4) return false;
        for (var i = 0; i < 4; i++) {
            if (pageItem.pathPoints[i].pointType !== pointType) return false;
        }
        return true;
    }

    /**
     * 長方形のパス（4点すべてがコーナー）か判定する
     * @param {PageItem} pageItem - 対象オブジェクト
     * @returns {boolean} 長方形なら true
     */
    function isRectanglePath(pageItem) {
        return isFourPointPathOfType(pageItem, PointType.CORNER);
    }

    /**
     * 楕円らしいパス（4点すべてがスムーズ）か判定する
     * @param {PageItem} pageItem - 対象オブジェクト
     * @returns {boolean} 楕円らしければ true
     */
    function isEllipseLikePath(pageItem) {
        return isFourPointPathOfType(pageItem, PointType.SMOOTH);
    }

    /**
     * 並べる列数と行数を決める。指定が 0 のときは容器に収まるだけ並べる
     * @param {Object} containerInfo - 容器の境界情報
     * @param {number} stepX - 横の送り（pt）
     * @param {number} stepY - 縦の送り（pt）
     * @param {boolean} isBrick - レンガ状に並べるか
     * @param {number} fixedColumns - 指定の列数（0 で自動）
     * @param {number} fixedRows - 指定の行数（0 で自動）
     * @returns {{columns: number, rows: number}} 列数と行数
     */
    function computeGridSize(containerInfo, stepX, stepY, isBrick, fixedColumns, fixedRows) {
        return {
            columns: (fixedColumns > 0)
                ? Math.round(fixedColumns)
                : Math.max(1, Math.ceil((containerInfo.width + (isBrick ? stepX * 0.5 : 0)) / stepX)),
            rows: (fixedRows > 0)
                ? Math.round(fixedRows)
                : Math.max(1, Math.ceil(containerInfo.height / stepY))
        };
    }

    /**
     * グリッドの各マスの左上座標を、行ごとに左から順に渡す（レンガ状なら奇数行を半個ずらす）
     * @param {Object} containerInfo - 容器の境界情報
     * @param {number} stepX - 横の送り（pt）
     * @param {number} stepY - 縦の送り（pt）
     * @param {{columns: number, rows: number}} gridSize - 列数と行数
     * @param {boolean} isBrick - レンガ状に並べるか
     * @param {Function} visitCell - (xLeft, yTop) を受け取る関数
     * @returns {void}
     */
    function forEachGridCell(containerInfo, stepX, stepY, gridSize, isBrick, visitCell) {
        for (var r = 0; r < gridSize.rows; r++) {
            var yTop = containerInfo.top - r * stepY;
            for (var c = 0; c < gridSize.columns; c++) {
                var xLeft = containerInfo.left + c * stepX + ((isBrick && (r % 2 === 1)) ? stepX * 0.5 : 0);
                visitCell(xLeft, yTop);
            }
        }
    }

    /**
     * タイルを残すかどうかの判定に使う、容器の形とマージンを差し引いた範囲をまとめる
     * @param {PageItem} container - 容器
     * @param {Object} containerInfo - 容器の境界情報
     * @param {number} marginPt - マージン（pt）
     * @returns {Object} 容器の形（isRect / isEllipse）と判定に使う範囲
     */
    function buildContainerShape(container, containerInfo, marginPt) {
        var isPath = (container.typename === 'PathItem');
        return {
            isRect: isPath && isRectanglePath(container),
            isEllipse: isPath && isEllipseLikePath(container),
            left: containerInfo.left + marginPt,
            right: containerInfo.right - marginPt,
            top: containerInfo.top - marginPt,
            bottom: containerInfo.bottom + marginPt,
            centerX: (containerInfo.left + containerInfo.right) / 2.0,
            centerY: (containerInfo.top + containerInfo.bottom) / 2.0,
            radiusX: Math.max(0, Math.abs(containerInfo.right - containerInfo.left) / 2.0 - marginPt),
            radiusY: Math.max(0, Math.abs(containerInfo.top - containerInfo.bottom) / 2.0 - marginPt)
        };
    }

    /**
     * タイルが容器の内側に収まっているか判定する。
     * 長方形は外接矩形が丸ごと内側、楕円は4隅が楕円の内側、それ以外は中心が内側なら残す
     * @param {Object} tileBounds - タイルの境界情報
     * @param {Object} containerShape - buildContainerShape() の結果
     * @returns {boolean} 残すなら true
     */
    function isTileInside(tileBounds, containerShape) {
        if (containerShape.isRect) {
            return (tileBounds.left >= containerShape.left && tileBounds.right <= containerShape.right &&
                tileBounds.top <= containerShape.top && tileBounds.bottom >= containerShape.bottom);
        }
        if (containerShape.isEllipse) {
            if (!(containerShape.radiusX > 0 && containerShape.radiusY > 0)) return false;
            var corners = [
                [tileBounds.left, tileBounds.top],
                [tileBounds.right, tileBounds.top],
                [tileBounds.left, tileBounds.bottom],
                [tileBounds.right, tileBounds.bottom]
            ];
            for (var k = 0; k < 4; k++) {
                var dx = (corners[k][0] - containerShape.centerX) / containerShape.radiusX;
                var dy = (corners[k][1] - containerShape.centerY) / containerShape.radiusY;
                if (!((dx * dx + dy * dy) <= 1.000001)) return false;
            }
            return true;
        }
        var tileCenterX = (tileBounds.left + tileBounds.right) / 2.0;
        var tileCenterY = (tileBounds.top + tileBounds.bottom) / 2.0;
        return (tileCenterX >= containerShape.left && tileCenterX <= containerShape.right &&
            tileCenterY <= containerShape.top && tileCenterY >= containerShape.bottom);
    }

    /**
     * グループ内のタイルのうち、容器に収まらないものを削除する（マスクは使わない）
     * @param {GroupItem} tileGroup - タイルを入れたグループ
     * @param {Object} containerShape - buildContainerShape() の結果
     * @returns {void}
     */
    function trimTilesOutside(tileGroup, containerShape) {
        for (var i = tileGroup.pageItems.length - 1; i >= 0; i--) {
            var gridItem = tileGroup.pageItems[i];
            if (!isTileInside(getBoundsInfo(gridItem), containerShape)) gridItem.remove();
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * テキストフィールドに↑↓キーでの値の増減を組み込む
     * ↑↓で±1、shift併用で±10（10の倍数にスナップ）、option併用で±0.1
     * @param {EditText} editText - 対象のテキストフィールド
     * @param {boolean} allowNegative - 負の値を許すか
     * @param {Function} onValueChanged - キーを押したあとに呼ぶ関数
     * @returns {void}
     */
    function changeValueByArrowKey(editText, allowNegative, onValueChanged) {
        allowNegative = !!allowNegative;
        editText.addEventListener("keydown", function(event) {
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
                    if (!allowNegative && value < 0) value = 0;
                    event.preventDefault();
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                if (event.keyName == "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= delta;
                    event.preventDefault();
                }
            } else {
                delta = 1;
                if (event.keyName == "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= delta;
                    if (!allowNegative && value < 0) value = 0;
                    event.preventDefault();
                }
            }

            if (keyboard.altKey) {
                value = Math.round(value * 10) / 10;
            } else {
                value = Math.round(value);
            }

            editText.text = value;
            onValueChanged();
        });
    }

    /**
     * 敷き詰め設定のダイアログを表示する
     * @param {number} tileWidthPt - タイルの幅（pt）。間隔の初期値に使う
     * @param {Function} onPreview - 入力が変わるたびに設定（pt 換算済み）を受け取る関数
     * @returns {{gap: number, margin: number, isBrick: boolean, symbolize: boolean, columns: number, rows: number}|null} 設定（pt）。キャンセル時は null
     */
    function showFillDialog(tileWidthPt, onPreview) {
        /* ダイアログのタイトルはラベル＋バージョン / Dialog title = label + version */
        var fillDialog = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);

        /* ウィンドウ見た目 / Window appearance */
        fillDialog.opacity = DIALOG_OPACITY;
        shiftDialogPosition(fillDialog, DIALOG_OFFSET_X, 0);
        fillDialog.alignChildren = 'fill';

        var rulerUnit = getUnitInfo("rulerType");
        var unitLabel = rulerUnit.label;

        /* グリッド数（0 で容器に合わせて自動）/ Grid count (0 fits the container automatically) */
        var gridCountRow = addFieldRow(fillDialog, 'fieldLabel.gridCount');
        var columnsField = addNumberField(gridCountRow, '0', 'tooltip.columns');
        gridCountRow.add('statictext', undefined, getLabel('unit.columns'));
        var rowsField = addNumberField(gridCountRow, '0', 'tooltip.rows');
        gridCountRow.add('statictext', undefined, getLabel('unit.rows'));

        /* 間隔 / Spacing */
        var spacingRow = addFieldRow(fillDialog, 'fieldLabel.spacing');
        var defaultGapValue = Math.round(tileWidthPt / rulerUnit.pointsPerUnit * DEFAULT_GAP_RATIO);
        var gapField = addNumberField(spacingRow, String(defaultGapValue), 'tooltip.spacing');
        spacingRow.add('statictext', undefined, unitSuffix(unitLabel));
        gapField.active = true;
        changeValueByArrowKey(gapField, false, updatePreviewFromFields);

        /* マージン / Margin */
        var marginRow = addFieldRow(fillDialog, 'fieldLabel.margin');
        var marginField = addNumberField(marginRow, '0', 'tooltip.margin');
        marginRow.add('statictext', undefined, unitSuffix(unitLabel));
        changeValueByArrowKey(marginField, true, updatePreviewFromFields); /* マージンは負OK / margin can be negative */

        /* 入力欄の値を pt に換算して読む（チェックボックスは作成前なら false）
           Read the fields converted to points (checkboxes read false before they exist) */
        function readFillSettings() {
            var pointsPerUnit = getUnitInfo("rulerType").pointsPerUnit;
            var gapValue = Math.max(0, parseFloat(gapField.text) || 0);
            var marginValue = parseFloat(marginField.text);
            if (isNaN(marginValue)) marginValue = 0;
            return {
                gap: gapValue * pointsPerUnit,
                margin: marginValue * pointsPerUnit,
                isBrick: !!(brickCheckbox && brickCheckbox.value),
                symbolize: !!(symbolizeCheckbox && symbolizeCheckbox.value),
                columns: Math.max(0, parseInt(columnsField.text, 10) || 0),
                rows: Math.max(0, parseInt(rowsField.text, 10) || 0)
            };
        }

        /* プレビュー更新 / Update preview */
        function updatePreviewFromFields() {
            onPreview(readFillSettings());
        }
        gapField.onChanging = updatePreviewFromFields;
        columnsField.onChanging = updatePreviewFromFields;
        rowsField.onChanging = updatePreviewFromFields;
        changeValueByArrowKey(columnsField, false, updatePreviewFromFields);
        changeValueByArrowKey(rowsField, false, updatePreviewFromFields);
        updatePreviewFromFields();

        /* レンガ状・シンボル化は中央にまとめる / Brick pattern and Symbolize sit centered */
        var optionRow = fillDialog.add('group');
        optionRow.alignment = ['fill', 'top'];
        optionRow.alignChildren = ['center', 'center'];

        /* レンガ状 / Brick pattern */
        var brickGroup = optionRow.add('group');
        brickGroup.alignChildren = ['left', 'center'];
        var brickCheckbox = brickGroup.add('checkbox', undefined, getLabel('checkbox.brick'));
        brickCheckbox.helpTip = getLabel('tooltip.brick');
        brickCheckbox.value = false;
        brickCheckbox.onClick = updatePreviewFromFields;

        /* シンボル化して複製 / Duplicate as Symbol */
        var symbolizeGroup = optionRow.add('group');
        symbolizeGroup.alignChildren = ['left', 'center'];
        var symbolizeCheckbox = symbolizeGroup.add('checkbox', undefined, getLabel('checkbox.symbolize'));
        symbolizeCheckbox.helpTip = getLabel('tooltip.symbolize');
        symbolizeCheckbox.value = false;

        /* ボタン / Buttons */
        var btnRowGroup = fillDialog.add('group');
        btnRowGroup.alignment = 'right';
        btnRowGroup.add('button', undefined, getLabel('button.cancel'), { name: 'cancel' });
        btnRowGroup.add('button', undefined, getLabel('button.ok'), { name: 'ok' });

        if (fillDialog.show() !== 1) return null;

        /* OKで確定値をptに変換 / Convert to pt on OK */
        return readFillSettings();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択した2つのオブジェクトを容器とタイルに振り分け、ダイアログで設定して敷き詰める
     * @returns {void}
     */
    function main() {
        var doc = app.documents.length ? app.activeDocument : null;
        if (!doc) {
            alert('ドキュメントが開いていません / No active document.');
            return;
        }

        if (!doc.selection || doc.selection.length !== 2) {
            alert('ちょうど2つを選択してください。/ Please select exactly two objects.');
            return;
        }

        /* 容器とタイルを面積で判定 / Detect container and tile by area */
        var firstItem = doc.selection[0];
        var secondItem = doc.selection[1];
        var firstInfo = getBoundsInfo(firstItem);
        var secondInfo = getBoundsInfo(secondItem);
        var isFirstContainer = (firstInfo.area >= secondInfo.area);
        var fillTarget = {
            container: isFirstContainer ? firstItem : secondItem,
            tile: isFirstContainer ? secondItem : firstItem,
            containerInfo: isFirstContainer ? firstInfo : secondInfo,
            tileInfo: isFirstContainer ? secondInfo : firstInfo,
            previewGroup: null
        };

        if (fillTarget.tileInfo.width <= 0 || fillTarget.tileInfo.height <= 0) {
            alert('タイルのサイズが不正です / Invalid tile size.');
            return;
        }

        var fillSettings = showFillDialog(fillTarget.tileInfo.width, function(previewSettings) {
            renderPreview(fillTarget, previewSettings);
        });
        clearPreview(fillTarget);
        if (!fillSettings) {
            app.userInteractionLevel = UserInteractionLevel.DISPLAYALERTS;
            return;
        }

        placeTileGrid(doc, fillTarget, fillSettings);
    }

    /**
     * プレビューを削除する（参照で消し、さらに名前で残りを掃除する）
     * @param {Object} fillTarget - 容器・タイル・プレビューの参照
     * @returns {void}
     */
    function clearPreview(fillTarget) {
        try {
            if (fillTarget.previewGroup && fillTarget.previewGroup.isValid) {
                fillTarget.previewGroup.remove();
            }
        } catch (e1) {}
        fillTarget.previewGroup = null;

        /* 名前で余分なプレビューを掃除 / Sweep stray previews */
        try {
            var containerLayer = fillTarget.container.layer;
            for (var i = containerLayer.groupItems.length - 1; i >= 0; i--) {
                var groupItem = containerLayer.groupItems[i];
                if (groupItem.name.indexOf(PREVIEW_GROUP_NAME) === 0) {
                    try {
                        groupItem.remove();
                    } catch (e2) {}
                }
            }
        } catch (e3) {}
    }

    /**
     * プレビューを描画する（半透明のグループにタイルを複製し、はみ出すものを削除）
     * @param {Object} fillTarget - 容器・タイル・プレビューの参照
     * @param {Object} previewSettings - showFillDialog() と同じ形の設定（pt）
     * @returns {void}
     */
    function renderPreview(fillTarget, previewSettings) {
        var marginPt = previewSettings.margin || 0;
        var isBrick = !!previewSettings.isBrick;
        clearPreview(fillTarget);

        var stepX = fillTarget.tileInfo.width + previewSettings.gap;
        var stepY = fillTarget.tileInfo.height + previewSettings.gap;
        if (stepX <= 0 || stepY <= 0) return;
        var gridSize = computeGridSize(fillTarget.containerInfo, stepX, stepY, isBrick, previewSettings.columns, previewSettings.rows);

        var previewGroup = fillTarget.container.layer.groupItems.add();
        previewGroup.name = PREVIEW_GROUP_NAME;
        fillTarget.previewGroup = previewGroup;

        var containerShape = buildContainerShape(fillTarget.container, fillTarget.containerInfo, marginPt);

        forEachGridCell(fillTarget.containerInfo, stepX, stepY, gridSize, isBrick, function(xLeft, yTop) {
            var tileCopy = fillTarget.tile.duplicate(previewGroup, ElementPlacement.PLACEATBEGINNING);
            tileCopy.position = [xLeft, yTop];
        });

        /* プレビューもトリム / Trim preview too */
        trimTilesOutside(previewGroup, containerShape);

        try {
            previewGroup.opacity = PREVIEW_OPACITY;
        } catch (e) {}
        app.redraw();
    }

    /**
     * タイルをシンボルとして登録する
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} tile - タイル
     * @returns {Symbol|null} 登録したシンボル。失敗時は null（通常の複製に切り替える）
     */
    function createTileSymbol(doc, tile) {
        try {
            var symbolSource = tile.duplicate();
            var symbolDefinition = doc.symbols.add(symbolSource);
            try { symbolSource.remove(); } catch (eSource) {}
            return symbolDefinition;
        } catch (eSymbol) {
            return null; /* 失敗時は通常複製 / fallback */
        }
    }

    /**
     * 確定した設定でタイルを敷き詰め、容器からはみ出すものを削除する
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} fillTarget - 容器・タイルの参照
     * @param {Object} fillSettings - showFillDialog() の結果（pt）
     * @returns {void}
     */
    function placeTileGrid(doc, fillTarget, fillSettings) {
        var tile = fillTarget.tile;
        var tileInfo = fillTarget.tileInfo;
        var containerInfo = fillTarget.containerInfo;

        /* 必要ならシンボル作成 / Symbolize if needed */
        var symbolDefinition = fillSettings.symbolize ? createTileSymbol(doc, tile) : null;

        /* 配置数を計算 / Compute counts */
        var stepX = tileInfo.width + fillSettings.gap;
        var stepY = tileInfo.height + fillSettings.gap;
        var gridSize = computeGridSize(containerInfo, stepX, stepY, fillSettings.isBrick, fillSettings.columns, fillSettings.rows);

        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;
        app.redraw();

        /* 複製グループを作成 / Create group for final duplicates */
        var gridGroup = fillTarget.container.layer.groupItems.add();
        gridGroup.name = GRID_GROUP_NAME;

        /* 敷き詰め実行 / Duplicate and place in grid */
        forEachGridCell(containerInfo, stepX, stepY, gridSize, fillSettings.isBrick, function(xLeft, yTop) {
            /* 同じ場所に元タイルがある場合はスキップ / Skip if same as original tile */
            if (Math.abs(xLeft - tileInfo.left) < 0.01 && Math.abs(yTop - tileInfo.top) < 0.01) return;
            if (fillSettings.symbolize && symbolDefinition) {
                var symbolItem = doc.symbolItems.add(symbolDefinition);
                symbolItem.move(gridGroup, ElementPlacement.PLACEATBEGINNING);
                symbolItem.position = [xLeft, yTop];
            } else {
                var tileCopy = tile.duplicate(gridGroup, ElementPlacement.PLACEATBEGINNING);
                tileCopy.position = [xLeft, yTop];
            }
        });

        /* マスクなしトリム / Trim without mask */
        trimTilesOutside(gridGroup, buildContainerShape(fillTarget.container, containerInfo, fillSettings.margin));

        app.userInteractionLevel = UserInteractionLevel.DISPLAYALERTS;
        app.redraw();
    }

    /**
     * メイン処理を実行し、例外はアラートで知らせる
     * @returns {void}
     */
    function runMainWithErrorAlert() {
        try {
            main();
        } catch (e) {
            alert('[TileSmallIntoLarge] Error:\n' + e);
        }
    }

    /* メインを1アクションで実行（ScriptLanguage / UndoModes がある環境のみ）/ Run main in single undo where the enums exist */
    try {
        if (typeof app.doScript === 'function' && typeof ScriptLanguage !== 'undefined' && typeof UndoModes !== 'undefined') {
            app.doScript(runMainWithErrorAlert, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, 'TileSmallIntoLarge');
        } else {
            runMainWithErrorAlert();
        }
    } catch (e) {
        runMainWithErrorAlert();
    }

})();
