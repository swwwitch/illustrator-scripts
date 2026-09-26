#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);
#targetengine "DialogEngine"

/*

### 概要

作業アートボード、すべてのアートボード、または番号で指定したアートボードを、プレビューしながら指定の幅・高さに変更します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ResizeArtboardsAll.md

### Overview

Resizes the active artboard, every artboard, or the artboards you list by number to a given width and height, with a live preview.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ResizeArtboardsAll.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ResizeArtboardsAll";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-29";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ResizeArtboardsAll.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ResizeArtboardsAll.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================

    var WINDOW_MARGINS = 15;                 /* ウィンドウ外周の余白 / window margin */
    var PANEL_MARGINS  = [15, 20, 15, 10];   /* パネル余白 [左,上,右,下] / panel margins */
    var COLUMN_SPACING = 15;                 /* 2カラムの間隔 / gap between columns */
    var COLUMN_PANEL_SPACING = 10;           /* カラム内のパネル間隔 / gap between panels in a column */
    var SIZE_ROW_SPACING = 6;                /* 幅・高さの行間 / gap between the width and height rows */
    var ANCHOR_RADIO_SPACING = 12;           /* 基準点のラジオの間隔 / gap between the reference point radios */
    var BUTTON_ROW_TOP_MARGIN = 5;           /* ボタンエリアの上余白 / top margin of the button row */
    var SIZE_FIELD_CHARACTERS = 5;           /* 幅・高さ欄の桁数 / width & height field characters */
    var SPECIFY_FIELD_CHARACTERS = 12;       /* 番号指定欄の桁数 / artboard number field characters */
    var LABEL_WIDTH_PADDING = 6;             /* 項目名の幅に足す余白 / padding added to the measured label width */

    /* ダイアログの不透明度と、初回表示時の画面中央からの横オフセット / Dialog opacity and first-run offset from screen center */
    var DIALOG_OPACITY = 0.95;
    var DIALOG_FIRST_RUN_OFFSET_X = 300;

    /* プレビューの再描画の最短間隔（ms） / Minimum interval between preview redraws (ms) */
    var REDRAW_INTERVAL_MS = 40;

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

    /* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
       Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        /* 級（Q）と歯（H）は同じ長さだが、文字サイズは「Q」、距離は「H」と呼び分ける */
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    /**
     * pt・px のように整数で表示する単位か判定する
     * @param {{label: string}} unitInfo - 単位の情報
     * @returns {boolean} 整数で表示するなら true
     */
    function isIntegerUnit(unitInfo) {
        return unitInfo.label === "pt" || unitInfo.label === "px";
    }

    /**
     * pt の値を、単位に合わせた表示用の文字列にする（pt・px は整数、それ以外は小数2桁まで）
     * @param {number} valuePt - pt の値
     * @param {{label: string, pointsPerUnit: number}} unitInfo - 単位の情報
     * @returns {string} 表示用の文字列
     */
    function formatSizeValue(valuePt, unitInfo) {
        var unitValue = valuePt / unitInfo.pointsPerUnit;
        if (isIntegerUnit(unitInfo)) return String(Math.round(unitValue));
        return String(Math.round(unitValue * 100) / 100);
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    var uiLang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        dialog: {
            title: { ja: "アートボードのサイズを変更", en: "Resize Artboards" }
        },
        panel: {
            size: { ja: "サイズ", en: "Size" },
            anchor: { ja: "基準点", en: "Reference Point" },
            target: { ja: "対象のアートボード", en: "Target Artboards" }
        },
        fieldLabel: {
            width: { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" }
        },
        radio: {
            anchorTopLeft: { ja: "左上", en: "Top-Left" },
            anchorCenter: { ja: "中央", en: "Center" },
            activeArtboard: { ja: "作業アートボードのみ", en: "Active artboard only" },
            allArtboards: { ja: "すべてのアートボード", en: "All artboards" },
            specify: { ja: "指定", en: "Specify" }
        },
        tooltip: {
            sizeField: {
                ja: "↑↓で±1、Shift+↑↓で10の倍数にスナップします。",
                en: "Up/Down: ±1. Shift+Up/Down: snap to a multiple of 10."
            },
            anchorTopLeft: { ja: "左上の位置を保ったまま、右と下に伸び縮みさせます。", en: "Keep the top-left corner and resize to the right and down." },
            anchorCenter: { ja: "中心の位置を保ったまま伸び縮みさせます。", en: "Keep the center and resize around it." },
            specifyField: {
                ja: "アートボードの番号（1始まり）を範囲やカンマ区切りで指定します。\n例：1-3 / 1,3 / 2-4,7",
                en: "Artboard numbers (starting at 1), as ranges or comma-separated.\ne.g. 1-3 / 1,3 / 2-4,7"
            }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            errorOccurred: { ja: "エラーが発生しました：", en: "An error occurred: " }
        }
    };

    /**
     * ローカライズ文字列を取得する（キー漏れ時は英語へフォールバック）
     * @param {Object} labelSet - { ja, en } のラベル
     * @returns {string} 表示言語の文字列
     */
    function getLabel(labelSet) {
        if (!labelSet) return "";
        if (labelSet[uiLang] != null) return labelSet[uiLang];
        return (labelSet.en != null) ? labelSet.en : "";
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {Object} labelSet - ラベル
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 単位を括弧で添えたパネル名を返す（日本語は全角括弧、英語は半角）
     * @param {Object} labelSet - パネル名のラベル
     * @param {string} unitLabel - 単位の表示ラベル
     * @returns {string} 単位付きのパネル名
     */
    function panelTitleWithUnit(labelSet, unitLabel) {
        return (uiLang === "ja")
            ? getLabel(labelSet) + "（" + unitLabel + "）"
            : getLabel(labelSet) + " (" + unitLabel + ")";
    }

    // =========================================
    // エラー処理 / Error handling
    // =========================================

    /**
     * Error を行番号・ファイル名付きで読みやすく整形する
     * @param {Error} error - 例外
     * @returns {string} 整形した文字列
     */
    function formatError(error) {
        var messageText = (error && error.message) ? String(error.message) : String(error);
        var lineText = (error && error.line) ? (" line " + error.line) : "";
        var fileText = (error && error.fileName) ? (" (" + error.fileName + ")") : "";
        return messageText + lineText + fileText;
    }

    // =========================================
    // ダイアログ位置の記憶 / Dialog position persistence
    // =========================================
    // #targetengine の $.global に置くので、Illustrator を終了するまで位置が残る。
    // Kept in $.global of the named engine, so it lasts until Illustrator quits.

    var DIALOG_POSITION_KEY = "__ResizeArtboardsAll_Dialog";

    /**
     * 保存済みのダイアログ位置を取得する
     * @param {string} storageKey - $.global のキー
     * @returns {number[]|null} [x, y]。無ければ null
     */
    function getStoredLocation(storageKey) {
        return $.global[storageKey] && $.global[storageKey].length === 2 ? $.global[storageKey] : null;
    }

    /**
     * ダイアログ位置をセッションに保存する
     * @param {string} storageKey - $.global のキー
     * @param {number[]} location - [x, y]
     * @returns {void}
     */
    function storeLocation(storageKey, location) {
        $.global[storageKey] = [location[0], location[1]];
    }

    /**
     * 位置を画面内に収める
     * @param {number[]} location - [x, y]
     * @returns {number[]} 画面内に収めた [x, y]
     */
    function clampLocationToScreen(location) {
        /* 画面情報が取れない環境では元の位置のまま / keep the location when screen info is unavailable */
        try {
            var visibleBounds = ($.screens && $.screens.length) ? $.screens[0].visibleBounds : [0, 0, 1920, 1080];
            var clampedX = Math.max(visibleBounds[0] + 10, Math.min(location[0], visibleBounds[2] - 10));
            var clampedY = Math.max(visibleBounds[1] + 10, Math.min(location[1], visibleBounds[3] - 10));
            return [clampedX, clampedY];
        } catch (e) {
            return location;
        }
    }

    /**
     * ダイアログ位置の記憶を設定し、保存関数を返す
     * 保存位置があれば表示時に復元し、無ければ初回は画面中央から横にずらして表示する
     * @param {Window} dialogWindow - 対象のダイアログ
     * @param {string} positionKey - $.global のキー
     * @param {number} firstRunOffsetX - 初回表示時の中央からの横オフセット
     * @returns {function} 現在位置を保存する関数
     */
    function attachPositionPersistence(dialogWindow, positionKey, firstRunOffsetX) {
        var savedLocation = getStoredLocation(positionKey);

        var persist = function () {
            storeLocation(positionKey, [dialogWindow.location[0], dialogWindow.location[1]]);
        };

        if (savedLocation) {
            dialogWindow.onShow = function () {
                dialogWindow.location = clampLocationToScreen(savedLocation);
            };
        } else {
            dialogWindow.onShow = function () {
                dialogWindow.layout.layout(true);
                var screenWidth = $.screens[0].right - $.screens[0].left;
                var screenHeight = $.screens[0].bottom - $.screens[0].top;
                var centerX = screenWidth / 2 - dialogWindow.bounds.width / 2;
                var centerY = screenHeight / 2 - dialogWindow.bounds.height / 2;
                dialogWindow.location = [centerX + firstRunOffsetX, centerY];
            };
        }

        dialogWindow.onMove = persist;
        return persist;
    }

    // =========================================
    // 入力の解釈 / Input parsing
    // =========================================

    /**
     * 入力欄の文字列から数値を取り出す（数字・小数点・マイナス以外は無視）
     * @param {string} inputText - 入力欄の文字列
     * @returns {number} 数値（読み取れなければ NaN）
     */
    function parseNumberInput(inputText) {
        return parseFloat(String(inputText).replace(/[^0-9.\-]/g, ""));
    }

    /**
     * 「1-3,5,7-9」形式の番号指定を、0始まりのインデックス配列にする
     * 範囲外の番号は無視し、重複を除いて昇順に並べる
     * @param {string} specifyText - 入力された番号指定（1始まり）
     * @param {number} artboardCount - アートボードの数
     * @returns {number[]} 0始まりのインデックス
     */
    function parseArtboardNumbers(specifyText, artboardCount) {
        var compactText = String(specifyText || "").replace(/\s+/g, "");
        if (!compactText) return [];

        var isIndexListed = {};

        /**
         * 1始まりの番号を範囲内なら登録する
         * @param {number} artboardNumber - 1始まりの番号
         * @returns {void}
         */
        function addArtboardNumber(artboardNumber) {
            var artboardIndex = artboardNumber - 1;
            if (artboardIndex >= 0 && artboardIndex < artboardCount) isIndexListed[artboardIndex] = true;
        }

        var specifyParts = compactText.split(",");
        for (var i = 0; i < specifyParts.length; i++) {
            if (!specifyParts[i]) continue;
            var rangeMatch = specifyParts[i].match(/^(\d+)-(\d+)$/);
            if (rangeMatch) {
                var rangeStart = parseInt(rangeMatch[1], 10);
                var rangeEnd = parseInt(rangeMatch[2], 10);
                if (rangeStart > rangeEnd) {
                    var swapValue = rangeStart;
                    rangeStart = rangeEnd;
                    rangeEnd = swapValue;
                }
                for (var artboardNumber = rangeStart; artboardNumber <= rangeEnd; artboardNumber++) {
                    addArtboardNumber(artboardNumber);
                }
            } else {
                var singleNumber = parseInt(specifyParts[i], 10);
                if (!isNaN(singleNumber)) addArtboardNumber(singleNumber);
            }
        }

        var artboardIndexes = [];
        for (var indexKey in isIndexListed) {
            if (isIndexListed.hasOwnProperty(indexKey)) artboardIndexes.push(parseInt(indexKey, 10));
        }
        artboardIndexes.sort(function (a, b) { return a - b; });
        return artboardIndexes;
    }

    /**
     * 数値入力欄に ↑↓ キーでの増減を付ける
     * ↑↓で±1、Shift+↑↓で10の倍数にスナップ、Option(Alt)+↑↓で±0.1（最後に整数へ丸める）。0未満にはしない
     * @param {EditText} editText - 対象の入力欄
     * @param {function} onUpdate - 値を変えたあとに呼ぶ関数（新しい文字列を渡す）
     * @returns {void}
     */
    function changeValueByArrowKey(editText, onUpdate) {
        editText.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;

            var currentValue = parseNumberInput(editText.text);
            if (isNaN(currentValue)) currentValue = 0;

            var keyboardState = ScriptUI.environment.keyboardState;
            var isOptionPressed = (keyboardState.altKey || keyboardState.optionKey);
            var isShiftPressed = keyboardState.shiftKey && !isOptionPressed; // Option を優先 / Option wins
            var isUp = (event.keyName === "Up");

            var nextValue;
            if (isOptionPressed) {
                nextValue = currentValue + (isUp ? 0.1 : -0.1);
            } else if (isShiftPressed) {
                /* 10の倍数にスナップ（倍数上ならさらに±10） / Snap to the next multiple of 10 */
                var baseValue = Math.max(0, currentValue);
                var isMultipleOfTen = (baseValue % 10 === 0);
                if (isUp) {
                    nextValue = isMultipleOfTen ? (baseValue + 10) : (Math.ceil(baseValue / 10) * 10);
                } else {
                    nextValue = isMultipleOfTen ? Math.max(0, baseValue - 10) : (Math.floor(baseValue / 10) * 10);
                }
            } else {
                nextValue = currentValue + (isUp ? 1 : -1);
            }

            if (nextValue < 0) nextValue = 0;
            /* ↑↓操作では常に整数へ丸める / Arrow keys always land on an integer */
            nextValue = Math.round(nextValue);

            event.preventDefault();
            editText.text = String(nextValue);
            if (typeof onUpdate === "function") onUpdate(editText.text);
        });
    }

    // =========================================
    // アートボードの変更 / Artboard resizing
    // =========================================

    /**
     * すべてのアートボードの矩形を控える
     * @param {Document} targetDocument - 対象のドキュメント
     * @returns {Array<number[]>} [左, 上, 右, 下] の配列
     */
    function captureArtboardRects(targetDocument) {
        var artboardRects = [];
        for (var i = 0; i < targetDocument.artboards.length; i++) {
            artboardRects.push(targetDocument.artboards[i].artboardRect.slice());
        }
        return artboardRects;
    }

    /**
     * 控えた矩形に、すべてのアートボードを戻す
     * @param {Document} targetDocument - 対象のドキュメント
     * @param {Array<number[]>} artboardRects - 控えた [左, 上, 右, 下] の配列
     * @returns {void}
     */
    function restoreArtboardRects(targetDocument, artboardRects) {
        for (var i = 0; i < artboardRects.length; i++) {
            targetDocument.artboards[i].artboardRect = artboardRects[i].slice();
        }
    }

    /**
     * 元の矩形と基準点から、新しいサイズの矩形を求める
     * px 単位のときは、左上を整数座標にそろえる
     * @param {number[]} sourceRect - 元の [左, 上, 右, 下]
     * @param {number} widthPt - 新しい幅（pt）
     * @param {number} heightPt - 新しい高さ（pt）
     * @param {boolean} isTopLeftAnchor - 左上基準なら true、中央基準なら false
     * @param {boolean} snapToPixel - 左上を整数座標にそろえるなら true
     * @returns {number[]} 新しい [左, 上, 右, 下]
     */
    function computeResizedRect(sourceRect, widthPt, heightPt, isTopLeftAnchor, snapToPixel) {
        var leftEdge, topEdge;
        if (isTopLeftAnchor) {
            leftEdge = sourceRect[0];
            topEdge = sourceRect[1];
        } else {
            leftEdge = (sourceRect[0] + sourceRect[2]) / 2 - widthPt / 2;
            topEdge = (sourceRect[1] + sourceRect[3]) / 2 + heightPt / 2;
        }
        if (snapToPixel) {
            leftEdge = Math.round(leftEdge);
            topEdge = Math.round(topEdge);
        }
        return [leftEdge, topEdge, leftEdge + widthPt, topEdge - heightPt];
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 項目名と数値入力欄の行を追加する
     * @param {Group} parentGroup - 追加先のグループ
     * @param {Object} labelSet - 項目名のラベル
     * @param {string} initialText - 入力欄の初期値
     * @returns {{label: StaticText, input: EditText}} 追加した項目名と入力欄
     */
    function addSizeRow(parentGroup, labelSet, initialText) {
        var sizeRowGroup = parentGroup.add("group");
        sizeRowGroup.orientation = "row";
        var sizeLabel = sizeRowGroup.add("statictext", undefined, labelText(labelSet));
        sizeLabel.justify = "right";
        var sizeInput = sizeRowGroup.add("edittext", undefined, initialText);
        sizeInput.characters = SIZE_FIELD_CHARACTERS;
        sizeInput.helpTip = getLabel(LABELS.tooltip.sizeField);
        return { label: sizeLabel, input: sizeInput };
    }

    /**
     * 項目名の幅を、長いほうに合わせてそろえる（右揃えで入力欄の左端が縦にそろう）
     * @param {Window} dialogWindow - 文字幅を測るウィンドウ
     * @param {StaticText[]} labelControls - そろえる項目名
     * @returns {void}
     */
    function equalizeLabelWidths(dialogWindow, labelControls) {
        var maxWidth = 0;
        for (var i = 0; i < labelControls.length; i++) {
            var labelWidth;
            /* 表示前は measureString が使えない環境があるため、文字数から見積もる / estimate when measuring fails */
            try {
                labelWidth = Math.ceil(dialogWindow.graphics.measureString(labelControls[i].text)[0]);
            } catch (e) {
                labelWidth = labelControls[i].text.length * 7;
            }
            if (labelWidth > maxWidth) maxWidth = labelWidth;
        }
        for (var j = 0; j < labelControls.length; j++) {
            labelControls[j].preferredSize.width = maxWidth + LABEL_WIDTH_PADDING;
        }
    }

    /**
     * タイトル付きのパネルを追加する
     * @param {Group} parentGroup - 追加先のグループ
     * @param {string} panelTitle - パネル名
     * @param {string} childOrientation - 中のグループの並び（"row" / "column"）
     * @returns {Group} コントロールを入れるグループ
     */
    function addTitledPanel(parentGroup, panelTitle, childOrientation) {
        var titledPanel = parentGroup.add("panel", undefined, panelTitle);
        titledPanel.orientation = "row";
        titledPanel.alignChildren = ["left", "top"];
        titledPanel.margins = PANEL_MARGINS;
        var contentGroup = titledPanel.add("group");
        contentGroup.orientation = childOrientation;
        contentGroup.alignChildren = ["left", "center"];
        return contentGroup;
    }

    /**
     * 列のグループを追加する
     * @param {Group} parentGroup - 追加先のグループ
     * @returns {Group} 追加した列
     */
    function addColumn(parentGroup) {
        var columnGroup = parentGroup.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = "fill";
        columnGroup.spacing = COLUMN_PANEL_SPACING;
        return columnGroup;
    }

    /**
     * サイズ変更のダイアログを表示し、プレビューしながらアートボードを変更する
     * OK で変更を確定し、キャンセルで開いた時点のサイズに戻す
     * @param {Document} targetDocument - 対象のドキュメント
     * @param {{label: string, pointsPerUnit: number}} unitInfo - 定規の単位
     * @returns {boolean} OK で閉じたら true
     */
    function showResizeDialog(targetDocument, unitInfo) {
        var artboardCount = targetDocument.artboards.length;
        /* プレビューのたびに開いた時点の矩形へ戻してから適用する / Every preview starts from the original rects */
        var originalRects = captureArtboardRects(targetDocument);
        var activeRect = originalRects[targetDocument.artboards.getActiveArtboardIndex()];

        var resizeDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        resizeDialog.orientation = "column";
        resizeDialog.alignChildren = "fill";
        resizeDialog.margins = WINDOW_MARGINS;
        resizeDialog.opacity = DIALOG_OPACITY;
        var persistLocation = attachPositionPersistence(resizeDialog, DIALOG_POSITION_KEY, DIALOG_FIRST_RUN_OFFSET_X);

        var columnsGroup = resizeDialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.spacing = COLUMN_SPACING;
        var leftColumn = addColumn(columnsGroup);
        var rightColumn = addColumn(columnsGroup);

        /* サイズパネル：作業アートボードの今のサイズを初期値にする / Size panel, seeded with the active artboard */
        var sizeGroup = addTitledPanel(leftColumn, panelTitleWithUnit(LABELS.panel.size, unitInfo.label), "column");
        sizeGroup.spacing = SIZE_ROW_SPACING;
        var widthRow = addSizeRow(sizeGroup, LABELS.fieldLabel.width, formatSizeValue(Math.abs(activeRect[2] - activeRect[0]), unitInfo));
        var heightRow = addSizeRow(sizeGroup, LABELS.fieldLabel.height, formatSizeValue(Math.abs(activeRect[1] - activeRect[3]), unitInfo));
        var widthInput = widthRow.input;
        var heightInput = heightRow.input;
        equalizeLabelWidths(resizeDialog, [widthRow.label, heightRow.label]);

        /* 基準点パネル / Reference point panel */
        var anchorGroup = addTitledPanel(leftColumn, getLabel(LABELS.panel.anchor), "row");
        anchorGroup.spacing = ANCHOR_RADIO_SPACING;
        var anchorTopLeftRadio = anchorGroup.add("radiobutton", undefined, getLabel(LABELS.radio.anchorTopLeft));
        var anchorCenterRadio = anchorGroup.add("radiobutton", undefined, getLabel(LABELS.radio.anchorCenter));
        anchorTopLeftRadio.helpTip = getLabel(LABELS.tooltip.anchorTopLeft);
        anchorCenterRadio.helpTip = getLabel(LABELS.tooltip.anchorCenter);
        anchorTopLeftRadio.value = true;

        /* 対象パネル / Target panel */
        var targetGroup = addTitledPanel(rightColumn, getLabel(LABELS.panel.target), "column");
        targetGroup.alignChildren = ["left", "top"];
        var activeArtboardRadio = targetGroup.add("radiobutton", undefined, getLabel(LABELS.radio.activeArtboard));
        var allArtboardsRadio = targetGroup.add("radiobutton", undefined, getLabel(LABELS.radio.allArtboards));
        var specifyRadio = targetGroup.add("radiobutton", undefined, getLabel(LABELS.radio.specify));
        var specifyInput = targetGroup.add("edittext", undefined, "");
        specifyInput.characters = SPECIFY_FIELD_CHARACTERS;
        specifyInput.helpTip = getLabel(LABELS.tooltip.specifyField);
        activeArtboardRadio.value = true;

        /* アートボードが1つなら「すべて」「指定」は選べない / Only one artboard: nothing else to pick */
        if (artboardCount <= 1) {
            allArtboardsRadio.enabled = false;
            specifyRadio.enabled = false;
        }

        /* ボタンエリア / Button row */
        var btnRowGroup = resizeDialog.add("group");
        btnRowGroup.alignment = "center";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        var btnCancel = btnRowGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = btnRowGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        var lastRedrawTime = 0;

        /**
         * 再描画を間引く（連続入力で画面がもたつかないように）
         * @returns {void}
         */
        function throttledRedraw() {
            var currentTime = (new Date()).getTime();
            if (currentTime - lastRedrawTime >= REDRAW_INTERVAL_MS) {
                app.redraw();
                lastRedrawTime = currentTime;
            }
        }

        /**
         * 番号指定の入力欄を、［指定］を選んでいるときだけ使えるようにする
         * @returns {void}
         */
        function updateSpecifyEnabled() {
            specifyInput.enabled = (specifyRadio.value === true);
        }

        /**
         * 選択中の対象から、変更するアートボードのインデックスを決める
         * ［指定］で番号が1つも拾えないときは作業アートボードにする
         * @returns {number[]} 0始まりのインデックス
         */
        function resolveTargetIndexes() {
            if (specifyRadio.value) {
                var specifiedIndexes = parseArtboardNumbers(specifyInput.text, artboardCount);
                if (specifiedIndexes.length) return specifiedIndexes;
            } else if (allArtboardsRadio.value) {
                var allIndexes = [];
                for (var i = 0; i < artboardCount; i++) allIndexes.push(i);
                return allIndexes;
            }
            return [targetDocument.artboards.getActiveArtboardIndex()];
        }

        /**
         * 入力中の幅・高さで、対象のアートボードを変更する（プレビュー）
         * 幅・高さが数値として読めないあいだは何もしない
         * @returns {void}
         */
        function applyResizePreview() {
            var widthValue = parseNumberInput(widthInput.text);
            var heightValue = parseNumberInput(heightInput.text);
            if (isNaN(widthValue) || isNaN(heightValue) || widthValue <= 0 || heightValue <= 0) return;

            var widthPt = widthValue * unitInfo.pointsPerUnit;
            var heightPt = heightValue * unitInfo.pointsPerUnit;
            var isTopLeftAnchor = (anchorTopLeftRadio.value === true);
            var snapToPixel = (unitInfo.label === "px");

            restoreArtboardRects(targetDocument, originalRects);
            var targetIndexes = resolveTargetIndexes();
            for (var i = 0; i < targetIndexes.length; i++) {
                var artboardIndex = targetIndexes[i];
                targetDocument.artboards[artboardIndex].artboardRect =
                    computeResizedRect(originalRects[artboardIndex], widthPt, heightPt, isTopLeftAnchor, snapToPixel);
            }
            throttledRedraw();
        }

        /**
         * 対象の選び直しを反映する
         * @returns {void}
         */
        function onTargetChanged() {
            updateSpecifyEnabled();
            applyResizePreview();
        }

        widthInput.onChanging = applyResizePreview;
        heightInput.onChanging = applyResizePreview;
        changeValueByArrowKey(widthInput, applyResizePreview);
        changeValueByArrowKey(heightInput, applyResizePreview);
        anchorTopLeftRadio.onClick = applyResizePreview;
        anchorCenterRadio.onClick = applyResizePreview;
        activeArtboardRadio.onClick = onTargetChanged;
        allArtboardsRadio.onClick = onTargetChanged;
        specifyRadio.onClick = onTargetChanged;
        /* 番号は［指定］を選んでいるときだけ効く / The numbers only matter in Specify mode */
        specifyInput.onChanging = function () { if (specifyRadio.value) applyResizePreview(); };
        specifyInput.onChange = specifyInput.onChanging;

        var isConfirmed = false;
        btnOK.onClick = function () {
            persistLocation();
            applyResizePreview();
            isConfirmed = true;
            resizeDialog.close(1);
        };
        btnCancel.onClick = function () {
            persistLocation();
            restoreArtboardRects(targetDocument, originalRects);
            app.redraw();
            resizeDialog.close(0);
        };

        /* 表示時の位置合わせに続けて、幅の欄にフォーカスを置く / Focus the width field after positioning */
        var positionOnShow = resizeDialog.onShow;
        resizeDialog.onShow = function () {
            positionOnShow();
            widthInput.active = true;
        };

        updateSpecifyEnabled();
        applyResizePreview();
        resizeDialog.show();
        return isConfirmed;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ドキュメントを確かめてダイアログを開く（変更はダイアログ内のプレビューで確定する）
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }
        showResizeDialog(app.activeDocument, getUnitInfo());
    }

    try {
        main();
    } catch (e) {
        $.writeln("[" + SCRIPT_NAME + "] ERROR: " + formatError(e));
        alert(getLabel(LABELS.alert.errorOccurred) + formatError(e));
    }

    app.selectTool("Adobe Select Tool");

})();
