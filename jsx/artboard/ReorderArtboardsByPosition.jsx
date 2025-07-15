#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

$.localize = true;

/*
### スクリプト名：

ReorderArtboardsByPosition.jsx

### 概要

- アートボードを名前順や位置順（左上、右上、上左、上右）で並べ替えるスクリプト
- カンバス（ドキュメント）上の見た目の並びを基に、［アートボード］パネルの順序を変更

### 主な機能

- アートボードを名前順または位置順にソート
- 許容差（Tolerance）の設定により微妙なズレを許容
- UI でソート方法と許容差を選択可能

### 処理の流れ

1. ダイアログでソート方法と許容差を選択
2. OK ボタンで並べ替えを実行
3. 結果を即時反映

### オリジナル、謝辞

- m1b 氏: https://community.adobe.com/t5/illustrator-discussions/randomly-order-artboards/m-p/12692397
- https://community.adobe.com/t5/illustrator-discussions/illustrator-script-to-renumber-reorder-the-artboards-with-there-position/m-p/12752568

### 更新履歴

- v1.0 (20231115) : 初期バージョン（Andrew_BJ による UI 改良と上限拡張）
- v1.0 (20231116) : 許容差の自動計算機能とスライダーを追加、ロジック整理

---

### Script Name:

ReorderArtboardsByPosition.jsx

### Overview

- Script to reorder artboards by name or position (Top Left, Top Right, Left Top, Right Top)
- Reorders the Artboards panel order based on the visual (canvas) arrangement on the document

### Main Features

- Sort artboards by name or various position orders
- Adjust tolerance to allow slight misalignments
- Choose sorting method and tolerance via UI

### Process Flow

1. Select sort method and tolerance in dialog
2. Execute sorting with OK button
3. Changes are applied immediately

### Original / Credit

- m1b: https://community.adobe.com/t5/illustrator-discussions/randomly-order-artboards/m-p/12692397
- https://community.adobe.com/t5/illustrator-discussions/illustrator-script-to-renumber-reorder-the-artboards-with-there-position/m-p/12752568

### Changelog

- v1.0 (20231115): Initial version (UI improvements and limit extension by Andrew_BJ)
- v1.0 (20231116): Added tolerance auto-calculation feature, slider support, and logic cleanup
*/

/*
ラベル定義 (日本語 / 英語)
UI構築、ソートメソッド定義、メイン処理、ソート関数群を含む
*/


var LABELS = {
    dialogTitle: {
        ja: "アートボードを並べ替え v1.1",
        en: "Sort Artboards"
    },
    sortMethod: {
        ja: "ソート方法",
        en: "Sort Method"
    },
    sort: {
        ja: "ソート",
        en: "Sort"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    tolerance: {
        ja: "許容差",
        en: "Tolerance"
    },
    noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
    noArtboards: { ja: "アートボードが存在しません。", en: "No artboards found." }
};

var SortMethod = {
    NAME: {
        sorter: sortByName,
        displayName: {
            ja: '名前順',
            en: 'By Name'
        }
    },
    POSITION_TOP_LEFT: {
        sorter: sortByPositionTopLeft,
        displayName: {
            ja: '位置順：左上',
            en: 'By Position Top Left'
        }
    },
    POSITION_TOP_RIGHT: {
        sorter: sortByPositionTopRight,
        displayName: {
            ja: '位置順：右上',
            en: 'By Position Top Right'
        }
    },
    POSITION_LEFT_TOP: {
        sorter: sortByPositionLeftTop,
        displayName: {
            ja: '位置順：上左',
            en: 'By Position Left Top'
        }
    },
    POSITION_RIGHT_TOP: {
        sorter: sortByPositionRightTop,
        displayName: {
            ja: '位置順：上右',
            en: 'By Position Right Top'
        }
    }
};

var settings = {
    action: sortArtboards,
    sortMethod: SortMethod.POSITION_TOP_LEFT,
    decimalPlaces: 3
};

main();

function main() {
    if (app.documents.length === 0) {
        alert(LABELS.noDocument.ja);
        return;
    }
    var doc = app.activeDocument;
    if (doc.artboards.length === 0) {
        alert(LABELS.noArtboards.ja);
        return;
    }

    var sortMethods = [];
    for (var key in SortMethod) {
        if (SortMethod.hasOwnProperty(key)) {
            sortMethods.push(SortMethod[key]);
        }
    }

    var dialog = new Window("dialog", LABELS.dialogTitle);
    dialog.orientation = "row";
    dialog.spacing = 15;

    var leftColumn = dialog.add("group");
    leftColumn.orientation = "column";
    leftColumn.alignChildren = ['fill', 'top'];

    // radioGroup を leftColumn 内に panel で追加し、タイトルを LABELS.sortMethod に設定
    var radioGroup = leftColumn.add("panel", undefined, LABELS.sortMethod);
    radioGroup.orientation = "column";
    radioGroup.alignChildren = ['left', 'top'];
    radioGroup.margins = [15, 20, 15, 10];

    // 許容差パネルを leftColumn 内に追加
    var toleranceGroup = leftColumn.add("panel", undefined, LABELS.tolerance);
    toleranceGroup.orientation = "column";
    toleranceGroup.alignChildren = ['fill', 'top'];
    toleranceGroup.margins = [15, 20, 15, 10];

    var tolRow = toleranceGroup.add("group");
    tolRow.orientation = "row";
    tolRow.alignChildren = ['left', 'center'];

    // スライダー初期値を自動計算値×1.1に設定
    var tempArtboards = [];
    for (var i = 0; i < doc.artboards.length; i++) {
        tempArtboards.push({
            artboardRect: doc.artboards[i].artboardRect
        });
    }
    var autoTol = calculateAutoTolerance(tempArtboards);
    var defaultTolerance = Math.round(autoTol * 1.1);

    var toleranceSlider = tolRow.add("slider", undefined, defaultTolerance, 0, 200);

    var toleranceValue = toleranceGroup.add("statictext", undefined, defaultTolerance);
    toleranceValue.justify = 'center';

    toleranceSlider.onChanging = function() {
        toleranceValue.text = Math.round(toleranceSlider.value);
    };

    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "column";
    buttonGroup.alignment = ['fill', 'top'];

    var radioButtons = [];
    var select = 0;
    for (var i = 0; i < sortMethods.length; i++) {
        var method = sortMethods[i];
        var rb = radioGroup.add("radiobutton", undefined, method.displayName);
        radioButtons.push(rb);
        if (settings.sortMethod.displayName && settings.sortMethod.displayName === method.displayName) {
            select = i;
        }
    }
    radioButtons[select].value = true;

    // ラジオボタンの onClick で「名前順」選択時に許容差をディム表示
    for (var i = 0; i < radioButtons.length; i++) {
        (function(index) {
            radioButtons[index].onClick = function() {
                var selectedLabel = sortMethods[index].displayName;
                var isNameSort = selectedLabel.indexOf("名前") !== -1 || selectedLabel.indexOf("Name") !== -1;
                toleranceSlider.enabled = !isNameSort;
                toleranceValue.enabled = !isNameSort;
            };
        })(i);
    }
    // 初期状態も反映
    (function() {
        var selectedLabel = sortMethods[select].displayName;
        var isNameSort = selectedLabel.indexOf("名前") !== -1 || selectedLabel.indexOf("Name") !== -1;
        toleranceSlider.enabled = !isNameSort;
        toleranceValue.enabled = !isNameSort;
    })();

    var okBtn = buttonGroup.add('button', undefined, LABELS.sort, {
        name: 'ok'
    });
    var cancelBtn = buttonGroup.add('button', undefined, LABELS.cancel, {
        name: 'cancel'
    });

    var btnWidth = Math.max(okBtn.preferredSize.width, cancelBtn.preferredSize.width);
    okBtn.preferredSize.width = btnWidth;
    cancelBtn.preferredSize.width = btnWidth;

    okBtn.onClick = function() {
        var tolerance = Math.round(toleranceSlider.value) || 0;
        for (var i = 0; i < radioButtons.length; i++) {
            if (radioButtons[i].value) {
                settings.sortMethod = sortMethods[i];
                // Top Left または Top Right の場合、許容差機能付きソートに置き換え
                if (
                    settings.sortMethod.displayName.indexOf("左上") !== -1 ||
                    settings.sortMethod.displayName.indexOf("Top Left") !== -1
                ) {
                    settings.sortMethod.sorter = function(_artboards, decimalPlaces) {
                        sortByPositionTopLeftWithTolerance(_artboards, decimalPlaces, tolerance);
                    };
                } else if (
                    settings.sortMethod.displayName.indexOf("右上") !== -1 ||
                    settings.sortMethod.displayName.indexOf("Top Right") !== -1
                ) {
                    settings.sortMethod.sorter = function(_artboards, decimalPlaces) {
                        sortByPositionTopRightWithTolerance(_artboards, decimalPlaces, tolerance);
                    };
                }
                // 上左・上右のときはスライダー値の tolerance を使う
                else if (
                    settings.sortMethod.displayName.indexOf("上左") !== -1 ||
                    settings.sortMethod.displayName.indexOf("Left Top") !== -1
                ) {
                    settings.sortMethod.sorter = function(_artboards, decimalPlaces) {
                        sortByPositionLeftTopWithTolerance(_artboards, decimalPlaces, tolerance);
                    };
                } else if (
                    settings.sortMethod.displayName.indexOf("上右") !== -1 ||
                    settings.sortMethod.displayName.indexOf("Right Top") !== -1
                ) {
                    settings.sortMethod.sorter = function(_artboards, decimalPlaces) {
                        sortByPositionRightTopWithTolerance(_artboards, decimalPlaces, tolerance);
                    };
                }
                break;
            }
        }
        dialog.close(1);
    };
// 横方向（left座標）の差から tolerance を自動計算
function calculateAutoToleranceHorizontal(_artboards) {
    var lefts = [];
    for (var i = 0; i < _artboards.length; i++) {
        lefts.push(_artboards[i].artboardRect[0]);
    }
    lefts.sort(function(a, b) {
        return a - b;
    }); // 左から右に並べる

    var diffs = [];
    for (var i = 1; i < lefts.length; i++) {
        var diff = Math.abs(lefts[i] - lefts[i - 1]);
        if (diff > 0) {
            diffs.push(diff);
        }
    }

    if (diffs.length === 0) {
        return 5; // 差が無ければデフォルト
    }

    var minDiff = Math.min.apply(null, diffs);
    return minDiff + 2;
}

    cancelBtn.onClick = function() {
        dialog.close(-1);
    };

    dialog.center();
    var result = dialog.show();
    if (result === 1) {
        settings.action(app.activeDocument, settings.sortMethod.sorter);
    }
}

function sortArtboards(doc, sortFunction) {
    var decimalPlaces = Math.pow(10, settings.decimalPlaces);

    // アートボード情報を配列に格納
    var _artboards = [];
    for (var i = 0; i < doc.artboards.length; i++) {
        var a = doc.artboards[i];
        _artboards.push({
            name: a.name,
            artboardRect: a.artboardRect,
            srcIndex: i
        });
    }

    // ソート実行
    sortFunction(_artboards, decimalPlaces);

    // アートボードを再配置しつつsrcIndexを更新（最適化の余地あり）
    for (var i = 0; i < _artboards.length; i++) {
        var s = _artboards[i];
        var a = doc.artboards[s.srcIndex];
        var b = doc.artboards.add(a.artboardRect);
        b.name = a.name;
        b.rulerOrigin = a.rulerOrigin;
        b.rulerPAR = a.rulerPAR;
        b.showCenter = a.showCenter;
        b.showCrossHairs = a.showCrossHairs;
        b.showSafeAreas = a.showSafeAreas;
        a.remove();

        // srcIndex更新を1回のループ内で実行
        for (var j = 0; j < _artboards.length; j++) {
            if (_artboards[j].srcIndex > s.srcIndex) {
                _artboards[j].srcIndex--;
            }
        }
    }
}

// 名前順ソート
function sortByName(_artboards) {
    _artboards.sort(function(a, b) {
        if (a.name < b.name) return -1;
        if (a.name > b.name) return 1;
        return 0;
    });
}

// 共通位置ソート関数
function sortByPosition(_artboards, decimalPlaces, primaryIndex, primaryAsc, secondaryIndex, secondaryAsc) {
    decimalPlaces = decimalPlaces || 1000;
    _artboards.sort(function(a, b) {
        var primary = Math.round((a.artboardRect[primaryIndex] - b.artboardRect[primaryIndex]) * decimalPlaces) / decimalPlaces;
        var secondary = Math.round((a.artboardRect[secondaryIndex] - b.artboardRect[secondaryIndex]) * decimalPlaces) / decimalPlaces;
        return (primaryAsc ? primary : -primary) || (secondaryAsc ? secondary : -secondary);
    });
}

// 位置順: 左上基準
function sortByPositionTopLeft(_artboards, decimalPlaces) {
    sortByPosition(_artboards, decimalPlaces, 1, false, 0, true);
}

// 位置順: 右上基準
function sortByPositionTopRight(_artboards, decimalPlaces) {
    sortByPosition(_artboards, decimalPlaces, 1, false, 0, false);
}

// 許容差付き 左上基準ソート
function sortByPositionTopLeftWithTolerance(_artboards, decimalPlaces, tolerance) {
    sortByPositionWithTolerance(_artboards, decimalPlaces, 1, false, 0, true, tolerance);
}

// 許容差付き 右上基準ソート
function sortByPositionTopRightWithTolerance(_artboards, decimalPlaces, tolerance) {
    sortByPositionWithTolerance(_artboards, decimalPlaces, 1, false, 0, false, tolerance);
}

// 許容差を考慮した共通位置ソート関数
function sortByPositionWithTolerance(_artboards, decimalPlaces, primaryIndex, primaryAsc, secondaryIndex, secondaryAsc, tolerance) {
    decimalPlaces = decimalPlaces || 1000;
    _artboards.sort(function(a, b) {
        var primaryA = Math.round(a.artboardRect[primaryIndex] * decimalPlaces) / decimalPlaces;
        var primaryB = Math.round(b.artboardRect[primaryIndex] * decimalPlaces) / decimalPlaces;
        var secondaryA = Math.round(a.artboardRect[secondaryIndex] * decimalPlaces) / decimalPlaces;
        var secondaryB = Math.round(b.artboardRect[secondaryIndex] * decimalPlaces) / decimalPlaces;

        // 上辺の差が tolerance 以内なら、secondary で比較
        if (Math.abs(primaryA - primaryB) <= tolerance) {
            return (secondaryAsc ? secondaryA - secondaryB : secondaryB - secondaryA);
        }
        return (primaryAsc ? primaryA - primaryB : primaryB - primaryA);
    });
}

// 位置順: 上左基準
function sortByPositionLeftTop(_artboards, decimalPlaces) {
    sortByPosition(_artboards, decimalPlaces, 0, true, 1, false);
}

// 位置順: 上右基準
function sortByPositionRightTop(_artboards, decimalPlaces) {
    sortByPosition(_artboards, decimalPlaces, 0, false, 1, false);
}

// 許容差自動計算関数
function calculateAutoTolerance(_artboards) {
    var tops = [];
    for (var i = 0; i < _artboards.length; i++) {
        tops.push(_artboards[i].artboardRect[1]);
    }
    tops.sort(function(a, b) {
        return b - a;
    }); // 上から下に並べる

    var diffs = [];
    for (var i = 1; i < tops.length; i++) {
        var diff = Math.abs(tops[i] - tops[i - 1]);
        if (diff > 0) {
            diffs.push(diff);
        }
    }

    if (diffs.length === 0) {
        return 5; // 差が無ければデフォルト
    }

    var minDiff = Math.min.apply(null, diffs);
    return minDiff + 2; // 少しマージンを加える
}
// 許容差付き 上左基準ソート
function sortByPositionLeftTopWithTolerance(_artboards, decimalPlaces, tolerance) {
    sortByPositionWithTolerance(_artboards, decimalPlaces, 0, true, 1, false, tolerance);
}

// 許容差付き 上右基準ソート
function sortByPositionRightTopWithTolerance(_artboards, decimalPlaces, tolerance) {
    sortByPositionWithTolerance(_artboards, decimalPlaces, 0, false, 1, false, tolerance);
}