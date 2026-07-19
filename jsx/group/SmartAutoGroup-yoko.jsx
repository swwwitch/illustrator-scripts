#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

- 選択オブジェクトを水平方向の行ごとに自動グループ化します。
- 縦のズレが「縦のズレ許容」以内なら、横に離れていても同じ行とみなします。

### 主な機能

- 水平方向グループ化（同じ行＝縦が揃った横並びを検出。横の距離は問わない）
- 「縦のズレ許容」スライダーによる行判定の調整
- グループ化後の自動選択
- 未グループオブジェクトの再実行確認
- 日本語／英語インターフェース対応

### 処理の流れ

1. ダイアログで「縦のズレ許容」を設定
2. DFS探索で縦のズレが許容内のオブジェクトを同じ行として連結
3. Illustrator標準 group コマンドで結合
4. グループ化されなかったオブジェクトがある場合に再実行を促す

### 更新履歴

- v1.0.0 (20250611) : 初期バージョン
- v1.1.0 (20260609) : 水平方向のみに特化（モード選択を廃止）。判定を「縦のズレ許容」しきい値に変更し、横の距離は無制限に
- v1.2.0 (20260609) : 「アートボードごと」チェックボックスを追加（ONで別アートボードのオブジェクトは同じ行でも別グループにする）。選択が複数アートボードにまたがる場合のみデフォルトON、アートボードが1つだけならディム表示。グループ化実行中にプログレスバーを表示。スライダーをラベルの下の行に配置

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartAutoGroup-yoko";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function() {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 縦のズレ許容の初期値・範囲（pt） / Default and range of vertical tolerance (pt) */
    var DEFAULT_VERTICAL_TOLERANCE = 10;
    var TOLERANCE_MIN = 0;
    var TOLERANCE_MAX = 200;

    /* 「アートボードごと」の初期値 / Default of "group within each artboard" */
    var DEFAULT_GROUP_PER_ARTBOARD = false;

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 表示言語を判定（ja で始まれば日本語、それ以外は英語） / Detect UI language */
    function getCurrentLang() {
        return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
    }

    var lang = getCurrentLang();

    /* 日英ラベル定義（カテゴリ別） / Bilingual labels grouped by category */
    var LABELS = {
        dialog: {
            title: {
                ja: "水平方向グループ化",
                en: "Horizontal Grouping"
            }
        },
        panel: {
            condition: {
                ja: "グループ化の条件",
                en: "Grouping Condition"
            }
        },
        field: {
            verticalTolerance: {
                ja: "縦のズレ許容：",
                en: "Vertical tolerance:"
            },
            verticalToleranceTip: {
                ja: "選択オブジェクトの上下のズレがこの値（pt）以内なら、横にどれだけ離れていても同じ行とみなしてグループ化します。0 にすると縦が重なっているものだけをまとめます。",
                en: "Objects whose vertical offset is within this value (pt) are treated as the same row and grouped regardless of horizontal distance. Set to 0 to group only vertically overlapping objects."
            },
            groupPerArtboard: {
                ja: "アートボードごと",
                en: "Within each artboard"
            },
            groupPerArtboardTip: {
                ja: "ONにすると、別のアートボードにあるオブジェクトは同じ行でも別々のグループにします。オブジェクトの中心がどのアートボードに含まれるかで判定します。",
                en: "When on, objects located on different artboards are placed in separate groups even if they are in the same row. The artboard is determined by each object's center point."
            }
        },
        button: {
            group: {
                ja: "グループ化",
                en: "Group"
            },
            cancel: {
                ja: "キャンセル",
                en: "Cancel"
            }
        },
        progress: {
            title: {
                ja: "グループ化中…",
                en: "Grouping…"
            }
        },
        message: {
            result: {
                ja: "○個のグループ化を行いました。",
                en: " groups have been created."
            },
            retry: {
                ja: "グループ化されなかったオブジェクトが {0} 個あります。\n再実行しますか？",
                en: "{0} objects were not grouped.\nDo you want to retry?"
            }
        }
    };

    /* ラベル取得（ドット区切りキーで現在の言語の文字列を返す） / Get the label by dotted key */
    function L(key) {
        var parts = key.split('.');
        var entry = LABELS;
        for (var i = 0; i < parts.length; i++) {
            entry = entry[parts[i]];
        }
        return entry[lang];
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* visibleBounds はポイント（pt）単位。スライダー値もそのまま pt として扱う。 */
    var UNIT_LABEL = "pt";

    /* スライダー値を表示用文字列に整形 / Format a slider value for display */
    function formatValue(value) {
        return Math.round(value) + " " + UNIT_LABEL;
    }

    // =========================================
    // 状態 / State
    // =========================================

    /* 現在の縦のズレ許容（行判定に使用） / Current vertical tolerance used for row detection */
    var verticalTolerance;

    /* アートボードごとにグループ化するか / Whether to group within each artboard */
    var groupPerArtboard;

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /* ラベル行＋スライダー行（ラベルは上、スライダー＋数値は次の行）を作るヘルパー
       / Build a caption line, then a slider + value line below it */
    function addSliderRow(parentPanel, captionLabel, initialValue, helpTip) {
        var captionText = parentPanel.add("statictext", undefined, captionLabel);

        var row = parentPanel.add("group");
        row.orientation = "row";
        row.alignChildren = "center";

        var slider = row.add("slider", undefined, initialValue, TOLERANCE_MIN, TOLERANCE_MAX);
        slider.preferredSize.width = 200;

        var valueText = row.add("statictext", undefined, formatValue(initialValue));
        valueText.characters = 5;

        if (helpTip) {
            captionText.helpTip = helpTip;
            slider.helpTip = helpTip;
        }

        slider.onChanging = function() {
            valueText.text = formatValue(slider.value);
        };

        return slider;
    }

    /* ダイアログUIの表示とユーザー選択取得 / Show the dialog and return the user's choice */
    function showDialog(initialTolerance, initialPerArtboard, hasMultipleArtboards) {
        var dialog = new Window("dialog", L('dialog.title') + " " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = "fill";
        dialog.margins = [15, 20, 15, 10];

        var conditionPanel = dialog.add("panel", undefined, L('panel.condition'));
        conditionPanel.orientation = "column";
        conditionPanel.alignChildren = "left";
        conditionPanel.margins = [15, 20, 15, 10];

        var toleranceSlider = addSliderRow(
            conditionPanel,
            L('field.verticalTolerance'),
            initialTolerance,
            L('field.verticalToleranceTip')
        );

        var perArtboardCheck = conditionPanel.add("checkbox", undefined, L('field.groupPerArtboard'));
        perArtboardCheck.value = initialPerArtboard;
        perArtboardCheck.helpTip = L('field.groupPerArtboardTip');
        // アートボードが1つしかない場合は意味がないのでディム表示
        perArtboardCheck.enabled = hasMultipleArtboards;

        var buttonRow = dialog.add("group");
        buttonRow.orientation = "row";
        buttonRow.alignment = "right";
        buttonRow.margins = [0, 6, 0, 6];
        var cancelButton = buttonRow.add("button", undefined, L('button.cancel'));
        var groupButton = buttonRow.add("button", undefined, L('button.group'), {
            name: "ok"
        });

        cancelButton.onClick = function() {
            dialog.close();
        };
        groupButton.onClick = function() {
            verticalTolerance = toleranceSlider.value;
            groupPerArtboard = perArtboardCheck.value;
            dialog.close(1);
        };

        return dialog.show();
    }

    // =========================================
    // 行判定・グループ抽出 / Row Detection & Grouping
    // =========================================

    /* 2つのバウンディングボックスが同じ横の行に並んでいるかを判定する。
       縦のズレ（重なっていなければそのすき間）が verticalTolerance 以内なら true。
       横方向にどれだけ離れていても、縦が揃っていれば同じ行とみなす。 */
    function isSameRow(boundsA, boundsB) {
        var topA = boundsA[1],
            bottomA = boundsA[3];
        var topB = boundsB[1],
            bottomB = boundsB[3];

        var verticalGap = Math.max(0, Math.max(bottomB - topA, bottomA - topB));
        return verticalGap <= verticalTolerance;
    }

    /* バウンディングボックスの中心が含まれるアートボードの番号を返す（該当なしは -1） */
    function getArtboardIndexForBounds(bounds, artboardRects) {
        var centerX = (bounds[0] + bounds[2]) / 2;
        var centerY = (bounds[1] + bounds[3]) / 2;
        for (var i = 0; i < artboardRects.length; i++) {
            var rect = artboardRects[i]; // [left, top, right, bottom]
            if (centerX >= rect[0] && centerX <= rect[2] && centerY <= rect[1] && centerY >= rect[3]) {
                return i;
            }
        }
        return -1;
    }

    /* 各オブジェクトが属するアートボード番号の配列を返す（items と同じ並び） */
    function computeArtboardIndices(items) {
        var artboards = app.activeDocument.artboards;
        var artboardRects = [];
        for (var i = 0; i < artboards.length; i++) {
            artboardRects.push(artboards[i].artboardRect);
        }
        var indices = [];
        for (var k = 0; k < items.length; k++) {
            indices.push(getArtboardIndexForBounds(items[k].visibleBounds, artboardRects));
        }
        return indices;
    }

    /* 選択オブジェクトが2つ以上のアートボードにまたがっているかを判定する */
    function selectionSpansMultipleArtboards(items) {
        var indices = computeArtboardIndices(items);
        var seen = {};
        var count = 0;
        for (var i = 0; i < indices.length; i++) {
            if (indices[i] < 0) continue; // どのアートボードにも乗っていないものは無視
            if (!seen[indices[i]]) {
                seen[indices[i]] = true;
                count++;
                if (count > 1) return true;
            }
        }
        return false;
    }

    /* DFS探索で、起点と同じ行に並ぶオブジェクトを1つの行グループに集める */
    function collectSameRowMembers(startIndex, items, artboardIndices, visited, rowMembers) {
        visited[startIndex] = true;
        rowMembers.push(items[startIndex]);
        var startBounds = items[startIndex].visibleBounds;
        for (var j = 0; j < items.length; j++) {
            if (visited[j]) continue;
            // アートボードごとにグループ化する場合、別アートボードのオブジェクトは連結しない
            if (groupPerArtboard && artboardIndices[startIndex] !== artboardIndices[j]) continue;
            if (isSameRow(startBounds, items[j].visibleBounds)) {
                collectSameRowMembers(j, items, artboardIndices, visited, rowMembers);
            }
        }
    }

    /* 選択オブジェクトを「同じ行」ごとの配列にまとめて返す（DFSによる連結成分抽出） */
    function collectHorizontalRows(items) {
        var artboardIndices = groupPerArtboard ? computeArtboardIndices(items) : [];
        var rows = [];
        var visited = []; // 未訪問は undefined（falsy）として扱う
        for (var i = 0; i < items.length; i++) {
            if (visited[i]) continue;
            var rowMembers = [];
            collectSameRowMembers(i, items, artboardIndices, visited, rowMembers);
            rows.push(rowMembers);
        }
        return rows;
    }

    // =========================================
    // グループ化処理 / Grouping Actions
    // =========================================

    /* 指定オブジェクト群を選択してグループ化し、元のレイヤーに残して新グループを返す */
    function createGroupFromItems(rowMembers) {
        var originalLayer = rowMembers[0].layer;

        app.executeMenuCommand('deselectall');
        for (var j = 0; j < rowMembers.length; j++) {
            rowMembers[j].selected = true;
        }
        app.executeMenuCommand('group');

        var newGroup = app.activeDocument.selection[0];
        newGroup.layer = originalLayer; // グループ化で最前面レイヤーへ移るのを元に戻す
        return newGroup;
    }

    /* 進捗バーのパレットを作成して表示する / Create and show a progress palette */
    function createProgress(maxValue) {
        var win = new Window("palette", L('dialog.title') + " " + SCRIPT_VERSION);
        win.orientation = "column";
        win.alignChildren = "fill";
        win.margins = 16;

        win.add("statictext", undefined, L('progress.title'));

        var bar = win.add("progressbar", undefined, 0, maxValue);
        bar.preferredSize.width = 280;

        win.show();
        win.update();
        return { win: win, bar: bar };
    }

    /* 指定したオブジェクト群だけを選択状態にする */
    function selectItems(objects) {
        app.activeDocument.selection = null;
        for (var i = 0; i < objects.length; i++) {
            objects[i].selected = true;
        }
    }

    /* 水平方向（同じ行）でのグループ化を実行する */
    function groupHorizontally() {
        if (!app.documents.length) return;
        var items = app.activeDocument.selection;
        if (!items || items.length === 0) return;

        var rows = collectHorizontalRows(items);

        var progress = createProgress(rows.length);

        var newGroups = [];
        var ungroupedCount = 0; // 1個だけの行＝グループ化されない単独オブジェクト
        for (var i = 0; i < rows.length; i++) {
            if (rows[i].length > 1) {
                newGroups.push(createGroupFromItems(rows[i]));
            } else {
                ungroupedCount++; // 1個だけの行＝グループ化されない単独オブジェクト
            }
            progress.bar.value = i + 1;
            progress.win.update();
        }

        progress.win.close();

        // グループ化されなかった単独オブジェクトがあれば再実行を促す
        if (ungroupedCount > 0) {
            var retryMessage = L('message.retry').replace("{0}", ungroupedCount);
            if (confirm(retryMessage)) {
                main(verticalTolerance, groupPerArtboard);
                return;
            }
        }

        app.redraw();
        alert(L('message.result').replace("○", newGroups.length));
        selectItems(newGroups);
        return newGroups;
    }

    // =========================================
    // エントリーポイント / Entry Point
    // =========================================

    function main(initialTolerance, initialPerArtboard) {
        verticalTolerance = (typeof initialTolerance === "number") ? initialTolerance : DEFAULT_VERTICAL_TOLERANCE;
        if (!app.documents.length) return;
        var selection = app.activeDocument.selection;
        if (!selection || selection.length === 0) return;
        // アートボードが1つだけならチェックボックスは無効（OFF固定）
        var hasMultipleArtboards = app.activeDocument.artboards.length > 1;
        if (typeof initialPerArtboard === "boolean") {
            groupPerArtboard = initialPerArtboard && hasMultipleArtboards;
        } else {
            // 選択が複数アートボードにまたがる場合だけデフォルトON
            groupPerArtboard = hasMultipleArtboards && selectionSpansMultipleArtboards(selection);
        }
        if (showDialog(verticalTolerance, groupPerArtboard, hasMultipleArtboards) !== 1) return;
        // verticalTolerance / groupPerArtboard は showDialog() 内で更新される
        groupHorizontally();
    }

    // 第2引数は省略（undefined）。初回は選択がアートボードをまたぐかで自動判定する
    main(DEFAULT_VERTICAL_TOLERANCE);

})();
