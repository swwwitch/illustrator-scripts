#target illustrator
#targetengine "SortTextByColumnEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

テキストフレーム内のタブ区切りテキストを、指定した列の値で並び替えます。
数値・文字列の昇順／降順とランダム順に対応し、見出し行は自動で判定します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SortTextByColumn.md

### Overview

Sorts the tab-separated text inside a text frame by the values in a chosen column.
Numeric and string sorting (ascending or descending) and random order are supported, and a header row is detected automatically.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SortTextByColumn.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SortTextByColumn";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.10";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-06-15";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SortTextByColumn.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SortTextByColumn.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 区切り文字 / Separators
    // =========================================

    var LINE_SEPARATOR = "\r"; /* 行（段落）の区切り / line (paragraph) separator */
    var CELL_SEPARATOR = "\t"; /* 列の区切り / column separator */

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ボタン行（再利用パーツ） / Button row (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ダイアログを作る関数より前）に貼る。
    //    識別子は BUTTON_ROW_* / addButtonRow
    // 2. ダイアログの最後で行を作り、ボタンは btn 接頭辞の変数で左右のグループに足す（キャンセル → OK の順）
    //      var buttonRow = addButtonRow(dialog);
    //      var btnPreferences = buttonRow.leftGroup.add("button", undefined, getLabel("button.preferences"));
    //      var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    //      var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    //    左右中央に並べるときは addButtonRow(dialog, { centered: true }) にして、buttonRow.rowGroup に直接足す
    // 3. 行の上の余白は BUTTON_ROW_TOP_MARGIN で決める。左右の余白はダイアログの margins に任せる
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */

    /**
     * ダイアログ下部のボタン行を作る。
     * 通常は「左のグループ・伸びるスペーサー・右のグループ」、centered なら行そのものを左右中央に置く
     * @param {Window|Group|Panel} parent - 行を足す先（ふつうはダイアログ）
     * @param {Object} [rowOptions] - { centered: true } で左右中央に並べる
     * @returns {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} 行と左右のグループ（centered のときは左右が null）
     */
    function addButtonRow(parent, rowOptions) {
        var isCentered = !!(rowOptions && rowOptions.centered);
        var btnRowGroup = parent.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.spacing = BUTTON_ROW_SPACING;

        if (isCentered) {
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            return { rowGroup: btnRowGroup, leftGroup: null, rightGroup: null };
        }

        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        btnLeftGroup.spacing = BUTTON_ROW_SPACING;

        /* 余りの幅を吸って、右のグループを右端に寄せる / Absorbs the extra width so the right group sits at the right edge */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.spacing = BUTTON_ROW_SPACING;

        return { rowGroup: btnRowGroup, leftGroup: btnLeftGroup, rightGroup: btnRightGroup };
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ボタン行（再利用パーツ）ここまで / End of the reusable button row
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内のローカライズ節（LABELS の直前）に貼る。
    //    uiLang を使うコード（StepperButtons・LinkToggle の部品など）より前に置く
    // 2. 識別子は uiLang / getCurrentLang / getLabel / labelText / labelValueText / fillLabelPlaceholders。
    //    同じ役割の既存の関数・変数（getCurrentLanguage、currentLanguage、formatLabel など）は消して、これに寄せる
    // 3. 呼び出しはどちらの形でもよい（混ぜてもよい）
    //      getLabel("dialog.title")        … パス
    //      getLabel(LABELS.dialog.title)   … { ja, en } を直接
    //      getLabel("alert.count", { count: 3 })  … "{count} 個" の {count} を差し込む
    //      getLabel("alert.range", [1, 10])       … "%1〜%2" の %1・%2 を差し込む
    //      labelText("fieldLabel.width")   … 末尾にコロン（日本語は全角「：」、英語は半角「:」）
    //      labelValueText("message.count", 5) … 「件数：5」／「Count: 5」（値が続く1行。英語はコロンのあとに空白）
    // 4. 見つからないパスはパスの文字列をそのまま返す（表示で気づけるように）。{ ja, en } が無いときは空文字
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    /**
     * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /**
     * LABELS から今の UI 言語の文言を取り出す。
     * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
     * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
     * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
     */
    function getLabel(labelRef, placeholderValues) {
        var labelEntry = labelRef;
        if (typeof labelRef === "string") {
            var labelPathKeys = labelRef.split(".");
            labelEntry = LABELS;
            for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
                labelEntry = labelEntry[labelPathKeys[i]];
            }
        }
        var labelString;
        if (typeof labelEntry === "string") labelString = labelEntry;
        else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
        else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
        else return (typeof labelRef === "string") ? labelRef : "";
        return fillLabelPlaceholders(String(labelString), placeholderValues);
    }

    /**
     * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
    }

    /**
     * 文言の {name} や %1 に値を差し込む
     * @param {string} labelString - 文言
     * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
     * @returns {string} 差し込んだ文言
     */
    function fillLabelPlaceholders(labelString, placeholderValues) {
        if (placeholderValues == null) return labelString;
        if (placeholderValues instanceof Array) {
            /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
            for (var i = placeholderValues.length; i >= 1; i--) {
                labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
            }
            return labelString;
        }
        for (var placeholderKey in placeholderValues) {
            if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
            labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
        }
        return labelString;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /* UIラベル（表示順：列 → 順序 → 見出し） */
    var LABELS = {
        dialog: {
            title: { ja: "列を基準に並べ替え", en: "Sort by Column" }
        },
        panel: {
            sortColumn: { ja: "ソート対象の列", en: "Sort Target Column" },
            sortOrder:  { ja: "ソート方法", en: "Sort Order" }
        },
        radio: {
            column:     { ja: "【列%1】", en: "[Row %1] " },
            ascending:  { ja: "昇順", en: "Ascending" },
            descending: { ja: "降順", en: "Descending" },
            random:     { ja: "ランダム", en: "Random" }
        },
        checkbox: {
            header:     { ja: "1行目を見出し行として扱う", en: "Treat first row as header" }
        },
        button: {
            ok:         { ja: "OK", en: "OK" },
            cancel:     { ja: "キャンセル", en: "Cancel" }
        },
        tooltip: {
            column:     { ja: "この列の値を基準に、行を並べ替えます。", en: "The rows are sorted by the values in this column." },
            ascending:  { ja: "値の小さい順（あいうえお順）に並べます。", en: "Sorts in ascending order." },
            descending: { ja: "値の大きい順に並べます。", en: "Sorts in descending order." },
            random:     { ja: "値と関係なく、行の順序をシャッフルします。", en: "Shuffles the rows regardless of their values." },
            header:     { ja: "1行目は並べ替えず、見出しとして先頭に残します。", en: "Keeps the first row at the top instead of sorting it." }
        }
    };

    // =========================================
    // 並べ替え / Sorting
    // =========================================

    /**
     * テキストから最初に現れる数値（整数・小数・カンマ区切り）を取り出す
     * @param {string} cellText - セルの文字列
     * @returns {number|null} 数値。見つからなければ null
     */
    function extractFirstNumber(cellText) {
        var numberMatch = cellText.match(/[\d,]+(\.\d+)?/);
        if (numberMatch) {
            return parseFloat(numberMatch[0].replace(/,/g, ""));
        }
        return null;
    }

    /**
     * 配列の順序をシャッフルする（Fisher–Yates、直接書き換える）
     * @param {Object[]} lineEntries - { line, key } の配列
     * @returns {void}
     */
    function shuffleLineEntries(lineEntries) {
        for (var i = lineEntries.length - 1; i > 0; i--) {
            var j = Math.floor(Math.random() * (i + 1));
            var swapEntry = lineEntries[i];
            lineEntries[i] = lineEntries[j];
            lineEntries[j] = swapEntry;
        }
    }

    /**
     * 数値キーを昇順・降順で比べる
     * @param {Object} a - { line, key } の要素
     * @param {Object} b - { line, key } の要素
     * @param {string} sortOrder - "asc" / "desc"
     * @returns {number} 並べ替え用の差
     */
    function compareNumericKeys(a, b, sortOrder) {
        return sortOrder === "asc" ? (a.key - b.key) : (b.key - a.key);
    }

    /**
     * 文字列キーを昇順・降順で比べる（大文字小文字は区別しない）
     * @param {Object} a - { line, key } の要素
     * @param {Object} b - { line, key } の要素
     * @param {string} sortOrder - "asc" / "desc"
     * @returns {number} -1 / 0 / 1
     */
    function compareStringKeys(a, b, sortOrder) {
        var aText = String(a.key).toLowerCase();
        var bText = String(b.key).toLowerCase();
        if (sortOrder === "asc") {
            return aText < bText ? -1 : aText > bText ? 1 : 0;
        }
        return aText > bText ? -1 : aText < bText ? 1 : 0;
    }

    /**
     * 指定列と並び順で行を並べ替える。セルに数値があれば数値、無ければ文字列で比べる
     * @param {string[]} lines - 並べ替える行
     * @param {number} columnIndex - 基準にする列（0 始まり）
     * @param {string} sortOrder - "asc" / "desc" / "random"
     * @returns {string[]} 並べ替えた行
     */
    function generateSortedLines(lines, columnIndex, sortOrder) {
        var lineEntries = [];
        for (var i = 0; i < lines.length; i++) {
            var cellText = lines[i].split(CELL_SEPARATOR)[columnIndex] || "";
            var firstNumber = extractFirstNumber(cellText);
            lineEntries.push({
                line: lines[i],
                key: firstNumber !== null ? firstNumber : cellText
            });
        }

        if (sortOrder === "random") {
            shuffleLineEntries(lineEntries);
        } else {
            lineEntries.sort(function (a, b) {
                if (typeof a.key === "number" && typeof b.key === "number") {
                    return compareNumericKeys(a, b, sortOrder);
                }
                return compareStringKeys(a, b, sortOrder);
            });
        }

        var sortedLines = [];
        for (var j = 0; j < lineEntries.length; j++) {
            sortedLines.push(lineEntries[j].line);
        }
        return sortedLines;
    }

    // =========================================
    // 見出しと既定の列 / Header and default column
    // =========================================

    /**
     * 各列の値を先頭から3件ずつ集める（列のラジオボタンに表示する）
     * @param {number} columnCount - 列の数
     * @param {string[]} lines - 行
     * @returns {string[][]} 列ごとの値
     */
    function collectColumnPreviews(columnCount, lines) {
        var columnPreviews = [];
        for (var i = 0; i < columnCount; i++) {
            columnPreviews[i] = [];
        }
        for (var r = 0; r < lines.length; r++) {
            var cells = lines[r].split(CELL_SEPARATOR);
            for (var c = 0; c < columnCount; c++) {
                if (columnPreviews[c].length < 3 && cells[c]) {
                    columnPreviews[c].push(cells[c]);
                }
            }
        }
        return columnPreviews;
    }

    /**
     * 1行目が見出しらしいかを調べる（1行目は数値でなく、2行目以降に数値がある列があるか）
     * @param {number} columnCount - 列の数
     * @param {string[]} lines - 行
     * @returns {boolean} 見出しらしければ true
     */
    function looksLikeHeaderRow(columnCount, lines) {
        for (var i = 0; i < columnCount; i++) {
            var hasNumericBelow = false;
            for (var r = 1; r < lines.length; r++) {
                if (extractFirstNumber(lines[r].split(CELL_SEPARATOR)[i]) !== null) {
                    hasNumericBelow = true;
                    break;
                }
            }
            var topIsNumber = extractFirstNumber(lines[0].split(CELL_SEPARATOR)[i]) !== null;
            if (!topIsNumber && hasNumericBelow) return true;
        }
        return false;
    }

    /**
     * 最初に選んでおく列を決める（先頭3行がすべて数値の最初の列。無ければ1列目）
     * @param {number} columnCount - 列の数
     * @param {string[]} lines - 行
     * @param {number} startRow - 調べ始める行（見出しありなら 1）
     * @returns {number} 列の番号（0 始まり）
     */
    function findDefaultColumn(columnCount, lines, startRow) {
        for (var c = 0; c < columnCount; c++) {
            var numericOnly = true;
            for (var r = startRow; r < Math.min(startRow + 3, lines.length); r++) {
                if (extractFirstNumber(lines[r].split(CELL_SEPARATOR)[c]) === null) {
                    numericOnly = false;
                    break;
                }
            }
            if (numericOnly) return c;
        }
        return 0;
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子は DIALOG_* / prepareDialogWindow / *DialogLeft* / getSelectionViewSpan の名前
    // 2. スクリプトの先頭（#target の次の行）に #targetengine "<SCRIPT_NAME>Engine" を置く。
    //    #targetengine が無いと $.global が実行ごとに消え、位置を覚えられない。すでにあればそのまま使う
    // 3. ダイアログの show() の直前で prepareDialogWindow(dialog, SCRIPT_NAME) を呼ぶ。
    //    それまでに入れた onShow / onMove / onClose はそのまま生かし、あとに位置の復元・記録をつなぐ
    //      prepareDialogWindow(mainDialog, SCRIPT_NAME);
    //      var dialogResult = mainDialog.show();
    //    同じスクリプトで複数のダイアログを開くときは、2つ目以降のキーを変える（SCRIPT_NAME + "_colorPicker" など）
    //    同じダイアログを何度も開くときも、毎回 show() の直前で呼んでよい（2回目からは選択範囲を測り直すだけ）
    // 4. 初めて開くとき（記録が無いとき）は、スクリプト側の配置（中央・オフセットなど）がそのまま効く
    // 5. 開く位置が選択中のオブジェクトに重なりそうなら左右の反対側へずらす（Illustrator のみ）。
    //    ずらした位置は記録せず、ユーザーが動かしたときだけ記録する
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var DIALOG_OPACITY = 0.98;       /* ダイアログの不透明度 / dialog opacity */
    var DIALOG_AVOID_MARGIN = 60;    /* 選択範囲の推定位置の両側に取る余裕（px）/ margin on each side of the estimated selection (px) */
    var DIALOG_AVOID_MAX_ITEMS = 100; /* 選択範囲を測るオブジェクトの上限 / max items measured for the selection bounds */

    /**
     * ダイアログの不透明度を設定し、前回閉じた位置で開いて、動かした位置を記録するようにする。
     * 開く位置が選択中のオブジェクトに重なりそうなときは、左右の反対側へずらす（Illustrator のみ）。
     * 既存の onShow / onMove / onClose は先に呼んでから、位置の復元・記録を行う。
     * @param {Window} dialog - 対象のダイアログ
     * @param {string} storageKey - 位置を覚えるキー（ふつうは SCRIPT_NAME）
     * @returns {void}
     */
    function prepareDialogWindow(dialog, storageKey) {
        /* 同じダイアログを開き直すときは、選択範囲を測り直すだけにする（ハンドラーを重ねない）
           When the same dialog is shown again, only re-measure the selection (don't stack handlers) */
        if (dialog.dialogWindowState) {
            dialog.dialogWindowState.selectionSpan = getSelectionViewSpan();
            dialog.dialogWindowState.avoidedLocation = null;
            return;
        }
        var locationKey = "__" + storageKey + "_DialogLocation";
        var previousOnShow = dialog.onShow;
        var previousOnMove = dialog.onMove;
        var previousOnClose = dialog.onClose;
        var windowState = {
            selectionSpan: getSelectionViewSpan(), /* 選択範囲は show() の前に測る / measured before show() */
            screenWidth: null,                     /* 最初に開いたときに推定する / estimated on the first show */
            avoidedLocation: null                  /* 避けるためにずらした位置（記録しない）/ location set to avoid the selection (not remembered) */
        };
        dialog.dialogWindowState = windowState;

        dialog.opacity = DIALOG_OPACITY;

        /* 今の位置を記録する / Remember the current location */
        function rememberDialogLocation() {
            var currentLocation = [dialog.location[0], dialog.location[1]];
            var avoidedLocation = windowState.avoidedLocation;
            if (avoidedLocation && currentLocation[0] === avoidedLocation[0] && currentLocation[1] === avoidedLocation[1]) return;
            $.global[locationKey] = currentLocation;
        }

        dialog.onShow = function () {
            /* 最初に開くときの既定の位置は画面の横中央なので、画面の幅を逆算できる。2回目からは前回の位置なので使い回す
               On the first show the default location is centered horizontally, which gives the screen width; reuse it afterwards */
            if (windowState.screenWidth === null) windowState.screenWidth = dialog.location[0] * 2 + dialog.bounds.width;
            if (previousOnShow) previousOnShow.apply(this, arguments);
            /* $.screens は実際の画面の大きさと合わない（Mac で 1280×524 など）ので、画面内かは判定しない
               $.screens does not match the real display (e.g. 1280x524 on a Mac), so no on-screen check */
            var savedLocation = $.global[locationKey];
            if (savedLocation) dialog.location = [savedLocation[0], savedLocation[1]];
            if (windowState.selectionSpan) {
                var avoidLeft = findDialogLeftAvoidingSelection(dialog.location[0], dialog.bounds.width, windowState.screenWidth, windowState.selectionSpan);
                if (avoidLeft !== null) {
                    dialog.location = [avoidLeft, dialog.location[1]];
                    /* 代入後の値で比べる（丸められることがある）/ Compare with the value after assignment, which may be rounded */
                    windowState.avoidedLocation = [dialog.location[0], dialog.location[1]];
                }
            }
        };
        dialog.onMove = function () {
            if (previousOnMove) previousOnMove.apply(this, arguments);
            rememberDialogLocation();
        };
        dialog.onClose = function () {
            rememberDialogLocation();
            /* false を返すと閉じるのを取りやめるので、戻り値は元の onClose のものを返す
               Returning false cancels the close, so pass the original onClose result through */
            if (previousOnClose) return previousOnClose.apply(this, arguments);
        };
    }

    /**
     * 選択中のオブジェクトが、ドキュメントの表示域の左端から画面上で何 px の範囲にあるかを返す。
     * @returns {{left: number, right: number, viewWidth: number}|null} 選択が無い・測れないときは null
     */
    function getSelectionViewSpan() {
        try {
            if (app.name !== "Adobe Illustrator" || !app.documents.length) return null;
            var targetDoc = app.activeDocument;
            var selectedItems = targetDoc.selection;
            if (!selectedItems || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
            var itemCount = Math.min(selectedItems.length, DIALOG_AVOID_MAX_ITEMS);
            var spanLeft = Infinity;
            var spanRight = -Infinity;
            for (var i = 0; i < itemCount; i++) {
                var itemBounds = selectedItems[i].visibleBounds;
                if (itemBounds[0] < spanLeft) spanLeft = itemBounds[0];
                if (itemBounds[2] > spanRight) spanRight = itemBounds[2];
            }
            var activeView = targetDoc.activeView; /* 複数ウィンドウで開いていても今のウィンドウ / the current window even with multiple windows */
            var viewBounds = activeView.bounds;
            var zoom = activeView.zoom;
            var viewWidth = (viewBounds[2] - viewBounds[0]) * zoom;
            /* 表示域の外にはみ出した部分は数えない / Ignore the part outside the view */
            var left = Math.max(0, (spanLeft - viewBounds[0]) * zoom);
            var right = Math.min(viewWidth, (spanRight - viewBounds[0]) * zoom);
            if (right <= left) return null;
            return { left: left, right: right, viewWidth: viewWidth };
        } catch (e) {
            /* テキスト編集中など測れないときは避けない / Do not avoid when it cannot be measured, e.g. while editing text */
            return null;
        }
    }

    /**
     * ダイアログが選択範囲に重なるなら、重ならない左端の位置を返す。
     * 表示域は画面の横中央にあるとみなし、ずれは DIALOG_AVOID_MARGIN で吸収する。
     * @param {number} dialogLeft - 今のダイアログの左端
     * @param {number} dialogWidth - ダイアログの幅
     * @param {number} screenWidth - 画面の幅
     * @param {{left: number, right: number, viewWidth: number}} selectionSpan - getSelectionViewSpan() の結果
     * @returns {number|null} ずらした左端。重ならない・どちらにも収まらないときは null
     */
    function findDialogLeftAvoidingSelection(dialogLeft, dialogWidth, screenWidth, selectionSpan) {
        var viewLeft = (screenWidth - selectionSpan.viewWidth) / 2;
        var avoidLeft = viewLeft + selectionSpan.left - DIALOG_AVOID_MARGIN;
        var avoidRight = viewLeft + selectionSpan.right + DIALOG_AVOID_MARGIN;
        if (dialogLeft + dialogWidth <= avoidLeft || dialogLeft >= avoidRight) return null;

        var leftSideLeft = avoidLeft - dialogWidth;   /* 選択範囲の左に置くとき / placed left of the selection */
        var rightSideLeft = avoidRight;               /* 選択範囲の右に置くとき / placed right of the selection */
        var fitsLeft = leftSideLeft >= 0;
        var fitsRight = rightSideLeft + dialogWidth <= screenWidth;
        /* 選択範囲が画面の右寄りなら左へ、左寄りなら右へ逃がす / Move away from the side the selection leans to */
        var preferLeft = (avoidLeft + avoidRight) / 2 > screenWidth / 2;
        if (preferLeft && fitsLeft) return leftSideLeft;
        if (fitsRight) return rightSideLeft;
        if (fitsLeft) return leftSideLeft;
        return null;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 並び替えの列、順序、見出し行の有無を選ぶダイアログを出す
     * @param {number} columnCount - 列の数
     * @param {string[]} lines - 行
     * @param {boolean} hasHeaderCandidate - 1行目だけタブが無く、見出しと判断済みか
     * @returns {Object|null} { column, order, useHeader }。キャンセル時は null
     */
    function showSortOptionsDialog(columnCount, lines, hasHeaderCandidate) {
        var sortDialog = new Window("dialog", getLabel("dialog.title"));
        sortDialog.orientation = "column";
        sortDialog.alignChildren = "left";

        var columnPanel = sortDialog.add("panel", undefined, getLabel("panel.sortColumn"));
        columnPanel.orientation = "column";
        columnPanel.alignChildren = "left";
        columnPanel.margins = [10, 25, 10, 10];

        var columnRadioGroup = columnPanel.add("group");
        columnRadioGroup.orientation = "column";
        columnRadioGroup.alignChildren = "left";

        /* 列ごとのラジオボタンに先頭3件の値を表示 / show the first three values on each column radio */
        var columnPreviews = collectColumnPreviews(columnCount, lines);
        var columnRadios = [];
        for (var i = 0; i < columnCount; i++) {
            var columnLabel = getLabel("radio.column").replace("%1", i + 1) + columnPreviews[i].join(", ") + "…";
            var columnRadio = columnRadioGroup.add("radiobutton", undefined, columnLabel);
            columnRadio.helpTip = getLabel("tooltip.column");
            columnRadios.push(columnRadio);
        }

        var orderPanel = sortDialog.add("panel", undefined, getLabel("panel.sortOrder"));
        orderPanel.orientation = "column";
        orderPanel.alignChildren = "left";
        orderPanel.margins = [10, 25, 10, 10];

        var orderRadioGroup = orderPanel.add("group");
        orderRadioGroup.orientation = "row";
        var ascendingRadio = orderRadioGroup.add("radiobutton", undefined, getLabel("radio.ascending"));
        ascendingRadio.helpTip = getLabel("tooltip.ascending");
        var descendingRadio = orderRadioGroup.add("radiobutton", undefined, getLabel("radio.descending"));
        descendingRadio.helpTip = getLabel("tooltip.descending");
        var randomRadio = orderRadioGroup.add("radiobutton", undefined, getLabel("radio.random"));
        randomRadio.helpTip = getLabel("tooltip.random");
        ascendingRadio.value = true;

        var headerCheckbox = sortDialog.add("checkbox", undefined, getLabel("checkbox.header"));
        headerCheckbox.helpTip = getLabel("tooltip.header");
        headerCheckbox.value = hasHeaderCandidate || looksLikeHeaderRow(columnCount, lines);

        /* 見出し行があれば2行目以降で既定の列を探す / skip the header row when choosing the default column */
        columnRadios[findDefaultColumn(columnCount, lines, headerCheckbox.value ? 1 : 0)].value = true;

        /* ボタン行（右：キャンセル・OK） / Button row (right: Cancel and OK) */
        var buttonRow = addButtonRow(sortDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "OK" });

        prepareDialogWindow(sortDialog, SCRIPT_NAME);
        if (sortDialog.show() !== 1) return null;

        var selectedColumn = 0;
        for (var j = 0; j < columnRadios.length; j++) {
            if (columnRadios[j].value) {
                selectedColumn = j;
                break;
            }
        }

        return {
            column: selectedColumn,
            order: randomRadio.value ? "random" : (descendingRadio.value ? "desc" : "asc"),
            useHeader: headerCheckbox.value
        };
    }

    // =========================================
    // テキストフレームの書き換え / Rewriting text frames
    // =========================================

    /**
     * テキストフレームを位置（Y座標の降順、同じなら X座標の昇順）で並べた配列を返す
     * @param {TextFrame[]} frameList - テキストフレーム
     * @returns {TextFrame[]} 並べ替えた新しい配列（失敗時は元の配列）
     */
    function sortTextFramesByPosition(frameList) {
        try {
            /* position の読み取りに失敗したら元の順のまま / keep the original order if a position cannot be read */
            var sortedList = frameList.slice();
            sortedList.sort(function (a, b) {
                var aPosition = a.position;
                var bPosition = b.position;
                if (aPosition[1] > bPosition[1]) return -1;
                if (aPosition[1] < bPosition[1]) return 1;
                if (aPosition[0] < bPosition[0]) return -1;
                if (aPosition[0] > bPosition[0]) return 1;
                return 0;
            });
            return sortedList;
        } catch (e) {
            alert("ソート中にエラーが発生しました: " + e.message);
            return frameList;
        }
    }

    /**
     * 上下に並んだテキストフレームを、上から順に1つのテキストフレームへまとめる
     * @param {TextFrame[]} frames - まとめるテキストフレーム
     * @returns {void}
     */
    function mergeTextFramesVertically(frames) {
        if (frames.length < 2) return;

        /* いったん1行ずつのフレームに分ける / split into one frame per line first */
        var sortedFrames = sortTextFramesByPosition(frames);
        var lineFrames = [];
        for (var i = 0; i < sortedFrames.length; i++) {
            var frameLines = sortedFrames[i].contents.split(LINE_SEPARATOR);
            for (var j = 0; j < frameLines.length; j++) {
                if (frameLines[j] !== "") {
                    var lineFrame = sortedFrames[i].duplicate();
                    lineFrame.contents = frameLines[j];
                    lineFrame.top -= j * 20; /* 位置調整（ソート用） / offset so the lines sort in order */
                    lineFrames.push(lineFrame);
                }
            }
            sortedFrames[i].remove();
        }
        sortedFrames = sortTextFramesByPosition(lineFrames);

        var baseFrame = sortedFrames[0];
        for (var k = 1; k < sortedFrames.length; k++) {
            baseFrame.paragraphs.add('\n');
            var sourceParagraphs = sortedFrames[k].paragraphs;
            for (var p = 0; p < sourceParagraphs.length; p++) {
                sourceParagraphs[p].duplicate(baseFrame);
            }
            sortedFrames[k].remove();
        }
    }

    /**
     * 見出し行を残して本文だけを並べ替えた内容に置き換える
     * 複製で見出しと本文のフレームに分け、書き換えたあと1つにまとめ直す
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {string[]} sortedLines - 並べ替えた本文の行
     * @returns {void}
     */
    function replaceBodyKeepingHeader(textFrame, sortedLines) {
        var headerFrame = textFrame;
        var bodyFrame = textFrame.duplicate();

        /* 本文フレームを1行分下に移動（フォントサイズに基づく） / move the body down one line by font size */
        var textSize = bodyFrame.textRange.characterAttributes.size;
        if (!isNaN(textSize)) {
            bodyFrame.top -= textSize * 1.5;
        }

        /* 見出し：2行目以降を削除 / header: remove every line but the first */
        var headerParagraphs = headerFrame.textRange.paragraphs;
        for (var i = headerParagraphs.length - 1; i >= 1; i--) {
            headerParagraphs[i].remove();
        }

        /* 本文：1行目を削除 / body: remove the first line */
        var bodyParagraphs = bodyFrame.textRange.paragraphs;
        if (bodyParagraphs.length > 1) {
            bodyParagraphs[0].remove();
        }

        bodyFrame.contents = sortedLines.join(LINE_SEPARATOR);
        mergeTextFramesVertically([headerFrame, bodyFrame]);
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択したテキストフレームのタブ区切りテキストを、選んだ列で並べ替える
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert("ドキュメントを開いてください。");
            return;
        }

        var docSelection = app.activeDocument.selection;
        if (docSelection.length !== 1 || docSelection[0].typename !== "TextFrame") {
            alert("1つのテキストオブジェクトを選択してください。");
            return;
        }

        var textFrame = docSelection[0];

        /* 空行や空白のみの行を除く / drop empty and whitespace-only lines */
        var allLines = textFrame.contents.split(LINE_SEPARATOR);
        var lines = [];
        for (var i = 0; i < allLines.length; i++) {
            if (allLines[i].replace(/\s/g, "").length > 0) {
                lines.push(allLines[i]);
            }
        }

        /* 1行目だけタブが無く2行目がタブ区切りなら、1行目を見出しとみなす / a tab-less first line over tabbed lines is a header */
        var hasHeaderCandidate = (lines.length >= 2 && lines[0].indexOf(CELL_SEPARATOR) === -1 && lines[1].indexOf(CELL_SEPARATOR) !== -1);
        var columnCount = lines[hasHeaderCandidate ? 1 : 0].split(CELL_SEPARATOR).length;

        var sortOptions = showSortOptionsDialog(columnCount, lines, hasHeaderCandidate);
        if (sortOptions === null) return;

        var dataLines = sortOptions.useHeader ? lines.slice(1) : lines;
        var sortedLines = generateSortedLines(dataLines, sortOptions.column, sortOptions.order);

        if (sortOptions.useHeader) {
            replaceBodyKeepingHeader(textFrame, sortedLines);
        } else {
            textFrame.contents = sortedLines.join(LINE_SEPARATOR);
        }
        app.redraw();
    }

    main();

})();
