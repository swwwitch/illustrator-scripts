#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アセット書き出しパネルに登録されたアセットの名前を、一括で変更します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RenameAssets.md

### Overview

Renames the assets registered in the Asset Export panel in bulk.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RenameAssets.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "RenameAssets";                 /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RenameAssets.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RenameAssets.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_MARGINS = 16;                  /* ダイアログの余白 / Dialog margins */
    var DIALOG_SPACING = 12;                  /* ダイアログの要素間隔 / Dialog spacing */
    var DIALOG_OFFSET_X = 300;                /* 表示位置を右へずらす量 / Horizontal shift of the dialog */
    var DIALOG_OFFSET_Y = 0;                  /* 表示位置を下へずらす量 / Vertical shift of the dialog */
    var DIALOG_OPACITY = 0.98;                /* ダイアログの不透明度 / Dialog opacity */
    var PANEL_MARGINS = [15, 20, 15, 10];     /* パネル余白 [左,上,右,下] / Panel margins [L,T,R,B] */
    var PANEL_SPACING = 16;                   /* パネル内の要素間隔 / Panel spacing */
    var FIELD_LABEL_WIDTH = 120;              /* 項目名の幅（揃える） / Width of the field labels */
    var FIELD_CHARS = 30;                     /* 入力欄の幅（文字数） / Width of the text fields */
    var STATUS_CHARS = 30;                    /* ステータス表示の幅（文字数） / Width of the status text */
    var BUTTON_ROW_MARGINS = [10, 10, 10, 0]; /* ボタン行の余白 / Button row margins */

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

    /* 日英ラベル定義。status.preview は件数を受け取る関数 / Japanese-English labels; status.preview takes counts */
    var LABELS = {
        dialog: {
            title: { ja: "グラフィックスタイルなどのアセット名のリネーム", en: "Rename Asset Names (Graphic Styles, etc.)" }
        },
        panel: {
            target: { ja: "対象", en: "Target" }
        },
        radio: {
            style: { ja: "グラフィックスタイル", en: "Graphic Styles" },
            brush: { ja: "ブラシ", en: "Brushes" },
            swatch: { ja: "スウォッチ", en: "Swatches" },
            symbol: { ja: "シンボル", en: "Symbols" }
        },
        fieldLabel: {
            find: { ja: "検索文字列", en: "Find" },
            replace: { ja: "置換文字列", en: "Replace" }
        },
        checkbox: {
            regex: { ja: "正規表現", en: "Regular Expression" },
            ignoreCase: { ja: "大文字／小文字を無視", en: "Ignore Case" }
        },
        status: {
            none: { ja: "—", en: "—" },
            noFind: { ja: "検索文字列なし", en: "Find string is empty" },
            preview: {
                ja: function (changed, total) {
                    return "プレビュー: 変更 " + changed + "／全 " + total;
                },
                en: function (changed, total) {
                    return "Preview: " + changed + " changed / " + total + " total";
                }
            }
        },
        tooltip: {
            find: { ja: "アセット名の中から探す文字列です。", en: "Text to look for in the asset names." },
            replace: {
                ja: "置き換える文字列です。空欄にすると検索文字列を削除します。",
                en: "Replacement text. Leave blank to delete the found text."
            },
            regex: {
                ja: "検索文字列を正規表現として扱います（$1 などの後方参照も使えます）。",
                en: "Treats the search text as a regular expression, including back-references such as $1."
            },
            ignoreCase: { ja: "大文字と小文字を区別せずに探します。", en: "Matches without regard to letter case." },
            target: { ja: "リネームするアセットの種類です。", en: "Which kind of asset to rename." },
            preview: {
                ja: "置換後の名前をアセットに仮に反映し、変更される件数を表示します。",
                en: "Applies the new names to the assets temporarily and shows how many change."
            }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            preview: { ja: "プレビュー", en: "Preview" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noAssets: {
                ja: "このドキュメントには対象となる項目（スタイル／ブラシ／スウォッチ／シンボル）がありません。",
                en: "No styles/brushes/swatches/symbols in this document."
            },
            noFind: { ja: "「検索」文字列が空です。", en: "Find string is empty." },
            noTarget: { ja: "選択した対象に名前付き項目がありません。", en: "No items found for selected target." }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語の値を取り出す
     * @param {string} labelPath - "dialog.title" のようなドット区切りのキー
     * @returns {string|Function} 表示言語の文言（status.preview は文言を作る関数）。見つからない場合は labelPath
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
            if (!labelNode) {
                return labelPath;
            }
        }
        return (uiLang === "ja") ? labelNode.ja : labelNode.en;
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角＋空白）
     * @param {string} labelPath - ラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === 'ja' ? '：' : ': ');
    }

    // =========================================
    // 名前の置換 / Name replacement
    // =========================================

    /**
     * 対象の種類に対応するアセットのコレクションを返す
     * @param {string} targetKey - "style" / "brush" / "swatch" / "symbol"
     * @returns {Object|null} コレクション（ドキュメントが無ければ null）
     */
    function getCollectionByKey(targetKey) {
        var doc = app.activeDocument;
        if (!doc) return null;
        if (targetKey === 'brush') return doc.brushes;
        if (targetKey === 'swatch') return doc.swatches;
        if (targetKey === 'symbol') return doc.symbols;
        return doc.graphicStyles;
    }

    /**
     * 置換関数を作る
     * @param {string} findStr - 検索文字列
     * @param {string} replStr - 置換文字列
     * @param {boolean} useRegex - 正規表現として扱うか
     * @param {boolean} ignoreCase - 大文字と小文字を区別しないか
     * @returns {Function} 元の名前を受け取り、置換後の名前（一致しなければ null）を返す関数
     */
    function buildReplacer(findStr, replStr, useRegex, ignoreCase) {
        if (useRegex || ignoreCase) {
            /* ［大文字／小文字を無視］だけのときも検索文字列を正規表現として組み立てる
               With only Ignore Case, the find text is still compiled as a pattern */
            var findPattern = null;
            try {
                findPattern = new RegExp(findStr, (useRegex && !ignoreCase) ? 'g' : 'gi');
            } catch (e) {
                /* 正規表現として不正な検索文字列 / invalid pattern */
                findPattern = null;
            }
            if (!findPattern) {
                return function () {
                    return null;
                };
            }
            return function (oldName) {
                return findPattern.test(oldName) ? oldName.replace(findPattern, replStr) : null;
            };
        }
        return function (oldName) {
            return oldName.indexOf(findStr) > -1 ? oldName.split(findStr).join(replStr) : null;
        };
    }

    /**
     * 角かっこで囲まれた既定の名前（[なし] など）かを返す
     * @param {string} assetName - アセット名
     * @returns {boolean} 既定の名前なら true
     */
    function isBracketedName(assetName) {
        /* 正規表現はリテラルで書く（ExtendScript の版によってはスコープの問題が出るため）
           Keep the literal inline to avoid scope issues in some ExtendScript versions */
        return /^\[.*\]$/.test(assetName);
    }

    /**
     * 名前の重複を避けるため、" (2)"、" (3)" … を付けて使われていない名前にする
     * @param {string} baseName - 希望する名前
     * @param {Object} usedNames - 使用済みの名前（名前 → true）
     * @returns {string} 使われていない名前（9999 を超えたらそこで打ち切る）
     */
    function ensureUniqueName(baseName, usedNames) {
        if (!usedNames[baseName]) return baseName;
        var i = 2;
        var candidate = baseName + " (" + i + ")";
        while (usedNames[candidate]) {
            i++;
            candidate = baseName + " (" + i + ")";
            /* 念のための上限 / safeguard */
            if (i > 9999) break;
        }
        return candidate;
    }

    /**
     * コレクションの名前を文字列として読み取る（名前が無いものは空文字）
     * @param {Object} collection - アセットのコレクション
     * @returns {string[]} 名前の配列
     */
    function readAssetNames(collection) {
        var assetNames = [];
        var totalCount = collection.length;
        for (var i = 0; i < totalCount; i++) {
            var rawName = collection[i].name;
            assetNames.push((rawName === undefined || rawName === null) ? "" : String(rawName));
        }
        return assetNames;
    }

    /**
     * 名前の配列から、使用済みの名前の表を作る
     * @param {string[]} assetNames - 名前の配列
     * @returns {Object} 名前 → true の表
     */
    function buildUsedNameMap(assetNames) {
        var usedNames = {};
        for (var i = 0; i < assetNames.length; i++) usedNames[assetNames[i]] = true;
        return usedNames;
    }

    /**
     * 変更後の名前を計算する（アセットはまだ変更しない）
     * @param {Object} collection - アセットのコレクション
     * @param {string} findStr - 検索文字列
     * @param {string} replStr - 置換文字列
     * @param {boolean} useRegex - 正規表現として扱うか
     * @param {boolean} ignoreCase - 大文字と小文字を区別しないか
     * @returns {{names: string[], changed: number}} 変更後の名前と、変更される件数
     */
    function computeProposedNames(collection, findStr, replStr, useRegex, ignoreCase) {
        var originalNames = readAssetNames(collection);
        var usedNames = buildUsedNameMap(originalNames);
        var proposedNames = originalNames.slice(0);
        var changedCount = 0;

        var replacer = buildReplacer(findStr, replStr, useRegex, ignoreCase);
        for (var j = 0; j < originalNames.length; j++) {
            var oldStr = originalNames[j];
            if (isBracketedName(oldStr)) continue;

            var replacedName = replacer(oldStr);
            if (replacedName !== null && replacedName !== oldStr) {
                /* 元の名前と、先に決めた名前の両方と重ならないようにする / avoid both original and earlier names */
                var uniqueName = ensureUniqueName(replacedName, usedNames);
                proposedNames[j] = uniqueName;
                usedNames[uniqueName] = true;
                changedCount++;
            }
        }
        return { names: proposedNames, changed: changedCount };
    }

    /**
     * 計算した名前をアセットに反映する
     * @param {Object} collection - アセットのコレクション
     * @param {string[]} newNames - 反映する名前
     * @returns {number} 変更した件数
     */
    function applyNames(collection, newNames) {
        var appliedCount = 0;
        for (var i = 0; i < collection.length && i < newNames.length; i++) {
            if (collection[i].name !== newNames[i]) {
                /* 名前を変更できないアセットがある / some assets reject a new name */
                try {
                    collection[i].name = newNames[i];
                    appliedCount++;
                } catch (e) { }
            }
        }
        return appliedCount;
    }

    /**
     * コレクション内のアセット名を置換する（確定実行）
     * @param {Object} collection - アセットのコレクション
     * @param {string} findStr - 検索文字列
     * @param {string} replStr - 置換文字列
     * @param {boolean} useRegex - 正規表現として扱うか
     * @param {boolean} ignoreCase - 大文字と小文字を区別しないか
     * @returns {number} 変更した件数
     */
    function renameItemsByCollection(collection, findStr, replStr, useRegex, ignoreCase) {
        /* 既存の名前を控えて重複を避ける / snapshot existing names for collision checks */
        var usedNames = buildUsedNameMap(readAssetNames(collection));
        var replacer = buildReplacer(findStr, replStr, useRegex, ignoreCase);
        var totalCount = collection.length;
        var changedCount = 0;

        for (var j = 0; j < totalCount; j++) {
            var asset = collection[j];
            var oldName = asset.name;
            if (oldName === undefined || oldName === null) continue;

            var oldStr = String(oldName);
            if (isBracketedName(oldStr)) continue;

            var replacedName = replacer(oldStr);
            if (replacedName === null || replacedName === oldStr) continue;

            var uniqueName = ensureUniqueName(replacedName, usedNames);
            if (uniqueName === oldStr) continue;

            /* 名前を変更できないアセットは飛ばす / skip assets that reject the new name */
            try {
                asset.name = uniqueName;
                /* 元の名前は表に残るが、新しい重複を防げれば十分 / old names stay in the map; that is fine */
                usedNames[uniqueName] = true;
                changedCount++;
            } catch (e) { }
        }
        return changedCount;
    }

    // =========================================
    // ダイアログの部品 / Dialog helpers
    // =========================================

    /**
     * ダイアログを開いたときに表示位置をずらす
     * @param {Window} targetDialog - 対象ダイアログ
     * @param {number} offsetX - 右へずらす量
     * @param {number} offsetY - 下へずらす量
     * @returns {void}
     */
    function shiftDialogPosition(targetDialog, offsetX, offsetY) {
        targetDialog.onShow = function () {
            var currentX = targetDialog.location[0];
            var currentY = targetDialog.location[1];
            targetDialog.location = [currentX + offsetX, currentY + offsetY];
        };
    }

    /**
     * ダイアログの不透明度を設定する
     * @param {Window} targetDialog - 対象ダイアログ
     * @param {number} opacityValue - 不透明度（0〜1）
     * @returns {void}
     */
    function setDialogOpacity(targetDialog, opacityValue) {
        /* 環境によっては opacity を設定できない / opacity is not supported everywhere */
        try {
            targetDialog.opacity = opacityValue;
        } catch (e) { }
    }

    /**
     * タイトル付きのパネルを追加する
     * @param {Object} parentGroup - 追加先のコンテナ
     * @param {string} title - パネルのタイトル
     * @param {string} [orientation] - 並びの向き（省略時は "column"）
     * @param {number[]} [margins] - 余白 [左,上,右,下]
     * @param {number} [spacing] - 要素間隔
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parentGroup, title, orientation, margins, spacing) {
        var titledPanel = parentGroup.add("panel", undefined, title);
        titledPanel.orientation = orientation || "column";
        titledPanel.alignChildren = ["fill", "top"];
        if (margins) titledPanel.margins = margins;
        if (spacing != null) titledPanel.spacing = spacing;
        return titledPanel;
    }

    /**
     * 横一列のグループを追加する
     * @param {Object} parentGroup - 追加先のコンテナ
     * @param {string[]} [alignChildren] - 子要素の揃え（["left", "center"] など）
     * @param {string|string[]} [alignment] - グループ自身の揃え（["fill", "bottom"] など）
     * @param {number} [spacing] - 要素間隔
     * @returns {Group} 追加したグループ
     */
    function addRow(parentGroup, alignChildren, alignment, spacing) {
        var rowGroup = parentGroup.add("group");
        rowGroup.orientation = "row";
        if (alignChildren) rowGroup.alignChildren = alignChildren;
        if (alignment) rowGroup.alignment = alignment;
        if (spacing != null) rowGroup.spacing = spacing;
        return rowGroup;
    }

    /**
     * 右揃えの項目名と入力欄の行を追加する
     * @param {Object} parentGroup - 追加先のコンテナ
     * @param {string} fieldLabelText - 項目名
     * @param {number} labelWidth - 項目名の幅
     * @param {number} [editChars] - 入力欄の幅（文字数。省略時は 30）
     * @returns {EditText} 追加した入力欄
     */
    function addLabeledEdit(parentGroup, fieldLabelText, labelWidth, editChars) {
        var rowGroup = addRow(parentGroup, ["fill", "center"]);
        var fieldLabel = rowGroup.add("statictext", undefined, fieldLabelText);
        fieldLabel.preferredSize.width = labelWidth;
        fieldLabel.justify = "right";
        var inputField = rowGroup.add("edittext", undefined, "");
        inputField.characters = editChars || 30;
        return inputField;
    }

    /**
     * チェックボックスを追加する
     * @param {Object} parentGroup - 追加先のコンテナ
     * @param {string} checkboxLabel - 表示名
     * @param {boolean} defaultValue - 初期値
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addCheckbox(parentGroup, checkboxLabel, defaultValue) {
        var checkboxControl = parentGroup.add("checkbox", undefined, checkboxLabel);
        checkboxControl.value = !!defaultValue;
        return checkboxControl;
    }

    /**
     * ラジオボタンのグループを追加する
     * @param {Object} parentGroup - 追加先のコンテナ
     * @param {string[]} radioLabels - 各ラジオボタンの表示名
     * @param {string} [orientation] - 並びの向き（省略時は "row"）
     * @param {number[]} [margins] - 余白
     * @param {number} [spacing] - 要素間隔
     * @returns {{group: Group, buttons: RadioButton[], select: Function, getSelectedIndex: Function, onChange: Function}} グループと操作用の関数
     */
    function addRadioGroup(parentGroup, radioLabels, orientation, margins, spacing) {
        var radioGroup = parentGroup.add("group");
        radioGroup.orientation = orientation || "row";
        if (margins) radioGroup.margins = margins;
        if (spacing != null) radioGroup.spacing = spacing;

        var radioButtons = [];
        for (var i = 0; i < radioLabels.length; i++) {
            radioButtons[i] = radioGroup.add("radiobutton", undefined, radioLabels[i]);
        }

        var changeHandlers = [];

        /**
         * 選択中のラジオボタンの番号を返す
         * @returns {number} 番号（選択なしは -1）
         */
        function getSelectedIndex() {
            for (var i = 0; i < radioButtons.length; i++) {
                if (radioButtons[i].value) return i;
            }
            return -1;
        }

        /**
         * 登録された変更時の処理を呼ぶ
         * @returns {void}
         */
        function fireChange() {
            var selectedIndex = getSelectedIndex();
            for (var h = 0; h < changeHandlers.length; h++) changeHandlers[h](selectedIndex);
        }

        for (var j = 0; j < radioButtons.length; j++) {
            radioButtons[j].onClick = fireChange;
        }

        return {
            group: radioGroup,
            buttons: radioButtons,
            select: function (selectedIndex, notify) {
                for (var i = 0; i < radioButtons.length; i++) radioButtons[i].value = (i === selectedIndex);
                if (notify === true) fireChange();
            },
            getSelectedIndex: getSelectedIndex,
            onChange: function (handler) {
                if (typeof handler === 'function') changeHandlers[changeHandlers.length] = handler;
            }
        };
    }

    /**
     * 中央寄せのテキスト（ステータス表示など）を追加する
     * @param {Object} parentGroup - 追加先のコンテナ
     * @param {number} [chars] - 幅（文字数。省略時は 30）
     * @returns {StaticText} 追加したテキスト
     */
    function addCenteredStatic(parentGroup, chars) {
        var rowGroup = addRow(parentGroup, ["center", "center"], "center");
        var centeredText = rowGroup.add("statictext", undefined, "");
        centeredText.characters = chars || 30;
        centeredText.alignment = "center";
        centeredText.justify = "center";
        return centeredText;
    }

    /**
     * ボタン行（左・スペーサー・右）を追加する
     * @param {Object} parentGroup - 追加先のコンテナ
     * @returns {{rowGroup: Group, leftGroup: Group, rightGroup: Group}} 行と左右のグループ
     */
    function addButtonRow(parentGroup) {
        var btnRowGroup = addRow(parentGroup, null, ["fill", "bottom"]);
        btnRowGroup.margins = BUTTON_ROW_MARGINS;

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];

        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];

        return { rowGroup: btnRowGroup, leftGroup: btnLeftGroup, rightGroup: btnRightGroup };
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ダイアログを表示し、OK なら置換の設定を返す（プレビューは名前を仮に変更し、キャンセルで戻す）
     * @returns {{find: string, repl: string, regex: boolean, ignoreCase: boolean, target: string}|null} 設定。キャンセル時は null
     */
    function showDialog() {
        var renameDialog = new Window("dialog", getLabel('dialog.title') + " " + SCRIPT_VERSION);
        setDialogOpacity(renameDialog, DIALOG_OPACITY);
        shiftDialogPosition(renameDialog, DIALOG_OFFSET_X, DIALOG_OFFSET_Y);
        renameDialog.orientation = "column";
        renameDialog.alignChildren = ["fill", "top"];
        renameDialog.margins = DIALOG_MARGINS;
        renameDialog.spacing = DIALOG_SPACING;

        /* 対象（上段）/ Target radio group */
        var targetPanel = addPanel(renameDialog, getLabel('panel.target'), 'row', PANEL_MARGINS, PANEL_SPACING);
        var targetRadios = addRadioGroup(targetPanel,
            [getLabel('radio.style'), getLabel('radio.brush'), getLabel('radio.swatch'), getLabel('radio.symbol')], 'row');
        var brushRadio = targetRadios.buttons[1];
        var swatchRadio = targetRadios.buttons[2];
        var symbolRadio = targetRadios.buttons[3];
        for (var r = 0; r < targetRadios.buttons.length; r++) {
            targetRadios.buttons[r].helpTip = getLabel('tooltip.target');
        }
        targetRadios.select(0);

        var findInput = addLabeledEdit(renameDialog, labelText('fieldLabel.find'), FIELD_LABEL_WIDTH, FIELD_CHARS);
        findInput.helpTip = getLabel('tooltip.find');
        var replaceInput = addLabeledEdit(renameDialog, labelText('fieldLabel.replace'), FIELD_LABEL_WIDTH, FIELD_CHARS);
        replaceInput.helpTip = getLabel('tooltip.replace');

        /* 正規表現と大文字／小文字（横並び・中央）/ Regex and ignore-case options, centered */
        var optionsRow = addRow(renameDialog, ["center", "center"], "center");
        var regexCheckbox = addCheckbox(optionsRow, getLabel('checkbox.regex'), false);
        regexCheckbox.helpTip = getLabel('tooltip.regex');
        var ignoreCaseCheckbox = addCheckbox(optionsRow, getLabel('checkbox.ignoreCase'), false);
        ignoreCaseCheckbox.helpTip = getLabel('tooltip.ignoreCase');

        /* プレビューの結果（ボタンで実行）/ Preview status, updated by the button */
        var statusText = addCenteredStatic(renameDialog, STATUS_CHARS);

        /* プレビュー前の名前と、その対象 / Names before the preview and their target */
        var snapshotNames = null;
        var snapshotTargetKey = null;
        var accepted = false;

        /**
         * 選択中の対象の種類を返す
         * @returns {string} "style" / "brush" / "swatch" / "symbol"
         */
        function getTargetKey() {
            return brushRadio.value ? 'brush' : (swatchRadio.value ? 'swatch' : (symbolRadio.value ? 'symbol' : 'style'));
        }

        /**
         * 控えた名前に戻してから処理を実行する（プレビューが重ならないように）。控えが無ければ先に取る
         * @param {string} targetKey - 対象の種類
         * @param {Function} workFn - 戻したコレクションを受け取る処理
         * @returns {Object|null} 対象のコレクション
         */
        function withSnapshot(targetKey, workFn) {
            var targetCollection = getCollectionByKey(targetKey);
            if (!targetCollection) return null;
            if (!snapshotNames || snapshotTargetKey !== targetKey) {
                snapshotNames = [];
                for (var i = 0; i < targetCollection.length; i++) snapshotNames[i] = targetCollection[i].name;
                snapshotTargetKey = targetKey;
            }
            restoreSnapshotNames(targetCollection);
            /* 名前の読み書きで例外が出てもダイアログは閉じない / keep the dialog alive on DOM errors */
            try {
                workFn && workFn(targetCollection);
            } catch (e) { }
            return targetCollection;
        }

        /**
         * 控えた名前をコレクションに書き戻す
         * @param {Object} targetCollection - 書き戻すコレクション
         * @returns {void}
         */
        function restoreSnapshotNames(targetCollection) {
            for (var i = 0; i < targetCollection.length && i < snapshotNames.length; i++) {
                /* 名前を変更できないアセットがある / some assets reject a new name */
                try {
                    targetCollection[i].name = snapshotNames[i];
                } catch (e) { }
            }
        }

        /**
         * キャンセル時にプレビュー前の名前へ戻す
         * @returns {void}
         */
        function restoreSnapshot() {
            if (!snapshotNames) return;
            var targetCollection = getCollectionByKey(snapshotTargetKey);
            if (!targetCollection) return;
            restoreSnapshotNames(targetCollection);
        }

        /**
         * 置換後の名前を仮に反映し、件数を表示する
         * @returns {void}
         */
        function updatePreview() {
            var findText = String(findInput.text || '');
            var replaceText = String(replaceInput.text || '');
            var useRegex = regexCheckbox.value === true;
            var ignoreCase = ignoreCaseCheckbox.value === true;
            if (!findText) {
                statusText.text = getLabel('status.noFind');
                return;
            }

            var previewPlan = null;
            var targetCollection = withSnapshot(getTargetKey(), function (restoredCollection) {
                previewPlan = computeProposedNames(restoredCollection, findText, replaceText, useRegex, ignoreCase);
                applyNames(restoredCollection, previewPlan.names);
            });
            if (!targetCollection) {
                statusText.text = getLabel('status.none');
                return;
            }
            statusText.text = getLabel('status.preview')(previewPlan.changed, targetCollection.length);
            app.redraw();
        }

        /* 対象を切り替えたら控えを捨て、次のプレビューで取り直す / a new target needs a new snapshot */
        targetRadios.onChange(function () {
            snapshotNames = null;
        });

        /* ボタン行（左=キャンセル／右=プレビュー・OK）/ Button row */
        var buttonRow = addButtonRow(renameDialog);
        var btnCancel = buttonRow.leftGroup.add("button", undefined, getLabel('button.cancel'), { name: "cancel" });
        var btnPreview = buttonRow.rightGroup.add("button", undefined, getLabel('button.preview'), { name: "preview" });
        btnPreview.helpTip = getLabel('tooltip.preview');
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel('button.ok'), { name: "ok" });

        btnOK.onClick = function () {
            /* プレビューで変えた名前はそのまま残す / keep the previewed names */
            accepted = true;
            snapshotNames = null;
            renameDialog.close();
        };
        btnCancel.onClick = function () {
            /* プレビューを反映していたら元の名前に戻す / revert the preview */
            try {
                restoreSnapshot();
            } catch (e) { }
            renameDialog.close();
        };
        btnPreview.onClick = function () {
            updatePreview();
        };

        findInput.active = true;

        renameDialog.show();

        return accepted ? {
            find: String(findInput.text || ""),
            repl: String(replaceInput.text || ""),
            regex: regexCheckbox.value,
            ignoreCase: ignoreCaseCheckbox.value,
            target: getTargetKey()
        } : null;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ドキュメントを確認し、ダイアログの設定でアセット名を置換する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel('alert.noDocument'));
            return;
        }
        var doc = app.activeDocument;
        /* 対象になるコレクションが1つでもあるか / at least one eligible collection */
        var hasAnyAssets = (doc.graphicStyles && doc.graphicStyles.length) ||
            (doc.brushes && doc.brushes.length) ||
            (doc.swatches && doc.swatches.length) ||
            (doc.symbols && doc.symbols.length);
        if (!hasAnyAssets) {
            alert(getLabel('alert.noAssets'));
            return;
        }

        var renameOptions = showDialog();
        if (!renameOptions) return;

        if (!renameOptions.find) {
            alert(getLabel('alert.noFind'));
            return;
        }

        var targetCollection = getCollectionByKey(renameOptions.target);
        if (!targetCollection || targetCollection.length === 0) {
            alert(getLabel('alert.noTarget'));
            return;
        }

        renameItemsByCollection(targetCollection, renameOptions.find, renameOptions.repl, renameOptions.regex, renameOptions.ignoreCase);
    }

    main();

})();
