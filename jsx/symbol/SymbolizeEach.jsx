#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトを1つずつ、またはまとめてシンボルとして登録し、元のオブジェクトをシンボルインスタンスに置き換えます。
名前はテキスト内容・レイヤー名・メモ・連番から自動で付けるか、Illustrator 標準の［新規シンボル］ダイアログで確認しながら登録できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SymbolizeEach.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nce9ec30232a0

### Overview

Registers the selected objects as symbols, one by one or as a single symbol, and replaces the originals with symbol instances.
Names are assigned automatically from the text contents, layer name, note, or a sequence number, or confirmed in Illustrator's native New Symbol dialog.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SymbolizeEach.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SymbolizeEach";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-10";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-15";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SymbolizeEach.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SymbolizeEach.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nce9ec30232a0"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 名前を連番で付けるときの既定の接頭辞 / Default prefix for sequence-numbered names */
    var DEFAULT_PREFIX = 'Symbol_';

    /* 連番の既定の桁数（1〜3）/ Default number of sequence digits (1–3) */
    var DEFAULT_SEQUENCE_DIGITS = 3;

    /* ［テキスト内容をシンボル名に使う］の初期状態 / Initial state of "Use text contents as symbol name" */
    var DEFAULT_USE_TEXT_AS_NAME = true;

    /* 基準点の初期選択（0〜8 の行優先、4 = 中央）/ Initial registration point (0–8, row-major; 4 = center) */
    var DEFAULT_REFERENCE_POINT_INDEX = 4;

    /* 選択範囲の扱いの初期選択（'asGroup' = まとめて1つ / 'eachItem' = オブジェクトごと）/ Initial selection handling */
    var DEFAULT_GROUP_MODE = 'eachItem';

    /* 登録方法の初期選択（'individual' = 標準ダイアログで確認 / 'batch' = 自動登録）/ Initial registration method */
    var DEFAULT_SYMBOLIZE_MODE = 'batch';

    /* リンク画像の初期選択（'ignore' = 無視 / 'embed' = 埋め込んで登録）/ Initial linked-image policy */
    var DEFAULT_LINKED_IMAGE_POLICY = 'embed';

    /* シンボル名の最大文字数 / Maximum length of a symbol name */
    var SYMBOL_NAME_MAX_LENGTH = 80;

    // =========================================
    // レイアウト / Layout
    // =========================================

    var WINDOW_MARGINS     = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING     = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS      = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING      = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING     = 12;                 /* 2カラムの間隔 / column spacing */
    var FIELD_LABEL_WIDTH  = 64;                 /* 項目名の幅 / field label width */
    var PREFIX_INPUT_WIDTH = 180;                /* 接頭辞の入力欄の幅 / prefix field width */

    var ANCHOR_WIDGET_SIZE = 66;                 /* 基準点ウィジェット全体の大きさ / anchor widget size */
    var ANCHOR_CELL_SIZE   = 9;                  /* 基準点ウィジェットの□1個の大きさ / anchor cell size */
    var ANCHOR_CELL_GAP    = 7.5;                /* 基準点ウィジェットの□どうしの間隔 / gap between anchor cells */

    /* 基準点ウィジェットの配色（選択セルの塗りは UI の明暗で上書き）/ Anchor widget colors (selected fill follows UI brightness) */
    var ANCHOR_LINE_COLOR    = [0.6, 0.6, 0.6, 1];
    var ANCHOR_SELECTED_FILL = [0.4, 0.4, 0.4, 1];
    var ANCHOR_DISABLED_LINE = [0.75, 0.75, 0.75, 1];
    var ANCHOR_DISABLED_FILL = [0.75, 0.75, 0.75, 1];

    /* 中央(4)を除く外周の□どうしをつなぐケイ線の組み合わせ / Pairs of outer cells (center excluded) joined by rules */
    var ANCHOR_CONNECTIONS = [[0, 1], [1, 2], [6, 7], [7, 8], [0, 3], [3, 6], [2, 5], [5, 8]];

    /**
     * ウィンドウの共通レイアウトを設定する
     * @param {Window} targetWindow - 対象のウィンドウ
     * @returns {void}
     */
    function setupWindow(targetWindow) {
        targetWindow.orientation = 'column';
        targetWindow.alignChildren = ['fill', 'top'];
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = WINDOW_SPACING;
    }

    /**
     * タイトル付きのパネルを追加し、共通レイアウトを設定する
     * @param {object} parent - 追加先のウィンドウまたはグループ
     * @param {object} titleSet - パネル名のラベル（ja/en）
     * @param {number} [spacing] - パネル内の要素間隔（省略時は PANEL_SPACING）
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parent, titleSet, spacing) {
        var newPanel = parent.add('panel', undefined, getLabel(titleSet));
        newPanel.orientation = 'column';
        newPanel.alignChildren = ['left', 'top'];
        newPanel.alignment = ['fill', 'fill'];
        newPanel.margins = PANEL_MARGINS;
        newPanel.spacing = (typeof spacing === 'number') ? spacing : PANEL_SPACING;
        return newPanel;
    }

    /**
     * 行グループの共通レイアウトを設定する（横位置と天地を対で指定し、子は伸ばさない）
     * @param {Group} rowGroup - 対象のグループ
     * @param {string} [alignment] - 行の横位置（省略時は 'left'）
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(rowGroup, alignment, spacing) {
        rowGroup.orientation = 'row';
        rowGroup.alignment = [alignment || 'left', 'center'];
        rowGroup.alignChildren = ['left', 'center'];
        rowGroup.spacing = (typeof spacing === 'number') ? spacing : PANEL_SPACING;
    }

    /**
     * 行の先頭に右揃えの項目名を追加する
     * @param {Group} rowGroup - 追加先の行グループ
     * @param {object} labelSet - 項目名のラベル（ja/en）
     * @returns {StaticText} 追加した項目名
     */
    function addFieldLabel(rowGroup, labelSet) {
        var fieldLabel = rowGroup.add('statictext', undefined, labelText(labelSet));
        fieldLabel.preferredSize.width = FIELD_LABEL_WIDTH;
        fieldLabel.justify = 'right';
        return fieldLabel;
    }

    // =========================================
    // 定数 / Constants
    // =========================================

    /* 選択範囲の扱い / Selection handling */
    var GROUP_MODE = { AS_GROUP: 'asGroup', EACH_ITEM: 'eachItem' };

    /* 登録方法 / Registration method */
    var SYMBOLIZE_MODE = { INDIVIDUAL: 'individual', BATCH: 'batch' };

    /* リンク画像の扱い / Linked-image policy */
    var LINKED_IMAGE_POLICY = { IGNORE: 'ignore', EMBED: 'embed' };

    /* 連番の桁数の選択肢（"0" / "00" / "000"）/ Sequence-digit choices */
    var SEQUENCE_DIGIT_OPTIONS = [1, 2, 3];

    /* 3×3 の基準点と SymbolRegistrationPoint の対応（行優先：上 → 中 → 下、列：左 → 中 → 右）/ Row-major 3×3 registration points */
    var REFERENCE_POINTS = [
        SymbolRegistrationPoint.SYMBOLTOPLEFTPOINT,
        SymbolRegistrationPoint.SYMBOLTOPMIDDLEPOINT,
        SymbolRegistrationPoint.SYMBOLTOPRIGHTPOINT,
        SymbolRegistrationPoint.SYMBOLMIDDLELEFTPOINT,
        SymbolRegistrationPoint.SYMBOLCENTERPOINT,
        SymbolRegistrationPoint.SYMBOLMIDDLERIGHTPOINT,
        SymbolRegistrationPoint.SYMBOLBOTTOMLEFTPOINT,
        SymbolRegistrationPoint.SYMBOLBOTTOMMIDDLEPOINT,
        SymbolRegistrationPoint.SYMBOLBOTTOMRIGHTPOINT
    ];

    /* 名前の候補にしない既定のレイヤー名 / Default layer names that are not used as symbol names */
    var DEFAULT_LAYER_NAME_PATTERN = /^(Layer|レイヤー)\s*\d+$/;

    /* 「新規シンボル…」のメニューコマンド / Menu command for "New Symbol..." */
    var NEW_SYMBOL_MENU_COMMAND = 'Adobe New Symbol Shortcut';

    /* 標準ダイアログ表示前に対象へ寄せるビュー（対象がビューの半分を占める倍率、上下限つき）/ View focus before the native dialog */
    var FOCUS_VIEW_RATIO = 0.5;
    var MIN_VIEW_ZOOM = 0.05;
    var MAX_VIEW_ZOOM = 64;

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 現在の UI 言語 / Current UI language */
    var uiLang = ($.locale.indexOf('ja') === 0) ? 'ja' : 'en';

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: 'シンボル化', en: 'Symbolize' }
        },
        panel: {
            groupMode: { ja: '選択範囲の扱い', en: 'Selection handling' },
            symbolizeMode: { ja: '登録方法', en: 'Registration method' },
            symbolName: { ja: 'シンボル名', en: 'Symbol name' },
            referencePoint: { ja: '基準点', en: 'Registration point' },
            linkedImage: { ja: 'リンク画像', en: 'Linked images' }
        },
        radio: {
            asGroup: { ja: 'まとめて1つのシンボルにする', en: 'Create one symbol from selection' },
            eachItem: { ja: 'オブジェクトごとにシンボル化', en: 'Create symbols per object' },
            individual: { ja: '標準ダイアログで確認', en: 'Confirm with native dialog' },
            batch: { ja: '自動登録', en: 'Register automatically' },
            ignoreLinked: { ja: '無視', en: 'Ignore' },
            embedLinked: { ja: '埋め込んで登録', en: 'Embed and register' }
        },
        checkbox: {
            useTextAsName: { ja: 'テキスト内容をシンボル名に使う', en: 'Use text contents as symbol name' }
        },
        fieldLabel: {
            prefix: { ja: '接頭辞', en: 'Prefix' },
            sequence: { ja: '連番', en: 'Sequence' }
        },
        tooltip: {
            asGroup: {
                ja: '現在の選択全体を1つのシンボルとして登録します。この場合、登録方法は標準ダイアログでの確認になります。',
                en: 'Registers the current selection as one symbol. In this mode, the native dialog is used for confirmation.'
            },
            eachItem: { ja: '選択中の各オブジェクトを個別のシンボルとして登録します。', en: 'Registers each selected object as its own symbol.' },
            individual: {
                ja: '対象ごとに Illustrator 標準の「新規シンボル」ダイアログを開きます。名前や基準点を毎回確認したい場合に使います。',
                en: 'Opens Illustrator\'s native New Symbol dialog for each target. Use this when you want to confirm the name and registration point each time.'
            },
            batch: {
                ja: '接頭辞・連番・テキスト流用・基準点の設定に従って、確認なしで自動登録します。',
                en: 'Registers symbols automatically without confirmation, using the prefix, sequence, text-reuse, and registration-point settings.'
            },
            prefix: {
                ja: 'テキスト内容・レイヤー名・メモから名前を取得できない場合に使う接頭辞です。',
                en: 'Prefix used when the symbol name cannot be taken from text contents, the layer name, or the item note.'
            },
            sequence: {
                ja: '接頭辞に続ける連番の桁数です。例：0、00、000。',
                en: 'Number of digits for the sequence appended to the prefix, such as 0, 00, or 000.'
            },
            useTextAsName: {
                ja: 'TextFrame、またはグループ内で最初に見つかった TextFrame の内容をシンボル名に使います。空の場合は次の候補に進みます。',
                en: 'Uses the TextFrame contents, or the first TextFrame found inside a group, as the symbol name. If empty, the next naming source is used.'
            },
            referencePoint: {
                ja: '自動登録時に使うシンボルの基準点です。標準ダイアログで確認する場合は Illustrator 側で指定します。',
                en: 'Registration point used for automatic registration. When using the native dialog, set it in Illustrator.'
            },
            linkedImage: {
                ja: 'リンク画像（PlacedItem）の扱いです。［無視］は選択に残して登録対象から外します。［埋め込んで登録］はシンボル化前に埋め込みます。',
                en: 'Policy for linked images (PlacedItem). Ignore keeps them selected but excludes them. Embed and register embeds them before symbolization.'
            }
        },
        button: {
            cancel: { ja: 'キャンセル', en: 'Cancel' },
            ok: { ja: 'OK', en: 'OK' }
        },
        alert: {
            noDocument: { ja: 'ドキュメントが開かれていません。', en: 'No document is open.' },
            noSelection: { ja: 'シンボル化したいオブジェクトを選択してください。', en: 'Select the objects you want to symbolize.' },
            completed: { ja: '完了しました。', en: 'Completed.' },
            createdCount: { ja: '新規作成したシンボル数', en: 'Created symbols' },
            existingCount: { ja: 'スルーした既存シンボル数', en: 'Skipped existing symbols' },
            ignoredLinkedCount: { ja: '無視したリンク画像数', en: 'Ignored linked images' },
            failedCount: { ja: '失敗', en: 'Failed' },
            failureDetails: { ja: '失敗の詳細', en: 'Failure details' }
        }
    };

    /**
     * ラベル（ja/en）を現在の UI 言語の文字列にする
     * @param {object} labelSet - ja/en を持つラベル
     * @returns {string} 現在の言語の文字列
     */
    function getLabel(labelSet) {
        return (labelSet && labelSet[uiLang]) || '';
    }

    /**
     * 項目名にコロンを付ける（日本語は全角、英語は半角）
     * @param {object} labelSet - ja/en を持つラベル
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === 'ja' ? '：' : ':');
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * @typedef {object} SymbolizeSettings
     * @property {string} groupMode - 選択範囲の扱い（GROUP_MODE）
     * @property {string} symbolizeMode - 登録方法（SYMBOLIZE_MODE）
     * @property {string} defaultPrefix - 連番で名前を付けるときの接頭辞
     * @property {number} sequenceDigits - 連番の桁数
     * @property {boolean} useTextAsName - テキスト内容をシンボル名に使うか
     * @property {SymbolRegistrationPoint} referencePoint - 自動登録時の基準点
     * @property {string} linkedImagePolicy - リンク画像の扱い（LINKED_IMAGE_POLICY）
     */

    /**
     * @typedef {object} ResultStats
     * @property {PageItem[]} finalSelection - 処理後に選択するアイテム
     * @property {number} createdCount - 新規作成したシンボル数
     * @property {number} existingSymbolCount - スルーした既存シンボルインスタンス数
     * @property {number} ignoredLinkedImageCount - 無視したリンク画像数
     * @property {string[]} failureMessages - 失敗の詳細（件数は length）
     */

    /**
     * 選択を確認し、ダイアログの設定に従ってシンボル化して結果を通知する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var doc = app.activeDocument;
        if (!doc.selection || doc.selection.length === 0) {
            alert(getLabel(LABELS.alert.noSelection));
            return;
        }

        var originalSelection = snapshotSelection(doc);

        var symbolizeSettings = showSettingsDialog();
        if (!symbolizeSettings) return;

        doc.selection = null;

        var resultStats = (symbolizeSettings.groupMode === GROUP_MODE.AS_GROUP)
            ? processSelectionAsGroup(doc, originalSelection, symbolizeSettings)
            : processEachItem(doc, originalSelection, symbolizeSettings);

        applySelection(doc, resultStats.finalSelection);
        showResultSummary(resultStats);
    }

    // =========================================
    // シンボル化の流れ / Symbolize flow
    // =========================================

    /**
     * 現在の選択を配列としてコピーする
     * @param {Document} doc - 対象ドキュメント
     * @returns {PageItem[]} 選択アイテムの配列
     */
    function snapshotSelection(doc) {
        var selectionItems = [];
        for (var i = 0; i < doc.selection.length; i++) {
            selectionItems.push(doc.selection[i]);
        }
        return selectionItems;
    }

    /**
     * 空の集計オブジェクトを作る
     * @returns {ResultStats} 集計オブジェクト
     */
    function createResultStats() {
        return {
            finalSelection: [],
            createdCount: 0,
            existingSymbolCount: 0,
            ignoredLinkedImageCount: 0,
            failureMessages: []
        };
    }

    /**
     * 既存のシンボルインスタンスと［無視］指定のリンク画像を、処理せず選択に残して数える
     * @param {PageItem} item - 判定するアイテム
     * @param {SymbolizeSettings} symbolizeSettings - ダイアログの設定
     * @param {ResultStats} resultStats - 集計先
     * @returns {boolean} 処理対象外なら true
     */
    function skipExcludedItem(item, symbolizeSettings, resultStats) {
        if (item.typename === 'SymbolItem') {
            resultStats.existingSymbolCount++;
        } else if (item.typename === 'PlacedItem' && symbolizeSettings.linkedImagePolicy === LINKED_IMAGE_POLICY.IGNORE) {
            resultStats.ignoredLinkedImageCount++;
        } else {
            return false;
        }
        resultStats.finalSelection.push(item);
        return true;
    }

    /**
     * 失敗したアイテムの詳細を集計に加える
     * @param {ResultStats} resultStats - 集計先
     * @param {PageItem} item - 失敗したアイテム
     * @param {Error} err - 発生したエラー
     * @returns {void}
     */
    function recordFailure(resultStats, item, err) {
        resultStats.failureMessages.push(formatFailureMessage(item, err));
    }

    /**
     * オブジェクトごとにシンボル化する（自動登録、または対象ごとに標準ダイアログで確認）
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} items - 処理するアイテム
     * @param {SymbolizeSettings} symbolizeSettings - ダイアログの設定
     * @returns {ResultStats} 集計結果
     */
    function processEachItem(doc, items, symbolizeSettings) {
        var resultStats = createResultStats();
        var useNativeDialog = (symbolizeSettings.symbolizeMode === SYMBOLIZE_MODE.INDIVIDUAL);

        /* 標準ダイアログでは対象へビューを寄せるので、終了時に元のビューへ戻す / Restore the view after focusing on each target */
        var savedView = useNativeDialog ? captureViewState(doc) : null;

        try {
            for (var j = 0; j < items.length; j++) {
                var originalItem = items[j];
                if (skipExcludedItem(originalItem, symbolizeSettings, resultStats)) continue;

                try {
                    if (useNativeDialog) {
                        runNewSymbolDialog(doc, [ensureEmbedded(doc, originalItem)], resultStats);
                    } else {
                        resultStats.finalSelection.push(symbolizeOneItem(doc, originalItem, j, symbolizeSettings));
                        resultStats.createdCount++;
                    }
                } catch (err) {
                    recordFailure(resultStats, originalItem, err);
                }
            }
        } finally {
            restoreViewState(doc, savedView);
        }

        return resultStats;
    }

    /**
     * 選択全体をまとめて1つのシンボルとして、標準ダイアログで登録する
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} items - 処理するアイテム
     * @param {SymbolizeSettings} symbolizeSettings - ダイアログの設定
     * @returns {ResultStats} 集計結果
     */
    function processSelectionAsGroup(doc, items, symbolizeSettings) {
        var resultStats = createResultStats();

        /* リンク画像の扱いを反映しながら、まとめて登録するアイテムを集める / Collect the targets, embedding linked images as needed */
        var targetItems = [];
        for (var j = 0; j < items.length; j++) {
            var originalItem = items[j];
            if (skipExcludedItem(originalItem, symbolizeSettings, resultStats)) continue;

            try {
                targetItems.push(ensureEmbedded(doc, originalItem));
            } catch (err) {
                recordFailure(resultStats, originalItem, err);
            }
        }
        if (targetItems.length === 0) return resultStats;

        var savedView = captureViewState(doc);
        try {
            runNewSymbolDialog(doc, targetItems, resultStats);
        } catch (err) {
            recordFailure(resultStats, targetItems[0], err);
        } finally {
            restoreViewState(doc, savedView);
        }

        return resultStats;
    }

    /**
     * 対象だけを選択してビューを寄せ、Illustrator 標準の「新規シンボル…」ダイアログを開く
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} targetItems - シンボルにするアイテム
     * @param {ResultStats} resultStats - 集計先
     * @returns {void}
     */
    function runNewSymbolDialog(doc, targetItems, resultStats) {
        applySelection(doc, targetItems);
        if (!doc.selection || doc.selection.length === 0) {
            throw new Error('Target items could not be selected.');
        }

        /* モーダル表示前に redraw して選択のハイライトを描く / Redraw so the highlight shows before the modal dialog */
        focusViewOnItems(doc, targetItems);
        app.redraw();

        var symbolCountBefore = doc.symbols.length;
        app.executeMenuCommand(NEW_SYMBOL_MENU_COMMAND);
        if (doc.symbols.length > symbolCountBefore) {
            resultStats.createdCount++;
        }

        /* 実行後の選択をそのまま結果に取り込む（キャンセル時は元のアイテムが残る）/ Keep whatever is selected afterward (cancel leaves the originals) */
        var selectionAfter = doc.selection;
        if (!selectionAfter) return;
        for (var k = 0; k < selectionAfter.length; k++) {
            resultStats.finalSelection.push(selectionAfter[k]);
        }
    }

    /**
     * 失敗したアイテムの種類・名前・エラー内容を1行にまとめる
     * @param {PageItem} item - 失敗したアイテム
     * @param {Error} err - 発生したエラー
     * @returns {string} 失敗の詳細
     */
    function formatFailureMessage(item, err) {
        var itemType = 'Unknown';
        var itemName = '';

        /* 削除済みなどで参照が無効なときは種類・名前を取れないまま続ける / Carry on without type/name when the reference is invalid */
        try {
            itemType = item.typename;
            itemName = item.name || '';
        } catch (e) { }

        var errorMessage = (err && err.message) ? err.message : String(err);
        return itemType + (itemName !== '' ? ' "' + itemName + '"' : '') + ': ' + errorMessage;
    }

    /**
     * 指定したアイテムだけを選択する（既存の選択は解除）
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} items - 選択するアイテム
     * @returns {void}
     */
    function applySelection(doc, items) {
        doc.selection = null;
        for (var k = 0; k < items.length; k++) {
            /* 参照が無効になったアイテムは飛ばす / Skip items whose reference is no longer valid */
            try {
                items[k].selected = true;
            } catch (e) { }
        }
    }

    /**
     * 処理結果の件数と失敗の詳細を通知する
     * @param {ResultStats} resultStats - 集計結果
     * @returns {void}
     */
    function showResultSummary(resultStats) {
        var countLines = [
            formatCountLine(LABELS.alert.createdCount, resultStats.createdCount),
            formatCountLine(LABELS.alert.existingCount, resultStats.existingSymbolCount),
            formatCountLine(LABELS.alert.ignoredLinkedCount, resultStats.ignoredLinkedImageCount),
            formatCountLine(LABELS.alert.failedCount, resultStats.failureMessages.length)
        ];
        var message = getLabel(LABELS.alert.completed) + '\n\n' + countLines.join('\n');

        if (resultStats.failureMessages.length > 0) {
            message += '\n\n' + getLabel(LABELS.alert.failureDetails) + '\n' + resultStats.failureMessages.join('\n');
        }

        alert(message);
    }

    /**
     * 件数の行を作る（英語はコロンの後に空白を入れる）
     * @param {object} labelSet - 項目名のラベル（ja/en）
     * @param {number} count - 件数
     * @returns {string} 「項目名：件数」の行
     */
    function formatCountLine(labelSet, count) {
        return labelText(labelSet) + (uiLang === 'ja' ? '' : ' ') + count;
    }

    // =========================================
    // 1オブジェクトのシンボル化 / Symbolize one item
    // =========================================

    /**
     * 1オブジェクトを新規シンボルとして登録し、元オブジェクトをそのインスタンスで置き換える
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} originalItem - シンボルにするアイテム
     * @param {number} index - 選択内での位置（連番に使う）
     * @param {SymbolizeSettings} symbolizeSettings - ダイアログの設定
     * @returns {SymbolItem} 置き換えたシンボルインスタンス
     */
    function symbolizeOneItem(doc, originalItem, index, symbolizeSettings) {
        /* 埋め込みで境界が変わることがあるので、元の位置を先に控える / Capture bounds first because embedding may change them */
        var originalBounds = originalItem.geometricBounds;
        var workingItem = ensureEmbedded(doc, originalItem);

        var symbolName = resolveSymbolName(doc, workingItem, index, symbolizeSettings);
        var newSymbol = createSymbolFromItem(doc, workingItem, symbolName, symbolizeSettings.referencePoint);
        var newInstance = doc.symbolItems.add(newSymbol);

        /* 元オブジェクトの直前へ移して重なり順を近づける（失敗しても位置合わせは続ける）/ Keep the z-order close to the original */
        try {
            newInstance.move(workingItem, ElementPlacement.PLACEBEFORE);
        } catch (e) { }

        alignToBounds(newInstance, originalBounds);
        workingItem.remove();

        return newInstance;
    }

    /**
     * リンク画像なら埋め込み、置き換わったアイテムを返す（それ以外はそのまま返す）
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} item - 対象のアイテム
     * @returns {PageItem} 以後の処理に使うアイテム
     */
    function ensureEmbedded(doc, item) {
        if (item.typename !== 'PlacedItem') return item;

        /* 埋め込みの前後で親の子アイテムを比べ、増えたものを置き換え後のアイテムとする / Diff the parent's children around embed() */
        var parentContainer = item.parent;
        var itemsBeforeEmbed = snapshotChildPageItems(parentContainer);

        item.embed();

        var embeddedItem = findNewPageItem(parentContainer, itemsBeforeEmbed);
        if (embeddedItem) return embeddedItem;

        /* 見つからないときは、埋め込み後に選択される置き換えアイテムを使う / Fall back to the replacement Illustrator selects after embedding */
        if (doc.selection && doc.selection.length > 0) {
            return doc.selection[0];
        }

        throw new Error('Embedded replacement item could not be located.');
    }

    /**
     * 親コンテナ直下のアイテムを配列として控える
     * @param {object} parentContainer - Layer または GroupItem
     * @returns {PageItem[]} 直下のアイテム
     */
    function snapshotChildPageItems(parentContainer) {
        var childItems = [];
        for (var i = 0; i < parentContainer.pageItems.length; i++) {
            childItems.push(parentContainer.pageItems[i]);
        }
        return childItems;
    }

    /**
     * 控えた一覧に無いアイテムを親コンテナ直下から探す
     * @param {object} parentContainer - Layer または GroupItem
     * @param {PageItem[]} itemsBefore - 埋め込み前に控えたアイテム
     * @returns {PageItem|null} 新しく増えたアイテム（見つからないときは null）
     */
    function findNewPageItem(parentContainer, itemsBefore) {
        /* 走査に失敗したときは null を返し、呼び出し側のフォールバックに任せる / Return null so the caller can fall back */
        try {
            for (var i = 0; i < parentContainer.pageItems.length; i++) {
                if (!containsPageItem(itemsBefore, parentContainer.pageItems[i])) {
                    return parentContainer.pageItems[i];
                }
            }
        } catch (e) { }
        return null;
    }

    /**
     * 配列に同じアイテムの参照が含まれるかを調べる
     * @param {PageItem[]} items - 調べる配列
     * @param {PageItem} targetItem - 探すアイテム
     * @returns {boolean} 含まれていれば true
     */
    function containsPageItem(items, targetItem) {
        for (var i = 0; i < items.length; i++) {
            if (items[i] === targetItem) return true;
        }
        return false;
    }

    /**
     * アイテムの複製をシンボルとして登録し、複製は破棄する
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} sourceItem - 元のアイテム
     * @param {string} symbolName - シンボル名
     * @param {SymbolRegistrationPoint} referencePoint - 基準点
     * @returns {Symbol} 登録したシンボル
     */
    function createSymbolFromItem(doc, sourceItem, symbolName, referencePoint) {
        var duplicatedItem = sourceItem.duplicate();
        var newSymbol = doc.symbols.add(duplicatedItem, referencePoint);
        newSymbol.name = symbolName;

        /* 登録で複製が取り込まれて消えている場合もあるので、失敗は無視する / The duplicate may already be gone after registration */
        try {
            duplicatedItem.remove();
        } catch (e) { }

        return newSymbol;
    }

    /**
     * アイテムの左上が指定した境界の左上に合うよう移動する
     * @param {PageItem} item - 移動するアイテム
     * @param {number[]} targetBounds - 合わせる境界 [左, 上, 右, 下]
     * @returns {void}
     */
    function alignToBounds(item, targetBounds) {
        var currentBounds = item.geometricBounds;
        var dx = targetBounds[0] - currentBounds[0];
        var dy = targetBounds[1] - currentBounds[1];
        item.translate(dx, dy);
    }

    // =========================================
    // ビュー / View
    // =========================================

    /**
     * 現在のビューのズーム倍率と中心を控える
     * @param {Document} doc - 対象ドキュメント
     * @returns {object|null} { zoom, centerPoint }（取得できないときは null）
     */
    function captureViewState(doc) {
        try {
            var view = doc.views[0];
            return {
                zoom: view.zoom,
                centerPoint: [view.centerPoint[0], view.centerPoint[1]]
            };
        } catch (e) {
            return null;
        }
    }

    /**
     * 控えたビューの状態に戻す
     * @param {Document} doc - 対象ドキュメント
     * @param {object|null} viewState - captureViewState() の戻り値
     * @returns {void}
     */
    function restoreViewState(doc, viewState) {
        if (!viewState) return;
        try {
            var view = doc.views[0];
            view.zoom = viewState.zoom;
            view.centerPoint = viewState.centerPoint;
        } catch (e) { }
    }

    /**
     * アイテム全体がビューの中央でおよそ半分を占めるよう、ズームと中心を合わせる
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} items - 対象のアイテム
     * @returns {void}
     */
    function focusViewOnItems(doc, items) {
        /* ビューを寄せられなくても登録は続ける / Registration continues even if the view cannot be focused */
        try {
            var itemBounds = getCombinedBounds(items);
            var itemWidth = itemBounds[2] - itemBounds[0];
            var itemHeight = itemBounds[1] - itemBounds[3];
            if (itemWidth <= 0 || itemHeight <= 0) return;

            var view = doc.views[0];
            var viewBounds = view.bounds;
            var viewWidth = viewBounds[2] - viewBounds[0];
            var viewHeight = viewBounds[1] - viewBounds[3];

            var targetZoom = Math.min(
                FOCUS_VIEW_RATIO * viewWidth * view.zoom / itemWidth,
                FOCUS_VIEW_RATIO * viewHeight * view.zoom / itemHeight
            );
            view.zoom = Math.max(MIN_VIEW_ZOOM, Math.min(MAX_VIEW_ZOOM, targetZoom));
            view.centerPoint = [(itemBounds[0] + itemBounds[2]) / 2, (itemBounds[1] + itemBounds[3]) / 2];
        } catch (e) { }
    }

    /**
     * 複数アイテムを囲む境界を求める
     * @param {PageItem[]} items - 対象のアイテム（1つ以上）
     * @returns {number[]} 境界 [左, 上, 右, 下]
     */
    function getCombinedBounds(items) {
        var firstBounds = items[0].geometricBounds;
        var combinedBounds = [firstBounds[0], firstBounds[1], firstBounds[2], firstBounds[3]];
        for (var i = 1; i < items.length; i++) {
            var bounds = items[i].geometricBounds;
            combinedBounds[0] = Math.min(combinedBounds[0], bounds[0]);
            combinedBounds[1] = Math.max(combinedBounds[1], bounds[1]);
            combinedBounds[2] = Math.max(combinedBounds[2], bounds[2]);
            combinedBounds[3] = Math.min(combinedBounds[3], bounds[3]);
        }
        return combinedBounds;
    }

    // =========================================
    // シンボル名 / Symbol name
    // =========================================

    /**
     * シンボル名を決める。優先順位はテキスト内容（オン時）→ レイヤー名 → メモ → 接頭辞＋連番
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} item - 対象のアイテム
     * @param {number} index - 選択内での位置（連番に使う）
     * @param {SymbolizeSettings} symbolizeSettings - ダイアログの設定
     * @returns {string} ドキュメント内で重複しないシンボル名
     */
    function resolveSymbolName(doc, item, index, symbolizeSettings) {
        var baseName = '';

        if (symbolizeSettings.useTextAsName) {
            baseName = sanitizeSymbolName(findTextContents(item));
        }
        if (baseName === '') {
            baseName = sanitizeSymbolName(getLayerNameFromItem(item));
        }
        if (baseName === '') {
            baseName = sanitizeSymbolName(getNoteFromItem(item));
        }
        if (baseName === '') {
            baseName = symbolizeSettings.defaultPrefix + zeroPadding(index + 1, symbolizeSettings.sequenceDigits);
        }

        return getUniqueSymbolName(doc, baseName);
    }

    /**
     * テキストフレームの内容、またはグループ内で最初に見つかったテキストフレームの内容を返す
     * @param {PageItem} item - 対象のアイテム
     * @returns {string|null} テキスト内容（テキストが無いときは null）
     */
    function findTextContents(item) {
        if (item.typename === 'TextFrame') return item.contents;
        if (item.typename !== 'GroupItem') return null;

        for (var i = 0; i < item.pageItems.length; i++) {
            var childItem = item.pageItems[i];
            if (childItem.typename === 'TextFrame') return childItem.contents;
            if (childItem.typename === 'GroupItem') {
                var nestedContents = findTextContents(childItem);
                if (nestedContents) return nestedContents;
            }
        }
        return null;
    }

    /**
     * アイテムが属するレイヤー名を返す（「レイヤー 1」などの既定名は除く）
     * @param {PageItem} item - 対象のアイテム
     * @returns {string|null} レイヤー名（使えないときは null）
     */
    function getLayerNameFromItem(item) {
        try {
            var layerName = item.layer.name;
            if (!layerName || DEFAULT_LAYER_NAME_PATTERN.test(layerName)) return null;
            return layerName;
        } catch (e) {
            return null;
        }
    }

    /**
     * アイテムのメモ（note）を返す（空白だけのものは除く）
     * @param {PageItem} item - 対象のアイテム
     * @returns {string|null} メモ（使えないときは null）
     */
    function getNoteFromItem(item) {
        try {
            var noteText = item.note;
            if (!noteText || trimString(String(noteText)) === '') return null;
            return noteText;
        } catch (e) {
            return null;
        }
    }

    /**
     * シンボル名として使えるよう整える（改行・タブを空白に、前後の空白を除き、最大文字数で切る）
     * @param {string|null} name - 元の文字列
     * @returns {string} 整えた名前（元が無いときは空文字）
     */
    function sanitizeSymbolName(name) {
        if (name === null || name === undefined) return '';
        var cleanName = trimString(String(name).replace(/[\r\n\t]+/g, ' '));
        return cleanName.substring(0, SYMBOL_NAME_MAX_LENGTH);
    }

    /**
     * 数値を指定した桁数までゼロで埋める
     * @param {number} num - 数値
     * @param {number} length - 桁数
     * @returns {string} ゼロ埋めした文字列
     */
    function zeroPadding(num, length) {
        var paddedText = String(num);
        while (paddedText.length < length) paddedText = '0' + paddedText;
        return paddedText;
    }

    /**
     * 既存のシンボル名と重ならない名前を返す（重なるときは "_2"、"_3" … を付ける）
     * @param {Document} doc - 対象ドキュメント
     * @param {string} baseName - 元の名前
     * @returns {string} 重複しない名前
     */
    function getUniqueSymbolName(doc, baseName) {
        var uniqueName = baseName;
        var suffixNumber = 2;
        while (symbolExists(doc, uniqueName)) {
            uniqueName = baseName + '_' + suffixNumber;
            suffixNumber++;
        }
        return uniqueName;
    }

    /**
     * 同じ名前のシンボルがあるかを調べる
     * @param {Document} doc - 対象ドキュメント
     * @param {string} symbolName - シンボル名
     * @returns {boolean} あれば true
     */
    function symbolExists(doc, symbolName) {
        /* getByName は見つからないと例外を投げる / getByName throws when not found */
        try {
            doc.symbols.getByName(symbolName);
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * 前後の空白を除く（ES3 に String.trim が無いため）
     * @param {string} text - 元の文字列
     * @returns {string} 前後の空白を除いた文字列
     */
    function trimString(text) {
        return text.replace(/^\s+|\s+$/g, '');
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 設定ダイアログを表示する
     * @returns {SymbolizeSettings|null} 設定（キャンセル時は null）
     */
    function showSettingsDialog() {
        var settingsDialog = new Window('dialog', getLabel(LABELS.dialog.title) + ' ' + SCRIPT_VERSION);
        setupWindow(settingsDialog);

        /* 選択セルの塗り：ライトは濃いグレー、ダークは明るいグレー / Selected cell: dark gray in light UI, bright gray in dark UI */
        ANCHOR_SELECTED_FILL = isLightUI() ? [0.4, 0.4, 0.4, 1] : [0.8, 0.8, 0.8, 1];

        var dialogControls = {
            groupModeRadios: buildGroupModePanel(settingsDialog),
            symbolizeModeRadios: buildSymbolizeModePanel(settingsDialog),
            symbolNameControls: buildSymbolNamePanel(settingsDialog)
        };

        /* 基準点とリンク画像を横に並べる / Place registration point and linked images side by side */
        var columnsGroup = settingsDialog.add('group');
        columnsGroup.orientation = 'row';
        columnsGroup.alignChildren = ['fill', 'fill'];
        columnsGroup.spacing = COLUMN_SPACING;
        dialogControls.anchorWidget = buildReferencePointPanel(columnsGroup);
        dialogControls.linkedImageRadios = buildLinkedImagePanel(columnsGroup);

        bindModeControls(dialogControls);

        var btnRowGroup = settingsDialog.add('group');
        setupRow(btnRowGroup, 'right');
        var btnCancel = btnRowGroup.add('button', undefined, getLabel(LABELS.button.cancel), { name: 'cancel' });
        var btnOK = btnRowGroup.add('button', undefined, getLabel(LABELS.button.ok), { name: 'ok' });
        settingsDialog.cancelElement = btnCancel;
        settingsDialog.defaultElement = btnOK;

        if (settingsDialog.show() !== 1) return null;
        return readDialogSettings(dialogControls);
    }

    /**
     * ダイアログのコントロールから設定を読み取る
     * @param {object} dialogControls - showSettingsDialog() で集めたコントロール
     * @returns {SymbolizeSettings} 設定
     */
    function readDialogSettings(dialogControls) {
        var nameControls = dialogControls.symbolNameControls;
        var prefixValue = trimString(nameControls.prefixInput.text);

        return {
            groupMode: getSelectedRadioValue(dialogControls.groupModeRadios),
            symbolizeMode: getSelectedRadioValue(dialogControls.symbolizeModeRadios),
            defaultPrefix: (prefixValue === '') ? DEFAULT_PREFIX : prefixValue,
            sequenceDigits: getSelectedRadioValue(nameControls.sequenceRadios),
            useTextAsName: nameControls.useTextCheckbox.value,
            referencePoint: REFERENCE_POINTS[dialogControls.anchorWidget.selectedAnchorIndex],
            linkedImagePolicy: getSelectedRadioValue(dialogControls.linkedImageRadios)
        };
    }

    /**
     * 選択範囲の扱いと登録方法の連動、自動登録専用の項目の有効／無効を設定する
     * @param {object} dialogControls - showSettingsDialog() で集めたコントロール
     * @returns {void}
     */
    function bindModeControls(dialogControls) {
        var groupModeRadios = dialogControls.groupModeRadios;
        var symbolizeModeRadios = dialogControls.symbolizeModeRadios;
        var nameControls = dialogControls.symbolNameControls;
        var anchorWidget = dialogControls.anchorWidget;
        var i;

        /**
         * 現在の選択に合わせて、登録方法と自動登録専用の項目を切り替える
         * @returns {void}
         */
        function refreshModeUI() {
            var asGroupMode = (getSelectedRadioValue(groupModeRadios) === GROUP_MODE.AS_GROUP);

            /* まとめて1つにするときは標準ダイアログでの確認に固定する / One symbol from selection always uses the native dialog */
            if (asGroupMode) selectRadioByValue(symbolizeModeRadios, SYMBOLIZE_MODE.INDIVIDUAL);
            for (var k = 0; k < symbolizeModeRadios.length; k++) {
                symbolizeModeRadios[k].enabled = !asGroupMode;
            }

            /* 接頭辞・連番・テキスト流用・基準点は自動登録のときだけ使う / Prefix, sequence, text reuse and registration point apply to automatic registration only */
            var batchEnabled = (getSelectedRadioValue(symbolizeModeRadios) === SYMBOLIZE_MODE.BATCH);
            nameControls.prefixRow.enabled = batchEnabled;
            nameControls.sequenceRow.enabled = batchEnabled;
            nameControls.useTextCheckbox.enabled = batchEnabled;
            /* パネルの enabled は描画に伝わらないので、ウィジェットにも設定する / The panel's enabled state does not reach onDraw */
            anchorWidget.parent.enabled = batchEnabled;
            anchorWidget.enabled = batchEnabled;
        }

        for (i = 0; i < groupModeRadios.length; i++) {
            groupModeRadios[i].onClick = function () {
                selectRadioByValue(groupModeRadios, this.optionValue);
                /* オブジェクトごとに戻したときは自動登録を選び直す / Switch back to automatic registration for per-object mode */
                if (this.optionValue === GROUP_MODE.EACH_ITEM) {
                    selectRadioByValue(symbolizeModeRadios, SYMBOLIZE_MODE.BATCH);
                }
                refreshModeUI();
            };
        }
        for (i = 0; i < symbolizeModeRadios.length; i++) {
            symbolizeModeRadios[i].onClick = function () {
                selectRadioByValue(symbolizeModeRadios, this.optionValue);
                refreshModeUI();
            };
        }

        refreshModeUI();
    }

    /**
     * 選択範囲の扱いパネル（まとめて1つ／オブジェクトごと）を追加する
     * @param {Window} parent - 追加先
     * @returns {RadioButton[]} ラジオボタン
     */
    function buildGroupModePanel(parent) {
        var groupModePanel = addPanel(parent, LABELS.panel.groupMode);
        return addOptionRadios(groupModePanel, [
            { value: GROUP_MODE.AS_GROUP, label: LABELS.radio.asGroup, tooltip: LABELS.tooltip.asGroup },
            { value: GROUP_MODE.EACH_ITEM, label: LABELS.radio.eachItem, tooltip: LABELS.tooltip.eachItem }
        ], DEFAULT_GROUP_MODE);
    }

    /**
     * 登録方法パネル（標準ダイアログで確認／自動登録）を追加する
     * @param {Window} parent - 追加先
     * @returns {RadioButton[]} ラジオボタン
     */
    function buildSymbolizeModePanel(parent) {
        var symbolizeModePanel = addPanel(parent, LABELS.panel.symbolizeMode);
        var symbolizeModeRow = symbolizeModePanel.add('group');
        setupRow(symbolizeModeRow);
        return addOptionRadios(symbolizeModeRow, [
            { value: SYMBOLIZE_MODE.INDIVIDUAL, label: LABELS.radio.individual, tooltip: LABELS.tooltip.individual },
            { value: SYMBOLIZE_MODE.BATCH, label: LABELS.radio.batch, tooltip: LABELS.tooltip.batch }
        ], DEFAULT_SYMBOLIZE_MODE);
    }

    /**
     * シンボル名パネル（接頭辞・連番の桁数・テキスト流用）を追加する
     * @param {Window} parent - 追加先
     * @returns {object} { prefixRow, prefixInput, sequenceRow, sequenceRadios, useTextCheckbox }
     */
    function buildSymbolNamePanel(parent) {
        var symbolNamePanel = addPanel(parent, LABELS.panel.symbolName, 8);

        var prefixRow = symbolNamePanel.add('group');
        setupRow(prefixRow, 'left', 8);
        addFieldLabel(prefixRow, LABELS.fieldLabel.prefix).helpTip = getLabel(LABELS.tooltip.prefix);
        var prefixInput = prefixRow.add('edittext', undefined, DEFAULT_PREFIX);
        prefixInput.preferredSize.width = PREFIX_INPUT_WIDTH;
        prefixInput.helpTip = getLabel(LABELS.tooltip.prefix);
        prefixInput.active = true;

        var sequenceRow = symbolNamePanel.add('group');
        setupRow(sequenceRow, 'left', 8);
        addFieldLabel(sequenceRow, LABELS.fieldLabel.sequence).helpTip = getLabel(LABELS.tooltip.sequence);
        var sequenceOptions = [];
        for (var i = 0; i < SEQUENCE_DIGIT_OPTIONS.length; i++) {
            sequenceOptions.push({
                value: SEQUENCE_DIGIT_OPTIONS[i],
                labelString: zeroPadding(0, SEQUENCE_DIGIT_OPTIONS[i]),
                tooltip: LABELS.tooltip.sequence
            });
        }
        var sequenceRadios = addOptionRadios(sequenceRow, sequenceOptions, DEFAULT_SEQUENCE_DIGITS);

        var useTextCheckbox = symbolNamePanel.add('checkbox', undefined, getLabel(LABELS.checkbox.useTextAsName));
        useTextCheckbox.value = DEFAULT_USE_TEXT_AS_NAME;
        useTextCheckbox.helpTip = getLabel(LABELS.tooltip.useTextAsName);

        return {
            prefixRow: prefixRow,
            prefixInput: prefixInput,
            sequenceRow: sequenceRow,
            sequenceRadios: sequenceRadios,
            useTextCheckbox: useTextCheckbox
        };
    }

    /**
     * 基準点パネル（9軸ウィジェット）を追加する
     * @param {Group} parent - 追加先
     * @returns {Button} 基準点ウィジェット（選択は selectedAnchorIndex）
     */
    function buildReferencePointPanel(parent) {
        var referencePointPanel = addPanel(parent, LABELS.panel.referencePoint);
        referencePointPanel.alignChildren = ['center', 'center'];
        referencePointPanel.helpTip = getLabel(LABELS.tooltip.referencePoint);

        var anchorWidget = addAnchorWidget(referencePointPanel, DEFAULT_REFERENCE_POINT_INDEX);
        anchorWidget.helpTip = getLabel(LABELS.tooltip.referencePoint);
        return anchorWidget;
    }

    /**
     * リンク画像パネル（無視／埋め込んで登録）を追加する
     * @param {Group} parent - 追加先
     * @returns {RadioButton[]} ラジオボタン
     */
    function buildLinkedImagePanel(parent) {
        var linkedImagePanel = addPanel(parent, LABELS.panel.linkedImage, 6);
        linkedImagePanel.helpTip = getLabel(LABELS.tooltip.linkedImage);
        return addOptionRadios(linkedImagePanel, [
            { value: LINKED_IMAGE_POLICY.IGNORE, label: LABELS.radio.ignoreLinked, tooltip: LABELS.tooltip.linkedImage },
            { value: LINKED_IMAGE_POLICY.EMBED, label: LABELS.radio.embedLinked, tooltip: LABELS.tooltip.linkedImage }
        ], DEFAULT_LINKED_IMAGE_POLICY);
    }

    // =========================================
    // ラジオボタン / Radio buttons
    // =========================================

    /**
     * 選択肢ごとにラジオボタンを追加し、クリックで排他選択になるようにする
     * @param {object} parent - 追加先
     * @param {object[]} optionDefs - { value, label（ja/en）または labelString, tooltip（ja/en） } の配列
     * @param {*} selectedValue - 最初に選ぶ値
     * @returns {RadioButton[]} ラジオボタン
     */
    function addOptionRadios(parent, optionDefs, selectedValue) {
        var radios = [];
        for (var i = 0; i < optionDefs.length; i++) {
            var optionDef = optionDefs[i];
            var radio = parent.add('radiobutton', undefined, optionDef.labelString || getLabel(optionDef.label));
            radio.optionValue = optionDef.value;
            radio.helpTip = getLabel(optionDef.tooltip);
            radio.onClick = function () {
                selectRadioByValue(radios, this.optionValue);
            };
            radios.push(radio);
        }
        selectRadioByValue(radios, selectedValue);
        return radios;
    }

    /**
     * 指定した値のラジオボタンだけを選択する
     * @param {RadioButton[]} radios - ラジオボタン
     * @param {*} optionValue - 選ぶ値
     * @returns {void}
     */
    function selectRadioByValue(radios, optionValue) {
        for (var i = 0; i < radios.length; i++) {
            radios[i].value = (radios[i].optionValue === optionValue);
        }
    }

    /**
     * 選択中のラジオボタンの値を返す
     * @param {RadioButton[]} radios - ラジオボタン
     * @returns {*} 選択中の値（未選択のときは先頭の値）
     */
    function getSelectedRadioValue(radios) {
        for (var i = 0; i < radios.length; i++) {
            if (radios[i].value) return radios[i].optionValue;
        }
        return radios[0].optionValue;
    }

    // =========================================
    // 基準点ウィジェット / Anchor widget
    // =========================================

    /**
     * UI が明るいテーマかを判定する
     * @returns {boolean} 明るいテーマなら true（取得できないときは暗い側）
     */
    function isLightUI() {
        try {
            return app.preferences.getRealPreference('uiBrightness') > 0.5;
        } catch (e) {
            return false;
        }
    }

    /**
     * 基準点を選ぶ 3×3 のウィジェットを追加する
     * @param {Panel} parent - 追加先
     * @param {number} anchorIndex - 最初に選ぶセル（0〜8）
     * @returns {Button} 追加したウィジェット
     */
    function addAnchorWidget(parent, anchorIndex) {
        var anchorWidget = parent.add('button', undefined, '');
        anchorWidget.minimumSize = [ANCHOR_WIDGET_SIZE, ANCHOR_WIDGET_SIZE];
        anchorWidget.preferredSize = [ANCHOR_WIDGET_SIZE, ANCHOR_WIDGET_SIZE];
        anchorWidget.maximumSize = [ANCHOR_WIDGET_SIZE, ANCHOR_WIDGET_SIZE];
        anchorWidget.selectedAnchorIndex = anchorIndex;
        anchorWidget.onDraw = function () {
            drawAnchorWidget(this);
        };

        /* クリックしたセルを mousedown で判定する（座標はコントロール基準）/ Hit-test the clicked cell on mousedown (control-relative coordinates) */
        try {
            anchorWidget.addEventListener('mousedown', function (event) {
                if (anchorWidget.enabled === false) return;
                var col = Math.max(0, Math.min(2, Math.floor(event.clientX / (anchorWidget.size[0] / 3))));
                var row = Math.max(0, Math.min(2, Math.floor(event.clientY / (anchorWidget.size[1] / 3))));
                anchorWidget.selectedAnchorIndex = row * 3 + col;
                try { anchorWidget.notify('onDraw'); } catch (e) { }
            });
        } catch (e) { }

        return anchorWidget;
    }

    /**
     * 基準点ウィジェットを描く（外周の□をケイ線でつなぎ、中央は独立）
     * @param {Button} anchorWidget - 描画するウィジェット
     * @returns {void}
     */
    function drawAnchorWidget(anchorWidget) {
        var graphics = anchorWidget.graphics;
        var widgetWidth = anchorWidget.size[0];
        var widgetHeight = anchorWidget.size[1];

        /* 背景をコントロールの地色で塗り、パネルに溶け込ませる / Paint the control's background so the widget blends into the panel */
        try {
            graphics.newPath();
            graphics.rectPath(0, 0, widgetWidth, widgetHeight);
            graphics.fillPath(graphics.backgroundColor);
        } catch (e) { }

        var cellStep = ANCHOR_CELL_SIZE + ANCHOR_CELL_GAP;
        var gridSize = ANCHOR_CELL_SIZE * 3 + ANCHOR_CELL_GAP * 2;
        var originX = Math.round((widgetWidth - gridSize) / 2);
        var originY = Math.round((widgetHeight - gridSize) / 2);

        var cellPositions = [];
        var i;
        for (i = 0; i < 9; i++) {
            cellPositions.push([originX + (i % 3) * cellStep, originY + Math.floor(i / 3) * cellStep]);
        }

        var isEnabled = (anchorWidget.enabled !== false);
        var halfCell = ANCHOR_CELL_SIZE / 2;
        var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, isEnabled ? ANCHOR_LINE_COLOR : ANCHOR_DISABLED_LINE, 1);
        for (i = 0; i < ANCHOR_CONNECTIONS.length; i++) {
            var cellA = cellPositions[ANCHOR_CONNECTIONS[i][0]];
            var cellB = cellPositions[ANCHOR_CONNECTIONS[i][1]];
            graphics.newPath();
            if (ANCHOR_CONNECTIONS[i][1] - ANCHOR_CONNECTIONS[i][0] === 1) {
                /* 横方向：右隣の□へ / Horizontal: to the cell on the right */
                graphics.moveTo(cellA[0] + ANCHOR_CELL_SIZE, cellA[1] + halfCell);
                graphics.lineTo(cellB[0], cellB[1] + halfCell);
            } else {
                /* 縦方向：下の□へ / Vertical: to the cell below */
                graphics.moveTo(cellA[0] + halfCell, cellA[1] + ANCHOR_CELL_SIZE);
                graphics.lineTo(cellB[0] + halfCell, cellB[1]);
            }
            graphics.strokePath(linePen);
        }

        for (i = 0; i < cellPositions.length; i++) {
            drawAnchorCell(graphics, cellPositions[i][0], cellPositions[i][1], i === anchorWidget.selectedAnchorIndex, isEnabled);
        }
    }

    /**
     * 基準点ウィジェットの□を1つ描く（選択中だけ塗る）
     * @param {ScriptUIGraphics} graphics - 描画先
     * @param {number} cellX - 左端
     * @param {number} cellY - 上端
     * @param {boolean} isSelected - 選択中なら true
     * @param {boolean} isEnabled - ウィジェットが有効なら true
     * @returns {void}
     */
    function drawAnchorCell(graphics, cellX, cellY, isSelected, isEnabled) {
        /* 枠を上に重ねるので塗りを先に描く / Fill first so the border draws on top */
        if (isSelected) {
            graphics.newPath();
            graphics.rectPath(cellX, cellY, ANCHOR_CELL_SIZE, ANCHOR_CELL_SIZE);
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, isEnabled ? ANCHOR_SELECTED_FILL : ANCHOR_DISABLED_FILL));
        }
        graphics.newPath();
        graphics.rectPath(cellX, cellY, ANCHOR_CELL_SIZE, ANCHOR_CELL_SIZE);
        graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, isEnabled ? ANCHOR_LINE_COLOR : ANCHOR_DISABLED_LINE, 1));
    }

    main();

})();
