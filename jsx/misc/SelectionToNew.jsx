#targetengine "SelectionToNewEngine"
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトから、新規レイヤー・新規アートボード・新規ドキュメントを作成します。

- 作成対象はラジオボタンで選択（`L` / `A` / `D` キーでも切り替え）
- 作成するレイヤー名・アートボード名・ドキュメント名を指定
- ［元のオブジェクトを残す］で、移動と複製を切り替え
- アートボードは既存の並びを引き継ぎ、方向と間隔を指定して挿入
- ドキュメントはロック・非表示オブジェクトを残すかどうかを選択（保存済みのドキュメントのみ）

詳しい仕様と注意事項は README を参照してください。

*/

/*

### Overview

Creates a new layer, a new artboard or a new document from the selected objects.

- The target is chosen with radio buttons (or the `L` / `A` / `D` keys)
- The new layer, artboard or document is named in the dialog
- Keep original objects switches between moving and duplicating
- Artboards inherit the existing arrangement; direction and spacing are set in the dialog
- Documents can keep locked and hidden objects (saved documents only)

See the README for the full specification and notes.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SelectionToNew";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-07-29";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-07-29";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SelectionToNew.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SelectionToNew.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n0f02f73a748d"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// アートボードの配置ロジック / Artboard layout logic
// AddArtboardPlus.jsx（jsx/artboard/AddArtboardPlus.jsx）から移植
// Ported from AddArtboardPlus.jsx
// Original: Copyright (c) 2018 Takeshi Umeda (noellabo)
// https://dtp-discourse.jp/t/illustrator/99

// =========================================
// セッション保持 / Session state
// =========================================

/* 前回のダイアログ設定。#targetengine を指定しているので、この変数は
   スクリプトの実行が終わってもIllustratorの起動中は残る。
   名前はドキュメントごとに変わるので保持しない
   The last dialog settings. The #targetengine directive keeps this variable alive
   between runs while Illustrator is running. Names are per-document, so they are
   deliberately not kept */
var sessionCreateOptions = (typeof sessionCreateOptions !== "undefined") ? sessionCreateOptions : null;

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 新規アートボードの挿入位置 / Insert position of the new artboard */
    /* true = 現在のアートボードの次 / false = 末尾 */
    /* true = after the current artboard, false = at the end */
    var ARTBOARD_INSERT_AFTER_CURRENT = true;

    /* ダイアログの［方向］の初期選択 / Initial selection of the dialog's Direction */
    /* 0 = 右（横並び） / 1 = 下（縦並び） */
    /* 0 = right (horizontal), 1 = down (vertical) */
    var ARTBOARD_DIRECTION_AXIS = 0;

    /* 新規アートボード名の初期値に付ける接尾辞 / Suffix for the new artboard's default name */
    var ARTBOARD_NAME_SUFFIX = "_new";

    /* 複製ドキュメントのファイル名の初期値に付ける接尾辞
       Suffix for the duplicated document's default file name */
    var DOCUMENT_NAME_SUFFIX = "_selection";

    /* 作成対象を切り替えるキーボードショートカット
       Keyboard shortcuts that switch the create target */
    var SHORTCUT_TARGETS = { L: "layer", A: "artboard", D: "document" };

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 実行環境の言語を判定します。
     *
     * @returns {string} 日本語環境なら "ja"、それ以外は "en"。
     */
    function getCurrentLang() {
        /* ja で始まるロケールだけ日本語として扱う / Treat only ja-prefixed locales as Japanese */
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var lang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "選択オブジェクトから新規作成", en: "New from Selection" }
        },
        panel: {
            createTarget: { ja: "作成するものと名前", en: "Create" },
            artboardLayout: { ja: "アートボードの配置", en: "Artboard placement" },
            options: { ja: "オプション", en: "Options" }
        },
        label: {
            direction: { ja: "方向：", en: "Direction:" },
            spacing: { ja: "間隔：", en: "Spacing:" }
        },
        radio: {
            targetLayer: { ja: "レイヤー", en: "Layer" },
            targetArtboard: { ja: "アートボード", en: "Artboard" },
            targetDocument: { ja: "ドキュメント", en: "Document" },
            directionRight: { ja: "右", en: "Right" },
            directionDown: { ja: "下", en: "Down" }
        },
        checkbox: {
            duplicate: { ja: "元のオブジェクトを残す", en: "Keep original objects" },
            includeLocked: { ja: "ロックされたオブジェクトを含める", en: "Include locked objects" },
            includeHidden: { ja: "非表示オブジェクトを含める", en: "Include hidden objects" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        artwork: {
            newLayerName: { ja: "新規レイヤー", en: "New Layer" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "オブジェクトが選択されていません。", en: "No objects are selected." },
            artboardLimit: {
                ja: "アートボードの最大数を超えるため、作成できません。",
                en: "Cannot create: the maximum number of artboards would be exceeded."
            },
            noSpace: {
                ja: "十分なスペースがないため、アートボードを作成できません。",
                en: "Cannot create the artboard: there is not enough space."
            },
            needsSave: {
                ja: "ドキュメントを複製するため、先に保存してください。",
                en: "Save the document first so it can be duplicated."
            },
            sameAsSource: {
                ja: "元のファイルと同じ名前は指定できません。別の名前を入力してください。",
                en: "The name must differ from the source file. Enter another name."
            },
            overwrite: {
                ja: "同名のファイルがすでにあります。上書きしますか？",
                en: "A file with that name already exists. Overwrite it?"
            },
            unexpected: {
                ja: "エラーが発生したため、処理を実行できませんでした。\nエラー内容：",
                en: "Processing could not be completed because an error occurred.\nError:"
            }
        },
        tip: {
            targetLayer: {
                ja: "選択オブジェクトを収めた新規レイヤーを、最前面に作成します。",
                en: "Creates a new layer at the top holding the selected objects."
            },
            targetArtboard: {
                ja: "アクティブアートボードと同じサイズの新規アートボードを作成し、相対位置を保って配置します。",
                en: "Creates a new artboard matching the active one and keeps the relative position."
            },
            targetDocument: {
                ja: "選択オブジェクトだけを残した複製ドキュメントを作成します。保存済みのドキュメントでのみ実行できます。",
                en: "Creates a duplicate document containing only the selection. Saved documents only."
            },
            extension: {
                ja: "元のファイルと同じ拡張子で保存します。",
                en: "Saved with the same extension as the source file."
            },
            ok: {
                ja: "選択した内容で作成します。",
                en: "Create with the selected settings."
            },
            cancel: {
                ja: "何もせずに閉じます。",
                en: "Close without doing anything."
            },
            layerName: {
                ja: "同名のレイヤーがある場合は連番が付きます。",
                en: "A numeric suffix is added when a layer with that name exists."
            },
            artboardName: {
                ja: "現在のアートボードの次に、同じサイズで作成します。",
                en: "Created next to the current artboard, at the same size."
            },
            documentName: {
                ja: "元のファイルと同じ場所に作成します。同名のファイルがある場合は確認します。",
                en: "Created next to the source file. An existing file prompts for confirmation."
            },
            direction: {
                ja: "新規アートボードを現在のアートボードのどちら側に並べるかを選びます。",
                en: "Choose which side of the current artboard the new one goes."
            },
            spacing: {
                ja: "アートボードどうしの間隔。初期値は既存の並びから推定した値で、定規の単位で入力します。",
                en: "Gap between artboards, in the ruler unit. The default is inferred from the existing layout."
            },
            duplicate: {
                ja: "ドキュメントは元のドキュメントを残すため、常に複製になります。",
                en: "Document always keeps the source document, so it is always a copy."
            },
            includeObjects: {
                ja: "ドキュメント作成時のみ。現在のアートボード上にある最上位のオブジェクトが対象です。",
                en: "Document only. Applies to top-level objects on the current artboard."
            }
        },
        progress: {
            title: { ja: "処理中", en: "Working" },
            duplicating: { ja: "ドキュメントを複製しています…", en: "Duplicating the document..." },
            deleting: { ja: "選択オブジェクト以外を削除しています…", en: "Removing everything but the selection..." },
            reopening: { ja: "元のドキュメントを開き直しています…", en: "Reopening the source document..." }
        }
    };

    /**
     * ラベル定義のリーフ（ja/en）を、現在の言語の文言に解決します。
     *
     * @param {object} labelEntry - LABELS のリーフ（{ ja, en }）。
     * @returns {string} 現在の言語の文言。{slash} は / に置き換えます。
     */
    function L(labelEntry) {
        /* 日本語が未定義なら英語にフォールバック / Fall back to English when ja is missing */
        var text = (labelEntry && labelEntry[lang]) || (labelEntry && labelEntry.en) || "";
        return text.replace(/\{slash\}/g, "/");
    }

    // =========================================
    // 矩形の共通処理 / Rect helpers
    // =========================================

    /**
     * 矩形の中心座標を求めます。
     *
     * @param {Array<number>} rect - [左, 上, 右, 下]。
     * @returns {Array<number>} [X, Y] の中心座標。
     */
    function getRectCenter(rect) {
        return [(rect[0] + rect[2]) / 2, (rect[1] + rect[3]) / 2];
    }

    /**
     * 座標が矩形の内側にあるかを判定します。
     *
     * @param {Array<number>} point - [X, Y] の座標。
     * @param {Array<number>} rect - [左, 上, 右, 下]。
     * @returns {boolean} 内側にある場合は true。
     */
    function isPointInsideRect(point, rect) {
        /* Y軸は上→下なので、上端以下・下端以上で判定する
           The Y axis is top-down, so compare against top as max and bottom as min */
        return point[0] >= rect[0] && point[0] <= rect[2] &&
               point[1] <= rect[1] && point[1] >= rect[3];
    }

    /**
     * 複数オブジェクトを囲む矩形を求めます。
     *
     * @param {Array<PageItem>} items - 対象のオブジェクト。
     * @returns {Array<number>|null} [左, 上, 右, 下]。対象が空の場合は null。
     */
    function getUnionBounds(items) {
        if (items.length === 0) return null;

        /* 1つずつ取り込みながら外側へ広げる / Grow the rect outward item by item */
        var firstBounds = items[0].geometricBounds;
        var unionBounds = [firstBounds[0], firstBounds[1], firstBounds[2], firstBounds[3]];

        for (var i = 1; i < items.length; i++) {
            var itemBounds = items[i].geometricBounds;
            if (itemBounds[0] < unionBounds[0]) unionBounds[0] = itemBounds[0];
            if (itemBounds[1] > unionBounds[1]) unionBounds[1] = itemBounds[1];
            if (itemBounds[2] > unionBounds[2]) unionBounds[2] = itemBounds[2];
            if (itemBounds[3] < unionBounds[3]) unionBounds[3] = itemBounds[3];
        }
        return unionBounds;
    }

    // =========================================
    // 選択オブジェクトの取得 / Collecting the selection
    // =========================================

    /**
     * 選択オブジェクトを配列にスナップショットします。
     *
     * `doc.selection` は参照するたびに現在の選択状態から作り直されるため、
     * move() で選択が変化するループ内で直接参照すると対象を取りこぼします。
     *
     * @param {Document} doc - 対象のドキュメント。
     * @returns {Array<PageItem>} 選択オブジェクトの配列。
     */
    function snapshotSelection(doc) {
        /* ライブな selection を配列へ写し取る / Copy the live selection into a plain array */
        var selectedItems = [];
        var liveSelection = doc.selection;

        for (var i = 0; i < liveSelection.length; i++) {
            selectedItems.push(liveSelection[i]);
        }
        return selectedItems;
    }

    /**
     * 重ね順の比較に使うソートキーを取得します。
     *
     * レイヤーとコンテナそれぞれの zOrderPosition を組み合わせ、
     * 値が大きいほど前面（上）になるようにします。
     *
     * @param {PageItem} item - 対象のオブジェクト。
     * @returns {{layerOrder: number, itemOrder: number}} 並べ替え用のキー。
     */
    function getStackOrderKey(item) {
        /* zOrderPosition を持たないアイテムもあるため、まとめて保護する
           Some items expose no zOrderPosition, so guard the whole lookup */
        try {
            return { layerOrder: item.layer.zOrderPosition, itemOrder: item.zOrderPosition };
        } catch (e) {
            return { layerOrder: 0, itemOrder: 0 };
        }
    }

    /**
     * オブジェクトを前面から背面の順（元の重ね順）に並べ替えます。
     *
     * @param {Array<PageItem>} items - 並べ替えるオブジェクトの配列。
     * @returns {Array<PageItem>} 前面が先頭になるよう並べ替えた配列。
     */
    function sortByStackOrder(items) {
        /* 並べ替えキーを先に作る / Build the sort keys up front */
        var sortEntries = [];

        for (var i = 0; i < items.length; i++) {
            sortEntries.push({ item: items[i], key: getStackOrderKey(items[i]), index: i });
        }

        sortEntries.sort(function (entryA, entryB) {
            if (entryA.key.layerOrder !== entryB.key.layerOrder) {
                return entryB.key.layerOrder - entryA.key.layerOrder;
            }
            if (entryA.key.itemOrder !== entryB.key.itemOrder) {
                return entryB.key.itemOrder - entryA.key.itemOrder;
            }
            /* キーが同じときは元の並び順を保つ / Keep the original order for ties */
            return entryA.index - entryB.index;
        });

        var sortedItems = [];
        for (var j = 0; j < sortEntries.length; j++) {
            sortedItems.push(sortEntries[j].item);
        }
        return sortedItems;
    }

    /**
     * オブジェクト群の中心を含むアートボードのインデックスを求めます。
     *
     * @param {Document} doc - 対象のドキュメント。
     * @param {Array<PageItem>} items - 対象のオブジェクト。
     * @param {number} fallbackIndex - 見つからないときに返すインデックス。
     * @returns {number} アートボードのインデックス。
     */
    function findArtboardIndexForItems(doc, items, fallbackIndex) {
        var unionBounds = getUnionBounds(items);
        if (unionBounds === null) return fallbackIndex;

        /* 中心が載っているアートボードを探す / Find the artboard holding the center */
        var center = getRectCenter(unionBounds);
        for (var i = 0; i < doc.artboards.length; i++) {
            if (isPointInsideRect(center, doc.artboards[i].artboardRect)) return i;
        }
        return fallbackIndex;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ダイアログの設定値。
     *
     * @typedef {object} CreateOptions
     * @property {string} createTarget - 作成対象（"layer" / "artboard" / "document"）。
     * @property {string} createName - 新規レイヤー名／アートボード名／ファイル名（拡張子なし）。
     * @property {number} directionAxis - アートボードを並べる方向（0=右 / 1=下）。
     * @property {number} spacingPt - アートボードどうしの間隔（pt）。
     * @property {boolean} useDuplicate - 元のオブジェクトを残して複製する場合は true。
     * @property {boolean} includeLocked - ロックされたオブジェクトも残す場合は true。
     * @property {boolean} includeHidden - 非表示オブジェクトも残す場合は true。
     */

    // =========================================
    // 単位 / Unit
    // =========================================

    /**
     * 現在の定規単位の pt 換算係数とラベルを取得します。
     *
     * @returns {{factor: number, label: string}} 1単位あたりのpt数と単位名。
     */
    function getRulerUnitInfo() {
        var rulerType = app.preferences.getIntegerPreference('rulerType');

        switch (rulerType) {
            case 0: return { factor: 72.0, label: 'inch' };            /* インチ / inch */
            case 1: return { factor: 72.0 / 25.4, label: 'mm' };       /* ミリ / mm */
            case 2: return { factor: 1.0, label: 'pt' };               /* ポイント / point */
            case 3: return { factor: 12.0, label: 'pica' };            /* パイカ / pica */
            case 4: return { factor: 72.0 / 2.54, label: 'cm' };       /* センチ / cm */
            case 5: return { factor: 72.0 / 25.4 * 0.25, label: 'Q' }; /* 歯 / Q */
            case 6: return { factor: 1.0, label: 'px' };               /* ピクセル / pixel */
            default: return { factor: 1.0, label: 'pt' };
        }
    }

    /**
     * 表示用に小数を整えます。
     *
     * @param {number} value - 整形する数値。
     * @returns {string} 小数第2位までに丸めた文字列。
     */
    function formatSpacingValue(value) {
        return String(Math.round(value * 100) / 100);
    }

    /**
     * 入力欄の値を↑↓キーで増減できるようにします。
     *
     * Shiftで±10（10の倍数にスナップ）、Option/Altで±0.1、通常は±1。負の値は0で止めます。
     *
     * @param {EditText} editText - 対象の入力欄。
     * @returns {void}
     */
    function changeValueByArrowKey(editText) {
        editText.addEventListener('keydown', function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;

            if (keyboard.shiftKey) {
                /* Shift押下時は10の倍数にスナップ / Snap to multiples of 10 with Shift */
                if (event.keyName == 'Up') {
                    value = Math.ceil((value + 1) / 10) * 10;
                    event.preventDefault();
                } else if (event.keyName == 'Down') {
                    value = Math.floor((value - 1) / 10) * 10;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            } else if (keyboard.altKey) {
                /* Option押下時は0.1単位 / Step by 0.1 with Option */
                if (event.keyName == 'Up') {
                    value += 0.1;
                    event.preventDefault();
                } else if (event.keyName == 'Down') {
                    value -= 0.1;
                    event.preventDefault();
                }
            } else {
                if (event.keyName == 'Up') {
                    value += 1;
                    event.preventDefault();
                } else if (event.keyName == 'Down') {
                    value -= 1;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            }

            if (value < 0) value = 0;
            editText.text = keyboard.altKey
                ? String(Math.round(value * 10) / 10)
                : String(Math.round(value));
        });
    }

    /**
     * 作成対象ごとの名前の初期値を組み立てます。
     *
     * @param {Document} doc - 対象のドキュメント。
     * @returns {{layer: string, artboard: string, document: string}} 対象ごとの初期値。
     */
    function buildDefaultNames(doc) {
        /* アートボードは基準になるアートボード名、ドキュメントは元のファイル名から作る
           The artboard follows the active artboard, the document follows the file name */
        var activeArtboard = doc.artboards[doc.artboards.getActiveArtboardIndex()];

        return {
            layer: L(LABELS.artwork.newLayerName),
            artboard: activeArtboard.name + ARTBOARD_NAME_SUFFIX,
            document: getFileBaseName(doc.name) + DOCUMENT_NAME_SUFFIX
        };
    }

    /* ラジオボタンの幅。日英どちらの文言も収まる値にそろえて、
       3行の入力欄が縦に並ぶようにする
       Radio width: fixed so the three input fields line up, and wide enough
       for both languages */
    var TARGET_RADIO_WIDTH = 120;

    /**
     * ラジオボタンと名前の入力欄を横一列に並べた、作成対象1つ分の行を作ります。
     *
     * 何の名前かはラジオボタンのラベルで分かるので、入力欄に見出しは付けません。
     *
     * @param {Panel} parentPanel - 追加先のパネル。
     * @param {object} radioEntry - ラジオボタンのラベル定義。
     * @param {string} defaultName - 入力欄の初期値。
     * @param {object} inputTipEntry - 入力欄のツールチップ定義。
     * @param {object} radioTipEntry - ラジオボタンのツールチップ定義。
     * @returns {{radio: RadioButton, input: EditText, group: Group}} 作った部品。
     */
    function addTargetRow(parentPanel, radioEntry, defaultName, inputTipEntry, radioTipEntry) {
        var rowGroup = parentPanel.add("group");
        rowGroup.orientation = "row";
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.spacing = 8;

        var targetRadio = rowGroup.add("radiobutton", undefined, L(radioEntry));
        targetRadio.preferredSize.width = TARGET_RADIO_WIDTH;
        targetRadio.helpTip = L(radioTipEntry);

        var nameInput = rowGroup.add("edittext", undefined, defaultName);
        nameInput.characters = 18;
        nameInput.helpTip = L(inputTipEntry);

        return { radio: targetRadio, input: nameInput, group: rowGroup };
    }

    /**
     * 何を作成するかを選択するダイアログを表示します。
     *
     * @param {Document} doc - 対象のドキュメント。名前の初期値を作るのに使います。
     * @returns {CreateOptions|null} 選択された設定。キャンセルされた場合は null。
     */
    function showCreateTargetDialog(doc) {
        var defaultNames = buildDefaultNames(doc);
        var documentExtension = getFileExtension(doc.name) || ".ai";

        /* タイトルにバージョンを添える / Show the version next to the title */
        var dialog = new Window("dialog", L(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];
        dialog.spacing = 12;
        dialog.margins = 16;

        /* 作成対象パネル。1行につき「ラジオボタン＋名前の入力欄」を並べる
           Create target panel: one row per target, radio + name input */
        var createTargetPanel = dialog.add("panel", undefined, L(LABELS.panel.createTarget));
        createTargetPanel.orientation = "column";
        createTargetPanel.alignChildren = ["left", "top"];
        createTargetPanel.spacing = 8;
        createTargetPanel.margins = [15, 20, 15, 15];

        var layerRow = addTargetRow(createTargetPanel, LABELS.radio.targetLayer,
            defaultNames.layer, LABELS.tip.layerName, LABELS.tip.targetLayer);
        var artboardRow = addTargetRow(createTargetPanel, LABELS.radio.targetArtboard,
            defaultNames.artboard, LABELS.tip.artboardName, LABELS.tip.targetArtboard);
        var documentRow = addTargetRow(createTargetPanel, LABELS.radio.targetDocument,
            defaultNames.document, LABELS.tip.documentName, LABELS.tip.targetDocument);

        /* ドキュメントだけは拡張子を添えて示す / Show the extension next to the document name */
        var extensionLabel = documentRow.group.add("statictext", undefined, documentExtension);
        extensionLabel.helpTip = L(LABELS.tip.extension);

        /* アートボードの配置パネル / Artboard placement panel */
        var rulerUnit = getRulerUnitInfo();
        var autoSpacingPt = computeAutoSpacingPt(doc,
            app.preferences.getRealPreference('plugin/ArtboardRearrange/ArtboardSpacing'));

        var artboardLayoutPanel = dialog.add("panel", undefined, L(LABELS.panel.artboardLayout));
        artboardLayoutPanel.orientation = "row";
        artboardLayoutPanel.alignChildren = ["left", "center"];
        artboardLayoutPanel.spacing = 8;
        artboardLayoutPanel.margins = [15, 20, 15, 15];

        var directionLabel = artboardLayoutPanel.add("statictext", undefined, L(LABELS.label.direction));

        /* ラジオボタンは同じグループに入れて排他にする
           Keep both radios in one group so ScriptUI makes them exclusive */
        var directionGroup = artboardLayoutPanel.add("group");
        directionGroup.orientation = "row";
        directionGroup.spacing = 8;
        var directionRightRadio = directionGroup.add("radiobutton", undefined, L(LABELS.radio.directionRight));
        var directionDownRadio = directionGroup.add("radiobutton", undefined, L(LABELS.radio.directionDown));

        var spacingLabel = artboardLayoutPanel.add("statictext", undefined, L(LABELS.label.spacing));
        var spacingInput = artboardLayoutPanel.add("edittext", undefined,
            formatSpacingValue(autoSpacingPt / rulerUnit.factor));
        spacingInput.characters = 5;
        changeValueByArrowKey(spacingInput);
        var spacingUnitLabel = artboardLayoutPanel.add("statictext", undefined, rulerUnit.label);

        directionRightRadio.helpTip = L(LABELS.tip.direction);
        directionDownRadio.helpTip = L(LABELS.tip.direction);
        spacingInput.helpTip = L(LABELS.tip.spacing);

        /* オプションパネル / Options panel */
        var optionPanel = dialog.add("panel", undefined, L(LABELS.panel.options));
        optionPanel.orientation = "column";
        optionPanel.alignChildren = ["left", "top"];
        optionPanel.spacing = 8;
        optionPanel.margins = [15, 20, 15, 15];

        var duplicateCheckbox = optionPanel.add("checkbox", undefined, L(LABELS.checkbox.duplicate));
        var includeLockedCheckbox = optionPanel.add("checkbox", undefined, L(LABELS.checkbox.includeLocked));
        var includeHiddenCheckbox = optionPanel.add("checkbox", undefined, L(LABELS.checkbox.includeHidden));

        /* ディム表示になる理由はUIから読み取れないので、ツールチップで補う
           Nothing on screen explains why a control is dimmed, so tooltips fill that in */
        duplicateCheckbox.helpTip = L(LABELS.tip.duplicate);
        includeLockedCheckbox.helpTip = L(LABELS.tip.includeObjects);
        includeHiddenCheckbox.helpTip = L(LABELS.tip.includeObjects);

        /* 前回の設定があれば復元する / Restore the previous settings when there are any */
        var savedOptions = sessionCreateOptions;
        duplicateCheckbox.value = (savedOptions !== null) && savedOptions.useDuplicate;
        includeLockedCheckbox.value = (savedOptions !== null) && savedOptions.includeLocked;
        includeHiddenCheckbox.value = (savedOptions !== null) && savedOptions.includeHidden;

        var savedDirectionAxis = (savedOptions !== null) ? savedOptions.directionAxis : ARTBOARD_DIRECTION_AXIS;
        directionDownRadio.value = (savedDirectionAxis === 1);
        directionRightRadio.value = !directionDownRadio.value;

        /**
         * 選ばれている作成対象を読み取ります。
         *
         * @returns {string} "layer" / "artboard" / "document"。
         */
        function readCreateTarget() {
            if (artboardRow.radio.value) return "artboard";
            if (documentRow.radio.value) return "document";
            return "layer";
        }

        /**
         * 名前欄とオプションの有効・無効を作成対象に合わせて切り替えます。
         *
         * 名前欄は3つとも並べたまま、対象の行だけを有効にします。
         * ［元のオブジェクトを残す］はドキュメント作成では常に複製になるため、
         * ［アートボードの配置］はアートボード作成のときだけ意味があるので、
         * それぞれ対象外のときはディム表示にします。
         * ［ロック／非表示オブジェクト］はレイヤーとアートボードでは対象範囲を
         * 定義できないため、ドキュメント作成のときだけ有効にします。
         *
         * @returns {void}
         */
        function syncOptionStates() {
            var createTarget = readCreateTarget();
            var isDocumentTarget = (createTarget === "document");

            /* 名前欄は行ごとに、対象のものだけ有効にする / Enable the matching row only */
            layerRow.input.enabled = (createTarget === "layer");
            artboardRow.input.enabled = (createTarget === "artboard");
            documentRow.input.enabled = isDocumentTarget;

            /* 配置はアートボード作成のときだけ意味がある
               Placement only applies when creating an artboard */
            var isArtboardTarget = (createTarget === "artboard");
            directionLabel.enabled = isArtboardTarget;
            directionRightRadio.enabled = isArtboardTarget;
            directionDownRadio.enabled = isArtboardTarget;
            spacingLabel.enabled = isArtboardTarget;
            spacingInput.enabled = isArtboardTarget;
            spacingUnitLabel.enabled = isArtboardTarget;

            duplicateCheckbox.enabled = !isDocumentTarget;
            includeLockedCheckbox.enabled = isDocumentTarget;
            includeHiddenCheckbox.enabled = isDocumentTarget;
        }

        /**
         * 作成対象を1つだけ選ばれた状態にします。
         *
         * ScriptUIのラジオボタンは同じ親コンテナ内でしか排他になりません。
         * ここでは行ごとにグループを分けているため、自前で他をオフにします。
         *
         * @param {object} activeRow - 選択された行（addTargetRow の戻り値）。
         * @returns {void}
         */
        function selectCreateTarget(activeRow) {
            layerRow.radio.value = (activeRow === layerRow);
            artboardRow.radio.value = (activeRow === artboardRow);
            documentRow.radio.value = (activeRow === documentRow);
            syncOptionStates();
        }

        layerRow.radio.onClick = function () { selectCreateTarget(layerRow); };
        artboardRow.radio.onClick = function () { selectCreateTarget(artboardRow); };
        documentRow.radio.onClick = function () { selectCreateTarget(documentRow); };

        var rowsByTarget = { layer: layerRow, artboard: artboardRow, document: documentRow };
        var savedTarget = (savedOptions !== null) ? savedOptions.createTarget : "layer";
        var initialRow = rowsByTarget[savedTarget] || layerRow;
        selectCreateTarget(initialRow);

        /* L / A / D で作成対象を切り替える。名前欄の入力中は文字として入ってほしいので、
           イベントの発生元が入力欄のときは何もしない
           L / A / D switch the target. While a name field has focus the keys must type
           normally, so bail out when the event comes from an edittext */
        dialog.addEventListener("keydown", function (event) {
            if (event.target && event.target.type === "edittext") return;

            var pressedKey = event.keyName ? String(event.keyName).toUpperCase() : "";
            var shortcutTarget = SHORTCUT_TARGETS[pressedKey];
            if (!shortcutTarget) return;

            selectCreateTarget(rowsByTarget[shortcutTarget]);
            event.preventDefault();
        });

        /* ショートカットが最初から効くよう、フォーカスはラジオボタンに置く
           Focus a radio so the shortcuts work without clicking first */
        dialog.onShow = function () {
            initialRow.radio.active = true;
        };

        /* ボタン行 / Button row */
        var dialogButtonGroup = dialog.add("group");
        dialogButtonGroup.orientation = "row";
        dialogButtonGroup.alignment = ["right", "top"];
        var cancelButton = dialogButtonGroup.add("button", undefined, L(LABELS.button.cancel), { name: "cancel" });
        var okButton = dialogButtonGroup.add("button", undefined, L(LABELS.button.ok), { name: "ok" });
        cancelButton.helpTip = L(LABELS.tip.cancel);
        okButton.helpTip = L(LABELS.tip.ok);

        if (dialog.show() !== 1) return null;

        var createTarget = readCreateTarget();

        /* 空欄のときは初期値に戻す / Fall back to the default when left blank */
        var createName = rowsByTarget[createTarget].input.text.replace(/^\s+|\s+$/g, "");
        if (createName === "") createName = defaultNames[createTarget];

        /* 次回のために保持する。ディム中の値もそのまま覚えておかないと、
           作成対象を切り替えて戻したときにチェックが消えてしまう
           Remember the raw values: zeroing the dimmed ones would clear the
           checkboxes when the user switches targets and comes back */
        var directionAxis = directionDownRadio.value ? 1 : 0;

        /* 入力が空・不正なら推定値に戻す / Fall back to the inferred value on blank or bad input */
        var spacingValue = parseFloat(spacingInput.text);
        var spacingPt = isNaN(spacingValue) ? autoSpacingPt : (spacingValue * rulerUnit.factor);

        /* 間隔はドキュメントごとに変わるので保持しない
           The spacing is document-specific, so it is not remembered */
        sessionCreateOptions = {
            createTarget: createTarget,
            directionAxis: directionAxis,
            useDuplicate: duplicateCheckbox.value,
            includeLocked: includeLockedCheckbox.value,
            includeHidden: includeHiddenCheckbox.value
        };

        /* ディム中の値は無視する / Ignore the values while the controls are dimmed */
        return {
            createTarget: createTarget,
            createName: createName,
            directionAxis: directionAxis,
            spacingPt: spacingPt,
            useDuplicate: (duplicateCheckbox.enabled && duplicateCheckbox.value),
            includeLocked: (includeLockedCheckbox.enabled && includeLockedCheckbox.value),
            includeHidden: (includeHiddenCheckbox.enabled && includeHiddenCheckbox.value)
        };
    }

    // =========================================
    // レイヤーを作成 / Create a layer
    // =========================================

    /**
     * 同じ名前のレイヤーがすでに存在するかを判定します。
     *
     * @param {Document} doc - 対象のドキュメント。
     * @param {string} layerName - 探すレイヤー名。
     * @returns {boolean} 存在する場合は true。
     */
    function layerNameExists(doc, layerName) {
        /* 最上位レイヤーだけを見る / Only top-level layers are checked */
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === layerName) return true;
        }
        return false;
    }

    /**
     * 重複しないレイヤー名を作ります（「新規レイヤー 2」のように連番を付与）。
     *
     * @param {Document} doc - 対象のドキュメント。
     * @param {string} baseName - 基準となるレイヤー名。
     * @returns {string} 重複しないレイヤー名。
     */
    function makeUniqueLayerName(doc, baseName) {
        if (!layerNameExists(doc, baseName)) return baseName;

        /* 空いている連番を探す / Look for a free sequence number */
        var suffixNumber = 2;
        while (layerNameExists(doc, baseName + " " + suffixNumber)) {
            suffixNumber++;
        }
        return baseName + " " + suffixNumber;
    }

    /**
     * 選択オブジェクトを収めた新規レイヤーを作成します。
     *
     * @param {Document} doc - 対象のドキュメント。
     * @param {boolean} useDuplicate - true なら元のオブジェクトを残して複製する。
     * @param {string} layerName - 新規レイヤーの名前。
     * @returns {void}
     */
    function createLayerFromSelection(doc, useDuplicate, layerName) {
        /* ループ前にスナップショットを取る（move で選択が変化するため）
           Snapshot before the loop: move() changes the live selection */
        var selectedItems = snapshotSelection(doc);
        if (selectedItems.length === 0) {
            alert(L(LABELS.alert.noSelection));
            return;
        }

        selectedItems = sortByStackOrder(selectedItems);

        /* 最前面に新規レイヤーを作る / Add the new layer at the top */
        var newLayer = doc.layers.add();
        newLayer.name = makeUniqueLayerName(doc, layerName);

        /* 前面のものから PLACEATEND で送ると元の重ね順が保たれる
           Sending front-to-back with PLACEATEND preserves the original order */
        var placedItems = [];
        for (var i = 0; i < selectedItems.length; i++) {
            if (useDuplicate) {
                placedItems.push(selectedItems[i].duplicate(newLayer, ElementPlacement.PLACEATEND));
            } else {
                selectedItems[i].move(newLayer, ElementPlacement.PLACEATEND);
                placedItems.push(selectedItems[i]);
            }
        }

        /* 新規レイヤー側のオブジェクトを選択状態にする
           Select the items that ended up on the new layer */
        doc.selection = placedItems;
    }

    // =========================================
    // アートボードの配置計算 / Artboard layout math
    // =========================================

    /**
     * 軸方向のアートボードサイズを取得します。
     *
     * @param {Array<number>} artboardRect - artboardRect（[左, 上, 右, 下]）。
     * @param {number} axisIndex - 0 なら幅、1 なら高さ。
     * @returns {number} 指定軸のサイズ。
     */
    function getArtboardAxisSize(artboardRect, axisIndex) {
        /* Y軸は上→下なので絶対値をとる / The Y axis is top-down, so take the absolute value */
        return (axisIndex === 0)
            ? (artboardRect[2] - artboardRect[0])
            : Math.abs(artboardRect[3] - artboardRect[1]);
    }

    /**
     * 隣り合う2枚のアートボードから主軸を判定します。
     *
     * @param {Array<number>} firstRect - 1枚目の artboardRect。
     * @param {Array<number>} secondRect - 2枚目の artboardRect。
     * @returns {number} 左端が同じなら 1（縦並び）、違えば 0（横並び）。
     */
    function detectPrimaryAxisIndex(firstRect, secondRect) {
        /* 左端が揃っていれば縦に積まれている / A shared left edge means a vertical stack */
        return (firstRect[0] === secondRect[0]) ? 1 : 0;
    }

    /**
     * 既存アートボードの並びから現在の間隔（pt）を推定します。
     *
     * @param {Document} doc - 対象のドキュメント。
     * @param {number} fallbackSpacingPt - 2枚未満のときに使う値（pt）。
     * @returns {number} 推定した間隔（pt）。
     */
    function computeAutoSpacingPt(doc, fallbackSpacingPt) {
        var artboardList = doc.artboards;
        if (artboardList.length < 2) return fallbackSpacingPt;

        /* ピッチからアートボード自身のサイズを引いた残りが隙間
           The gap is the pitch minus the artboard's own size */
        var firstRect = artboardList[0].artboardRect;
        var secondRect = artboardList[1].artboardRect;
        var primaryAxisIndex = detectPrimaryAxisIndex(firstRect, secondRect);
        var pitch = Math.abs(secondRect[primaryAxisIndex] - firstRect[primaryAxisIndex]);
        var spacing = pitch - getArtboardAxisSize(firstRect, primaryAxisIndex);
        return (spacing < 0) ? 0 : spacing;
    }

    /**
     * 最大カンバス範囲を取得します。
     *
     * Original idea by OMOTI
     * https://forums.adobe.com/thread/2459293
     *
     * @param {Document} doc - 対象のドキュメント。
     * @returns {Array<number>} [左, 上, 右, 下]。
     */
    function getLargestCanvasBounds(doc) {
        var LARGEST_SIZE = 16383;

        /* 一時テキストの変換行列からカンバス左上を得る
           Read the canvas origin from a temporary text frame's matrix */
        var tempLayer = doc.layers.add();
        var tempText = tempLayer.textFrames.add();
        var canvasLeft = tempText.matrix.mValueTX;
        var canvasTop = tempText.matrix.mValueTY;
        tempLayer.remove();

        return [canvasLeft, canvasTop, canvasLeft + LARGEST_SIZE, canvasTop - LARGEST_SIZE];
    }

    /**
     * 既存の並びから引き継げるグリッドを検出します。
     *
     * 既存の並びが指定方向と一致する場合だけ、そのピッチ（gridStep）と
     * 折り返し位置（columns）を読み取ります。
     *
     * @param {Document} doc - 対象のドキュメント。
     * @param {Array<number>} gridStep - 既定のグリッド移動量 [X, Y]。破壊的に更新します。
     * @param {number} primaryAxisIndex - 主軸（0=横 / 1=縦）。
     * @param {number} secondaryAxisIndex - 副軸。
     * @returns {{canInherit: boolean, columns: number}} 検出結果。
     */
    function detectInheritedGrid(doc, gridStep, primaryAxisIndex, secondaryAxisIndex) {
        var artboards = doc.artboards;
        if (artboards.length < 2) return { canInherit: false, columns: 0 };

        var firstRect = artboards[0].artboardRect;
        var secondRect = artboards[1].artboardRect;
        if (detectPrimaryAxisIndex(firstRect, secondRect) !== primaryAxisIndex) {
            return { canInherit: false, columns: 0 };
        }

        /* 実際のピッチを主軸の移動量として採用 / Adopt the real pitch as the primary step */
        gridStep[primaryAxisIndex] = secondRect[primaryAxisIndex] - firstRect[primaryAxisIndex];

        /* 副軸の座標が変わる位置が折り返し＝列数 / The wrap point gives the column count */
        for (var i = 2; i < artboards.length; i++) {
            var scannedRect = artboards[i].artboardRect;
            if (firstRect[secondaryAxisIndex] !== scannedRect[secondaryAxisIndex]) {
                gridStep[secondaryAxisIndex] = scannedRect[secondaryAxisIndex] - firstRect[secondaryAxisIndex];
                return { canInherit: true, columns: i };
            }
        }
        return { canInherit: true, columns: 0 };
    }

    /**
     * カンバスに収まるグリッドの行数・列数を求めます。
     *
     * @param {Document} doc - 対象のドキュメント。
     * @param {Array<number>} firstRect - 先頭アートボードの artboardRect。
     * @param {Array<number>} gridStep - グリッド1セルあたりの移動量 [X, Y]。
     * @param {number} primaryAxisIndex - 主軸（0=横 / 1=縦）。
     * @param {number} secondaryAxisIndex - 副軸。
     * @returns {{columns: number, rows: number}} 収まる列数と行数。
     */
    function countGridCapacity(doc, firstRect, gridStep, primaryAxisIndex, secondaryAxisIndex) {
        var canvasRect = getLargestCanvasBounds(doc);

        /* 1セル分の枠をカンバスの端と比べて、何個並ぶかを数える
           Count how many cells fit between one cell and the canvas edge */
        var gridUnitRect = [
            firstRect[0] + Math.abs(gridStep[0]),
            firstRect[1] - Math.abs(gridStep[1]),
            firstRect[0],
            firstRect[1]
        ];

        var primaryEdgeIndex =
            (primaryAxisIndex ^ +(gridStep[primaryAxisIndex] < 0)) ? primaryAxisIndex : primaryAxisIndex + 2;
        var secondaryEdgeIndex =
            (secondaryAxisIndex ^ +(gridStep[secondaryAxisIndex] < 0)) ? secondaryAxisIndex : secondaryAxisIndex + 2;

        return {
            columns: Math.abs(Math.floor(
                (canvasRect[primaryEdgeIndex] - gridUnitRect[primaryEdgeIndex]) / gridStep[primaryAxisIndex])),
            rows: Math.abs(Math.floor(
                (canvasRect[secondaryEdgeIndex] - gridUnitRect[secondaryEdgeIndex]) / gridStep[secondaryAxisIndex]))
        };
    }

    /**
     * アートボードの配置計画。
     *
     * @typedef {object} ArtboardLayout
     * @property {boolean} canInherit - 既存の並びのグリッドを引き継げる場合は true。
     * @property {Array<number>} gridStep - グリッド1セルあたりの移動量 [X, Y]。
     * @property {number} columns - グリッドの列数。
     * @property {number} primaryAxisIndex - 主軸（0=横 / 1=縦）。
     * @property {number} secondaryAxisIndex - 副軸。
     * @property {number} primarySign - 主軸方向の符号（+1 / -1）。
     * @property {number} spacing - アートボード間の間隔（pt）。
     * @property {Array<number>} firstRect - 先頭アートボードの artboardRect。
     * @property {Array<number>} referenceRect - サイズの基準にする artboardRect。
     */

    /**
     * 既存の並びを解析し、新規アートボードの配置計画を作ります。
     *
     * @param {Document} doc - 対象のドキュメント。
     * @param {number} referenceIndex - サイズの基準にするアートボードのインデックス。
     * @param {number} directionAxis - 並べる方向（0=右 / 1=下）。ダイアログで指定します。
     * @param {number} spacing - アートボードどうしの間隔（pt）。ダイアログで指定します。
     * @returns {ArtboardLayout|null} 配置計画。スペースが足りない場合は null。
     */
    function planArtboardLayout(doc, referenceIndex, directionAxis, spacing) {
        var artboards = doc.artboards;
        var firstRect = artboards[0].artboardRect;
        var primaryAxisIndex = directionAxis;
        var secondaryAxisIndex = 1 - primaryAxisIndex;

        /* グリッド1セルあたりの移動量（Y軸は上→下なので間隔を引く）
           Grid step per cell (the Y axis is top-down, so the spacing is subtracted) */
        var gridStep = [
            (firstRect[2] - firstRect[0]) + spacing,
            (firstRect[3] - firstRect[1]) - spacing
        ];

        var inheritedGrid = detectInheritedGrid(doc, gridStep, primaryAxisIndex, secondaryAxisIndex);
        var capacity = countGridCapacity(doc, firstRect, gridStep, primaryAxisIndex, secondaryAxisIndex);

        /* 既存の並びから列数が読めたときはそちらを優先
           Prefer the column count read from the existing layout */
        var columns = inheritedGrid.columns || capacity.columns;
        if (artboards.length + 1 > columns * capacity.rows) return null;

        return {
            canInherit: inheritedGrid.canInherit,
            gridStep: gridStep,
            columns: columns,
            primaryAxisIndex: primaryAxisIndex,
            secondaryAxisIndex: secondaryAxisIndex,
            primarySign: (gridStep[primaryAxisIndex] < 0) ? -1 : 1,
            spacing: spacing,
            firstRect: firstRect,
            referenceRect: artboards[referenceIndex].artboardRect
        };
    }

    /**
     * グリッド上のインデックスに対応する位置を求めます。
     *
     * @param {ArtboardLayout} artboardLayout - 配置計画。
     * @param {number} gridIndex - グリッド上のインデックス。
     * @returns {Array<number>} [左, 上] の座標。
     */
    function getArtboardGridPosition(artboardLayout, gridIndex) {
        /* 主軸は列内の位置、副軸は何行目かで決まる
           The primary axis gives the column, the secondary axis the row */
        var offset = [];
        offset[artboardLayout.primaryAxisIndex] =
            (gridIndex % artboardLayout.columns) * artboardLayout.gridStep[artboardLayout.primaryAxisIndex];
        offset[artboardLayout.secondaryAxisIndex] =
            Math.floor(gridIndex / artboardLayout.columns) * artboardLayout.gridStep[artboardLayout.secondaryAxisIndex];

        return [artboardLayout.firstRect[0] + offset[0], artboardLayout.firstRect[1] + offset[1]];
    }

    /**
     * グリッドを引き継げないときの配置位置を求めます。
     *
     * 先頭ではなく、挿入位置の直前のアートボードを基準に指定方向へ1枚分進めます。
     * グリッドのインデックスを歩数に使うと、既存の並びと軸が違う場合に
     * 枚数分だけ離れた位置へ飛んでしまうため、実位置から積み上げます。
     *
     * @param {Document} doc - 対象のドキュメント。
     * @param {ArtboardLayout} artboardLayout - 配置計画。
     * @param {number} anchorIndex - 基準にするアートボードのインデックス。
     * @returns {Array<number>} [左, 上] の座標。
     */
    function getAnchoredArtboardPosition(doc, artboardLayout, anchorIndex) {
        /* 基準アートボードの実位置から1枚分だけ進める / Step one artboard from the anchor */
        var anchorRect = doc.artboards[anchorIndex].artboardRect;
        var advance = getArtboardAxisSize(anchorRect, artboardLayout.primaryAxisIndex) + artboardLayout.spacing;
        var position = [anchorRect[0], anchorRect[1]];
        position[artboardLayout.primaryAxisIndex] += artboardLayout.primarySign * advance;
        return position;
    }

    // =========================================
    // アートボードを作成 / Create an artboard
    // =========================================

    /**
     * アートボード上のアイテムを収集します。
     *
     * doc.pageItems はグループ・複合パスの子まで再帰的に含むため、最上位
     * （親がレイヤー）のアイテムのみを対象にします。子は親と一緒に動くので、
     * ここで拾うと二重に処理されてしまいます。
     *
     * @param {Document} doc - 対象のドキュメント。
     * @param {Array<number>} artboardRect - 対象アートボードの矩形。
     * @returns {Array<PageItem>} アートボードに属するアイテム。
     */
    function getItemsAssignedToArtboard(doc, artboardRect) {
        /* 重心がアートボード内にあるものを所属とみなす
           An item belongs to the artboard when its center falls inside */
        var assignedItems = [];

        for (var i = 0; i < doc.pageItems.length; i++) {
            var item = doc.pageItems[i];
            if (item.parent.typename !== 'Layer') continue;

            if (isPointInsideRect(getRectCenter(item.geometricBounds), artboardRect)) {
                assignedItems.push(item);
            }
        }
        return assignedItems;
    }

    /**
     * アイテムと祖先レイヤーのロック／表示状態を一時的に解除します。
     *
     * ロック／非表示のアイテムは translate() が例外になります。途中で例外になると
     * そこまで動かしたアートボードだけが残って崩れるため、一時解除してから処理します。
     *
     * @param {PageItem} item - 対象のオブジェクト。
     * @returns {Array<object>} 復元用の情報リスト。
     */
    function unlockItemTemporarily(item) {
        /* 祖先レイヤーから順に解除し、戻すための記録を残す
           Clear the ancestors first, recording what to restore */
        var restoreList = [];
        var ancestorLayer = item.parent;

        while (ancestorLayer && ancestorLayer.typename === 'Layer') {
            if (ancestorLayer.locked) {
                ancestorLayer.locked = false;
                restoreList.push({ target: ancestorLayer, property: 'locked', value: true });
            }
            if (!ancestorLayer.visible) {
                ancestorLayer.visible = true;
                restoreList.push({ target: ancestorLayer, property: 'visible', value: false });
            }
            ancestorLayer = ancestorLayer.parent;
        }
        if (item.locked) {
            item.locked = false;
            restoreList.push({ target: item, property: 'locked', value: true });
        }
        if (item.hidden) {
            item.hidden = false;
            restoreList.push({ target: item, property: 'hidden', value: true });
        }
        return restoreList;
    }

    /**
     * 一時解除したロック／表示状態を元に戻します。
     *
     * @param {Array<object>} restoreList - unlockItemTemporarily の戻り値。
     * @returns {void}
     */
    function restoreLockAndVisibility(restoreList) {
        /* 解除と逆順に戻す（内側→外側）/ Restore in reverse order (inner → outer) */
        for (var i = restoreList.length - 1; i >= 0; i--) {
            restoreList[i].target[restoreList[i].property] = restoreList[i].value;
        }
    }

    /**
     * 挿入位置以降のアートボードを1枚分だけ後ろへずらし、アートワークも一緒に運びます。
     *
     * アイテムの帰属は移動前の位置でまとめて取得（スナップショット）してから動かすため、
     * 移動順による取り違えが起きません。
     *
     * @param {Document} doc - 対象のドキュメント。
     * @param {ArtboardLayout} artboardLayout - 配置計画。
     * @param {number} fromIndex - ずらし始めるインデックス。
     * @returns {void}
     */
    function relayoutExistingArtboards(doc, artboardLayout, fromIndex) {
        var artboards = doc.artboards;

        /* 先に移動計画を立てる（動かしながら判定すると帰属がずれる）
           Plan every move first: judging while moving would misassign items */
        var plannedMoves = [];
        for (var i = fromIndex; i < artboards.length; i++) {
            var currentRect = artboards[i].artboardRect;
            var targetPosition = getArtboardGridPosition(artboardLayout, i + 1);
            plannedMoves.push({
                index: i,
                dx: targetPosition[0] - currentRect[0],
                dy: targetPosition[1] - currentRect[1],
                items: getItemsAssignedToArtboard(doc, currentRect)
            });
        }

        /* アートワークを動かしてから、アートボード自体を動かす
           Move the artwork, then the artboard itself */
        for (var j = 0; j < plannedMoves.length; j++) {
            var plannedMove = plannedMoves[j];

            for (var k = 0; k < plannedMove.items.length; k++) {
                var restoreList = unlockItemTemporarily(plannedMove.items[k]);
                try {
                    plannedMove.items[k].translate(plannedMove.dx, plannedMove.dy);
                } finally {
                    restoreLockAndVisibility(restoreList);
                }
            }

            var movedRect = artboards[plannedMove.index].artboardRect;
            artboards[plannedMove.index].artboardRect = [
                movedRect[0] + plannedMove.dx, movedRect[1] + plannedMove.dy,
                movedRect[2] + plannedMove.dx, movedRect[3] + plannedMove.dy
            ];
        }
    }

    /**
     * 末尾に追加されたアートボードを、パネル上の挿入位置へ並べ替えます。
     *
     * @param {Document} doc - 対象のドキュメント。
     * @param {number} insertIndex - 挿入位置のインデックス。
     * @param {number} originalCount - 追加前のアートボード数。
     * @returns {void}
     */
    function reorderAppendedArtboard(doc, insertIndex, originalCount) {
        var artboards = doc.artboards;

        /* 追加分の枠と名前を退避（後続シフトで上書きされる前に）
           Save the appended rect and name before the shift clobbers them */
        var appendedRect = artboards[originalCount].artboardRect;
        var appendedName = artboards[originalCount].name;

        /* 後続を1枚分だけ後ろへ（高位から処理して上書き衝突を回避）
           Shift trailing artboards back by one (high → low to avoid clobbering) */
        for (var i = originalCount - 1; i >= insertIndex; i--) {
            artboards[i + 1].artboardRect = artboards[i].artboardRect;
            artboards[i + 1].name = artboards[i].name;
        }

        artboards[insertIndex].artboardRect = appendedRect;
        artboards[insertIndex].name = appendedName;
    }

    /**
     * 選択オブジェクトを指定量だけ移動します。複製が指定された場合は複製側を移動します。
     *
     * @param {Array<PageItem>} items - 対象のオブジェクト。
     * @param {number} dx - X方向の移動量。
     * @param {number} dy - Y方向の移動量。
     * @param {boolean} useDuplicate - true なら元のオブジェクトを残して複製する。
     * @returns {Array<PageItem>} 移動先に置かれたオブジェクト。
     */
    function moveOrDuplicateItems(items, dx, dy, useDuplicate) {
        /* 複製時は元を残し、複製した側だけを動かす
           When duplicating, keep the originals and move only the copies */
        var placedItems = [];

        for (var i = 0; i < items.length; i++) {
            var targetItem = useDuplicate ? items[i].duplicate() : items[i];
            targetItem.translate(dx, dy);
            placedItems.push(targetItem);
        }
        return placedItems;
    }

    /**
     * 選択オブジェクトから新規アートボードを作成します。
     *
     * 新規アートボードのサイズはアクティブアートボードと同じで、選択オブジェクトは
     * 元のアートボード内での相対位置を保ったまま移動（または複製）されます。
     *
     * @param {Document} doc - 対象のドキュメント。
     * @param {boolean} useDuplicate - true なら元のオブジェクトを残して複製する。
     * @param {string} artboardName - 新規アートボードの名前。
     * @param {number} directionAxis - 並べる方向（0=右 / 1=下）。
     * @param {number} spacingPt - アートボードどうしの間隔（pt）。
     * @returns {void}
     */
    function createArtboardFromSelection(doc, useDuplicate, artboardName, directionAxis, spacingPt) {
        var selectedItems = snapshotSelection(doc);
        if (selectedItems.length === 0) {
            alert(L(LABELS.alert.noSelection));
            return;
        }

        /* 上限を超えるなら何も変更せずに終える / Bail out before touching anything */
        var artboards = doc.artboards;
        var originalArtboardCount = artboards.length;
        var artboardLimit = (parseFloat(app.version) >= 22) ? 1000 : 100;

        if (originalArtboardCount + 1 > artboardLimit) {
            alert(L(LABELS.alert.artboardLimit));
            return;
        }

        var activeArtboardIndex = artboards.getActiveArtboardIndex();
        var insertArtboardIndex = ARTBOARD_INSERT_AFTER_CURRENT
            ? (activeArtboardIndex + 1)
            : originalArtboardCount;

        var artboardLayout = planArtboardLayout(doc, activeArtboardIndex, directionAxis, spacingPt);
        if (artboardLayout === null) {
            alert(L(LABELS.alert.noSpace));
            return;
        }

        /* 基準になるアートボードと選択位置は、ずらす前に記録しておく
           Record the source artboard and the selection's position before anything shifts */
        var sourceArtboardIndex = findArtboardIndexForItems(doc, selectedItems, activeArtboardIndex);
        var sourceRect = artboards[sourceArtboardIndex].artboardRect;
        var boundsBeforeShift = getUnionBounds(selectedItems);

        /* 既存アートボードを後ろへずらして挿入スペースを空ける。
           グリッドを引き継げないときは既存の並びを動かさない（別軸のグリッドへ
           流し込むと並びが崩れるため）
           Free the insert space. When the grid can't be inherited, leave the existing
           artboards alone — re-flowing onto the other axis would break the arrangement */
        if (artboardLayout.canInherit) {
            relayoutExistingArtboards(doc, artboardLayout, insertArtboardIndex);
        }

        /* 選択オブジェクトが実際にずれた量を測る。
           アートボード内に重心があるものだけがずれるため、ずれ幅は0のこともある。
           この実測値を引くことで、重心がアートボード外にあるオブジェクトでも
           元のアートボードに対する相対位置が正しく保たれる
           Measure how far the selection actually moved: only art whose center sits inside
           the artboard is shifted, so this can be zero. Subtracting the measured amount
           keeps the relative position correct even for art centered outside its artboard */
        var boundsAfterShift = getUnionBounds(selectedItems);
        var shiftX = boundsAfterShift[0] - boundsBeforeShift[0];
        var shiftY = boundsAfterShift[1] - boundsBeforeShift[1];

        var newPosition = artboardLayout.canInherit
            ? getArtboardGridPosition(artboardLayout, insertArtboardIndex)
            : getAnchoredArtboardPosition(doc, artboardLayout, insertArtboardIndex - 1);

        /* アクティブアートボードと同じサイズで作る / Match the active artboard's size */
        var referenceRect = artboardLayout.referenceRect;
        artboards.add([
            newPosition[0],
            newPosition[1],
            newPosition[0] + (referenceRect[2] - referenceRect[0]),
            newPosition[1] + (referenceRect[3] - referenceRect[1])
        ]);

        /* パネル上の順序を挿入位置へ並べ替え / Reorder in the panel */
        reorderAppendedArtboard(doc, insertArtboardIndex, originalArtboardCount);
        artboards[insertArtboardIndex].name = artboardName;

        /* 相対位置を保ったまま新規アートボードへ（すでにずれた分は差し引く）
           Carry the relative position over, minus the shift already applied */
        var placedItems = moveOrDuplicateItems(
            selectedItems,
            newPosition[0] - sourceRect[0] - shiftX,
            newPosition[1] - sourceRect[1] - shiftY,
            useDuplicate);

        artboards.setActiveArtboardIndex(insertArtboardIndex);
        doc.selection = placedItems;
        app.redraw();
    }

    // =========================================
    // ドキュメントの走査 / Walking a document
    // =========================================

    /**
     * 進捗を示すパレットを開きます。
     *
     * ドキュメントの複製は保存・削除・再オープンと時間のかかる処理が続くため、
     * 無反応に見えないよう進捗を出します。
     *
     * @returns {{setMessage: function, close: function}} 進捗パレットの操作口。
     */
    function openProgressPalette() {
        var palette = new Window("palette", L(LABELS.progress.title));
        palette.orientation = "column";
        palette.alignChildren = ["fill", "center"];
        palette.margins = 20;

        var messageText = palette.add("statictext", undefined, "");
        messageText.preferredSize.width = 280;
        palette.show();

        return {
            /**
             * 表示するメッセージを差し替えます。
             *
             * @param {object} labelEntry - 表示する文言のラベル定義。
             * @returns {void}
             */
            setMessage: function (labelEntry) {
                messageText.text = L(labelEntry);
                palette.update();
            },

            /**
             * パレットを閉じます。
             *
             * @returns {void}
             */
            close: function () {
                palette.close();
            }
        };
    }

    /**
     * ファイル名から拡張子を除いた部分を取り出します。
     *
     * @param {string} fileName - ファイル名（拡張子込み）。
     * @returns {string} 拡張子を除いた部分。
     */
    function getFileBaseName(fileName) {
        /* 「my.logo.ai」のような名前でも最後のドットだけで切る
           Split at the last dot only, even for names like "my.logo.ai" */
        var dotIndex = fileName.lastIndexOf('.');
        return (dotIndex > 0) ? fileName.substring(0, dotIndex) : fileName;
    }

    /**
     * ファイル名から拡張子を取り出します。
     *
     * @param {string} fileName - ファイル名（拡張子込み）。
     * @returns {string} ドットを含む拡張子。拡張子がない場合は空文字。
     */
    function getFileExtension(fileName) {
        var dotIndex = fileName.lastIndexOf('.');
        return (dotIndex > 0) ? fileName.substring(dotIndex) : '';
    }

    /**
     * すべてのレイヤーのロック・非表示を一時的に解除します。
     *
     * 反転選択はロック・非表示のレイヤーを拾わないため、削除の前に解除します。
     * レイヤーの表示状態は成果物に残るので、あとで元に戻せるよう記録します。
     *
     * @param {Layers} layers - 対象のレイヤーコレクション。
     * @param {Array<object>} restoreList - 復元用の情報を追加する配列。
     * @returns {void}
     */
    function unlockLayersTemporarily(layers, restoreList) {
        /* サブレイヤーまで再帰的に解除する / Clear sub-layers recursively too */
        for (var i = 0; i < layers.length; i++) {
            restoreList.push({ layer: layers[i], locked: layers[i].locked, visible: layers[i].visible });
            layers[i].locked = false;
            layers[i].visible = true;
            unlockLayersTemporarily(layers[i].layers, restoreList);
        }
    }

    /**
     * 一時解除したレイヤーのロック・非表示を元に戻します。
     *
     * @param {Array<object>} restoreList - unlockLayersTemporarily で記録した情報。
     * @returns {void}
     */
    function restoreLayerStates(restoreList) {
        /* 解除と逆順に戻す（内側→外側）/ Restore in reverse order (inner → outer) */
        for (var i = restoreList.length - 1; i >= 0; i--) {
            restoreList[i].layer.locked = restoreList[i].locked;
            restoreList[i].layer.visible = restoreList[i].visible;
        }
    }

    /**
     * 記録済みのレイヤー状態から、指定レイヤーの分を探します。
     *
     * @param {Array<object>} layerStates - unlockLayersTemporarily で記録した情報。
     * @param {Layer} layer - 探すレイヤー。
     * @returns {object|null} 見つかった記録。見つからない場合は null。
     */
    function findLayerState(layerStates, layer) {
        /* レイヤー数はたかが知れているので線形探索で足りる
           Layer counts stay small, so a linear scan is enough */
        for (var i = 0; i < layerStates.length; i++) {
            if (layerStates[i].layer === layer) return layerStates[i];
        }
        return null;
    }

    /**
     * 反転削除の前に、残したいオブジェクトだけをロック／非表示のままにします。
     *
     * 反転選択はロック・非表示のオブジェクトを拾わないため、ロック／非表示のままに
     * したものは削除されずに残ります。逆に残さないものはここで解除しておきます。
     * レイヤー単位のロック・非表示も、そのオブジェクト自身の状態として扱います。
     *
     * 対象は最上位（親がレイヤー）のオブジェクトだけです。グループ内のオブジェクトは
     * ロック／非表示にしても、親グループが反転で選択されればまとめて削除されるため、
     * 残す保証ができません。状態には触れず、親の扱いに従わせます。
     *
     * @param {Document} targetDoc - 対象のドキュメント。
     * @param {Array<object>} layerStates - 解除前に記録したレイヤー状態。
     * @param {boolean} includeLocked - ロックされたオブジェクトを残す場合は true。
     * @param {boolean} includeHidden - 非表示オブジェクトを残す場合は true。
     * @returns {Array<object>} 残す印を付けたオブジェクトと、その元の状態。
     */
    function protectItemsFromDelete(targetDoc, layerStates, includeLocked, includeHidden) {
        /* コレクションはループ外に退避する（毎回DOMを経由するため）
           Hoist the collection: each access crosses the DOM bridge */
        var pageItems = targetDoc.pageItems;
        var protectedEntries = [];

        for (var i = 0; i < pageItems.length; i++) {
            var item = pageItems[i];

            /* グループ・複合パスの中身は親の扱いに任せる
               Leave the contents of groups and compound paths to their parent */
            if (item.parent.typename !== 'Layer') continue;

            var layerState = findLayerState(layerStates, item.layer);
            var ownLocked = item.locked;
            var ownHidden = item.hidden;
            var wasLocked = ownLocked || (layerState !== null && layerState.locked);
            var wasHidden = ownHidden || (layerState !== null && !layerState.visible);

            var keepLocked = includeLocked && wasLocked;
            var keepHidden = includeHidden && wasHidden;

            if (keepLocked || keepHidden) {
                protectedEntries.push({ item: item, locked: ownLocked, hidden: ownHidden });
            }

            /* 変化するときだけ書き込む / Write only when the value actually changes */
            if (ownLocked !== keepLocked) item.locked = keepLocked;
            if (ownHidden !== keepHidden) item.hidden = keepHidden;
        }
        return protectedEntries;
    }

    /**
     * 残したオブジェクトを、アートボードの内外で仕分けます。
     *
     * アートボード内のものは自身のロック・非表示状態を元に戻し、
     * 外にあるものは削除します。
     *
     * @param {Array<object>} protectedEntries - protectItemsFromDelete の戻り値。
     * @param {Array<number>} artboardRect - 残すアートボードの矩形。
     * @returns {void}
     */
    function finalizeProtectedItems(protectedEntries, artboardRect) {
        for (var i = 0; i < protectedEntries.length; i++) {
            var entry = protectedEntries[i];

            /* 親グループごと削除されている場合があるため保護
               The parent group may already be gone, so guard the access */
            try {
                if (isPointInsideRect(getRectCenter(entry.item.geometricBounds), artboardRect)) {
                    entry.item.locked = entry.locked;
                    entry.item.hidden = entry.hidden;
                } else {
                    entry.item.locked = false;
                    entry.item.hidden = false;
                    entry.item.remove();
                }
            } catch (e) {}
        }
    }

    // =========================================
    // ドキュメントを作成 / Create a document
    // =========================================

    /**
     * 指定した1枚を除いて、すべてのアートボードを削除します。
     *
     * @param {Document} targetDoc - 対象のドキュメント。
     * @param {number} keepIndex - 残すアートボードのインデックス。
     * @returns {void}
     */
    function keepOnlyArtboard(targetDoc, keepIndex) {
        /* 降順に処理するので、削除しても未処理側のインデックスはずれない
           A descending loop keeps the pending indexes valid */
        var artboards = targetDoc.artboards;

        for (var i = artboards.length - 1; i >= 0; i--) {
            if (i !== keepIndex) artboards.remove(i);
        }
    }

    /**
     * 選択オブジェクトから新規ドキュメントを作成します。
     *
     * 別名保存で複製ドキュメントを作り、複製側で選択範囲を反転して削除し、
     * 現在のアートボード以外を削除します。スウォッチ・シンボル・ドキュメント設定は
     * そのまま引き継がれます。
     *
     * `saveAs()` は開いているドキュメント自体を保存先に紐づけ直すため、複製側に
     * 選択状態がそのまま残ります。これを使うと「選択以外」をIllustrator自身の
     * 反転選択で求められるので、オブジェクトの突き合わせが不要になります。
     * 元ファイルはディスク上に保存時のまま残るので、最後に開き直します。
     *
     * @param {Document} doc - 対象のドキュメント。
     * @param {boolean} includeLocked - ロックされたオブジェクトも残す場合は true。
     * @param {boolean} includeHidden - 非表示オブジェクトも残す場合は true。
     * @param {string} fileBaseName - 複製ドキュメントのファイル名（拡張子なし）。
     * @returns {void}
     */
    function createDocumentFromSelection(doc, includeLocked, includeHidden, fileBaseName) {
        if (snapshotSelection(doc).length === 0) {
            alert(L(LABELS.alert.noSelection));
            return;
        }

        /* 未保存・変更ありだと、開き直したときに変更が失われる
           Unsaved changes would be lost when the source document is reopened */
        var originalFile = doc.saved ? doc.fullName : null;
        if (originalFile === null || !originalFile.exists) {
            alert(L(LABELS.alert.needsSave));
            return;
        }

        /* 保存先は元ファイルと同じ場所、拡張子も元のまま
           The duplicate sits next to the source, keeping its extension */
        var duplicateFile = new File(
            originalFile.parent.fsName + "/" + fileBaseName + getFileExtension(originalFile.name));

        /* 元ファイルに上書きすると、開き直す先が複製になってしまう
           Overwriting the source would leave nothing to reopen */
        if (duplicateFile.fsName === originalFile.fsName) {
            alert(L(LABELS.alert.sameAsSource));
            return;
        }
        if (duplicateFile.exists && !confirm(L(LABELS.alert.overwrite) + "\n" + duplicateFile.name)) {
            return;
        }

        var activeArtboardIndex = doc.artboards.getActiveArtboardIndex();

        /* 保存・削除・再オープンと時間がかかるので、進捗を出しておく
           Saving, deleting and reopening all take time, so show the progress */
        var progressPalette = openProgressPalette();

        /* 何があってもパレットは閉じ、元ドキュメントは必ず開き直す
           Always close the palette and reopen the source document */
        var layerRestoreList = [];
        try {
            progressPalette.setMessage(LABELS.progress.duplicating);

            /* 別名保存すると、このドキュメント自体が複製ファイルに紐づく（選択状態はそのまま）
               Saving as rebinds this very document to the duplicate, selection intact */
            doc.saveAs(duplicateFile);

            progressPalette.setMessage(LABELS.progress.deleting);

            /* オブジェクト単位で状態を扱えるよう、レイヤーはいったんすべて解除する
               Clear every layer first so each item can be addressed individually */
            unlockLayersTemporarily(doc.layers, layerRestoreList);

            /* 残すものだけロック／非表示のままにする（反転が拾わない＝削除されない）
               Keep only the survivors locked or hidden: Inverse skips them */
            var protectedEntries = protectItemsFromDelete(doc, layerRestoreList, includeLocked, includeHidden);

            /* メニューコマンドはアクティブドキュメントに働くので、対象を明示しておく
               Menu commands act on the active document, so make the target explicit */
            app.activeDocument = doc;

            /* 選択範囲を反転して削除 / Invert the selection and delete it */
            app.executeMenuCommand('Inverse menu item');
            app.executeMenuCommand('clear');

            keepOnlyArtboard(doc, activeArtboardIndex);

            /* 残したもののうち、アートボード外にあるものは削除する
               Drop the survivors that ended up outside the artboard */
            finalizeProtectedItems(protectedEntries, doc.artboards[0].artboardRect);

            /* レイヤーの表示・ロック状態は元ドキュメントのまま残す
               Leave the layers' visibility and lock state as they were */
            restoreLayerStates(layerRestoreList);
        } finally {
            /* saveAs より前で失敗した場合は複製に紐づいていないので開き直さない
               Nothing to reopen when the failure happened before saveAs */
            if (doc.fullName.fsName === duplicateFile.fsName) {
                progressPalette.setMessage(LABELS.progress.reopening);
                app.open(originalFile);
            }
            progressPalette.close();
        }

        app.activeDocument = doc;
        app.redraw();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログで作成対象を選び、対応する処理を実行します。
     *
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(L(LABELS.alert.noDocument));
            return;
        }

        var doc = app.activeDocument;

        /* キャンセルされたら何もしない / Do nothing when cancelled */
        var createOptions = showCreateTargetDialog(doc);
        if (createOptions === null) return;

        /* 選ばれた作成対象へ振り分ける / Dispatch to the chosen target */
        if (createOptions.createTarget === "artboard") {
            createArtboardFromSelection(doc, createOptions.useDuplicate, createOptions.createName,
                createOptions.directionAxis, createOptions.spacingPt);
        } else if (createOptions.createTarget === "document") {
            createDocumentFromSelection(doc, createOptions.includeLocked, createOptions.includeHidden,
                createOptions.createName);
        } else {
            createLayerFromSelection(doc, createOptions.useDuplicate, createOptions.createName);
        }
    }

    /* 想定外のエラーはここで受け止める。個々の処理には try を置かない方針なので、
       捕まえ損ねるとExtendScriptの生のエラーダイアログが出てしまう
       Catch the unexpected here: the individual steps deliberately avoid try blocks,
       so anything uncaught would surface as a raw ExtendScript error dialog */
    try {
        main();
    } catch (err) {
        alert(L(LABELS.alert.unexpected) + err);
    }

})();
