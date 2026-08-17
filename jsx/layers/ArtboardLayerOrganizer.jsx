#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### 概要

- ドキュメント内のオブジェクトをアートボード単位で振り分け、「番号_アートボード名」のレイヤーに整理します。
- 所属アートボードは各オブジェクトの重心位置で判定します。
- 詳細な機能・オプションはREADMEを参照してください。

### Overview

- Distributes objects in the document by artboard and organizes them into "number_artboard name" layers.
- Each object is assigned to an artboard by its centroid.
- See the README for the full feature and option list.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ArtboardLayerOrganizer";       /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-04";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-17";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArtboardLayerOrganizer.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ArtboardLayerOrganizer.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nadb8b8ba49fe"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* システム管理レイヤー名（空でも削除しない） / System-managed layer names (never deleted) */
    var GUIDE_LAYER_NAME = "_guide";
    var PASTEBOARD_LAYER_NAME = "_pasteboard";

    /* 区切り文字ドロップダウンの選択位置と実際の文字 / Separator dropdown indexes and their characters */
    var SEPARATOR_UNDERSCORE = 0;
    var SEPARATOR_HYPHEN     = 1;
    var SEPARATOR_SPACE      = 2;
    var SEPARATOR_NONE       = 3;
    var SEPARATOR_CHARACTERS = ["_", "-", " ", ""];

    /* ダイアログの初期値。必要に応じて編集 / Dialog defaults; edit as needed */
    var DEFAULT_OPTIONS = {
        removeEmptyLayers: true,
        includeArtboardNumber: true,
        includeArtboardName: true,
        useSeparator: true,
        layerNameSeparatorIndex: SEPARATOR_UNDERSCORE,
        ignoreLockedLayers: true,
        ignoreLockedObjects: true,
        ignoreHiddenLayers: true,
        ignoreHiddenObjects: true,
        excludedLayerNames: "bg"    // , または 、 区切り / Comma-separated
    };

    // =========================================
    // レイアウト設定 / Layout Settings
    // =========================================

    /* ダイアログ外周の余白 / Dialog margins */
    var DIALOG_MARGINS = [15, 20, 15, 15];

    /* パネル共通の余白と行間 / Common panel margins and spacing */
    var PANEL_MARGINS = [15, 20, 15, 10];
    var PANEL_SPACING = 8;

    /* 「ロック」「非表示」サブパネル間の間隔 / Gap between the locked / hidden sub-panels */
    var EXCLUSION_SUBPANEL_SPACING = 15;

    /* レイヤー名プレビューの最大文字数（ダイアログ幅を広げないための上限） / Preview length cap, keeps the dialog from widening */
    var LAYER_NAME_PREVIEW_MAX_LENGTH = 16;

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

    var uiLanguage = detectUILanguage();

    /* ラベル内の {slash} などを言語別に置き換えるための記号表 / Symbol table for {slash}-style placeholders */
    var LOCALIZED_SYMBOLS = {
        slash:      { ja: "／", en: "/" },
        colon:      { ja: "：", en: ":" },
        comma:      { ja: "、", en: ", " },
        openParen:  { ja: "（", en: "(" },
        closeParen: { ja: "）", en: ")" }
    };

    var LABELS = {
        dialog: {
            title: { ja: "アートボードごとにレイヤーを整理", en: "Organize Layers by Artboard" }
        },
        panel: {
            targetArtboards: { ja: "対象のアートボード", en: "Target Artboards" },
            layerName: { ja: "レイヤー名の構成", en: "Layer Name Format" },
            exclusion: { ja: "対象外にする", en: "Exclude" },
            locked: { ja: "ロック中", en: "Locked" },
            hidden: { ja: "非表示", en: "Hidden" },
            postProcess: { ja: "整理後の処理", en: "After Organizing" }
        },
        radio: {
            currentArtboardOnly: { ja: "現在のアートボード", en: "Current artboard" },
            allArtboards: { ja: "すべて", en: "All" }
        },
        checkbox: {
            includeArtboardNumber: { ja: "番号を含める", en: "Include number" },
            includeArtboardName: { ja: "名前を含める", en: "Include name" },
            useSeparator: { ja: "区切り文字", en: "Separator" },
            excludeLayer: { ja: "レイヤー", en: "Layer" },
            excludeObject: { ja: "オブジェクト", en: "Object" },
            removeEmpty: { ja: "空のレイヤー{slash}サブレイヤーを削除", en: "Remove empty layers{slash}sub-layers" }
        },
        dropdown: {
            separatorUnderscore: { ja: "アンダースコア (_)", en: "Underscore (_)" },
            separatorHyphen: { ja: "ハイフン (-)", en: "Hyphen (-)" },
            separatorSpace: { ja: "半角スペース", en: "Space" },
            separatorNone: { ja: "なし", en: "None" }
        },
        fieldLabel: {
            specifiedLayers: { ja: "レイヤー名で指定", en: "Specify by name" },
            layerNamePreview: { ja: "例", en: "Example" }
        },
        tooltip: {
            targetArtboards: {
                ja: "アートボードが1つのときは「現在のアートボード」に固定されます。「現在のアートボード」では、どのアートボードにも乗っていないオブジェクトは移動しません",
                en: "Locked to \"Current artboard\" when the document has just one artboard. In that mode, objects outside every artboard are left where they are"
            },
            includeArtboardNumber: {
                ja: "レイヤー名の先頭にアートボードの通し番号（1, 2, 3…）を付けます",
                en: "Prefixes the layer name with the artboard number (1, 2, 3 ...)"
            },
            includeArtboardName: {
                ja: "レイヤー名にアートボード名を含めます。名前が空のときは「アートボード」を使います",
                en: "Includes the artboard name; falls back to \"Artboard\" when it is empty"
            },
            useSeparator: {
                ja: "アートボード番号とアートボード名の間に入れる文字",
                en: "Character inserted between the artboard number and the artboard name"
            },
            exclusionPanel: {
                ja: "ガイドはここでの指定に関係なく、常に _guide レイヤーに集約されます",
                en: "Guides are always gathered into the _guide layer regardless of these settings"
            },
            lockedExclusion: {
                ja: "ロックされたレイヤー{slash}オブジェクトを整理対象から除外します",
                en: "Leaves locked layers{slash}objects out of the organizing"
            },
            hiddenExclusion: {
                ja: "非表示のレイヤー{slash}オブジェクトを整理対象から除外します",
                en: "Leaves hidden layers{slash}objects out of the organizing"
            },
            specifiedLayers: {
                ja: "カンマ または 「、」 区切りでレイヤー名を指定（例: bg, temp）",
                en: "Layer names separated by comma (e.g. bg, temp)"
            },
            removeEmpty: {
                ja: "整理後に空になったレイヤー{slash}サブレイヤーを削除します（_guide と _pasteboard は削除しません）",
                en: "Removes layers{slash}sub-layers left empty after organizing (_guide and _pasteboard are kept)"
            }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        fallbackName: {
            artboard: { ja: "アートボード", en: "Artboard" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            moveFailed: {
                ja: "{count}件のオブジェクトを移動できませんでした。\nロックを解除できないオブジェクトが含まれている可能性があります。",
                en: "{count} object(s) could not be moved.\nSome of them may not allow unlocking."
            }
        }
    };

    /**
     * 言語別の記号を返す（日本語＝全角、英語＝半角）
     * @param {string} symbolName - 記号名（slash / colon / comma / openParen / closeParen）
     * @returns {string} 表示言語に合わせた記号。未定義の記号名なら空文字
     */
    function getLocalizedSymbol(symbolName) {
        var symbolVariants = LOCALIZED_SYMBOLS[symbolName];
        if (!symbolVariants) return "";
        return symbolVariants[uiLanguage] || symbolVariants.en;
    }

    /**
     * {slash} などのプレースホルダを言語別の記号に展開する
     * 記号表に無いプレースホルダ（{count} など）はそのまま残す
     * @param {string} labelText - 展開前のテキスト
     * @returns {string} 展開後のテキスト
     */
    function expandSymbolPlaceholders(labelText) {
        return labelText.replace(/\{(\w+)\}/g, function (matchedText, symbolName) {
            var symbolText = getLocalizedSymbol(symbolName);
            return (symbolText === "") ? matchedText : symbolText;
        });
    }

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "panel.layerName" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
            if (labelNode === undefined || labelNode === null) return labelPath;
        }
        var labelText = labelNode[uiLanguage];
        if (typeof labelText !== "string") labelText = labelNode.en;
        if (typeof labelText !== "string") return labelPath;
        return expandSymbolPlaceholders(labelText);
    }

    // =========================================
    // 前提チェック / Preconditions
    // =========================================

    if (app.documents.length === 0) {
        alert(getLabel("alert.noDocument"));
        return;
    }

    var activeDocument = app.activeDocument;
    var documentArtboards = activeDocument.artboards;
    var guideLayer = null; // _guide レイヤー参照（必要時に取得・作成） / _guide layer reference, resolved on demand

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * パネル共通の見た目をまとめて設定する
     * @param {Panel} targetPanel - 対象パネル
     * @param {Array<string>} [panelAlignment] - パネル自身の配置（省略時は横も縦も fill）
     * @returns {void}
     */
    function applyPanelLayout(targetPanel, panelAlignment) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = panelAlignment || ["fill", "top"];
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = PANEL_SPACING;
    }

    /**
     * ダイアログを構築・表示し、選択結果の処理設定を返す
     * @returns {Object} 処理設定。キャンセル時は null
     */
    function showOptionsDialog() {
        var optionsDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        optionsDialog.orientation = "column";
        optionsDialog.alignChildren = ["fill", "top"];
        optionsDialog.margins = DIALOG_MARGINS;

        var targetControls = buildTargetArtboardPanel(optionsDialog);
        var layerNameControls = buildLayerNamePanel(optionsDialog);
        var exclusionControls = buildExclusionPanel(optionsDialog);
        var postProcessControls = buildPostProcessPanel(optionsDialog);
        buildDialogButtonRow(optionsDialog);

        if (optionsDialog.show() !== 1) {
            return null;
        }

        return {
            removeEmptyLayers: postProcessControls.removeEmptyLayersCheckbox.value,
            currentArtboardOnly: targetControls.currentArtboardOnlyRadio.value,
            includeArtboardNumber: layerNameControls.artboardNumberCheckbox.value,
            includeArtboardName: layerNameControls.artboardNameCheckbox.value,
            layerNameSeparatorIndex: layerNameControls.getSeparatorIndex(),
            ignoreLockedLayers: exclusionControls.lockedLayerCheckbox.value,
            ignoreLockedObjects: exclusionControls.lockedObjectCheckbox.value,
            ignoreHiddenLayers: exclusionControls.hiddenLayerCheckbox.value,
            ignoreHiddenObjects: exclusionControls.hiddenObjectCheckbox.value,
            excludedLayerNames: parseExcludedLayerNames(exclusionControls.excludedNamesField.text)
        };
    }

    /**
     * 対象アートボードパネル（現在のみ／すべて）を構築する
     * @param {Window} parentDialog - 追加先のダイアログ
     * @returns {Object} ラジオボタン参照をまとめたオブジェクト
     */
    function buildTargetArtboardPanel(parentDialog) {
        var targetArtboardPanel = parentDialog.add("panel", undefined, getLabel("panel.targetArtboards"));
        applyPanelLayout(targetArtboardPanel);
        targetArtboardPanel.helpTip = getLabel("tooltip.targetArtboards");

        var artboardScopeRow = targetArtboardPanel.add("group");
        artboardScopeRow.orientation = "row";
        artboardScopeRow.alignment = ["left", "top"];
        artboardScopeRow.alignChildren = ["left", "center"];

        var currentArtboardOnlyRadio = artboardScopeRow.add("radiobutton", undefined, getLabel("radio.currentArtboardOnly"));
        var allArtboardsRadio = artboardScopeRow.add("radiobutton", undefined, getLabel("radio.allArtboards"));
        currentArtboardOnlyRadio.helpTip = targetArtboardPanel.helpTip;
        allArtboardsRadio.helpTip = targetArtboardPanel.helpTip;

        if (documentArtboards.length <= 1) {
            currentArtboardOnlyRadio.value = true;
            allArtboardsRadio.enabled = false;
        } else {
            allArtboardsRadio.value = true;
        }

        return { currentArtboardOnlyRadio: currentArtboardOnlyRadio };
    }

    /**
     * レイヤー名パネル（番号・名前・区切り文字）を構築する
     * @param {Window} parentDialog - 追加先のダイアログ
     * @returns {Object} チェックボックス参照と区切り文字インデックス取得関数
     */
    function buildLayerNamePanel(parentDialog) {
        var layerNamePanel = parentDialog.add("panel", undefined, getLabel("panel.layerName"));
        applyPanelLayout(layerNamePanel);

        var artboardNumberCheckbox = layerNamePanel.add("checkbox", undefined, getLabel("checkbox.includeArtboardNumber"));
        artboardNumberCheckbox.value = DEFAULT_OPTIONS.includeArtboardNumber;
        artboardNumberCheckbox.helpTip = getLabel("tooltip.includeArtboardNumber");

        var separatorRow = layerNamePanel.add("group");
        separatorRow.orientation = "row";
        separatorRow.alignment = ["left", "top"];
        separatorRow.alignChildren = ["left", "center"];
        separatorRow.helpTip = getLabel("tooltip.useSeparator");
        var useSeparatorCheckbox = separatorRow.add("checkbox", undefined, getLabel("checkbox.useSeparator"));
        useSeparatorCheckbox.value = DEFAULT_OPTIONS.useSeparator;
        useSeparatorCheckbox.helpTip = separatorRow.helpTip;
        var separatorDropdown = separatorRow.add("dropdownlist", undefined, [
            getLabel("dropdown.separatorUnderscore"),
            getLabel("dropdown.separatorHyphen"),
            getLabel("dropdown.separatorSpace"),
            getLabel("dropdown.separatorNone")
        ]);
        separatorDropdown.selection = DEFAULT_OPTIONS.layerNameSeparatorIndex;
        separatorDropdown.helpTip = separatorRow.helpTip;

        var artboardNameCheckbox = layerNamePanel.add("checkbox", undefined, getLabel("checkbox.includeArtboardName"));
        artboardNameCheckbox.value = DEFAULT_OPTIONS.includeArtboardName;
        artboardNameCheckbox.helpTip = getLabel("tooltip.includeArtboardName");

        var layerNamePreviewText = layerNamePanel.add("statictext", undefined, "");
        layerNamePreviewText.alignment = ["fill", "top"];

        /**
         * 現在の設定で選択されている区切り文字のインデックスを返す
         * @returns {number} 区切り文字のインデックス
         */
        function getSeparatorIndex() {
            if (useSeparatorCheckbox.value && separatorDropdown.selection) {
                return separatorDropdown.selection.index;
            }
            return SEPARATOR_NONE;
        }

        /**
         * 1番目のアートボードを例にしたレイヤー名のプレビュー文字列を作る
         * 長いアートボード名でダイアログが広がらないよう、一定の長さで省略する
         * @returns {string} プレビュー用のテキスト
         */
        function buildLayerNamePreview() {
            var sampleLayerName = getArtboardLayerName(0, {
                includeArtboardNumber: artboardNumberCheckbox.value,
                includeArtboardName: artboardNameCheckbox.value,
                layerNameSeparatorIndex: getSeparatorIndex()
            });
            if (sampleLayerName.length > LAYER_NAME_PREVIEW_MAX_LENGTH) {
                sampleLayerName = sampleLayerName.substring(0, LAYER_NAME_PREVIEW_MAX_LENGTH) + "…";
            }
            return getLabel("fieldLabel.layerNamePreview") + getLocalizedSymbol("colon") + sampleLayerName;
        }

        /**
         * 番号と名前の両方が外れないようにし、区切り文字の有効／無効とプレビューを更新する
         * 区切り文字は番号と名前の両方を出すときだけ意味を持つ
         * @param {Checkbox} keepOnCheckbox - 両方外れたときに戻すチェックボックス
         * @returns {void}
         */
        function updateLayerNameControls(keepOnCheckbox) {
            if (!artboardNumberCheckbox.value && !artboardNameCheckbox.value) {
                keepOnCheckbox.value = true;
            }
            var separatorApplies = artboardNumberCheckbox.value && artboardNameCheckbox.value;
            separatorRow.enabled = separatorApplies;
            separatorDropdown.enabled = separatorApplies && useSeparatorCheckbox.value;
            layerNamePreviewText.text = buildLayerNamePreview();
        }

        artboardNumberCheckbox.onClick = function () { updateLayerNameControls(artboardNameCheckbox); };
        artboardNameCheckbox.onClick = function () { updateLayerNameControls(artboardNumberCheckbox); };
        useSeparatorCheckbox.onClick = function () { updateLayerNameControls(artboardNumberCheckbox); };
        separatorDropdown.onChange = function () { updateLayerNameControls(artboardNumberCheckbox); };
        updateLayerNameControls(artboardNumberCheckbox);

        return {
            artboardNumberCheckbox: artboardNumberCheckbox,
            artboardNameCheckbox: artboardNameCheckbox,
            getSeparatorIndex: getSeparatorIndex
        };
    }

    /**
     * 対象外パネル（ロック／非表示／指定レイヤー名）を構築する
     * @param {Window} parentDialog - 追加先のダイアログ
     * @returns {Object} チェックボックスと入力欄の参照をまとめたオブジェクト
     */
    function buildExclusionPanel(parentDialog) {
        var exclusionPanel = parentDialog.add("panel", undefined, getLabel("panel.exclusion"));
        applyPanelLayout(exclusionPanel);
        exclusionPanel.helpTip = getLabel("tooltip.exclusionPanel");

        var lockHiddenRow = exclusionPanel.add("group");
        lockHiddenRow.orientation = "row";
        lockHiddenRow.alignment = ["left", "top"];
        lockHiddenRow.alignChildren = ["left", "fill"];
        lockHiddenRow.spacing = EXCLUSION_SUBPANEL_SPACING;

        var lockedControls = buildExclusionSubPanel(lockHiddenRow, "panel.locked", "tooltip.lockedExclusion", DEFAULT_OPTIONS.ignoreLockedLayers, DEFAULT_OPTIONS.ignoreLockedObjects);
        var hiddenControls = buildExclusionSubPanel(lockHiddenRow, "panel.hidden", "tooltip.hiddenExclusion", DEFAULT_OPTIONS.ignoreHiddenLayers, DEFAULT_OPTIONS.ignoreHiddenObjects);

        var excludedNamesRow = exclusionPanel.add("group");
        excludedNamesRow.orientation = "row";
        excludedNamesRow.alignment = ["fill", "top"];
        excludedNamesRow.alignChildren = ["left", "center"];
        excludedNamesRow.helpTip = getLabel("tooltip.specifiedLayers");
        var excludedNamesLabel = excludedNamesRow.add("statictext", undefined, getLabel("fieldLabel.specifiedLayers") + getLocalizedSymbol("colon"));
        excludedNamesLabel.helpTip = excludedNamesRow.helpTip;
        var excludedNamesField = excludedNamesRow.add("edittext", undefined, DEFAULT_OPTIONS.excludedLayerNames);
        excludedNamesField.alignment = ["fill", "center"];
        excludedNamesField.helpTip = excludedNamesRow.helpTip;

        return {
            lockedLayerCheckbox: lockedControls.layerCheckbox,
            lockedObjectCheckbox: lockedControls.objectCheckbox,
            hiddenLayerCheckbox: hiddenControls.layerCheckbox,
            hiddenObjectCheckbox: hiddenControls.objectCheckbox,
            excludedNamesField: excludedNamesField
        };
    }

    /**
     * ロック／非表示のサブパネル（レイヤー・オブジェクトの2択）を構築する
     * @param {Group} parentGroup - 追加先のグループ
     * @param {string} titleLabelPath - パネルタイトルの LABELS パス
     * @param {string} tooltipLabelPath - パネルと各チェックボックスに設定する tooltip の LABELS パス
     * @param {boolean} layerDefaultValue - レイヤー側チェックボックスの初期値
     * @param {boolean} objectDefaultValue - オブジェクト側チェックボックスの初期値
     * @returns {Object} チェックボックス参照をまとめたオブジェクト
     */
    function buildExclusionSubPanel(parentGroup, titleLabelPath, tooltipLabelPath, layerDefaultValue, objectDefaultValue) {
        var exclusionSubPanel = parentGroup.add("panel", undefined, getLabel(titleLabelPath));
        applyPanelLayout(exclusionSubPanel, ["left", "fill"]);
        exclusionSubPanel.helpTip = getLabel(tooltipLabelPath);

        var layerCheckbox = exclusionSubPanel.add("checkbox", undefined, getLabel("checkbox.excludeLayer"));
        layerCheckbox.value = layerDefaultValue;
        layerCheckbox.helpTip = exclusionSubPanel.helpTip;
        var objectCheckbox = exclusionSubPanel.add("checkbox", undefined, getLabel("checkbox.excludeObject"));
        objectCheckbox.value = objectDefaultValue;
        objectCheckbox.helpTip = exclusionSubPanel.helpTip;

        return { layerCheckbox: layerCheckbox, objectCheckbox: objectCheckbox };
    }

    /**
     * 整理後パネル（空レイヤー削除）を構築する
     * @param {Window} parentDialog - 追加先のダイアログ
     * @returns {Object} チェックボックス参照をまとめたオブジェクト
     */
    function buildPostProcessPanel(parentDialog) {
        var postProcessPanel = parentDialog.add("panel", undefined, getLabel("panel.postProcess"));
        applyPanelLayout(postProcessPanel);
        var removeEmptyLayersCheckbox = postProcessPanel.add("checkbox", undefined, getLabel("checkbox.removeEmpty"));
        removeEmptyLayersCheckbox.value = DEFAULT_OPTIONS.removeEmptyLayers;
        removeEmptyLayersCheckbox.helpTip = getLabel("tooltip.removeEmpty");
        return { removeEmptyLayersCheckbox: removeEmptyLayersCheckbox };
    }

    /**
     * OK / キャンセルのボタン列を構築する
     * @param {Window} optionsDialog - 対象ダイアログ
     * @returns {void}
     */
    function buildDialogButtonRow(optionsDialog) {
        var dialogButtonRow = optionsDialog.add("group");
        dialogButtonRow.orientation = "row";
        dialogButtonRow.alignment = ["center", "top"];
        dialogButtonRow.alignChildren = ["center", "center"];
        var cancelButton = dialogButtonRow.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var okButton = dialogButtonRow.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        optionsDialog.defaultElement = okButton;
        optionsDialog.cancelElement = cancelButton;
    }

    // =========================================
    // 除外判定 / Exclusion rules
    // =========================================

    /**
     * "bg, temp" のような文字列をレイヤー名の配列に分解する（, または 、 区切り）
     * @param {string} excludedNamesText - 指定レイヤー欄の入力値
     * @returns {Array<string>} 前後の空白を除いたレイヤー名の配列
     */
    function parseExcludedLayerNames(excludedNamesText) {
        if (!excludedNamesText) return [];
        var rawLayerNames = excludedNamesText.split(/[,、]/);
        var parsedLayerNames = [];
        for (var i = 0; i < rawLayerNames.length; i++) {
            var trimmedLayerName = rawLayerNames[i].replace(/^\s+|\s+$/g, "");
            if (trimmedLayerName.length > 0) parsedLayerNames.push(trimmedLayerName);
        }
        return parsedLayerNames;
    }

    /**
     * レイヤー名が指定除外リストに含まれるか判定する
     * @param {string} layerName - 判定するレイヤー名
     * @param {Array<string>} excludedLayerNames - 除外レイヤー名の配列
     * @returns {boolean} 含まれていれば true
     */
    function isNameInExcludedList(layerName, excludedLayerNames) {
        if (!excludedLayerNames) return false;
        for (var i = 0; i < excludedLayerNames.length; i++) {
            if (excludedLayerNames[i] === layerName) return true;
        }
        return false;
    }

    /**
     * レイヤー自身が除外対象か判定する（指定名・ロック・非表示）
     * @param {Layer} targetLayer - 判定するレイヤー
     * @param {Object} organizeOptions - 処理設定
     * @returns {boolean} 除外対象なら true
     */
    function isLayerExcluded(targetLayer, organizeOptions) {
        if (!targetLayer || targetLayer.typename !== "Layer") return false;
        if (isNameInExcludedList(targetLayer.name, organizeOptions.excludedLayerNames)) return true;
        if (organizeOptions.ignoreLockedLayers && targetLayer.locked) return true;
        if (organizeOptions.ignoreHiddenLayers && !targetLayer.visible) return true;
        return false;
    }

    /**
     * 親方向に辿って除外対象のレイヤーがあるか判定する
     * @param {PageItem} targetItem - 判定するオブジェクト
     * @param {Object} organizeOptions - 処理設定
     * @param {boolean} skipNameExclusion - true のとき指定レイヤー名による除外を無視し、ロック／非表示だけを見る
     * @returns {boolean} 除外対象の祖先レイヤーがあれば true
     */
    function hasExcludedAncestorLayer(targetItem, organizeOptions, skipNameExclusion) {
        var ancestor = targetItem.parent;
        while (ancestor && ancestor.typename !== "Document") {
            if (ancestor.typename === "Layer") {
                if (organizeOptions.ignoreLockedLayers && ancestor.locked) return true;
                if (organizeOptions.ignoreHiddenLayers && !ancestor.visible) return true;
                if (!skipNameExclusion && isNameInExcludedList(ancestor.name, organizeOptions.excludedLayerNames)) return true;
            }
            ancestor = ancestor.parent;
        }
        return false;
    }

    /**
     * PageItem の真偽値プロパティを安全に読む
     * guides のように種類によっては存在しないプロパティがあるため、読めない場合は false を返す
     * @param {PageItem} targetItem - 対象オブジェクト
     * @param {string} propertyName - プロパティ名（locked / hidden / guides）
     * @returns {boolean} プロパティが true のときのみ true
     */
    function readItemFlag(targetItem, propertyName) {
        try {
            return targetItem[propertyName] === true;
        } catch (e) {
            return false;
        }
    }

    /**
     * オブジェクトがガイドか判定する
     * @param {PageItem} targetItem - 対象オブジェクト
     * @returns {boolean} ガイドなら true
     */
    function isGuideItem(targetItem) {
        return readItemFlag(targetItem, "guides");
    }

    /**
     * オブジェクト自身が除外対象か判定する（ロック／非表示）
     * @param {PageItem} targetItem - 判定するオブジェクト
     * @param {Object} organizeOptions - 処理設定
     * @returns {boolean} 除外対象なら true
     */
    function isObjectExcluded(targetItem, organizeOptions) {
        if (organizeOptions.ignoreLockedObjects && readItemFlag(targetItem, "locked")) return true;
        if (organizeOptions.ignoreHiddenObjects && readItemFlag(targetItem, "hidden")) return true;
        return false;
    }

    // =========================================
    // レイヤー操作 / Layer helpers
    // =========================================

    /**
     * 名前でトップレベルレイヤーを検索する
     * @param {string} layerName - 探すレイヤー名
     * @returns {Layer} 見つかったレイヤー。無ければ null
     */
    function findTopLevelLayerByName(layerName) {
        var topLevelLayers = activeDocument.layers;
        for (var i = 0; i < topLevelLayers.length; i++) {
            if (topLevelLayers[i].name === layerName) return topLevelLayers[i];
        }
        return null;
    }

    /**
     * 指定名のトップレベルレイヤーを取得し、無ければ作成する
     * @param {string} layerName - レイヤー名
     * @returns {Layer} 取得または作成したレイヤー
     */
    function getOrCreateTopLevelLayer(layerName) {
        var foundLayer = findTopLevelLayerByName(layerName);
        if (foundLayer) return foundLayer;
        var createdLayer = activeDocument.layers.add();
        createdLayer.name = layerName;
        return createdLayer;
    }

    /**
     * _guide レイヤーを必要になった時点で取得・作成する
     * @returns {Layer} _guide レイヤー
     */
    function getOrCreateGuideLayer() {
        if (!guideLayer) {
            guideLayer = getOrCreateTopLevelLayer(GUIDE_LAYER_NAME);
        }
        return guideLayer;
    }

    /**
     * トップレベルレイヤーを配列に控える
     * 処理中のレイヤー増減でインデックスがずれるのを防ぐために使う
     * @returns {Array<Layer>} トップレベルレイヤーの配列
     */
    function getTopLevelLayerSnapshot() {
        var topLevelLayers = [];
        for (var i = 0; i < activeDocument.layers.length; i++) {
            topLevelLayers.push(activeDocument.layers[i]);
        }
        return topLevelLayers;
    }

    /**
     * システム管理レイヤー（_guide / _pasteboard）か判定する
     * @param {Layer} targetLayer - 判定するレイヤー
     * @returns {boolean} 保護対象なら true
     */
    function isProtectedSystemLayer(targetLayer) {
        if (!targetLayer || targetLayer.typename !== "Layer") return false;
        return targetLayer.name === GUIDE_LAYER_NAME || targetLayer.name === PASTEBOARD_LAYER_NAME;
    }

    /**
     * レイヤーが空（オブジェクトもサブレイヤーも無い）か判定する
     * @param {Layer} targetLayer - 判定するレイヤー
     * @returns {boolean} 空なら true
     */
    function isLayerEmpty(targetLayer) {
        return targetLayer.pageItems.length === 0 && targetLayer.layers.length === 0;
    }

    /**
     * レイヤーを一時的に書き込み可（ロック解除・表示）にして処理を実行し、終了時に元の状態へ戻す
     * ここで削除されるレイヤーは無い前提のため、復元は素通しで行う
     * @param {Layer} targetLayer - 対象レイヤー
     * @param {function():*} runWithLayer - 書き込み可の状態で実行する処理
     * @returns {*} runWithLayer の戻り値
     */
    function withWritableLayer(targetLayer, runWithLayer) {
        if (!targetLayer || targetLayer.typename !== "Layer") {
            return runWithLayer();
        }
        var wasLocked = targetLayer.locked;
        var wasHidden = !targetLayer.visible;
        if (wasLocked) targetLayer.locked = false;
        if (wasHidden) targetLayer.visible = true;
        try {
            return runWithLayer();
        } finally {
            if (wasLocked) targetLayer.locked = true;
            if (wasHidden) targetLayer.visible = false;
        }
    }

    /**
     * レイヤーを削除する（ロックされていれば解除してから削除する）
     * 親レイヤーのロックなどで削除できない場合は何もしない
     * @param {Layer} targetLayer - 削除するレイヤー
     * @returns {boolean} 削除できたら true
     */
    function removeLayerSafely(targetLayer) {
        try {
            if (targetLayer.locked) targetLayer.locked = false;
            targetLayer.remove();
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * レイヤーを最前面へ移動する（ロック／非表示は一時解除）
     * @param {Layer} targetLayer - 対象レイヤー
     * @returns {void}
     */
    function bringLayerToFront(targetLayer) {
        if (!targetLayer) return;
        withWritableLayer(targetLayer, function () {
            targetLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
        });
    }

    // =========================================
    // ロック／非表示の一時解除 / Suspending lock and hidden state
    // =========================================

    /**
     * 移動の妨げになるロック／非表示を一時解除し、復元用の記録を返す
     * 「対象外にする」が有効な側は移動候補に含まれないため、その分は解除しない
     * @param {Array<PageItem>} movableItems - 移動候補のオブジェクト配列
     * @param {Object} organizeOptions - 処理設定
     * @returns {Array<Object>} 復元用の記録（上位レイヤー→下位レイヤー→オブジェクトの順）
     */
    function suspendLockAndHidden(movableItems, organizeOptions) {
        var suspendedEntries = [];
        var unlockLayers = !organizeOptions.ignoreLockedLayers;
        var showLayers = !organizeOptions.ignoreHiddenLayers;
        if (unlockLayers || showLayers) {
            suspendLayerLockAndHidden(activeDocument, unlockLayers, showLayers, suspendedEntries);
        }
        var unlockItems = !organizeOptions.ignoreLockedObjects;
        var showItems = !organizeOptions.ignoreHiddenObjects;
        if (unlockItems || showItems) {
            suspendItemLockAndHidden(movableItems, unlockItems, showItems, suspendedEntries);
        }
        return suspendedEntries;
    }

    /**
     * レイヤーツリーを上位から辿ってロック／非表示を解除する
     * 親を先に解除しないと子の解除が効かないため、必ず上位から処理する
     * @param {Document|Layer} layerContainer - 探索するドキュメントまたはレイヤー
     * @param {boolean} unlockLocked - ロックを解除するなら true
     * @param {boolean} showHidden - 非表示を表示にするなら true
     * @param {Array<Object>} suspendedEntries - 記録の追加先
     * @returns {void}
     */
    function suspendLayerLockAndHidden(layerContainer, unlockLocked, showHidden, suspendedEntries) {
        for (var i = 0; i < layerContainer.layers.length; i++) {
            var childLayer = layerContainer.layers[i];
            var wasLocked = unlockLocked && childLayer.locked;
            var wasHidden = showHidden && !childLayer.visible;
            if (wasLocked || wasHidden) {
                if (wasLocked) childLayer.locked = false;
                if (wasHidden) childLayer.visible = true;
                suspendedEntries.push({ node: childLayer, isLayer: true, wasLocked: wasLocked, wasHidden: wasHidden });
            }
            suspendLayerLockAndHidden(childLayer, unlockLocked, showHidden, suspendedEntries);
        }
    }

    /**
     * 移動候補オブジェクトのロック／非表示を解除する
     * 親レイヤーの解除後に呼ぶこと
     * @param {Array<PageItem>} movableItems - 移動候補のオブジェクト配列
     * @param {boolean} unlockLocked - ロックを解除するなら true
     * @param {boolean} showHidden - 非表示を表示にするなら true
     * @param {Array<Object>} suspendedEntries - 記録の追加先
     * @returns {void}
     */
    function suspendItemLockAndHidden(movableItems, unlockLocked, showHidden, suspendedEntries) {
        for (var i = 0; i < movableItems.length; i++) {
            var targetItem = movableItems[i];
            var wasLocked = unlockLocked && readItemFlag(targetItem, "locked");
            var wasHidden = showHidden && readItemFlag(targetItem, "hidden");
            if (!wasLocked && !wasHidden) continue;
            if (wasLocked) targetItem.locked = false;
            if (wasHidden) targetItem.hidden = false;
            suspendedEntries.push({ node: targetItem, isLayer: false, wasLocked: wasLocked, wasHidden: wasHidden });
        }
    }

    /**
     * suspendLockAndHidden で解除したロック／非表示を元へ戻す
     * 子より親を後に戻す必要があるため、記録の逆順で処理する
     * @param {Array<Object>} suspendedEntries - 復元する記録の配列
     * @returns {void}
     */
    function restoreLockAndHidden(suspendedEntries) {
        for (var i = suspendedEntries.length - 1; i >= 0; i--) {
            var suspendedEntry = suspendedEntries[i];
            try {
                if (suspendedEntry.wasLocked) suspendedEntry.node.locked = true;
                if (suspendedEntry.wasHidden) {
                    if (suspendedEntry.isLayer) suspendedEntry.node.visible = false;
                    else suspendedEntry.node.hidden = true;
                }
            } catch (e) {
                // 統合処理で削除されたレイヤーなど、書き戻せないものはスキップ
            }
        }
    }

    // =========================================
    // 移動処理 / Moving items
    // =========================================

    /**
     * 収集した移動エントリを移動先レイヤーへ移動する
     * 成否にかかわらず処理済みフラグを立て、同じオブジェクトを二重に数えないようにする
     * @param {Array<Object>} moveEntries - { item: PageItem, index: number } の配列
     * @param {Layer} targetLayer - 移動先レイヤー
     * @param {Array<boolean>} handledFlags - 処理済みフラグ（不要なら null）
     * @returns {number} 移動できなかった件数
     */
    function moveEntriesToLayer(moveEntries, targetLayer, handledFlags) {
        if (moveEntries.length === 0) return 0;
        return withWritableLayer(targetLayer, function () {
            var failedMoves = 0;
            var i = moveEntries.length;
            while (i--) {
                try {
                    moveEntries[i].item.move(targetLayer, ElementPlacement.PLACEATBEGINNING);
                } catch (e) {
                    failedMoves++;
                }
                if (handledFlags && typeof moveEntries[i].index === "number") {
                    handledFlags[moveEntries[i].index] = true;
                }
            }
            return failedMoves;
        });
    }

    /**
     * 未処理のアイテムをガイドとそれ以外に振り分け、移動エントリを作る
     * @param {Array<PageItem>} movableItems - 移動候補のオブジェクト配列
     * @param {Array<boolean>} handledFlags - 処理済みフラグ
     * @param {function(number):boolean} shouldMoveItem - 対象に含めるか判定する関数（null なら未処理すべて）
     * @returns {{normal: Array<Object>, guide: Array<Object>}} 通常オブジェクトとガイドの移動エントリ
     */
    function collectMoveEntries(movableItems, handledFlags, shouldMoveItem) {
        var moveEntries = { normal: [], guide: [] };
        for (var i = 0; i < movableItems.length; i++) {
            if (handledFlags[i]) continue;
            if (shouldMoveItem && !shouldMoveItem(i)) continue;
            var moveEntry = { item: movableItems[i], index: i };
            if (isGuideItem(movableItems[i])) moveEntries.guide.push(moveEntry);
            else moveEntries.normal.push(moveEntry);
        }
        return moveEntries;
    }

    /**
     * 通常オブジェクトとガイドをそれぞれの移動先レイヤーへ送る
     * @param {{normal: Array<Object>, guide: Array<Object>}} moveEntries - 振り分け済みの移動エントリ
     * @param {Layer} normalTargetLayer - 通常オブジェクトの移動先レイヤー
     * @param {Array<boolean>} handledFlags - 処理済みフラグ
     * @returns {number} 移動に失敗した件数
     */
    function moveEntriesToTargets(moveEntries, normalTargetLayer, handledFlags) {
        var failedMoves = moveEntriesToLayer(moveEntries.normal, normalTargetLayer, handledFlags);
        if (moveEntries.guide.length > 0) {
            failedMoves += moveEntriesToLayer(moveEntries.guide, getOrCreateGuideLayer(), handledFlags);
        }
        return failedMoves;
    }

    /**
     * レイヤー以下の pageItem を再帰的に集める
     * layer.pageItems はサブレイヤーの中身を含まないため、サブレイヤーは個別に辿る
     * @param {Layer} targetLayer - 探索するレイヤー
     * @param {Array<Object>} collectedEntries - 収集先の配列（{ item: PageItem } を追加）
     * @returns {void}
     */
    function collectPageItemEntriesRecursive(targetLayer, collectedEntries) {
        for (var i = 0; i < targetLayer.layers.length; i++) {
            collectPageItemEntriesRecursive(targetLayer.layers[i], collectedEntries);
        }
        for (var j = 0; j < targetLayer.pageItems.length; j++) {
            var pageItem = targetLayer.pageItems[j];
            if (pageItem.parent !== targetLayer) continue;
            collectedEntries.push({ item: pageItem });
        }
    }

    // =========================================
    // アートボード判定 / Artboard resolution
    // =========================================

    /**
     * 移動候補オブジェクトの重心座標を先にまとめて求める
     * アートボードごとに geometricBounds を取り直さないための前計算
     * @param {Array<PageItem>} movableItems - 移動候補のオブジェクト配列
     * @returns {Array<Array<number>>} [x, y] の配列（座標を取得できなかった要素は null）
     */
    function buildItemCentroids(movableItems) {
        var itemCentroids = [];
        for (var i = 0; i < movableItems.length; i++) {
            var itemBounds = null;
            try {
                itemBounds = movableItems[i].geometricBounds; // [left, top, right, bottom]
            } catch (e) {
                itemBounds = null;
            }
            if (itemBounds) {
                itemCentroids.push([(itemBounds[0] + itemBounds[2]) / 2, (itemBounds[1] + itemBounds[3]) / 2]);
            } else {
                itemCentroids.push(null);
            }
        }
        return itemCentroids;
    }

    /**
     * 重心がアートボード矩形に含まれるか判定する
     * @param {Array<number>} itemCentroid - [x, y]（取得できなかった場合は null）
     * @param {Array<number>} artboardRect - [left, top, right, bottom]
     * @returns {boolean} 含まれていれば true
     */
    function isCentroidInsideArtboard(itemCentroid, artboardRect) {
        if (!itemCentroid) return false;
        return (
            itemCentroid[0] >= artboardRect[0] &&
            itemCentroid[0] <= artboardRect[2] &&
            itemCentroid[1] <= artboardRect[1] &&
            itemCentroid[1] >= artboardRect[3]
        );
    }

    /**
     * 指定アートボードに重心が入るかを判定するフィルター関数を作る
     * @param {Array<Array<number>>} itemCentroids - 重心座標の配列
     * @param {Array<number>} artboardRect - [left, top, right, bottom]
     * @returns {function(number):boolean} インデックスを受け取る判定関数
     */
    function buildCentroidFilter(itemCentroids, artboardRect) {
        return function (itemIndex) {
            return isCentroidInsideArtboard(itemCentroids[itemIndex], artboardRect);
        };
    }

    /**
     * オプション設定に応じたレイヤー名の区切り文字を返す
     * @param {Object} organizeOptions - 処理設定
     * @returns {string} 区切り文字
     */
    function getSeparatorString(organizeOptions) {
        var separatorText = SEPARATOR_CHARACTERS[organizeOptions.layerNameSeparatorIndex];
        return (typeof separatorText === "string") ? separatorText : SEPARATOR_CHARACTERS[SEPARATOR_UNDERSCORE];
    }

    /**
     * アートボードの表示名を返す（空名のときは代替名）
     * @param {number} artboardIndex - アートボードのインデックス
     * @returns {string} アートボード名
     */
    function getArtboardDisplayName(artboardIndex) {
        var artboardName = documentArtboards[artboardIndex].name;
        if (!artboardName) return getLabel("fallbackName.artboard");
        return artboardName;
    }

    /**
     * アートボード番号と名前からレイヤー名を組み立てる
     * @param {number} artboardIndex - アートボードのインデックス
     * @param {Object} organizeOptions - 処理設定
     * @returns {string} 組み立てたレイヤー名
     */
    function getArtboardLayerName(artboardIndex, organizeOptions) {
        var nameParts = [];
        if (organizeOptions.includeArtboardNumber) nameParts.push(String(artboardIndex + 1));
        if (organizeOptions.includeArtboardName) nameParts.push(getArtboardDisplayName(artboardIndex));
        if (nameParts.length === 0) nameParts.push(String(artboardIndex + 1));
        return nameParts.join(getSeparatorString(organizeOptions));
    }

    /**
     * 旧仕様レイヤー名（アートボード名のみ）からアートボードのインデックスを引く
     * @param {string} layerName - 判定するレイヤー名
     * @returns {number} 一致したアートボードのインデックス。無ければ -1
     */
    function findLegacyArtboardIndexByLayerName(layerName) {
        for (var i = 0; i < documentArtboards.length; i++) {
            if (getArtboardDisplayName(i) === layerName) return i;
        }
        return -1;
    }

    /**
     * 処理対象のアートボード範囲（start <= index < end）を返す
     * @param {Object} organizeOptions - 処理設定
     * @returns {{start: number, end: number}} 処理対象の範囲
     */
    function getTargetArtboardRange(organizeOptions) {
        if (organizeOptions.currentArtboardOnly) {
            var activeIndex = documentArtboards.getActiveArtboardIndex();
            return { start: activeIndex, end: activeIndex + 1 };
        }
        return { start: 0, end: documentArtboards.length };
    }

    // =========================================
    // 振り分け / Distribution
    // =========================================

    /**
     * トップレベルの移動候補オブジェクトを集める
     * ガイドは指定レイヤー名による除外を無視して常に回収する（ロック／非表示の祖先は尊重）
     * @param {Object} organizeOptions - 処理設定
     * @returns {Array<PageItem>} 移動候補のオブジェクト配列
     */
    function collectMovableItems(organizeOptions) {
        var movableItems = [];
        var allPageItems = activeDocument.pageItems;
        var pageItemCount = allPageItems.length;
        for (var i = 0; i < pageItemCount; i++) {
            var pageItem = allPageItems[i];
            var itemParentType = pageItem.parent.typename;
            if (itemParentType !== "Layer" && itemParentType !== "Document") continue;

            if (isGuideItem(pageItem)) {
                if (!hasExcludedAncestorLayer(pageItem, organizeOptions, true)) movableItems.push(pageItem);
                continue;
            }
            if (hasExcludedAncestorLayer(pageItem, organizeOptions, false)) continue;
            if (isObjectExcluded(pageItem, organizeOptions)) continue;
            movableItems.push(pageItem);
        }
        return movableItems;
    }

    /**
     * 各アートボードに対応するレイヤーへオブジェクトを振り分ける
     * @param {Array<PageItem>} movableItems - 移動候補のオブジェクト配列
     * @param {Array<Array<number>>} itemCentroids - 重心座標の配列
     * @param {Object} organizeOptions - 処理設定
     * @param {Array<boolean>} handledFlags - 処理済みフラグ
     * @param {{start: number, end: number}} artboardRange - 処理対象のアートボード範囲
     * @returns {number} 移動に失敗した件数
     */
    function assignItemsToArtboardLayers(movableItems, itemCentroids, organizeOptions, handledFlags, artboardRange) {
        var failedMoves = 0;
        for (var artboardIndex = artboardRange.start; artboardIndex < artboardRange.end; artboardIndex++) {
            var artboardLayer = getOrCreateTopLevelLayer(getArtboardLayerName(artboardIndex, organizeOptions));
            var centroidFilter = buildCentroidFilter(itemCentroids, documentArtboards[artboardIndex].artboardRect);
            var moveEntries = collectMoveEntries(movableItems, handledFlags, centroidFilter);
            failedMoves += moveEntriesToTargets(moveEntries, artboardLayer, handledFlags);
        }
        return failedMoves;
    }

    /**
     * どのアートボードにも属さなかったオブジェクトを _pasteboard / _guide へ振り分ける
     * @param {Array<PageItem>} movableItems - 移動候補のオブジェクト配列
     * @param {Array<boolean>} handledFlags - 処理済みフラグ
     * @returns {number} 移動に失敗した件数
     */
    function assignLeftoverItems(movableItems, handledFlags) {
        var moveEntries = collectMoveEntries(movableItems, handledFlags, null);
        if (moveEntries.normal.length === 0 && moveEntries.guide.length === 0) return 0;
        /* 通常オブジェクトが無いときは _pasteboard を作らない / Skip creating _pasteboard when only guides are left */
        var pasteboardLayer = (moveEntries.normal.length > 0) ? getOrCreateTopLevelLayer(PASTEBOARD_LAYER_NAME) : null;
        return moveEntriesToTargets(moveEntries, pasteboardLayer, handledFlags);
    }

    /**
     * 対象アートボードのレイヤーを上から 1→2→3… の順に並べ、_guide を最前面に置く
     * @param {Object} organizeOptions - 処理設定
     * @param {{start: number, end: number}} artboardRange - 処理対象のアートボード範囲
     * @returns {void}
     */
    function applyLayerOrder(organizeOptions, artboardRange) {
        for (var i = artboardRange.end - 1; i >= artboardRange.start; i--) {
            bringLayerToFront(getOrCreateTopLevelLayer(getArtboardLayerName(i, organizeOptions)));
        }
        bringLayerToFront(guideLayer);
    }

    /**
     * 旧仕様（アートボード名のみ）のレイヤーを新仕様レイヤーへ統合する
     * @param {Object} organizeOptions - 処理設定
     * @returns {number} 移動に失敗した件数
     */
    function mergeLegacyLayers(organizeOptions) {
        var failedMoves = 0;
        var topLevelLayers = getTopLevelLayerSnapshot();
        for (var i = topLevelLayers.length - 1; i >= 0; i--) {
            var legacyLayer = topLevelLayers[i];
            if (isProtectedSystemLayer(legacyLayer)) continue;
            if (isLayerExcluded(legacyLayer, organizeOptions)) continue;
            var legacyArtboardIndex = findLegacyArtboardIndexByLayerName(legacyLayer.name);
            if (legacyArtboardIndex < 0) continue;

            var targetLayer = getOrCreateTopLevelLayer(getArtboardLayerName(legacyArtboardIndex, organizeOptions));
            if (legacyLayer === targetLayer) continue;

            failedMoves += mergeSingleLegacyLayer(legacyLayer, targetLayer);
        }
        return failedMoves;
    }

    /**
     * 旧仕様レイヤー1枚を移動先へ統合し、空になったら削除する
     * @param {Layer} legacyLayer - 統合元の旧仕様レイヤー
     * @param {Layer} targetLayer - 統合先のレイヤー
     * @returns {number} 移動に失敗した件数
     */
    function mergeSingleLegacyLayer(legacyLayer, targetLayer) {
        var collectedEntries = [];
        collectPageItemEntriesRecursive(legacyLayer, collectedEntries);
        var failedMoves = withWritableLayer(legacyLayer, function () {
            return moveEntriesToLayer(collectedEntries, targetLayer, null);
        });
        if (isLayerEmpty(legacyLayer) && activeDocument.layers.length > 1) {
            removeLayerSafely(legacyLayer);
        }
        return failedMoves;
    }

    // =========================================
    // 空レイヤーの削除 / Removing empty layers
    // =========================================

    /**
     * 空になったサブレイヤーを再帰的に削除する（除外対象は中身ごと残す）
     * @param {Layer} parentLayer - 親レイヤー
     * @param {Object} organizeOptions - 処理設定
     * @returns {void}
     */
    function removeEmptySubLayers(parentLayer, organizeOptions) {
        for (var i = parentLayer.layers.length - 1; i >= 0; i--) {
            var childLayer = parentLayer.layers[i];
            if (isLayerExcluded(childLayer, organizeOptions)) continue;
            removeEmptySubLayers(childLayer, organizeOptions);
            if (isProtectedSystemLayer(childLayer)) continue;
            if (isLayerEmpty(childLayer)) removeLayerSafely(childLayer);
        }
    }

    /**
     * 処理後に残った空レイヤーを削除する（_guide / _pasteboard は保護）
     * @param {Object} organizeOptions - 処理設定
     * @returns {void}
     */
    function cleanupEmptyLayers(organizeOptions) {
        var topLevelLayers = getTopLevelLayerSnapshot();
        for (var i = topLevelLayers.length - 1; i >= 0; i--) {
            var topLevelLayer = topLevelLayers[i];
            if (isLayerExcluded(topLevelLayer, organizeOptions)) continue;
            removeEmptySubLayers(topLevelLayer, organizeOptions);
            if (isProtectedSystemLayer(topLevelLayer)) continue;
            if (isLayerEmpty(topLevelLayer) && activeDocument.layers.length > 1) {
                removeLayerSafely(topLevelLayer);
            }
        }
    }

    // =========================================
    // メイン / Main
    // =========================================

    /**
     * 設定に従ってドキュメント全体のレイヤー整理を実行する
     * @param {Object} organizeOptions - 処理設定
     * @returns {number} 移動に失敗した件数
     */
    function organizeDocumentLayers(organizeOptions) {
        var movableItems = collectMovableItems(organizeOptions);
        var itemCentroids = buildItemCentroids(movableItems);
        var handledFlags = [];
        var artboardRange = getTargetArtboardRange(organizeOptions);
        var failedMoves = 0;

        guideLayer = findTopLevelLayerByName(GUIDE_LAYER_NAME); // 既存があれば先に拾う / capture existing if any

        var suspendedEntries = suspendLockAndHidden(movableItems, organizeOptions);
        try {
            failedMoves += assignItemsToArtboardLayers(movableItems, itemCentroids, organizeOptions, handledFlags, artboardRange);
            if (!organizeOptions.currentArtboardOnly) {
                failedMoves += assignLeftoverItems(movableItems, handledFlags);
                failedMoves += mergeLegacyLayers(organizeOptions);
            }
            applyLayerOrder(organizeOptions, artboardRange);
        } finally {
            restoreLockAndHidden(suspendedEntries);
        }

        if (organizeOptions.removeEmptyLayers) {
            cleanupEmptyLayers(organizeOptions);
        }
        return failedMoves;
    }

    var organizeOptions = showOptionsDialog();
    if (!organizeOptions) {
        return;
    }

    var failedMoves = organizeDocumentLayers(organizeOptions);
    if (failedMoves > 0) {
        alert(getLabel("alert.moveFailed").replace("{count}", String(failedMoves)));
    }

})();
