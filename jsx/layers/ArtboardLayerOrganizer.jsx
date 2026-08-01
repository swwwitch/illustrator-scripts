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
var SCRIPT_VERSION  = "v1.3.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-04";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-01";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArtboardLayerOrganizer.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ArtboardLayerOrganizer.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n50aacdeb4908"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* システム管理レイヤー名（空でも削除しない） / System-managed layer names (never deleted) */
    var GUIDE_LAYER_NAME = "_guide";
    var PASTEBOARD_LAYER_NAME = "_pasteboard";

    /* パネル共通レイアウト / Common panel layout */
    var PANEL_MARGINS = [15, 20, 15, 10];
    var PANEL_SPACING = 8;

    /* 区切り文字ドロップダウンの選択位置 / Separator dropdown indexes */
    var SEPARATOR_UNDERSCORE = 0;
    var SEPARATOR_HYPHEN = 1;
    var SEPARATOR_SPACE = 2;
    var SEPARATOR_NONE = 3;

    /* ダイアログの初期値。必要に応じて編集 / Dialog defaults; edit as needed */
    var DEFAULT_OPTIONS = {
        removeEmptyLayers: true,
        includeArtboardNumber: true,
        includeArtboardName: true,
        useSeparator: true,
        separatorIndex: SEPARATOR_UNDERSCORE,
        ignoreLockedLayers: true,
        ignoreLockedObjects: true,
        ignoreHiddenLayers: true,
        ignoreHiddenObjects: true,
        excludedLayerNames: "bg"    // , または 、 区切り / Comma-separated
    };

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * Illustrator の UI 言語から表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = detectUILang();

    var LABELS = {
        dialog: {
            title: { ja: "アートボード単位でレイヤー整理", en: "Artboard Layer Organizer" }
        },
        panel: {
            targetArtboards: { ja: "対象のアートボード", en: "Target Artboards" },
            layerName: { ja: "レイヤー名の構成", en: "Layer Name Format" },
            exclusion: { ja: "対象外にする", en: "Exclude" },
            locked: { ja: "ロック", en: "Locked" },
            hidden: { ja: "非表示", en: "Hidden" },
            postProcess: { ja: "整理後", en: "After Organizing" }
        },
        radio: {
            currentArtboardOnly: { ja: "現在のみ", en: "Current only" },
            allArtboards: { ja: "すべて", en: "All" }
        },
        checkbox: {
            includeArtboardNumber: { ja: "アートボード番号", en: "Artboard Number" },
            includeArtboardName: { ja: "アートボード名", en: "Artboard Name" },
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
            specifiedLayers: { ja: "除外するレイヤー名", en: "Layer names to exclude" }
        },
        tooltip: {
            targetArtboards: {
                ja: "アートボードが1つのときは「現在のみ」に固定されます",
                en: "Locked to \"Current only\" when the document has just one artboard"
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
                ja: "一部のオブジェクトを移動できませんでした。\n移動できなかったオブジェクト: {count}件",
                en: "Some objects could not be moved.\nObjects that could not be moved: {count}"
            }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "panel.layerName" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var pathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        var labelText = labelNode[uiLang] || labelNode["en"];
        if (!labelText) return labelPath;
        return expandSymbolPlaceholders(labelText);
    }

    /**
     * {slash} などのプレースホルダを言語別の記号に展開する
     * @param {string} labelText - 展開前のテキスト
     * @returns {string} 展開後のテキスト
     */
    function expandSymbolPlaceholders(labelText) {
        return labelText
            .replace(/\{slash\}/g, getLocalizedSymbol("slash"))
            .replace(/\{colon\}/g, getLocalizedSymbol("colon"))
            .replace(/\{comma\}/g, getLocalizedSymbol("comma"))
            .replace(/\{openParen\}/g, getLocalizedSymbol("openParen"))
            .replace(/\{closeParen\}/g, getLocalizedSymbol("closeParen"));
    }

    /**
     * 言語別の記号を返す（日本語＝全角、英語＝半角）
     * @param {string} symbolName - 記号名（slash / colon / comma / openParen / closeParen）
     * @returns {string} 表示言語に合わせた記号
     */
    function getLocalizedSymbol(symbolName) {
        if (uiLang === "ja") {
            switch (symbolName) {
                case "slash": return "／";
                case "colon": return "：";
                case "comma": return "、";
                case "openParen": return "（";
                case "closeParen": return "）";
            }
        }
        switch (symbolName) {
            case "slash": return "/";
            case "colon": return ":";
            case "comma": return ", ";
            case "openParen": return "(";
            case "closeParen": return ")";
        }
        return "";
    }

    // =========================================
    // 前提チェック / Preconditions
    // =========================================

    if (app.documents.length === 0) {
        alert(getLabel("alert.noDocument"));
        return;
    }

    var activeDoc = app.activeDocument;
    var artboards = activeDoc.artboards;
    var guideLayer = null; // _guide レイヤー参照（必要時に取得・作成） / _guide layer reference, resolved on demand

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * パネル共通の見た目をまとめて設定する
     * @param {Panel} panel - 対象パネル
     * @param {number} [spacing] - 子要素の間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function applyPanelLayout(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ['fill', 'top'];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ダイアログを構築・表示し、選択結果の処理設定を返す
     * @returns {Object} 処理設定。キャンセル時は null
     */
    function showOptionsDialog() {
        var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];
        dialog.margins = [15, 20, 15, 15];

        var targetControls = buildTargetArtboardPanel(dialog);
        var layerNameControls = buildLayerNamePanel(dialog);
        var exclusionControls = buildExclusionPanel(dialog);
        var postProcessControls = buildPostProcessPanel(dialog);
        buildDialogButtonRow(dialog);

        if (dialog.show() !== 1) {
            return null;
        }

        return {
            removeEmptyLayers: postProcessControls.removeEmptyCheckbox.value,
            currentArtboardOnly: targetControls.currentArtboardRadio.value,
            includeArtboardNumber: layerNameControls.artboardNumberCheckbox.value,
            includeArtboardName: layerNameControls.artboardNameCheckbox.value,
            layerNameSeparatorIndex: layerNameControls.getSeparatorIndex(),
            ignoreLockedLayers: exclusionControls.lockedLayerCheckbox.value,
            ignoreLockedObjects: exclusionControls.lockedObjectCheckbox.value,
            ignoreHiddenLayers: exclusionControls.hiddenLayerCheckbox.value,
            ignoreHiddenObjects: exclusionControls.hiddenObjectCheckbox.value,
            excludedLayerNames: parseExcludedLayerNames(exclusionControls.specifiedLayersField.text)
        };
    }

    /**
     * 対象アートボードパネル（現在のみ／すべて）を構築する
     * @param {Window} parentContainer - 追加先のダイアログ
     * @returns {Object} ラジオボタン参照をまとめたオブジェクト
     */
    function buildTargetArtboardPanel(parentContainer) {
        var targetPanel = parentContainer.add("panel", undefined, getLabel("panel.targetArtboards"));
        applyPanelLayout(targetPanel);

        var scopeGroup = targetPanel.add("group");
        scopeGroup.orientation = "row";
        scopeGroup.alignChildren = ["left", "center"];

        var currentArtboardRadio = scopeGroup.add("radiobutton", undefined, getLabel("radio.currentArtboardOnly"));
        var allArtboardsRadio = scopeGroup.add("radiobutton", undefined, getLabel("radio.allArtboards"));
        targetPanel.helpTip = getLabel("tooltip.targetArtboards");
        currentArtboardRadio.helpTip = targetPanel.helpTip;
        allArtboardsRadio.helpTip = targetPanel.helpTip;

        if (artboards.length <= 1) {
            currentArtboardRadio.value = true;
            allArtboardsRadio.enabled = false;
        } else {
            allArtboardsRadio.value = true;
        }

        return {
            currentArtboardRadio: currentArtboardRadio,
            allArtboardsRadio: allArtboardsRadio
        };
    }

    /**
     * レイヤー名パネル（番号・名前・区切り文字）を構築する
     * @param {Window} parentContainer - 追加先のダイアログ
     * @returns {Object} チェックボックス参照と区切り文字インデックス取得関数
     */
    function buildLayerNamePanel(parentContainer) {
        var layerNamePanel = parentContainer.add("panel", undefined, getLabel("panel.layerName"));
        applyPanelLayout(layerNamePanel);

        var artboardNumberCheckbox = layerNamePanel.add("checkbox", undefined, getLabel("checkbox.includeArtboardNumber"));
        artboardNumberCheckbox.value = DEFAULT_OPTIONS.includeArtboardNumber;
        artboardNumberCheckbox.helpTip = getLabel("tooltip.includeArtboardNumber");

        var separatorGroup = layerNamePanel.add("group");
        separatorGroup.orientation = "row";
        separatorGroup.alignChildren = ["left", "center"];
        separatorGroup.helpTip = getLabel("tooltip.useSeparator");
        var useSeparatorCheckbox = separatorGroup.add("checkbox", undefined, getLabel("checkbox.useSeparator"));
        useSeparatorCheckbox.value = DEFAULT_OPTIONS.useSeparator;
        useSeparatorCheckbox.helpTip = separatorGroup.helpTip;
        var separatorDropdown = separatorGroup.add("dropdownlist", undefined, [
            getLabel("dropdown.separatorUnderscore"),
            getLabel("dropdown.separatorHyphen"),
            getLabel("dropdown.separatorSpace"),
            getLabel("dropdown.separatorNone")
        ]);
        separatorDropdown.selection = DEFAULT_OPTIONS.separatorIndex;
        separatorDropdown.helpTip = separatorGroup.helpTip;

        /**
         * 区切り文字まわりの有効／無効を、アートボード番号の状態に合わせて更新する
         * @returns {void}
         */
        function updateSeparatorEnabled() {
            var numberIncluded = artboardNumberCheckbox.value;
            separatorGroup.enabled = numberIncluded;
            separatorDropdown.enabled = numberIncluded && useSeparatorCheckbox.value;
        }

        updateSeparatorEnabled();
        artboardNumberCheckbox.onClick = updateSeparatorEnabled;
        useSeparatorCheckbox.onClick = updateSeparatorEnabled;

        var artboardNameCheckbox = layerNamePanel.add("checkbox", undefined, getLabel("checkbox.includeArtboardName"));
        artboardNameCheckbox.value = DEFAULT_OPTIONS.includeArtboardName;
        artboardNameCheckbox.helpTip = getLabel("tooltip.includeArtboardName");

        return {
            artboardNumberCheckbox: artboardNumberCheckbox,
            artboardNameCheckbox: artboardNameCheckbox,
            getSeparatorIndex: function () {
                if (useSeparatorCheckbox.value && separatorDropdown.selection) {
                    return separatorDropdown.selection.index;
                }
                return SEPARATOR_NONE;
            }
        };
    }

    /**
     * 対象外パネル（ロック／非表示／指定レイヤー）を構築する
     * @param {Window} parentContainer - 追加先のダイアログ
     * @returns {Object} チェックボックスと入力欄の参照をまとめたオブジェクト
     */
    function buildExclusionPanel(parentContainer) {
        var exclusionPanel = parentContainer.add("panel", undefined, getLabel("panel.exclusion"));
        applyPanelLayout(exclusionPanel);

        var stateGroup = exclusionPanel.add("group");
        stateGroup.orientation = "row";
        stateGroup.alignChildren = ["left", "top"];
        stateGroup.spacing = 15;

        var lockedControls = buildExclusionSubPanel(stateGroup, "panel.locked", "tooltip.lockedExclusion", DEFAULT_OPTIONS.ignoreLockedLayers, DEFAULT_OPTIONS.ignoreLockedObjects);
        var hiddenControls = buildExclusionSubPanel(stateGroup, "panel.hidden", "tooltip.hiddenExclusion", DEFAULT_OPTIONS.ignoreHiddenLayers, DEFAULT_OPTIONS.ignoreHiddenObjects);

        /* 指定レイヤー入力欄（対象外パネル最下部） / Specified layer names input */
        var specifiedGroup = exclusionPanel.add("group");
        specifiedGroup.orientation = "row";
        specifiedGroup.alignChildren = ["left", "center"];
        specifiedGroup.helpTip = getLabel("tooltip.specifiedLayers");
        var specifiedStaticText = specifiedGroup.add("statictext", undefined, getLabel("fieldLabel.specifiedLayers") + getLocalizedSymbol("colon"));
        specifiedStaticText.helpTip = specifiedGroup.helpTip;
        var specifiedLayersField = specifiedGroup.add("edittext", undefined, DEFAULT_OPTIONS.excludedLayerNames);
        specifiedLayersField.preferredSize.width = 200;
        specifiedLayersField.helpTip = specifiedGroup.helpTip;

        return {
            lockedLayerCheckbox: lockedControls.layerCheckbox,
            lockedObjectCheckbox: lockedControls.objectCheckbox,
            hiddenLayerCheckbox: hiddenControls.layerCheckbox,
            hiddenObjectCheckbox: hiddenControls.objectCheckbox,
            specifiedLayersField: specifiedLayersField
        };
    }

    /**
     * ロック／非表示のサブパネル（レイヤー・オブジェクトの2択）を構築する
     * @param {Group} parentContainer - 追加先のグループ
     * @param {string} titleLabelPath - パネルタイトルの LABELS パス
     * @param {string} tooltipLabelPath - パネルと各チェックボックスに設定する tooltip の LABELS パス
     * @param {boolean} layerDefaultValue - レイヤー側チェックボックスの初期値
     * @param {boolean} objectDefaultValue - オブジェクト側チェックボックスの初期値
     * @returns {Object} チェックボックス参照をまとめたオブジェクト
     */
    function buildExclusionSubPanel(parentContainer, titleLabelPath, tooltipLabelPath, layerDefaultValue, objectDefaultValue) {
        var subPanel = parentContainer.add("panel", undefined, getLabel(titleLabelPath));
        applyPanelLayout(subPanel);
        subPanel.helpTip = getLabel(tooltipLabelPath);

        var layerCheckbox = subPanel.add("checkbox", undefined, getLabel("checkbox.excludeLayer"));
        layerCheckbox.value = layerDefaultValue;
        layerCheckbox.helpTip = subPanel.helpTip;
        var objectCheckbox = subPanel.add("checkbox", undefined, getLabel("checkbox.excludeObject"));
        objectCheckbox.value = objectDefaultValue;
        objectCheckbox.helpTip = subPanel.helpTip;

        return { layerCheckbox: layerCheckbox, objectCheckbox: objectCheckbox };
    }

    /**
     * 後処理パネル（空レイヤー削除）を構築する
     * @param {Window} parentContainer - 追加先のダイアログ
     * @returns {Object} チェックボックス参照をまとめたオブジェクト
     */
    function buildPostProcessPanel(parentContainer) {
        var postProcessPanel = parentContainer.add("panel", undefined, getLabel("panel.postProcess"));
        applyPanelLayout(postProcessPanel);
        var removeEmptyCheckbox = postProcessPanel.add("checkbox", undefined, getLabel("checkbox.removeEmpty"));
        removeEmptyCheckbox.value = DEFAULT_OPTIONS.removeEmptyLayers;
        removeEmptyCheckbox.helpTip = getLabel("tooltip.removeEmpty");
        return { removeEmptyCheckbox: removeEmptyCheckbox };
    }

    /**
     * OK / キャンセルのボタン列を構築する
     * @param {Window} dialog - 対象ダイアログ
     * @returns {void}
     */
    function buildDialogButtonRow(dialog) {
        var buttonGroup = dialog.add("group");
        buttonGroup.orientation = "row";
        buttonGroup.alignment = ["center", "top"];
        var cancelButton = buttonGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var okButton = buttonGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        dialog.defaultElement = okButton;
        dialog.cancelElement = cancelButton;
    }

    // =========================================
    // 除外判定 / Exclusion rules
    // =========================================

    /**
     * "bg, temp" のような文字列をレイヤー名の配列に分解する（, または 、 区切り）
     * @param {string} inputText - 指定レイヤー欄の入力値
     * @returns {Array<string>} 前後の空白を除いたレイヤー名の配列
     */
    function parseExcludedLayerNames(inputText) {
        if (!inputText) return [];
        var rawNames = inputText.split(/[,、]/);
        var layerNames = [];
        for (var i = 0; i < rawNames.length; i++) {
            var trimmedName = rawNames[i].replace(/^\s+|\s+$/g, "");
            if (trimmedName.length > 0) layerNames.push(trimmedName);
        }
        return layerNames;
    }

    /**
     * レイヤー名が指定除外リストに含まれるか判定する
     * @param {string} layerName - 判定するレイヤー名
     * @param {Array<string>} excludedNames - 除外レイヤー名の配列
     * @returns {boolean} 含まれていれば true
     */
    function isNameInExcludedList(layerName, excludedNames) {
        if (!excludedNames || excludedNames.length === 0) return false;
        for (var i = 0; i < excludedNames.length; i++) {
            if (excludedNames[i] === layerName) return true;
        }
        return false;
    }

    /**
     * レイヤー自身が除外対象か判定する（指定名・ロック・非表示）
     * @param {Layer} layer - 判定するレイヤー
     * @param {Object} organizeOptions - 処理設定
     * @returns {boolean} 除外対象なら true
     */
    function isLayerExcluded(layer, organizeOptions) {
        if (!layer || layer.typename !== "Layer") return false;
        if (isNameInExcludedList(layer.name, organizeOptions.excludedLayerNames)) return true;
        if (organizeOptions.ignoreLockedLayers && layer.locked) return true;
        if (organizeOptions.ignoreHiddenLayers && !layer.visible) return true;
        return false;
    }

    /**
     * 親方向に辿って除外対象のレイヤーがあるか判定する
     * @param {PageItem} item - 判定するオブジェクト
     * @param {Object} organizeOptions - 処理設定
     * @param {boolean} lockAndHiddenOnly - true のとき指定レイヤー名による除外を無視し、ロック／非表示だけを見る
     * @returns {boolean} 除外対象の祖先レイヤーがあれば true
     */
    function hasExcludedAncestorLayer(item, organizeOptions, lockAndHiddenOnly) {
        var ancestor = item.parent;
        while (ancestor && ancestor.typename !== "Document") {
            if (ancestor.typename === "Layer") {
                if (organizeOptions.ignoreLockedLayers && ancestor.locked) return true;
                if (organizeOptions.ignoreHiddenLayers && !ancestor.visible) return true;
                if (!lockAndHiddenOnly && isNameInExcludedList(ancestor.name, organizeOptions.excludedLayerNames)) return true;
            }
            ancestor = ancestor.parent;
        }
        return false;
    }

    /**
     * PageItem の真偽値プロパティを安全に読む
     * 一部の PageItem は locked / hidden / guides の参照で例外になるため、取得できない場合は false を返す
     * @param {PageItem} item - 対象オブジェクト
     * @param {string} propertyName - プロパティ名（locked / hidden / guides）
     * @returns {boolean} プロパティが true のときのみ true
     */
    function readItemFlag(item, propertyName) {
        try {
            return item[propertyName] === true;
        } catch (e) {
            return false;
        }
    }

    /**
     * オブジェクトがガイドか判定する
     * @param {PageItem} item - 対象オブジェクト
     * @returns {boolean} ガイドなら true
     */
    function isGuideItem(item) {
        return readItemFlag(item, "guides");
    }

    /**
     * オブジェクト自身が除外対象か判定する（ロック／非表示）
     * @param {PageItem} item - 判定するオブジェクト
     * @param {Object} organizeOptions - 処理設定
     * @returns {boolean} 除外対象なら true
     */
    function isObjectExcluded(item, organizeOptions) {
        if (organizeOptions.ignoreLockedObjects && readItemFlag(item, "locked")) return true;
        if (organizeOptions.ignoreHiddenObjects && readItemFlag(item, "hidden")) return true;
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
        for (var i = 0; i < activeDoc.layers.length; i++) {
            if (activeDoc.layers[i].name === layerName) {
                return activeDoc.layers[i];
            }
        }
        return null;
    }

    /**
     * 指定名のトップレベルレイヤーを取得し、無ければ作成する
     * @param {string} layerName - レイヤー名
     * @returns {Layer} 取得または作成したレイヤー
     */
    function getOrCreateTopLevelLayer(layerName) {
        var layer = findTopLevelLayerByName(layerName);
        if (layer) {
            return layer;
        }
        layer = activeDoc.layers.add();
        layer.name = layerName;
        return layer;
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
     * システム管理レイヤー（_guide / _pasteboard）か判定する
     * @param {Layer} layer - 判定するレイヤー
     * @returns {boolean} 保護対象なら true
     */
    function isProtectedSystemLayer(layer) {
        if (!layer || layer.typename !== "Layer") return false;
        return layer.name === GUIDE_LAYER_NAME || layer.name === PASTEBOARD_LAYER_NAME;
    }

    /**
     * レイヤーを一時的に書き込み可（ロック解除・表示）にして処理を実行し、終了時に元の状態へ戻す
     * 処理内でレイヤーが削除されていた場合は復元せず黙って抜ける
     * @param {Layer} layer - 対象レイヤー
     * @param {function():*} runWithLayer - 書き込み可の状態で実行する処理
     * @returns {*} runWithLayer の戻り値
     */
    function withWritableLayer(layer, runWithLayer) {
        if (!layer || layer.typename !== "Layer") {
            return runWithLayer();
        }
        var wasLocked = layer.locked;
        var wasVisible = layer.visible;
        if (wasLocked) layer.locked = false;
        if (!wasVisible) layer.visible = true;
        try {
            return runWithLayer();
        } finally {
            try {
                if (wasLocked) layer.locked = true;
                if (!wasVisible) layer.visible = false;
            } catch (e) {
                // 処理内でレイヤー削除済み等で参照不能ならスキップ
            }
        }
    }

    /**
     * レイヤーを最前面へ移動する（ロック／非表示は一時解除）
     * @param {Layer} layer - 対象レイヤー
     * @returns {void}
     */
    function bringLayerToFront(layer) {
        if (!layer) return;
        withWritableLayer(layer, function () {
            layer.zOrder(ZOrderMethod.BRINGTOFRONT);
        });
    }

    /**
     * 空になったサブレイヤーを再帰的に削除する
     * @param {Layer} parentLayer - 親レイヤー
     * @param {Object} organizeOptions - 処理設定
     * @returns {void}
     */
    function removeEmptySubLayers(parentLayer, organizeOptions) {
        for (var i = parentLayer.layers.length - 1; i >= 0; i--) {
            var subLayer = parentLayer.layers[i];
            removeEmptySubLayers(subLayer, organizeOptions);
            if (isLayerExcluded(subLayer, organizeOptions)) continue;
            if (subLayer.pageItems.length === 0 && subLayer.layers.length === 0) {
                subLayer.remove();
            }
        }
    }

    // =========================================
    // 移動処理 / Moving items
    // =========================================

    /**
     * 収集した移動エントリを移動先レイヤーへ移動する
     * @param {Array<Object>} moveEntries - { item: PageItem, index: number } の配列
     * @param {Layer} targetLayer - 移動先レイヤー
     * @param {Array<boolean>} movedFlags - 移動済みフラグ（不要なら null）
     * @returns {{moved: number, failed: number}} 移動できた件数と失敗件数
     */
    function moveEntriesToLayer(moveEntries, targetLayer, movedFlags) {
        var moveResult = { moved: 0, failed: 0 };
        if (moveEntries.length === 0) return moveResult;
        return withWritableLayer(targetLayer, function () {
            var i = moveEntries.length;
            while (i--) {
                try {
                    moveEntries[i].item.move(targetLayer, ElementPlacement.PLACEATBEGINNING);
                    if (movedFlags && typeof moveEntries[i].index === "number") {
                        movedFlags[moveEntries[i].index] = true;
                    }
                    moveResult.moved++;
                } catch (e) {
                    moveResult.failed++;
                }
            }
            return moveResult;
        });
    }

    /**
     * 未処理のアイテムをガイドとそれ以外に振り分け、移動エントリを作る
     * @param {Array<PageItem>} movableItems - 移動候補のオブジェクト配列
     * @param {Array<boolean>} movedFlags - 移動済みフラグ
     * @param {function(PageItem):boolean} acceptItem - 対象に含めるか判定する関数（null なら未処理すべて）
     * @returns {{normal: Array<Object>, guide: Array<Object>}} 通常オブジェクトとガイドの移動エントリ
     */
    function buildMoveEntries(movableItems, movedFlags, acceptItem) {
        var moveEntries = { normal: [], guide: [] };
        for (var i = 0; i < movableItems.length; i++) {
            if (movedFlags[i]) continue;
            var candidateItem = movableItems[i];
            if (acceptItem && !acceptItem(candidateItem)) continue;
            var moveEntry = { item: candidateItem, index: i };
            if (isGuideItem(candidateItem)) {
                moveEntries.guide.push(moveEntry);
            } else {
                moveEntries.normal.push(moveEntry);
            }
        }
        return moveEntries;
    }

    /**
     * レイヤー以下の pageItem を再帰的に集める
     * @param {Layer} layer - 探索するレイヤー
     * @param {Array<Object>} collectedEntries - 収集先の配列（{ item: PageItem } を追加）
     * @returns {void}
     */
    function collectPageItemsRecursive(layer, collectedEntries) {
        for (var i = 0; i < layer.layers.length; i++) {
            collectPageItemsRecursive(layer.layers[i], collectedEntries);
        }
        for (var j = 0; j < layer.pageItems.length; j++) {
            var pageItem = layer.pageItems[j];
            if (pageItem.parent !== layer) continue;
            collectedEntries.push({ item: pageItem });
        }
    }

    /**
     * レイヤー内の全 pageItem を別レイヤーへ移動する
     * @param {Layer} sourceLayer - 移動元レイヤー
     * @param {Layer} targetLayer - 移動先レイヤー
     * @returns {{moved: number, failed: number}} 移動できた件数と失敗件数
     */
    function moveAllItemsToLayer(sourceLayer, targetLayer) {
        var collectedEntries = [];
        collectPageItemsRecursive(sourceLayer, collectedEntries);
        return moveEntriesToLayer(collectedEntries, targetLayer, null);
    }

    // =========================================
    // アートボード判定 / Artboard resolution
    // =========================================

    /**
     * オブジェクトの重心がアートボード矩形に含まれるか判定する
     * @param {PageItem} item - 判定するオブジェクト
     * @param {Array<number>} artboardRect - [left, top, right, bottom]
     * @returns {boolean} 含まれていれば true（座標を取得できない場合は false）
     */
    function isCentroidInsideArtboard(item, artboardRect) {
        var itemBounds;
        try {
            itemBounds = item.geometricBounds; // [left, top, right, bottom]
        } catch (e) {
            return false;
        }
        var centroidX = (itemBounds[0] + itemBounds[2]) / 2;
        var centroidY = (itemBounds[1] + itemBounds[3]) / 2;

        return (
            centroidX >= artboardRect[0] &&
            centroidX <= artboardRect[2] &&
            centroidY <= artboardRect[1] &&
            centroidY >= artboardRect[3]
        );
    }

    /**
     * オプション設定に応じたレイヤー名の区切り文字を返す
     * @param {Object} organizeOptions - 処理設定
     * @returns {string} 区切り文字
     */
    function getSeparatorString(organizeOptions) {
        var separatorIndex = SEPARATOR_UNDERSCORE;
        if (organizeOptions && typeof organizeOptions.layerNameSeparatorIndex === "number") {
            separatorIndex = organizeOptions.layerNameSeparatorIndex;
        }
        switch (separatorIndex) {
            case SEPARATOR_HYPHEN: return "-";
            case SEPARATOR_SPACE: return " ";
            case SEPARATOR_NONE: return "";
            default: return "_";
        }
    }

    /**
     * アートボードの表示名を返す（空名のときは代替名）
     * @param {number} artboardIndex - アートボードのインデックス
     * @returns {string} アートボード名
     */
    function getArtboardDisplayName(artboardIndex) {
        var artboardName = artboards[artboardIndex].name;
        if (!artboardName || artboardName === "") {
            return getLabel("fallbackName.artboard");
        }
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
        var artboardName = getArtboardDisplayName(artboardIndex);

        if (!organizeOptions || organizeOptions.includeArtboardNumber !== false) {
            nameParts.push(String(artboardIndex + 1));
        }
        if (!organizeOptions || organizeOptions.includeArtboardName !== false) {
            nameParts.push(artboardName);
        }
        if (nameParts.length === 0) {
            nameParts.push(String(artboardIndex + 1));
            nameParts.push(artboardName);
        }
        return nameParts.join(getSeparatorString(organizeOptions));
    }

    /**
     * 旧仕様レイヤー名（アートボード名のみ）からアートボードのインデックスを引く
     * @param {string} layerName - 判定するレイヤー名
     * @returns {number} 一致したアートボードのインデックス。無ければ -1
     */
    function findLegacyArtboardIndexByLayerName(layerName) {
        for (var i = 0; i < artboards.length; i++) {
            if (getArtboardDisplayName(i) === layerName) {
                return i;
            }
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
            var activeIndex = activeDoc.artboards.getActiveArtboardIndex();
            return { start: activeIndex, end: activeIndex + 1 };
        }
        return { start: 0, end: artboards.length };
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
        for (var i = 0; i < activeDoc.pageItems.length; i++) {
            var pageItem = activeDoc.pageItems[i];
            var parentNode = pageItem.parent;
            if (parentNode.typename !== "Layer" && parentNode.typename !== "Document") continue;

            if (isGuideItem(pageItem)) {
                if (hasExcludedAncestorLayer(pageItem, organizeOptions, true)) continue;
                movableItems.push(pageItem);
                continue;
            }

            if (hasExcludedAncestorLayer(pageItem, organizeOptions, false)) continue;
            if (isObjectExcluded(pageItem, organizeOptions)) continue;
            movableItems.push(pageItem);
        }
        return movableItems;
    }

    /**
     * 通常オブジェクトとガイドをそれぞれの移動先レイヤーへ送る
     * @param {{normal: Array<Object>, guide: Array<Object>}} moveEntries - 振り分け済みの移動エントリ
     * @param {Layer} normalTargetLayer - 通常オブジェクトの移動先レイヤー
     * @param {Array<boolean>} movedFlags - 移動済みフラグ
     * @returns {number} 移動に失敗した件数
     */
    function dispatchMoveEntries(moveEntries, normalTargetLayer, movedFlags) {
        var failedMoves = 0;
        if (moveEntries.normal.length > 0) {
            failedMoves += moveEntriesToLayer(moveEntries.normal, normalTargetLayer, movedFlags).failed;
        }
        if (moveEntries.guide.length > 0) {
            failedMoves += moveEntriesToLayer(moveEntries.guide, getOrCreateGuideLayer(), movedFlags).failed;
        }
        return failedMoves;
    }

    /**
     * 各アートボードに対応するレイヤーへオブジェクトを振り分ける
     * @param {Array<PageItem>} movableItems - 移動候補のオブジェクト配列
     * @param {Object} organizeOptions - 処理設定
     * @param {Array<boolean>} movedFlags - 移動済みフラグ
     * @param {{start: number, end: number}} artboardRange - 処理対象のアートボード範囲
     * @returns {number} 移動に失敗した件数
     */
    function assignItemsToArtboardLayers(movableItems, organizeOptions, movedFlags, artboardRange) {
        var failedMoves = 0;
        for (var artboardIndex = artboardRange.start; artboardIndex < artboardRange.end; artboardIndex++) {
            var artboardLayer = getOrCreateTopLevelLayer(getArtboardLayerName(artboardIndex, organizeOptions));
            var artboardRect = artboards[artboardIndex].artboardRect;
            var moveEntries = buildMoveEntries(movableItems, movedFlags, function (candidateItem) {
                return isCentroidInsideArtboard(candidateItem, artboardRect);
            });
            failedMoves += dispatchMoveEntries(moveEntries, artboardLayer, movedFlags);
        }
        return failedMoves;
    }

    /**
     * どのアートボードにも属さなかったオブジェクトを _pasteboard / _guide へ振り分ける
     * @param {Array<PageItem>} movableItems - 移動候補のオブジェクト配列
     * @param {Object} organizeOptions - 処理設定
     * @param {Array<boolean>} movedFlags - 移動済みフラグ
     * @returns {number} 移動に失敗した件数
     */
    function assignLeftoverItems(movableItems, organizeOptions, movedFlags) {
        var moveEntries = buildMoveEntries(movableItems, movedFlags, null);
        if (moveEntries.normal.length === 0 && moveEntries.guide.length === 0) return 0;
        var pasteboardLayer = null;
        if (moveEntries.normal.length > 0) {
            pasteboardLayer = getOrCreateTopLevelLayer(PASTEBOARD_LAYER_NAME);
        }
        return dispatchMoveEntries(moveEntries, pasteboardLayer, movedFlags);
    }

    /**
     * 対象アートボードのレイヤーを上から 1→2→3… の順に並べる
     * @param {Object} organizeOptions - 処理設定
     * @param {{start: number, end: number}} artboardRange - 処理対象のアートボード範囲
     * @returns {void}
     */
    function reorderArtboardLayers(organizeOptions, artboardRange) {
        for (var i = artboardRange.end - 1; i >= artboardRange.start; i--) {
            bringLayerToFront(getOrCreateTopLevelLayer(getArtboardLayerName(i, organizeOptions)));
        }
    }

    /**
     * 旧仕様（アートボード名のみ）のレイヤーを新仕様レイヤーへ統合する
     * @param {Object} organizeOptions - 処理設定
     * @returns {number} 移動に失敗した件数
     */
    function mergeLegacyLayers(organizeOptions) {
        var failedMoves = 0;
        for (var i = activeDoc.layers.length - 1; i >= 0; i--) {
            var legacyLayer = activeDoc.layers[i];
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
        return withWritableLayer(legacyLayer, function () {
            var failedMoves = moveAllItemsToLayer(legacyLayer, targetLayer).failed;
            if (legacyLayer.pageItems.length === 0 && legacyLayer.layers.length === 0 && activeDoc.layers.length > 1) {
                legacyLayer.remove();
            }
            return failedMoves;
        });
    }

    /**
     * 処理後に残った空レイヤーを削除する（_guide / _pasteboard は保護）
     * @param {Object} organizeOptions - 処理設定
     * @returns {void}
     */
    function cleanupEmptyLayers(organizeOptions) {
        for (var i = activeDoc.layers.length - 1; i >= 0; i--) {
            var topLayer = activeDoc.layers[i];
            if (isLayerExcluded(topLayer, organizeOptions)) continue;
            removeEmptySubLayers(topLayer, organizeOptions);
            if (isProtectedSystemLayer(topLayer)) continue;
            if (topLayer.pageItems.length === 0 && topLayer.layers.length === 0 && activeDoc.layers.length > 1) {
                topLayer.remove();
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
        var movedFlags = [];
        var artboardRange = getTargetArtboardRange(organizeOptions);
        var failedMoves = 0;

        guideLayer = findTopLevelLayerByName(GUIDE_LAYER_NAME); // 既存があれば先に拾う / capture existing if any

        failedMoves += assignItemsToArtboardLayers(movableItems, organizeOptions, movedFlags, artboardRange);
        reorderArtboardLayers(organizeOptions, artboardRange);

        if (!organizeOptions.currentArtboardOnly) {
            failedMoves += assignLeftoverItems(movableItems, organizeOptions, movedFlags);
        }

        bringLayerToFront(guideLayer);

        if (!organizeOptions.currentArtboardOnly) {
            failedMoves += mergeLegacyLayers(organizeOptions);
        }

        reorderArtboardLayers(organizeOptions, artboardRange);
        bringLayerToFront(guideLayer);

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
