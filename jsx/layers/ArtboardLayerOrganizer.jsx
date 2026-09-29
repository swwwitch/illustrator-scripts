#target illustrator
#targetengine "ArtboardLayerOrganizerEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ドキュメント内のオブジェクトをアートボード単位で振り分け、「番号_アートボード名」のレイヤーに整理します。
所属アートボードは各オブジェクトの重心位置で判定します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArtboardLayerOrganizer.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nadb8b8ba49fe

### Overview

Distributes the objects in the document by artboard and organizes them into "number_artboard name" layers.
Each object is assigned to the artboard that contains its centroid.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ArtboardLayerOrganizer.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ArtboardLayerOrganizer";       /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.4.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-04";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArtboardLayerOrganizer.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ArtboardLayerOrganizer.md"; /* README (English) */
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
        layerNameSeparatorIndex: SEPARATOR_UNDERSCORE,
        ignoreLockedLayers: true,
        ignoreLockedObjects: true,
        ignoreHiddenLayers: true,
        ignoreHiddenObjects: true,
        ignoreLockedGuides: false,
        ignoreHiddenGuides: false,
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

    /**
     * パネル共通の見た目をまとめて設定する
     * @param {Panel} targetPanel - 対象パネル
     * @param {string[]} [panelAlignment] - パネル自身の配置（省略時は横も縦も fill）
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
     * 横並びの行グループを追加する
     * @param {Object} parentContainer - 追加先のダイアログ／パネル
     * @param {string} horizontalAlignment - 行自身の横方向の配置（"left" / "fill" など）
     * @param {string} [childVerticalAlignment] - 子の縦方向の揃え（省略時は "center"）
     * @returns {Group} 追加した行グループ
     */
    function addRowGroup(parentContainer, horizontalAlignment, childVerticalAlignment) {
        var rowGroup = parentContainer.add("group");
        rowGroup.orientation = "row";
        rowGroup.alignment = [horizontalAlignment, "top"];
        rowGroup.alignChildren = ["left", childVerticalAlignment || "center"];
        return rowGroup;
    }

    /**
     * 複数のコントロールに同じ tooltip を設定する
     * @param {string} tooltipText - 表示するテキスト
     * @param {Object[]} targetControls - 設定先のコントロール
     * @returns {void}
     */
    function setSharedHelpTip(tooltipText, targetControls) {
        for (var i = 0; i < targetControls.length; i++) {
            targetControls[i].helpTip = tooltipText;
        }
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ローカライズ（再利用パーツ） / Localization (reusable)

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

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

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
            excludeLayer: { ja: "レイヤー", en: "Layers" },
            excludeObject: { ja: "オブジェクト", en: "Objects" },
            excludeGuide: { ja: "ガイド", en: "Guides" },
            removeEmpty: { ja: "空のレイヤー／サブレイヤーを削除", en: "Remove empty layers/sub-layers" }
        },
        dropdown: {
            separatorUnderscore: { ja: "アンダースコア（_）", en: "Underscore (_)" },
            separatorHyphen: { ja: "ハイフン（-）", en: "Hyphen (-)" },
            separatorSpace: { ja: "半角スペース", en: "Space" },
            separatorNone: { ja: "なし", en: "None" }
        },
        fieldLabel: {
            specifiedLayers: { ja: "レイヤー名で指定", en: "Layer names" },
            separator: { ja: "区切り文字", en: "Separator" },
            layerNamePreview: { ja: "例", en: "Example" }
        },
        tooltip: {
            targetArtboards: {
                ja: "アートボードが1つのときは「現在のアートボード」に固定されます。「現在のアートボード」では、どのアートボードにも乗っていないオブジェクトは移動しません",
                en: "Locked to \"Current artboard\" when the document has just one artboard. In that mode, objects outside every artboard are left where they are"
            },
            includeArtboardNumber: {
                ja: "レイヤー名の先頭にアートボードの通し番号（1, 2, 3…）を付けます",
                en: "Prefixes the layer name with the artboard number (1, 2, 3…)"
            },
            includeArtboardName: {
                ja: "レイヤー名にアートボード名を含めます。名前が空のときは「アートボード」を使います",
                en: "Includes the artboard name; falls back to \"Artboard\" when it is empty"
            },
            separator: {
                ja: "アートボード番号とアートボード名の間に入れる文字。「なし」を選ぶと直接つなぎます",
                en: "Character inserted between the artboard number and the artboard name; \"None\" joins them directly"
            },
            exclusionPanel: {
                ja: "ガイドは「レイヤー名で指定」に関係なく _guide レイヤーに集めます（対象外にしたロック中／非表示のレイヤー内にあるものは除く）",
                en: "Guides go to the _guide layer even inside layers listed in \"Layer names\" (except those inside excluded locked/hidden layers)"
            },
            lockedExclusion: {
                ja: "チェックしたものはロック中なら整理対象から除外します。外したものはロックを一時解除して移動し、処理後にロックし直します",
                en: "Checked items are left out when locked. Unchecked ones are unlocked for the move and locked again afterwards"
            },
            hiddenExclusion: {
                ja: "チェックしたものは非表示なら整理対象から除外します。外したものは一時的に表示して移動し、処理後に非表示に戻します",
                en: "Checked items are left out when hidden. Unchecked ones are shown for the move and hidden again afterwards"
            },
            specifiedLayers: {
                ja: "対象外にするレイヤー名をカンマ（,）または読点（、）で区切って指定。サブレイヤーも名前で一致します（例：bg, temp）",
                en: "Names of layers to exclude, separated by commas. Sub-layers match by name too (e.g. bg, temp)"
            },
            removeEmpty: {
                ja: "整理後に空になったレイヤー／サブレイヤーを削除します（_guide と _pasteboard は削除しません）",
                en: "Removes layers/sub-layers left empty after organizing (_guide and _pasteboard are kept)"
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

    // =========================================
    // 前提チェック / Preconditions
    // =========================================

    if (app.documents.length === 0) {
        alert(getLabel("alert.noDocument"));
        return;
    }

    var documentRef = app.activeDocument;
    var documentArtboards = documentRef.artboards;
    var guideLayer = null; // _guide レイヤー参照（必要時に取得・作成） / _guide layer reference, resolved on demand

    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)

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

    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity

    // ボタン行（再利用パーツ） / Button row (reusable)

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

    // ボタン行（再利用パーツ）ここまで / End of the reusable button row

    // =========================================
    // ダイアログ / Dialog
    // =========================================

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

        var buttonRow = addButtonRow(optionsDialog, { centered: true });
        var btnCancel = buttonRow.rowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        optionsDialog.defaultElement = btnOK;
        optionsDialog.cancelElement = btnCancel;

        prepareDialogWindow(optionsDialog, SCRIPT_NAME);
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
            ignoreLockedGuides: exclusionControls.lockedGuideCheckbox.value,
            ignoreHiddenGuides: exclusionControls.hiddenGuideCheckbox.value,
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

        var artboardScopeRow = addRowGroup(targetArtboardPanel, "center");
        var currentArtboardOnlyRadio = artboardScopeRow.add("radiobutton", undefined, getLabel("radio.currentArtboardOnly"));
        var allArtboardsRadio = artboardScopeRow.add("radiobutton", undefined, getLabel("radio.allArtboards"));
        setSharedHelpTip(getLabel("tooltip.targetArtboards"), [targetArtboardPanel, currentArtboardOnlyRadio, allArtboardsRadio]);

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

        var separatorRow = addRowGroup(layerNamePanel, "left");
        var separatorLabel = separatorRow.add("statictext", undefined, labelText("fieldLabel.separator"));
        var separatorDropdown = separatorRow.add("dropdownlist", undefined, [
            getLabel("dropdown.separatorUnderscore"),
            getLabel("dropdown.separatorHyphen"),
            getLabel("dropdown.separatorSpace"),
            getLabel("dropdown.separatorNone")
        ]);
        separatorDropdown.selection = DEFAULT_OPTIONS.layerNameSeparatorIndex;
        setSharedHelpTip(getLabel("tooltip.separator"), [separatorRow, separatorLabel, separatorDropdown]);

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
            return separatorDropdown.selection ? separatorDropdown.selection.index : SEPARATOR_NONE;
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
            return labelText("fieldLabel.layerNamePreview") + sampleLayerName;
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
            layerNamePreviewText.text = buildLayerNamePreview();
        }

        artboardNumberCheckbox.onClick = function () { updateLayerNameControls(artboardNameCheckbox); };
        artboardNameCheckbox.onClick = function () { updateLayerNameControls(artboardNumberCheckbox); };
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

        var lockHiddenRow = addRowGroup(exclusionPanel, "left", "fill");
        lockHiddenRow.spacing = EXCLUSION_SUBPANEL_SPACING;

        var lockedControls = buildExclusionSubPanel(lockHiddenRow, "panel.locked", "tooltip.lockedExclusion", {
            layer: DEFAULT_OPTIONS.ignoreLockedLayers,
            object: DEFAULT_OPTIONS.ignoreLockedObjects,
            guide: DEFAULT_OPTIONS.ignoreLockedGuides
        });
        var hiddenControls = buildExclusionSubPanel(lockHiddenRow, "panel.hidden", "tooltip.hiddenExclusion", {
            layer: DEFAULT_OPTIONS.ignoreHiddenLayers,
            object: DEFAULT_OPTIONS.ignoreHiddenObjects,
            guide: DEFAULT_OPTIONS.ignoreHiddenGuides
        });

        var excludedNamesRow = addRowGroup(exclusionPanel, "fill");
        var excludedNamesLabel = excludedNamesRow.add("statictext", undefined, labelText("fieldLabel.specifiedLayers"));
        var excludedNamesField = excludedNamesRow.add("edittext", undefined, DEFAULT_OPTIONS.excludedLayerNames);
        excludedNamesField.alignment = ["fill", "center"];
        setSharedHelpTip(getLabel("tooltip.specifiedLayers"), [excludedNamesRow, excludedNamesLabel, excludedNamesField]);

        return {
            lockedLayerCheckbox: lockedControls.layerCheckbox,
            lockedObjectCheckbox: lockedControls.objectCheckbox,
            hiddenLayerCheckbox: hiddenControls.layerCheckbox,
            hiddenObjectCheckbox: hiddenControls.objectCheckbox,
            lockedGuideCheckbox: lockedControls.guideCheckbox,
            hiddenGuideCheckbox: hiddenControls.guideCheckbox,
            excludedNamesField: excludedNamesField
        };
    }

    /**
     * ロック／非表示のサブパネル（レイヤー・オブジェクト・ガイド）を構築する
     * @param {Group} parentGroup - 追加先のグループ
     * @param {string} titleLabelPath - パネルタイトルの LABELS パス
     * @param {string} tooltipLabelPath - パネルと各チェックボックスに設定する tooltip の LABELS パス
     * @param {{layer: boolean, object: boolean, guide: boolean}} defaultValues - 各チェックボックスの初期値
     * @returns {Object} チェックボックス参照をまとめたオブジェクト
     */
    function buildExclusionSubPanel(parentGroup, titleLabelPath, tooltipLabelPath, defaultValues) {
        var exclusionSubPanel = parentGroup.add("panel", undefined, getLabel(titleLabelPath));
        applyPanelLayout(exclusionSubPanel, ["left", "fill"]);

        var layerCheckbox = exclusionSubPanel.add("checkbox", undefined, getLabel("checkbox.excludeLayer"));
        layerCheckbox.value = defaultValues.layer;
        var objectCheckbox = exclusionSubPanel.add("checkbox", undefined, getLabel("checkbox.excludeObject"));
        objectCheckbox.value = defaultValues.object;
        var guideCheckbox = exclusionSubPanel.add("checkbox", undefined, getLabel("checkbox.excludeGuide"));
        guideCheckbox.value = defaultValues.guide;
        setSharedHelpTip(getLabel(tooltipLabelPath), [exclusionSubPanel, layerCheckbox, objectCheckbox, guideCheckbox]);

        return { layerCheckbox: layerCheckbox, objectCheckbox: objectCheckbox, guideCheckbox: guideCheckbox };
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

    // =========================================
    // 除外判定 / Exclusion rules
    // =========================================

    /**
     * "bg, temp" のような文字列をレイヤー名の配列に分解する（, または 、 区切り）
     * @param {string} excludedNamesText - 指定レイヤー欄の入力値
     * @returns {string[]} 前後の空白を除いたレイヤー名の配列
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
     * @param {string[]} excludedLayerNames - 除外レイヤー名の配列
     * @returns {boolean} 含まれていれば true
     */
    function isNameInExcludedList(layerName, excludedLayerNames) {
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
        return isNameInExcludedList(targetLayer.name, organizeOptions.excludedLayerNames) ||
            isLayerExcludedByState(targetLayer, organizeOptions);
    }

    /**
     * レイヤーがロック／非表示の状態によって除外対象か判定する
     * @param {Layer} targetLayer - 判定するレイヤー
     * @param {Object} organizeOptions - 処理設定
     * @returns {boolean} 除外対象なら true
     */
    function isLayerExcludedByState(targetLayer, organizeOptions) {
        return (organizeOptions.ignoreLockedLayers && targetLayer.locked) ||
            (organizeOptions.ignoreHiddenLayers && !targetLayer.visible);
    }

    /**
     * 親方向に辿って除外対象のレイヤーがあるか判定する
     * @param {PageItem} targetItem - 判定するオブジェクト
     * @param {Object} organizeOptions - 処理設定
     * @param {boolean} skipNameExclusion - true のとき指定レイヤー名による除外を無視し、ロック／非表示だけを見る
     * @returns {boolean} 除外対象の祖先レイヤーがあれば true
     */
    function hasExcludedAncestorLayer(targetItem, organizeOptions, skipNameExclusion) {
        var isAncestorExcluded = skipNameExclusion ? isLayerExcludedByState : isLayerExcluded;
        var ancestorNode = targetItem.parent;
        while (ancestorNode && ancestorNode.typename !== "Document") {
            if (ancestorNode.typename === "Layer" && isAncestorExcluded(ancestorNode, organizeOptions)) return true;
            ancestorNode = ancestorNode.parent;
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
     * ガイドはオブジェクトとは別の設定で判定する
     * @param {PageItem} targetItem - 判定するオブジェクト
     * @param {Object} organizeOptions - 処理設定
     * @returns {boolean} 除外対象なら true
     */
    function isObjectExcluded(targetItem, organizeOptions) {
        var isGuide = isGuideItem(targetItem);
        var ignoreLocked = isGuide ? organizeOptions.ignoreLockedGuides : organizeOptions.ignoreLockedObjects;
        var ignoreHidden = isGuide ? organizeOptions.ignoreHiddenGuides : organizeOptions.ignoreHiddenObjects;
        if (ignoreLocked && readItemFlag(targetItem, "locked")) return true;
        if (ignoreHidden && readItemFlag(targetItem, "hidden")) return true;
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
        var topLevelLayers = documentRef.layers;
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
        var createdLayer = documentRef.layers.add();
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
     * @returns {Layer[]} トップレベルレイヤーの配列
     */
    function getTopLevelLayerSnapshot() {
        var topLevelLayers = [];
        for (var i = 0; i < documentRef.layers.length; i++) {
            topLevelLayers.push(documentRef.layers[i]);
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
     * 「対象外にする」に該当するオブジェクトは移動候補に含まれないため、移動候補のロック／非表示はすべて解除してよい
     * @param {PageItem[]} movableItems - 移動候補のオブジェクト配列
     * @param {Object} organizeOptions - 処理設定
     * @returns {Object[]} 復元用の記録（上位レイヤー→下位レイヤー→オブジェクトの順）
     */
    function suspendLockAndHidden(movableItems, organizeOptions) {
        var suspendedEntries = [];
        var unlockLayers = !organizeOptions.ignoreLockedLayers;
        var showLayers = !organizeOptions.ignoreHiddenLayers;
        if (unlockLayers || showLayers) {
            suspendLayerLockAndHidden(documentRef, unlockLayers, showLayers, suspendedEntries);
        }
        suspendItemLockAndHidden(movableItems, suspendedEntries);
        return suspendedEntries;
    }

    /**
     * レイヤーツリーを上位から辿ってロック／非表示を解除する
     * 親を先に解除しないと子の解除が効かないため、必ず上位から処理する
     * @param {Document|Layer} layerContainer - 探索するドキュメントまたはレイヤー
     * @param {boolean} unlockLocked - ロックを解除するなら true
     * @param {boolean} showHidden - 非表示を表示にするなら true
     * @param {Object[]} suspendedEntries - 記録の追加先
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
     * @param {PageItem[]} movableItems - 移動候補のオブジェクト配列
     * @param {Object[]} suspendedEntries - 記録の追加先
     * @returns {void}
     */
    function suspendItemLockAndHidden(movableItems, suspendedEntries) {
        for (var i = 0; i < movableItems.length; i++) {
            var targetItem = movableItems[i];
            var wasLocked = readItemFlag(targetItem, "locked");
            var wasHidden = readItemFlag(targetItem, "hidden");
            if (!wasLocked && !wasHidden) continue;
            if (wasLocked) targetItem.locked = false;
            if (wasHidden) targetItem.hidden = false;
            suspendedEntries.push({ node: targetItem, isLayer: false, wasLocked: wasLocked, wasHidden: wasHidden });
        }
    }

    /**
     * suspendLockAndHidden で解除したロック／非表示を元へ戻す
     * 子より親を後に戻す必要があるため、記録の逆順で処理する
     * @param {Object[]} suspendedEntries - 復元する記録の配列
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
     * @param {Object[]} moveEntries - { item: PageItem, index: number } の配列
     * @param {Layer} targetLayer - 移動先レイヤー
     * @param {boolean[]} handledFlags - 処理済みフラグ（不要なら null）
     * @returns {number} 移動できなかった件数
     */
    function moveEntriesToLayer(moveEntries, targetLayer, handledFlags) {
        if (moveEntries.length === 0) return 0;
        return withWritableLayer(targetLayer, function () {
            var failedMoveCount = 0;
            var i = moveEntries.length;
            while (i--) {
                try {
                    moveEntries[i].item.move(targetLayer, ElementPlacement.PLACEATBEGINNING);
                } catch (e) {
                    failedMoveCount++;
                }
                if (handledFlags && typeof moveEntries[i].index === "number") {
                    handledFlags[moveEntries[i].index] = true;
                }
            }
            return failedMoveCount;
        });
    }

    /**
     * 未処理のアイテムをガイドとそれ以外に振り分け、移動エントリを作る
     * @param {PageItem[]} movableItems - 移動候補のオブジェクト配列
     * @param {boolean[]} handledFlags - 処理済みフラグ
     * @param {function(number):boolean} shouldMoveItem - 対象に含めるか判定する関数（null なら未処理すべて）
     * @returns {{normal: Object[], guide: Object[]}} 通常オブジェクトとガイドの移動エントリ
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
     * @param {{normal: Object[], guide: Object[]}} moveEntries - 振り分け済みの移動エントリ
     * @param {Layer} normalTargetLayer - 通常オブジェクトの移動先レイヤー
     * @param {boolean[]} handledFlags - 処理済みフラグ
     * @returns {number} 移動に失敗した件数
     */
    function moveEntriesToTargets(moveEntries, normalTargetLayer, handledFlags) {
        var failedMoveCount = moveEntriesToLayer(moveEntries.normal, normalTargetLayer, handledFlags);
        if (moveEntries.guide.length > 0) {
            failedMoveCount += moveEntriesToLayer(moveEntries.guide, getOrCreateGuideLayer(), handledFlags);
        }
        return failedMoveCount;
    }

    /**
     * レイヤー以下の pageItem を再帰的に集める
     * layer.pageItems はサブレイヤーの中身を含まないため、サブレイヤーは個別に辿る
     * @param {Layer} targetLayer - 探索するレイヤー
     * @param {Object[]} collectedEntries - 収集先の配列（{ item: PageItem } を追加）
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
     * @param {PageItem[]} movableItems - 移動候補のオブジェクト配列
     * @returns {number[][]} [x, y] の配列（座標を取得できなかった要素は null）
     */
    function buildItemCentroids(movableItems) {
        var itemCentroids = [];
        for (var i = 0; i < movableItems.length; i++) {
            /* 中身の無いオブジェクトなどは geometricBounds で例外になる / geometricBounds can throw on empty items */
            try {
                var itemBounds = movableItems[i].geometricBounds; // [left, top, right, bottom]
                itemCentroids.push([(itemBounds[0] + itemBounds[2]) / 2, (itemBounds[1] + itemBounds[3]) / 2]);
            } catch (e) {
                itemCentroids.push(null);
            }
        }
        return itemCentroids;
    }

    /**
     * 重心がアートボード矩形に含まれるか判定する
     * @param {number[]} itemCentroid - [x, y]（取得できなかった場合は null）
     * @param {number[]} artboardRect - [left, top, right, bottom]
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
     * @param {number[][]} itemCentroids - 重心座標の配列
     * @param {number[]} artboardRect - [left, top, right, bottom]
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
    function getSeparatorCharacter(organizeOptions) {
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
        return nameParts.join(getSeparatorCharacter(organizeOptions));
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
     * ガイドは指定レイヤー名による除外を無視する（ロック／非表示の祖先とガイド自身の設定は尊重）
     * @param {Object} organizeOptions - 処理設定
     * @returns {PageItem[]} 移動候補のオブジェクト配列
     */
    function collectMovableItems(organizeOptions) {
        var movableItems = [];
        var allPageItems = documentRef.pageItems;
        var pageItemCount = allPageItems.length;
        for (var i = 0; i < pageItemCount; i++) {
            var pageItem = allPageItems[i];
            var itemParentType = pageItem.parent.typename;
            if (itemParentType !== "Layer" && itemParentType !== "Document") continue;

            /* ガイドは指定レイヤー名による除外を受けない / Guides ignore the name-based exclusion */
            if (hasExcludedAncestorLayer(pageItem, organizeOptions, isGuideItem(pageItem))) continue;
            if (isObjectExcluded(pageItem, organizeOptions)) continue;
            movableItems.push(pageItem);
        }
        return movableItems;
    }

    /**
     * 各アートボードに対応するレイヤーへオブジェクトを振り分ける
     * @param {PageItem[]} movableItems - 移動候補のオブジェクト配列
     * @param {number[][]} itemCentroids - 重心座標の配列
     * @param {Object} organizeOptions - 処理設定
     * @param {boolean[]} handledFlags - 処理済みフラグ
     * @param {{start: number, end: number}} artboardRange - 処理対象のアートボード範囲
     * @returns {number} 移動に失敗した件数
     */
    function assignItemsToArtboardLayers(movableItems, itemCentroids, organizeOptions, handledFlags, artboardRange) {
        var failedMoveCount = 0;
        for (var artboardIndex = artboardRange.start; artboardIndex < artboardRange.end; artboardIndex++) {
            var artboardLayer = getOrCreateTopLevelLayer(getArtboardLayerName(artboardIndex, organizeOptions));
            var centroidFilter = buildCentroidFilter(itemCentroids, documentArtboards[artboardIndex].artboardRect);
            var moveEntries = collectMoveEntries(movableItems, handledFlags, centroidFilter);
            failedMoveCount += moveEntriesToTargets(moveEntries, artboardLayer, handledFlags);
        }
        return failedMoveCount;
    }

    /**
     * どのアートボードにも属さなかったオブジェクトを _pasteboard / _guide へ振り分ける
     * @param {PageItem[]} movableItems - 移動候補のオブジェクト配列
     * @param {boolean[]} handledFlags - 処理済みフラグ
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
        var failedMoveCount = 0;
        var topLevelLayers = getTopLevelLayerSnapshot();
        for (var i = topLevelLayers.length - 1; i >= 0; i--) {
            var legacyLayer = topLevelLayers[i];
            if (isProtectedSystemLayer(legacyLayer)) continue;
            if (isLayerExcluded(legacyLayer, organizeOptions)) continue;
            var legacyArtboardIndex = findLegacyArtboardIndexByLayerName(legacyLayer.name);
            if (legacyArtboardIndex < 0) continue;

            var targetLayer = getOrCreateTopLevelLayer(getArtboardLayerName(legacyArtboardIndex, organizeOptions));
            if (legacyLayer === targetLayer) continue;

            failedMoveCount += mergeSingleLegacyLayer(legacyLayer, targetLayer);
        }
        return failedMoveCount;
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
        var failedMoveCount = withWritableLayer(legacyLayer, function () {
            return moveEntriesToLayer(collectedEntries, targetLayer, null);
        });
        if (isLayerEmpty(legacyLayer) && documentRef.layers.length > 1) {
            removeLayerSafely(legacyLayer);
        }
        return failedMoveCount;
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
            if (isLayerEmpty(topLevelLayer) && documentRef.layers.length > 1) {
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
        var failedMoveCount = 0;

        guideLayer = findTopLevelLayerByName(GUIDE_LAYER_NAME); // 既存があれば先に拾う / capture existing if any

        var suspendedEntries = suspendLockAndHidden(movableItems, organizeOptions);
        try {
            failedMoveCount += assignItemsToArtboardLayers(movableItems, itemCentroids, organizeOptions, handledFlags, artboardRange);
            if (!organizeOptions.currentArtboardOnly) {
                failedMoveCount += assignLeftoverItems(movableItems, handledFlags);
                failedMoveCount += mergeLegacyLayers(organizeOptions);
            }
            applyLayerOrder(organizeOptions, artboardRange);
        } finally {
            restoreLockAndHidden(suspendedEntries);
        }

        if (organizeOptions.removeEmptyLayers) {
            cleanupEmptyLayers(organizeOptions);
        }
        return failedMoveCount;
    }

    var organizeOptions = showOptionsDialog();
    if (!organizeOptions) {
        return;
    }

    var failedMoveCount = organizeDocumentLayers(organizeOptions);
    if (failedMoveCount > 0) {
        alert(getLabel("alert.moveFailed").replace("{count}", String(failedMoveCount)));
    }

})();
