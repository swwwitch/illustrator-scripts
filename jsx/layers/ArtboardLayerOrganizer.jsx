#target illustrator
#targetengine "ArtboardLayerOrganizerEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.ArtboardLayerOrganizer
Meta-Default-language=ja
Name=アートボード単位でオブジェクトをレイヤーに整理
Name[en]=Sort objects into layers by artboard
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.5.3
Release-Date=2026-04-04
Tags=layer,レイヤー
X-com.swwwitch-Script-Name=ArtboardLayerOrganizer
X-com.swwwitch-Updated=2026-11-20
X-com.swwwitch-README-JA=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArtboardLayerOrganizer.md
X-com.swwwitch-README-EN=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ArtboardLayerOrganizer.md
X-com.swwwitch-Article-URL=https://note.com/dtp_tranist/n/nadb8b8ba49fe
X-com.swwwitch-Source-URL=https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/layers/ArtboardLayerOrganizer.jsx
Description-BEGIN
[ja]
ドキュメント内のオブジェクトをアートボード単位で振り分け、「番号_アートボード名」のレイヤーに整理します。
所属アートボードは各オブジェクトの重心位置で判定します。
[en]
Distributes the objects in the document by artboard and organizes them into "number_artboard name" layers.
Each object is assigned to the artboard that contains its centroid.
Description-END
SCRIPTMETA-END
*/

// =========================================
// 基本情報 / Basic info
// =========================================
// スクリプト情報（再利用パーツ） / Script info (reusable)
/* 基本情報は先頭の SCRIPTMETA が元。実行時に読み、getScriptMetaValue("Version") のように直接読む（値は SCRIPTMETA を書き換え、helpers/sync-scriptmeta.py で同期）
   The basic info lives in the SCRIPTMETA block at the top; read its fields directly with getScriptMetaValue() */

/**
 * このスクリプトの先頭にある SCRIPTMETA の項目（Description の手前まで）を読む
 * @returns {Object} 項目名 → 値（同じ項目は最初の値）。ファイルが読めなければ空のオブジェクト
 */
function readScriptMeta() {
    var metaFields = {};
    var scriptFile = new File($.fileName);
    scriptFile.encoding = "UTF-8";
    if (!scriptFile.exists || !scriptFile.open("r")) return metaFields;
    /* 境界の文字列は分けて書く（コードの中に境界があると SCRIPTMETA を読むアプリが誤読するため）
       Marker strings are split so apps that scan for SCRIPTMETA do not find them in the code */
    var blockBegin = "SCRIPTMETA" + "-BEGIN";
    var blockEnd = "SCRIPTMETA" + "-END";
    var descriptionBegin = "Description" + "-BEGIN";
    var isInBlock = false;
    for (var lineCount = 0; lineCount < 200 && !scriptFile.eof; lineCount++) {
        var metaLine = scriptFile.readln().replace(/\r$/, "");
        if (metaLine === blockBegin) {
            isInBlock = true;
        } else if (isInBlock) {
            if (metaLine === blockEnd || metaLine === descriptionBegin) break;
            var separatorIndex = metaLine.indexOf("=");
            var fieldKey = metaLine.substring(0, separatorIndex);
            if (separatorIndex > 0 && !metaFields.hasOwnProperty(fieldKey)) metaFields[fieldKey] = metaLine.substring(separatorIndex + 1);
        }
    }
    scriptFile.close();
    return metaFields;
}

var SCRIPT_META = readScriptMeta();

/**
 * SCRIPTMETA の項目の値を返す
 * @param {string} fieldKey - 項目名（"Version"、"X-com.swwwitch-Updated" など）
 * @returns {string} 値。項目が無ければ空文字
 */
function getScriptMetaValue(fieldKey) {
    return SCRIPT_META.hasOwnProperty(fieldKey) ? SCRIPT_META[fieldKey] : "";
}

/* SCRIPTMETA が読めないとき（ファイルを経由しない実行）は、スクリプト名だけファイル名から補う / fall back to the file name */
var SCRIPT_NAME = getScriptMetaValue("X-com.swwwitch-Script-Name") || decodeURI(new File($.fileName).name).replace(/\.[^.]+$/, ""); /* スクリプト名（設定の保存名） / script name, the key for saved settings */
// スクリプト情報（再利用パーツ）ここまで / End of the reusable script info

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* システム管理レイヤー名（空でも削除しない） / System-managed layer names (never deleted) */
    var GUIDE_LAYER_NAME = "_guide";
    var PASTEBOARD_LAYER_NAME = "_pasteboard";

    /* ルーラーガイドを線とみなす厚みの許容値（pt） / Thickness tolerance for treating a ruler guide as a line (pt) */
    var BOUNDS_TOLERANCE = 0.001;

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
        collectGuidesToGuideLayer: true,   // false: ガイドも各アートボードのレイヤーへ / false: guides go to artboard layers too
        excludedLayerNames: "bg"    // , または 、 区切り / Comma-separated
    };

    // =========================================
    // レイアウト設定 / Layout Settings
    // =========================================

    // UIレイアウト（再利用パーツ） / UI layout (reusable)

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
    var TAB_MARGINS    = [15, 20, 5, 10];    /* タブ余白 [左,上,右,下] / tab margins */

    /**
     * ウィンドウの共通設定
     * @param {Window} targetWindow - 対象のウィンドウ
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(targetWindow, spacing) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルの共通設定（子は幅いっぱい。ボタンは alignment = "left" で広げない）
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * タブの共通設定
     * @param {Tab} targetTab - 対象のタブ
     * @param {number} [spacing] - 要素間隔（省略時は変えない）
     * @returns {void}
     */
    function setupTab(targetTab, spacing) {
        targetTab.orientation = "column";
        targetTab.alignChildren = "fill";
        targetTab.margins = TAB_MARGINS;
        if (typeof spacing === "number") targetTab.spacing = spacing;
    }

    /**
     * 横並びの行グループの共通設定（ボタン列など）。
     * alignment と alignChildren を対で指定し、中のボタンが横に伸びたり天地がずれたりしないようにする
     * @param {Group} rowGroup - 対象のグループ
     * @param {string|string[]} [rowAlignment] - 横方向の alignment（省略時は "left"）。配列ならそのまま使う
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(rowGroup, rowAlignment, spacing) {
        rowGroup.orientation = "row";
        rowGroup.alignment = (rowAlignment instanceof Array) ? rowAlignment : [rowAlignment || "left", "center"];
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ボタンの高さを指定した px だけ詰める（レイアウトが決まったあとに呼ぶ）
     * @param {Button} targetButton - 対象のボタン
     * @param {number} trimPixels - 詰める量（px）
     * @returns {void}
     */
    function trimButtonHeight(targetButton, trimPixels) {
        /* レイアウト前は size が無い / size is not set until the layout runs */
        if (!targetButton.size) return;
        targetButton.size = [targetButton.size.width, targetButton.size.height - trimPixels];
    }

    // UIレイアウト（再利用パーツ）ここまで / End of the reusable UI layout

    /* レイヤー名プレビューの最大文字数（ダイアログ幅を広げないための上限） / Preview length cap, keeps the dialog from widening */
    var LAYER_NAME_PREVIEW_MAX_LENGTH = 16;

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
     * 項目名の文言の末尾にコロンを付ける（日本語は半角スペース＋半角コロン「 :」、英語は「:」。Illustrator の線パネルなどの項目名に合わせる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? " :" : ":");
    }

    /**
     * 「項目名 : 値」の1行を返す（日本語は「件数 : 5」、英語は「Count: 5」。どちらもコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + " " + value;
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
            guideDestination: { ja: "ガイドの移動先", en: "Guide Destination" },
            exclusion: { ja: "対象外にする", en: "Exclude" },
            locked: { ja: "ロック中", en: "Locked" },
            hidden: { ja: "非表示", en: "Hidden" },
            postProcess: { ja: "整理後の処理", en: "After Organizing" }
        },
        radio: {
            currentArtboardOnly: { ja: "現在のアートボード", en: "Current artboard" },
            allArtboards: { ja: "すべて", en: "All" },
            guideToGuideLayer: { ja: "_guide レイヤー", en: "_guide layer" },
            guideToArtboardLayers: { ja: "各レイヤー", en: "Each layer" }
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
                ja: "レイヤー名の先頭にアートボードの通し番号（1, 2, 3...）を付けます",
                en: "Prefixes the layer name with the artboard number (1, 2, 3...)"
            },
            includeArtboardName: {
                ja: "レイヤー名にアートボード名を含めます。名前が空のときは「アートボード」を使います",
                en: "Includes the artboard name; falls back to \"Artboard\" when it is empty"
            },
            separator: {
                ja: "アートボード番号とアートボード名の間に入れる文字。「なし」を選ぶと直接つなぎます",
                en: "Character inserted between the artboard number and the artboard name; \"None\" joins them directly"
            },
            guideDestination: {
                ja: "「_guide レイヤー」ではガイドをすべて _guide に集めます。「各レイヤー」ではほかのオブジェクトと同じく各アートボードのレイヤーへ移します（_guide にあるものも含む）。ルーラーから引いたガイドは、横位置（垂直ガイドは縦位置）が重なるアートボードが1つだけのときに移し、それ以外は _guide に残します",
                en: "\"_guide layer\" collects every guide on the _guide layer. \"Each layer\" moves guides to the artboard layers like other objects, including those already on _guide. A ruler guide moves only when exactly one artboard spans its position; otherwise it stays on _guide"
            },
            exclusionPanel: {
                ja: "ガイドは「レイヤー名で指定」に関係なく移動します（対象外にしたロック中／非表示のレイヤー内にあるものは除く）",
                en: "Guides are moved even inside layers listed in \"Layer names\" (except those inside excluded locked/hidden layers)"
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
            /* 文字ツールで文字を選択しているときは TextRange が返り、[0] が無い / Selecting characters with the Type tool returns a TextRange, which has no [0] */
            if (!selectedItems || selectedItems.typename === "TextRange" || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
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
    var BUTTON_ROW_BOTTOM_MARGIN = 14; /* ボタン行の下の余白。ダイアログの下余白と合わせて約30px（Illustrator 標準のダイアログに合わせる） / bottom margin; with the dialog margin about 30px, like Illustrator's own dialogs */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
    var BUTTON_ROW_CENTER_MAX_WIDTH = 240; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

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
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, BUTTON_ROW_BOTTOM_MARGIN];
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

    /**
     * 左のグループにボタンが無い（右のボタンだけの）行を、ダイアログの幅に合わせて揃える。
     * 内側の幅（左右の余白を除く）が BUTTON_ROW_CENTER_MAX_WIDTH 以下なら左右中央、それより広ければ右揃えのまま。
     * 幅はレイアウトが決まるまで分からないので、ダイアログを表示した時点（show イベント）で判定する。
     * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
     * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
     * @returns {void}
     */
    function alignRightOnlyButtonRow(buttonRow) {
        if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
        var dialogWindow = buttonRow.rowGroup.window;
        dialogWindow.addEventListener("show", function () {
            if (!buttonRow.leftGroup) return;
            var btnRowGroup = buttonRow.rowGroup;
            /* 行の幅＝ダイアログの内側の幅（左右の余白を除く）/ The row spans the dialog's inner width (margins excluded) */
            if (!btnRowGroup.size || btnRowGroup.size.width > BUTTON_ROW_CENTER_MAX_WIDTH) return;
            /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
            btnRowGroup.remove(buttonRow.leftGroup);
            btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            buttonRow.leftGroup = null;
            dialogWindow.layout.layout(true);
        });
    }

    // ボタン行（再利用パーツ）ここまで / End of the reusable button row

    // UI の明暗（再利用パーツ） / UI theme (reusable)

    /**
     * UI がダークテーマかどうかを判定する（Illustrator は uiBrightness、InDesign は uiBrightnessPreference）
     * @returns {boolean} ダークなら true。取得できない環境では false（明るいUI扱い）
     */
    function isDarkUI() {
        try {
            if (app.preferences && app.preferences.getRealPreference) {
                return app.preferences.getRealPreference("uiBrightness") <= 0.5; /* Illustrator */
            }
            return app.generalPreferences.uiBrightnessPreference <= 0.5; /* InDesign */
        } catch (e) {
            return false;
        }
    }

    // UI の明暗（再利用パーツ）ここまで / End of the reusable UI theme

    // バージョン表示（再利用パーツ） / Version footer (reusable)

    var VERSION_FOOTER_TOP_MARGIN = 4;    /* バージョンの行の上に足す余白（px。ウィンドウの間隔に足す）/ extra space above the version row, added to the window spacing */
    var VERSION_FOOTER_BOTTOM_MARGIN = 8; /* バージョンの行の下の余白（px。ウィンドウの下余白をこの値にする）/ space below the version row (replaces the window's bottom margin) */
    var VERSION_FOOTER_PIXEL_SIZE = 1;    /* ドット1つの大きさ（px）/ size of one dot (px) */
    var VERSION_FOOTER_LETTER_GAP = 1;    /* 字の間（ドット）/ gap between glyphs (dots) */
    var VERSION_FOOTER_BRAND = "DTP Transit"; /* 右下に出す名前（字形にある文字だけ）/ name shown at the bottom right (glyph characters only) */
    var VERSION_FOOTER_COLOR = isDarkUI() ? [0.45, 0.45, 0.45, 1] : [0.62, 0.62, 0.62, 1]; /* 文字の色（ディム表示に近い灰色）/ text color, a gray close to the dimmed look */

    /* 字形：Sigi 7px Regular（Rasmus Andersson、CC0。https://github.com/rsms/sigi-pixel-font）の大文字・数字・「. - _ +」・空白を
       9 行のドット（# が点。上7行が大文字の高さ、下2行は小文字の下に出る部分）に写したもの。Sigi に小文字が無いので、a〜z は同じ大きさで描き足した。I は l と見分けるよう上下に横棒を足した
       / Glyphs: uppercase, digits, ". - _ +" and space from Sigi 7px Regular by Rasmus Andersson (CC0, https://github.com/rsms/sigi-pixel-font),
         nine rows each ("#" is a dot; seven rows of cap height plus two for descenders). Sigi has no lowercase, so a-z were drawn to match. I has top and bottom bars to tell it from l */
    var VERSION_FOOTER_GLYPH_HEIGHT = 9;  /* 字の高さ（ドット。下2行は g・j・p・q・y の下に出る部分）/ glyph height (dots), including two descender rows */
    var VERSION_FOOTER_GLYPHS = {
        "A": [".###.", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
        "B": ["####.", "#...#", "#...#", "####.", "#...#", "#...#", "####.", ".....", "....."],
        "C": [".###.", "#...#", "#....", "#....", "#....", "#...#", ".###.", ".....", "....."],
        "D": ["####.", "#...#", "#...#", "#...#", "#...#", "#...#", "####.", ".....", "....."],
        "E": ["#####", "#....", "#....", "####.", "#....", "#....", "#####", ".....", "....."],
        "F": ["#####", "#....", "#....", "####.", "#....", "#....", "#....", ".....", "....."],
        "G": [".###.", "#...#", "#....", "#..##", "#...#", "#...#", ".###.", ".....", "....."],
        "H": ["#...#", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
        "I": ["###", ".#.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
        "J": ["....#", "....#", "....#", "....#", "....#", "#...#", ".###.", ".....", "....."],
        "K": ["#...#", "#..#.", "#.#..", "##...", "#.#..", "#..#.", "#...#", ".....", "....."],
        "L": ["#...", "#...", "#...", "#...", "#...", "#...", "####", "....", "...."],
        "M": ["#...#", "##.##", "#.#.#", "#...#", "#...#", "#...#", "#...#", ".....", "....."],
        "N": ["#...#", "#...#", "##..#", "#.#.#", "#..##", "#...#", "#...#", ".....", "....."],
        "O": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "P": ["####.", "#...#", "#...#", "####.", "#....", "#....", "#....", ".....", "....."],
        "Q": [".###.", "#...#", "#...#", "#...#", "#...#", "#..##", ".####", ".....", "....."],
        "R": ["####.", "#...#", "#...#", "####.", "#.#..", "#..#.", "#...#", ".....", "....."],
        "S": [".###.", "#...#", "#....", ".###.", "....#", "#...#", ".###.", ".....", "....."],
        "T": ["#####", "..#..", "..#..", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
        "U": ["#...#", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "V": ["#...#", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", "..#..", ".....", "....."],
        "W": ["#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", ".##.##.", ".......", "......."],
        "X": ["#...#", "#...#", ".#.#.", "..#..", ".#.#.", "#...#", "#...#", ".....", "....."],
        "Y": ["#...#", "#...#", ".#.#.", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
        "Z": ["#####", "....#", "...#.", "..#..", ".#...", "#....", "#####", ".....", "....."],
        "a": ["....", "....", ".##.", "...#", ".###", "#..#", ".###", "....", "...."],
        "b": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "###.", "....", "...."],
        "c": ["....", "....", ".###", "#...", "#...", "#...", ".###", "....", "...."],
        "d": ["...#", "...#", ".###", "#..#", "#..#", "#..#", ".###", "....", "...."],
        "e": ["....", "....", ".##.", "#..#", "####", "#...", ".###", "....", "...."],
        "f": ["..#", ".#.", "###", ".#.", ".#.", ".#.", ".#.", "...", "..."],
        "g": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "###."],
        "h": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
        "i": ["#", ".", "#", "#", "#", "#", "#", ".", "."],
        "j": [".#", "..", ".#", ".#", ".#", ".#", ".#", ".#", "#."],
        "k": ["#...", "#...", "#..#", "#.#.", "##..", "#.#.", "#..#", "....", "...."],
        "l": ["#.", "#.", "#.", "#.", "#.", "#.", ".#", "..", ".."],
        "m": [".....", ".....", "####.", "#.#.#", "#.#.#", "#.#.#", "#.#.#", ".....", "....."],
        "n": ["....", "....", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
        "o": ["....", "....", ".##.", "#..#", "#..#", "#..#", ".##.", "....", "...."],
        "p": ["....", "....", "###.", "#..#", "#..#", "#..#", "###.", "#...", "#..."],
        "q": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "...#"],
        "r": ["...", "...", "#.#", "##.", "#..", "#..", "#..", "...", "..."],
        "s": ["....", "....", ".###", "#...", ".##.", "...#", "###.", "....", "...."],
        "t": ["...", ".#.", "###", ".#.", ".#.", ".#.", ".##", "...", "..."],
        "u": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "....", "...."],
        "v": [".....", ".....", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", ".....", "....."],
        "w": [".....", ".....", "#...#", "#...#", "#.#.#", "#.#.#", ".#.#.", ".....", "....."],
        "x": ["....", "....", "#..#", "#..#", ".##.", "#..#", "#..#", "....", "...."],
        "y": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "...#", "###."],
        "z": ["....", "....", "####", "...#", ".##.", "#...", "####", "....", "...."],
        "0": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "1": [".#.", "##.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
        "2": [".###.", "#...#", "....#", "...#.", "..#..", ".#...", "#####", ".....", "....."],
        "3": [".###.", "#...#", "....#", "..##.", "....#", "#...#", ".###.", ".....", "....."],
        "4": ["...#.", "..##.", ".#.#.", "#..#.", "#####", "...#.", "...#.", ".....", "....."],
        "5": ["#####", "#....", "####.", "....#", "....#", "#...#", ".###.", ".....", "....."],
        "6": [".###.", "#...#", "#....", "####.", "#...#", "#...#", ".###.", ".....", "....."],
        "7": ["#####", "....#", "....#", "...#.", "...#.", "..#..", "..#..", ".....", "....."],
        "8": [".###.", "#...#", "#...#", ".###.", "#...#", "#...#", ".###.", ".....", "....."],
        "9": [".###.", "#...#", "#...#", ".####", "....#", "#...#", ".###.", ".....", "....."],
        ".": [".", ".", ".", ".", ".", ".", "#", ".", "."],
        "-": ["...", "...", "...", "###", "...", "...", "...", "...", "..."],
        "_": ["....", "....", "....", "....", "....", "....", "####", "....", "...."],
        "+": [".....", "..#..", "..#..", "#####", "..#..", "..#..", ".....", ".....", "....."],
        " ": ["...", "...", "...", "...", "...", "...", "...", "...", "..."]
    };

    /**
     * ウィンドウの下端の左にスクリプト名とバージョン（SCRIPT_NAME・SCRIPTMETA の Version）、右に VERSION_FOOTER_BRAND を小さく薄く入れる。部品をすべて足したあと、show() の直前に呼ぶ。
     * 直前の行（ボタン行など）の下余白を 0 にし、ウィンドウの下余白を VERSION_FOOTER_BOTTOM_MARGIN にする。
     * InDesign の ScriptUI は文字サイズの指定を無視するので、onDraw でドットを並べて描く（UITheme の isDarkUI() を使う）
     * @param {Window} targetWindow - ダイアログまたはパレット（縦並び）
     * @returns {Group|null} バージョンを描くグループ。すでに足してあれば null
     */
    function addVersionFooter(targetWindow) {
        if (targetWindow.versionFooter) return null;

        /* ボタン行の下余白（BUTTON_ROW_BOTTOM_MARGIN）は、この行が引き受ける / this row takes over the button row's bottom margin */
        var lastChild = targetWindow.children[targetWindow.children.length - 1];
        if (lastChild && lastChild.type === "group") {
            var rowMargins = lastChild.margins;
            lastChild.margins = [rowMargins[0], rowMargins[1], rowMargins[2], 0];
        }
        var windowMargins = targetWindow.margins;
        targetWindow.margins = [windowMargins[0], windowMargins[1], windowMargins[2], VERSION_FOOTER_BOTTOM_MARGIN];

        var footerGroup = targetWindow.add("group");
        footerGroup.orientation = "row";
        footerGroup.alignment = ["fill", "bottom"];
        footerGroup.alignChildren = ["left", "center"];
        footerGroup.margins = [0, VERSION_FOOTER_TOP_MARGIN, 0, 0];

        /* 左は「スクリプト名 バージョン」。字形に無い文字（日本語など）を含む名前や、幅が足りないときはバージョンだけ。
           左の欄は余りの幅まで伸ばし、ダイアログの幅はバージョンの分しか求めない（名前でダイアログを広げない）
           / Left: "script name version", or just the version when the name has characters without glyphs (e.g. Japanese)
             or does not fit. The left area takes the extra width and only asks for the version's width, so the name never widens the dialog */
        var versionText = getVersionInfoValue("version");
        var leftCandidates = [versionText];
        if (typeof SCRIPT_NAME === "string" && hasVersionFooterGlyphs(SCRIPT_NAME)) leftCandidates.unshift(SCRIPT_NAME + " " + versionText);
        var versionCanvas = addVersionFooterCanvas(footerGroup, leftCandidates);
        versionCanvas.alignment = ["fill", "center"];
        bindVersionInfoClick(versionCanvas); /* クリックで「このスクリプトについて」を開く / click opens About This Script */
        var brandCanvas = addVersionFooterCanvas(footerGroup, [VERSION_FOOTER_BRAND]);
        brandCanvas.alignment = ["right", "center"];
        brandCanvas.maximumSize = brandCanvas.preferredSize;
        targetWindow.versionFooter = footerGroup;
        return versionCanvas;
    }

    /**
     * 文字列のすべての文字に字形があるかを返す
     * @param {string} dotText - 文字列
     * @returns {boolean} すべてあれば true
     */
    function hasVersionFooterGlyphs(dotText) {
        for (var i = 0; i < dotText.length; i++) {
            if (!VERSION_FOOTER_GLYPHS.hasOwnProperty(dotText.charAt(i))) return false;
        }
        return dotText.length > 0;
    }

    /**
     * ドットで文字列を描くグループを足す。候補のうち、欄の幅に収まる最初のものを描く
     * @param {Group} parentGroup - 足す先
     * @param {string[]} dotTexts - 描く文字列の候補（長い順。いちばん短いものの幅を欄の幅として求める）
     * @returns {Group} 描くグループ
     */
    function addVersionFooterCanvas(parentGroup, dotTexts) {
        var dotCanvas = parentGroup.add("group");
        var shortestText = dotTexts[dotTexts.length - 1];
        dotCanvas.preferredSize = [measureVersionFooterDots(shortestText) * VERSION_FOOTER_PIXEL_SIZE, VERSION_FOOTER_GLYPH_HEIGHT * VERSION_FOOTER_PIXEL_SIZE];
        dotCanvas.helpTip = dotTexts[0];
        dotCanvas.onDraw = function () {
            var availableWidth = this.size ? this.size.width : 0;
            var drawText = shortestText;
            for (var i = 0; i < dotTexts.length; i++) {
                if (measureVersionFooterDots(dotTexts[i]) * VERSION_FOOTER_PIXEL_SIZE <= availableWidth) {
                    drawText = dotTexts[i];
                    break;
                }
            }
            drawVersionFooterDots(this.graphics, drawText);
        };
        return dotCanvas;
    }

    /**
     * 文字列の幅をドットの数で返す（字形に無い文字は数えない）
     * @param {string} versionText - 文字列
     * @returns {number} 幅（ドット）
     */
    function measureVersionFooterDots(versionText) {
        var dotCount = 0;
        for (var i = 0; i < versionText.length; i++) {
            var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
            if (!glyphRows) continue;
            if (dotCount > 0) dotCount += VERSION_FOOTER_LETTER_GAP;
            dotCount += glyphRows[0].length;
        }
        return dotCount;
    }

    /**
     * 文字列をドットで描く。横に続くドットは1つの四角にまとめて塗る
     * @param {ScriptUIGraphics} canvasGraphics - 描く先の graphics
     * @param {string} versionText - 文字列
     * @returns {void}
     */
    function drawVersionFooterDots(canvasGraphics, versionText) {
        var footerBrush = canvasGraphics.newBrush(canvasGraphics.BrushType.SOLID_COLOR, VERSION_FOOTER_COLOR);
        var dotSize = VERSION_FOOTER_PIXEL_SIZE;
        var cursorDots = 0;
        canvasGraphics.newPath();
        for (var i = 0; i < versionText.length; i++) {
            var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
            if (!glyphRows) continue;
            if (cursorDots > 0) cursorDots += VERSION_FOOTER_LETTER_GAP;
            for (var row = 0; row < glyphRows.length; row++) {
                var rowText = glyphRows[row];
                var runStart = -1;
                for (var col = 0; col <= rowText.length; col++) {
                    var isDot = col < rowText.length && rowText.charAt(col) === "#";
                    if (isDot && runStart < 0) runStart = col;
                    if (!isDot && runStart >= 0) {
                        canvasGraphics.rectPath((cursorDots + runStart) * dotSize, row * dotSize, (col - runStart) * dotSize, dotSize);
                        runStart = -1;
                    }
                }
            }
            cursorDots += glyphRows[0].length;
        }
        canvasGraphics.fillPath(footerBrush);
    }

    // -----------------------------------------
    // このスクリプトについて / About this script
    // -----------------------------------------
    /* 左下のスクリプト名・バージョンをクリックすると開く。SCRIPTMETA の項目を getScriptMetaValue() で直接読み（スクリプト情報の部品）、
       この部品だけで動くよう、文言（VERSION_INFO_LABELS）とレイアウトの値も部品の中に持つ
       Opened by clicking the name and version at the bottom left. Reads the SCRIPTMETA fields directly with getScriptMetaValue(),
       and keeps its own labels and layout values so the part works on its own */

    var VERSION_INFO_LANG = (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en"; /* 文言の言語 / label language */
    var VERSION_INFO_IS_INDESIGN = (String(app.name).indexOf("InDesign") >= 0);         /* InDesign で動いているか / running in InDesign */
    /* 実行中のスクリプトのファイル。ボタンを押すころには $.fileName が別のものを指すことがあるので、読み込んだ時点で控える
       The running script file, captured at load time; $.fileName may point elsewhere by the time the button is clicked */
    var VERSION_INFO_SCRIPT_FILE = new File($.fileName);
    var VERSION_INFO_REVEAL_APP_PATH = "/Applications/RevealInFinder.app";   /* Finder で選択表示する補助アプリ / helper app that reveals a file in the Finder */
    var VERSION_INFO_REVEAL_PATH_FILE = "/tmp/illustrator_reveal_path.txt";  /* 補助アプリが読むパス（アプリの旧名のまま。変えると読めない）/ path the helper reads (keeps the app's old name) */
    var VERSION_INFO_FILE_VIEWER_APP_PATH = "/Applications/OpenInFileViewer.app";   /* Path Finder が起動中なら Path Finder、そうでなければ Finder で選択表示する補助アプリ / helper app that reveals a file in Path Finder when it is running, otherwise in the Finder */
    var VERSION_INFO_FILE_VIEWER_PATH_FILE = "/tmp/open_in_file_viewer_path.txt"; /* OpenInFileViewer.app が読むパス / path the helper reads */

    var VERSION_INFO_WINDOW_MARGINS = 16;               /* ウィンドウの余白 / window margin */
    var VERSION_INFO_WINDOW_SPACING = 12;               /* ウィンドウ内の間隔 / window spacing */
    var VERSION_INFO_PANEL_MARGINS = [16, 20, 16, 12];  /* パネルの余白 [左,上,右,下] / panel margins */
    var VERSION_INFO_ROW_SPACING = 6;                   /* 項目名と値の間隔 / gap between a label and its value */
    var VERSION_INFO_BUTTON_SPACING = 10;               /* ボタンどうしの間隔 / gap between buttons */
    var VERSION_INFO_BUTTON_ROW_MARGINS = [0, 5, 0, 14]; /* ボタン行の余白（Illustrator 標準のダイアログに合わせる）/ button row margins */
    var VERSION_INFO_COPY_BUTTON_SIZE = 22;             /* コピーボタンの一辺 / copy button size */
    /* コピーボタンの配色。∧∨（ステップボタン）と同じく、背景に重ねる半透明 / Copy button colors, translucent overlays like the stepper */
    var VERSION_INFO_COPY_FILL_COLOR    = isDarkUI() ? [0, 0, 0, 0.10] : [1, 1, 1, 0.50]; /* 地 / background */
    var VERSION_INFO_COPY_FRAME_COLOR   = isDarkUI() ? [1, 1, 1, 0.12] : [0, 0, 0, 0.15]; /* 枠線 / frame */
    var VERSION_INFO_COPY_PRESSED_COLOR = isDarkUI() ? [1, 1, 1, 0.12] : [0, 0, 0, 0.13]; /* 押下中 / pressed */
    var VERSION_INFO_COPY_ICON_COLOR    = isDarkUI() ? [1, 1, 1, 0.85] : [0, 0, 0, 0.70]; /* アイコン / icon */

    var VERSION_INFO_LABELS = {
        dialog: {
            title: { ja: "このスクリプトについて", en: "About This Script" }
        },
        panel: {
            generalInfo: { ja: "基本情報", en: "General" },
            links: { ja: "リンク", en: "Links" }
        },
        fieldLabel: {
            scriptName: { ja: "スクリプト名", en: "Script Name" },
            scriptFeature: { ja: "機能", en: "Feature" },
            scriptVersion: { ja: "バージョン", en: "Version" }
        },
        button: {
            openReadme: { ja: "README", en: "README" },
            openSourceCode: { ja: "ソースコード", en: "Source Code" },
            openArticle: { ja: "記事（note）", en: "Article (note)" },
            revealInFinder: { ja: "Finder に表示", en: "Show in Finder" },
            revealInExplorer: { ja: "エクスプローラーに表示", en: "Show in Explorer" },
            close: { ja: "閉じる", en: "Close" }
        },
        tooltip: {
            versionInfo: { ja: "クリックで、このスクリプトについての情報を表示", en: "Click for info about this script" },
            copyScriptName: { ja: "スクリプト名をコピー", en: "Copy Script Name" }
        },
        alert: {
            copyFailed: { ja: "スクリプト名をコピーできませんでした。", en: "Could not copy the script name." }
        }
    };

    /**
     * VERSION_INFO_LABELS の文言を今の言語で返す
     * @param {Object} labelEntry - { ja, en }
     * @returns {string} 文言
     */
    function getVersionInfoLabel(labelEntry) {
        return labelEntry[VERSION_INFO_LANG] || labelEntry.en;
    }

    /**
     * 「このスクリプトについて」とフッターに出す値を返す。SCRIPTMETA の項目を getScriptMetaValue() で直接読み、名前は SCRIPT_NAME。
     * getScriptMetaValue() が無いスクリプトは従来の SCRIPT_* を typeof で読む（未定義でも例外にならない）
     * @param {string} variableKey - "name" / "version" / "feature" / "readme" / "source" / "article"
     * @returns {string} 値（バージョンは「v1.2.0」の形）。無ければ空文字
     */
    function getVersionInfoValue(variableKey) {
        var infoValue = "";
        if (variableKey === "name") {
            infoValue = (typeof SCRIPT_NAME !== "undefined") ? SCRIPT_NAME : "";
        } else if (typeof getScriptMetaValue === "function") {
            var isJapanese = (VERSION_INFO_LANG === "ja");
            if (variableKey === "version") infoValue = getScriptMetaValue("Version") ? "v" + getScriptMetaValue("Version") : "";
            else if (variableKey === "feature") infoValue = getScriptMetaValue("Name");
            else if (variableKey === "readme") infoValue = getScriptMetaValue(isJapanese ? "X-com.swwwitch-README-JA" : "X-com.swwwitch-README-EN");
            else if (variableKey === "source") infoValue = getScriptMetaValue("X-com.swwwitch-Source-URL");
            else if (variableKey === "article") infoValue = getScriptMetaValue("X-com.swwwitch-Article-URL");
        } else {
            if (variableKey === "version") infoValue = (typeof SCRIPT_VERSION !== "undefined") ? SCRIPT_VERSION : "";
            else if (variableKey === "feature") infoValue = (typeof SCRIPT_DESCRIPTION_JA !== "undefined") ? SCRIPT_DESCRIPTION_JA : "";
            else if (variableKey === "readme") {
                if (VERSION_INFO_LANG === "ja") infoValue = (typeof SCRIPT_README_JA !== "undefined") ? SCRIPT_README_JA : "";
                else infoValue = (typeof SCRIPT_README_EN !== "undefined") ? SCRIPT_README_EN : "";
            }
            else if (variableKey === "source") infoValue = (typeof SCRIPT_SOURCE_URL !== "undefined") ? SCRIPT_SOURCE_URL : "";
            else if (variableKey === "article") infoValue = (typeof SCRIPT_ARTICLE_URL !== "undefined") ? SCRIPT_ARTICLE_URL : "";
        }
        return (infoValue === undefined || infoValue === null) ? "" : String(infoValue);
    }

    /**
     * 左下のスクリプト名・バージョンをクリックしたら、「このスクリプトについて」のダイアログを開くようにする
     * @param {Group} versionCanvas - addVersionFooterCanvas() で作った左の欄
     * @returns {void}
     */
    function bindVersionInfoClick(versionCanvas) {
        versionCanvas.helpTip = getVersionInfoLabel(VERSION_INFO_LABELS.tooltip.versionInfo);
        /* mousedown で開くと、離したときのイベントが閉じたあとの欄に届くので mouseup で開く
           Open on mouseup so the release does not land on the footer after the dialog closes */
        versionCanvas.addEventListener("mouseup", function () {
            showVersionInfoDialog(versionCanvas.window.type === "palette");
        });
    }

    /**
     * 「このスクリプトについて」のダイアログを開く。スクリプト名（コピーボタン付き）・機能・バージョンの［基本情報］パネル、
     * README・ソースコード・記事を開く［リンク］パネル、［Finder に表示］［閉じる］を並べる
     * @param {boolean} isFromPalette - 常駐パレットから開いたなら true（Illustrator のコピーを BridgeTalk 経由にする）
     * @returns {void}
     */
    function showVersionInfoDialog(isFromPalette) {
        var infoDialog = new Window("dialog", getVersionInfoLabel(VERSION_INFO_LABELS.dialog.title));
        infoDialog.orientation = "column";
        infoDialog.alignChildren = "fill";
        infoDialog.margins = VERSION_INFO_WINDOW_MARGINS;
        infoDialog.spacing = VERSION_INFO_WINDOW_SPACING;

        var infoPanel = addVersionInfoPanel(infoDialog, VERSION_INFO_LABELS.panel.generalInfo);
        /* コピーボタンはスクリプト名だけ / Only the script name gets a copy button */
        addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptName, getVersionInfoValue("name"), isFromPalette);
        /* 機能は README の見出しで日本語だけなので、日本語の画面だけに出す / The feature is the Japanese README title, shown only in the Japanese UI */
        if (VERSION_INFO_LANG === "ja" && getVersionInfoValue("feature") !== "") {
            addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptFeature, getVersionInfoValue("feature"), null);
        }
        addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptVersion, getVersionInfoValue("version"), null);
        alignVersionInfoLabels(infoPanel);

        /* URL が無いボタンはディム表示にして、並びはスクリプトごとに変えない / Dim buttons without a URL; keep the same buttons everywhere */
        var linksPanel = addVersionInfoPanel(infoDialog, VERSION_INFO_LABELS.panel.links);
        var linkButtonRow = linksPanel.add("group");
        linkButtonRow.orientation = "row";
        linkButtonRow.alignment = ["left", "center"];
        linkButtonRow.alignChildren = ["left", "center"];
        linkButtonRow.spacing = VERSION_INFO_BUTTON_SPACING;
        addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openReadme, getVersionInfoValue("readme"));
        addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openSourceCode, getVersionInfoValue("source"));
        addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openArticle, getVersionInfoValue("article"));

        /* ボタン行：左に［Finder に表示］、右に［閉じる］ / Button row: Show in Finder at the left, Close at the right */
        var buttonRow = infoDialog.add("group");
        buttonRow.orientation = "row";
        buttonRow.alignment = ["fill", "bottom"];
        buttonRow.alignChildren = ["left", "center"];
        buttonRow.margins = VERSION_INFO_BUTTON_ROW_MARGINS;
        buttonRow.spacing = VERSION_INFO_BUTTON_SPACING;
        var revealLabel = (File.fs === "Windows") ? VERSION_INFO_LABELS.button.revealInExplorer : VERSION_INFO_LABELS.button.revealInFinder;
        var revealButton = buttonRow.add("button", undefined, getVersionInfoLabel(revealLabel));
        if (VERSION_INFO_SCRIPT_FILE.exists) {
            revealButton.helpTip = VERSION_INFO_SCRIPT_FILE.fsName;
            revealButton.onClick = function () {
                revealVersionInfoScriptFile(VERSION_INFO_SCRIPT_FILE);
                infoDialog.close();
            };
        } else {
            revealButton.enabled = false;
        }
        /* 余りの幅を吸って、［閉じる］を右端に寄せる / Absorbs the extra width so Close sits at the right edge */
        var buttonSpacer = buttonRow.add("group");
        buttonSpacer.alignment = ["fill", "fill"];
        buttonSpacer.minimumSize.width = 0;
        var closeButton = buttonRow.add("button", undefined, getVersionInfoLabel(VERSION_INFO_LABELS.button.close), { name: "ok" });
        closeButton.alignment = ["right", "center"];
        /* return でも esc でも閉じる / Close with both Return and Esc */
        infoDialog.defaultElement = closeButton;
        infoDialog.cancelElement = closeButton;
        infoDialog.onShow = function () {
            pinVersionInfoCopyButton(infoPanel);
        };
        infoDialog.show();
    }

    /**
     * 「このスクリプトについて」のダイアログにパネルを足す
     * @param {Window} infoDialog - 足す先のダイアログ
     * @param {Object} titleEntry - パネル名の文言 { ja, en }
     * @returns {Panel} パネル
     */
    function addVersionInfoPanel(infoDialog, titleEntry) {
        var infoPanel = infoDialog.add("panel", undefined, getVersionInfoLabel(titleEntry));
        infoPanel.orientation = "column";
        infoPanel.alignChildren = ["fill", "top"];
        infoPanel.alignment = "fill";
        infoPanel.margins = VERSION_INFO_PANEL_MARGINS;
        infoPanel.spacing = VERSION_INFO_WINDOW_SPACING;
        return infoPanel;
    }

    /**
     * ［基本情報］パネルに「項目名 : 値」の1行を足す（値は見せるだけなので statictext）
     * @param {Panel} infoPanel - 足す先のパネル
     * @param {Object} labelEntry - 項目名の文言 { ja, en }
     * @param {string} valueText - 表示する値
     * @param {boolean|null} copyFromPalette - 値の右にコピーボタンを置くなら true / false（常駐パレットから開いたか）、置かないなら null
     * @returns {StaticText} 値の statictext
     */
    function addVersionInfoRow(infoPanel, labelEntry, valueText, copyFromPalette) {
        var infoRow = infoPanel.add("group");
        infoRow.orientation = "row";
        infoRow.alignment = ["fill", "center"];
        infoRow.alignChildren = ["left", "center"];
        infoRow.spacing = VERSION_INFO_ROW_SPACING;
        /* 項目名のコロンは日本語が「 :」、英語が「:」 / Colon: " :" in Japanese, ":" in English */
        var rowLabel = infoRow.add("statictext", undefined, getVersionInfoLabel(labelEntry) + (VERSION_INFO_LANG === "ja" ? " :" : ":"));
        rowLabel.justify = "right";
        var valueLabel = infoRow.add("statictext", undefined, valueText);
        if (copyFromPalette === null) return valueLabel;
        /* statictext の fill では伸びないことがあるので、伸びるスペーサーでコピーボタンをパネルの右端に寄せる
           A stretching spacer (statictext fill is unreliable) pushes the copy button to the panel's right edge */
        var copySpacer = infoRow.add("group");
        copySpacer.alignment = ["fill", "fill"];
        copySpacer.minimumSize.width = 0;
        infoRow.copyButton = addVersionInfoCopyButton(infoRow, valueLabel, copyFromPalette);
        return valueLabel;
    }

    /**
     * コピーボタンを［基本情報］パネルの右端に置き直す（表示時、レイアウトが決まったあとに呼ぶ）。
     * 行の fill・スペーサーだけでは、項目の値が短い行がパネルの幅まで伸びないことがあるので、行の幅とボタンの位置を直接入れる
     * @param {Panel} infoPanel - ［基本情報］パネル
     * @returns {void}
     */
    function pinVersionInfoCopyButton(infoPanel) {
        var innerWidth = infoPanel.size.width - VERSION_INFO_PANEL_MARGINS[0] - VERSION_INFO_PANEL_MARGINS[2];
        for (var i = 0; i < infoPanel.children.length; i++) {
            var infoRow = infoPanel.children[i];
            if (!infoRow.copyButton || !infoRow.size || infoRow.size.width >= innerWidth) continue;
            infoRow.size = [innerWidth, infoRow.size.height];
            infoRow.copyButton.location = [innerWidth - infoRow.copyButton.size.width, infoRow.copyButton.location[1]];
        }
    }

    /**
     * ［基本情報］パネルの項目名を、いちばん長いものの幅にそろえる
     * @param {Panel} infoPanel - 対象のパネル
     * @returns {void}
     */
    function alignVersionInfoLabels(infoPanel) {
        var rowLabels = [];
        var labelWidth = 0;
        for (var i = 0; i < infoPanel.children.length; i++) {
            var rowLabel = infoPanel.children[i].children[0]; /* 行の先頭は項目名 / each row starts with its label */
            rowLabels.push(rowLabel);
            labelWidth = Math.max(labelWidth, rowLabel.graphics.measureString(rowLabel.text)[0]);
        }
        for (var j = 0; j < rowLabels.length; j++) {
            rowLabels[j].preferredSize.width = Math.ceil(labelWidth) + 2;
        }
    }

    /**
     * URL を開くボタンを足す。URL が無ければディム表示にする。開いたらダイアログは閉じる
     * @param {Window} infoDialog - ボタンを置くダイアログ
     * @param {Group} buttonGroup - 足す先のグループ
     * @param {Object} labelEntry - ボタンの文言 { ja, en }
     * @param {string} targetUrl - 開く URL（空ならディム表示）
     * @returns {void}
     */
    function addVersionInfoLinkButton(infoDialog, buttonGroup, labelEntry, targetUrl) {
        var linkButton = buttonGroup.add("button", undefined, getVersionInfoLabel(labelEntry));
        if (!targetUrl) {
            linkButton.enabled = false;
            return;
        }
        linkButton.helpTip = targetUrl;
        linkButton.onClick = function () {
            openVersionInfoUrl(targetUrl);
            infoDialog.close();
        };
    }

    /**
     * 値をクリップボードへコピーするアイコンのボタンを足す。コピーしたらアイコンをチェックマークに変え、ポインタが外れたら戻す
     * @param {Group} infoRow - 足す先の行
     * @param {StaticText} valueLabel - コピーする値の statictext
     * @param {boolean} isFromPalette - 常駐パレットから開いたなら true
     * @returns {Group} ボタンとして使う group
     */
    function addVersionInfoCopyButton(infoRow, valueLabel, isFromPalette) {
        var buttonSize = VERSION_INFO_COPY_BUTTON_SIZE;
        var copyButton = infoRow.add("group");
        copyButton.preferredSize = [buttonSize, buttonSize];
        copyButton.minimumSize = [buttonSize, buttonSize];
        copyButton.maximumSize = [buttonSize, buttonSize];
        copyButton.helpTip = getVersionInfoLabel(VERSION_INFO_LABELS.tooltip.copyScriptName);
        copyButton.isPressed = false;
        copyButton.isCopied = false;

        copyButton.onDraw = function () {
            var buttonGraphics = copyButton.graphics;
            var fillColor = copyButton.isPressed ? VERSION_INFO_COPY_PRESSED_COLOR : VERSION_INFO_COPY_FILL_COLOR;
            buttonGraphics.newPath();
            buttonGraphics.rectPath(0, 0, buttonSize, buttonSize);
            buttonGraphics.fillPath(buttonGraphics.newBrush(buttonGraphics.BrushType.SOLID_COLOR, fillColor));
            buttonGraphics.newPath();
            buttonGraphics.rectPath(0.5, 0.5, buttonSize - 1, buttonSize - 1);
            buttonGraphics.strokePath(buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_FRAME_COLOR, 1));
            if (copyButton.isCopied) drawVersionInfoCheckIcon(buttonGraphics, buttonSize);
            else drawVersionInfoCopyIcon(buttonGraphics, buttonSize);
        };

        /**
         * 状態を変えて描き直す（group には notify() が無いので隠して再表示する）
         * @param {boolean} isPressed - 押下中なら true
         * @param {boolean} isCopied - コピー済みなら true
         * @returns {void}
         */
        function repaint(isPressed, isCopied) {
            if (copyButton.isPressed === isPressed && copyButton.isCopied === isCopied) return;
            copyButton.isPressed = isPressed;
            copyButton.isCopied = isCopied;
            copyButton.hide();
            copyButton.show();
        }
        copyButton.addEventListener("mousedown", function () { repaint(true, false); });
        copyButton.addEventListener("mouseup", function () {
            if (!copyButton.isPressed) return;
            var hasCopied = copyVersionInfoText(valueLabel.text, isFromPalette);
            repaint(false, hasCopied);
            if (!hasCopied) alert(getVersionInfoLabel(VERSION_INFO_LABELS.alert.copyFailed));
        });
        /* 押したまま外へ出たとき・コピーのあと離れたときは元の見た目に戻す / Reset when the pointer leaves */
        copyButton.addEventListener("mouseout", function () { repaint(false, false); });
        return copyButton;
    }

    /**
     * コピーのアイコン（後ろの四角に前の四角を重ねた形）を描く
     * @param {ScriptUIGraphics} buttonGraphics - 描画先
     * @param {number} buttonSize - ボタンの一辺
     * @returns {void}
     */
    function drawVersionInfoCopyIcon(buttonGraphics, buttonSize) {
        var iconPen = buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_ICON_COLOR, 1);
        var sheetSize = 9;
        var offset = 3;
        var frontLeft = Math.round((buttonSize - sheetSize - offset) / 2) + offset + 0.5;
        var frontTop = Math.round((buttonSize - sheetSize - offset) / 2) + offset + 0.5;
        var backLeft = frontLeft - offset;
        var backTop = frontTop - offset;
        /* 後ろの四角は、前の四角に隠れる部分を描かない（上と左の辺だけ）/ Draw only the visible top and left of the back sheet */
        buttonGraphics.newPath();
        buttonGraphics.moveTo(frontLeft, backTop + sheetSize);
        buttonGraphics.lineTo(backLeft, backTop + sheetSize);
        buttonGraphics.lineTo(backLeft, backTop);
        buttonGraphics.lineTo(backLeft + sheetSize, backTop);
        buttonGraphics.lineTo(backLeft + sheetSize, frontTop);
        buttonGraphics.strokePath(iconPen);
        buttonGraphics.newPath();
        buttonGraphics.rectPath(frontLeft, frontTop, sheetSize, sheetSize);
        buttonGraphics.strokePath(iconPen);
    }

    /**
     * コピーしたあとのチェックマークを描く
     * @param {ScriptUIGraphics} buttonGraphics - 描画先
     * @param {number} buttonSize - ボタンの一辺
     * @returns {void}
     */
    function drawVersionInfoCheckIcon(buttonGraphics, buttonSize) {
        var center = buttonSize / 2;
        buttonGraphics.newPath();
        buttonGraphics.moveTo(center - 4.5, center);
        buttonGraphics.lineTo(center - 1.5, center + 3);
        buttonGraphics.lineTo(center + 4.5, center - 3.5);
        buttonGraphics.strokePath(buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_ICON_COLOR, 1.5));
    }

    /**
     * 文字列をクリップボードにコピーする。
     * InDesign は AppleScript（Windows は VBScript）の doScript で書き込む。
     * Illustrator は文字列を直接送る API が無いので、一時レイヤーのテキストフレームを選んでコピーする（プレーンテキストも入る）。
     * Illustrator の常駐パレットからは DOM に触れないので、同じ処理を BridgeTalk でメインエンジンに渡す
     * @param {string} copyText - コピーする文字列
     * @param {boolean} isFromPalette - 常駐パレットから開いたなら true
     * @returns {boolean} コピーの操作まで進めたら true
     */
    function copyVersionInfoText(copyText, isFromPalette) {
        try {
            if (VERSION_INFO_IS_INDESIGN) {
                if (File.fs === "Windows") {
                    app.doScript('CreateObject("htmlfile").parentWindow.clipboardData.setData "text", "' + copyText.replace(/"/g, '""') + '"', ScriptLanguage.VISUAL_BASIC);
                } else {
                    app.doScript('set the clipboard to "' + copyText.replace(/\\/g, "\\\\").replace(/"/g, '\\"') + '"', ScriptLanguage.APPLESCRIPT_LANGUAGE);
                }
                return true;
            }
            var copyScript = buildVersionInfoCopyScript(copyText);
            if (!isFromPalette) return eval(copyScript) === true;
            var copyMessage = new BridgeTalk();
            copyMessage.target = BridgeTalk.appSpecifier; /* 複数バージョンを起動していても、今のバージョンに送る / the running version, even with several installed */
            copyMessage.body = copyScript;
            copyMessage.send();
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * Illustrator でコピーする処理を文字列で組み立てる（BridgeTalk でも送れるよう、関数の toString() は使わない）。
     * 一時レイヤーにテキストフレームを作り、フレームを選んでメニューコマンドでコピーして、レイヤーごと消す。
     * 文字として選択（textRange.select()）してコピーすると「オブジェクトをコピーできません」の警告が出るので使わない。
     * 選択と作業レイヤーは元に戻す。ドキュメントが無ければ一時ドキュメントを作って閉じる
     * @param {string} copyText - コピーする文字列
     * @returns {string} 実行すると true / false を返すスクリプト
     */
    function buildVersionInfoCopyScript(copyText) {
        return [
            "(function (copyText) {",
            "    var usingTempDoc = (app.documents.length === 0);",
            "    var targetDoc = usingTempDoc ? app.documents.add() : app.activeDocument;",
            "    var savedSelection = (!usingTempDoc && targetDoc.selection instanceof Array) ? targetDoc.selection : [];",
            "    var savedActiveLayer = usingTempDoc ? null : targetDoc.activeLayer;",
            "    var tempLayer = null;",
            "    var hasCopied = false;",
            "    try {",
            "        tempLayer = targetDoc.layers.add();",
            "        var tempFrame = tempLayer.textFrames.add();",
            "        tempFrame.contents = copyText;",
            "        app.redraw();",
            "        app.executeMenuCommand('deselectall');",
            "        tempFrame.selected = true;",
            "        app.redraw();",
            "        app.executeMenuCommand('copy');",
            "        app.redraw();",
            "        hasCopied = true;",
            "    } catch (e) { }",
            "    if (tempLayer) { try { tempLayer.remove(); } catch (e) { } }",
            "    if (usingTempDoc) { targetDoc.close(SaveOptions.DONOTSAVECHANGES); return hasCopied; }",
            "    try { targetDoc.activeLayer = savedActiveLayer; } catch (e) { }",
            "    targetDoc.selection = null;",
            "    for (var i = 0; i < savedSelection.length; i++) { try { savedSelection[i].selected = true; } catch (e) { } }",
            "    app.redraw();",
            "    return hasCopied;",
            "})(" + quoteVersionInfoString(copyText) + ");"
        ].join(" "); /* 行コメントを含まないので1行につなぐ（BridgeTalk の転送で改行が化けない）/ joined on one line; there are no line comments */
    }

    /**
     * 文字列を、スクリプトに埋め込める二重引用符のリテラルにする
     * @param {string} sourceText - 文字列
     * @returns {string} "…" の形のリテラル
     */
    function quoteVersionInfoString(sourceText) {
        return '"' + String(sourceText).replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/\r/g, "\\r").replace(/\n/g, "\\n") + '"';
    }

    /**
     * ファイルを選択した状態で表示する。OpenInFileViewer.app があれば Path Finder が起動中なら Path Finder、そうでなければ Finder で表示し、
     * 無ければ RevealInFinder.app で Finder に表示する。どちらも無いときは入っているフォルダーを開く
     * @param {File} scriptFile - 表示するファイル
     * @returns {void}
     */
    function revealVersionInfoScriptFile(scriptFile) {
        if (launchVersionInfoHelperApp(VERSION_INFO_FILE_VIEWER_APP_PATH, VERSION_INFO_FILE_VIEWER_PATH_FILE, scriptFile.fsName)) return;
        if (launchVersionInfoHelperApp(VERSION_INFO_REVEAL_APP_PATH, VERSION_INFO_REVEAL_PATH_FILE, scriptFile.fsName)) return;
        /* アプリが無い・起動できないときは、何も起きないままにせずフォルダーを開く / Otherwise open the enclosing folder */
        scriptFile.parent.execute();
    }

    /**
     * 補助アプリが読む一時ファイルにパスを書き出してから、アプリを起動する
     * @param {string} appPath - 補助アプリ（.app）のパス
     * @param {string} pathFilePath - アプリが読む一時ファイルのパス
     * @param {string} targetPath - 表示するファイルの絶対パス
     * @returns {boolean} 起動できたら true。Mac 以外・アプリが無い・起動できないときは false
     */
    function launchVersionInfoHelperApp(appPath, pathFilePath, targetPath) {
        /* .app は実体がフォルダーなので File.exists が false になる環境がある。Folder でも確かめる / An .app is a folder, so check with Folder too */
        if (File.fs !== "Macintosh" || !(new Folder(appPath).exists || new File(appPath).exists)) return false;
        return writeVersionInfoRevealPath(pathFilePath, targetPath) && new File(appPath).execute();
    }

    /**
     * 補助アプリに渡すパスを一時ファイルへ書き出す（UTF-8、改行なし）
     * @param {string} pathFilePath - 一時ファイルのパス
     * @param {string} targetPath - 絶対パス
     * @returns {boolean} 書き出せたら true
     */
    function writeVersionInfoRevealPath(pathFilePath, targetPath) {
        var pathFile = new File(pathFilePath);
        pathFile.encoding = "UTF-8";
        pathFile.lineFeed = "Unix";
        if (!pathFile.open("w")) return false;
        var hasWritten = pathFile.write(targetPath);
        pathFile.close();
        return hasWritten;
    }

    /**
     * URL を既定のブラウザーで開く（一時的なショートカットファイルを作って開く）
     * @param {string} targetUrl - 開く URL
     * @returns {void}
     */
    function openVersionInfoUrl(targetUrl) {
        var isMac = (File.fs === "Macintosh");
        var shortcutFile = new File(Folder.temp.fsName + "/VersionInfoLink" + (isMac ? ".webloc" : ".url"));
        shortcutFile.encoding = "UTF-8";
        shortcutFile.open("w");
        if (isMac) {
            var escapedUrl = targetUrl.replace(/&/g, "&amp;").replace(/</g, "&lt;");
            shortcutFile.write('<?xml version="1.0" encoding="UTF-8"?>\n'
                + '<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">\n'
                + '<plist version="1.0"><dict><key>URL</key><string>' + escapedUrl + '</string></dict></plist>\n');
        } else {
            shortcutFile.write("[InternetShortcut]\r\nURL=" + targetUrl + "\r\n");
        }
        shortcutFile.close();
        shortcutFile.execute();
    }

    // バージョン表示（再利用パーツ）ここまで / End of the reusable version footer

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ダイアログを構築・表示し、選択結果の処理設定を返す
     * @returns {Object} 処理設定。キャンセル時は null
     */
    function showOptionsDialog() {
        var optionsDialog = new Window("dialog", getLabel("dialog.title"));
        setupWindow(optionsDialog);

        var targetControls = buildTargetArtboardPanel(optionsDialog);
        var layerNameControls = buildLayerNamePanel(optionsDialog);
        var guideDestinationControls = buildGuideDestinationPanel(optionsDialog);
        var exclusionControls = buildExclusionPanel(optionsDialog);
        var postProcessControls = buildPostProcessPanel(optionsDialog);

        var buttonRow = addButtonRow(optionsDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);
        optionsDialog.defaultElement = btnOK;
        optionsDialog.cancelElement = btnCancel;

        addVersionFooter(optionsDialog);
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
            collectGuidesToGuideLayer: guideDestinationControls.guideLayerRadio.value,
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
        setupPanel(targetArtboardPanel, 6);

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
     * ガイドの移動先パネル（_guide レイヤー／各レイヤー）を構築する
     * @param {Window} parentDialog - 追加先のダイアログ
     * @returns {Object} ラジオボタン参照をまとめたオブジェクト
     */
    function buildGuideDestinationPanel(parentDialog) {
        var guideDestinationPanel = parentDialog.add("panel", undefined, getLabel("panel.guideDestination"));
        setupPanel(guideDestinationPanel, 6);

        var guideDestinationRow = addRowGroup(guideDestinationPanel, "center");
        var guideLayerRadio = guideDestinationRow.add("radiobutton", undefined, getLabel("radio.guideToGuideLayer"));
        var artboardLayersRadio = guideDestinationRow.add("radiobutton", undefined, getLabel("radio.guideToArtboardLayers"));
        setSharedHelpTip(getLabel("tooltip.guideDestination"), [guideDestinationPanel, guideLayerRadio, artboardLayersRadio]);

        guideLayerRadio.value = DEFAULT_OPTIONS.collectGuidesToGuideLayer;
        artboardLayersRadio.value = !DEFAULT_OPTIONS.collectGuidesToGuideLayer;

        return { guideLayerRadio: guideLayerRadio };
    }

    /**
     * レイヤー名パネル（番号・名前・区切り文字）を構築する
     * @param {Window} parentDialog - 追加先のダイアログ
     * @returns {Object} チェックボックス参照と区切り文字インデックス取得関数
     */
    function buildLayerNamePanel(parentDialog) {
        var layerNamePanel = parentDialog.add("panel", undefined, getLabel("panel.layerName"));
        setupPanel(layerNamePanel, 6);

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
                sampleLayerName = sampleLayerName.substring(0, LAYER_NAME_PREVIEW_MAX_LENGTH) + "...";
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
        setupPanel(exclusionPanel, 6);
        exclusionPanel.helpTip = getLabel("tooltip.exclusionPanel");

        var lockHiddenRow = addRowGroup(exclusionPanel, "left", "fill");
        lockHiddenRow.spacing = COLUMN_SPACING;

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
        setupPanel(exclusionSubPanel, 6);
        exclusionSubPanel.alignment = ["left", "fill"];

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
        setupPanel(postProcessPanel, 6);
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
     * @param {boolean} separateGuides - true のときガイドを _guide 行きに分ける。false なら通常オブジェクトと同じ扱い
     * @returns {{normal: Object[], guide: Object[]}} 通常オブジェクトとガイドの移動エントリ
     */
    function collectMoveEntries(movableItems, handledFlags, shouldMoveItem, separateGuides) {
        var moveEntries = { normal: [], guide: [] };
        for (var i = 0; i < movableItems.length; i++) {
            if (handledFlags[i]) continue;
            if (shouldMoveItem && !shouldMoveItem(i)) continue;
            var moveEntry = { item: movableItems[i], index: i };
            if (separateGuides && isGuideItem(movableItems[i])) moveEntries.guide.push(moveEntry);
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
     * すべてのアートボードを囲む矩形を返す
     * @returns {number[]} [left, top, right, bottom]
     */
    function getArtboardUnionRect() {
        var unionRect = documentArtboards[0].artboardRect.slice(0);
        for (var i = 1; i < documentArtboards.length; i++) {
            var artboardRect = documentArtboards[i].artboardRect;
            unionRect[0] = Math.min(unionRect[0], artboardRect[0]);
            unionRect[1] = Math.max(unionRect[1], artboardRect[1]);
            unionRect[2] = Math.max(unionRect[2], artboardRect[2]);
            unionRect[3] = Math.min(unionRect[3], artboardRect[3]);
        }
        return unionRect;
    }

    /**
     * ルーラーから引いたガイド（すべてのアートボードを貫く水平／垂直線）の所属アートボードを求める
     * 重心はカンバスの中央付近になり当てにならないため、線の位置（水平なら y、垂直なら x）が
     * 範囲に入るアートボードが1つだけのときにそのアートボードとする
     * @param {number[]} itemBounds - ガイドの geometricBounds [left, top, right, bottom]
     * @param {number[]} unionRect - すべてのアートボードを囲む矩形
     * @returns {number} アートボードのインデックス。決まらなければ -1、ルーラーガイドでなければ null
     */
    function findRulerGuideArtboardIndex(itemBounds, unionRect) {
        var isHorizontal = Math.abs(itemBounds[1] - itemBounds[3]) < BOUNDS_TOLERANCE &&
            itemBounds[0] <= unionRect[0] && itemBounds[2] >= unionRect[2];
        var isVertical = Math.abs(itemBounds[2] - itemBounds[0]) < BOUNDS_TOLERANCE &&
            itemBounds[1] >= unionRect[1] && itemBounds[3] <= unionRect[3];
        if (!isHorizontal && !isVertical) return null;

        var matchedIndex = -1;
        for (var i = 0; i < documentArtboards.length; i++) {
            var artboardRect = documentArtboards[i].artboardRect;
            var isInside = isHorizontal ?
                (itemBounds[1] <= artboardRect[1] && itemBounds[1] >= artboardRect[3]) :
                (itemBounds[0] >= artboardRect[0] && itemBounds[0] <= artboardRect[2]);
            if (!isInside) continue;
            if (matchedIndex >= 0) return -1;
            matchedIndex = i;
        }
        return matchedIndex;
    }

    /**
     * 各アートボードのレイヤーへ送るガイドのうち、ルーラーガイドの所属アートボードを先に求める
     * @param {PageItem[]} movableItems - 移動候補のオブジェクト配列
     * @returns {Array} インデックスごとのアートボード番号（-1 は決まらない、null はルーラーガイド以外）
     */
    function buildRulerGuideArtboardIndexes(movableItems) {
        var rulerGuideIndexes = [];
        var unionRect = getArtboardUnionRect();
        for (var i = 0; i < movableItems.length; i++) {
            rulerGuideIndexes.push(null);
            if (!isGuideItem(movableItems[i])) continue;
            try {
                rulerGuideIndexes[i] = findRulerGuideArtboardIndex(movableItems[i].geometricBounds, unionRect);
            } catch (e) {
                // 座標を取れないものは重心判定に任せる
            }
        }
        return rulerGuideIndexes;
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
     * 指定アートボードに属するかを判定するフィルター関数を作る
     * ルーラーガイドは先に求めた所属アートボードで、それ以外は重心で判定する
     * @param {number[][]} itemCentroids - 重心座標の配列
     * @param {Array} rulerGuideIndexes - ルーラーガイドの所属アートボード（null はルーラーガイド以外）
     * @param {number} artboardIndex - 判定するアートボードのインデックス
     * @returns {function(number):boolean} インデックスを受け取る判定関数
     */
    function buildArtboardFilter(itemCentroids, rulerGuideIndexes, artboardIndex) {
        var artboardRect = documentArtboards[artboardIndex].artboardRect;
        return function (itemIndex) {
            var rulerGuideIndex = rulerGuideIndexes[itemIndex];
            if (typeof rulerGuideIndex === "number") return rulerGuideIndex === artboardIndex;
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
        var separateGuides = organizeOptions.collectGuidesToGuideLayer;
        /* _guide に集めるときはガイドも重心で判定してよい（行き先は同じ） / Ruler-guide handling only matters when guides follow artboards */
        var rulerGuideIndexes = separateGuides ? [] : buildRulerGuideArtboardIndexes(movableItems);
        for (var artboardIndex = artboardRange.start; artboardIndex < artboardRange.end; artboardIndex++) {
            var artboardLayer = getOrCreateTopLevelLayer(getArtboardLayerName(artboardIndex, organizeOptions));
            var artboardFilter = buildArtboardFilter(itemCentroids, rulerGuideIndexes, artboardIndex);
            var moveEntries = collectMoveEntries(movableItems, handledFlags, artboardFilter, separateGuides);
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
        var moveEntries = collectMoveEntries(movableItems, handledFlags, null, true);
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
