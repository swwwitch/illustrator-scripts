#target illustrator
#targetengine "RelinkMissingLinksEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

リンク切れの配置画像を検出し、指定したフォルダーから同名ファイルを探して自動的に再リンクします。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RelinkMissingLinks.md

### Overview

Detects missing linked images and relinks them automatically from a folder you choose, matching by file name.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RelinkMissingLinks.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "RelinkMissingLinks";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.4.6";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-18";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RelinkMissingLinks.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RelinkMissingLinks.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 照合方法（先頭が初期値）/ Match modes; the first one is the default */
    var MATCH_MODES = [
        { labelKey: "exact", mode: "exact" },
        { labelKey: "nameOnly", mode: "nameOnly" },
        { labelKey: "preferPng", mode: "priority", ext: "png" },
        { labelKey: "preferPsd", mode: "priority", ext: "psd" },
        { labelKey: "preferJpg", mode: "priority", ext: "jpg" }
    ];

    // =========================================
    // レイアウト / Layout
    // =========================================
    var FOLDER_PANEL_MARGINS = [5, 20, 5, 10];   /* フォルダーパネルの余白 [左,上,右,下] / folder panel margins */
    var PANEL_MARGINS = [15, 20, 15, 10];        /* パネル余白 [左,上,右,下] / panel margins */
    var COLUMN_SPACING = 20;                     /* 2カラムの間隔 / gap between the two columns */
    var FOLDER_FIELD_CHARS = 30;                 /* フォルダー欄の文字数 / width of the folder field */
    var CANDIDATE_LIST_BOUNDS = [0, 0, 400, 150];  /* 候補リストの大きさ / size of the candidate list */

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

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "リンク切れの再リンク", en: "Relink Missing Links" },
            chooseCandidate: { ja: "候補を選択", en: "Choose a Candidate" }
        },
        panel: {
            folder: { ja: "再リンク用フォルダー", en: "Relink Folder" },
            matchMode: { ja: "拡張子の扱い", en: "Extension Handling" },
            target: { ja: "対象", en: "Target" }
        },
        checkbox: {
            missing: { ja: "リンク切れの画像", en: "Missing Links" },
            linked: { ja: "リンクが有効な画像", en: "Working Links" }
        },
        radio: {
            exact: { ja: "完全一致", en: "Exact Match" },
            nameOnly: { ja: "ファイル名のみ", en: "Name Only" },
            preferPng: { ja: "pngを優先", en: "Prefer PNG" },
            preferPsd: { ja: "psdを優先", en: "Prefer PSD" },
            preferJpg: { ja: "jpgを優先", en: "Prefer JPG" }
        },
        fieldLabel: {
            candidate: { ja: "再リンクするファイルを選んでください", en: "Choose the file to relink to" }
        },
        tooltip: {
            folder: { ja: "リンクし直す画像を探すフォルダーです。", en: "The folder searched for the images to relink." },
            missing: { ja: "リンク切れになっている画像を対象にします。", en: "Targets the images whose link is broken." },
            linked: {
                ja: "リンク切れでない画像も、同名のファイルが見つかれば張り替えます。",
                en: "Also relinks images that are not broken, when a file of the same name is found."
            },
            matchMode: {
                ja: "フォルダー内のどのファイルを同じ画像とみなすかの決め方です。",
                en: "How a file in the folder is matched to the linked image."
            }
        },
        button: {
            choose: { ja: "指定", en: "Choose" },
            relink: { ja: "再リンク", en: "Relink" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noFolder: { ja: "再リンク用のフォルダーを指定してください。", en: "Choose a folder to relink from." },
            invalidFolder: { ja: "有効なフォルダーを指定してください。", en: "Choose a folder that exists." },
            relinkFailed: { ja: "再リンク失敗：", en: "Relink failed: " }
        }
    };

    // =========================================
    // リンクの判定 / Link checks
    // =========================================

    /**
     * リンク切れかどうか
     * @param {PlacedItem} item - 配置画像
     * @returns {boolean} リンク切れなら true
     */
    function isLinkBroken(item) {
        try {
            if (!item.file) return true;  /* file が無い = リンク切れ / No file means a missing link */
            return !item.file.exists;
        } catch (e) {
            /* リンク切れでは file の参照が例外になる / Accessing file throws on a missing link */
            return true;
        }
    }

    /**
     * ［対象］の指定に照らして、再リンクする画像か
     * @param {PlacedItem} item - 配置画像
     * @param {Object} relinkOptions - showRelinkDialog() の結果
     * @returns {boolean} 再リンクするなら true
     */
    function shouldRelinkItem(item, relinkOptions) {
        var isBroken = isLinkBroken(item);
        return (relinkOptions.targetMissing && isBroken) || (relinkOptions.targetLinked && !isBroken);
    }

    /**
     * 照合に使う元のファイル名を返す（リンク切れで file が読めないときは XMP の名前、無ければ画像の名前）
     * @param {PlacedItem} item - 配置画像
     * @param {string} xmpFileName - XMP から取ったファイル名（無ければ空）
     * @returns {string} ファイル名。分からなければ空文字
     */
    function resolveLinkFileName(item, xmpFileName) {
        try {
            if (item.file && item.file.name) return item.file.name;
        } catch (e) {
            /* リンク切れでは file の参照が例外になる / Accessing file throws on a missing link */
            return xmpFileName || "";
        }
        if (!isLinkBroken(item)) return "";
        return xmpFileName || item.name || "";
    }

    /**
     * XMP メタデータの stRef:filePath からファイル名の一覧を作る
     * @param {Document} doc - 対象ドキュメント
     * @returns {string[]} ファイル名（XMP の並び順）
     */
    function collectXmpFileNames(doc) {
        var fileNames = [];
        var xmp;
        try {
            xmp = new XML(doc.XMPString);
        } catch (e) {
            /* XMP が読めなければ名前なしで続ける / Carry on without names when the XMP cannot be parsed */
            return fileNames;
        }
        try {
            var filePathNodes = xmp.xpath("//stRef:filePath");
            for (var i = 0; i < filePathNodes.length(); i++) {
                fileNames.push(filePathNodes[i].toString().replace(/^.*[\/\\]/, ""));
            }
        } catch (e) {
            /* 名前空間が無いなどで xpath が失敗したら名前なし / No names when xpath fails (e.g. missing namespace) */
        }
        return fileNames;
    }

    // =========================================
    // 照合 / Matching
    // =========================================

    /**
     * 拡張子を除去する
     * @param {string} fileName - ファイル名
     * @returns {string} 拡張子を除いた名前
     */
    function stripExt(fileName) {
        return fileName.replace(/\.[^\.]+$/, "");
    }

    /**
     * 候補ファイル名が元のファイル名に一致するか（名前は小文字で渡す）
     * @param {string} candidateName - 候補のファイル名
     * @param {string} originalName - 元のファイル名
     * @param {string} originalBase - 元のファイル名（拡張子なし）
     * @param {string} matchMode - "exact" / "nameOnly" / "priority"
     * @param {string} priorityExt - "priority" のときの拡張子
     * @returns {boolean} 一致すれば true
     */
    function matchCandidate(candidateName, originalName, originalBase, matchMode, priorityExt) {
        if (matchMode === "exact") {
            /* 拡張子まで完全一致 / Full name including the extension */
            return candidateName === originalName;
        } else if (matchMode === "nameOnly") {
            return stripExt(candidateName) === originalBase;
        } else if (matchMode === "priority") {
            /* ベース名＋優先拡張子と完全一致 / Base name plus the preferred extension */
            return candidateName === originalBase + "." + priorityExt;
        }
        return false;
    }

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
     * フォルダー欄と［指定］ボタンのパネルを作る
     * @param {Window} dialog - ダイアログ
     * @returns {EditText} フォルダーのパス欄
     */
    function buildFolderPanel(dialog) {
        var folderPanel = dialog.add("panel", undefined, getLabel("panel.folder"));
        folderPanel.orientation = "column";
        folderPanel.alignment = "fill";
        folderPanel.margins = FOLDER_PANEL_MARGINS;

        var folderPathInput = folderPanel.add("edittext", undefined, "");
        folderPathInput.helpTip = getLabel("tooltip.folder");
        folderPathInput.characters = FOLDER_FIELD_CHARS;

        var btnChooseFolder = folderPanel.add("button", undefined, getLabel("button.choose"));
        btnChooseFolder.onClick = function () {
            var chosenFolder = Folder.selectDialog(getLabel("panel.folder"));
            if (chosenFolder) folderPathInput.text = chosenFolder.fsName;
        };
        return folderPathInput;
    }

    /**
     * ［対象］パネルを作り、ドキュメント内のリンクの状態で初期値を決める
     * @param {Group} parent - 追加先
     * @param {Document} doc - 対象ドキュメント
     * @returns {Object} { missingCheckbox, linkedCheckbox }
     */
    function buildTargetPanel(parent, doc) {
        var targetPanel = parent.add("panel", undefined, getLabel("panel.target"));
        targetPanel.orientation = "column";
        targetPanel.alignment = "top";
        targetPanel.margins = PANEL_MARGINS;

        var missingCheckbox = targetPanel.add("checkbox", undefined, getLabel("checkbox.missing"));
        missingCheckbox.helpTip = getLabel("tooltip.missing");
        missingCheckbox.alignment = "left";
        missingCheckbox.value = true;

        var linkedCheckbox = targetPanel.add("checkbox", undefined, getLabel("checkbox.linked"));
        linkedCheckbox.helpTip = getLabel("tooltip.linked");
        linkedCheckbox.alignment = "left";

        /* リンク切れ／有効なリンクが無ければその項目を無効に / Disable a checkbox when no such link exists */
        var hasMissing = false;
        var hasLinked = false;
        for (var i = 0; i < doc.placedItems.length; i++) {
            if (isLinkBroken(doc.placedItems[i])) {
                hasMissing = true;
            } else {
                hasLinked = true;
            }
        }
        if (!hasMissing) {
            missingCheckbox.enabled = false;
            missingCheckbox.value = false;
        }
        /* 有効なリンクがあれば初期状態で ON / On by default when working links exist */
        linkedCheckbox.enabled = hasLinked;
        linkedCheckbox.value = hasLinked;

        return { missingCheckbox: missingCheckbox, linkedCheckbox: linkedCheckbox };
    }

    /**
     * ［拡張子の扱い］パネルを作る
     * @param {Group} parent - 追加先
     * @returns {RadioButton[]} MATCH_MODES と同じ並びのラジオボタン
     */
    function buildMatchModePanel(parent) {
        var matchModePanel = parent.add("panel", undefined, getLabel("panel.matchMode"));
        matchModePanel.orientation = "column";
        matchModePanel.alignment = "top";
        matchModePanel.margins = PANEL_MARGINS;

        var matchModeRadios = [];
        for (var i = 0; i < MATCH_MODES.length; i++) {
            matchModeRadios[i] = matchModePanel.add("radiobutton", undefined, getLabel("radio." + MATCH_MODES[i].labelKey));
            matchModeRadios[i].helpTip = getLabel("tooltip.matchMode");
            matchModeRadios[i].alignment = "left";
        }
        matchModeRadios[0].value = true;
        return matchModeRadios;
    }

    /**
     * ダイアログを表示し、フォルダーが有効になるまで繰り返す
     * @param {Window} dialog - ダイアログ
     * @param {EditText} folderPathInput - フォルダーのパス欄
     * @returns {Folder|null} 再リンク用フォルダー。キャンセル時は null
     */
    function showUntilValidFolder(dialog, folderPathInput) {
        while (true) {
            if (dialog.show() != 1) return null;
            if (folderPathInput.text === "") {
                alert(getLabel("alert.noFolder"));
                continue;
            }
            var targetFolder = new Folder(folderPathInput.text);
            if (!targetFolder.exists) {
                alert(getLabel("alert.invalidFolder"));
                continue;
            }
            return targetFolder;
        }
    }

    /**
     * 設定ダイアログを表示する
     * @param {Document} doc - 対象ドキュメント
     * @returns {Object|null} { targetFolder, matchMode, priorityExt, targetMissing, targetLinked }。キャンセル時は null
     */
    function showRelinkDialog(doc) {
        var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        dialog.alignChildren = "fill";

        var folderPathInput = buildFolderPanel(dialog);

        var optionColumns = dialog.add("group");
        optionColumns.orientation = "row";
        optionColumns.alignment = "center";
        optionColumns.spacing = COLUMN_SPACING;

        var matchModeRadios = buildMatchModePanel(optionColumns);
        var targetCheckboxes = buildTargetPanel(optionColumns, doc);

        var buttonRow = addButtonRow(dialog, { centered: true });
        var btnCancel = buttonRow.rowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnRelink = buttonRow.rowGroup.add("button", undefined, getLabel("button.relink"), { name: "ok" });

        prepareDialogWindow(dialog, SCRIPT_NAME);
        var targetFolder = showUntilValidFolder(dialog, folderPathInput);
        if (!targetFolder) return null;

        var selectedMode = MATCH_MODES[0];
        for (var i = 0; i < MATCH_MODES.length; i++) {
            if (matchModeRadios[i].value) {
                selectedMode = MATCH_MODES[i];
                break;
            }
        }

        return {
            targetFolder: targetFolder,
            matchMode: selectedMode.mode,
            priorityExt: selectedMode.ext,
            targetMissing: targetCheckboxes.missingCheckbox.value,
            targetLinked: targetCheckboxes.linkedCheckbox.value
        };
    }

    /**
     * 候補が複数あるとき、どれに再リンクするかを選ばせる
     * @param {File[]} candidates - 候補のファイル
     * @returns {File|null} 選んだファイル。キャンセル時は null
     */
    function chooseCandidateFile(candidates) {
        var chooseDialog = new Window("dialog", getLabel("dialog.chooseCandidate"));
        chooseDialog.alignChildren = "fill";
        chooseDialog.add("statictext", undefined, labelText("fieldLabel.candidate"));

        var candidateList = chooseDialog.add("listbox", CANDIDATE_LIST_BOUNDS);
        for (var i = 0; i < candidates.length; i++) {
            candidateList.add("item", candidates[i].name);
        }
        candidateList.selection = 0;

        var buttonRow = addButtonRow(chooseDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        prepareDialogWindow(chooseDialog, SCRIPT_NAME + "_chooseCandidate");
        if (chooseDialog.show() == 1 && candidateList.selection) {
            return candidates[candidateList.selection.index];
        }
        return null;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 1つの画像を、フォルダー内で一致したファイルへ再リンクする（候補が複数なら選ばせる）
     * @param {PlacedItem} item - 配置画像
     * @param {string} linkFileName - 照合に使う元のファイル名
     * @param {File[]} folderFiles - フォルダー内のファイル
     * @param {Object} relinkOptions - showRelinkDialog() の結果
     * @returns {void}
     */
    function relinkItem(item, linkFileName, folderFiles, relinkOptions) {
        if (!linkFileName) return;
        var originalName = linkFileName.toLowerCase();
        var originalBase = stripExt(originalName);

        var candidates = [];
        for (var i = 0; i < folderFiles.length; i++) {
            var candidateName = folderFiles[i].name.toLowerCase();
            if (matchCandidate(candidateName, originalName, originalBase, relinkOptions.matchMode, relinkOptions.priorityExt)) {
                candidates.push(folderFiles[i]);
            }
        }
        if (candidates.length === 0) return;

        var chosenFile = (candidates.length === 1) ? candidates[0] : chooseCandidateFile(candidates);
        if (!chosenFile) return;
        try {
            item.file = chosenFile;
        } catch (e) {
            alert(getLabel("alert.relinkFailed") + chosenFile.name + "\n" + e);
        }
    }

    /**
     * フォルダー内のファイル（フォルダーを除く）を返す
     * @param {Folder} targetFolder - 対象フォルダー
     * @returns {File[]} ファイル
     */
    function getFolderFiles(targetFolder) {
        var entries = targetFolder.getFiles();
        var folderFiles = [];
        for (var i = 0; i < entries.length; i++) {
            if (entries[i] instanceof File) folderFiles.push(entries[i]);
        }
        return folderFiles;
    }

    /**
     * メイン処理
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;

        var relinkOptions = showRelinkDialog(doc);
        if (!relinkOptions) return;

        var xmpFileNames = collectXmpFileNames(doc);
        var folderFiles = getFolderFiles(relinkOptions.targetFolder);

        for (var i = 0; i < doc.placedItems.length; i++) {
            var item = doc.placedItems[i];
            if (!shouldRelinkItem(item, relinkOptions)) continue;
            relinkItem(item, resolveLinkFileName(item, xmpFileNames[i]), folderFiles, relinkOptions);
        }
    }

    main();

})();
