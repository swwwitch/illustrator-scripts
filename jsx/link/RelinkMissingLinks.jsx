#target illustrator
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
var SCRIPT_VERSION  = "v1.4.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-18";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

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

    /**
     * UI言語を返す
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

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

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "panel.folder" のようなパス
     * @returns {string} 表示言語のテキスト
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        return LABELS[labelPathKeys[0]][labelPathKeys[1]][uiLang];
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
    }

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

        var btnRowGroup = dialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = "center";
        btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnRowGroup.add("button", undefined, getLabel("button.relink"), { name: "ok" });

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

        var btnRowGroup = chooseDialog.add("group");
        btnRowGroup.alignment = "right";
        btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

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
