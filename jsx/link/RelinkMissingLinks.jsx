#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
リンク切れの配置画像を検出し、ユーザーに再リンク用フォルダーを指定させ、
同名ファイルが見つかった場合に自動で再リンクを行います。

Detect missing linked images in Illustrator, prompt user to specify a folder,
and automatically relink if a file with the same name is found.

更新日 / Last Updated: 2025-08-02

更新履歴 / Update History:
- v1.0 (20250718): 初版作成 / Initial version
- v1.1 (20250730): ダイアログオプション（拡張子完全一致/ファイル名のみ/拡張子優先）追加
- v1.2 (20250731): 処理対象「すべて」を選択した場合、リンクが有効なファイルも更新可能に変更
*/

// スクリプトバージョン
var SCRIPT_VERSION = "v1.2";

function getCurrentLang() {
  return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* ラベル定義 / Label definitions */
var LABELS = {
    dialogTitle: {
        ja: "リンク切れの再リンク " + SCRIPT_VERSION,
        en: "Relink Missing Links " + SCRIPT_VERSION
    },
    folderLabel: {
        ja: "再リンク用フォルダー:",
        en: "Relink Folder:"
    },
    chooseButton: {
        ja: "フォルダー指定",
        en: "Choose Folder"
    },
    matchGroup: {
        ja: "拡張子の扱い",
        en: "Extension Handling"
    },
    targetGroup: {
        ja: "対象",
        en: "Target"
    },
    chkMissing: {
        ja: "リンク切れのみ",
        en: "Missing Links Only"
    },
    chkAll: {
        ja: "リンク変更",
        en: "Relink Existing"
    },
    ok: {
        ja: "OK",
        en: "OK"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    }
};

/* ダイアログを表示してユーザーに設定を選ばせる / Show dialog and get user settings */
function showRelinkDialog() {
    var dialog = new Window("dialog", LABELS.dialogTitle[lang]);
    dialog.alignChildren = "fill";

    var folderGroup = dialog.add("group");
    folderGroup.orientation = "column";
    folderGroup.margins = [0, 3, 0, 10];
    folderGroup.add("statictext", undefined, LABELS.folderLabel[lang]);

    var folderPath = folderGroup.add("edittext", undefined, "");
    folderPath.characters = 20;
    var chooseBtn = folderGroup.add("button", undefined, LABELS.chooseButton[lang]);
    chooseBtn.onClick = function() {
        var target = Folder.selectDialog(LABELS.folderLabel[lang]);
        if (target) folderPath.text = target.fsName;
    };

    var matchGroup = dialog.add("panel", undefined, LABELS.matchGroup[lang]);
    matchGroup.orientation = "column";
    matchGroup.alignment = "left";
    matchGroup.margins = [15, 20, 15, 10];

    // --- 対象パネル ---
    var targetGroup = dialog.add("panel", undefined, LABELS.targetGroup[lang]);
    targetGroup.orientation = "column";
    targetGroup.alignment = "left";
    targetGroup.margins = [15, 20, 15, 10];

    var chkMissingOnly = targetGroup.add("checkbox", undefined, LABELS.chkMissing[lang]);
    chkMissingOnly.alignment = "left";
    chkMissingOnly.value = true; // デフォルト

    var chkAll = targetGroup.add("checkbox", undefined, LABELS.chkAll[lang]);
    chkAll.alignment = "left";

    // --- 追加: リンク切れ/有効リンクの有無をチェック ---
    try {
        var doc = app.activeDocument;
        var hasMissing = false;
        var hasAlive = false;
        for (var i = 0; i < doc.placedItems.length; i++) {
            var pi = doc.placedItems[i];
            if (isLinkBroken(pi)) {
                hasMissing = true;
            } else {
                hasAlive = true;
            }
        }
        if (!hasMissing) {
            chkMissingOnly.enabled = false;
            chkMissingOnly.value = false;
        }
        if (!hasAlive) {
            chkAll.enabled = false;
            chkAll.value = false;
        } else {
            // リンクが生きているものがあればONにする
            chkAll.value = true;
        }
    } catch (e) {}

    var options = [
        { label: "完全一致", value: "exact" },
        { label: "ファイル名のみで調べる", value: "nameOnly" },
        { label: "pngを優先", value: "priority", ext: "png" },
        { label: "psdを優先", value: "priority", ext: "psd" },
        { label: "jpgを優先", value: "priority", ext: "jpg" }
    ];

    var radioButtons = [];
    for (var i = 0; i < options.length; i++) {
        radioButtons[i] = matchGroup.add("radiobutton", undefined, options[i].label);
        radioButtons[i].alignment = "left";
    }
    radioButtons[0].value = true;

    var btnGroup = dialog.add("group");
    btnGroup.orientation = "row";
    btnGroup.alignment = "right";
    btnGroup.add("button", undefined, LABELS.cancel[lang], {name:"cancel"});
    var okBtn = btnGroup.add("button", undefined, LABELS.ok[lang]);

    return validateDialogInput(dialog, folderPath, options, radioButtons, chkMissingOnly, chkAll);
}

// ダイアログ入力を検証 / Validate dialog input
function validateDialogInput(dialog, folderPath, options, radioButtons, chkMissingOnly, chkAll) {
    var targetFolder, mode, priorityExt;
    while (true) {
        if (dialog.show() != 1) return null;
        if (folderPath.text === "") {
            alert("再リンク用のフォルダーを指定してください。");
            continue;
        }
        targetFolder = new Folder(folderPath.text);
        if (!targetFolder.exists) {
            alert("有効なフォルダーを指定してください。");
            continue;
        }
        for (var i = 0; i < options.length; i++) {
            if (radioButtons[i].value) {
                mode = options[i].value;
                if (options[i].ext) priorityExt = options[i].ext;
                break;
            }
        }
        break;
    }
    return {
        targetFolder: targetFolder,
        mode: mode,
        priorityExt: priorityExt,
        targetMissingOnly: chkMissingOnly.value,
        targetAll: chkAll.value
    };
}

// メイン処理 / Main process
function main() {
    var dialogResult = showRelinkDialog();
    if (!dialogResult) return;

    var doc = app.activeDocument;

    // XMPメタデータからfilePathノードを抽出しファイル名リスト化
    var xmp;
    try {
        xmp = new XML(doc.XMPString);
    } catch (e) {
        xmp = null;
    }
    var paths = [];
    if (xmp) {
        try {
            var nodes = xmp.xpath('//stRef:filePath');
            for (var i = 0; i < nodes.length(); i++) {
                var pathStr = nodes[i].toString();
                var fname = pathStr.replace(/^.*[\/\\]/, "");
                paths.push(fname);
            }
        } catch (e) {}
    }

    for (var j = 0; j < doc.placedItems.length; j++) {
        var pi = doc.placedItems[j];

        var shouldRelink = false;
        if (dialogResult.targetAll) {
            shouldRelink = true;
        } else if (dialogResult.targetMissingOnly && isLinkBroken(pi)) {
            shouldRelink = true;
        }

        if (shouldRelink) {
            var fname = "";
            try {
                if (pi.file && pi.file.name) {
                    fname = pi.file.name;
                } else if (isLinkBroken(pi)) {
                    // XMPから収集したファイル名を利用
                    if (paths[j]) {
                        fname = paths[j];
                    } else if (pi.name) {
                        fname = pi.name;
                    }
                }
            } catch (e) {
                if (paths[j]) {
                    fname = paths[j];
                }
            }
            relinkSingleItem(
                pi,
                fname,
                dialogResult.targetFolder,
                dialogResult.mode,
                dialogResult.priorityExt,
                dialogResult
            );
        }
    }
}

main();

// 個別再リンク処理 / Function to relink a single missing file
function relinkSingleItem(item, brokenName, targetFolder, mode, priorityExt, dialogResult) {
    if (!(item instanceof PlacedItem)) return;

    var shouldRelink = false;

    // 「リンク切れのみ」の場合 → リンク切れのみ対象
    if (dialogResult.targetMissingOnly && isLinkBroken(item)) {
        shouldRelink = true;
    }

    // 「リンク変更」の場合 → リンクが生きている場合のみ対象
    if (dialogResult.targetAll && !isLinkBroken(item)) {
        shouldRelink = true;
    }

    if (!shouldRelink) return;

    var originalName = "";
    try {
        if (item.file && item.file.name) {
            originalName = item.file.name.toLowerCase();
        } else if (brokenName) {
            originalName = brokenName.toLowerCase();
        }
    } catch (e) {
        if (brokenName) originalName = brokenName.toLowerCase();
    }
    if (!originalName) return;

    var originalBase = stripExt(originalName);

    var filesInFolder = targetFolder.getFiles();
    var candidates = [];
    for (var k = 0; k < filesInFolder.length; k++) {
        var candidate = filesInFolder[k];
        if (!(candidate instanceof File)) continue;

        var candidateName = candidate.name.toLowerCase();
        if (matchCandidate(candidateName, originalName, originalBase, mode, priorityExt)) {
            candidates.push(candidate);
        }
    }

    if (candidates.length === 1) {
        try {
            item.file = candidates[0];
        } catch (e) {
            alert("再リンク失敗：" + candidates[0].name + "\n" + e);
        }
    } else if (candidates.length > 1) {
        var chooseDlg = new Window("dialog", "候補を選択");
        chooseDlg.alignChildren = "fill";
        chooseDlg.add("statictext", undefined, "再リンクするファイルを選んでください:");

        var list = chooseDlg.add("listbox", [0,0,400,150]);
        for (var c = 0; c < candidates.length; c++) {
            list.add("item", candidates[c].name);
        }
        list.selection = 0;

        var btnGroup = chooseDlg.add("group");
        btnGroup.alignment = "right";
        btnGroup.add("button", undefined, "キャンセル", {name:"cancel"});
        var okBtn = btnGroup.add("button", undefined, "OK");

        if (chooseDlg.show() == 1 && list.selection) {
            try {
                item.file = candidates[list.selection.index];
            } catch (e) {
                alert("再リンク失敗：" + candidates[list.selection.index].name + "\n" + e);
            }
        }
    }
}

// リンク切れ判定 / Check if link is broken
function isLinkBroken(item) {
    try {
        if (!item || !(item instanceof PlacedItem)) {
            return false; // PlacedItem以外は対象外 / Not target if not PlacedItem
        }
        if (!item.file) {
            return true; // fileプロパティが存在しない=リンク切れ / No file property means broken link
        }
        return !item.file.exists; // fileが存在しない / file does not exist
    } catch (e) {
        return true; // アクセス時例外=リンク切れ / Exception means broken link
    }
}

// 候補ファイルと比較 / Match candidate file
function matchCandidate(candidateName, brokenName, brokenBase, mode, priorityExt) {
    var candidateBase = stripExt(candidateName); // キャッシュ

    if (mode === "exact") {
        // 拡張子を含めたファイル名が完全一致の場合のみ
        return candidateName === brokenName;
    } else if (mode === "nameOnly") {
        return candidateBase === brokenBase;
    } else if (mode === "priority") {
        // ベース名に優先拡張子をつけたファイル名と完全一致
        var expectedName = brokenBase + "." + priorityExt;
        return candidateName === expectedName;
    }
    return false;
}

// 拡張子を除去 / Remove extension
function stripExt(filename) {
    return filename.replace(/\.[^\.]+$/, "");
}

// 拡張子を取得 / Get extension
function getExt(filename) {
    var match = filename.match(/\.([^\.]+)$/);
    return match ? match[1].toLowerCase() : "";
}