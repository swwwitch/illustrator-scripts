#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
リンク切れの配置画像を検出し、ユーザーに再リンク用フォルダーを指定させ、
同名ファイルが見つかった場合に自動で再リンクを行います。

Detect missing linked images in Illustrator, prompt user to specify a folder,
and automatically relink if a file with the same name is found.

更新履歴 / Update History:
- v1.0 (20250718): 初版作成 / Initial version
- v1.1 (20250730): ダイアログオプション（拡張子完全一致/ファイル名のみ/拡張子優先）追加
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
        ja: "すべて",
        en: "All"
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

    // --- ここから追加: 対象パネル ---
    var targetGroup = dialog.add("panel", undefined, LABELS.targetGroup[lang]);
    targetGroup.orientation = "column";
    targetGroup.alignment = "left";
    targetGroup.margins = [15, 20, 15, 10];

    var chkMissingOnly = targetGroup.add("checkbox", undefined, LABELS.chkMissing[lang]);
    chkMissingOnly.alignment = "left";
    chkMissingOnly.value = true; // デフォルト

    var chkAll = targetGroup.add("checkbox", undefined, LABELS.chkAll[lang]);
    chkAll.alignment = "left";
    // --- ここまで追加 ---

    var options = [{
            label: "完全一致",
            value: "exact"
        },
        {
            label: "ファイル名のみで調べる",
            value: "nameOnly"
        },
        {
            label: "pngを優先",
            value: "priority",
            ext: "png"
        },
        {
            label: "psdを優先",
            value: "priority",
            ext: "psd"
        },
        {
            label: "jpgを優先",
            value: "priority",
            ext: "jpg"
        }
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
    var x = new XML(doc.XMPString);
    var m = x.xpath('//stRef:filePath');

    if (m !== '') {
        for (var i = 0, len = m.length(); i < len; i++) {
            var pathStr = m[i].toString();
            var fname = pathStr.replace(/^.*[\/\\]/, "");
            var fileObj = File(pathStr);
            // ユーザーの選択に応じて処理対象を分岐
            if (
                dialogResult.targetAll ||
                (!fileObj.exists && dialogResult.targetMissingOnly)
            ) {
                relinkSingleItem(
                    doc.placedItems,
                    fname,
                    dialogResult.targetFolder,
                    dialogResult.mode,
                    dialogResult.priorityExt
                );
            }
        }
    }
}

main();

// 個別再リンク処理 / Function to relink a single missing file
// 個別再リンク処理 / Function to relink a single missing file
function relinkSingleItem(items, brokenName, targetFolder, mode, priorityExt) {
    var filesInFolder = targetFolder.getFiles();

    var brokenNameLower = brokenName.toLowerCase();
    var brokenBase = stripExt(brokenNameLower);

    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        if (!(item instanceof PlacedItem)) continue;

        if (isLinkBroken(item)) {
            var candidates = [];
            for (var k = 0; k < filesInFolder.length; k++) {
                var candidate = filesInFolder[k];
                if (!(candidate instanceof File)) continue;

                var candidateName = candidate.name.toLowerCase();
                if (matchCandidate(candidateName, brokenNameLower, brokenBase, mode, priorityExt)) {
                    candidates.push(candidate);
                }
            }

            if (candidates.length === 1) {
                // 候補が1件 → 即決
                try {
                    item.file = candidates[0];
                } catch (e) {
                    alert("再リンク失敗：" + candidates[0].name + "\n" + e);
                }
                return;
            } else if (candidates.length > 1) {
                // 候補が複数 → ダイアログで選択
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
                return;
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