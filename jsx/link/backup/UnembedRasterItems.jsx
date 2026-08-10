#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

- ドキュメント内のすべての埋め込みラスター画像をリンク画像に置き換えます。
- 実行時のダイアログで、処理方法を「元ファイルに再リンク」「強制的に埋め込み解除（常にPSDを書き出し）」から選べます。
- 「元ファイルに再リンク」では、XMPマニフェストに記録された埋め込み前のファイルが残っていれば、書き出さずにそのファイルへリンクします。見つからない場合はPSDを書き出します。
- 書き出し先は、Illustratorドキュメントと同じ階層の「Links」フォルダーです（存在しない場合は自動作成）。
- 書き出すファイル名には、埋め込み前の元のファイル名を使用します。レイヤー名 → 拡張子付きの親グループ名 → XMPマニフェストの順に探し、どれも取得できない場合は image1、image2… の連番になります。
- 同名ファイルがある場合は連番の接尾辞を付けるため、既存ファイルを上書きしません。
- 配置サイズのまま書き出し、元の回転角・位置を維持したままリンク画像として再配置します。
- 書き出し解像度は配置倍率から逆算するため、元の画像の解像度（ピクセル数）が保持されます。

### 注意

- 効果（ドロップシャドウなど）が適用された埋め込み画像の変換は推奨しません。
- 効果はラスタライズされてリンク画像に含まれ、解像度は埋め込み画像側の解像度になります。
- 効果によっては、リンク画像の位置が元の埋め込み画像とずれることがあります。
- CMYK／RGB／グレースケール以外のカラースペースの画像はスキップします。

### Overview

- Replaces every embedded raster image in the document with a linked image.
- A dialog lets you choose between relinking to the original file and always exporting a PSD.
- When relinking, the pre-embed file recorded in the XMP manifest is used directly if it still exists; otherwise a PSD is exported.
- Files are written to a "Links" folder next to the Illustrator document (created automatically when missing).
- Exported files reuse the original file name from before embedding, looked up in this order: layer name, parent group name with an extension, XMP manifest; when none is available, image1, image2… is used instead.
- An incremental suffix is added when a file of the same name exists, so existing files are never overwritten.
- The image is exported at its placed size, and the linked item keeps the rotation and position of the original.
- The export resolution is derived from the placement scale, so the original pixel dimensions are preserved.

### Notes

- Unembedding raster items with effects (e.g. drop shadow) applied is not recommended.
- The effect is rasterized into the linked image, and its resolution follows the embedded image, not the document raster settings.
- Some effects shift the position of the linked image compared to the embedded image.
- Images whose color space is not CMYK, RGB or Grayscale are skipped.
- Opacity, blending mode and other appearance settings on the embedded item are not carried over to the linked item.

### オリジナルとの差異

- 書き出すファイル名に、埋め込み前の元のファイル名を使用（オリジナルは image1、image2… の連番のみ）。レイヤー名・親グループ名・XMPマニフェストの3つから探す。
- 未対応のカラースペース、書き出し失敗、拡大率を取得できない画像をスキップし、結果にまとめて表示。
- ドキュメント未オープン・未保存・埋め込み画像なしを実行前に判定して警告。
- メッセージを日本語／英語に対応。
- 効果などでサイズが変化した画像の扱いを、コード内の `if (false)` からユーザー設定 `SKIP_RESIZED_ITEMS` に変更。
- 書き出しフォルダー名・解像度などをユーザー設定として冒頭に集約。
- 元ファイルが残っている場合は書き出さずに再リンクする処理方法を追加（ダイアログで選択）。
- 処理を関数に分割し、全関数にJSDocを付与。

### Differences from the Original

- Uses the pre-embed original file name for the exported file, from the layer name, the parent group name or the XMP manifest; the original script always used image1, image2…
- Skips unsupported color spaces, failed exports and items whose scale cannot be determined, and reports them together.
- Checks for no open document, an unsaved document and documents without embedded images before running.
- Japanese / English messages.
- The handling of items resized by effects moved from an `if (false)` branch to the `SKIP_RESIZED_ITEMS` user setting.
- Export folder name, resolution and other options are collected as user settings at the top.
- Added a mode that relinks to the original file instead of exporting, chosen in a dialog.
- Split into smaller functions, each documented with JSDoc.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "UnembedRasterItems";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.4.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "m1b";                          /* 作者 / author */
var SCRIPT_MODIFIED = "Masahiro Takano (@swwwitch)";  /* 改変 / modified by */
var SCRIPT_RELEASED = "2022-07-23";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-07-27";                   /* 更新日 / last updated */

/**
 * @author m1b
 * @discussion https://community.adobe.com/t5/illustrator-discussions/is-it-possible-to-convert-rasteritem-to-placeditem/m-p/13081172
 */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/UnembedRasterItems.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/UnembedRasterItems.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ユーザー設定 / User settings
// =========================================
var LINKS_FOLDER_NAME    = "Links"; /* 書き出し先フォルダー名 / name of the export folder */
var USE_XMP_NAMES        = true;    /* レイヤー名が空のとき、XMPマニフェストの元ファイル名を使う / use the original file names from the XMP manifest when the layer name is empty */

/*
 書き出し解像度。画像は配置サイズのまま書き出すため、配置倍率から逆算した実効解像度で
 書き出すと元のピクセル数がそのまま保持される（例：24%で配置＝300ppi）。
 PRESERVE_RESOLUTION を false にすると EXPORT_RESOLUTION の固定値で書き出す（リサンプルされる）。
 / The image is exported at its placed size, so exporting at the effective resolution derived from
 the placement scale keeps the original pixel dimensions (e.g. placed at 24% = 300 ppi).
 Set PRESERVE_RESOLUTION to false to export at the fixed EXPORT_RESOLUTION instead (resamples).
*/
var PRESERVE_RESOLUTION  = true;
var EXPORT_RESOLUTION    = 300;
var SIZE_TOLERANCE       = 0.01;    /* サイズ一致とみなす許容値（pt） / tolerance for size comparison (pt) */
var SKIP_RESIZED_ITEMS   = false;   /* サイズが一致しない画像は変換せず元に戻す / revert items whose size does not match */
var KEEP_EXPORT_DOC_OPEN = false;   /* 書き出し用の一時ドキュメントを開いたままにする（動作確認用） / keep the temp document open (for testing) */

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在のUI言語を判定します。
     *
     * @returns {string} 日本語環境なら "ja"、それ以外は "en"。
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    /* 日英ラベル定義（カテゴリ分け） / Japanese-English label definitions (by category) */
    var LABELS = {
        dialog: {
            title: { ja: "埋め込み解除", en: "Unembed Raster Items" },
            scopeLabel: { ja: "処理対象", en: "Target" },
            scopeSelection: { ja: "選択している画像のみ（{count}件）", en: "Selected images only ({count})" },
            scopeAll: { ja: "すべての埋め込み画像", en: "Every embedded image" },
            modeLabel: { ja: "処理方法", en: "Method" },
            modeRelink: {
                ja: "元ファイルに再リンク（見つからない場合はPSDを書き出し）",
                en: "Relink to the original file (export a PSD when it is missing)"
            },
            modeExport: {
                ja: "強制的に埋め込み解除（常にPSDを書き出し）",
                en: "Force unembed (always export a PSD)"
            },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        message: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            unsavedDocument: {
                ja: "ドキュメントが未保存です。保存してから実行してください。",
                en: "The document has not been saved yet. Please save it before running this script."
            },
            noRasterItem: { ja: "埋め込み画像が見つかりません。", en: "No embedded raster image was found." },
            folderCreateFailed: {
                ja: "書き出し先フォルダーを作成できません：{path}",
                en: "Could not create the export folder: {path}"
            }
        },
        error: {
            invalidScale: {
                ja: "「{name}」は拡大率を取得できないためスキップしました。",
                en: "Skipped \"{name}\" because its scale could not be determined."
            },
            tooLarge: {
                ja: "「{name}」は配置サイズが {width} × {height} pt あり、アートボードの上限（{limit} pt）を超えるためスキップしました。",
                en: "Skipped \"{name}\" because its placed size of {width} x {height} pt is beyond the artboard limit ({limit} pt)."
            },
            resolutionClamped: {
                ja: "警告：「{name}」は元の解像度が {resolution} ppi と高いため、{limit} ppi に下げて書き出しました。",
                en: "Warning: \"{name}\" has a high original resolution ({resolution} ppi), so it was exported at {limit} ppi."
            },
            unsupportedColorSpace: {
                ja: "「{name}」は未対応のカラースペースのためスキップしました。({colorSpace})",
                en: "Skipped \"{name}\" because its color space is not supported. ({colorSpace})"
            },
            exportFailed: {
                ja: "「{name}」の書き出しに失敗しました。({reason})",
                en: "Failed to export \"{name}\". ({reason})"
            },
            sizeMismatchSkipped: {
                ja: "「{name}」はサイズを正しく再現できないため変換しませんでした。効果やアピアランスを解除してから再実行してください。",
                en: "Did not unembed \"{name}\" because it could not be sized correctly. Try removing effects or special appearance and run again."
            },
            sizeMismatchWarning: {
                ja: "警告：「{name}」は効果などの影響でサイズが変化しています。",
                en: "Warning: \"{name}\" has altered dimensions, probably due to an effect or special appearance."
            }
        },
        result: {
            summary: {
                ja: "{total}件中{count}件の埋め込み画像をリンクに変換しました。",
                en: "Unembedded {count} of {total} raster items."
            },
            breakdown: {
                ja: "（元ファイルに再リンク：{relinked}件／PSDを書き出し：{exported}件）",
                en: "(relinked to the original file: {relinked}, exported as PSD: {exported})"
            }
        }
    };

    /**
     * カテゴリとキーから現在の言語のラベルを取得し、{キー} を値に置換します。
     *
     * @param {string} category - LABELS のカテゴリ名。
     * @param {string} key - カテゴリ内のキー名。
     * @param {Object} [values] - ラベル内の {キー} に差し込む値の組。
     * @returns {string} 現在の言語のラベル。見つからない場合は key をそのまま返します。
     */
    function L(category, key, values) {
        var entry = LABELS[category] && LABELS[category][key],
            text = entry ? entry[currentLanguage] : key;

        /* ExtendScript では undefined を for-in に渡すと例外になる / Passing undefined to for-in throws in ExtendScript */
        if (values)
            for (var placeholder in values)
                text = text.split('{' + placeholder + '}').join(values[placeholder]);

        return text;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /* Illustratorのアートボードの上限（pt） / Illustrator's maximum artboard size (pt) */
    var MAX_ARTBOARD_SIZE = 16383;

    /* 書き出し解像度の上限（ppi） / Maximum export resolution (ppi) */
    var MAX_EXPORT_RESOLUTION = 2400;

    /* 処理方法 / Processing modes */
    var MODE_RELINK = 'relink',  /* 元ファイルが見つかればそれに再リンク / relink to the original file when it still exists */
        MODE_EXPORT = 'export';  /* 常にPSDを書き出して置き換え / always export a PSD and replace */

    /* 処理対象 / Target scopes */
    var SCOPE_SELECTION = 'selection',  /* 選択している画像のみ / selected images only */
        SCOPE_ALL = 'all';              /* すべての埋め込み画像 / every embedded image */

    /**
     * ドキュメントの状態を確認して処理を開始します。
     *
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(L('message', 'noDocument'));
            return;
        }

        var targetDoc = app.activeDocument;

        if (!targetDoc.fullName.exists) {
            alert(L('message', 'unsavedDocument'));
            return;
        }

        if (targetDoc.rasterItems.length === 0) {
            alert(L('message', 'noRasterItem'));
            return;
        }

        /* 選択範囲に含まれる埋め込み画像（グループ内も探す） / Embedded images in the selection, including inside groups */
        var selectedRasterItems = getSelectedRasterItems(targetDoc);

        /* 処理対象と処理方法をユーザーに選ばせる / Let the user choose the scope and the method */
        var options = showOptionsDialog(selectedRasterItems.length);

        if (options == null)
            return;

        var linksFolder = getLinksFolder(targetDoc);

        if (linksFolder == null) {
            alert(L('message', 'folderCreateFailed', { path: targetDoc.fullName.parent.fsName + '/' + LINKS_FOLDER_NAME }));
            return;
        }

        unembedRasterItems(targetDoc, linksFolder, options.mode, (options.scope === SCOPE_SELECTION) ? selectedRasterItems : null);
    }

    /**
     * 処理対象と処理方法を選ぶダイアログを表示します。
     *
     * @param {number} selectedItemCount - 選択範囲に含まれる埋め込み画像の件数。
     * @returns {Object|null} { scope: string, mode: string }。キャンセル時は null。
     */
    function showOptionsDialog(selectedItemCount) {

        var dlg = new Window('dialog', L('dialog', 'title') + ' ' + SCRIPT_VERSION);
        dlg.orientation = 'column';
        dlg.alignChildren = 'fill';
        dlg.margins = 16;
        dlg.spacing = 12;

        /* 処理対象 / Target scope */
        var scopePanel = dlg.add('panel', undefined, L('dialog', 'scopeLabel'));
        scopePanel.orientation = 'column';
        scopePanel.alignChildren = 'left';
        scopePanel.margins = [14, 18, 14, 14];
        scopePanel.spacing = 8;

        var selectionRadio = scopePanel.add('radiobutton', undefined, L('dialog', 'scopeSelection', { count: selectedItemCount })),
            allRadio = scopePanel.add('radiobutton', undefined, L('dialog', 'scopeAll'));

        /* 選択範囲に埋め込み画像が無ければ選べない / Disable the selection scope when nothing applicable is selected */
        selectionRadio.enabled = (selectedItemCount > 0);
        selectionRadio.value = (selectedItemCount > 0);
        allRadio.value = (selectedItemCount === 0);

        /* 処理方法 / Method */
        var modePanel = dlg.add('panel', undefined, L('dialog', 'modeLabel'));
        modePanel.orientation = 'column';
        modePanel.alignChildren = 'left';
        modePanel.margins = [14, 18, 14, 14];
        modePanel.spacing = 8;

        var relinkRadio = modePanel.add('radiobutton', undefined, L('dialog', 'modeRelink'));
        modePanel.add('radiobutton', undefined, L('dialog', 'modeExport'));

        relinkRadio.value = true;

        var buttonGroup = dlg.add('group');
        buttonGroup.alignment = 'right';
        buttonGroup.add('button', undefined, L('dialog', 'cancel'), { name: 'cancel' });
        buttonGroup.add('button', undefined, 'OK', { name: 'ok' });

        if (dlg.show() !== 1)
            return null;

        return {
            scope: selectionRadio.value ? SCOPE_SELECTION : SCOPE_ALL,
            mode: relinkRadio.value ? MODE_RELINK : MODE_EXPORT
        };
    }

    /**
     * 選択範囲に含まれる埋め込み画像を、グループの中も辿って集めます。
     *
     * @param {Document} targetDoc - 対象のIllustratorドキュメント。
     * @returns {Array<RasterItem>} 選択されている埋め込み画像。
     */
    function getSelectedRasterItems(targetDoc) {

        var foundItems = [];

        collectRasterItems(targetDoc.selection, foundItems);

        return foundItems;
    }

    /**
     * ページアイテムから埋め込み画像を再帰的に集めます。
     *
     * @param {Array<PageItem>|PageItems} pageItems - 探索するページアイテム。
     * @param {Array<RasterItem>} foundItems - 見つかった画像を追加する配列。
     * @returns {void}
     */
    function collectRasterItems(pageItems, foundItems) {

        for (var i = 0; i < pageItems.length; i++) {

            if (pageItems[i].typename === 'RasterItem')
                foundItems.push(pageItems[i]);

            else if (pageItems[i].typename === 'GroupItem')
                collectRasterItems(pageItems[i].pageItems, foundItems);
        }
    }

    /**
     * ドキュメント内のすべての埋め込み画像をPSDへ書き出し、リンク画像に置き換えます。
     *
     * @param {Document} targetDoc - 対象のIllustratorドキュメント。
     * @param {Folder} linksFolder - PSDの書き出し先フォルダー。
     * @param {string} processMode - MODE_RELINK または MODE_EXPORT。
     * @param {Array<RasterItem>|null} selectedRasterItems - 選択している画像のみ処理する場合はその配列。すべて処理する場合は null。
     * @returns {void}
     */
    function unembedRasterItems(targetDoc, linksFolder, processMode, selectedRasterItems) {

        var errorMessages = [],
            unembeddedCount = 0,
            relinkedCount = 0;

        /* 名前と元ファイルのパスは処理前にまとめて決定する（マニフェストとの突き合わせにドキュメント全体が必要）
           / Resolve every name and source path up front, since matching against the manifest needs the whole document */
        var targets = buildTargets(targetDoc, selectedRasterItems),
            embeddedItemCount = targets.length;

        var previousInteractionLevel = app.userInteractionLevel;
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

        try {
            /* 後ろから処理すると、削除してもコレクションのインデックスがずれない
               / Processing backwards keeps the collection indices valid while items are removed */
            for (var i = embeddedItemCount - 1; i >= 0; i--) {

                var sourceRasterItem = targets[i].rasterItem,
                    rasterColorSpace = sourceRasterItem.imageColorSpace,
                    exportBaseName = targets[i].baseName;

                /* 未対応のカラースペースはスキップ / Skip unsupported color spaces */
                if (!isSupportedColorSpace(rasterColorSpace)) {
                    errorMessages.push(L('error', 'unsupportedColorSpace', { name: exportBaseName, colorSpace: rasterColorSpace }));
                    continue;
                }

                var scaleAndRotation = getLinkScaleAndRotation(sourceRasterItem);

                /* 拡大率が0や不正な値の画像はスキップ / Skip items whose scale is zero or invalid */
                if (!isUsableScale(scaleAndRotation)) {
                    errorMessages.push(L('error', 'invalidScale', { name: exportBaseName }));
                    continue;
                }

                /* 埋め込み前の元ファイルが残っていれば、書き出さずにそれへ再リンク
                   / Relink to the original file instead of exporting, when it still exists */
                var originalFile = (processMode === MODE_RELINK) ? getExistingFile(targets[i].filePath) : null;

                if (originalFile != null) {
                    var relinkedItem = placeLinkedItem(sourceRasterItem, originalFile, scaleAndRotation, false);

                    if (!hasSameSize(sourceRasterItem, relinkedItem)) {
                        /* 元ファイルが差し替えられているなどでサイズが合わない場合は書き出しに切り替える
                           / Fall back to exporting when the original file no longer matches in size */
                        relinkedItem.remove();
                    }
                    else {
                        sourceRasterItem.remove();
                        unembeddedCount++;
                        relinkedCount++;
                        continue;
                    }
                }

                /* 配置サイズがアートボードの上限を超える画像はスキップ / Skip items larger than the artboard limit */
                if (sourceRasterItem.width > MAX_ARTBOARD_SIZE || sourceRasterItem.height > MAX_ARTBOARD_SIZE) {
                    errorMessages.push(L('error', 'tooLarge', {
                        name: exportBaseName,
                        width: Math.round(sourceRasterItem.width),
                        height: Math.round(sourceRasterItem.height),
                        limit: MAX_ARTBOARD_SIZE
                    }));
                    continue;
                }

                /* 配置倍率から書き出し解像度を決める / Derive the export resolution from the placement scale */
                var exportResolution = getExportResolution(scaleAndRotation);

                if (exportResolution > MAX_EXPORT_RESOLUTION) {
                    errorMessages.push(L('error', 'resolutionClamped', {
                        name: exportBaseName,
                        resolution: Math.round(exportResolution),
                        limit: MAX_EXPORT_RESOLUTION
                    }));
                    exportResolution = MAX_EXPORT_RESOLUTION;
                }

                var exportPath = getFilePathWithOverwriteProtectionSuffix(linksFolder.fsName + '/' + exportBaseName + '.psd'),
                    exportedPSDFile;

                try {
                    exportedPSDFile = exportRasterItemAsPSD(sourceRasterItem, exportPath, exportBaseName, scaleAndRotation, exportResolution);
                } catch (err) {
                    errorMessages.push(L('error', 'exportFailed', { name: exportBaseName, reason: err.message }));
                    continue;
                }

                if (!exportedPSDFile.exists)
                    continue;

                var linkedPlacedItem = placeLinkedItem(sourceRasterItem, exportedPSDFile, scaleAndRotation, true);

                /* サイズが一致しない場合は効果などが原因 / A size mismatch usually means an effect was applied */
                if (!hasSameSize(sourceRasterItem, linkedPlacedItem)) {

                    if (SKIP_RESIZED_ITEMS) {
                        errorMessages.push(L('error', 'sizeMismatchSkipped', { name: exportBaseName }));
                        exportedPSDFile.remove();
                        linkedPlacedItem.remove();
                        continue;
                    }

                    errorMessages.push(L('error', 'sizeMismatchWarning', { name: exportBaseName }));
                }

                /* 置き換えが済んだ元の埋め込み画像を削除 / Remove the embedded item that has been replaced */
                sourceRasterItem.remove();
                unembeddedCount++;
            }

        } finally {
            /* 途中で例外が出ても操作レベルを必ず戻す / Always restore the interaction level, even on failure */
            app.userInteractionLevel = previousInteractionLevel;
        }

        /* 元のドキュメントに戻して再描画してから結果を表示
           / Bring the original document back and repaint before showing the result */
        app.activeDocument = targetDoc;
        app.redraw();

        showResult(unembeddedCount, relinkedCount, embeddedItemCount, errorMessages);
    }

    /**
     * 処理結果と警告・エラーをまとめて表示します。
     *
     * @param {number} unembeddedCount - リンクに変換できた画像数。
     * @param {number} relinkedCount - 元ファイルへ再リンクした画像数。
     * @param {number} embeddedItemCount - 対象となった埋め込み画像の総数。
     * @param {Array<string>} errorMessages - 警告・エラーメッセージ。
     * @returns {void}
     */
    function showResult(unembeddedCount, relinkedCount, embeddedItemCount, errorMessages) {

        var resultMessage = L('result', 'summary', { count: unembeddedCount, total: embeddedItemCount });

        /* 再リンクと書き出しが混在した場合は内訳を出す / Show the breakdown when both methods were used */
        if (relinkedCount > 0)
            resultMessage += '\n' + L('result', 'breakdown', {
                relinked: relinkedCount,
                exported: unembeddedCount - relinkedCount
            });

        if (errorMessages.length > 0)
            resultMessage += '\n' + errorMessages.join('\n');

        alert(resultMessage);
    }

    // =========================================
    // 書き出しと再リンク / Export & relink
    // =========================================

    /**
     * ドキュメントと同じ階層の書き出し先フォルダーを取得します（なければ作成）。
     *
     * @param {Document} targetDoc - 対象のIllustratorドキュメント。
     * @returns {Folder|null} 書き出し先フォルダー。作成できない場合は null。
     */
    function getLinksFolder(targetDoc) {
        var linksFolder = Folder(targetDoc.fullName.parent.fsName + '/' + LINKS_FOLDER_NAME + '/');

        if (!linksFolder.exists && !linksFolder.create())
            return null;

        return linksFolder;
    }

    /**
     * 埋め込み画像を一時ドキュメントへ複製し、配置サイズのままPSDとして書き出します。
     * 拡大率は変えず回転だけを打ち消すため、書き出したPSDは配置時と同じ実寸になります。
     *
     * @param {RasterItem} sourceRasterItem - 書き出す埋め込み画像。
     * @param {string} exportPath - 書き出し先のフルパス。
     * @param {string} exportBaseName - 一時ドキュメントのタイトルに使う名前。
     * @param {Array<number>} scaleAndRotation - [水平拡大率%, 垂直拡大率%, 回転角°]。
     * @param {number} exportResolution - 書き出し解像度（ppi）。
     * @returns {File} 書き出したPSDファイル。
     */
    function exportRasterItemAsPSD(sourceRasterItem, exportPath, exportBaseName, scaleAndRotation, exportResolution) {

        var rasterColorSpace = sourceRasterItem.imageColorSpace,
            exportDoc = createExportDocument(exportBaseName, rasterColorSpace),
            failedStep = 'duplicate';

        try {
            var exportRasterItem = sourceRasterItem.duplicate(exportDoc.layers[0], ElementPlacement.PLACEATBEGINNING);

            /* 倍率はそのままに、回転だけ0°に戻す / Keep the scale as placed, undo the rotation only */
            failedStep = 'transform';
            if (scaleAndRotation[2] !== 0)
                exportRasterItem.transform(app.getRotationMatrix(-scaleAndRotation[2]), true, true, true, true, true);

            /* アートボードを画像サイズに合わせる / Fit the artboard to the image */
            failedStep = 'artboard';
            exportRasterItem.position = [0, exportRasterItem.height];
            exportDoc.artboards[0].artboardRect = [0, exportRasterItem.height, exportRasterItem.width, 0];

            failedStep = 'export';
            return exportAsPSD(exportDoc, exportPath, rasterColorSpace, exportResolution);

        } catch (err) {
            /* どの段階で失敗したかを結果に残す / Report which step failed */
            throw new Error(failedStep + ' — ' + err.message);

        } finally {
            /* 失敗時も一時ドキュメントを閉じる / Close the temp document even on failure */
            if (!KEEP_EXPORT_DOC_OPEN)
                exportDoc.close(SaveOptions.DONOTSAVECHANGES);
        }
    }

    /**
     * ファイルを、元の埋め込み画像と同じ位置・重ね順でリンク配置します。
     *
     * @param {RasterItem} sourceRasterItem - 元の埋め込み画像。
     * @param {File} linkFile - リンクするファイル。
     * @param {Array<number>} scaleAndRotation - [水平拡大率%, 垂直拡大率%, 回転角°]。
     * @param {boolean} isPlacedSize - 配置サイズのまま書き出したPSDなら true（拡大縮小が不要）。
     *                                 元ファイルのように原寸で配置される場合は false。
     * @returns {PlacedItem} 配置したリンク画像。
     */
    function placeLinkedItem(sourceRasterItem, linkFile, scaleAndRotation, isPlacedSize) {

        var linkedPlacedItem = sourceRasterItem.layer.placedItems.add();
        linkedPlacedItem.file = linkFile;

        /* 原寸で配置されるファイルは、元の配置倍率に合わせる / Scale files that come in at their natural size */
        var transformMatrix = isPlacedSize
            ? app.getRotationMatrix(scaleAndRotation[2])
            : app.concatenateRotationMatrix(app.getScaleMatrix(scaleAndRotation[0], scaleAndRotation[1]), scaleAndRotation[2]);

        linkedPlacedItem.transform(transformMatrix, true, true, true, true, true);

        linkedPlacedItem.move(sourceRasterItem, ElementPlacement.PLACEAFTER);
        linkedPlacedItem.position = sourceRasterItem.position;

        return linkedPlacedItem;
    }

    /**
     * パスが空でなく、実ファイルが存在する場合だけ File を返します。
     *
     * @param {string} filePath - 確認するフルパス。
     * @returns {File|null} 存在するファイル。無い場合は null。
     */
    function getExistingFile(filePath) {

        if (!filePath)
            return null;

        var linkFile = File(filePath);

        return linkFile.exists ? linkFile : null;
    }

    /**
     * 書き出し用の一時ドキュメントを作成します。
     *
     * @param {string} docTitle - ドキュメントのタイトル。
     * @param {ImageColorSpace} docColorSpace - 埋め込み画像のカラースペース。
     * @returns {Document} 作成したドキュメント。
     */
    function createExportDocument(docTitle, docColorSpace) {

        var docPreset = new DocumentPreset(),
            isCMYK = (
                docColorSpace == ImageColorSpace.CMYK
                || docColorSpace == ImageColorSpace.GrayScale
            );

        docPreset.title = docTitle;
        docPreset.width = 1000;
        docPreset.height = 1000;
        docPreset.colorMode = isCMYK ? DocumentColorSpace.CMYK : DocumentColorSpace.RGB;

        return app.documents.addDocument(isCMYK ? DocumentPresetType.BasicCMYK : DocumentPresetType.BasicRGB, docPreset);
    }

    /**
     * ドキュメントをPSDとして書き出します。
     *
     * @param {Document} sourceDoc - 書き出すドキュメント。
     * @param {string} exportPath - 書き出し先のフルパス。
     * @param {ImageColorSpace} exportColorSpace - 書き出すカラースペース。
     * @param {number} exportResolution - 書き出し解像度（ppi）。
     * @returns {File} 書き出したPSDファイル。
     */
    function exportAsPSD(sourceDoc, exportPath, exportColorSpace, exportResolution) {

        var psdFile = File(exportPath),
            psdOptions = new ExportOptionsPhotoshop();

        psdOptions.antiAliasing = false;
        psdOptions.artBoardClipping = true;
        psdOptions.imageColorSpace = exportColorSpace;
        psdOptions.editableText = false;
        psdOptions.flatten = true;
        psdOptions.maximumEditability = false;
        psdOptions.resolution = exportResolution;
        psdOptions.warnings = false;
        psdOptions.writeLayers = false;

        sourceDoc.exportFile(psdFile, ExportType.PHOTOSHOP, psdOptions);

        return psdFile;
    }

    /**
     * 既存ファイルを上書きしないパスを返します（例：myFile(2).psd）。
     *
     * @param {string} desiredPath - 希望するフルパス。
     * @returns {string} 既存ファイルと重複しないフルパス。
     */
    function getFilePathWithOverwriteProtectionSuffix(desiredPath) {

        var suffixNumber = 1,
            pathParts = desiredPath.split(/(\.[^\.]+)$/),
            uniquePath = desiredPath;

        while (File(uniquePath).exists)
            uniquePath = pathParts[0] + '(' + (++suffixNumber) + ')' + pathParts[1];

        return uniquePath;
    }

    // =========================================
    // 書き出し名の決定 / Export names
    // =========================================

    /**
     * @typedef {Object} SourceInfo
     * @property {string} baseName - 書き出しに使う拡張子なしのファイル名。
     * @property {string} filePath - 埋め込み前の元ファイルのフルパス（不明な場合は空文字）。
     */

    /**
     * @typedef {Object} UnembedTarget
     * @property {RasterItem} rasterItem - 処理する埋め込み画像。
     * @property {string} baseName - 書き出しに使う拡張子なしのファイル名。
     * @property {string} filePath - 埋め込み前の元ファイルのフルパス（不明な場合は空文字）。
     */

    /**
     * 処理対象の一覧を作ります。
     * 名前とパスの解決はドキュメント内の全画像に対して行い（マニフェストとの突き合わせに全件必要）、
     * そのうえで選択されている画像だけに絞り込みます。
     *
     * @param {Document} targetDoc - 対象のIllustratorドキュメント。
     * @param {Array<RasterItem>|null} selectedRasterItems - 選択している画像のみ処理する場合はその配列。すべて処理する場合は null。
     * @returns {Array<UnembedTarget>} 処理対象の一覧。
     */
    function buildTargets(targetDoc, selectedRasterItems) {

        var allRasterItems = targetDoc.rasterItems,
            sourceInfo = resolveSourceInfo(targetDoc, allRasterItems),
            targets = [],
            i;

        if (selectedRasterItems == null) {

            for (i = 0; i < allRasterItems.length; i++)
                targets.push({
                    rasterItem: allRasterItems[i],
                    baseName: sourceInfo[i].baseName,
                    filePath: sourceInfo[i].filePath
                });

            return targets;
        }

        /* 選択された画像を、ドキュメント全体で解決した情報と突き合わせる / Match the selected items against the document-wide info */
        var infoByUUID = {};

        for (i = 0; i < allRasterItems.length; i++)
            infoByUUID[allRasterItems[i].uuid] = sourceInfo[i];

        for (i = 0; i < selectedRasterItems.length; i++) {
            var info = infoByUUID[selectedRasterItems[i].uuid];

            targets.push({
                rasterItem: selectedRasterItems[i],
                baseName: info ? info.baseName : 'image' + (i + 1),
                filePath: info ? info.filePath : ''
            });
        }

        return targets;
    }

    /**
     * すべての埋め込み画像について、書き出し名と元ファイルのパスを処理前にまとめて決定します。
     * 名前はレイヤー名 → 親グループ名 → XMPマニフェスト → 連番 の順に探します。
     *
     * @param {Document} targetDoc - 対象のIllustratorドキュメント。
     * @param {RasterItems} embeddedRasterItems - ドキュメント内の埋め込み画像。
     * @returns {Array<SourceInfo>} 埋め込み画像と同じ並びの情報。
     */
    function resolveSourceInfo(targetDoc, embeddedRasterItems) {

        var itemCount = embeddedRasterItems.length,
            fileNames = [],
            filePaths = [],
            sourceInfo = [],
            i;

        /* アイテム自身から取得できる名前（拡張子付きのまま） / Names available from the item itself (extension kept) */
        for (i = 0; i < itemCount; i++) {
            fileNames.push(getFileNameFromItem(embeddedRasterItems[i]));
            filePaths.push('');
        }

        /* 名前とパスをXMPマニフェストで補う / Fill in names and paths from the XMP manifest */
        if (USE_XMP_NAMES)
            fillFromManifest(targetDoc, fileNames, filePaths);

        for (i = 0; i < itemCount; i++)
            sourceInfo.push({
                baseName: toSafeBaseName(fileNames[i], 'image' + (i + 1)),
                filePath: filePaths[i]
            });

        return sourceInfo;
    }

    /**
     * 埋め込み画像自身が持つ名前を返します。
     * レイヤー名、無ければ拡張子付きの親グループ名（配置時にファイル名が残ることがある）。
     *
     * @param {RasterItem} rasterItem - 対象の埋め込み画像。
     * @returns {string} 拡張子付きの名前。取得できない場合は空文字。
     */
    function getFileNameFromItem(rasterItem) {

        if (rasterItem.name)
            return rasterItem.name;

        var parentItem = rasterItem.parent;

        if (parentItem == undefined || parentItem.typename !== 'GroupItem')
            return '';

        var parentName = parentItem.name || '';

        return /\.[a-z][a-z0-9]{1,4}\s*$/i.test(parentName) ? parentName : '';
    }

    /**
     * XMPマニフェストの情報で、名前と元ファイルのパスを補います。
     * 名前が判明している画像には同名のエントリからパスだけを与え、名前が無い画像には
     * 残った候補を割り当てます。割り当ては、候補の件数が名前未定の画像数と一致するときだけ行い、
     * 一致しない場合は取り違えを避けて何もしません。
     *
     * @param {Document} targetDoc - 対象のIllustratorドキュメント。
     * @param {Array<string>} fileNames - 画像ごとの名前。空文字の要素が埋められます。
     * @param {Array<string>} filePaths - 画像ごとの元ファイルのパス。判明した要素が埋められます。
     * @returns {void}
     */
    function fillFromManifest(targetDoc, fileNames, filePaths) {

        var manifestPaths = getManifestFilePaths(targetDoc);

        if (manifestPaths.length === 0)
            return;

        /* 重複を除いた候補と、ファイル名からパスを引くための索引 / Unique candidates, plus a name-to-path index */
        var candidatePaths = [],
            pathByName = {},
            foundNames = {},
            nameKey,
            i;

        for (i = 0; i < manifestPaths.length; i++) {
            nameKey = getFileNameFromPath(manifestPaths[i]).toLowerCase();

            if (nameKey === '' || foundNames[nameKey])
                continue;

            foundNames[nameKey] = true;
            pathByName[nameKey] = manifestPaths[i];
            candidatePaths.push(manifestPaths[i]);
        }

        /* 名前が判明している画像は、同名のエントリからパスだけ補う（順番に依存しない安全な照合）
           / Items with a known name get their path by name match, which does not depend on order */
        var knownNames = {},
            missingCount = 0;

        for (i = 0; i < fileNames.length; i++) {

            if (fileNames[i] === '') {
                missingCount++;
                continue;
            }

            nameKey = fileNames[i].toLowerCase();
            knownNames[nameKey] = true;

            if (pathByName[nameKey])
                filePaths[i] = pathByName[nameKey];
        }

        if (missingCount === 0)
            return;

        /* 名前未定の画像に割り当てる候補 / Candidates left for the unnamed items */
        var remainingPaths = [];

        for (i = 0; i < candidatePaths.length; i++)
            if (!knownNames[getFileNameFromPath(candidatePaths[i]).toLowerCase()])
                remainingPaths.push(candidatePaths[i]);

        /* 件数が一致しないときは、取り違えを避けるため使わない / Skip when the counts disagree, to avoid mismatched names */
        if (remainingPaths.length !== missingCount)
            return;

        var pathIndex = 0;

        for (i = 0; i < fileNames.length; i++) {
            if (fileNames[i] !== '')
                continue;

            filePaths[i] = remainingPaths[pathIndex++];
            fileNames[i] = getFileNameFromPath(filePaths[i]);
        }
    }

    /**
     * フルパスからファイル名部分を取り出します。
     *
     * @param {string} filePath - フルパス。
     * @returns {string} URLエンコードを解いたファイル名。
     */
    function getFileNameFromPath(filePath) {
        return decodeURI(String(filePath).replace(/^.*[\/\\]/, ''));
    }

    /**
     * 名前から拡張子と使用できない文字を取り除き、ファイル名として整えます。
     *
     * @param {string} fileName - 元の名前（拡張子付きでも可）。
     * @param {string} fallbackName - 名前が空になる場合に使う代替名。
     * @returns {string} ファイル名に使用できる文字列。
     */
    function toSafeBaseName(fileName, fallbackName) {

        /* 拡張子と前後の空白を除去（「2026.07.27」のような名前を壊さないよう英字始まりの2〜5文字に限定）
           / Strip the extension and surrounding whitespace (limited to 2-5 chars starting with a letter, so names like "2026.07.27" survive) */
        var baseName = (fileName || '').replace(/\.[a-z][a-z0-9]{1,4}$/i, '').replace(/^\s+|\s+$/g, '');

        /* ファイル名に使えない文字を置換 / Replace characters that are illegal in file names */
        baseName = baseName.replace(/[\\\/:*?"<>|]/g, '_');

        return baseName === '' ? fallbackName : baseName;
    }

    /**
     * ドキュメントのXMPマニフェストから、埋め込み前の元ファイルのパスを出現順に取得します。
     *
     * @param {Document} targetDoc - 対象のIllustratorドキュメント。
     * @returns {Array<string>} 元ファイルのフルパス。取得できない場合は空配列。
     */
    function getManifestFilePaths(targetDoc) {

        var manifestPaths = [],
            filePaths;

        try {
            var xmp = new XML(targetDoc.XMPString);

            /* 埋め込み参照のみを対象にし、取得できなければすべての参照を見る
               / Prefer manifest references only, and fall back to every reference */
            filePaths = xmp.xpath('//stMfs:reference/stRef:filePath');

            if (filePaths == null || filePaths.length() === 0)
                filePaths = xmp.xpath('//stRef:filePath');

        } catch (err) {
            return manifestPaths;
        }

        if (filePaths == null)
            return manifestPaths;

        for (var i = 0; i < filePaths.length(); i++) {
            var filePath = decodeURI(String(filePaths[i]));

            if (filePath !== '')
                manifestPaths.push(filePath);
        }

        return manifestPaths;
    }

    /**
     * PSDとして書き出せるカラースペースかどうかを判定します。
     *
     * @param {ImageColorSpace} imageColorSpace - 判定するカラースペース。
     * @returns {boolean} CMYK／RGB／グレースケールなら true。
     */
    function isSupportedColorSpace(imageColorSpace) {
        return (
            imageColorSpace == ImageColorSpace.CMYK
            || imageColorSpace == ImageColorSpace.RGB
            || imageColorSpace == ImageColorSpace.GrayScale
        );
    }

    /**
     * 配置画像・埋め込み画像の拡大率と回転角を返します。
     *
     * @author m1b
     * @param {PlacedItem|RasterItem} linkItem - 対象のアイテム。
     * @returns {Array<number>} [水平拡大率%, 垂直拡大率%, 回転角°]。
     */
    function getLinkScaleAndRotation(linkItem) {

        var itemMatrix = linkItem.matrix,
            flipPlacedItem = (linkItem.typename == 'PlacedItem') ? 1 : -1,
            rotatedAmount = 0;

        /* 回転角はタグとして保持されている（未回転の場合はタグなし） / The rotation is stored as a tag (absent when never rotated) */
        try {
            rotatedAmount = linkItem.tags.getByName('BBAccumRotation').value * 180 / Math.PI;
        } catch (err) { }

        var unrotatedMatrix = app.concatenateRotationMatrix(itemMatrix, rotatedAmount * flipPlacedItem);

        return [
            unrotatedMatrix.mValueA * 100,
            unrotatedMatrix.mValueD * -100 * flipPlacedItem,
            rotatedAmount
        ];
    }

    /**
     * 取得した拡大率が書き出しに使える値かどうかを判定します。
     * 0や NaN のままだと等倍に戻す計算が破綻するため、事前に弾きます。
     *
     * @param {Array<number>} scaleAndRotation - [水平拡大率%, 垂直拡大率%, 回転角°]。
     * @returns {boolean} 縦横とも0以外の有限値なら true。
     */
    function isUsableScale(scaleAndRotation) {
        return (
            scaleAndRotation[0] != 0 && isFinite(scaleAndRotation[0])
            && scaleAndRotation[1] != 0 && isFinite(scaleAndRotation[1])
        );
    }

    /**
     * 書き出し解像度（ppi）を返します。
     * 配置サイズのまま書き出すため、配置倍率から逆算した実効解像度で書き出すと
     * 元のピクセル数がそのまま保持されます（倍率24% → 7200÷24 = 300ppi）。
     *
     * @param {Array<number>} scaleAndRotation - [水平拡大率%, 垂直拡大率%, 回転角°]。
     * @returns {number} 書き出し解像度（ppi）。
     */
    function getExportResolution(scaleAndRotation) {

        if (!PRESERVE_RESOLUTION)
            return EXPORT_RESOLUTION;

        return 72 * 100 / Math.abs(scaleAndRotation[0]);
    }

    /**
     * 2つのアイテムの幅と高さが許容値内で一致するかを判定します。
     *
     * @param {PageItem} itemA - 比較するアイテム。
     * @param {PageItem} itemB - 比較するアイテム。
     * @returns {boolean} 幅と高さがともに一致すれば true。
     */
    function hasSameSize(itemA, itemB) {
        return (
            Math.abs(itemA.width - itemB.width) < SIZE_TOLERANCE
            && Math.abs(itemA.height - itemB.height) < SIZE_TOLERANCE
        );
    }

    main();

})();
