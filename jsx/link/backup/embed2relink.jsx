#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

// =========================================
// ユーザー設定 / User settings
// =========================================
var USE_XMP_PATHS  = true;  /* 元ファイルのパスをXMPマニフェストからも探す / also look up the original paths in the XMP manifest */
var SIZE_TOLERANCE = 0.01;  /* サイズ一致とみなす許容値（pt） / tolerance for size comparison (pt) */

(function() {

    // =========================================
    // 元ファイルのパス解決 / Source path resolution
    // =========================================

    /**
     * ドキュメントのXMPマニフェストから、埋め込み前の元ファイルのパスを出現順に取得する
     * Collect the pre-embed source paths from the document's XMP manifest, in order
     * @param {Document} targetDoc - 対象のIllustratorドキュメント
     * @returns {Array<string>} 元ファイルのフルパス。取得できない場合は空配列
     */
    function getManifestFilePaths(targetDoc) {

        var manifestPaths = [],
            filePaths;

        try {
            var xmp = new XML(targetDoc.XMPString);

            /* 埋め込み参照のみを対象にし、取得できなければすべての参照を見る
               Prefer manifest references only, and fall back to every reference */
            filePaths = xmp.xpath('//stMfs:reference/stRef:filePath');

            if (filePaths == null || filePaths.length() === 0) {
                filePaths = xmp.xpath('//stRef:filePath');
            }

        } catch (err) {
            return manifestPaths;
        }

        if (filePaths == null) {
            return manifestPaths;
        }

        for (var i = 0; i < filePaths.length(); i++) {
            var filePath = decodeURI(String(filePaths[i]));

            if (filePath !== "") {
                manifestPaths.push(filePath);
            }
        }

        return manifestPaths;
    }

    /**
     * フルパスからファイル名部分を取り出す
     * Extract the file name from a full path
     * @param {string} filePath - フルパス
     * @returns {string} URLエンコードを解いたファイル名
     */
    function getFileNameFromPath(filePath) {
        return decodeURI(String(filePath).replace(/^.*[\/\\]/, ""));
    }

    /**
     * 埋め込み画像自身が持つ名前を返す
     * レイヤー名、無ければ拡張子付きの親グループ名（配置時にファイル名が残ることがある）
     * Get the name carried by the embedded image itself
     * @param {RasterItem} rasterItem - 対象の埋め込み画像
     * @returns {string} 拡張子付きの名前。取得できない場合は空文字
     */
    function getFileNameFromItem(rasterItem) {

        if (rasterItem.name) {
            return rasterItem.name;
        }

        var parentItem = rasterItem.parent;

        if (parentItem == undefined || parentItem.typename !== "GroupItem") {
            return "";
        }

        var parentName = parentItem.name || "";

        return /\.[a-z][a-z0-9]{1,4}\s*$/i.test(parentName) ? parentName : "";
    }

    /**
     * XMPマニフェストの情報で、名前と元ファイルのパスを補う
     * 名前が判明している画像には同名のエントリからパスだけを与え、名前が無い画像には
     * 残った候補を割り当てる。割り当ては候補の件数が名前未定の画像数と一致するときだけ行い、
     * 一致しない場合は取り違えを避けて何もしない
     * Fill in names and source paths from the XMP manifest
     * @param {Document} targetDoc - 対象のIllustratorドキュメント
     * @param {Array<string>} fileNames - 画像ごとの名前。空文字の要素が埋められる
     * @param {Array<string>} filePaths - 画像ごとの元ファイルのパス。判明した要素が埋められる
     * @returns {void}
     */
    function fillFromManifest(targetDoc, fileNames, filePaths) {

        var manifestPaths = getManifestFilePaths(targetDoc);

        if (manifestPaths.length === 0) {
            return;
        }

        /* 重複を除いた候補と、ファイル名からパスを引くための索引 / Unique candidates, plus a name-to-path index */
        var candidatePaths = [],
            pathByName = {},
            foundNames = {},
            nameKey,
            i;

        for (i = 0; i < manifestPaths.length; i++) {
            nameKey = getFileNameFromPath(manifestPaths[i]).toLowerCase();

            if (nameKey === "" || foundNames[nameKey]) {
                continue;
            }

            foundNames[nameKey] = true;
            pathByName[nameKey] = manifestPaths[i];
            candidatePaths.push(manifestPaths[i]);
        }

        /* 名前が判明している画像は、同名のエントリからパスだけ補う（順番に依存しない安全な照合）
           Items with a known name get their path by name match, which does not depend on order */
        var knownNames = {},
            missingCount = 0;

        for (i = 0; i < fileNames.length; i++) {

            if (fileNames[i] === "") {
                missingCount++;
                continue;
            }

            nameKey = fileNames[i].toLowerCase();
            knownNames[nameKey] = true;

            if (pathByName[nameKey]) {
                filePaths[i] = pathByName[nameKey];
            }
        }

        if (missingCount === 0) {
            return;
        }

        /* 名前未定の画像に割り当てる候補 / Candidates left for the unnamed items */
        var remainingPaths = [];

        for (i = 0; i < candidatePaths.length; i++) {
            if (!knownNames[getFileNameFromPath(candidatePaths[i]).toLowerCase()]) {
                remainingPaths.push(candidatePaths[i]);
            }
        }

        /* 件数が一致しないときは、取り違えを避けるため使わない / Skip when the counts disagree, to avoid mismatched paths */
        if (remainingPaths.length !== missingCount) {
            return;
        }

        var pathIndex = 0;

        for (i = 0; i < fileNames.length; i++) {
            if (fileNames[i] !== "") {
                continue;
            }

            filePaths[i] = remainingPaths[pathIndex++];
            fileNames[i] = getFileNameFromPath(filePaths[i]);
        }
    }

    /**
     * ドキュメント内のすべての埋め込み画像について、元ファイルのパスをUUID索引で返す
     * 照合はドキュメント全体で行う必要があるため、選択の有無にかかわらず全件を解決する
     * Resolve the source path of every embedded image, indexed by UUID
     * @param {Document} targetDoc - 対象のIllustratorドキュメント
     * @returns {Object} UUIDをキー、元ファイルのフルパスを値とする索引
     */
    function resolveSourcePathsByUUID(targetDoc) {

        var allRasterItems = targetDoc.rasterItems,
            itemCount = allRasterItems.length,
            fileNames = [],
            filePaths = [],
            pathByUUID = {},
            i;

        /* アイテム自身から取得できる名前（拡張子付きのまま） / Names available from the item itself (extension kept) */
        for (i = 0; i < itemCount; i++) {
            fileNames.push(getFileNameFromItem(allRasterItems[i]));
            filePaths.push("");
        }

        if (USE_XMP_PATHS) {
            fillFromManifest(targetDoc, fileNames, filePaths);
        }

        for (i = 0; i < itemCount; i++) {
            pathByUUID[allRasterItems[i].uuid] = filePaths[i];
        }

        return pathByUUID;
    }

    /**
     * パスが空でなく、実ファイルが存在する場合だけ File を返す
     * Return a File only when the path is non-empty and the file exists
     * @param {string} filePath - 確認するフルパス
     * @returns {File} 存在するファイル。無い場合は null
     */
    function getExistingFile(filePath) {

        if (!filePath) {
            return null;
        }

        var linkFile = File(filePath);

        return linkFile.exists ? linkFile : null;
    }

    /**
     * 埋め込み画像の元ファイルを取得する
     * アイテムの file プロパティを優先し、無ければXMPマニフェスト由来のパスを使う
     * Get the source file of an embedded raster item
     * @param {RasterItem} item - 対象の埋め込み画像
     * @param {Object} pathByUUID - UUIDから元ファイルのパスを引く索引
     * @returns {File} 元ファイル。取得できない場合は null
     */
    function getSourceFile(item, pathByUUID) {

        try {
            if (item.file && item.file.exists) {
                return item.file;
            }
        } catch (e) {
            /* file プロパティ参照エラー時はマニフェスト側を見る / Fall back to the manifest on property error */
        }

        return getExistingFile(pathByUUID[item.uuid]);
    }

    // =========================================
    // 再リンク処理 / Relink
    // =========================================

    /**
     * 選択オブジェクトから埋め込み画像（RasterItem）を再帰的に収集する
     * Collect embedded raster items recursively
     * @param {Array} items - 走査対象のオブジェクト配列（selection や pageItems）
     * @param {Array<RasterItem>} result - 収集結果を格納する配列
     * @returns {Array<RasterItem>} 収集した埋め込み画像の配列
     */
    function collectRasterItems(items, result) {
        for (var i = 0; i < items.length; i++) {
            var obj = items[i];
            if (obj.typename === "RasterItem") {
                result.push(obj);
            } else if (obj.typename === "GroupItem") {
                /* グループ内も走査 / Traverse inside groups */
                collectRasterItems(obj.pageItems, result);
            }
        }
        return result;
    }

    /**
     * 配置画像・埋め込み画像の拡大率と回転角を返す
     * Get the scale and rotation of a placed or embedded image
     * @author m1b
     * @param {PlacedItem|RasterItem} linkItem - 対象のアイテム
     * @returns {Array<number>} [水平拡大率%, 垂直拡大率%, 回転角°]
     */
    function getLinkScaleAndRotation(linkItem) {

        var itemMatrix = linkItem.matrix,
            flipPlacedItem = (linkItem.typename == "PlacedItem") ? 1 : -1,
            rotatedAmount = 0;

        /* 回転角はタグとして保持されている（未回転の場合はタグなし） / The rotation is stored as a tag (absent when never rotated) */
        try {
            rotatedAmount = linkItem.tags.getByName("BBAccumRotation").value * 180 / Math.PI;
        } catch (err) {}

        var unrotatedMatrix = app.concatenateRotationMatrix(itemMatrix, rotatedAmount * flipPlacedItem);

        return [
            unrotatedMatrix.mValueA * 100,
            unrotatedMatrix.mValueD * -100 * flipPlacedItem,
            rotatedAmount
        ];
    }

    /**
     * 取得した拡大率が再配置に使える値かどうかを判定する
     * 0や NaN のままだと元の配置サイズに合わせる計算が破綻するため、事前に弾く
     * Check whether the scale can be used to place the link
     * @param {Array<number>} scaleAndRotation - [水平拡大率%, 垂直拡大率%, 回転角°]
     * @returns {boolean} 縦横とも0以外の有限値なら true
     */
    function isUsableScale(scaleAndRotation) {
        return (
            scaleAndRotation[0] != 0 && isFinite(scaleAndRotation[0])
            && scaleAndRotation[1] != 0 && isFinite(scaleAndRotation[1])
        );
    }

    /**
     * 2つのアイテムの幅と高さが許容値内で一致するかを判定する
     * Check whether two items have the same width and height
     * @param {PageItem} itemA - 比較するアイテム
     * @param {PageItem} itemB - 比較するアイテム
     * @returns {boolean} 幅と高さがともに一致すれば true
     */
    function hasSameSize(itemA, itemB) {
        return (
            Math.abs(itemA.width - itemB.width) < SIZE_TOLERANCE
            && Math.abs(itemA.height - itemB.height) < SIZE_TOLERANCE
        );
    }

    /**
     * ファイルを、元の埋め込み画像と同じ位置・重ね順・配置倍率でリンク配置する
     * Place a file as a link, matching the embedded image's position, stacking order and scale
     * @param {RasterItem} sourceRasterItem - 元の埋め込み画像
     * @param {File} linkFile - リンクするファイル
     * @param {Array<number>} scaleAndRotation - [水平拡大率%, 垂直拡大率%, 回転角°]
     * @returns {PlacedItem} 配置したリンク画像
     */
    function placeLinkedItem(sourceRasterItem, linkFile, scaleAndRotation) {

        var linkedPlacedItem = sourceRasterItem.layer.placedItems.add();
        linkedPlacedItem.file = linkFile;

        /* リンクファイルは原寸で配置されるため、元の配置倍率と回転に合わせる
           The link comes in at its natural size, so match the original scale and rotation */
        var transformMatrix = app.concatenateRotationMatrix(
            app.getScaleMatrix(scaleAndRotation[0], scaleAndRotation[1]),
            scaleAndRotation[2]
        );

        linkedPlacedItem.transform(transformMatrix, true, true, true, true, true);

        linkedPlacedItem.move(sourceRasterItem, ElementPlacement.PLACEAFTER);
        linkedPlacedItem.position = sourceRasterItem.position;

        return linkedPlacedItem;
    }

    /**
     * 埋め込み画像をリンク画像に置き換える
     * Replace an embedded image with a linked one
     * @param {RasterItem} item - 置き換え元の埋め込み画像
     * @param {File} targetFile - リンク先のファイル
     * @returns {{placedItem: PlacedItem, sizeChanged: boolean}} 置き換え後のリンク画像とサイズ差の有無
     */
    function relinkItem(item, targetFile) {

        var scaleAndRotation = getLinkScaleAndRotation(item);

        if (!isUsableScale(scaleAndRotation)) {
            throw new Error("配置倍率を取得できませんでした。");
        }

        var placedItem = placeLinkedItem(item, targetFile, scaleAndRotation),
            sizeChanged = !hasSameSize(item, placedItem);

        item.remove();

        return { placedItem: placedItem, sizeChanged: sizeChanged };
    }

    /**
     * メイン処理：選択中の埋め込み画像をすべてリンク画像に置き換える
     * Main routine: relink all selected embedded images
     * @returns {void}
     */
    function main() {
        /* ドキュメントの存在確認 / Check for an open document */
        if (app.documents.length === 0) {
            alert("ドキュメントが開かれていません。");
            return;
        }

        var doc = app.activeDocument;
        var sel = doc.selection;

        /* 選択チェック / Check the selection */
        if (!sel || sel.length === 0) {
            alert("画像を選択してください。");
            return;
        }

        var rasterItems = collectRasterItems(sel, []);

        if (rasterItems.length === 0) {
            /* 埋め込み画像がない場合、リンク画像かどうかで案内を分ける / Branch the message */
            var placedCount = 0;
            for (var i = 0; i < sel.length; i++) {
                if (sel[i].typename === "PlacedItem") {
                    placedCount++;
                }
            }
            if (placedCount > 0) {
                alert("選択されている画像はすでに「リンク画像」です。");
            } else {
                alert("埋め込み画像を選択してください。");
            }
            return;
        }

        /* 元パスの照合はドキュメント全体で行う必要があるため、処理前にまとめて解決する
           The manifest must be matched across the whole document, so resolve everything up front */
        var pathByUUID = resolveSourcePathsByUUID(doc);

        var relinkedFiles = [];
        var skippedCount = 0;
        var errorMessages = [];
        var warningMessages = [];
        var newItems = [];
        var askManual = true;

        for (var j = 0; j < rasterItems.length; j++) {
            var item = rasterItems[j];
            var itemLabel = (item.name && item.name !== "") ? item.name : "画像 " + (j + 1);
            var targetFile = getSourceFile(item, pathByUUID);

            /* 元ファイルが見つからない場合はダイアログで手動選択させる / Ask for manual selection */
            if (!targetFile && askManual) {
                var userChoice = confirm(
                    "［" + itemLabel + "］（" + (j + 1) + "/" + rasterItems.length + "）\n" +
                    "元のファイル情報が残っていないか、ファイルが見つかりません。\n" +
                    "手動でファイルを選択して再リンクしますか？\n\n" +
                    "「いいえ」を選ぶと、以降の画像も手動選択せずにスキップします。"
                );
                if (userChoice) {
                    targetFile = File.openDialog("［" + itemLabel + "］に再リンクする画像ファイルを選択してください");
                } else {
                    askManual = false;
                }
            }

            /* ファイルが決まらなかった画像はスキップ / Skip items without a target file */
            if (!targetFile) {
                skippedCount++;
                continue;
            }

            try {
                var result = relinkItem(item, targetFile);

                newItems.push(result.placedItem);
                relinkedFiles.push(targetFile.fsName);

                /* 元ファイルが差し替えられているなどでサイズが合わない場合は知らせる
                   Report a size mismatch, e.g. when the source file has been replaced */
                if (result.sizeChanged) {
                    warningMessages.push("［" + itemLabel + "］ 元の埋め込み画像とサイズが一致しません。");
                }
            } catch (e) {
                errorMessages.push("［" + itemLabel + "］ " + e.message);
            }
        }

        /* 新しいリンク画像を選択状態にする / Select the newly placed items */
        try {
            doc.selection = null;
            for (var k = 0; k < newItems.length; k++) {
                newItems[k].selected = true;
            }
        } catch (e) {
            /* 選択の復元に失敗しても処理結果には影響しない / Selection restore is best-effort */
        }

        /* 結果をまとめて表示 / Report the result */
        var message = "リンク画像への切り替えと再リンクが完了しました。\n";
        message += "対象: " + rasterItems.length + " 件 ／ 再リンク: " + relinkedFiles.length + " 件";
        if (skippedCount > 0) {
            message += " ／ スキップ: " + skippedCount + " 件";
        }
        if (errorMessages.length > 0) {
            message += " ／ エラー: " + errorMessages.length + " 件";
        }

        if (relinkedFiles.length > 0) {
            message += "\n\n【再リンク先】\n" + relinkedFiles.join("\n");
        }
        if (warningMessages.length > 0) {
            message += "\n\n【注意】\n" + warningMessages.join("\n");
        }
        if (errorMessages.length > 0) {
            message += "\n\n【エラー】\n" + errorMessages.join("\n");
        }

        alert(message);
    }

    main();
})();
