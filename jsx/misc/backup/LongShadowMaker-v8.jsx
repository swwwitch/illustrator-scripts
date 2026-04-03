#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
  Long Shadow Maker (Extrude and Merge with Preview)
  ダイアログで距離・角度・スケールを指定してロングシャドウを生成するスクリプト。プレビュー機能付き。
  サンプリング方式で影の側面を生成し、必要に応じてパスファインダーで合体・穴埋めを行います。

  対象：閉パスの PathItem / CompoundPathItem / GroupItem / TextFrame

  オリジナルアイデアとコード
  こじらせたクマー さん
  https://note.com/nice_lotus120/n/nf406fb3ae2b4


 更新日: 2026-02-09
*/

var SCRIPT_VERSION = "v1.0.9";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: {
        ja: "ロングシャドウメーカー",
        en: "Long Shadow Maker"
    },
    openDocument: {
        ja: "ドキュメントを開いてください。",
        en: "Please open a document."
    },
selectOneClosed: {
    ja: "閉パス（単一パス／複合パス／グループ）を1つだけ選択してください。",
    en: "Please select exactly one closed shape (path / compound path / group / text)."
},
    selectClosed: {
        ja: "閉パスを選択してください。",
        en: "Please select a closed path."
    },
    selectClosedGroup: {
        ja: "閉パスのグループを選択してください。",
        en: "Please select a group that consists of closed paths."
    },
    cannotBuildFromGroup: {
        ja: "グループから形状を作成できませんでした。閉パスのグループを選択してください。",
        en: "Could not build a shape from the group. Please select a group of closed paths."
    },
    notGroupItem: {
    ja: "GroupItem ではありません。",
    en: "This is not a GroupItem."
},
cannotGetMergedFromGroup: {
    ja: "グループの合体結果を取得できませんでした。",
    en: "Could not retrieve the merged result from the group."
},
groupMergeError: {
    ja: "グループの合体中にエラーが発生しました: ",
    en: "An error occurred while merging the group: "
},
notTextFrame: {
    ja: "TextFrame ではありません。",
    en: "This is not a TextFrame."
},
outlineFailed: {
    ja: "アウトライン化に失敗しました。",
    en: "Failed to create outlines."
},
textMergeError: {
    ja: "テキストの合体中にエラーが発生しました: ",
    en: "An error occurred while processing the text: "
},
    preset: {
        ja: "プリセット",
        en: "Preset"
    },
    settings: {
        ja: "設定",
        en: "Settings"
    },
    offset: {
        ja: "オフセット",
        en: "Offset"
    },
shape: {
    ja: "形状",
    en: "Join"
},
    distance: {
        ja: "距離",
        en: "Distance"
    },
    angle: {
        ja: "角度",
        en: "Angle"
    },
    scale: {
        ja: "スケール",
        en: "Scale"
    },
    preview: {
        ja: "プレビュー",
        en: "Preview"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    ok: {
        ja: "OK",
        en: "OK"
    },
    joinMiter: {
        ja: "マイター",
        en: "Miter"
    },
    joinRound: {
        ja: "ラウンド",
        en: "Round"
    },
    joinBevel: {
        ja: "ベベル",
        en: "Bevel"
    }
};

function L(key) {
    try {
        if (LABELS[key] && LABELS[key][lang]) return LABELS[key][lang];
        if (LABELS[key] && LABELS[key].en) return LABELS[key].en;
    } catch (_) { }
    return key;
}

(function () {
    if (app.documents.length === 0) {
        alert(L('openDocument'));
        return;
    }

    var doc = app.activeDocument;
    var sel = doc.selection;

    function isSupportedItem(it) {
        return it && (
            it.typename === "PathItem" ||
            it.typename === "CompoundPathItem" ||
            it.typename === "GroupItem" ||
            it.typename === "TextFrame"
        );
    }

    function getSubPaths(it) {
        if (!it) return [];

        if (it.typename === "PathItem") return [it];

        if (it.typename === "CompoundPathItem") {
            var arr = [];
            try {
                for (var i = 0; i < it.pathItems.length; i++) arr.push(it.pathItems[i]);
            } catch (_) { }
            return arr;
        }

        // GroupItem: 中にある PathItem / CompoundPathItem を再帰的に集める
        if (it.typename === "GroupItem") {
            var out = [];

            function collectFrom(container) {
                if (!container) return;
                var items = null;
                try { items = container.pageItems; } catch (_) { items = null; }
                if (!items) return;

                for (var j = 0; j < items.length; j++) {
                    var child = items[j];
                    if (!child) continue;

                    if (child.typename === "PathItem") {
                        out.push(child);
                    } else if (child.typename === "CompoundPathItem") {
                        try {
                            for (var k = 0; k < child.pathItems.length; k++) out.push(child.pathItems[k]);
                        } catch (_) { }
                    } else if (child.typename === "GroupItem") {
                        collectFrom(child);
                    }
                }
            }

            collectFrom(it);
            return out;
        }

        return [];
    }

    function isAllClosed(subPaths) {
        if (!subPaths || subPaths.length === 0) return false;
        for (var i = 0; i < subPaths.length; i++) {
            try {
                if (!subPaths[i].closed) return false;
            } catch (_) {
                return false;
            }
        }
        return true;
    }

    if (sel.length !== 1 || !isSupportedItem(sel[0])) {
        alert(L('selectOneClosed'));
        return;
    }

    var originalPath = sel[0];
    var originalSubPaths = getSubPaths(originalPath);

    // --- 一時ベース（②）残骸回収用タグ ---
    var TEMP_BASE_NAME = "__LongShadowTempBase__";

    function tagTempBaseDeep(it) {
        if (!it) return;
        try { it.name = TEMP_BASE_NAME; } catch (_) { }
        try {
            if (it.typename === 'GroupItem') {
                for (var i = 0; i < it.pageItems.length; i++) tagTempBaseDeep(it.pageItems[i]);
            } else if (it.typename === 'CompoundPathItem') {
                for (var j = 0; j < it.pathItems.length; j++) {
                    try { it.pathItems[j].name = TEMP_BASE_NAME; } catch (_) { }
                }
            }
        } catch (_) { }
    }

    function removeAllTempBasesByName() {
        // --- 一時レイヤー（__LongShadowTempBase__）の削除 ---
        function removeTempLayerByName() {
            try {
                var layers = doc.layers;
                for (var i = layers.length - 1; i >= 0; i--) {
                    try {
                        var lyr = layers[i];
                        if (!lyr) continue;
                        var nm = '';
                        try { nm = lyr.name; } catch (_) { nm = ''; }
                        if (nm === TEMP_BASE_NAME) {
                            // unlock/visible before remove
                            try { lyr.locked = false; } catch (_) { }
                            try { lyr.visible = true; } catch (_) { }
                            try { lyr.remove(); } catch (_) { }
                        }
                    } catch (_) { }
                }
            } catch (_) { }
        }
        function unlockChain(it) {
            try {
                // unlock self
                try { it.locked = false; } catch (_) { }
                try { it.hidden = false; } catch (_) { }

                // unlock parent chain (GroupItem/Layer)
                var p = null;
                try { p = it.parent; } catch (_) { p = null; }
                var guard = 0;
                while (p && guard++ < 50) {
                    try {
                        if (p.typename === 'Layer') {
                            try { p.locked = false; } catch (_) { }
                            try { p.visible = true; } catch (_) { }
                            break;
                        }
                        if (p.typename === 'GroupItem') {
                            try { p.locked = false; } catch (_) { }
                            try { p.hidden = false; } catch (_) { }
                        }
                        try { p = p.parent; } catch (_) { p = null; }
                    } catch (_) {
                        break;
                    }
                }
            } catch (_) { }
        }

        function forceRemove(it) {
            try { if (!it || !it.isValid) return; } catch (_) { return; }
            // first: unlock chain
            try { unlockChain(it); } catch (_) { }
            // try direct remove
            try { it.remove(); return; } catch (_) { }
            // fallback: select + clear
            try {
                doc.selection = null;
                it.selected = true;
                app.executeMenuCommand('clear');
            } catch (_) { }
            try { doc.selection = null; } catch (_) { }
            // last try
            try { if (it && it.isValid) it.remove(); } catch (_) { }
        }

        try {
            // Prefer groupItems scan (temp base is usually a group)
            var gi = doc.groupItems;
            for (var i = gi.length - 1; i >= 0; i--) {
                try {
                    var g = gi[i];
                    if (!g || !g.isValid) continue;
                    var nm = '';
                    try { nm = g.name; } catch (_) { nm = ''; }
                    if (nm === TEMP_BASE_NAME) {
                        forceRemove(g);
                    }
                } catch (_) { }
            }
        } catch (_) { }

        // Also sweep remaining pageItems (covers non-group leftovers)
        try {
            var items = doc.pageItems;
            for (var j = items.length - 1; j >= 0; j--) {
                try {
                    var it = items[j];
                    if (!it || !it.isValid) continue;
                    var nm2 = '';
                    try { nm2 = it.name; } catch (_) { nm2 = ''; }
                    if (nm2 === TEMP_BASE_NAME) {
                        forceRemove(it);
                    }
                } catch (_) { }
            }
        } catch (_) { }
    }

    // Path/Compound はここで閉パス検証。Group/Text は実行時に一時パスへ変換して検証する。
    if (originalPath.typename !== "GroupItem" && originalPath.typename !== "TextFrame") {
        if (!isAllClosed(originalSubPaths)) {
            alert(L('selectClosed'));
            return;
        }
    }
    /* グループの一時合体 / Build temporary merged item from GroupItem */
    function executePathfinderAddAndExpand() {
        app.executeMenuCommand('Live Pathfinder Add');
        app.executeMenuCommand('expandStyle');
    }

    // GroupItem を「一時的に」単一パス/複合パスへ変換して返す（元グループは残す）
    // 戻り値: { item: PageItem, cleanup: Function, ok: Boolean, message: String }
    function buildMergedItemFromGroup(groupItem) {
        var res = { item: null, cleanup: function () { }, ok: false, message: "" };

        if (!groupItem || groupItem.typename !== "GroupItem") {
            res.message = L('notGroupItem');
            return res;
        }

        var tempDup = null;
        var merged = null;
        var trash = []; // 一時生成物を全部ここに入れて最後に削除する

        res.cleanup = function () {
            try { doc.selection = null; } catch (_) { }
            try {
                for (var ti = trash.length - 1; ti >= 0; ti--) {
                    try {
                        var it = trash[ti];
                        if (it && it.isValid) it.remove();
                    } catch (_) { }
                }
            } catch (_) { }

            // 念のため
            try { if (merged && merged.isValid) merged.remove(); } catch (_) { }
            try { if (tempDup && tempDup.isValid) tempDup.remove(); } catch (_) { }
        };

        try {
            tempDup = groupItem.duplicate();
            try { trash.push(tempDup); } catch (_) { }

            // duplicate を選択して PathFinder → Expand
            doc.selection = null;
            tempDup.selected = true;
            executePathfinderAddAndExpand();

            var items = doc.selection;
            doc.selection = null;
            try {
                if (items && items.length) {
                    for (var t0 = 0; t0 < items.length; t0++) trash.push(items[t0]);
                }
            } catch (_) { }

            if (!items || items.length === 0) {
                res.message = L('cannotGetMergedFromGroup');
                return res;
            }

            if (items.length === 1) {
                merged = items[0];
            } else {
                // 複数残った場合は一度グループ化して再合体
                var g = doc.groupItems.add();
                try { trash.push(g); } catch (_) { }
                for (var i = 0; i < items.length; i++) {
                    try { items[i].move(g, ElementPlacement.PLACEATEND); } catch (_) { }
                }
                doc.selection = null;
                g.selected = true;
                executePathfinderAddAndExpand();

                var items2 = doc.selection;
                doc.selection = null;
                try {
                    if (items2 && items2.length) {
                        for (var t1 = 0; t1 < items2.length; t1++) trash.push(items2[t1]);
                    }
                } catch (_) { }

                if (items2 && items2.length === 1) {
                    merged = items2[0];
                } else {
                    merged = (items2 && items2.length > 0) ? items2[0] : items[0];
                }
            }

            try { if (merged) trash.push(merged); } catch (_) { }
            // 一時ベース（②）にタグ付け（子要素まで）
            try { tagTempBaseDeep(merged); } catch (_) { }
            res.item = merged;
            res.ok = true;
            return res;

        } catch (e) {
            res.message = L('groupMergeError') + e;
            return res;
        }
    }
    // TextFrame を「一時的に」アウトライン化→合体して単一パス/複合パスへ変換して返す（元テキストは残す）
    // 戻り値: { item: PageItem, cleanup: Function, ok: Boolean, message: String }
    function buildMergedItemFromText(textFrame) {
        var res = { item: null, cleanup: function () { }, ok: false, message: "" };

        if (!textFrame || textFrame.typename !== "TextFrame") {
            res.message = L('notTextFrame');
            return res;
        }

        var tempDup = null;
        var merged = null;
        var trash = [];

        res.cleanup = function () {
            try { doc.selection = null; } catch (_) { }
            try {
                for (var ti = trash.length - 1; ti >= 0; ti--) {
                    try {
                        var it = trash[ti];
                        if (it && it.isValid) it.remove();
                    } catch (_) { }
                }
            } catch (_) { }
            try { if (merged && merged.isValid) merged.remove(); } catch (_) { }
            try { if (tempDup && tempDup.isValid) tempDup.remove(); } catch (_) { }
        };

        try {
            // ① テキストを複製
            tempDup = textFrame.duplicate();
            try { trash.push(tempDup); } catch (_) { }

            // ② アウトライン化（v15方式：TextFrame#createOutline を使用）
            // createOutline() は元の TextFrame を削除して GroupItem/CompoundPathItem を返す
            // ここでは複製に対して実行するので、オリジナルは残る
            var outlined = null;
            try {
                if (tempDup && typeof tempDup.createOutline === 'function') {
                    outlined = tempDup.createOutline();
                }
            } catch (eOutline) {
                outlined = null;
            }

            // createOutline() 成功時は tempDup が消えるので trash から外す必要はない（cleanup側でisValidを見て消す）

            if (!outlined) {
                res.message = L('outlineFailed');
                return res;
            }

            // cleanup 対象は「アウトライン化で生成されたルート」だけにする。
            // 子要素まで個別に trash に積むと、後段処理で移動・合体された要素を誤って削除して
            // 生成済みのロングシャドウが消えることがあるため。
            try { trash.push(outlined); } catch (_) { }

            // 選択状態が残ると後段の処理に巻き込まれるので解除
            try { doc.selection = null; } catch (_) { }
            try { outlined.selected = false; } catch (_) { }

            merged = outlined;
            // 一時ベース（②）にタグ付け（子要素まで）
            try { tagTempBaseDeep(merged); } catch (_) { }
            res.item = merged;
            res.ok = true;
            return res;

        } catch (e) {
            res.message = L('textMergeError') + e;
            return res;
        }
    }

    var isPreviewing = false;

    var __applyOffsetNow = true; // プレビュー時は false にしてオフセットを無視

    /* ダイアログ作成 / Build dialog */
    var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
    dlg.orientation = "column";
    dlg.alignChildren = "fill";

    // プリセット（スケール＋角度） ※距離は含めない（ダイアログ最上部）
    // 外側グループで左右中央に配置し、内側グループにラベル＋プルダウンをまとめる
    var grpPresetWrap = dlg.add("group");
    grpPresetWrap.orientation = "row";
    grpPresetWrap.alignChildren = ["center", "center"];
    grpPresetWrap.alignment = "center";

    var grpPreset = grpPresetWrap.add("group");
    grpPreset.orientation = "row";
    grpPreset.alignChildren = ["left", "center"];

    grpPreset.add("statictext", undefined, L('preset'));

    var ddPreset = grpPreset.add("dropdownlist", undefined, [
        "100% /  45°",
        "100% /  30°",
        "100% /  60°",
        " 50% /  90°",
        "  1% /  90°",
        "100% / 135°",
        "100% / 120°",
        "100% / 150°"
    ]);
    ddPreset.selection = 0; // デフォルト：45° / 100%

    // --- 1カラム（プリセットは貫通） ---
    var cols = dlg.add("group");
    cols.orientation = "column";
    cols.alignChildren = ["fill", "top"];
    cols.alignment = "fill";

    // 左：現状（設定）
    var mainGroup = cols.add("panel", undefined, L('settings'));
    mainGroup.orientation = "column";
    mainGroup.alignChildren = "left";
    mainGroup.margins = [15, 20, 15, 10];

    // 右：オフセット（中身は後で実装）
    var offsetPanel = cols.add("panel", undefined, L('offset'));
    offsetPanel.orientation = "column";
    offsetPanel.alignChildren = "left";
    offsetPanel.margins = [15, 20, 15, 10];

    // --- オフセットpanel内 2カラム ---
    var offCols = offsetPanel.add("group");
    offCols.orientation = "row";
    offCols.alignChildren = ["fill", "top"];
    // offCols.alignment = "fill";

    // 左：□［　］pt
    var offLeft = offCols.add("group");
    offLeft.orientation = "column";
    offLeft.alignChildren = ["left", "top"];

    // オフセット（チェック＋値を1行に）
    var gOff = offLeft.add("group");
    gOff.orientation = "row";
    gOff.alignChildren = ["left", "center"];

    var cbOffsetRow = gOff.add("checkbox", undefined, "");
    cbOffsetRow.value = false;

    // オフセット初期値を選択オブジェクト（originalPath）のサイズから計算（pt）
    var initialOffset = 0;
    try {
        var bOff = null;
        try { bOff = originalPath.geometricBounds; } catch (_) { bOff = null; }
        if (!bOff) {
            try { bOff = originalPath.visibleBounds; } catch (_) { bOff = null; }
        }
        if (bOff && bOff.length === 4) {
            var wPtOff = Math.abs(bOff[2] - bOff[0]);
            var hPtOff = Math.abs(bOff[1] - bOff[3]);

            // 幅と高さの合計（pt）
            var sizeSumPt = wPtOff + hPtOff;
            // 平均サイズ（pt）
            var halfSizePt = sizeSumPt / 2;

            // PathItem 想定：divisor は 20
            var divisor = 20;

            // オフセット基準値（pt）
            var offsetBasePt = halfSizePt / divisor;

            if (!isNaN(offsetBasePt)) {
                initialOffset = Math.round(offsetBasePt);
            }
        }
    } catch (_) { }

    var etOff = gOff.add("edittext", undefined, String(initialOffset));
    etOff.characters = 4;
    changeValueByArrowKey(etOff, false, false);
    var stOffUnit = gOff.add("statictext", undefined, "pt");

    // 右：形状panel
    var offRight = offCols.add("group");
    offRight.orientation = "column";
    offRight.alignChildren = ["fill", "top"];

    // 形状（角の処理） - オフセットpanel内（2カラム：左=見出し / 右=ラジオ）
    var pJoin = offRight.add("group");
    pJoin.orientation = "row";
    pJoin.alignChildren = ["left", "top"];
    pJoin.margins = [0, 0, 0, 0];

    // 左：見出し
    var gJoinLabel = pJoin.add("group");
    gJoinLabel.orientation = "column";
    gJoinLabel.alignChildren = ["left", "top"];
    var stJoin = gJoinLabel.add("statictext", undefined, L('shape'));
    stJoin.preferredSize.width = 60;
    stJoin.justify = "right";

    // 右：ラジオ（垂直並び）
    var gJoinCol = pJoin.add("group");
    gJoinCol.orientation = "column";
    gJoinCol.alignChildren = ["left", "center"];

    var rbMiter = gJoinCol.add("radiobutton", undefined, L('joinMiter'));
    var rbRound = gJoinCol.add("radiobutton", undefined, L('joinRound'));
    var rbBevel = gJoinCol.add("radiobutton", undefined, L('joinBevel'));
    rbMiter.value = false;
    rbRound.value = true;  // default = Round
    rbBevel.value = false;

    // オフセットUIの有効/無効（ディム表示）
    function updateOffsetEnabled() {
        var on = !!cbOffsetRow.value;
        try { etOff.enabled = on; } catch (_) { }
        try { rbMiter.enabled = on; } catch (_) { }
        try { rbRound.enabled = on; } catch (_) { }
        try { rbBevel.enabled = on; } catch (_) { }
        try { stOffUnit.enabled = on; } catch (_) { }
        try { pJoin.enabled = on; } catch (_) { }
    }

    cbOffsetRow.onClick = function () {
        updateOffsetEnabled();
    };

    // 初期状態：OFF
    updateOffsetEnabled();

    // 距離入力
    var grpDistance = mainGroup.add("group");
    var lblDistance = grpDistance.add("statictext", undefined, L('distance'));
    lblDistance.preferredSize.width = 60;
    lblDistance.justify = "right";

    // オリジナルの幅＋高さを足した値をデフォルト距離にする（pt）
    var gb = originalPath.geometricBounds; // [left, top, right, bottom]
    var w = gb[2] - gb[0];
    var h = gb[1] - gb[3];
    var defaultDistance = (w + h);

    var inputDistance = grpDistance.add("edittext", undefined, Math.round(defaultDistance).toString());
    changeValueByArrowKey(inputDistance, false, true);
    inputDistance.characters = 4;
    var stDistUnit = grpDistance.add("statictext", undefined, "pt");
    stDistUnit.preferredSize.width = 24;

    // スライダー（距離）※右側に配置
    var maxDist = Math.max(500, Math.round(defaultDistance * 3));
    var slDistance = grpDistance.add("slider", undefined, Math.round(defaultDistance), 0, maxDist);
    slDistance.preferredSize.width = 170;

    // 角度入力
    var grpAngle = mainGroup.add("group");
    var lblAngle = grpAngle.add("statictext", undefined, L('angle'));
    lblAngle.preferredSize.width = 60;
    lblAngle.justify = "right";
    var inputAngle = grpAngle.add("edittext", undefined, "45");
    changeValueByArrowKey(inputAngle, true, true);
    inputAngle.characters = 4;
    var stAngleUnit = grpAngle.add("statictext", undefined, "°");
    stAngleUnit.preferredSize.width = 24;

    // スライダー（角度）※右側に配置
    var slAngle = grpAngle.add("slider", undefined, 45, -180, 180);
    slAngle.preferredSize.width = 170;

    // スケール入力
    var grpScale = mainGroup.add("group");
    var lblScale = grpScale.add("statictext", undefined, L('scale'));
    lblScale.preferredSize.width = 60;
    lblScale.justify = "right";
    var inputScale = grpScale.add("edittext", undefined, "100"); // デフォルト 100%
    changeValueByArrowKey(inputScale, false, true);
    inputScale.characters = 4;
    var stScaleUnit = grpScale.add("statictext", undefined, "%");
    stScaleUnit.preferredSize.width = 24;


    // スライダー（スケール）※右側に配置
    var slScale = grpScale.add("slider", undefined, 100, 1, 300);
    slScale.preferredSize.width = 170;

    // スライダー連動（edittext <-> slider）
    bindSliderNumber(inputDistance, slDistance, 0, maxDist, false);
    bindSliderNumber(inputAngle, slAngle, -180, 180, true);
    bindSliderNumber(inputScale, slScale, 1, 300, false);

    // プリセット選択時：スケールと角度に反映（距離は変更しない）
    ddPreset.onChange = function () {
        if (!ddPreset.selection) return;

        // 例: "100% / 45°"
        var s = ddPreset.selection.text;
        var m = s.match(/^\s*([\-]?\d+(?:\.\d+)?)\s*%\s*\/\s*([\-]?\d+(?:\.\d+)?)\s*°\s*$/);
        if (!m) return;

        var scl = m[1];
        var ang = m[2];

        inputAngle.text = ang;
        inputScale.text = scl;

        if (chkPreview && chkPreview.value) updatePreview();
    };

    // Buttons（左：プレビュー（無効） / 右：キャンセル・OK）
    var gBtn = dlg.add("group");
    gBtn.orientation = "row";
    gBtn.alignment = ["fill", "top"];
    gBtn.alignChildren = ["right", "center"];

    // 左端：プレビュー
    var chkPreview = gBtn.add("checkbox", undefined, L('preview'));
    chkPreview.value = true;   // デフォルトON
    chkPreview.enabled = true; // 有効

    // スペーサー（左右を分離）
    var spacer = gBtn.add("group");
    spacer.alignment = ["fill", "fill"];

    // 右端：キャンセル / OK
    var btnCancel = gBtn.add("button", undefined, L('cancel'), { name: "cancel" });
    var btnOk = gBtn.add("button", undefined, L('ok'), { name: "ok" });
    btnOk.active = true;

    // --- ↑↓キーで数値を増減（Shift=±10, Option=±0.1） ---
    function changeValueByArrowKey(editText, allowNegative, enablePreviewUpdate) {
        editText.addEventListener("keydown", function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                // Shiftキー押下時は10の倍数にスナップ
                if (event.keyName == "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value = Math.floor((value - 1) / delta) * delta;
                    event.preventDefault();
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                // Optionキー押下時は0.1単位で増減
                if (event.keyName == "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= delta;
                    event.preventDefault();
                }
            } else {
                delta = 1;
                if (event.keyName == "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= delta;
                    event.preventDefault();
                }
            }

            if (keyboard.altKey) {
                // 小数第1位までに丸め
                value = Math.round(value * 10) / 10;
            } else {
                // 整数に丸め
                value = Math.round(value);
            }

            if (!allowNegative && value < 0) value = 0;

            // 反映
            editText.text = value;

            // プレビュー更新（必要なフィールドのみ／プレビューONの時のみ）
            if (enablePreviewUpdate !== false && typeof updatePreview === "function" && chkPreview && chkPreview.value) {
                updatePreview();
            }
        });
    }

    // --- スライダー連動（距離/スケール/角度） ---
    function clamp(v, min, max) {
        return Math.max(min, Math.min(max, v));
    }

    // slider は整数のみなので、必要に応じて丸める
    function bindSliderNumber(editText, slider, min, max, isAngle) {
        var syncing = false;

        function setEditFromSlider() {
            if (syncing) return;
            syncing = true;
            try {
                var v = slider.value;
                if (!isAngle) {
                    // 距離/スケールは整数
                    v = Math.round(v);
                } else {
                    // 角度も整数扱い
                    v = Math.round(v);
                }
                editText.text = String(v);
            } catch (_) { }
            syncing = false;

            // プレビュー更新（ONの時のみ）
            if (typeof updatePreview === "function" && chkPreview && chkPreview.value) {
                updatePreview();
            }
        }

        function setSliderFromEdit() {
            if (syncing) return;
            syncing = true;
            try {
                var v = Number(editText.text);
                if (isNaN(v)) v = 0;
                v = clamp(v, min, max);
                slider.value = Math.round(v);
            } catch (_) { }
            syncing = false;
        }

        // 初期同期
        setSliderFromEdit();

        slider.onChanging = setEditFromSlider;
        slider.onChange = setEditFromSlider; // マウス離したときも反映

        // editText側の変化に追随（既存のonChangingは残しつつ、同期だけ追加）
        var prevOnChanging = editText.onChanging;
        editText.onChanging = function () {
            try { setSliderFromEdit(); } catch (_) { }
            if (typeof prevOnChanging === "function") prevOnChanging();
        };
    }

    /* 色処理 / Color utilities */
    // A（オリジナル）の塗り色を元に、彩度を下げた色を作る
    // factor: 0〜1（1=元の彩度、0=完全に無彩色）
    function desaturateColorFromFill(fillColor, factor) {
        if (!fillColor) return null;
        if (factor === undefined || factor === null) factor = 0.7;

        // Grayはそのまま
        if (fillColor.typename === "GrayColor") {
            var g = new GrayColor();
            g.gray = fillColor.gray;
            return g;
        }

        // CMYKは一旦RGBに近似変換して処理（戻しはRGBで適用）
        var rgb = null;
        if (fillColor.typename === "RGBColor") {
            rgb = { r: fillColor.red, g: fillColor.green, b: fillColor.blue };
        } else if (fillColor.typename === "CMYKColor") {
            // 0-100 -> 0-1
            var c = fillColor.cyan / 100.0;
            var m = fillColor.magenta / 100.0;
            var y = fillColor.yellow / 100.0;
            var k = fillColor.black / 100.0;
            // 簡易 CMYK -> RGB
            var rr = 255 * (1 - c) * (1 - k);
            var gg = 255 * (1 - m) * (1 - k);
            var bb = 255 * (1 - y) * (1 - k);
            rgb = { r: rr, g: gg, b: bb };
        } else {
            // 対応外はそのまま返す（NoColor等）
            return fillColor;
        }

        // RGB -> HSL で彩度調整（HSLのSをfactor倍）
        function clamp01(v) { return Math.max(0, Math.min(1, v)); }
        var r1 = clamp01(rgb.r / 255.0);
        var g1 = clamp01(rgb.g / 255.0);
        var b1 = clamp01(rgb.b / 255.0);

        var max = Math.max(r1, g1, b1);
        var min = Math.min(r1, g1, b1);
        var h = 0, s = 0, l = (max + min) / 2;
        var d = max - min;

        if (d !== 0) {
            s = d / (1 - Math.abs(2 * l - 1));
            switch (max) {
                case r1: h = ((g1 - b1) / d) % 6; break;
                case g1: h = ((b1 - r1) / d) + 2; break;
                case b1: h = ((r1 - g1) / d) + 4; break;
            }
            h = h * 60;
            if (h < 0) h += 360;
        }

        s = clamp01(s * factor);

        // HSL -> RGB
        function hue2rgb(p, q, t) {
            if (t < 0) t += 1;
            if (t > 1) t -= 1;
            if (t < 1 / 6) return p + (q - p) * 6 * t;
            if (t < 1 / 2) return q;
            if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
            return p;
        }

        var r2, g2, b2;
        if (s === 0) {
            r2 = g2 = b2 = l;
        } else {
            var q = l < 0.5 ? l * (1 + s) : (l + s - l * s);
            var p = 2 * l - q;
            var hk = h / 360;
            r2 = hue2rgb(p, q, hk + 1 / 3);
            g2 = hue2rgb(p, q, hk);
            b2 = hue2rgb(p, q, hk - 1 / 3);
        }

        var out = new RGBColor();
        out.red = Math.round(r2 * 255);
        out.green = Math.round(g2 * 255);
        out.blue = Math.round(b2 * 255);
        return out;
    }


    function bezierPoint(p0, p1, p2, p3, t) {
        var u = 1 - t;
        var tt = t * t;
        var uu = u * u;
        var uuu = uu * u;
        var ttt = tt * t;
        return [
            uuu * p0[0] + 3 * uu * t * p1[0] + 3 * u * tt * p2[0] + ttt * p3[0],
            uuu * p0[1] + 3 * uu * t * p1[1] + 3 * u * tt * p2[1] + ttt * p3[1]
        ];
    }


    /* 実パス生成（サンプリング面生成） / Build shadow by sampling & bridging */
    var CURVE_SAMPLE_COUNT = 8; // v15 相当（必要なら調整）

    function duplicateMoveScale(item, dx, dy, scalePercent) {
        if (!item) return null;
        var dup = null;
        try { dup = item.duplicate(); } catch (_) { dup = null; }
        if (!dup) return null;

        // scale: percent (100 = 等倍)
        try {
            var sp = (isFinite(scalePercent) ? Number(scalePercent) : 100);
            if (sp !== 100 && dup && typeof dup.resize === 'function') {
                try { dup.resize(sp, sp, true, true, true, true, true, Transformation.CENTER); }
                catch (eR1) {
                    try { dup.resize(sp, sp, true, true, true, true, true); } catch (_) { }
                }
            }
        } catch (_) { }

        try { dup.translate(dx, dy); } catch (_) { }
        return dup;
    }

    // PathItem を「曲線も含めて」点列にする（閉パス前提）
    function samplePathToPoints(pathItem, curveSamplesPerSeg) {
        if (!curveSamplesPerSeg) curveSamplesPerSeg = CURVE_SAMPLE_COUNT;
        var pts = [];
        if (!pathItem) return pts;

        try {
            var pp = pathItem.pathPoints;
            var n = pp.length;
            if (!pp || n < 2) return pts;

            for (var i = 0; i < n; i++) {
                var a = pp[i];
                var b = pp[(i + 1) % n];

                var p0 = a.anchor;
                var p1 = a.rightDirection;
                var p2 = b.leftDirection;
                var p3 = b.anchor;

                for (var s = 0; s < curveSamplesPerSeg; s++) {
                    var t = s / curveSamplesPerSeg;
                    pts.push(bezierPoint(p0, p1, p2, p3, t));
                }
            }
        } catch (_) { }

        return pts;
    }

    // PathItem/Compound/Group を「対応順序を保って」PathItem 配列として抽出（閉パスのみ）
    function collectClosedPathsForBridge(item) {
        var out = [];
        if (!item) return out;

        function pushIfClosed(p) {
            try { if (p && p.typename === 'PathItem' && p.closed) out.push(p); } catch (_) { }
        }

        try {
            if (item.typename === 'PathItem') {
                pushIfClosed(item);
                return out;
            }
            if (item.typename === 'CompoundPathItem') {
                for (var i = 0; i < item.pathItems.length; i++) pushIfClosed(item.pathItems[i]);
                return out;
            }
            if (item.typename === 'GroupItem') {
                // pageItems 順で再帰（見た目の重なり順に近い）
                for (var j = 0; j < item.pageItems.length; j++) {
                    var ch = item.pageItems[j];
                    if (!ch) continue;
                    if (ch.typename === 'PathItem') {
                        pushIfClosed(ch);
                    } else if (ch.typename === 'CompoundPathItem') {
                        for (var k = 0; k < ch.pathItems.length; k++) pushIfClosed(ch.pathItems[k]);
                    } else if (ch.typename === 'GroupItem') {
                        var sub = collectClosedPathsForBridge(ch);
                        for (var m = 0; m < sub.length; m++) out.push(sub[m]);
                    }
                }
                return out;
            }
        } catch (_) { }

        return out;
    }

    // 点列同士を結ぶ側面（四角形）を作る
    function buildSideFacesFromPointLists(parentGroup, ptsA, ptsB, baseFill) {
        var created = [];
        if (!parentGroup || !ptsA || !ptsB) return created;
        var n = Math.min(ptsA.length, ptsB.length);
        if (n < 2) return created;

        for (var i = 0; i < n; i++) {
            var next = (i + 1) % n;
            var p = null;
            try {
                p = parentGroup.pathItems.add();
                p.setEntirePath([ptsA[i], ptsB[i], ptsB[next], ptsA[next]]);
                p.closed = true;
                p.stroked = false;
                p.filled = true;
                if (baseFill) p.fillColor = baseFill;
                created.push(p);
            } catch (_) {
                try { if (p && p.isValid) p.remove(); } catch (_) { }
            }
        }
        return created;
    }

    // v15系：サンプリングで面を作る実生成
    // 戻り値: 生成した影グループ（GroupItem）
    function generateShadowBySampling(baseItem, dx, dy, scalePercent) {
        if (!baseItem) return null;

        // ベースの閉パス群を抽出
        var basePaths = collectClosedPathsForBridge(baseItem);
        if (!basePaths || basePaths.length === 0) return null;

        // 複製して移動/スケール（複製側）
        var dup = duplicateMoveScale(baseItem, dx, dy, scalePercent);
        if (!dup) return null;

        // 複製側の閉パス群を抽出
        var dupPaths = collectClosedPathsForBridge(dup);
        if (!dupPaths || dupPaths.length === 0) {
            try { dup.remove(); } catch (_) { }
            return null;
        }

        // ペアリング：数が違う場合は最小数で処理（落とさない優先）
        var pairCount = Math.min(basePaths.length, dupPaths.length);
        if (pairCount === 0) {
            try { dup.remove(); } catch (_) { }
            return null;
        }

        // 影用グループ
        var g = doc.groupItems.add();

        // ベースの塗りを基準に彩度を落とす
        var baseFill = null;
        try { baseFill = desaturateColorFromFill(basePaths[0].fillColor, 0.7); } catch (_) { baseFill = null; }

        for (var i = 0; i < pairCount; i++) {
            var a = basePaths[i];
            var b = dupPaths[i];

            // 曲線でも直線でも一定サンプルで点列化
            var ptsA = samplePathToPoints(a, CURVE_SAMPLE_COUNT);
            var ptsB = samplePathToPoints(b, CURVE_SAMPLE_COUNT);

            // 点数がずれた場合に備えて、短い方に合わせる
            if (ptsA.length < 2 || ptsB.length < 2) continue;

            buildSideFacesFromPointLists(g, ptsA, ptsB, baseFill);
        }

        // 複製側（dup）は最終結果には不要
        try { dup.remove(); } catch (_) { }

        return g;
    }

    /* オフセット（ライブ効果） / Offset Path live effect */
    // joinTypes: 0 = Round, 1 = Bevel, 2 = Miter
    function getJoinCode() {
        try {
            if (rbRound && rbRound.value) return 0;
            if (rbBevel && rbBevel.value) return 1;
        } catch (_) { }
        return 2; // マイター
    }

    function buildOffsetEffectXML(offsetPt, joinCode) {
        // mlim はマイター制限（既定 4）、ofst は pt
        return '<LiveEffect name="Adobe Offset Path"><Dict data="R mlim 4 R ofst ' + offsetPt + ' I jntp ' + joinCode + ' "/></LiveEffect>';
    }

    function groupItems(doc, items) {
        var g = doc.groupItems.add();
        // 受け取った順序で入れる
        for (var i = 0; i < items.length; i++) {
            try {
                items[i].move(g, ElementPlacement.PLACEATEND);
            } catch (_) { }
        }
        return g;
    }

    function applyOffsetEffectToCGroup(cGroup) {
        if (!cGroup) return;
        if (!__applyOffsetNow) return;
        if (!cbOffsetRow || !cbOffsetRow.value) return;

        var off = Number(etOff.text);
        if (isNaN(off)) off = 0;

        // 0でもONなら適用する（結果が変わらないだけ）
        var joinCode = getJoinCode();
        var xml = buildOffsetEffectXML(off, joinCode);

        try {
            cGroup.applyEffect(xml);
        } catch (_) { }

        // 角丸（Round）の場合、後処理として Pathfinder Merge を実行
        if (joinCode === 0) {
            try {
                doc.selection = null;
                cGroup.selected = true;
                app.executeMenuCommand('Live Pathfinder Merge');
                doc.selection = null;
            } catch (_) { }
        }
    }


    // Scale special-case: treat 1% as 0.01% for ultra-small extrusion
    function normalizeScalePercent(rawScalePercent) {
        var v = Number(rawScalePercent);
        if (isNaN(v)) return 100;
        // User request: 1% should run as 0.01%
        if (v === 1) return 0.01;
        return v;
    }

    /* 実行関数 / Main execution */
    function doExtrude(applyOffset) {
        __applyOffsetNow = (applyOffset !== false);
        var dist = parseFloat(inputDistance.text) || 0;
        var scale = normalizeScalePercent(parseFloat(inputScale.text));
        if (isNaN(scale) || scale <= 0) scale = 100; // デフォルト 100%
        var ang = -(parseFloat(inputAngle.text) || 0); // 入力角度を反転

        var rad = ang * Math.PI / 180;
        var dx = dist * Math.cos(rad);
        var dy = dist * Math.sin(rad);

        // 実行対象（ベース形状）の確定
        // GroupItem の場合は「複製→合体→expand」した一時パス/複合パスを使う（元は残す）
        var baseItem = originalPath;
        var baseCleanup = function () { };
        var baseSubPaths = originalSubPaths;

        function cleanupTempBaseSafely() {
            try { baseCleanup(); } catch (_) { }
            try { removeAllTempBasesByName(); } catch (_) { }
        }

        if (originalPath.typename === "GroupItem") {
            var built = buildMergedItemFromGroup(originalPath);
            if (!built.ok || !built.item) {
                alert(built.message || L('cannotBuildFromGroup'));
                cleanupTempBaseSafely();
                return;
            }
            baseItem = built.item;
            baseCleanup = built.cleanup;
            baseSubPaths = getSubPaths(baseItem);
            if (!isAllClosed(baseSubPaths)) {
                cleanupTempBaseSafely();
                alert(L('selectClosedGroup'));
                return;
            }
        }

        if (__applyOffsetNow && originalPath.typename === "TextFrame") {
            var builtT = buildMergedItemFromText(originalPath);
            if (!builtT.ok || !builtT.item) {
                alert(builtT.message || L('selectClosed'));
                cleanupTempBaseSafely();
                return;
            }
            baseItem = builtT.item;
            baseCleanup = builtT.cleanup;
            baseSubPaths = getSubPaths(baseItem);
            if (!isAllClosed(baseSubPaths)) {
                cleanupTempBaseSafely();
                alert(L('selectClosed'));
                return;
            }
        }

        // --- プレビュー簡易化 ---
        // プレビュー時は「合体した影」を作らず、複製オブジェクトのみ表示する
        if (!__applyOffsetNow) {
            // Path/Compound/Group（baseItem）用のプレビュー複製（f=1 相当）
            var pathB = null;
            try {
                pathB = baseItem.duplicate();
                // scale: percent (100 = 等倍)
                try { pathB.resize(scale, scale, true, true, true, true, true, Transformation.CENTER); } catch (_) { }
                try { pathB.translate(dx, dy); } catch (_) { }
            } catch (_) { pathB = null; }
            // TextFrame はアウトライン化せず、テキストそのものを複製して表示
            if (originalPath.typename === "TextFrame") {
                function makeTextPreviewDup(f, op) {
                    var d = null;
                    try { d = originalPath.duplicate(); } catch (_) { d = null; }
                    if (!d) return null;

                    // scale interpolation: 100 -> scale
                    var s = 100 + (scale - 100) * f;
                    try { d.resize(s, s, true, true, true, true, true, Transformation.CENTER); } catch (_) { }
                    try { d.translate(dx * f, dy * f); } catch (_) { }
                    try { d.move(originalPath, ElementPlacement.PLACEBEFORE); } catch (_) { }
                    try { d.opacity = op; } catch (_) { }
                    return d;
                }

                // 中間表示（scale=100%でも表示する）
                // 生成順：薄い→濃い（手前が見やすい）
                makeTextPreviewDup(1, 20);
                makeTextPreviewDup(0.75, 40);
                makeTextPreviewDup(0.5, 60);
                makeTextPreviewDup(0.25, 80);

                try { doc.selection = null; } catch (_) { }
                app.redraw();
                return;
            }

            function makePreviewDup(f, op) {
                var d = null;
                try { d = baseItem.duplicate(); } catch (_) { d = null; }
                if (!d) return null;

                var s = 100 + (scale - 100) * f;
                try { d.resize(s, s, true, true, true, true, true, Transformation.CENTER); } catch (_) { }
                try { d.translate(dx * f, dy * f); } catch (_) { }
                try { d.move(originalPath, ElementPlacement.PLACEBEFORE); } catch (_) { }
                try { d.opacity = op; } catch (_) { }
                return d;
            }

            // 中間表示（scale=100%でも表示する）
            // 生成順：薄い→濃い（手前が見やすい）
            try { if (pathB) pathB.move(originalPath, ElementPlacement.PLACEBEFORE); } catch (_) { }
            try { if (pathB) pathB.opacity = 20; } catch (_) { }

            makePreviewDup(0.75, 40);
            makePreviewDup(0.5, 60);
            makePreviewDup(0.25, 80);

            // 選択解除
            try { doc.selection = null; } catch (_) { }

            // Group 由来の一時ベースがある場合は片付け（pathBは残す）
            cleanupTempBaseSafely();

            app.redraw();
            return;
        }

        // --- 実生成（サンプリング面生成） ---
        // v15系の「元形状と複製形状の間を面でつなぐ」方式に差し替え
        var shadowGroup = generateShadowBySampling(baseItem, dx, dy, scale);
        if (!shadowGroup) {
            cleanupTempBaseSafely();
            alert(L('selectClosed'));
            return;
        }

        // パスの穴を潰すため、Merge → Add → Expand の順で実行
        try {
            doc.selection = null;
            shadowGroup.selected = true;
            app.executeMenuCommand('Live Pathfinder Merge');
            app.executeMenuCommand('Live Pathfinder Add');
            app.executeMenuCommand('expandStyle');
        } catch (_) { }

        // 合体後の選択結果をまとめてグループ化（オフセット適用の受け皿）
        var resultItems = null;
        try { resultItems = doc.selection; } catch (_) { resultItems = null; }
        try { doc.selection = null; } catch (_) { }
        var cGroup = null;
        try {
            if (resultItems && resultItems.length) {
                cGroup = groupItems(doc, resultItems);
            } else {
                cGroup = shadowGroup;
            }
        } catch (_) { cGroup = shadowGroup; }


        // オフセット（ライブ効果）をCグループに適用
        applyOffsetEffectToCGroup(cGroup);

        // C を A の背面へ（TextFrame の場合に SENDTOBACK すると見えなくなることがあるため、同一レイヤー内へ配置を寄せる）
        (function placeShadowBehindOriginal(shadowItem, originalItem) {
            if (!shadowItem || !originalItem) return;

            // まずは通常通り「元オブジェクトの直後」へ
            try {
                shadowItem.move(originalItem, ElementPlacement.PLACEAFTER);
                return;
            } catch (_) { }

            // TextFrame や特殊な親のとき move が失敗する場合があるので、同一レイヤーへ移して背面側に置く
            try {
                var lyr = null;
                try { lyr = originalItem.layer; } catch (_) { lyr = null; }
                if (lyr) {
                    // レイヤー先頭（背面側）へ。ドキュメント全体の最背面には送らない。
                    shadowItem.move(lyr, ElementPlacement.PLACEATBEGINNING);
                    return;
                }
            } catch (_) { }

            // 最後の手段：ドキュメント全体の背面（これで見えなくなるケースもあるため最後に）
            try { shadowItem.zOrder(ZOrderMethod.SENDTOBACK); } catch (_) { }
        })(cGroup, originalPath);

        // （TextFrame/GroupItem の一時ベース削除は baseCleanup() 側に統合）

        // 一時ベース（Group由来など）の後始末
        cleanupTempBaseSafely();

        // 最終保険：Pathfinder/expand後に残った一時ベースを完全掃除
        try { removeAllTempBasesByName(); } catch (_) { }
        // 最終保険：一時レイヤー（__LongShadowTempBase__）が残っていれば削除
        try { removeTempLayerByName(); } catch (_) { }

        // 単一パス（PathItem）実行時は、生成したシャドウ（Cグループ）を選択状態にする
        try {
            doc.selection = null;
            if (originalPath && originalPath.typename === "PathItem" && cGroup && cGroup.isValid) {
                doc.selection = [cGroup];
            }
        } catch (_) { }

        app.redraw();
        return;

    }

    /* プレビュー更新 / Update preview */
    function updatePreview() {
        if (isPreviewing) {
            app.undo(); // 前回のプレビューを取り消し
            isPreviewing = false;
        }
        if (chkPreview.value) {
            doExtrude(false);
            isPreviewing = true;
        }
        app.redraw();
    }

    /* イベントリスナー / Event listeners */
    // 値が変わった時
    inputDistance.onChanging = function () { if (chkPreview.value) updatePreview(); };
    inputAngle.onChanging = function () { if (chkPreview.value) updatePreview(); };
    inputScale.onChanging = function () { if (chkPreview.value) updatePreview(); };

    // プレビューチェックボックスがクリックされた時
    chkPreview.onClick = updatePreview;

    // OKボタン
    btnOk.onClick = function () {
        if (isPreviewing) {
            app.undo(); // 一旦プレビューを消してから確定実行
        }
        doExtrude(true);
        dlg.close();
    };

    // キャンセルボタン
    btnCancel.onClick = function () {
        if (isPreviewing) {
            app.undo();
        }
        try { removeAllTempBasesByName(); } catch (_) { }
        try { removeTempLayerByName(); } catch (_) { }
        dlg.close();
    };

    // ダイアログを開いた時点でプレビューを表示
    if (chkPreview && chkPreview.value) {
        updatePreview();
    }
    dlg.show();

})();
