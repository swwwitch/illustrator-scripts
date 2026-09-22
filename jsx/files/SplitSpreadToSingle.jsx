#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

見開きページ相当のオブジェクトを検出し、左右2つの片ページに分割します。
対象は選択オブジェクトのみ、またはドキュメント内のすべてから選べます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SplitSpreadToSingle.md

### Overview

Detects spread-like objects and splits them into left and right single pages.
You can process only the selection, or every matching object in the document.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SplitSpreadToSingle.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SplitSpreadToSingle";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-21";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SplitSpreadToSingle.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SplitSpreadToSingle.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 片ページ／見開きと判定する幅の許容差（基準ページ幅に対する比率） / Width tolerance for single / spread detection (ratio of the reference page width) */
    var PAGE_WIDTH_TOLERANCE = 0.15;

    /* 再配置の間隔の初期値 / Default rearrangement spacing */
    var DEFAULT_SPACING_TEXT = "20";

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* パネルの余白 / Panel margins */
    var PANEL_MARGINS = [15, 20, 15, 10];

    /* 間隔欄の幅（文字数） / Width of the spacing fields in characters */
    var SPACING_FIELD_CHARACTERS = 5;

    // =========================================
    // 識別文字列 / Markers
    // =========================================

    /* 分割したグループの note に書く識別子 / Identifier written to the note of split groups */
    var SPLIT_GROUP_NOTE_PREFIX = "__SplitSpreadToSingle__";

    /* 再配置中にアートボードごとのオブジェクトへ一時的に付ける目印 / Temporary marker added to notes while rearranging */
    var REARRANGE_NOTE_MARKER = "__SplitSpreadToSingle_Rearrange__";

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在のUI言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "見開きページを片ページに", en: "Split Spread Pages into Single Pages" }
        },
        panel: {
            target: { ja: "対象", en: "Target" },
            evenPage: { ja: "偶数ページ", en: "Even Pages" },
            postProcess: { ja: "後処理", en: "Post-Process" }
        },
        radio: {
            selectionOnly: { ja: "選択したオブジェクトのみ", en: "Selected Objects Only" },
            all: { ja: "すべて", en: "All" },
            sideRight: { ja: "右", en: "Right" },
            sideLeft: { ja: "左", en: "Left" }
        },
        checkbox: {
            renameArtboards: { ja: "アートボード名を連番でリネーム", en: "Rename Artboards Sequentially" },
            rearrangeArtboards: { ja: "アートボードの再配置", en: "Rearrange Artboards" }
        },
        fieldLabel: {
            spacingHorizontal: { ja: "左右", en: "Horizontal" },
            spacingVertical: { ja: "上下", en: "Vertical" }
        },
        tooltip: {
            modeSelection: { ja: "選択しているアートボードだけを分割します。", en: "Splits only the selected artboards." },
            modeAll: { ja: "ドキュメント内のすべてのアートボードを分割します。", en: "Splits every artboard in the document." },
            sideRight: {
                ja: "偶数ページを見開きの右側として扱います。左綴じ（横書き）向けです。",
                en: "Treats even pages as the right-hand side of the spread, for left-bound documents."
            },
            sideLeft: {
                ja: "偶数ページを見開きの左側として扱います。右綴じ（縦書き）向けです。",
                en: "Treats even pages as the left-hand side of the spread, for right-bound documents."
            },
            rename: { ja: "分割後のアートボードに、ページ番号で名前を付け直します。", en: "Renames the resulting artboards with their page numbers." },
            rearrange: {
                ja: "分割後のアートボードを並べ直します。間隔は下の欄で指定します。",
                en: "Lays the resulting artboards out again. The fields below set the spacing."
            },
            spacingHorizontal: { ja: "並べ直すときの横方向の間隔です。", en: "Horizontal spacing used when the artboards are laid out." },
            spacingVertical: { ja: "並べ直すときの縦方向の間隔です。", en: "Vertical spacing used when the artboards are laid out." }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        generatedName: {
            leftHalfGroup: { ja: "左半分", en: "Left_Half" },
            rightHalfGroup: { ja: "右半分", en: "Right_Half" },
            artboardPrefix: { ja: "アートボード ", en: "Artboard " }
        },
        alert: {
            invalidSpacing: { ja: "アートボードの再配置の間隔には数値を入力してください。", en: "Enter numeric values for artboard rearrangement spacing." },
            noSpreadFoundAll: {
                ja: "見開きページ相当の PlacedItem / GroupItem / RasterItem が見つかりませんでした。",
                en: "No spread-like PlacedItem / GroupItem / RasterItem was found."
            },
            selectSpreadObject: { ja: "分割したい見開きオブジェクトを選択してください。", en: "Select a spread object to split." },
            selectionIsSingle: {
                ja: "選択オブジェクトは片ページ相当です。見開きページ相当のオブジェクトを選択してください。",
                en: "The selected object looks like a single page. Select a spread-like object."
            },
            selectValidSpread: {
                ja: "見開きページ相当の PlacedItem / GroupItem / RasterItem を選択してください。",
                en: "Select a spread-like PlacedItem / GroupItem / RasterItem."
            }
        }
    };

    /**
     * ラベルを取得する（ドット区切りキー）
     * @param {string} labelPath - "panel.target" のようなドット区切りキー
     * @returns {string} 現在のUI言語のラベル（見つからなければキーそのもの）
     */
    function getLabel(labelPath) {
        var pathKeys = String(labelPath).split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return (labelNode[uiLang] != null) ? labelNode[uiLang] : labelPath;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 見出し付きパネルを追加する
     * @param {Window} parentWindow - 追加先
     * @param {string} titleLabelPath - パネルタイトルの LABELS パス
     * @param {string} orientation - "column" または "row"
     * @returns {Panel} 追加したパネル
     */
    function addTitledPanel(parentWindow, titleLabelPath, orientation) {
        var newPanel = parentWindow.add("panel", undefined, getLabel(titleLabelPath));
        newPanel.orientation = orientation;
        newPanel.alignChildren = "left";
        newPanel.margins = PANEL_MARGINS;
        return newPanel;
    }

    /**
     * tooltip 付きのコントロールを追加する
     * @param {Object} parentGroup - 追加先
     * @param {string} controlType - "radiobutton" / "checkbox" など
     * @param {string} textLabelPath - 表示文字の LABELS パス
     * @param {string} tooltipLabelPath - tooltip の LABELS パス
     * @returns {Object} 追加したコントロール
     */
    function addControlWithTip(parentGroup, controlType, textLabelPath, tooltipLabelPath) {
        var newControl = parentGroup.add(controlType, undefined, getLabel(textLabelPath));
        newControl.helpTip = getLabel(tooltipLabelPath);
        return newControl;
    }

    /**
     * 間隔の入力欄を項目名つきで追加する
     * @param {Group} parentGroup - 追加先
     * @param {string} labelKey - LABELS.fieldLabel / LABELS.tooltip のキー
     * @returns {{label: StaticText, field: EditText}} 項目名と入力欄
     */
    function addSpacingField(parentGroup, labelKey) {
        var spacingLabel = parentGroup.add("statictext", undefined, getLabel("fieldLabel." + labelKey));
        var spacingField = parentGroup.add("edittext", undefined, DEFAULT_SPACING_TEXT);
        spacingField.helpTip = getLabel("tooltip." + labelKey);
        spacingField.characters = SPACING_FIELD_CHARACTERS;
        return { label: spacingLabel, field: spacingField };
    }

    /**
     * 設定ダイアログを表示し、選ばれた設定を返す
     * @returns {Object|null} 設定（キャンセル時は null）
     */
    function showSplitDialog() {
        var splitDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        splitDialog.alignChildren = "fill";

        var targetPanel = addTitledPanel(splitDialog, "panel.target", "column");
        addControlWithTip(targetPanel, "radiobutton", "radio.selectionOnly", "tooltip.modeSelection");
        var rbAll = addControlWithTip(targetPanel, "radiobutton", "radio.all", "tooltip.modeAll");
        rbAll.value = true;

        var evenPagePanel = addTitledPanel(splitDialog, "panel.evenPage", "row");
        var rbEvenRight = addControlWithTip(evenPagePanel, "radiobutton", "radio.sideRight", "tooltip.sideRight");
        var rbEvenLeft = addControlWithTip(evenPagePanel, "radiobutton", "radio.sideLeft", "tooltip.sideLeft");
        rbEvenLeft.value = true;

        var postProcessPanel = addTitledPanel(splitDialog, "panel.postProcess", "column");
        var cbRenameArtboards = addControlWithTip(postProcessPanel, "checkbox", "checkbox.renameArtboards", "tooltip.rename");
        cbRenameArtboards.value = true;
        var cbRearrangeArtboards = addControlWithTip(postProcessPanel, "checkbox", "checkbox.rearrangeArtboards", "tooltip.rearrange");
        cbRearrangeArtboards.value = true;

        var spacingGroup = postProcessPanel.add("group");
        spacingGroup.orientation = "row";
        spacingGroup.alignChildren = ["left", "center"];
        var horizontalSpacing = addSpacingField(spacingGroup, "spacingHorizontal");
        var verticalSpacing = addSpacingField(spacingGroup, "spacingVertical");

        /* 間隔欄は再配置がオンのときだけ有効 / Spacing fields apply only when rearranging */
        function updateSpacingEnabled() {
            var rearrangeEnabled = cbRearrangeArtboards.value;
            horizontalSpacing.label.enabled = rearrangeEnabled;
            horizontalSpacing.field.enabled = rearrangeEnabled;
            verticalSpacing.label.enabled = rearrangeEnabled;
            verticalSpacing.field.enabled = rearrangeEnabled;
        }
        cbRearrangeArtboards.onClick = updateSpacingEnabled;
        updateSpacingEnabled();

        var btnRowGroup = splitDialog.add("group");
        btnRowGroup.alignment = "center";
        btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        if (splitDialog.show() !== 1) return null;

        return {
            processAll: rbAll.value,
            evenOnRight: rbEvenRight.value,
            renameArtboards: cbRenameArtboards.value,
            rearrangeArtboards: cbRearrangeArtboards.value,
            spacingHorizontal: parseFloat(horizontalSpacing.field.text),
            spacingVertical: parseFloat(verticalSpacing.field.text)
        };
    }

    // =========================================
    // 見開きの判定 / Spread detection
    // =========================================

    /**
     * 分割できる種類のオブジェクトか（配置画像・グループ・ラスター画像）
     * @param {PageItem} pageItem - 調べるオブジェクト
     * @returns {boolean} 対象なら true
     */
    function isTargetItem(pageItem) {
        return pageItem && (
            pageItem.typename === "PlacedItem" ||
            pageItem.typename === "GroupItem" ||
            pageItem.typename === "RasterItem"
        );
    }

    /**
     * ドキュメント内の分割できる種類のオブジェクトをすべて集める
     * @param {Document} doc - 対象ドキュメント
     * @returns {PageItem[]} 対象のオブジェクト
     */
    function collectTargetItems(doc) {
        var targetItems = [];
        for (var i = 0; i < doc.pageItems.length; i++) {
            var pageItem = doc.pageItems[i];
            if (isTargetItem(pageItem)) {
                targetItems.push(pageItem);
            }
        }
        return targetItems;
    }

    /**
     * 最も重なり面積の大きいアートボードを返す
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} pageItem - 調べるオブジェクト
     * @returns {number} アートボードの番号（重ならなければ -1）
     */
    function getPrimaryArtboardIndex(doc, pageItem) {
        var bounds = pageItem.geometricBounds;
        var bestIndex = -1;
        var bestArea = 0;

        for (var i = 0; i < doc.artboards.length; i++) {
            var abRect = doc.artboards[i].artboardRect;
            var overlapWidth = Math.min(bounds[2], abRect[2]) - Math.max(bounds[0], abRect[0]);
            var overlapHeight = Math.min(bounds[1], abRect[1]) - Math.max(bounds[3], abRect[3]);

            if (overlapWidth > 0 && overlapHeight > 0) {
                var overlapArea = overlapWidth * overlapHeight;
                if (overlapArea > bestArea) {
                    bestArea = overlapArea;
                    bestIndex = i;
                }
            }
        }
        return bestIndex;
    }

    /**
     * ドキュメント内の基準ページ幅（最小アートボード幅）を返す
     * @param {Document} doc - 対象ドキュメント
     * @returns {number} 基準ページ幅（アートボードが無ければ 0）
     */
    function getReferencePageWidth(doc) {
        var minWidth = null;
        for (var i = 0; i < doc.artboards.length; i++) {
            var abRect = doc.artboards[i].artboardRect;
            var abWidth = abRect[2] - abRect[0];
            if (abWidth <= 0) continue;
            if (minWidth === null || abWidth < minWidth) {
                minWidth = abWidth;
            }
        }
        return (minWidth === null) ? 0 : minWidth;
    }

    /**
     * 幅の比率が目標にほぼ一致するか
     * @param {number} ratio - 幅の比率
     * @param {number} targetRatio - 目標の比率（1 = 片ページ、2 = 見開き）
     * @returns {boolean} 許容差の範囲内なら true
     */
    function isNearRatio(ratio, targetRatio) {
        return Math.abs(ratio - targetRatio) <= PAGE_WIDTH_TOLERANCE;
    }

    /**
     * 指定アートボードが片ページ相当か見開き相当かを返す
     * @param {Document} doc - 対象ドキュメント
     * @param {number} artboardIndex - アートボードの番号
     * @param {number} referencePageWidth - 基準ページ幅
     * @returns {string} "single" / "spread" / "other"
     */
    function getArtboardType(doc, artboardIndex, referencePageWidth) {
        var abRect = doc.artboards[artboardIndex].artboardRect;
        var ratio = (abRect[2] - abRect[0]) / referencePageWidth;

        if (isNearRatio(ratio, 1)) return "single";
        if (isNearRatio(ratio, 2)) return "spread";
        return "other";
    }

    /**
     * オブジェクト幅と基準ページ幅から片ページ／見開き／その他を判定する
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} pageItem - 調べるオブジェクト
     * @returns {{artboardIndex: number, pageType: string, artboardType: string}} 判定結果
     */
    function getPageTypeInfo(doc, pageItem) {
        var artboardIndex = getPrimaryArtboardIndex(doc, pageItem);
        var referencePageWidth = (artboardIndex < 0) ? 0 : getReferencePageWidth(doc);
        if (referencePageWidth <= 0) {
            return { artboardIndex: artboardIndex, pageType: "other", artboardType: "other" };
        }

        var bounds = pageItem.geometricBounds;
        var itemWidth = bounds[2] - bounds[0];
        var abRect = doc.artboards[artboardIndex].artboardRect;
        var ratio = itemWidth / referencePageWidth;
        var fitToArtboardRatio = itemWidth / (abRect[2] - abRect[0]);
        var artboardType = getArtboardType(doc, artboardIndex, referencePageWidth);
        var pageType = "other";

        /* 見開きアートボード上でアートボード幅にほぼ一致するオブジェクトは見開き扱い / Treat objects that match artboard width on spread artboards as spread */
        if (artboardType === "spread" && isNearRatio(fitToArtboardRatio, 1)) {
            pageType = "spread";
        } else if (artboardType === "single" && isNearRatio(ratio, 1)) {
            pageType = "single";
        } else if (artboardType === "spread" && isNearRatio(ratio, 2)) {
            pageType = "spread";
        }

        return { artboardIndex: artboardIndex, pageType: pageType, artboardType: artboardType };
    }

    /**
     * 見開きアートボード上の見開きオブジェクトか
     * @param {Object} pageTypeInfo - getPageTypeInfo() の結果
     * @returns {boolean} 分割の対象なら true
     */
    function isSpreadOnSpreadArtboard(pageTypeInfo) {
        return pageTypeInfo.artboardIndex >= 0 &&
            pageTypeInfo.artboardType === "spread" &&
            pageTypeInfo.pageType === "spread";
    }

    /**
     * ドキュメント内の見開きオブジェクトを集める（見つからなければ警告）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Object[]|null} {item, abIndex} の配列（見つからなければ null）
     */
    function collectSpreadsInDocument(doc) {
        var spreads = [];
        var candidates = collectTargetItems(doc);
        for (var i = 0; i < candidates.length; i++) {
            var pageTypeInfo = getPageTypeInfo(doc, candidates[i]);
            if (isSpreadOnSpreadArtboard(pageTypeInfo)) {
                spreads.push({ item: candidates[i], abIndex: pageTypeInfo.artboardIndex });
            }
        }

        if (spreads.length === 0) {
            alert(getLabel("alert.noSpreadFoundAll"));
            return null;
        }
        return spreads;
    }

    /**
     * 選択から見開きオブジェクトを集める（見つからなければ理由に合わせて警告）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Object[]|null} {item, abIndex} の配列（見つからなければ null）
     */
    function collectSpreadsInSelection(doc) {
        if (doc.selection.length < 1) {
            alert(getLabel("alert.selectSpreadObject"));
            return null;
        }

        var spreads = [];
        var singleCount = 0;
        var otherCount = 0;

        for (var i = 0; i < doc.selection.length; i++) {
            var selectedItem = doc.selection[i];
            if (!isTargetItem(selectedItem)) {
                otherCount++;
                continue;
            }
            var pageTypeInfo = getPageTypeInfo(doc, selectedItem);
            if (pageTypeInfo.artboardIndex < 0) {
                otherCount++;
            } else if (isSpreadOnSpreadArtboard(pageTypeInfo)) {
                spreads.push({ item: selectedItem, abIndex: pageTypeInfo.artboardIndex });
            } else if (pageTypeInfo.pageType === "single") {
                singleCount++;
            } else {
                otherCount++;
            }
        }

        if (spreads.length === 0) {
            if (singleCount > 0 && otherCount === 0) {
                alert(getLabel("alert.selectionIsSingle"));
            } else {
                alert(getLabel("alert.selectValidSpread"));
            }
            return null;
        }
        return spreads;
    }

    /**
     * アートボードごとに最初の1件だけを残し、番号の大きい順に並べる
     * （アートボードの挿入で番号がずれるため後ろから処理する）
     * @param {Object[]} spreads - {item, abIndex} の配列
     * @returns {Object[]} 処理順に並べた {item, abIndex} の配列
     */
    function buildWorkList(spreads) {
        var workList = [];
        var usedArtboardMap = {};
        for (var i = 0; i < spreads.length; i++) {
            if (usedArtboardMap[spreads[i].abIndex]) continue;
            usedArtboardMap[spreads[i].abIndex] = true;
            workList.push(spreads[i]);
        }
        workList.sort(function (a, b) { return b.abIndex - a.abIndex; });
        return workList;
    }

    // =========================================
    // 分割 / Splitting
    // =========================================

    /**
     * 矩形でクリップしたグループにオブジェクトを入れる
     * @param {PageItem} targetItem - クリップするオブジェクト
     * @param {number[]} clipRect - クリップ範囲 [left, top, right, bottom]
     * @param {string} groupName - グループ名
     * @returns {GroupItem} クリッピンググループ
     */
    function createClipGroup(targetItem, clipRect, groupName) {
        var parentContainer = targetItem.parent;

        var clipPath = parentContainer.pathItems.rectangle(
            clipRect[1],
            clipRect[0],
            clipRect[2] - clipRect[0],
            clipRect[1] - clipRect[3]
        );
        clipPath.stroked = false;
        clipPath.filled = false;

        var clipGroup = parentContainer.groupItems.add();
        clipGroup.name = groupName;

        targetItem.move(clipGroup, ElementPlacement.PLACEATEND);
        clipPath.move(clipGroup, ElementPlacement.PLACEATBEGINNING);

        clipPath.clipping = true;
        clipGroup.clipped = true;

        return clipGroup;
    }

    /**
     * 見開きオブジェクトを左右2つにクリップし、アートボードも2つに分ける
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} spreadItem - 見開きオブジェクト
     * @param {number} abIndex - 見開きアートボードの番号
     * @param {boolean} evenOnRight - 偶数ページを右にするか
     * @param {number} pageOffset - 分けたアートボードの間隔の半分
     * @returns {void}
     */
    function splitSpread(doc, spreadItem, abIndex, evenOnRight, pageOffset) {
        /* 元アートボードの位置とサイズを基準にする / Use the source artboard bounds as the base */
        var abRect = doc.artboards[abIndex].artboardRect;
        var left = abRect[0];
        var top = abRect[1];
        var right = abRect[2];
        var bottom = abRect[3];
        var centerX = left + (right - left) / 2;

        /* 複製（B） / Duplicate for right half */
        var rightHalfSource = spreadItem.duplicate();

        /* アートボード基準の左右ページ矩形 / Left and right page rects based on the artboard */
        var leftClipRect = [left, top, centerX, bottom];
        var rightClipRect = [centerX, top, right, bottom];

        /* 元オブジェクト(A) → 左半分、複製(B) → 右半分 / Original object to left half, duplicate to right half */
        var leftHalfGroup = createClipGroup(spreadItem, leftClipRect, getLabel("generatedName.leftHalfGroup"));
        var rightHalfGroup = createClipGroup(rightHalfSource, rightClipRect, getLabel("generatedName.rightHalfGroup"));

        var leftPageRect;
        var rightPageRect;
        var leftHalfTargetLeft;
        var rightHalfTargetLeft;

        if (evenOnRight) {
            rightPageRect = [centerX, top, right, bottom];
            leftPageRect = [left - pageOffset, top, centerX - pageOffset, bottom];
            /* 左半分（A）は右ページへ、右半分（B）は左ページへ移す / Move left half to right page and right half to left page */
            leftHalfTargetLeft = rightPageRect[0];
            rightHalfTargetLeft = leftPageRect[0];
        } else {
            /* 偶数ページが左（デフォルト） / Even pages on the left (default) */
            leftPageRect = [left, top, centerX, bottom];
            rightPageRect = [centerX + pageOffset, top, right + pageOffset, bottom];
            /* 左半分（A）は左ページのまま、右半分（B）は右ページへ移す / Keep left half on left page and move right half to right page */
            leftHalfTargetLeft = leftPageRect[0];
            rightHalfTargetLeft = rightPageRect[0];
        }

        /* 見た目の左→右に合わせて、元のアートボードを左ページ、追加アートボードを右ページにする / Original artboard becomes the left page, the inserted one the right page */
        doc.artboards[abIndex].artboardRect = leftPageRect;
        doc.artboards.setActiveArtboardIndex(abIndex);
        doc.artboards.insert(rightPageRect, abIndex + 1);

        /* クリップ範囲から配置先までの移動量 / Translation from the clip rects to the destination pages */
        leftHalfGroup.translate(leftHalfTargetLeft - leftClipRect[0], 0);
        rightHalfGroup.translate(rightHalfTargetLeft - rightClipRect[0], 0);

        leftHalfGroup.note = SPLIT_GROUP_NOTE_PREFIX + ":role=A";
        rightHalfGroup.note = SPLIT_GROUP_NOTE_PREFIX + ":role=B";

        leftHalfGroup.selected = true;
        rightHalfGroup.selected = true;
    }

    // =========================================
    // アートボード再配置 / Artboard rearrangement
    // =========================================

    /**
     * レイヤーとオブジェクトのロック・非表示をすべて解除する（入れ子もたどる）
     * @param {Object} container - Document / Layer / GroupItem
     * @returns {void}
     */
    function unlockAndUnhideAll(container) {
        if (!container) return;

        if (container.typename === "Document" || container.typename === "Layer") {
            var layers = container.layers;
            for (var i = 0; i < layers.length; i++) {
                var layer = layers[i];
                if (layer.locked) layer.locked = false;
                if (!layer.visible) layer.visible = true;
                unlockAndUnhideAll(layer);
            }
        }

        if (container.pageItems) {
            var pageItems = container.pageItems;
            for (var j = 0; j < pageItems.length; j++) {
                var pageItem = pageItems[j];
                if (pageItem.locked) pageItem.locked = false;
                if (pageItem.hidden) pageItem.hidden = false;
                if (pageItem.typename === "GroupItem") {
                    unlockAndUnhideAll(pageItem);
                }
            }
        }
    }

    /**
     * アートボードごとに乗っているオブジェクトを集める
     * （2つのアートボードにまたがるものは最初のアートボードだけに入れる）
     * @param {Document} doc - 対象ドキュメント
     * @returns {PageItem[][]} アートボード番号ごとのオブジェクト
     */
    function collectItemsByArtboard(doc) {
        var artboards = doc.artboards;
        var artboardItemMap = [];
        var markedItems = [];
        var originalNotes = [];

        for (var i = 0; i < artboards.length; i++) {
            artboards.setActiveArtboardIndex(i);
            doc.selection = null;
            doc.selectObjectsOnActiveArtboard();

            var itemsOnArtboard = [];
            for (var j = 0; j < doc.selection.length; j++) {
                var selectedItem = doc.selection[j];
                if (!selectedItem) continue;

                /* 目印付き＝前のアートボードで集めたもの / A marked item was already collected */
                var noteText = String(selectedItem.note || "");
                if (noteText.indexOf(REARRANGE_NOTE_MARKER) >= 0) continue;

                markedItems.push(selectedItem);
                originalNotes.push(noteText);
                selectedItem.note = noteText ? (noteText + "\n" + REARRANGE_NOTE_MARKER) : REARRANGE_NOTE_MARKER;
                itemsOnArtboard.push(selectedItem);
            }

            artboardItemMap[i] = itemsOnArtboard;
            doc.selection = null;
        }

        /* 目印を外して note を戻す / Restore the original notes */
        for (var k = 0; k < markedItems.length; k++) {
            markedItems[k].note = originalNotes[k];
        }
        return artboardItemMap;
    }

    /**
     * アートボードをページ順（見開きの並び）に並べ直し、乗っているオブジェクトも一緒に動かす
     * @param {Document} doc - 対象ドキュメント
     * @param {number} spacingX - 横方向の間隔
     * @param {number} spacingY - 縦方向の間隔
     * @param {boolean} evenOnRight - 偶数ページを右にするか
     * @returns {void}
     */
    function rearrangeArtboardsByPageOrder(doc, spacingX, spacingY, evenOnRight) {
        if (!doc || !doc.artboards || doc.artboards.length === 0) return;

        unlockAndUnhideAll(doc);

        var artboards = doc.artboards;
        var firstRect = artboards[0].artboardRect;
        var baseX = firstRect[0];
        var baseY = firstRect[1];
        var firstPageRight = !evenOnRight;
        var artboardItemMap = collectItemsByArtboard(doc);

        for (var i = 0; i < artboards.length; i++) {
            var artboard = artboards[i];
            var itemsOnArtboard = artboardItemMap[i] || [];
            var oldRect = artboard.artboardRect;
            var pageWidth = oldRect[2] - oldRect[0];
            var pageHeight = oldRect[1] - oldRect[3];
            var pageNumber = i + 1;
            var rowIndex;
            var newLeft;

            if (pageNumber === 1) {
                rowIndex = 0;
                newLeft = baseX;
            } else {
                var isThisPageRight;
                if (firstPageRight) {
                    isThisPageRight = (pageNumber % 2 !== 0);
                    newLeft = isThisPageRight ? baseX : (baseX - pageWidth - spacingX);
                } else {
                    isThisPageRight = (pageNumber % 2 === 0);
                    newLeft = isThisPageRight ? (baseX + pageWidth + spacingX) : baseX;
                }
                rowIndex = Math.floor(pageNumber / 2);
            }

            var newTop = baseY - rowIndex * (pageHeight + spacingY);
            var deltaX = newLeft - oldRect[0];
            var deltaY = newTop - oldRect[1];

            artboard.artboardRect = [newLeft, newTop, newLeft + pageWidth, newTop - pageHeight];

            if (deltaX !== 0 || deltaY !== 0) {
                for (var k = 0; k < itemsOnArtboard.length; k++) {
                    itemsOnArtboard[k].translate(deltaX, deltaY, true, true, true, true);
                }
            }
        }

        app.executeMenuCommand("fitall");
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 分けたアートボードの間隔の半分を、新規ドキュメントの「アートボードの間隔」から求める
     * @returns {number} 間隔の半分（読めなければ 5）
     */
    function getPageOffset() {
        /* 環境設定が読めなければ既定値 / fall back when the preference cannot be read */
        try {
            return app.preferences.getRealPreference("artnewdialog/artboardSpacing") / 2;
        } catch (e) {
            return 5;
        }
    }

    /**
     * 見開きオブジェクトを片ページに分割する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            return;
        }

        var doc = app.activeDocument;

        var splitOptions = showSplitDialog();
        if (!splitOptions) return;

        if (splitOptions.rearrangeArtboards && (isNaN(splitOptions.spacingHorizontal) || isNaN(splitOptions.spacingVertical))) {
            alert(getLabel("alert.invalidSpacing"));
            return;
        }

        var pageOffset = getPageOffset();

        /* 対象オブジェクトの収集 / Collect target objects */
        var spreads = splitOptions.processAll ? collectSpreadsInDocument(doc) : collectSpreadsInSelection(doc);
        if (!spreads) return;

        var workList = buildWorkList(spreads);

        /* 各オブジェクトを処理 / Process each object */
        doc.selection = null;
        for (var i = 0; i < workList.length; i++) {
            splitSpread(doc, workList[i].item, workList[i].abIndex, splitOptions.evenOnRight, pageOffset);
        }

        /* アートボード名をリネーム（オプション） / Rename artboards (optional) */
        if (splitOptions.renameArtboards) {
            for (var k = 0; k < doc.artboards.length; k++) {
                doc.artboards[k].name = getLabel("generatedName.artboardPrefix") + (k + 1);
            }
        }

        /* アートボード再配置（オプション） / Rearrange artboards (optional) */
        if (splitOptions.rearrangeArtboards) {
            rearrangeArtboardsByPageOrder(doc, splitOptions.spacingHorizontal, splitOptions.spacingVertical, splitOptions.evenOnRight);
        }
    }

    main();

})();
