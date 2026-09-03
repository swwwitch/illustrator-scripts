#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ポイント文字およびエリア内文字の［字形の境界に整列］を切り替えながら、垂直方向（上・中央・下）に整列します。
チェックボックスやラジオボタンの操作はそのつどプレビューへ反映され、T / M / B キーでも整列位置を切り替えられます。

詳細は README を参照してください。

### Overview

Aligns objects vertically — top, center or bottom — while toggling Align to Glyph Bounds for point text and area text.
Every checkbox and radio button refreshes the preview, and the T, M and B keys switch the alignment position.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartVerticalAlign";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-04";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-03";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartVerticalAlign.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartVerticalAlign.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n9ee716675032"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* ［プレビュー境界］の初期状態 / Initial state of the Preview Bounds checkbox */
    var DEFAULT_USE_PREVIEW_BOUNDS = false;

    // =========================================
    // 環境設定キー / Preference keys
    // =========================================

    /* 境界にプレビュー境界（線幅・効果）を含めるか / Include stroke and effects in bounds */
    var PREF_KEY_INCLUDE_STROKE_IN_BOUNDS = 'includeStrokeInBounds';

    // =========================================
    // 整列・字形境界の定義 / Alignment and glyph-bounds tables
    // =========================================

    /* 整列位置。ラジオボタン・メニューコマンド・ショートカットキーを1か所で対応付ける
       / Alignment options: radio button, menu command and shortcut key in one table */
    var ALIGN_OPTIONS = [
        { labelKey: 'radio.top',    menuCommand: 'Vertical Align Top',    shortcutKey: 'T' },
        { labelKey: 'radio.center', menuCommand: 'Vertical Align Center', shortcutKey: 'M' },
        { labelKey: 'radio.bottom', menuCommand: 'Vertical Align Bottom', shortcutKey: 'B' }
    ];
    var ALIGN_INDEX_CENTER = 1;
    var ALIGN_INDEX_BOTTOM = 2;

    /* ［字形の境界に整列］のチェックボックスと、対応する環境設定キー・文字種
       / Glyph-bounds checkboxes with their preference key and text kind */
    var GLYPH_BOUNDS_OPTIONS = [
        { labelKey: 'checkbox.pointText', prefKey: 'EnableActualPointTextSpaceAlign', textKind: TextType.POINTTEXT },
        { labelKey: 'checkbox.areaText',  prefKey: 'EnableActualAreaTextSpaceAlign',  textKind: TextType.AREATEXT }
    ];
    var GLYPH_INDEX_POINT_TEXT = 0;

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ダイアログの初期位置・不透明度 / Dialog position & opacity */
    var DIALOG_OFFSET_X = 300;  /* 右(+)／左(-) / shift right (+) / left (-) */
    var DIALOG_OFFSET_Y = 0;    /* 下(+)／上(-) / shift down (+) / up (-) */
    var DIALOG_OPACITY = 0.97;  /* 0.0 - 1.0 */

    /* 余白と間隔 / Margins and spacing */
    var PANEL_MARGINS = [15, 20, 15, 15];      /* パネル余白 [左,上,右,下] */
    var RADIO_SPACING = 5;                     /* ラジオボタンの行間 / spacing between radio buttons */
    var PREVIEW_ROW_MARGINS = [15, 0, 15, 0];  /* ［プレビュー境界］行の余白 */
    var BUTTON_WIDTH = 90;

    /**
     * パネルの共通設定
     * @param {Panel} panel - 対象パネル
     * @param {number} [spacing] - 要素間隔（省略時はScriptUIの既定値のまま）
     * @returns {void}
     */
    function setupPanel(panel, spacing) {
        panel.orientation = 'column';
        panel.alignChildren = ['left', 'top'];
        panel.margins = PANEL_MARGINS;
        if (typeof spacing === 'number') {
            panel.spacing = spacing;
        }
    }

    /**
     * ダイアログの表示位置をずらす
     * @param {Window} targetDialog - 対象ダイアログ
     * @param {number} offsetX - 横方向のオフセット
     * @param {number} offsetY - 縦方向のオフセット
     * @returns {void}
     */
    function shiftDialogPosition(targetDialog, offsetX, offsetY) {
        targetDialog.onShow = function () {
            targetDialog.location = [
                targetDialog.location[0] + offsetX,
                targetDialog.location[1] + offsetY
            ];
        };
    }

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

    /* ラベル定義（カテゴリ別）/ Label definitions (by category) */
    var LABELS = {
        dialog: {
            title: { ja: "垂直方向の整列", en: "Vertical Alignment" }
        },
        panel: {
            glyphBounds: { ja: "字形の境界に整列", en: "Align to Glyph Bounds" },
            alignment: { ja: "整列", en: "Alignment" }
        },
        checkbox: {
            pointText: { ja: "ポイント文字", en: "Point Text" },
            areaText: { ja: "エリア内文字", en: "Area Text" },
            previewBounds: { ja: "プレビュー境界", en: "Preview Bounds" }
        },
        radio: {
            top: { ja: "上", en: "Top" },
            center: { ja: "中央", en: "Center" },
            bottom: { ja: "下", en: "Bottom" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        }
    };

    /**
     * ラベルを取得する（ドット区切りキー）
     * @param {string} key - "panel.alignment" のようなドット区切りキー
     * @returns {string} 現在のUI言語のラベル（見つからなければキーそのもの）
     */
    function getLabel(key) {
        var labelNode = LABELS;
        var keyParts = String(key).split('.');
        for (var i = 0; i < keyParts.length; i++) {
            if (labelNode == null) break;
            labelNode = labelNode[keyParts[i]];
        }
        return String((labelNode && labelNode[uiLang] != null) ? labelNode[uiLang] : key);
    }

    // =========================================
    // 選択オブジェクトの取得 / Selection
    // =========================================

    /**
     * 整列の対象になるオブジェクトか判定する
     * @param {PageItem} pageItem - 判定するオブジェクト
     * @returns {boolean} 対象なら true
     */
    function isAlignableItem(pageItem) {
        if (!pageItem) return false;
        return (pageItem.typename === "TextFrame" ||
            pageItem.typename === "PathItem" ||
            pageItem.typename === "GroupItem" ||
            pageItem.typename === "CompoundPathItem");
    }

    /**
     * 選択中の整列対象オブジェクトを配列で取得する
     * @returns {Array<PageItem>} 整列対象のオブジェクト
     */
    function getAlignableSelection() {
        var currentSelection = app.activeDocument.selection;
        var alignableItems = [];
        for (var i = 0; i < currentSelection.length; i++) {
            if (isAlignableItem(currentSelection[i])) {
                alignableItems.push(currentSelection[i]);
            }
        }
        return alignableItems;
    }

    // =========================================
    // 整列プレビュー / Alignment preview
    // =========================================

    /* ALIGN_OPTIONS と同じ並びのラジオボタン。UI構築時に埋める
       / Radio buttons in ALIGN_OPTIONS order, filled while the dialog is built */
    var alignRadios = [];

    /**
     * 指定したインデックスの整列位置だけを選択状態にする
     * @param {number} optionIndex - ALIGN_OPTIONS のインデックス
     * @returns {void}
     */
    function selectAlignOption(optionIndex) {
        for (var i = 0; i < alignRadios.length; i++) {
            alignRadios[i].value = (i === optionIndex);
        }
    }

    /**
     * 選択中の整列位置のメニューコマンドを返す
     * @returns {string|null} メニューコマンド名（未選択なら null）
     */
    function getSelectedAlignCommand() {
        for (var i = 0; i < alignRadios.length; i++) {
            if (alignRadios[i].value) return ALIGN_OPTIONS[i].menuCommand;
        }
        return null;
    }

    /**
     * 現在の設定で整列を実行し、プレビューへ即時反映する
     * @returns {void}
     */
    function applyPreviewAlignment() {
        var menuCommand = getSelectedAlignCommand();
        if (menuCommand && getAlignableSelection().length > 0) {
            /* 環境設定の変更を境界に反映させてから整列する（OFFに戻したときも効かせるため）
               / Redraw first so the new preference is reflected in the bounds, including when it is turned OFF */
            app.redraw();
            app.executeMenuCommand(menuCommand);
        }
        app.redraw();
    }

    // =========================================
    // UI構築 / Dialog
    // =========================================

    /**
     * ［字形の境界に整列］パネルを追加する
     * @param {Window} targetDialog - 追加先のダイアログ
     * @param {Array<PageItem>} alignableItems - 選択中の整列対象オブジェクト
     * @returns {Array<Checkbox>} GLYPH_BOUNDS_OPTIONS と同じ並びのチェックボックス
     */
    function addGlyphBoundsPanel(targetDialog, alignableItems) {
        var glyphBoundsPanel = targetDialog.add('panel', undefined, getLabel('panel.glyphBounds'));
        setupPanel(glyphBoundsPanel);

        /* 選択中の文字種。ポイント文字／エリア内文字のときだけ他方をディムする
           / Kind of the selected text; dims the checkbox for the other kind */
        var selectedTextKind = (alignableItems.length > 0) ? alignableItems[0].kind : null;

        var glyphBoundsCheckboxes = [];
        for (var i = 0; i < GLYPH_BOUNDS_OPTIONS.length; i++) {
            glyphBoundsCheckboxes.push(addGlyphBoundsCheckbox(glyphBoundsPanel, GLYPH_BOUNDS_OPTIONS[i], selectedTextKind));
        }
        return glyphBoundsCheckboxes;
    }

    /**
     * ［字形の境界に整列］のチェックボックスを1つ追加し、環境設定とプレビューに結び付ける
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {Object} glyphOption - GLYPH_BOUNDS_OPTIONS の1項目
     * @param {TextType} selectedTextKind - 選択中の文字種（テキスト以外は null）
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addGlyphBoundsCheckbox(parentPanel, glyphOption, selectedTextKind) {
        var checkbox = parentPanel.add('checkbox', undefined, getLabel(glyphOption.labelKey));
        checkbox.value = app.preferences.getBooleanPreference(glyphOption.prefKey);

        /* 選択が別の文字種ならディム / Dim when the selection is the other text kind */
        if (selectedTextKind != null && selectedTextKind !== glyphOption.textKind) {
            checkbox.enabled = false;
        }

        /* ON/OFFで環境設定を書き換え、そのままプレビューを更新
           / Write the preference and refresh the preview on every toggle */
        checkbox.onClick = function () {
            app.preferences.setBooleanPreference(glyphOption.prefKey, checkbox.value === true);
            applyPreviewAlignment();
        };
        return checkbox;
    }

    /**
     * ［整列］パネルを追加し、ラジオボタンを alignRadios に登録する
     * @param {Window} targetDialog - 追加先のダイアログ
     * @returns {void}
     */
    function addAlignmentPanel(targetDialog) {
        var alignmentPanel = targetDialog.add('panel', undefined, getLabel('panel.alignment'));
        setupPanel(alignmentPanel, RADIO_SPACING);

        for (var i = 0; i < ALIGN_OPTIONS.length; i++) {
            var alignRadio = alignmentPanel.add('radiobutton', undefined, getLabel(ALIGN_OPTIONS[i].labelKey));
            alignRadio.onClick = applyPreviewAlignment;
            alignRadios.push(alignRadio);
        }
    }

    /**
     * ［プレビュー境界］の行を追加する
     * @param {Window} targetDialog - 追加先のダイアログ
     * @returns {void}
     */
    function addPreviewBoundsRow(targetDialog) {
        var previewBoundsRow = targetDialog.add('group');
        previewBoundsRow.orientation = 'row';
        previewBoundsRow.alignChildren = ['left', 'center'];
        previewBoundsRow.margins = PREVIEW_ROW_MARGINS;

        var previewBoundsCheckbox = previewBoundsRow.add('checkbox', undefined, getLabel('checkbox.previewBounds'));
        previewBoundsCheckbox.value = DEFAULT_USE_PREVIEW_BOUNDS;
        previewBoundsCheckbox.onClick = function () {
            /* ONで線幅・効果を境界に含める / ON: include stroke and effects in the bounds */
            app.preferences.setBooleanPreference(PREF_KEY_INCLUDE_STROKE_IN_BOUNDS, previewBoundsCheckbox.value === true);
            applyPreviewAlignment();
        };
    }

    /**
     * ボタンエリアを追加する
     * @param {Window} targetDialog - 追加先のダイアログ
     * @returns {void}
     */
    function addButtonRow(targetDialog) {
        var btnRowGroup = targetDialog.add('group');
        btnRowGroup.orientation = 'row';
        btnRowGroup.alignChildren = ['right', 'center'];
        btnRowGroup.alignment = ['right', 'bottom'];

        var btnCancel = btnRowGroup.add('button', undefined, getLabel('button.cancel'), { name: 'cancel' });
        btnCancel.preferredSize.width = BUTTON_WIDTH;
        btnCancel.onClick = function () {
            targetDialog.close();
        };

        var btnOK = btnRowGroup.add('button', undefined, getLabel('button.ok'), { name: 'ok' });
        btnOK.preferredSize.width = BUTTON_WIDTH;
        btnOK.onClick = function () {
            targetDialog.close();
        };
    }

    /**
     * T / M / B キーで整列位置を切り替えるハンドラーを登録する
     * @param {Window} targetDialog - 対象ダイアログ
     * @returns {void}
     */
    function addAlignmentKeyHandler(targetDialog) {
        targetDialog.addEventListener("keydown", function (event) {
            for (var i = 0; i < ALIGN_OPTIONS.length; i++) {
                if (event.keyName !== ALIGN_OPTIONS[i].shortcutKey) continue;
                selectAlignOption(i);
                applyPreviewAlignment();
                event.preventDefault();
                return;
            }
        });
    }

    /**
     * 選択内容に応じて整列位置とポイント文字チェックボックスの初期値を決める
     * @param {Array<PageItem>} alignableItems - 選択中の整列対象オブジェクト
     * @param {Checkbox} pointTextCheckbox - ポイント文字のチェックボックス
     * @returns {void}
     */
    function applyDefaultAlignment(alignableItems, pointTextCheckbox) {
        var hasTextFrame = false;
        var hasOtherItem = false;
        for (var i = 0; i < alignableItems.length; i++) {
            if (alignableItems[i].typename === "TextFrame") {
                hasTextFrame = true;
            } else {
                hasOtherItem = true;
            }
        }

        if (hasTextFrame && !hasOtherItem) {
            /* テキストのみ：下揃え / Text only: align bottom */
            pointTextCheckbox.value = false;
            selectAlignOption(ALIGN_INDEX_BOTTOM);
            return;
        }

        /* テキストと図形の混在、または図形のみ：中央揃え / Mixed or shapes only: align center */
        if (hasTextFrame) {
            pointTextCheckbox.value = true;
        }
        selectAlignOption(ALIGN_INDEX_CENTER);
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログを組み立てて表示する
     * @returns {void}
     */
    function main() {
        var alignableItems = getAlignableSelection();

        var dialog = new Window('dialog');
        dialog.text = getLabel('dialog.title') + ' ' + SCRIPT_VERSION;
        dialog.orientation = 'column';
        dialog.alignChildren = ['fill', 'top'];
        dialog.opacity = DIALOG_OPACITY;
        shiftDialogPosition(dialog, DIALOG_OFFSET_X, DIALOG_OFFSET_Y);

        var glyphBoundsCheckboxes = addGlyphBoundsPanel(dialog, alignableItems);
        addAlignmentPanel(dialog);
        addPreviewBoundsRow(dialog);
        addButtonRow(dialog);
        addAlignmentKeyHandler(dialog);

        applyDefaultAlignment(alignableItems, glyphBoundsCheckboxes[GLYPH_INDEX_POINT_TEXT]);

        dialog.show();
    }

    main();

})();
