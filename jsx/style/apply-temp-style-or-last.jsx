#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

    /*
      選択オブジェクトに既存のグラフィックスタイル「temp_style」を適用
      Apply the existing graphic style "temp_style" to selected objects
    
     ### オリジナルアイデア
    
    @comsk(asa me)さん
    https://qiita.com/comsk/items/87161b2b7d2336b161c4

    Gorolib Designさん
    https://gorolib.blog.jp/archives/73930467.html#google_vignette
    */

    (function () {

        var SCRIPT_VERSION = "v1.1.0";
        var TARGET_STYLE_NAME = "temp_style";

        /* === ローカライズ / Localization === */
        var lang = ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';

        var LABELS = {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "オブジェクトを選択してください", en: "Please select objects" },
            noStyles: { ja: "グラフィックスタイルが存在しません", en: "No graphic styles in this document" },
            noneApplied: { ja: "スタイルを適用できるオブジェクトがありません", en: "No object accepted the style" }
        };

        /* キー指定でラベル取得 / Get label by key */
        function getLabel(key) {
            var entry = LABELS[key];
            if (!entry) return key;
            return entry[lang] || entry.en || key;
        }

        /* === コアロジック / Core logic === */
        /* 指定名のスタイルがなければ最後のグラフィックスタイルを返す
           Fall back to the last graphic style if the named one is missing */
        function resolveStyleWithFallback(doc, preferredName) {
            try {
                return doc.graphicStyles.getByName(preferredName);
            } catch (e) { }
            var styles = doc.graphicStyles;
            if (styles.length === 0) return null;
            return styles[styles.length - 1];
        }

        function applyStyleToSelection(doc, style) {
            var sel = doc.selection;
            var appliedCount = 0;
            for (var i = 0; i < sel.length; i++) {
                try {
                    style.applyTo(sel[i]);
                    appliedCount++;
                } catch (e) {
                    // 適用できないアイテムはスキップ / skip items that cannot accept the style
                }
            }
            return appliedCount;
        }

        /* === エントリポイント / Entry point === */
        function main() {
            if (app.documents.length === 0) {
                alert(getLabel("noDocument"));
                return;
            }
            var doc = app.activeDocument;

            if (doc.selection.length === 0) {
                alert(getLabel("noSelection"));
                return;
            }

            var style = resolveStyleWithFallback(doc, TARGET_STYLE_NAME);
            if (!style) {
                alert(getLabel("noStyles"));
                return;
            }

            var appliedCount = applyStyleToSelection(doc, style);
            if (appliedCount === 0) {
                alert(getLabel("noneApplied"));
            }
        }

        main();

    })();
