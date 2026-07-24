#target illustrator

(function () {
    try {
        app.preferences.setBooleanPreference(
            "enableBackgroundSave",
            false
        );

        app.preferences.setBooleanPreference(
            "enableBackgroundExport",
            false
        );

        alert(
            "以下の設定をOFFにしました。\n\n" +
            "・バックグラウンドで保存\n" +
            "・バックグラウンドで書き出し\n\n" +
            "必要に応じてIllustratorを再起動してください。"
        );
    } catch (e) {
        alert(
            "設定を変更できませんでした。\n\n" +
            e.message
        );
    }
})();
