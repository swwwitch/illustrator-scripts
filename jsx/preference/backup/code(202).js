// Illustrator バックグラウンド保存・書き出しをOFFにするスクリプト

#target illustrator

try {
    // バックグラウンド保存をOFF
    app.preferences.setBooleanPreference("enableBackgroundSave", false);

    // バックグラウンド書き出しをOFF
    app.preferences.setBooleanPreference("enableBackgroundExport", false);

    alert("バックグラウンド保存：OFF\nバックグラウンド書き出し：OFF\n設定を変更しました。");
} catch (e) {
    alert("エラーが発生しました：" + e.message);
}
