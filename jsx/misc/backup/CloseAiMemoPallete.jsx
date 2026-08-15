#target illustrator
#targetengine "TextMemoEngine"

/* AiMemoPallete（$.global.__TextMemoWindow）だけを閉じる。
   パレットと同じ #targetengine を宣言して $.global を共有し、直接閉じる /
   Close only AiMemoPallete. Declaring the same #targetengine shares $.global
   with the palette, so it can be closed directly. */
if ($.global.__TextMemoWindow) {
    try { $.global.__TextMemoWindow.close(); } catch (e) {}
    $.global.__TextMemoWindow = null;
}
