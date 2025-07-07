# ドキュメントで使用されているフォント情報を書き出す

[![Direct](https://img.shields.io/badge/Direct%20Link-ExportFontInfoFromXMP.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/ExportFontInfoFromXMP.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

アクティブなIllustratorドキュメントに埋め込まれているXMPメタデータからすべてのフォント情報を抽出し、次のファイル形式でデスクトップに保存します。

- テキストファイル（タブ区切り）
- CSV
- Markdown

合成フォントの構成フォントも取り出せます。

```markdown
### sw-B

- fontName: ATC-73772d42
- fontFace: 
- fontType: 合成フォント
- composite: True
- fileName: sw-B

#### 構成フォント
- RyoGothicStd-Bold.otf
- NotoSansCJKjp-Light.otf
- ヒラギノ角ゴシック W2.ttc
- RyoGothicStd-Heavy.otf
```

テキストファイル

```xml
fontName:	ATC-73772d42
fontFamily:	sw-B
fontFace:	
fontType:	合成フォント
composite:	True
fileName:	sw-B
構成フォント:
・RyoGothicStd-Bold.otf
・NotoSansCJKjp-Light.otf
・ヒラギノ角ゴシック W2.ttc
・RyoGothicStd-Heavy.otf
・RyoGothicStd-Bold.otf
```

[【Illustrator】ドキュメントで使用されているフォント情報を書き出す｜DTP Transit 別館](https://note.com/dtp_tranist/n/n16e7e95652b6)