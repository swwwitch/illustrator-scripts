# RegisterGraphicStyleWithText

[![Direct](https://img.shields.io/badge/Direct%20Link-RegisterGraphicStyleWithText.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/style/RegisterGraphicStyleWithText.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RegisterGraphicStyleWithText.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Registers the appearance of the selected object as a graphic style, using the selected text's content as the style name.

### Features

- Accepts one text frame plus one object at the top level of the selection
- Also accepts a single group containing one text frame plus one object
- With several groups selected, each group is processed in turn
- An existing style with the same name is removed first, so repeated runs never create duplicates

### Usage

1. Select a text frame and the object whose appearance you want to register.
2. Run the script.

### Notes

- The text content becomes the style name, so keep it short and unique.
- Use register-temp-style.jsx when you want a fixed style name instead.

### Update History

- v1.1.0
