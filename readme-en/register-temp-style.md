# register-temp-style

[![Direct](https://img.shields.io/badge/Direct%20Link-register--temp--style.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/style/register-temp-style.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/register-temp-style.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Registers the appearance of the selected object as a graphic style with a fixed name, bypassing Illustrator's default "Untitled" registration.
- If a style with the same name already exists, it is removed first so repeated runs do not create duplicates.

### Process Flow

1. Exit silently unless exactly one object is selected.
2. Remove the existing style with the same name, if any.
3. Load, run, and unload a temporary action that appends an unnamed graphic style.
4. Rename the last style to `TEMP_STYLE_NAME`.

### Settings

- Edit `TEMP_STYLE_NAME` to change the registered style name.

### Original idea

@comsk (asa me)

https://qiita.com/comsk/items/87161b2b7d2336b161c4

### Script info

- Version: v1.0.0
