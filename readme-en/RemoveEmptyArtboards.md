# RemoveEmptyArtboards

[![Direct](https://img.shields.io/badge/Direct%20Link-RemoveEmptyArtboards.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/RemoveEmptyArtboards.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RemoveEmptyArtboards.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Finds the artboards with no visible objects on them and deletes them together.

### Features

- A toggle decides whether hidden layers and objects count as occupying an artboard (on by default, meaning they are ignored)
- The target count updates as soon as the toggle changes

### Usage

1. Open the document.
2. Run the script.
3. Check the count and run it.

### Notes

- Visibility is resolved through the ancestors, so hidden items inside groups, and items under hidden groups or layers, are treated the same way.

### Update History

- v1.0.0
