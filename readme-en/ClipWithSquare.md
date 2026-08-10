# ClipWithSquare

[![Direct](https://img.shields.io/badge/Direct%20Link-ClipWithSquare.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/mask/ClipWithSquare.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ClipWithSquare.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Creates the smallest square path centered on the selected image (linked or embedded), including images inside a clipping mask group
- Builds a new clipping group from that square

### Main Features

- Handles linked images, embedded images and the clipping masks that contain them
- Works on images on locked or template layers as well, by creating a working layer for the operation

### Process Flow

1. Walk the selected objects
2. Release any clipping mask and extract the image alone
3. Build the smallest square from `visibleBounds`
4. Put the square and the image into one group and turn it into a clipping group
5. Leave the created group selected

### Update History

- v1.0 (20231126): Initial version
- v1.1 (20250813): Stabilized the release-and-rebuild of clipping masks, and tidied the comments

### Script info

- Version: v1.1
