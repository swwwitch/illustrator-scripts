# RemoveOtherArtboards-v2

[![Direct](https://img.shields.io/badge/Direct%20Link-RemoveOtherArtboards--v2.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/RemoveOtherArtboards-v2.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RemoveOtherArtboards-v2.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Remove all non-active artboards
- Delete objects outside the active artboard (including guides)

### Key Features

- Delete non-active artboards
- Delete objects outside the active artboard
- Temporarily relax and restore lock/visible/template state of layers/sublayers/groups
- Delete object guides not overlapping the active artboard (ruler guides excluded)
- Temporarily relax and restore locked/hidden state of PageItems

### Processing Flow

1) Get the active document and active artboard
2) Remove non-active artboards from the end
3) Recursively collect and temporarily relax the state of layers/sublayers/groups/PageItems
4) Delete object guides not overlapping the active artboard
5) Inverse-select and delete objects outside the active artboard
6) Restore original states

### Update History

- v1.0 (202406XX) : Initial version
- v1.1 (20250815) : Expanded docs, notes on lock restore & guide checks, comment cleanup
- v1.2 (20250815) : Added PageItem locked/hidden handling, terminology unification, structured comments

### Script info

- Version: v1.2
