# Keep a history of what you locked and bring it back

[![Direct](https://img.shields.io/badge/Direct%20Link-LockHistoryPalette.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/select/LockHistoryPalette.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/LockHistoryPalette.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

A persistent palette that records every batch of objects you lock with its Lock button as one history entry.

Illustrator's Object > Unlock All releases every lock in the document at once. This script remembers what you locked and when, one entry at a time, so you can **release just the batch you locked last**.

The record is written into each object's **tag** (`PageItem.tags`, tag name `LockHistoryPalette`) and stored in the document. The history comes back after you close the palette, and after you close and reopen the file.

### Features

- Lock records the current selection as one history entry, then locks it (option+L)
- Lists each entry as "#1 (2 items / Path)" — the order you locked in, the count, and the type of the first object
- Selecting an entry marks its extent with a red frame in the document
- Unlock releases one entry at a time
- Unlock All releases everything in the history
- Remove from List and Remove All from List discard the record while leaving the locks in place
- The record lives in the document, so it survives reopening the file
- Deleted objects drop out at the next read
- Every button and the list carry a tooltip describing what they do
- Japanese / English UI

### How to use

1. Open a document.
2. Run `LockHistoryPalette.jsx`; the palette appears.
3. Select the objects you want to lock and press Lock in the palette. That batch becomes one history entry.
4. Select the entry you want back and press Unlock.
5. If you lose track of where an entry is, select it in the list; a red frame marks its whole extent.

### Buttons

| Button | What it does |
| --- | --- |
| Lock | Records the selected objects as one history entry, then locks them |
| Unlock | Unlocks the objects of the selected entry and deletes its record |
| Unlock All | Unlocks the objects of every entry and deletes all records |
| Remove from List | Deletes the selected entry's record but leaves the objects locked |
| Remove All from List | Deletes every record in the history but leaves the objects locked |

### Notes

- **Objects locked with `command`+`2` are not recorded.** Use the palette's Lock button for anything you want recorded. Illustrator gives scripts no notification when something is locked, and inferring it afterwards from the lock state misses cases, so the locking itself happens in the palette.
- **Select > Save Selection cannot be run from a script.** It is not exposed as a menu command, and there is no API for reading or writing saved selection sets. `PageItem.tags` provides the same "stored in the document" behaviour instead.
- **The records are re-read when you click the palette.** Illustrator has no resident timer (`app.scheduleTask` belongs to Adobe Bridge, and `$.setTimeout` is missing too), so the reading happens when the palette becomes active and right after each action. Editing the document means focusing the document window, so the list is current again the moment you come back.
- **Illustrator cannot select locked objects.** Selecting an entry therefore marks the spot with a red frame instead of selecting. The items are briefly unlocked so the bounds can be measured, then restored. The frame lives on a temporary top layer called `LockHistoryPalette Preview` and is removed, layer and all, when you return to the palette.
- **option+L only works while the palette has focus.** ScriptUI key events reach a palette only when it is frontmost and active, so pressing it with the artboard focused does nothing; click the palette first. For a shortcut that works anywhere in Illustrator, assign one to the `File > Scripts` menu entry under System Settings > Keyboard > Keyboard Shortcuts > App Shortcuts on macOS.
- **The tags stay in the document.** If you hand off the file without clearing the records, the tags travel with it. Use Remove All from List or Unlock All to remove them.
- **DOM work is delegated to the main engine.** A persistent palette's engine (`#targetengine`) cannot reach `app.activeDocument`, so locking, reading and writing tags, and drawing the frame are all handed to the main engine through BridgeTalk. Replies arrive asynchronously, so the list updates a moment after you press a button.
- **This is a persistent script.** It uses `#targetengine`, so after editing the source you must **close the palette before running it again** — otherwise the old code keeps running.

### Article

[Keep a history of what you locked in Illustrator and bring it back (Japanese)](https://note.com/dtp_tranist/n/n577d8a654ec1)

### Change log

- v1.0.0 (2026-09-23): First release

### Script info

- Version: v1.0.0
- First release: 2026-09-23
- Last updated: 2026-09-23
