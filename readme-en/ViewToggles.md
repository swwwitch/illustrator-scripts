# Switch view and snapping options on or off in one dialog

[![Direct](https://img.shields.io/badge/Direct%20Link-ViewToggles.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/preference/ViewToggles.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ViewToggles.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Switches the application frame, Control panel, guides, grid, snapping and more on or off from a single dialog.

Menu commands such as View > Show ... simply flip the current state, so you have to check it first or risk switching the wrong way. This script reads the current state before the dialog opens, so everything ends up exactly as checked.

### Features

- **Basic UI**: brightness (Dark / Medium Dark / Medium Light / Light), Application Frame, Application Bar, Contextual Task Bar, Control panel, Toolbar, Help Bar, Generative Menu
- **Objects**: Corner Widget, Edges, Bounding Box
- **Artboards**: Show Borders, Show Artboard Names, Show Generative AI Buttons
- **Guides**: Show Guides, Show Smart Guides, Lock, Snap (Snap to Point)
- **Grid**: Show, Snap
- **Other Snapping**: Pixel, Glyph
- **Save and Export**: Export in Background, Save in Background, Autosave Recovery Data
- **Tools**: Shape Builder Tool Color (Pick Color From in the tool options: Color Swatches / Artwork)
- **Presets**: save the current checkboxes and brightness under a name and recall them from the dropdown. Presets survive a restart

### Usage

1. Run the script. It reads the current states and opens the dialog
2. Change the checkboxes or radio buttons you want to switch
3. Click OK. Only the items you changed are switched

### Notes

- The helper app `/Applications/SetAiMenuState.app` reads and clicks the menu checkmarks. Its source is `helpers/SetAiMenuState.applescript` in ai-scripts; build it with `osacompile -o /Applications/SetAiMenuState.app helpers/SetAiMenuState.applescript`
- The helper needs Accessibility permission (System Settings > Privacy & Security > Accessibility). Rebuilding it may require granting the permission again
- Without the helper, only the items readable from preference keys can be switched; the others are greyed out
- Some items (Toolbar, Help Bar, brightness, Save and Export, Shape Builder, etc.) are switched by the helper through the menus, Preferences or tool options after the script ends. Leave Illustrator alone until it finishes
- Shape Builder Tool Color can be switched only while a document is open. Switching it makes the Shape Builder tool the active tool
- Brightness, Save and Export and the Shape Builder option work with the Japanese UI only. The English menu names are unverified
- macOS only

### Update History

- v1.1.0 (2026-10-09) Added the Tools panel to switch Pick Color From (Color Swatches / Artwork) of the Shape Builder tool
- v1.0.0 (2026-10-09) Initial release
