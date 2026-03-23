# Changelog — Tab Bar

All notable changes to this project are documented in this file.

---

## [0.3.5] — 2026-03-23

### Changed
- **"Enable Tab Bar" button renamed to "Refresh Tab Bar"** — the button bootstraps the extension on first load and re-triggers initialisation if tabs fall out of sync; the new name reflects this more accurately.

### Fixed
- **Clean uninstall** — no orphaned toolbar bar is left behind after removing the extension and restarting LibreOffice. Three complementary mechanisms now ensure cleanup:
  - `XModifyListener` on the extension manager fires when any extension is removed; if the extension is no longer installed (detected via `XPackageInformationProvider`), toolbar settings are removed from the Writer module config and flushed to disk immediately.
  - `XTerminateListener.notifyTermination` removes toolbar settings when LibreOffice closes normally.
  - `XLayoutManager.destroyElement` is called on every open frame to clear the LayoutManager's persisted visibility state, which is stored separately from the toolbar definition and was the root cause of the residual grey bar.
- **Extension identifier** changed from `com.zdedw.tabbar` (contained a machine-specific username) to `com.github.tabbar`, consistent with `com.github.clipboardpane`.

---

## [0.3.1] — 2026-03-22

### Added
- **"Enable Tab Bar" button icon** — the toolbar button now displays a small tab-bar icon (three tabs above a document body) at 16 px and 26 px (HiDPI), generated at build time with no external dependencies. Previously the button was text-only with an empty `ImageIdentifier`.

---

## [0.3.0] — 2026-03-22

### Added
- **Cross-platform config directory support** — config and log files are now stored in the appropriate system directory on each platform:
  - Windows: `%APPDATA%\LibreOffice\`
  - macOS: `~/Library/Application Support/LibreOffice/`
  - Linux: `$XDG_CONFIG_HOME/libreoffice/` (default: `~/.config/libreoffice/`)
- `_get_config_dir()` function handles all three platforms via `sys.platform`

### Changed
- Debug logging is now **off by default**. Set the environment variable `TABBAR_DEBUG=1` before launching LibreOffice to enable it. Log file is written to the platform config directory.
- `_SETS_FILE` now resolves through `_CONFIG_DIR` rather than a hardcoded Windows `%APPDATA%` path.

---

## [0.2.9] — 2026-03-21

### Added
- **Modified state detection in poll** — the 1-second background poll now checks `model.isModified()` against `_rendered_modified` and triggers a toolbar rebuild when the unsaved state changes, keeping the `*` suffix in sync even without a user action.
- `_rendered_modified` dict tracks the last-rendered modified state per frame.

### Fixed
- Internal storage keys (`__config__`, `__last_session__`) no longer appear as selectable sets in the Sets menu. `_load_sets()` now filters any key beginning with `__`.

---

## [0.2.8] — 2026-03-19

### Added
- **Tab status indicator** — unsaved (modified) documents are marked with a `*` suffix in the tab label.
- **Close All Others** in the ▾ tab menu — closes every document except the current one.
- **Close All** in the ▾ tab menu — closes all open documents.
- **Update a Set** in the ☰ Sets menu — overwrites an existing named set with the current open documents.
- **Restore Last Session** in the ☰ Sets menu — reopens the documents that were open when LibreOffice was last closed.
- **Auto-session save on close** — `_save_last_session()` is called whenever any document is closed, keeping the last session record current.
- **Keyboard tab switching** — Ctrl+Tab / Ctrl+Shift+Tab cycles through open documents.
- **Tab Key Switching toggle** in the ☰ Sets menu — enables or disables keyboard cycling; off by default to avoid conflicting with system or desktop shortcuts. Setting is persisted to config.

---

## [0.2.7] — 2026-03-18

### Added
- **Active tab visual highlight** — the currently focused document tab is prefixed with `●` (U+25CF). Inactive tabs are plain text.
- `TabWindowFocusListener` — registers `XFocusListener` on each frame's container window; `focusGained` updates `_active_frame_id` and rebuilds the toolbar.
- `_active_frame_id` module global tracks the `id()` of the frame holding OS focus; updated on tab click, window focus, and startup scan.

### Changed
- `addStatusListener` in `TabDispatch` sets `ev.State = uno.Any("boolean", is_active)` for semantic correctness alongside the label-based visual.

---

## [0.2.6] — 2026-03-17

### Added
- **Saved Tab Sets** (`☰ Sets` button, permanently anchored at the right end of the toolbar):
  - *Save Current Set…* — prompts for a name and saves all open document URLs as a named set.
  - *Open a set* — click any named set to add its documents to the current session (existing documents are not closed).
  - *Rename a Set…* — rename an existing set via a listbox picker.
  - *Delete a Set…* — remove a set from storage.
- Sets are stored as JSON in `%APPDATA%\LibreOffice\tabbar_sets.json` (Windows; see v0.3.0 for cross-platform path).
- Separator between tab buttons and the Sets button.

---

## [0.2.5] — 2026-03-16

### Added
- **Tab reordering** — Move Left and Move Right in the ▾ menu swap the tab's position in `_frames[]` and immediately rebuild the toolbar. The relevant direction is disabled at the leftmost or rightmost position.

### Changed
- Close button removed from its former position as a separate toolbar button per tab; it now lives exclusively in the ▾ dropdown menu alongside Rename, Save, and Save As.

---

## [0.2.4] — 2026-03-15

### Added
- **New Document** option in the ▾ tab menu — opens a new blank Writer document via `Desktop.loadComponentFromURL`.

---

## [0.2.3] — 2026-03-14

### Added
- **Auto-updating tab names** — a 1-second background poll (`threading.Timer`) calls `_check_title_changes()`, which compares `frame.Title` against the last-rendered title in `_rendered_titles`. If any tab has drifted, the toolbar is rebuilt. Catches title changes from Rename, Save As, and initial save of an untitled document.

---

## [0.2.2] — 2026-03-13

### Fixed
- **Rename now renames the actual file on disk.** Previously, Rename only updated the in-memory tab label. Now it calls `model.storeAsURL()` (which also updates the frame's URL and title) and removes the old file with `os.remove()`. For unsaved (untitled) documents, Rename falls through to Save As.
- **Save and Save As now work.** Both were dispatching commands that silently did nothing because the target frame was not active. Fixed by calling `frame.activate()` before dispatch. Save switched to `model.store()` called directly on the model object.

---

## [0.2.1] — 2026-03-12

### Added
- **Per-tab ▾ dropdown menu** — each tab gains a small `▾` button that opens a popup menu (blocking modal, via `XPopupMenu.execute()`). Initial menu items: Rename, Save, Save As, Close.
- Two-button-per-tab layout: `[Title (.uno:TabBar.Switch.N)]` + `[▾ (tabbar:menu.N)]`
- `tabbar:menu.N` and `tabbar:close.N` dispatch URLs registered in `ProtocolHandler.xcu`.
- Right-click menus implemented via a dedicated second button, since `XMouseListener` does not receive right-click events on VCL toolbar buttons.

---

## [0.2.0] — 2026-03-11

### Changed
- First public-facing iteration after initial private development. Toolbar creation path stabilised; `uno.invoke()` with `uno.Any("[]com.sun.star.beans.PropertyValue", item)` confirmed as the only reliable way to insert toolbar items (plain Python tuple raises `IllegalArgumentException`).
- 1-second polling timer used for new-document detection (all listener-based alternatives proved unreliable in the Python UNO bridge).

---

## [0.1.0] — 2026-03-10

### Added
- Initial release.
- Displays all open LibreOffice Writer documents as clickable tabs in a custom toolbar (`private:resource/toolbar/custom_tabtoolbar`).
- Clicking a tab activates that frame via `frame.activate()` and `getContainerWindow().setFocus()`.
- Startup triggered by `Jobs.xcu` (`onFirstVisibleTask` event) and by the "Enable Tab Bar" button (`Addons.xcu` → `tabbar:init`).
- `XDispatchProviderInterceptor` (`TabInterceptor`) installed per frame to handle `.uno:TabBar.Switch.N` dispatch commands.
- `XProtocolHandler` (`TabBarProtocolHandler`) handles all `tabbar:*` URLs.
- `TabFrameActionListener.disposing()` detects document close and removes the tab.
