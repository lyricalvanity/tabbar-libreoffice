# -*- coding: utf-8 -*-
"""
TabBar – LibreOffice Writer Document Tab Bar  v0.3.5

Shows all open Writer documents as clickable tabs in a toolbar.
Each tab has two buttons:  [Title] [▾ menu]
The ▾ menu offers: Rename, Save, Save As, Move Left, Move Right, New Document,
                   Close, Close All Others, Close All.
A permanent [☰ Sets] button at the toolbar's right edge manages saved tab sets.

Startup chain
-------------
1. Jobs.xcu fires onFirstVisibleTask at LO startup → bootstraps listeners.
2. Addons.xcu adds a "Tab Bar" toolbar with an "Enable Tab Bar" button that
   dispatches tabbar:init.  ProtocolHandler.xcu routes tabbar:* to
   TabBarProtocolHandler.  Either path calls _bootstrap().

Auto-update
-----------
Each tracked Writer frame gets:
  • TabFrameActionListener  – disposing() fires on close → purge dead tabs
  • TabWindowFocusListener  – focusLost() fires when user switches away
                              → scan desktop for newly opened frames
  • ▾ dropdown button per tab – left-click shows context menu (tabbar:menu.N)
  • 1-second polling timer  – catches new documents that slip past the above

Platform support
----------------
Windows : config in %APPDATA%\LibreOffice\
macOS   : config in ~/Library/Application Support/LibreOffice/
Linux   : config in $XDG_CONFIG_HOME/libreoffice/  (default ~/.config/libreoffice/)

Debug logging (off by default): set env var TABBAR_DEBUG=1 before launching LO.
Log file: <config dir>/tab_bar.log
"""

# ── Standard library first – these never fail ────────────────────────────────
import os
import json
import sys
import traceback
import threading

# ── Log helper defined BEFORE any UNO imports so import errors are captured ──


def _get_config_dir():
    if sys.platform == "win32":
        base = os.environ.get("APPDATA", os.path.expanduser("~"))
        return os.path.join(base, "LibreOffice")
    elif sys.platform == "darwin":
        return os.path.join(os.path.expanduser("~"),
                            "Library", "Application Support", "LibreOffice")
    else:  # Linux / other Unix — XDG standard
        xdg = os.environ.get("XDG_CONFIG_HOME",
                              os.path.join(os.path.expanduser("~"), ".config"))
        return os.path.join(xdg, "libreoffice")


_CONFIG_DIR = _get_config_dir()
_DEBUG      = os.environ.get("TABBAR_DEBUG", "").lower() in ("1", "true", "yes")
_LOG        = os.path.join(_CONFIG_DIR, "tab_bar.log")


def _log(msg):
    if not _DEBUG:
        return
    try:
        os.makedirs(_CONFIG_DIR, exist_ok=True)
        with open(_LOG, "a", encoding="utf-8") as f:
            f.write(msg + "\n")
    except Exception:
        pass


_log("tab_bar: file executing")   # proves the Python file is loaded at all

# ── UNO imports in a guarded block so failures are diagnosed ─────────────────

try:
    import uno
    import unohelper

    from com.sun.star.lang     import XServiceInfo, XEventListener, XInitialization
    from com.sun.star.task     import XJob
    from com.sun.star.frame    import (XDispatchProviderInterceptor, XDispatch,
                                       XDispatchProvider, XFrameActionListener,
                                       XTerminateListener)
    from com.sun.star.document import XDocumentEventListener
    from com.sun.star.awt      import XFocusListener, XKeyHandler
    from com.sun.star.awt      import Key as _Key, KeyModifier as _KeyMod
    from com.sun.star.util     import XModifyListener

    _log("tab_bar: UNO imports OK")

except Exception:
    _log("tab_bar: UNO import FAILED\n" + traceback.format_exc())
    raise   # let LO know the module is broken


# ──────────────────────────────────────────────────────────────────────────────
# Constants
# ──────────────────────────────────────────────────────────────────────────────

JOB_IMPL     = "com.github.tabbar.Job"
JOB_SVC      = "com.sun.star.task.Job"

HANDLER_IMPL = "com.github.tabbar.ProtocolHandler"
HANDLER_SVC  = "com.sun.star.frame.ProtocolHandler"

TOOLBAR_URL  = "private:resource/toolbar/custom_tabtoolbar"
CMD_PREFIX   = ".uno:TabBar.Switch."

WRITER_SVC   = "com.sun.star.text.TextDocument"
CMD_MENU     = "tabbar:menu."
CMD_CLOSE    = "tabbar:close."
CMD_SETS     = "tabbar:sets"
MAX_LABEL    = 30

_SETS_FILE   = os.path.join(_CONFIG_DIR, "tabbar_sets.json")
_LO_SUFFIXES = (
    " \u2013 LibreOffice Writer",   # en-dash (U+2013) – some LO/OS combos use this
    " - LibreOffice Writer",         # hyphen-minus
    " \u2014 LibreOffice Writer",   # em-dash
    " \u2013 LibreOffice",
    " - LibreOffice",
    " \u2014 LibreOffice",
)


# ──────────────────────────────────────────────────────────────────────────────
# Module-level state  (survives across job / handler instantiations)
# ──────────────────────────────────────────────────────────────────────────────

_frames              = []     # [XFrame, …] open Writer docs, in order
_interceptors        = {}     # id(frame) -> TabInterceptor
_frame_listeners     = {}     # id(frame) -> TabFrameActionListener (GC guard)
_focus_listeners     = {}     # id(frame) -> TabWindowFocusListener  (GC guard)
_custom_labels       = {}     # id(frame) -> user-set tab label (overrides frame.Title)
_rendered_titles     = {}     # id(frame) -> last label written to toolbar (change detection)
_active_frame_id     = None   # id(frame) of the frame that currently has focus
_rendered_modified   = {}     # id(frame) -> bool, last rendered modified state
_kb_tab_switch       = False  # Ctrl+Tab cycling enabled (user toggle)
_key_handlers        = {}     # id(frame) -> TabKeyHandler (GC guard)
_bootstrapped        = False  # have we registered global listeners yet?
_global_listener     = None   # GlobalEventBroadcaster listener (GC guard)
_desktop_listener    = None   # Desktop XFrameActionListener      (GC guard)
_terminate_listener  = None   # Desktop XTerminateListener        (GC guard)
_ext_modify_listener = None   # ExtensionManager XModifyListener  (GC guard)
_poll_timer          = None   # threading.Timer for periodic frame scan


# ──────────────────────────────────────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────────────────────────────────────

def _pv(name, value):
    p = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
    p.Name  = name
    p.Value = value
    return p


def _make_tab_items(index, title, is_active=False, is_modified=False):
    """Return TWO toolbar items per tab: [Title] [▾].

    Active tab is prefixed with ●.  Modified (unsaved) tab has a * suffix.
    Both are theme-independent visual indicators that work in plain text labels.
    """
    # Apply modified marker before truncation so it's always visible
    display = title + (" *" if is_modified else "")
    # Budget for the "● " active prefix (2 chars)
    budget  = MAX_LABEL - (2 if is_active else 0)
    if len(display) > budget:
        display = display[:budget - 1] + "\u2026"
    label = ("\u25cf " + display) if is_active else display   # ● BLACK CIRCLE
    switch_btn = (
        _pv("CommandURL", f"{CMD_PREFIX}{index}"),
        _pv("Label",      label),
        _pv("Type",       0),
        _pv("Style",      1),
        _pv("IsVisible",  True),
    )
    menu_btn = (
        _pv("CommandURL", f"tabbar:menu.{index}"),
        _pv("Label",      "\u25be"),   # ▾ small down-pointing triangle
        _pv("Type",       0),
        _pv("Style",      1),
        _pv("IsVisible",  True),
    )
    return switch_btn, menu_btn


def _strip_suffix(title):
    for s in _LO_SUFFIXES:
        if title.endswith(s):
            return title[:-len(s)]
    return title


def _sync_active_frame(ctx):
    """Ask the Desktop which frame is current and record its id."""
    global _active_frame_id
    try:
        desktop = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", ctx)
        current = desktop.getCurrentFrame()
        if current is not None:
            for frame in _frames:
                if frame == current:
                    _active_frame_id = id(frame)
                    return
    except Exception:
        pass


def _get_writer_cfg(ctx):
    try:
        sup = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.ui.ModuleUIConfigurationManagerSupplier", ctx)
        return sup.getUIConfigurationManager(WRITER_SVC)
    except Exception:
        _log("_get_writer_cfg failed:\n" + traceback.format_exc())
        return None


def _is_writer_frame(frame):
    try:
        ctrl = frame.getController()
        if ctrl is None:
            return False
        model = ctrl.getModel()
        return model is not None and model.supportsService(WRITER_SVC)
    except Exception:
        return False



# ──────────────────────────────────────────────────────────────────────────────
# Tab-sets  (saved sessions stored as JSON in %APPDATA%\LibreOffice\)
# ──────────────────────────────────────────────────────────────────────────────

def _read_raw_file():
    """Return the full JSON dict from disk, or {} on any error."""
    try:
        if os.path.exists(_SETS_FILE):
            with open(_SETS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        _log("_read_raw_file failed:\n" + traceback.format_exc())
    return {}


def _write_raw_file(data):
    """Write the full JSON dict to disk."""
    try:
        os.makedirs(os.path.dirname(_SETS_FILE), exist_ok=True)
        with open(_SETS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    except Exception:
        _log("_write_raw_file failed:\n" + traceback.format_exc())


def _load_sets():
    """Return only user-named sets (excludes internal __ keys)."""
    return {k: v for k, v in _read_raw_file().items()
            if not k.startswith("__")}


def _save_sets(sets):
    """Persist user-named sets, leaving internal __ keys untouched."""
    data = _read_raw_file()
    # Remove old user sets, keep internal keys
    for k in [k for k in data if not k.startswith("__")]:
        del data[k]
    data.update(sets)
    _write_raw_file(data)


def _load_config():
    """Return the __config__ dict."""
    return _read_raw_file().get("__config__", {})


def _save_config(config):
    """Write __config__ back without disturbing anything else."""
    data = _read_raw_file()
    data["__config__"] = config
    _write_raw_file(data)


def _save_last_session():
    """Persist the currently open saved documents as __last_session__."""
    data = _read_raw_file()
    data["__last_session__"] = _current_saved_urls()
    _write_raw_file(data)


def _open_last_session(ctx):
    """Open documents from __last_session__ that aren't already open."""
    urls = _read_raw_file().get("__last_session__", [])
    if not urls:
        return
    already = _open_urls()
    smgr    = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    for url in urls:
        if url in already:
            continue
        try:
            local = uno.fileUrlToSystemPath(url)
            if os.path.exists(local):
                desktop.loadComponentFromURL(url, "_blank", 0, ())
        except Exception:
            _log(f"_open_last_session: failed to open {url!r}:\n"
                 + traceback.format_exc())


# ── Small UNO dialogs ─────────────────────────────────────────────────────────

def _show_message(ctx, win, title, message):
    """Non-blocking informational dialog (OK button only)."""
    try:
        smgr    = ctx.ServiceManager
        toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)

        dm = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
        dm.Title    = title
        dm.Width    = 230
        dm.Height   = 72
        dm.Moveable = True

        lbl = dm.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
        lbl.Label     = message
        lbl.PositionX = 8
        lbl.PositionY = 8
        lbl.Width     = 214
        lbl.Height    = 36
        lbl.MultiLine = True
        dm.insertByName("lbl", lbl)

        btn = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
        btn.Label          = "OK"
        btn.PushButtonType = 1
        btn.DefaultButton  = True
        btn.PositionX      = 140
        btn.PositionY      = 50
        btn.Width          = 45
        btn.Height         = 14
        dm.insertByName("btn", btn)

        dlg = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
        dlg.setModel(dm)
        dlg.createPeer(toolkit, win)
        dlg.execute()
        dlg.dispose()
    except Exception:
        _log("_show_message failed:\n" + traceback.format_exc())


def _pick_from_list(ctx, win, title, message, options):
    """Show a listbox dialog. Returns the chosen string or None on Cancel."""
    try:
        smgr    = ctx.ServiceManager
        toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)

        dm = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
        dm.Title    = title
        dm.Width    = 220
        dm.Height   = 122
        dm.Moveable = True

        lbl = dm.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
        lbl.Label     = message
        lbl.PositionX = 8
        lbl.PositionY = 8
        lbl.Width     = 204
        lbl.Height    = 14
        dm.insertByName("lbl", lbl)

        lst = dm.createInstance("com.sun.star.awt.UnoControlListBoxModel")
        lst.PositionX      = 8
        lst.PositionY      = 25
        lst.Width          = 204
        lst.Height         = 58
        lst.MultiSelection = False
        dm.insertByName("lst", lst)

        btn_ok = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
        btn_ok.Label          = "OK"
        btn_ok.PushButtonType = 1
        btn_ok.DefaultButton  = True
        btn_ok.PositionX      = 115
        btn_ok.PositionY      = 97
        btn_ok.Width          = 45
        btn_ok.Height         = 14
        dm.insertByName("btn_ok", btn_ok)

        btn_cancel = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
        btn_cancel.Label          = "Cancel"
        btn_cancel.PushButtonType = 2
        btn_cancel.PositionX      = 167
        btn_cancel.PositionY      = 97
        btn_cancel.Width          = 45
        btn_cancel.Height         = 14
        dm.insertByName("btn_cancel", btn_cancel)

        dlg = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
        dlg.setModel(dm)
        dlg.createPeer(toolkit, win)

        lst_ctrl = dlg.getControl("lst")
        for opt in options:
            lst_ctrl.addItem(opt, lst_ctrl.getItemCount())
        if options:
            lst_ctrl.selectItemPos(0, True)

        result = None
        if dlg.execute() == 1:
            pos = lst_ctrl.getSelectedItemPos()
            if 0 <= pos < len(options):
                result = options[pos]
        dlg.dispose()
        return result
    except Exception:
        _log("_pick_from_list failed:\n" + traceback.format_exc())
        return None


# ── Tab-set operations ────────────────────────────────────────────────────────

def _current_saved_urls():
    """Return list of file:// URLs for every open, saved Writer document."""
    urls = []
    for frame in _frames:
        try:
            model = frame.getController().getModel()
            url   = model.getURL() if model else ""
            if url and not url.startswith("private:"):
                urls.append(url)
        except Exception:
            pass
    return urls


def _open_urls():
    """Return set of file:// URLs currently open."""
    return set(_current_saved_urls())


def _save_current_set(ctx, win):
    """Ask for a name, then persist the current open documents as a tab set."""
    urls = _current_saved_urls()
    if not urls:
        _show_message(ctx, win, "Save Tab Set",
                      "No saved documents are open.\n"
                      "Save your documents before creating a tab set.")
        return

    name = _get_input(ctx, win, "Save Tab Set", "Tab set name:", "")
    if not name or not name.strip():
        return
    name = name.strip()

    sets = _load_sets()
    if name in sets:
        # Overwrite silently — user typed the same name intentionally
        pass
    sets[name] = urls
    _save_sets(sets)
    _log(f"saved tab set {name!r}: {len(urls)} URL(s)")


def _open_set(ctx, set_name):
    """Open all documents in the named set that aren't already open."""
    sets = _load_sets()
    urls = sets.get(set_name, [])
    already_open = _open_urls()

    smgr    = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    missing = []
    for url in urls:
        if url in already_open:
            continue
        try:
            local = uno.fileUrlToSystemPath(url)
            if not os.path.exists(local):
                missing.append(os.path.basename(local))
                continue
            desktop.loadComponentFromURL(url, "_blank", 0, ())
        except Exception:
            _log(f"_open_set: failed to open {url!r}:\n" + traceback.format_exc())
            missing.append(url)

    if missing:
        _log(f"_open_set {set_name!r}: {len(missing)} file(s) not found: {missing}")


def _rename_set_dialog(ctx, win):
    """Pick a saved set by name, then rename it."""
    sets  = _load_sets()
    names = sorted(sets)
    if not names:
        _show_message(ctx, win, "Rename Set", "No saved tab sets.")
        return

    old = _pick_from_list(ctx, win, "Rename Set", "Choose set to rename:", names)
    if old is None:
        return

    new = _get_input(ctx, win, "Rename Set", f"New name for \u2018{old}\u2019:", old)
    if not new or not new.strip() or new.strip() == old:
        return
    new = new.strip()

    sets[new] = sets.pop(old)
    _save_sets(sets)
    _log(f"tab set renamed {old!r} → {new!r}")


def _update_set_dialog(ctx, win):
    """Overwrite an existing set with the currently open documents."""
    sets  = _load_sets()
    names = sorted(sets)
    if not names:
        _show_message(ctx, win, "Update Set", "No saved tab sets to update.")
        return
    name = _pick_from_list(ctx, win, "Update Set",
                           "Choose set to overwrite with current tabs:", names)
    if name is None:
        return
    urls = _current_saved_urls()
    if not urls:
        _show_message(ctx, win, "Update Set",
                      "No saved documents are open to update the set with.")
        return
    sets[name] = urls
    _save_sets(sets)
    _log(f"updated tab set {name!r}: {len(urls)} URL(s)")


def _delete_set_dialog(ctx, win):
    """Pick a saved set by name, then delete it."""
    sets  = _load_sets()
    names = sorted(sets)
    if not names:
        _show_message(ctx, win, "Delete Set", "No saved tab sets.")
        return

    name = _pick_from_list(ctx, win, "Delete Set", "Choose set to delete:", names)
    if name is None:
        return

    sets.pop(name, None)
    _save_sets(sets)
    _log(f"tab set deleted {name!r}")


def _show_sets_menu(ctx, frame, win, click_x, click_y):
    """Build and execute the ☰ Sets popup menu.

    Static IDs:
      1  Save Current Set…
      4  Update a Set…
      2  Rename a Set…
      3  Delete a Set…
      5  Restore Last Session
      6  Tab Key Switching  (toggle, checkmarked when on)
    100+ Named set items (open on click)
    """
    try:
        smgr  = ctx.ServiceManager
        popup = smgr.createInstanceWithContext("com.sun.star.awt.PopupMenu", ctx)

        sets      = _load_sets()
        names     = sorted(sets)
        raw       = _read_raw_file()
        has_last  = bool(raw.get("__last_session__"))

        # ── Top: save / update ───────────────────────────────────────────────
        popup.insertItem(1, "Save Current Set\u2026", 0, 0)
        popup.insertItem(4, "Update a Set\u2026",     0, 1)
        if not names:
            popup.enableItem(4, False)

        # ── Middle: named sets ───────────────────────────────────────────────
        if names:
            popup.insertSeparator(2)
            for i, name in enumerate(names):
                popup.insertItem(100 + i, name, 0, 3 + i)
            base = 3 + len(names)
            popup.insertSeparator(base)
            popup.insertItem(2, "Rename a Set\u2026", 0, base + 1)
            popup.insertItem(3, "Delete a Set\u2026", 0, base + 2)
            foot = base + 3
        else:
            foot = 2

        # ── Bottom: session + keyboard toggle ────────────────────────────────
        popup.insertSeparator(foot)
        popup.insertItem(5, "Restore Last Session", 0, foot + 1)
        if not has_last:
            popup.enableItem(5, False)

        kb_label = "Tab Key Switching"
        popup.insertItem(6, kb_label, 0, foot + 2)
        popup.checkItem(6, _kb_tab_switch)

        rect        = uno.createUnoStruct("com.sun.star.awt.Rectangle")
        rect.X      = click_x
        rect.Y      = click_y
        rect.Width  = 1
        rect.Height = 1

        selected = popup.execute(win, rect, 0)
        _log(f"sets menu: selected={selected}")

        if selected == 1:
            _save_current_set(ctx, win)
        elif selected == 4:
            _update_set_dialog(ctx, win)
        elif selected == 2:
            _rename_set_dialog(ctx, win)
        elif selected == 3:
            _delete_set_dialog(ctx, win)
        elif selected == 5:
            _open_last_session(ctx)
        elif selected == 6:
            _toggle_kb_tab_switch(ctx)
        elif 100 <= selected < 100 + len(names):
            _open_set(ctx, names[selected - 100])

    except Exception:
        _log("_show_sets_menu failed:\n" + traceback.format_exc())


# ──────────────────────────────────────────────────────────────────────────────
# Toolbar management
# ──────────────────────────────────────────────────────────────────────────────

def _rebuild_toolbar(ctx):
    cfg = _get_writer_cfg(ctx)
    if cfg is None:
        return

    items = []
    for i, frame in enumerate(_frames):
        try:
            # Use custom label if set, otherwise strip the LO suffix from title
            fid = id(frame)
            if fid in _custom_labels:
                title = _custom_labels[fid]
            else:
                raw = frame.Title or f"Document {i + 1}"
                title = _strip_suffix(raw)
                _log(f"  tab {i}: raw={raw!r}  stripped={title!r}")
            _rendered_titles[fid] = title   # record for change detection
        except Exception:
            title = f"Document {i + 1}"
        try:
            model       = frame.getController().getModel()
            is_modified = bool(model and model.isModified())
        except Exception:
            is_modified = False
        _rendered_modified[fid] = is_modified
        is_active = (fid == _active_frame_id)
        items.extend(_make_tab_items(i, title, is_active, is_modified))

    # Permanent ☰ Sets button — always the last item in the toolbar
    if items:
        items.append((
            _pv("CommandURL", ""),
            _pv("Type",       1),        # SEPARATOR_LINE
            _pv("IsVisible",  True),
        ))
    items.append((
        _pv("CommandURL", CMD_SETS),
        _pv("Label",      "\u2630 Sets"),
        _pv("Type",       0),
        _pv("Style",      1),
        _pv("IsVisible",  True),
    ))

    try:
        # insertByIndex(long, any) – must pass uno.Any in the argTuple so
        # pyuno.invoke applies the explicit SEQUENCE type (TypeClass 20).
        def _insert(container, i, item):
            uno.invoke(
                container, "insertByIndex",
                (i, uno.Any("[]com.sun.star.beans.PropertyValue", item)))

        if cfg.hasSettings(TOOLBAR_URL):
            container = cfg.getSettings(TOOLBAR_URL, True)
            while container.getCount() > 0:
                container.removeByIndex(0)
            for i, item in enumerate(items):
                _insert(container, i, item)
            cfg.replaceSettings(TOOLBAR_URL, container)
        else:
            container = cfg.createSettings()
            for i, item in enumerate(items):
                _insert(container, i, item)
            cfg.insertSettings(TOOLBAR_URL, container)

        _log(f"toolbar rebuilt: {len(items)} tab(s)")

    except Exception:
        _log("_rebuild_toolbar failed:\n" + traceback.format_exc())


def _show_toolbar_in_frame(ctx, frame):
    try:
        lm = frame.LayoutManager
        if not lm.isElementVisible(TOOLBAR_URL):
            lm.requestElement(TOOLBAR_URL)
        _log("toolbar shown in frame")
    except Exception:
        _log("_show_toolbar_in_frame failed:\n" + traceback.format_exc())


# ──────────────────────────────────────────────────────────────────────────────
# Context menu – right-click on a tab
# ──────────────────────────────────────────────────────────────────────────────

def _get_input(ctx, parent_win, title, message, default=""):
    """Show a simple UNO input dialog. Returns the entered string, or None on Cancel."""
    try:
        smgr = ctx.ServiceManager
        toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)

        # ── Dialog model ──────────────────────────────────────────────────────
        dlg_model = smgr.createInstanceWithContext(
            "com.sun.star.awt.UnoControlDialogModel", ctx)
        dlg_model.Title    = title
        dlg_model.Width    = 220
        dlg_model.Height   = 75
        dlg_model.Moveable = True

        # Label
        lbl = dlg_model.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
        lbl.Label     = message
        lbl.PositionX = 8
        lbl.PositionY = 8
        lbl.Width     = 204
        lbl.Height    = 14
        dlg_model.insertByName("lbl", lbl)

        # Edit field
        edt = dlg_model.createInstance("com.sun.star.awt.UnoControlEditModel")
        edt.Text      = default
        edt.PositionX = 8
        edt.PositionY = 26
        edt.Width     = 204
        edt.Height    = 14
        dlg_model.insertByName("edt", edt)

        # OK button
        btn_ok = dlg_model.createInstance("com.sun.star.awt.UnoControlButtonModel")
        btn_ok.Label         = "OK"
        btn_ok.PushButtonType = 1    # OK
        btn_ok.DefaultButton  = True
        btn_ok.PositionX      = 115
        btn_ok.PositionY      = 52
        btn_ok.Width          = 45
        btn_ok.Height         = 14
        dlg_model.insertByName("btn_ok", btn_ok)

        # Cancel button
        btn_cancel = dlg_model.createInstance("com.sun.star.awt.UnoControlButtonModel")
        btn_cancel.Label          = "Cancel"
        btn_cancel.PushButtonType = 2    # Cancel
        btn_cancel.PositionX      = 167
        btn_cancel.PositionY      = 52
        btn_cancel.Width          = 45
        btn_cancel.Height         = 14
        dlg_model.insertByName("btn_cancel", btn_cancel)

        # ── Create dialog ─────────────────────────────────────────────────────
        dlg = smgr.createInstanceWithContext(
            "com.sun.star.awt.UnoControlDialog", ctx)
        dlg.setModel(dlg_model)
        dlg.createPeer(toolkit, parent_win)

        ret = dlg.execute()   # 1 = OK, 0 = Cancel
        text = dlg.getControl("edt").getText() if ret == 1 else None
        dlg.dispose()
        return text

    except Exception:
        _log("_get_input failed:\n" + traceback.format_exc())
        return None


def _get_model(frame):
    """Return the XModel for a frame, or None."""
    try:
        ctrl = frame.getController()
        return ctrl.getModel() if ctrl else None
    except Exception:
        return None


def _dispatch_via_helper(ctx, frame, cmd):
    """Activate frame and fire a UNO command through DispatchHelper."""
    try:
        frame.activate()
        helper = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.DispatchHelper", ctx)
        helper.executeDispatch(frame, cmd, "_self", 0, ())
        _log(f"helper dispatched {cmd}")
    except Exception:
        _log(f"_dispatch_via_helper({cmd}) failed:\n" + traceback.format_exc())


def _save_document(ctx, frame):
    """Save the document. If unsaved (no URL), falls back to Save As."""
    try:
        model = _get_model(frame)
        if model is None:
            return
        url = ""
        try:
            url = model.getURL()
        except Exception:
            pass
        if url and not url.startswith("private:"):
            model.store()
            _log("document stored")
        else:
            _dispatch_via_helper(ctx, frame, ".uno:SaveAs")
    except Exception:
        _log("_save_document failed:\n" + traceback.format_exc())


def _rename_document(ctx, frame, parent_win):
    """Rename the document on disk.

    For unsaved documents, opens the Save As dialog so the user can
    name them for the first time.  For saved documents, prompts for a
    new filename (extension preserved), saves to the new path via
    storeAsURL, then deletes the old file.
    """
    try:
        model = _get_model(frame)
        if model is None:
            return

        current_url = ""
        try:
            current_url = model.getURL()
        except Exception:
            pass

        if not current_url or current_url.startswith("private:"):
            # Not saved yet – let Save As handle naming
            _dispatch_via_helper(ctx, frame, ".uno:SaveAs")
            return

        # ── Saved document: ask for a new filename ────────────────────────
        local_path = uno.fileUrlToSystemPath(current_url)
        dir_path   = os.path.dirname(local_path)
        basename   = os.path.basename(local_path)
        stem, ext  = os.path.splitext(basename)

        new_stem = _get_input(ctx, parent_win,
                              "Rename Document",
                              f"New name for \u2018{basename}\u2019:",
                              stem)
        if not new_stem or not new_stem.strip():
            return
        new_stem = new_stem.strip()

        # Preserve the extension unless the user explicitly typed one
        if not os.path.splitext(new_stem)[1]:
            new_stem += ext
        new_path = os.path.join(dir_path, new_stem)
        if new_path.lower() == local_path.lower():
            return  # nothing to do

        new_url = uno.systemPathToFileUrl(new_path)

        # Carry forward the current filter so the format is unchanged
        filter_args = ()
        try:
            for prop in model.getMediaDescriptor():
                if prop.Name == "FilterName":
                    fa = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
                    fa.Name  = "FilterName"
                    fa.Value = prop.Value
                    filter_args = (fa,)
                    break
        except Exception:
            pass

        model.storeAsURL(new_url, filter_args)
        _log(f"document renamed to {new_url!r}")

        # Remove the old file now that the document lives at the new path
        try:
            os.remove(local_path)
            _log(f"old file removed: {local_path!r}")
        except Exception:
            _log(f"could not remove old file:\n" + traceback.format_exc())

        # Drop any custom label so the tab auto-updates from the new title
        fid = id(frame)
        if fid in _custom_labels:
            del _custom_labels[fid]
        _rebuild_toolbar(ctx)

    except Exception:
        _log("_rename_document failed:\n" + traceback.format_exc())


def _close_others(ctx, keep_frame):
    """Close every tracked frame except keep_frame."""
    for f in list(_frames):
        if f != keep_frame:
            try:
                _dispatch_via_helper(ctx, f, ".uno:CloseDoc")
            except Exception:
                pass


def _close_all(ctx):
    """Close every tracked frame."""
    for f in list(_frames):
        try:
            _dispatch_via_helper(ctx, f, ".uno:CloseDoc")
        except Exception:
            pass


def _show_tab_context_menu(ctx, frame, win, click_x, click_y, tab_idx):
    """Build and execute a popup menu for a tab's ▾ button."""
    try:
        smgr  = ctx.ServiceManager
        popup = smgr.createInstanceWithContext("com.sun.star.awt.PopupMenu", ctx)

        n = len(_frames)

        # Menu layout:
        #   Rename…             (id=1)
        #   ─────────────
        #   Save                (id=2)
        #   Save As…            (id=3)
        #   ─────────────
        #   Move Left           (id=5)   disabled when already leftmost
        #   Move Right          (id=6)   disabled when already rightmost
        #   ─────────────
        #   New Document        (id=4)
        #   ─────────────
        #   Close               (id=7)
        #   Close All Others    (id=8)   disabled when only 1 tab
        #   Close All           (id=9)
        popup.insertItem(1, "Rename\u2026",        0, 0)
        popup.insertSeparator(1)
        popup.insertItem(2, "Save",                0, 2)
        popup.insertItem(3, "Save As\u2026",       0, 3)
        popup.insertSeparator(4)
        popup.insertItem(5, "Move Left",           0, 5)
        popup.insertItem(6, "Move Right",          0, 6)
        popup.insertSeparator(7)
        popup.insertItem(4, "New Document",        0, 8)
        popup.insertSeparator(9)
        popup.insertItem(7, "Close",               0, 10)
        popup.insertItem(8, "Close All Others",    0, 11)
        popup.insertItem(9, "Close All",           0, 12)

        if tab_idx == 0:
            popup.enableItem(5, False)
        if tab_idx >= n - 1:
            popup.enableItem(6, False)
        if n <= 1:
            popup.enableItem(8, False)

        rect        = uno.createUnoStruct("com.sun.star.awt.Rectangle")
        rect.X      = click_x
        rect.Y      = click_y
        rect.Width  = 1
        rect.Height = 1

        selected = popup.execute(win, rect, 0)
        _log(f"context menu: selected={selected}  tab_idx={tab_idx}")

        # Re-resolve the target frame (list may have changed during popup)
        if 0 <= tab_idx < len(_frames):
            target = _frames[tab_idx]
        else:
            return

        if selected == 1:
            _rename_document(ctx, target, win)
        elif selected == 2:
            _save_document(ctx, target)
        elif selected == 3:
            _dispatch_via_helper(ctx, target, ".uno:SaveAs")
        elif selected == 5 and tab_idx > 0:
            _frames[tab_idx], _frames[tab_idx - 1] = _frames[tab_idx - 1], _frames[tab_idx]
            _rebuild_toolbar(ctx)
        elif selected == 6 and tab_idx < len(_frames) - 1:
            _frames[tab_idx], _frames[tab_idx + 1] = _frames[tab_idx + 1], _frames[tab_idx]
            _rebuild_toolbar(ctx)
        elif selected == 4:
            try:
                desktop = smgr.createInstanceWithContext(
                    "com.sun.star.frame.Desktop", ctx)
                desktop.loadComponentFromURL(
                    "private:factory/swriter", "_blank", 0, ())
            except Exception:
                _log("New Document failed:\n" + traceback.format_exc())
        elif selected == 7:
            _dispatch_via_helper(ctx, target, ".uno:CloseDoc")
        elif selected == 8:
            _close_others(ctx, target)
        elif selected == 9:
            _close_all(ctx)

    except Exception:
        _log("_show_tab_context_menu failed:\n" + traceback.format_exc())


# ──────────────────────────────────────────────────────────────────────────────
# Frame lifecycle
# ──────────────────────────────────────────────────────────────────────────────

def _add_frame(ctx, frame):
    if frame is None:
        return
    if any(f == frame for f in _frames):
        return
    if not _is_writer_frame(frame):
        _log("frame is not Writer – skipping")
        return

    _frames.append(frame)
    _log(f"frame added, total={len(_frames)}")

    # Dispatch interceptor for tab-switch commands
    try:
        interceptor = TabInterceptor(frame)
        frame.registerDispatchProviderInterceptor(interceptor)
        _interceptors[id(frame)] = interceptor
        _log("interceptor registered")
    except Exception:
        _log("interceptor registration failed:\n" + traceback.format_exc())

    # XFrameActionListener – disposing() fires when the frame closes
    try:
        fa = TabFrameActionListener(ctx)
        frame.addFrameActionListener(fa)
        _frame_listeners[id(frame)] = fa        # prevent GC
        _log("frame-action listener registered")
    except Exception:
        _log("frame-action listener registration failed:\n" + traceback.format_exc())

    # XFocusListener on the container window – focusLost fires when the user
    # switches to another OS window (possibly a newly opened document)
    try:
        win = frame.getContainerWindow()
        if win is not None:
            fl = TabWindowFocusListener(ctx, frame)
            win.addFocusListener(fl)
            _focus_listeners[id(frame)] = fl    # prevent GC
            _log("window focus listener registered")
    except Exception:
        _log("window focus listener registration failed:\n" + traceback.format_exc())

    _rebuild_toolbar(ctx)
    _show_toolbar_in_frame(ctx, frame)
    if _kb_tab_switch:
        _register_key_handler(ctx, frame)


def _remove_frame(ctx, frame):
    if frame is None:
        return
    removed = False
    for i, f in enumerate(_frames):
        if f == frame:
            _frames.pop(i)
            removed = True
            break
    if not removed:
        return

    _log(f"frame removed, total={len(_frames)}")

    fid = id(frame)
    if fid in _interceptors:
        try:
            frame.deregisterDispatchProviderInterceptor(_interceptors[fid])
        except Exception:
            pass
        del _interceptors[fid]

    if fid in _frame_listeners:
        try:
            frame.removeFrameActionListener(_frame_listeners[fid])
        except Exception:
            pass
        del _frame_listeners[fid]

    if fid in _focus_listeners:
        try:
            frame.getContainerWindow().removeFocusListener(_focus_listeners[fid])
        except Exception:
            pass
        del _focus_listeners[fid]

    if fid in _custom_labels:
        del _custom_labels[fid]
    if fid in _rendered_titles:
        del _rendered_titles[fid]
    if fid in _rendered_modified:
        del _rendered_modified[fid]

    _unregister_key_handler(frame)
    _save_last_session()   # keep __last_session__ current after every close

    if _frames:
        _rebuild_toolbar(ctx)
    else:
        _remove_toolbar_settings(ctx)


def _clean_dead_frames(ctx):
    dead = []
    for frame in _frames:
        try:
            _ = frame.Title
        except Exception:
            dead.append(frame)
    for frame in dead:
        _remove_frame(ctx, frame)


def _scan_existing_frames(ctx):
    """Pick up Writer documents that were already open before we started."""
    try:
        _sync_active_frame(ctx)   # know which tab to highlight before rebuilding
        smgr    = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        frames  = desktop.getFrames()
        n       = frames.getCount()
        _log(f"scanning {n} existing frame(s)")
        for i in range(n):
            _add_frame(ctx, frames.getByIndex(i))
    except Exception:
        _log("_scan_existing_frames failed:\n" + traceback.format_exc())


def _check_title_changes(ctx):
    """Rebuild the toolbar if any frame's title or modified state has changed.

    Catches title updates from Save / Save As / Rename, and the unsaved-changes
    indicator (*) appearing or disappearing as the user edits.
    """
    try:
        changed = False
        for frame in _frames:
            fid = id(frame)
            # Title check (skip custom-labelled frames — their label is stable)
            if fid not in _custom_labels:
                try:
                    current = _strip_suffix(frame.Title or "")
                    if _rendered_titles.get(fid) != current:
                        _log(f"title drift frame {fid}: "
                             f"{_rendered_titles.get(fid)!r} → {current!r}")
                        changed = True
                        break
                except Exception:
                    pass
            # Modified-state check (applies to all frames, custom label or not)
            try:
                model   = frame.getController().getModel()
                is_mod  = bool(model and model.isModified())
                if _rendered_modified.get(fid) != is_mod:
                    changed = True
                    break
            except Exception:
                pass
        if changed:
            _rebuild_toolbar(ctx)
    except Exception:
        pass  # never raise from the poll thread


def _start_poll(ctx):
    """
    Start a 1-second repeating timer that:
      • picks up newly opened Writer documents
      • removes closed-document tabs
      • detects title changes from Save / Save As / Rename

    Every event-based mechanism tried (XFrameActionListener.frameAction,
    XDocumentEventListener.notifyDocumentEvent, XFocusListener.focusLost)
    is not reliably called by LO for new document windows.  Polling is the
    only mechanism that catches all cases.
    """
    global _poll_timer
    if _poll_timer is not None:
        return

    def _poll():
        global _poll_timer
        try:
            _scan_existing_frames(ctx)
            _clean_dead_frames(ctx)
            _check_title_changes(ctx)
        except Exception:
            pass        # never let the timer thread die from an exception
        finally:
            # reschedule unconditionally; daemon=True means it won't block LO exit
            _poll_timer = threading.Timer(1.0, _poll)
            _poll_timer.daemon = True
            _poll_timer.start()

    _poll_timer = threading.Timer(1.0, _poll)
    _poll_timer.daemon = True
    _poll_timer.start()
    _log("poll timer started (1 s interval)")


# ──────────────────────────────────────────────────────────────────────────────
# Listeners
# ──────────────────────────────────────────────────────────────────────────────

class TabDocumentEventListener(unohelper.Base, XDocumentEventListener, XEventListener):
    """Registered on GlobalEventBroadcaster (belt-and-suspenders)."""

    _OPEN  = frozenset({"onLoad", "onNew", "onDocumentOpened", "onDocumentNew",
                        "OnLoad", "OnNew", "OnCreate"})
    _CLOSE = frozenset({"onClose", "onUnload", "onDocumentClosed",
                        "OnClose", "OnUnload"})

    def __init__(self, ctx):
        self._ctx = ctx

    def notifyDocumentEvent(self, event):
        try:
            name = event.EventName
            _log(f"doc event: {name!r}")
            if name in self._OPEN:
                frame = self._frame_from(event)
                _add_frame(self._ctx, frame)
            elif name in self._CLOSE:
                _clean_dead_frames(self._ctx)
        except Exception:
            _log("notifyDocumentEvent failed:\n" + traceback.format_exc())

    def _frame_from(self, event):
        try:
            model = event.Source
            if model is None:
                return None
            ctrl = model.getCurrentController()
            return ctrl.getFrame() if ctrl else None
        except Exception:
            return None

    def disposing(self, e):
        pass


class TabFrameActionListener(unohelper.Base, XFrameActionListener):
    """
    Registered per Writer frame.  frameAction() is unreliable in the Python
    bridge; we rely only on disposing(), which fires when the frame closes.
    """

    def __init__(self, ctx):
        self._ctx = ctx

    def frameAction(self, event):
        pass   # not reliably called; TabWindowFocusListener handles open detection

    def disposing(self, e):
        """Frame closed – purge dead tabs immediately."""
        try:
            _log("frame disposing – cleaning dead frames")
            _clean_dead_frames(self._ctx)
        except Exception:
            pass


class TabWindowFocusListener(unohelper.Base, XFocusListener):
    """
    Registered on each Writer frame's container window (OS-level window).

    focusLost fires when the user switches away from this window – including
    when a newly opened document window steals focus.  We scan the desktop
    at that point to pick up any frames we don't know about yet.

    focusGained fires when this window becomes active.  We update the active
    frame id and rebuild the toolbar so the correct tab is highlighted.
    """

    def __init__(self, ctx, frame):
        self._ctx   = ctx
        self._frame = frame

    def focusLost(self, event):
        try:
            _log("window focusLost – scanning for new frames")
            _scan_existing_frames(self._ctx)
        except Exception:
            _log("focusLost error:\n" + traceback.format_exc())

    def focusGained(self, event):
        global _active_frame_id
        try:
            _active_frame_id = id(self._frame)
            _clean_dead_frames(self._ctx)
            _rebuild_toolbar(self._ctx)
        except Exception:
            pass

    def disposing(self, e):
        pass


# ──────────────────────────────────────────────────────────────────────────────
# Bootstrap – register global listeners once
# ──────────────────────────────────────────────────────────────────────────────

def _bootstrap(ctx):
    global _bootstrapped, _global_listener, _desktop_listener, _terminate_listener, _ext_modify_listener, _kb_tab_switch
    if _bootstrapped:
        return
    _bootstrapped = True

    # Restore persistent user preferences
    cfg = _load_config()
    _kb_tab_switch = bool(cfg.get("kb_tab_switch", False))
    _log(f"bootstrap: kb_tab_switch={_kb_tab_switch}")

    # GlobalEventBroadcaster (belt-and-suspenders; may not fire in all builds)
    try:
        broadcaster = ctx.getValueByName(
            "/singletons/com.sun.star.frame.theGlobalEventBroadcaster")
        if broadcaster is None:
            broadcaster = ctx.ServiceManager.createInstanceWithContext(
                "com.sun.star.frame.GlobalEventBroadcaster", ctx)
        _global_listener = TabDocumentEventListener(ctx)
        broadcaster.addDocumentEventListener(_global_listener)
        _log("GlobalEventBroadcaster listener registered")
    except Exception:
        _log("GlobalEventBroadcaster registration failed:\n" + traceback.format_exc())

    # Desktop frame-action listener + terminate listener
    try:
        smgr    = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        _desktop_listener = TabDesktopFrameActionListener(ctx)
        desktop.addFrameActionListener(_desktop_listener)
        _log("desktop frame-action listener registered")
        _terminate_listener = TabTerminateListener(ctx)
        desktop.addTerminateListener(_terminate_listener)
        _log("terminate listener registered")
    except Exception:
        _log("desktop listener setup failed:\n" + traceback.format_exc())

    # Extension manager modify listener — cleans up immediately when our
    # extension is removed, so no orphaned toolbar survives an LO restart.
    try:
        ext_mgr = ctx.getValueByName(
            "/singletons/com.sun.star.deployment.theExtensionManager")
        if ext_mgr is None:
            ext_mgr = ctx.ServiceManager.createInstanceWithContext(
                "com.sun.star.deployment.ExtensionManager", ctx)
        if ext_mgr is not None:
            _ext_modify_listener = TabExtensionModifyListener(ctx)
            ext_mgr.addModifyListener(_ext_modify_listener)
            _log("extension modify listener registered")
        else:
            _log("extension modify listener: could not obtain extension manager")
    except Exception:
        _log("extension modify listener failed:\n" + traceback.format_exc())

    # Polling timer – the only mechanism that reliably catches new documents
    _start_poll(ctx)


class TabKeyHandler(unohelper.Base, XKeyHandler, XEventListener):
    """Intercepts Ctrl+Tab / Ctrl+Shift+Tab to cycle through tabs.

    Registered on each Writer frame's controller via XUserInputInterception.
    Only active when _kb_tab_switch is True.
    """

    def __init__(self, ctx):
        self._ctx = ctx

    def keyPressed(self, event):
        try:
            if (event.KeyCode == _Key.TAB
                    and (event.Modifiers & _KeyMod.MOD1)   # Ctrl held
                    and _kb_tab_switch):
                _cycle_tab(self._ctx,
                            backward=bool(event.Modifiers & _KeyMod.SHIFT))
                return True   # consumed — do not propagate
        except Exception:
            _log("TabKeyHandler.keyPressed failed:\n" + traceback.format_exc())
        return False

    def keyReleased(self, event):
        return False

    def disposing(self, e):
        pass


def _cycle_tab(ctx, backward=False):
    """Switch to the next (or previous) tab, wrapping around."""
    global _active_frame_id
    if not _frames:
        return
    try:
        cur = next((i for i, f in enumerate(_frames)
                    if id(f) == _active_frame_id), 0)
        nxt = (cur - 1) % len(_frames) if backward else (cur + 1) % len(_frames)
        target = _frames[nxt]
        _active_frame_id = id(target)
        target.activate()
        target.getContainerWindow().setFocus()
    except Exception:
        _log("_cycle_tab failed:\n" + traceback.format_exc())


def _register_key_handler(ctx, frame):
    """Attach a TabKeyHandler to frame's controller (if not already done)."""
    fid = id(frame)
    if fid in _key_handlers:
        return
    try:
        ctrl = frame.getController()
        if ctrl is None:
            return
        kh = TabKeyHandler(ctx)
        ctrl.addKeyHandler(kh)
        _key_handlers[fid] = kh
        _log(f"key handler registered for frame {fid}")
    except Exception:
        _log(f"_register_key_handler failed:\n" + traceback.format_exc())


def _unregister_key_handler(frame):
    """Remove the TabKeyHandler from frame's controller."""
    fid = id(frame)
    if fid not in _key_handlers:
        return
    try:
        ctrl = frame.getController()
        if ctrl is not None:
            ctrl.removeKeyHandler(_key_handlers[fid])
    except Exception:
        pass
    del _key_handlers[fid]


def _toggle_kb_tab_switch(ctx):
    """Flip the Ctrl+Tab cycling toggle and persist the new value."""
    global _kb_tab_switch
    _kb_tab_switch = not _kb_tab_switch
    cfg = _load_config()
    cfg["kb_tab_switch"] = _kb_tab_switch
    _save_config(cfg)
    _log(f"kb_tab_switch toggled → {_kb_tab_switch}")
    if _kb_tab_switch:
        for frame in _frames:
            _register_key_handler(ctx, frame)
    else:
        for frame in list(_frames):
            _unregister_key_handler(frame)


def _is_our_extension_installed(ctx):
    """Return True if com.github.tabbar is still registered with the extension manager.

    Uses XPackageInformationProvider which takes no complex parameters and returns
    an empty string when the extension is not installed.
    """
    try:
        pkg_info = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.deployment.PackageInformationProvider", ctx)
        if pkg_info is None:
            return True  # can't check — assume still installed
        location = pkg_info.getPackageLocation("com.github.tabbar")
        return bool(location)
    except Exception:
        return True  # assume installed if anything goes wrong


def _remove_toolbar_settings(ctx):
    # Remove toolbar definition from the Writer module config and flush to disk.
    try:
        cfg = _get_writer_cfg(ctx)
        if cfg and cfg.hasSettings(TOOLBAR_URL):
            cfg.removeSettings(TOOLBAR_URL)
            try:
                cfg.store()   # XUIConfigurationPersistence — force immediate disk write
            except Exception:
                pass
            _log("toolbar settings removed and stored")
    except Exception:
        _log("_remove_toolbar_settings failed:\n" + traceback.format_exc())
    # Also destroy the element in every open frame's LayoutManager so the
    # visibility state is not persisted either — without this a grey orphan
    # bar appears on the next LO start even if the definition is gone.
    for frame in list(_frames):
        try:
            frame.LayoutManager.destroyElement(TOOLBAR_URL)
        except Exception:
            pass


class TabExtensionModifyListener(unohelper.Base, XModifyListener):
    """Fires whenever any extension is installed or removed.

    If our extension is no longer in the extension manager's list, the toolbar
    settings are removed from the Writer module config immediately — before LO
    restarts — so no orphaned toolbar bar is left behind.
    """
    def __init__(self, ctx):
        self._ctx = ctx

    def modified(self, event):
        try:
            if not _is_our_extension_installed(self._ctx):
                _log("extension removed: cleaning up toolbar settings")
                _remove_toolbar_settings(self._ctx)
        except Exception:
            _log("TabExtensionModifyListener.modified failed:\n" + traceback.format_exc())

    def disposing(self, event):
        pass


class TabTerminateListener(unohelper.Base, XTerminateListener):
    """Removes the toolbar settings from the Writer module config when LO closes.

    This ensures that if the extension is uninstalled after LO exits, the
    toolbar URL is not left orphaned in the user profile on the next LO start.
    """
    def __init__(self, ctx):
        self._ctx = ctx

    def queryTermination(self, event):
        pass  # never block termination

    def notifyTermination(self, event):
        _remove_toolbar_settings(self._ctx)

    def disposing(self, event):
        pass


class TabDesktopFrameActionListener(unohelper.Base, XFrameActionListener):
    """Belt-and-suspenders on the Desktop frame."""
    def __init__(self, ctx):
        self._ctx = ctx
    def frameAction(self, event):
        pass
    def disposing(self, e):
        pass


# ──────────────────────────────────────────────────────────────────────────────
# Dispatch: tab-switch commands  (.uno:TabBar.Switch.N)
# ──────────────────────────────────────────────────────────────────────────────

class TabDispatch(unohelper.Base, XDispatch, XEventListener):

    def __init__(self, index):
        self._index = index

    def dispatch(self, URL, Arguments):
        global _active_frame_id
        try:
            n = self._index
            if 0 <= n < len(_frames):
                target = _frames[n]
                _active_frame_id = id(target)   # update before activate so
                target.activate()               # the rebuild sees the right id
                try:
                    target.getContainerWindow().setFocus()
                except Exception:
                    pass
                _log(f"switched to tab {n}")
        except Exception:
            _log("TabDispatch.dispatch failed:\n" + traceback.format_exc())

    def addStatusListener(self, listener, URL):
        try:
            n         = self._index
            is_valid  = 0 <= n < len(_frames)
            is_active = is_valid and id(_frames[n]) == _active_frame_id
            ev = uno.createUnoStruct("com.sun.star.frame.FeatureStateEvent")
            ev.FeatureURL = URL
            ev.IsEnabled  = is_valid
            ev.State      = uno.Any("boolean", is_active)
            listener.statusChanged(ev)
        except Exception:
            pass

    def removeStatusListener(self, listener, URL): pass
    def disposing(self, e):                        pass


class TabInterceptor(unohelper.Base, XDispatchProviderInterceptor, XEventListener):

    def __init__(self, frame):
        self._frame  = frame
        self._master = None
        self._slave  = None

    def getMasterDispatchProvider(self):    return self._master
    def getSlaveDispatchProvider(self):     return self._slave
    def setMasterDispatchProvider(self, p): self._master = p
    def setSlaveDispatchProvider(self, p):  self._slave  = p

    def queryDispatch(self, URL, Target, Flags):
        if URL.Complete.startswith(CMD_PREFIX):
            try:
                idx = int(URL.Complete[len(CMD_PREFIX):])
                return TabDispatch(idx)
            except (ValueError, IndexError):
                pass
        return self._slave.queryDispatch(URL, Target, Flags) if self._slave else None

    def queryDispatches(self, Requests):
        return tuple(
            self.queryDispatch(r.FeatureURL, r.FrameName, r.SearchFlags)
            for r in Requests
        )

    def disposing(self, e): pass


# ──────────────────────────────────────────────────────────────────────────────
# Protocol handler  (tabbar:* URLs → reliable startup trigger)
# ──────────────────────────────────────────────────────────────────────────────

class _TabBarMenuDispatch(unohelper.Base, XDispatch, XEventListener):
    """Returned by TabBarProtocolHandler for tabbar:menu.N URLs.

    Shows a popup context menu for tab N when the ▾ dropdown button is clicked.
    The popup is shown relative to the frame's container window.
    """

    def __init__(self, ctx, frame, tab_idx):
        self._ctx     = ctx
        self._frame   = frame
        self._tab_idx = tab_idx

    def dispatch(self, URL, Arguments):
        _log(f"tabbar:menu.{self._tab_idx} dispatched")
        try:
            frame = self._frame
            if frame is None:
                return
            win = frame.getContainerWindow()

            # Try to position the popup just below the tab toolbar.
            popup_x, popup_y = 5, 35
            try:
                lm    = frame.LayoutManager
                tb_el = lm.getElement(TOOLBAR_URL)
                if tb_el:
                    tb_win = tb_el.getRealInterface()
                    if tb_win:
                        ps = tb_win.getPosSize()
                        # Estimate X under this tab's ▾ button.
                        # Each tab occupies 2 toolbar slots (title + ▾).
                        n = len(_frames)
                        if n > 0 and ps.Width > 0:
                            slot = self._tab_idx * 2 + 1   # ▾ is the 2nd slot
                            popup_x = max(0, int(slot * ps.Width / (n * 2)))
                        popup_y = ps.Y + ps.Height
                        win = tb_win   # use toolbar window as parent for better pos
            except Exception:
                pass

            _show_tab_context_menu(self._ctx, frame, win, popup_x, popup_y,
                                   self._tab_idx)
        except Exception:
            _log("_TabBarMenuDispatch.dispatch failed:\n" + traceback.format_exc())

    def addStatusListener(self, listener, URL):
        try:
            ev = uno.createUnoStruct("com.sun.star.frame.FeatureStateEvent")
            ev.FeatureURL = URL
            ev.IsEnabled  = 0 <= self._tab_idx < len(_frames)
            ev.State      = uno.Any("boolean", False)
            listener.statusChanged(ev)
        except Exception:
            pass

    def removeStatusListener(self, listener, URL): pass
    def disposing(self, e):                        pass


class _TabBarCloseDispatch(unohelper.Base, XDispatch, XEventListener):
    """Returned by TabBarProtocolHandler for tabbar:close.N URLs.

    Closes the document at tab index N via the standard UNO CloseDoc command,
    which presents LibreOffice's own "Save changes?" dialog if needed.
    """

    def __init__(self, ctx, frame, tab_idx):
        self._ctx     = ctx
        self._frame   = frame
        self._tab_idx = tab_idx

    def dispatch(self, URL, Arguments):
        _log(f"tabbar:close.{self._tab_idx} dispatched")
        try:
            if 0 <= self._tab_idx < len(_frames):
                target = _frames[self._tab_idx]
                _dispatch_via_helper(self._ctx, target, ".uno:CloseDoc")
        except Exception:
            _log("_TabBarCloseDispatch.dispatch failed:\n" + traceback.format_exc())

    def addStatusListener(self, listener, URL):
        try:
            ev = uno.createUnoStruct("com.sun.star.frame.FeatureStateEvent")
            ev.FeatureURL = URL
            ev.IsEnabled  = 0 <= self._tab_idx < len(_frames)
            ev.State      = uno.Any("boolean", False)
            listener.statusChanged(ev)
        except Exception:
            pass

    def removeStatusListener(self, listener, URL): pass
    def disposing(self, e):                        pass


class _TabBarSetsDispatch(unohelper.Base, XDispatch, XEventListener):
    """Returned by TabBarProtocolHandler for the tabbar:sets URL.

    Shows the ☰ Sets popup menu (save / open / rename / delete tab sets).
    """

    def __init__(self, ctx, frame):
        self._ctx   = ctx
        self._frame = frame

    def dispatch(self, URL, Arguments):
        _log("tabbar:sets dispatched")
        try:
            frame = self._frame
            if frame is None:
                return
            win = frame.getContainerWindow()

            # Position popup near the right end of the tab toolbar
            popup_x, popup_y = 5, 35
            try:
                lm    = frame.LayoutManager
                tb_el = lm.getElement(TOOLBAR_URL)
                if tb_el:
                    tb_win = tb_el.getRealInterface()
                    if tb_win:
                        ps = tb_win.getPosSize()
                        popup_x = max(5, ps.Width - 60)
                        popup_y = ps.Y + ps.Height
                        win = tb_win
            except Exception:
                pass

            _show_sets_menu(self._ctx, frame, win, popup_x, popup_y)
        except Exception:
            _log("_TabBarSetsDispatch.dispatch failed:\n" + traceback.format_exc())

    def addStatusListener(self, listener, URL):
        try:
            ev = uno.createUnoStruct("com.sun.star.frame.FeatureStateEvent")
            ev.FeatureURL = URL
            ev.IsEnabled  = True
            ev.State      = uno.Any("boolean", False)
            listener.statusChanged(ev)
        except Exception:
            pass

    def removeStatusListener(self, listener, URL): pass
    def disposing(self, e):                        pass


class _TabBarInitDispatch(unohelper.Base, XDispatch, XEventListener):
    """Returned by TabBarProtocolHandler for tabbar:init URL."""

    def __init__(self, ctx, frame):
        self._ctx   = ctx
        self._frame = frame

    def dispatch(self, URL, Arguments):
        _log(f"tabbar:init dispatched (frame={'<set>' if self._frame else None})")
        _bootstrap(self._ctx)
        _scan_existing_frames(self._ctx)
        if self._frame is not None:
            _add_frame(self._ctx, self._frame)

    def addStatusListener(self, listener, URL):
        try:
            ev = uno.createUnoStruct("com.sun.star.frame.FeatureStateEvent")
            ev.FeatureURL = URL
            ev.IsEnabled  = True
            ev.State      = uno.Any("boolean", False)
            listener.statusChanged(ev)
        except Exception:
            pass

    def removeStatusListener(self, listener, URL): pass
    def disposing(self, e):                        pass


class TabBarProtocolHandler(unohelper.Base,
                            XDispatchProvider,
                            XInitialization,
                            XServiceInfo):

    def __init__(self, ctx):
        self._ctx   = ctx
        self._frame = None
        _log("TabBarProtocolHandler created")

    def initialize(self, args):
        try:
            if args:
                self._frame = args[0]
            _log(f"TabBarProtocolHandler.initialize: frame={'<set>' if self._frame else None}")
        except Exception:
            _log("initialize failed:\n" + traceback.format_exc())

    def queryDispatch(self, URL, frame_name, search_flags):
        _log(f"TabBarProtocolHandler.queryDispatch: {URL.Complete!r}")
        if URL.Protocol.lower() == "tabbar:":
            complete = URL.Complete
            if complete.startswith("tabbar:menu."):
                try:
                    idx = int(complete[len("tabbar:menu."):])
                    return _TabBarMenuDispatch(self._ctx, self._frame, idx)
                except (ValueError, IndexError):
                    pass
            elif complete.startswith("tabbar:close."):
                try:
                    idx = int(complete[len("tabbar:close."):])
                    return _TabBarCloseDispatch(self._ctx, self._frame, idx)
                except (ValueError, IndexError):
                    pass
            elif complete == CMD_SETS:
                return _TabBarSetsDispatch(self._ctx, self._frame)
            return _TabBarInitDispatch(self._ctx, self._frame)
        return None

    def queryDispatches(self, Requests):
        return tuple(
            self.queryDispatch(r.FeatureURL, r.FrameName, r.SearchFlags)
            for r in Requests
        )

    def getImplementationName(self):    return HANDLER_IMPL
    def supportsService(self, n):       return n == HANDLER_SVC
    def getSupportedServiceNames(self): return (HANDLER_SVC,)


# ──────────────────────────────────────────────────────────────────────────────
# XJob – invoked by LibreOffice's Job Executor
# ──────────────────────────────────────────────────────────────────────────────

class TabBarJob(unohelper.Base, XJob, XServiceInfo):

    def __init__(self, ctx):
        self._ctx = ctx

    def execute(self, Args):
        _log(f"TabBarJob.execute: {len(Args)} arg(s)")
        try:
            event_name = None
            frame      = None
            for nv in Args:
                if nv.Name == "Environment":
                    for env in nv.Value:
                        if   env.Name == "EventName": event_name = env.Value
                        elif env.Name == "Frame":     frame      = env.Value

            _log(f"  event={event_name!r}  frame={'<set>' if frame else '<none>'}")
            _bootstrap(self._ctx)

            if event_name == "onFirstVisibleTask":
                _scan_existing_frames(self._ctx)
            elif event_name in ("onLoad", "onNew", "onDocumentOpened"):
                _add_frame(self._ctx, frame) if frame else None
            elif event_name in ("onClose", "onUnload", "onDocumentClosed"):
                _clean_dead_frames(self._ctx)

        except Exception:
            _log("TabBarJob.execute failed:\n" + traceback.format_exc())
        return None

    def getImplementationName(self):    return JOB_IMPL
    def supportsService(self, n):       return n == JOB_SVC
    def getSupportedServiceNames(self): return (JOB_SVC,)


# ──────────────────────────────────────────────────────────────────────────────
# Registration
# ──────────────────────────────────────────────────────────────────────────────

g_ImplementationHelper = unohelper.ImplementationHelper()

g_ImplementationHelper.addImplementation(
    TabBarJob, JOB_IMPL, (JOB_SVC,))

g_ImplementationHelper.addImplementation(
    TabBarProtocolHandler, HANDLER_IMPL, (HANDLER_SVC,))

_log("tab_bar: g_ImplementationHelper registered")
