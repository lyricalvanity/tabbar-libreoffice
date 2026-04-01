# -*- coding: utf-8 -*-
"""
TabBar – LibreOffice Writer Document Tab Bar  v0.4.0

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

# ──────────────────────────────────────────────────────────────────────────────
# Localisation
# ──────────────────────────────────────────────────────────────────────────────

_STRINGS = {
    "en": {
        "RENAME":               "Rename\u2026",
        "SAVE":                 "Save",
        "SAVE_AS":              "Save As\u2026",
        "MOVE_LEFT":            "Move Left",
        "MOVE_RIGHT":           "Move Right",
        "NEW_DOCUMENT":         "New Document",
        "CLOSE":                "Close",
        "CLOSE_ALL_OTHERS":     "Close All Others",
        "CLOSE_ALL":            "Close All",
        "SAVE_CURRENT_SET":     "Save Current Set\u2026",
        "UPDATE_A_SET":         "Update a Set\u2026",
        "RENAME_A_SET":         "Rename a Set\u2026",
        "DELETE_A_SET":         "Delete a Set\u2026",
        "RESTORE_LAST_SESSION": "Restore Last Session",
        "TAB_KEY_SWITCHING":    "Tab Key Switching",
        "SETS_BUTTON":          "\u2630 Sets",
        "DLG_SAVE_SET":         "Save Tab Set",
        "DLG_RENAME_SET":       "Rename Set",
        "DLG_UPDATE_SET":       "Update Set",
        "DLG_DELETE_SET":       "Delete Set",
        "DLG_RENAME_DOC":       "Rename Document",
        "PROMPT_SET_NAME":      "Tab set name:",
        "PROMPT_RENAME_SET":    "New name for \u2018{old}\u2019:",
        "PROMPT_OVERWRITE":     "Choose set to overwrite with current tabs:",
        "PROMPT_RENAME_WHICH":  "Choose set to rename:",
        "PROMPT_DELETE_WHICH":  "Choose set to delete:",
        "PROMPT_RENAME_DOC":    "New name for \u2018{name}\u2019:",
        "ERR_NO_DOCS":          "No saved documents are open.\nSave your documents before creating a tab set.",
        "ERR_NO_SETS":          "No saved tab sets.",
        "ERR_NO_SETS_UPDATE":   "No saved tab sets to update.",
        "ERR_NO_OPEN_DOCS":     "No saved documents are open to update the set with.",
    },
    "de": {
        "RENAME":               "Umbenennen\u2026",
        "SAVE":                 "Speichern",
        "SAVE_AS":              "Speichern unter\u2026",
        "MOVE_LEFT":            "Nach links",
        "MOVE_RIGHT":           "Nach rechts",
        "NEW_DOCUMENT":         "Neues Dokument",
        "CLOSE":                "Schlie\u00dfen",
        "CLOSE_ALL_OTHERS":     "Alle anderen schlie\u00dfen",
        "CLOSE_ALL":            "Alle schlie\u00dfen",
        "SAVE_CURRENT_SET":     "Aktuelle Gruppe speichern\u2026",
        "UPDATE_A_SET":         "Gruppe aktualisieren\u2026",
        "RENAME_A_SET":         "Gruppe umbenennen\u2026",
        "DELETE_A_SET":         "Gruppe l\u00f6schen\u2026",
        "RESTORE_LAST_SESSION": "Letzte Sitzung wiederherstellen",
        "TAB_KEY_SWITCHING":    "Tab-Taste zum Wechseln",
        "SETS_BUTTON":          "\u2630 Gruppen",
        "DLG_SAVE_SET":         "Registergruppe speichern",
        "DLG_RENAME_SET":       "Gruppe umbenennen",
        "DLG_UPDATE_SET":       "Gruppe aktualisieren",
        "DLG_DELETE_SET":       "Gruppe l\u00f6schen",
        "DLG_RENAME_DOC":       "Dokument umbenennen",
        "PROMPT_SET_NAME":      "Name der Gruppe:",
        "PROMPT_RENAME_SET":    "Neuer Name f\u00fcr \u2018{old}\u2019:",
        "PROMPT_OVERWRITE":     "Gruppe zum \u00dcberschreiben w\u00e4hlen:",
        "PROMPT_RENAME_WHICH":  "Umzubenennende Gruppe w\u00e4hlen:",
        "PROMPT_DELETE_WHICH":  "Zu l\u00f6schende Gruppe w\u00e4hlen:",
        "PROMPT_RENAME_DOC":    "Neuer Name f\u00fcr \u2018{name}\u2019:",
        "ERR_NO_DOCS":          "Keine gespeicherten Dokumente ge\u00f6ffnet.\nBitte speichern Sie Ihre Dokumente zuerst.",
        "ERR_NO_SETS":          "Keine gespeicherten Gruppen vorhanden.",
        "ERR_NO_SETS_UPDATE":   "Keine gespeicherten Gruppen zum Aktualisieren.",
        "ERR_NO_OPEN_DOCS":     "Keine gespeicherten Dokumente ge\u00f6ffnet.",
    },
    "fr": {
        "RENAME":               "Renommer\u2026",
        "SAVE":                 "Enregistrer",
        "SAVE_AS":              "Enregistrer sous\u2026",
        "MOVE_LEFT":            "D\u00e9placer \u00e0 gauche",
        "MOVE_RIGHT":           "D\u00e9placer \u00e0 droite",
        "NEW_DOCUMENT":         "Nouveau document",
        "CLOSE":                "Fermer",
        "CLOSE_ALL_OTHERS":     "Fermer les autres",
        "CLOSE_ALL":            "Tout fermer",
        "SAVE_CURRENT_SET":     "Enregistrer le groupe actuel\u2026",
        "UPDATE_A_SET":         "Mettre \u00e0 jour un groupe\u2026",
        "RENAME_A_SET":         "Renommer un groupe\u2026",
        "DELETE_A_SET":         "Supprimer un groupe\u2026",
        "RESTORE_LAST_SESSION": "Restaurer la derni\u00e8re session",
        "TAB_KEY_SWITCHING":    "Navigation par tabulation",
        "SETS_BUTTON":          "\u2630 Groupes",
        "DLG_SAVE_SET":         "Enregistrer le groupe",
        "DLG_RENAME_SET":       "Renommer le groupe",
        "DLG_UPDATE_SET":       "Mettre \u00e0 jour le groupe",
        "DLG_DELETE_SET":       "Supprimer le groupe",
        "DLG_RENAME_DOC":       "Renommer le document",
        "PROMPT_SET_NAME":      "Nom du groupe\u00a0:",
        "PROMPT_RENAME_SET":    "Nouveau nom pour \u00ab\u00a0{old}\u00a0\u00bb\u00a0:",
        "PROMPT_OVERWRITE":     "Choisir le groupe \u00e0 \u00e9craser\u00a0:",
        "PROMPT_RENAME_WHICH":  "Choisir le groupe \u00e0 renommer\u00a0:",
        "PROMPT_DELETE_WHICH":  "Choisir le groupe \u00e0 supprimer\u00a0:",
        "PROMPT_RENAME_DOC":    "Nouveau nom pour \u00ab\u00a0{name}\u00a0\u00bb\u00a0:",
        "ERR_NO_DOCS":          "Aucun document enregistr\u00e9 n\u2019est ouvert.\nVeuillez enregistrer vos documents d\u2019abord.",
        "ERR_NO_SETS":          "Aucun groupe enregistr\u00e9.",
        "ERR_NO_SETS_UPDATE":   "Aucun groupe \u00e0 mettre \u00e0 jour.",
        "ERR_NO_OPEN_DOCS":     "Aucun document enregistr\u00e9 ouvert.",
    },
    "es": {
        "RENAME":               "Renombrar\u2026",
        "SAVE":                 "Guardar",
        "SAVE_AS":              "Guardar como\u2026",
        "MOVE_LEFT":            "Mover a la izquierda",
        "MOVE_RIGHT":           "Mover a la derecha",
        "NEW_DOCUMENT":         "Nuevo documento",
        "CLOSE":                "Cerrar",
        "CLOSE_ALL_OTHERS":     "Cerrar los dem\u00e1s",
        "CLOSE_ALL":            "Cerrar todo",
        "SAVE_CURRENT_SET":     "Guardar conjunto actual\u2026",
        "UPDATE_A_SET":         "Actualizar un conjunto\u2026",
        "RENAME_A_SET":         "Renombrar un conjunto\u2026",
        "DELETE_A_SET":         "Eliminar un conjunto\u2026",
        "RESTORE_LAST_SESSION": "Restaurar \u00faltima sesi\u00f3n",
        "TAB_KEY_SWITCHING":    "Cambiar con tabulador",
        "SETS_BUTTON":          "\u2630 Conjuntos",
        "DLG_SAVE_SET":         "Guardar conjunto",
        "DLG_RENAME_SET":       "Renombrar conjunto",
        "DLG_UPDATE_SET":       "Actualizar conjunto",
        "DLG_DELETE_SET":       "Eliminar conjunto",
        "DLG_RENAME_DOC":       "Renombrar documento",
        "PROMPT_SET_NAME":      "Nombre del conjunto:",
        "PROMPT_RENAME_SET":    "Nuevo nombre para \u2018{old}\u2019:",
        "PROMPT_OVERWRITE":     "Elegir conjunto a sobreescribir:",
        "PROMPT_RENAME_WHICH":  "Elegir conjunto a renombrar:",
        "PROMPT_DELETE_WHICH":  "Elegir conjunto a eliminar:",
        "PROMPT_RENAME_DOC":    "Nuevo nombre para \u2018{name}\u2019:",
        "ERR_NO_DOCS":          "No hay documentos guardados abiertos.\nGuarde sus documentos primero.",
        "ERR_NO_SETS":          "No hay conjuntos guardados.",
        "ERR_NO_SETS_UPDATE":   "No hay conjuntos guardados para actualizar.",
        "ERR_NO_OPEN_DOCS":     "No hay documentos guardados abiertos.",
    },
    "it": {
        "RENAME":               "Rinomina\u2026",
        "SAVE":                 "Salva",
        "SAVE_AS":              "Salva come\u2026",
        "MOVE_LEFT":            "Sposta a sinistra",
        "MOVE_RIGHT":           "Sposta a destra",
        "NEW_DOCUMENT":         "Nuovo documento",
        "CLOSE":                "Chiudi",
        "CLOSE_ALL_OTHERS":     "Chiudi gli altri",
        "CLOSE_ALL":            "Chiudi tutto",
        "SAVE_CURRENT_SET":     "Salva gruppo corrente\u2026",
        "UPDATE_A_SET":         "Aggiorna un gruppo\u2026",
        "RENAME_A_SET":         "Rinomina un gruppo\u2026",
        "DELETE_A_SET":         "Elimina un gruppo\u2026",
        "RESTORE_LAST_SESSION": "Ripristina ultima sessione",
        "TAB_KEY_SWITCHING":    "Navigazione con Tab",
        "SETS_BUTTON":          "\u2630 Gruppi",
        "DLG_SAVE_SET":         "Salva gruppo schede",
        "DLG_RENAME_SET":       "Rinomina gruppo",
        "DLG_UPDATE_SET":       "Aggiorna gruppo",
        "DLG_DELETE_SET":       "Elimina gruppo",
        "DLG_RENAME_DOC":       "Rinomina documento",
        "PROMPT_SET_NAME":      "Nome del gruppo:",
        "PROMPT_RENAME_SET":    "Nuovo nome per \u2018{old}\u2019:",
        "PROMPT_OVERWRITE":     "Scegli il gruppo da sovrascrivere:",
        "PROMPT_RENAME_WHICH":  "Scegli il gruppo da rinominare:",
        "PROMPT_DELETE_WHICH":  "Scegli il gruppo da eliminare:",
        "PROMPT_RENAME_DOC":    "Nuovo nome per \u2018{name}\u2019:",
        "ERR_NO_DOCS":          "Nessun documento salvato aperto.\nSalva prima i tuoi documenti.",
        "ERR_NO_SETS":          "Nessun gruppo salvato.",
        "ERR_NO_SETS_UPDATE":   "Nessun gruppo da aggiornare.",
        "ERR_NO_OPEN_DOCS":     "Nessun documento salvato aperto.",
    },
    "nl": {
        "RENAME":               "Hernoemen\u2026",
        "SAVE":                 "Opslaan",
        "SAVE_AS":              "Opslaan als\u2026",
        "MOVE_LEFT":            "Naar links verplaatsen",
        "MOVE_RIGHT":           "Naar rechts verplaatsen",
        "NEW_DOCUMENT":         "Nieuw document",
        "CLOSE":                "Sluiten",
        "CLOSE_ALL_OTHERS":     "Alle andere sluiten",
        "CLOSE_ALL":            "Alles sluiten",
        "SAVE_CURRENT_SET":     "Huidige set opslaan\u2026",
        "UPDATE_A_SET":         "Set bijwerken\u2026",
        "RENAME_A_SET":         "Set hernoemen\u2026",
        "DELETE_A_SET":         "Set verwijderen\u2026",
        "RESTORE_LAST_SESSION": "Laatste sessie herstellen",
        "TAB_KEY_SWITCHING":    "Tab-toets om te wisselen",
        "SETS_BUTTON":          "\u2630 Sets",
        "DLG_SAVE_SET":         "Tabbladset opslaan",
        "DLG_RENAME_SET":       "Set hernoemen",
        "DLG_UPDATE_SET":       "Set bijwerken",
        "DLG_DELETE_SET":       "Set verwijderen",
        "DLG_RENAME_DOC":       "Document hernoemen",
        "PROMPT_SET_NAME":      "Naam van de set:",
        "PROMPT_RENAME_SET":    "Nieuwe naam voor \u2018{old}\u2019:",
        "PROMPT_OVERWRITE":     "Kies de te overschrijven set:",
        "PROMPT_RENAME_WHICH":  "Kies de te hernoemen set:",
        "PROMPT_DELETE_WHICH":  "Kies de te verwijderen set:",
        "PROMPT_RENAME_DOC":    "Nieuwe naam voor \u2018{name}\u2019:",
        "ERR_NO_DOCS":          "Geen opgeslagen documenten open.\nSla uw documenten eerst op.",
        "ERR_NO_SETS":          "Geen opgeslagen sets.",
        "ERR_NO_SETS_UPDATE":   "Geen sets om bij te werken.",
        "ERR_NO_OPEN_DOCS":     "Geen opgeslagen documenten open.",
    },
    "pl": {
        "RENAME":               "Zmie\u0144 nazw\u0119\u2026",
        "SAVE":                 "Zapisz",
        "SAVE_AS":              "Zapisz jako\u2026",
        "MOVE_LEFT":            "Przesu\u0144 w lewo",
        "MOVE_RIGHT":           "Przesu\u0144 w prawo",
        "NEW_DOCUMENT":         "Nowy dokument",
        "CLOSE":                "Zamknij",
        "CLOSE_ALL_OTHERS":     "Zamknij pozosta\u0142e",
        "CLOSE_ALL":            "Zamknij wszystkie",
        "SAVE_CURRENT_SET":     "Zapisz bie\u017c\u0105cy zestaw\u2026",
        "UPDATE_A_SET":         "Aktualizuj zestaw\u2026",
        "RENAME_A_SET":         "Zmie\u0144 nazw\u0119 zestawu\u2026",
        "DELETE_A_SET":         "Usu\u0144 zestaw\u2026",
        "RESTORE_LAST_SESSION": "Przywr\u00f3\u0107 ostatni\u0105 sesj\u0119",
        "TAB_KEY_SWITCHING":    "Prze\u0142\u0105czanie klawiszem Tab",
        "SETS_BUTTON":          "\u2630 Zestawy",
        "DLG_SAVE_SET":         "Zapisz zestaw kart",
        "DLG_RENAME_SET":       "Zmie\u0144 nazw\u0119 zestawu",
        "DLG_UPDATE_SET":       "Aktualizuj zestaw",
        "DLG_DELETE_SET":       "Usu\u0144 zestaw",
        "DLG_RENAME_DOC":       "Zmie\u0144 nazw\u0119 dokumentu",
        "PROMPT_SET_NAME":      "Nazwa zestawu:",
        "PROMPT_RENAME_SET":    "Nowa nazwa dla \u2018{old}\u2019:",
        "PROMPT_OVERWRITE":     "Wybierz zestaw do nadpisania:",
        "PROMPT_RENAME_WHICH":  "Wybierz zestaw do zmiany nazwy:",
        "PROMPT_DELETE_WHICH":  "Wybierz zestaw do usuni\u0119cia:",
        "PROMPT_RENAME_DOC":    "Nowa nazwa dla \u2018{name}\u2019:",
        "ERR_NO_DOCS":          "Brak otwartych zapisanych dokument\u00f3w.\nZapisz dokumenty przed utworzeniem zestawu.",
        "ERR_NO_SETS":          "Brak zapisanych zestaw\u00f3w.",
        "ERR_NO_SETS_UPDATE":   "Brak zestaw\u00f3w do aktualizacji.",
        "ERR_NO_OPEN_DOCS":     "Brak otwartych zapisanych dokument\u00f3w.",
    },
    "pt": {
        "RENAME":               "Renomear\u2026",
        "SAVE":                 "Guardar",
        "SAVE_AS":              "Guardar como\u2026",
        "MOVE_LEFT":            "Mover para a esquerda",
        "MOVE_RIGHT":           "Mover para a direita",
        "NEW_DOCUMENT":         "Novo documento",
        "CLOSE":                "Fechar",
        "CLOSE_ALL_OTHERS":     "Fechar os outros",
        "CLOSE_ALL":            "Fechar tudo",
        "SAVE_CURRENT_SET":     "Guardar conjunto atual\u2026",
        "UPDATE_A_SET":         "Atualizar um conjunto\u2026",
        "RENAME_A_SET":         "Renomear um conjunto\u2026",
        "DELETE_A_SET":         "Eliminar um conjunto\u2026",
        "RESTORE_LAST_SESSION": "Restaurar \u00faltima sess\u00e3o",
        "TAB_KEY_SWITCHING":    "Navegar com Tab",
        "SETS_BUTTON":          "\u2630 Conjuntos",
        "DLG_SAVE_SET":         "Guardar conjunto",
        "DLG_RENAME_SET":       "Renomear conjunto",
        "DLG_UPDATE_SET":       "Atualizar conjunto",
        "DLG_DELETE_SET":       "Eliminar conjunto",
        "DLG_RENAME_DOC":       "Renomear documento",
        "PROMPT_SET_NAME":      "Nome do conjunto:",
        "PROMPT_RENAME_SET":    "Novo nome para \u2018{old}\u2019:",
        "PROMPT_OVERWRITE":     "Escolher conjunto a substituir:",
        "PROMPT_RENAME_WHICH":  "Escolher conjunto a renomear:",
        "PROMPT_DELETE_WHICH":  "Escolher conjunto a eliminar:",
        "PROMPT_RENAME_DOC":    "Novo nome para \u2018{name}\u2019:",
        "ERR_NO_DOCS":          "Nenhum documento guardado aberto.\nGuarde os seus documentos primeiro.",
        "ERR_NO_SETS":          "Nenhum conjunto guardado.",
        "ERR_NO_SETS_UPDATE":   "Nenhum conjunto para atualizar.",
        "ERR_NO_OPEN_DOCS":     "Nenhum documento guardado aberto.",
    },
    "pt-br": {
        "RENAME":               "Renomear\u2026",
        "SAVE":                 "Salvar",
        "SAVE_AS":              "Salvar como\u2026",
        "MOVE_LEFT":            "Mover para a esquerda",
        "MOVE_RIGHT":           "Mover para a direita",
        "NEW_DOCUMENT":         "Novo documento",
        "CLOSE":                "Fechar",
        "CLOSE_ALL_OTHERS":     "Fechar os outros",
        "CLOSE_ALL":            "Fechar tudo",
        "SAVE_CURRENT_SET":     "Salvar conjunto atual\u2026",
        "UPDATE_A_SET":         "Atualizar um conjunto\u2026",
        "RENAME_A_SET":         "Renomear um conjunto\u2026",
        "DELETE_A_SET":         "Excluir um conjunto\u2026",
        "RESTORE_LAST_SESSION": "Restaurar \u00faltima sess\u00e3o",
        "TAB_KEY_SWITCHING":    "Navegar com Tab",
        "SETS_BUTTON":          "\u2630 Conjuntos",
        "DLG_SAVE_SET":         "Salvar conjunto",
        "DLG_RENAME_SET":       "Renomear conjunto",
        "DLG_UPDATE_SET":       "Atualizar conjunto",
        "DLG_DELETE_SET":       "Excluir conjunto",
        "DLG_RENAME_DOC":       "Renomear documento",
        "PROMPT_SET_NAME":      "Nome do conjunto:",
        "PROMPT_RENAME_SET":    "Novo nome para \u2018{old}\u2019:",
        "PROMPT_OVERWRITE":     "Escolher conjunto a sobrescrever:",
        "PROMPT_RENAME_WHICH":  "Escolher conjunto a renomear:",
        "PROMPT_DELETE_WHICH":  "Escolher conjunto a excluir:",
        "PROMPT_RENAME_DOC":    "Novo nome para \u2018{name}\u2019:",
        "ERR_NO_DOCS":          "Nenhum documento salvo aberto.\nSalve seus documentos primeiro.",
        "ERR_NO_SETS":          "Nenhum conjunto salvo.",
        "ERR_NO_SETS_UPDATE":   "Nenhum conjunto para atualizar.",
        "ERR_NO_OPEN_DOCS":     "Nenhum documento salvo aberto.",
    },
    "ru": {
        "RENAME":               "\u041f\u0435\u0440\u0435\u0438\u043c\u0435\u043d\u043e\u0432\u0430\u0442\u044c\u2026",
        "SAVE":                 "\u0421\u043e\u0445\u0440\u0430\u043d\u0438\u0442\u044c",
        "SAVE_AS":              "\u0421\u043e\u0445\u0440\u0430\u043d\u0438\u0442\u044c \u043a\u0430\u043a\u2026",
        "MOVE_LEFT":            "\u041f\u0435\u0440\u0435\u043c\u0435\u0441\u0442\u0438\u0442\u044c \u0432\u043b\u0435\u0432\u043e",
        "MOVE_RIGHT":           "\u041f\u0435\u0440\u0435\u043c\u0435\u0441\u0442\u0438\u0442\u044c \u0432\u043f\u0440\u0430\u0432\u043e",
        "NEW_DOCUMENT":         "\u041d\u043e\u0432\u044b\u0439 \u0434\u043e\u043a\u0443\u043c\u0435\u043d\u0442",
        "CLOSE":                "\u0417\u0430\u043a\u0440\u044b\u0442\u044c",
        "CLOSE_ALL_OTHERS":     "\u0417\u0430\u043a\u0440\u044b\u0442\u044c \u043e\u0441\u0442\u0430\u043b\u044c\u043d\u044b\u0435",
        "CLOSE_ALL":            "\u0417\u0430\u043a\u0440\u044b\u0442\u044c \u0432\u0441\u0435",
        "SAVE_CURRENT_SET":     "\u0421\u043e\u0445\u0440\u0430\u043d\u0438\u0442\u044c \u0442\u0435\u043a\u0443\u0449\u0438\u0439 \u043d\u0430\u0431\u043e\u0440\u2026",
        "UPDATE_A_SET":         "\u041e\u0431\u043d\u043e\u0432\u0438\u0442\u044c \u043d\u0430\u0431\u043e\u0440\u2026",
        "RENAME_A_SET":         "\u041f\u0435\u0440\u0435\u0438\u043c\u0435\u043d\u043e\u0432\u0430\u0442\u044c \u043d\u0430\u0431\u043e\u0440\u2026",
        "DELETE_A_SET":         "\u0423\u0434\u0430\u043b\u0438\u0442\u044c \u043d\u0430\u0431\u043e\u0440\u2026",
        "RESTORE_LAST_SESSION": "\u0412\u043e\u0441\u0441\u0442\u0430\u043d\u043e\u0432\u0438\u0442\u044c \u043f\u043e\u0441\u043b\u0435\u0434\u043d\u044e\u044e \u0441\u0435\u0441\u0441\u0438\u044e",
        "TAB_KEY_SWITCHING":    "\u041f\u0435\u0440\u0435\u043a\u043b\u044e\u0447\u0435\u043d\u0438\u0435 Tab",
        "SETS_BUTTON":          "\u2630 \u041d\u0430\u0431\u043e\u0440\u044b",
        "DLG_SAVE_SET":         "\u0421\u043e\u0445\u0440\u0430\u043d\u0438\u0442\u044c \u043d\u0430\u0431\u043e\u0440",
        "DLG_RENAME_SET":       "\u041f\u0435\u0440\u0435\u0438\u043c\u0435\u043d\u043e\u0432\u0430\u0442\u044c \u043d\u0430\u0431\u043e\u0440",
        "DLG_UPDATE_SET":       "\u041e\u0431\u043d\u043e\u0432\u0438\u0442\u044c \u043d\u0430\u0431\u043e\u0440",
        "DLG_DELETE_SET":       "\u0423\u0434\u0430\u043b\u0438\u0442\u044c \u043d\u0430\u0431\u043e\u0440",
        "DLG_RENAME_DOC":       "\u041f\u0435\u0440\u0435\u0438\u043c\u0435\u043d\u043e\u0432\u0430\u0442\u044c \u0434\u043e\u043a\u0443\u043c\u0435\u043d\u0442",
        "PROMPT_SET_NAME":      "\u041d\u0430\u0437\u0432\u0430\u043d\u0438\u0435 \u043d\u0430\u0431\u043e\u0440\u0430:",
        "PROMPT_RENAME_SET":    "\u041d\u043e\u0432\u043e\u0435 \u043d\u0430\u0437\u0432\u0430\u043d\u0438\u0435 \u0434\u043b\u044f \u00ab{old}\u00bb:",
        "PROMPT_OVERWRITE":     "\u0412\u044b\u0431\u0435\u0440\u0438\u0442\u0435 \u043d\u0430\u0431\u043e\u0440 \u0434\u043b\u044f \u043f\u0435\u0440\u0435\u0437\u0430\u043f\u0438\u0441\u0438:",
        "PROMPT_RENAME_WHICH":  "\u0412\u044b\u0431\u0435\u0440\u0438\u0442\u0435 \u043d\u0430\u0431\u043e\u0440 \u0434\u043b\u044f \u043f\u0435\u0440\u0435\u0438\u043c\u0435\u043d\u043e\u0432\u0430\u043d\u0438\u044f:",
        "PROMPT_DELETE_WHICH":  "\u0412\u044b\u0431\u0435\u0440\u0438\u0442\u0435 \u043d\u0430\u0431\u043e\u0440 \u0434\u043b\u044f \u0443\u0434\u0430\u043b\u0435\u043d\u0438\u044f:",
        "PROMPT_RENAME_DOC":    "\u041d\u043e\u0432\u043e\u0435 \u0438\u043c\u044f \u0434\u043b\u044f \u00ab{name}\u00bb:",
        "ERR_NO_DOCS":          "\u041d\u0435\u0442 \u043e\u0442\u043a\u0440\u044b\u0442\u044b\u0445 \u0441\u043e\u0445\u0440\u0430\u043d\u0451\u043d\u043d\u044b\u0445 \u0434\u043e\u043a\u0443\u043c\u0435\u043d\u0442\u043e\u0432.\n\u0421\u043d\u0430\u0447\u0430\u043b\u0430 \u0441\u043e\u0445\u0440\u0430\u043d\u0438\u0442\u0435 \u0434\u043e\u043a\u0443\u043c\u0435\u043d\u0442\u044b.",
        "ERR_NO_SETS":          "\u041d\u0435\u0442 \u0441\u043e\u0445\u0440\u0430\u043d\u0451\u043d\u043d\u044b\u0445 \u043d\u0430\u0431\u043e\u0440\u043e\u0432.",
        "ERR_NO_SETS_UPDATE":   "\u041d\u0435\u0442 \u043d\u0430\u0431\u043e\u0440\u043e\u0432 \u0434\u043b\u044f \u043e\u0431\u043d\u043e\u0432\u043b\u0435\u043d\u0438\u044f.",
        "ERR_NO_OPEN_DOCS":     "\u041d\u0435\u0442 \u043e\u0442\u043a\u0440\u044b\u0442\u044b\u0445 \u0441\u043e\u0445\u0440\u0430\u043d\u0451\u043d\u043d\u044b\u0445 \u0434\u043e\u043a\u0443\u043c\u0435\u043d\u0442\u043e\u0432.",
    },
    "sv": {
        "RENAME":               "Byt namn\u2026",
        "SAVE":                 "Spara",
        "SAVE_AS":              "Spara som\u2026",
        "MOVE_LEFT":            "Flytta v\u00e4nster",
        "MOVE_RIGHT":           "Flytta h\u00f6ger",
        "NEW_DOCUMENT":         "Nytt dokument",
        "CLOSE":                "St\u00e4ng",
        "CLOSE_ALL_OTHERS":     "St\u00e4ng \u00f6vriga",
        "CLOSE_ALL":            "St\u00e4ng alla",
        "SAVE_CURRENT_SET":     "Spara aktuell grupp\u2026",
        "UPDATE_A_SET":         "Uppdatera en grupp\u2026",
        "RENAME_A_SET":         "Byt namn p\u00e5 en grupp\u2026",
        "DELETE_A_SET":         "Ta bort en grupp\u2026",
        "RESTORE_LAST_SESSION": "\u00c5terst\u00e4ll senaste session",
        "TAB_KEY_SWITCHING":    "V\u00e4xla med Tab",
        "SETS_BUTTON":          "\u2630 Grupper",
        "DLG_SAVE_SET":         "Spara flikgrupp",
        "DLG_RENAME_SET":       "Byt namn p\u00e5 grupp",
        "DLG_UPDATE_SET":       "Uppdatera grupp",
        "DLG_DELETE_SET":       "Ta bort grupp",
        "DLG_RENAME_DOC":       "Byt namn p\u00e5 dokument",
        "PROMPT_SET_NAME":      "Gruppnamn:",
        "PROMPT_RENAME_SET":    "Nytt namn f\u00f6r \u2018{old}\u2019:",
        "PROMPT_OVERWRITE":     "V\u00e4lj grupp att skriva \u00f6ver:",
        "PROMPT_RENAME_WHICH":  "V\u00e4lj grupp att byta namn p\u00e5:",
        "PROMPT_DELETE_WHICH":  "V\u00e4lj grupp att ta bort:",
        "PROMPT_RENAME_DOC":    "Nytt namn f\u00f6r \u2018{name}\u2019:",
        "ERR_NO_DOCS":          "Inga sparade dokument \u00e4r \u00f6ppna.\nSpara dina dokument f\u00f6rst.",
        "ERR_NO_SETS":          "Inga sparade grupper.",
        "ERR_NO_SETS_UPDATE":   "Inga grupper att uppdatera.",
        "ERR_NO_OPEN_DOCS":     "Inga sparade dokument \u00e4r \u00f6ppna.",
    },
    "da": {
        "RENAME":               "Omd\u00f8b\u2026",
        "SAVE":                 "Gem",
        "SAVE_AS":              "Gem som\u2026",
        "MOVE_LEFT":            "Flyt til venstre",
        "MOVE_RIGHT":           "Flyt til h\u00f8jre",
        "NEW_DOCUMENT":         "Nyt dokument",
        "CLOSE":                "Luk",
        "CLOSE_ALL_OTHERS":     "Luk andre",
        "CLOSE_ALL":            "Luk alle",
        "SAVE_CURRENT_SET":     "Gem aktuel gruppe\u2026",
        "UPDATE_A_SET":         "Opdater en gruppe\u2026",
        "RENAME_A_SET":         "Omd\u00f8b en gruppe\u2026",
        "DELETE_A_SET":         "Slet en gruppe\u2026",
        "RESTORE_LAST_SESSION": "Gendan seneste session",
        "TAB_KEY_SWITCHING":    "Skift med Tab",
        "SETS_BUTTON":          "\u2630 Grupper",
        "DLG_SAVE_SET":         "Gem fanegruppe",
        "DLG_RENAME_SET":       "Omd\u00f8b gruppe",
        "DLG_UPDATE_SET":       "Opdater gruppe",
        "DLG_DELETE_SET":       "Slet gruppe",
        "DLG_RENAME_DOC":       "Omd\u00f8b dokument",
        "PROMPT_SET_NAME":      "Gruppenavn:",
        "PROMPT_RENAME_SET":    "Nyt navn for \u2018{old}\u2019:",
        "PROMPT_OVERWRITE":     "V\u00e6lg gruppe at overskrive:",
        "PROMPT_RENAME_WHICH":  "V\u00e6lg gruppe at omd\u00f8be:",
        "PROMPT_DELETE_WHICH":  "V\u00e6lg gruppe at slette:",
        "PROMPT_RENAME_DOC":    "Nyt navn for \u2018{name}\u2019:",
        "ERR_NO_DOCS":          "Ingen gemte dokumenter er \u00e5bne.\nGem dine dokumenter f\u00f8rst.",
        "ERR_NO_SETS":          "Ingen gemte grupper.",
        "ERR_NO_SETS_UPDATE":   "Ingen grupper at opdatere.",
        "ERR_NO_OPEN_DOCS":     "Ingen gemte dokumenter er \u00e5bne.",
    },
    "fi": {
        "RENAME":               "Nime\u00e4 uudelleen\u2026",
        "SAVE":                 "Tallenna",
        "SAVE_AS":              "Tallenna nimell\u00e4\u2026",
        "MOVE_LEFT":            "Siirr\u00e4 vasemmalle",
        "MOVE_RIGHT":           "Siirr\u00e4 oikealle",
        "NEW_DOCUMENT":         "Uusi asiakirja",
        "CLOSE":                "Sulje",
        "CLOSE_ALL_OTHERS":     "Sulje muut",
        "CLOSE_ALL":            "Sulje kaikki",
        "SAVE_CURRENT_SET":     "Tallenna nykyinen ryhm\u00e4\u2026",
        "UPDATE_A_SET":         "P\u00e4ivit\u00e4 ryhm\u00e4\u2026",
        "RENAME_A_SET":         "Nime\u00e4 ryhm\u00e4 uudelleen\u2026",
        "DELETE_A_SET":         "Poista ryhm\u00e4\u2026",
        "RESTORE_LAST_SESSION": "Palauta edellinen istunto",
        "TAB_KEY_SWITCHING":    "Vaihda Tab-n\u00e4pp\u00e4imell\u00e4",
        "SETS_BUTTON":          "\u2630 Ryhm\u00e4t",
        "DLG_SAVE_SET":         "Tallenna v\u00e4lilehtiryhm\u00e4",
        "DLG_RENAME_SET":       "Nime\u00e4 ryhm\u00e4 uudelleen",
        "DLG_UPDATE_SET":       "P\u00e4ivit\u00e4 ryhm\u00e4",
        "DLG_DELETE_SET":       "Poista ryhm\u00e4",
        "DLG_RENAME_DOC":       "Nime\u00e4 asiakirja uudelleen",
        "PROMPT_SET_NAME":      "Ryhm\u00e4n nimi:",
        "PROMPT_RENAME_SET":    "Uusi nimi kohteelle \u2018{old}\u2019:",
        "PROMPT_OVERWRITE":     "Valitse korvattava ryhm\u00e4:",
        "PROMPT_RENAME_WHICH":  "Valitse uudelleennimett\u00e4v\u00e4 ryhm\u00e4:",
        "PROMPT_DELETE_WHICH":  "Valitse poistettava ryhm\u00e4:",
        "PROMPT_RENAME_DOC":    "Uusi nimi kohteelle \u2018{name}\u2019:",
        "ERR_NO_DOCS":          "Ei avoimia tallennettuja asiakirjoja.\nTallenna asiakirjasi ensin.",
        "ERR_NO_SETS":          "Ei tallennettuja ryhmi\u00e4.",
        "ERR_NO_SETS_UPDATE":   "Ei ryhmi\u00e4 p\u00e4ivitett\u00e4v\u00e4ksi.",
        "ERR_NO_OPEN_DOCS":     "Ei avoimia tallennettuja asiakirjoja.",
    },
    "nb": {
        "RENAME":               "Gi nytt navn\u2026",
        "SAVE":                 "Lagre",
        "SAVE_AS":              "Lagre som\u2026",
        "MOVE_LEFT":            "Flytt til venstre",
        "MOVE_RIGHT":           "Flytt til h\u00f8yre",
        "NEW_DOCUMENT":         "Nytt dokument",
        "CLOSE":                "Lukk",
        "CLOSE_ALL_OTHERS":     "Lukk andre",
        "CLOSE_ALL":            "Lukk alle",
        "SAVE_CURRENT_SET":     "Lagre gjeldende gruppe\u2026",
        "UPDATE_A_SET":         "Oppdater en gruppe\u2026",
        "RENAME_A_SET":         "Gi gruppe nytt navn\u2026",
        "DELETE_A_SET":         "Slett en gruppe\u2026",
        "RESTORE_LAST_SESSION": "Gjenopprett forrige \u00f8kt",
        "TAB_KEY_SWITCHING":    "Bytt med Tab",
        "SETS_BUTTON":          "\u2630 Grupper",
        "DLG_SAVE_SET":         "Lagre fanergruppe",
        "DLG_RENAME_SET":       "Gi gruppe nytt navn",
        "DLG_UPDATE_SET":       "Oppdater gruppe",
        "DLG_DELETE_SET":       "Slett gruppe",
        "DLG_RENAME_DOC":       "Gi dokument nytt navn",
        "PROMPT_SET_NAME":      "Gruppenavn:",
        "PROMPT_RENAME_SET":    "Nytt navn for \u2018{old}\u2019:",
        "PROMPT_OVERWRITE":     "Velg gruppe \u00e5 overskrive:",
        "PROMPT_RENAME_WHICH":  "Velg gruppe \u00e5 gi nytt navn:",
        "PROMPT_DELETE_WHICH":  "Velg gruppe \u00e5 slette:",
        "PROMPT_RENAME_DOC":    "Nytt navn for \u2018{name}\u2019:",
        "ERR_NO_DOCS":          "Ingen lagrede dokumenter er \u00e5pne.\nLagre dokumentene dine f\u00f8rst.",
        "ERR_NO_SETS":          "Ingen lagrede grupper.",
        "ERR_NO_SETS_UPDATE":   "Ingen grupper \u00e5 oppdatere.",
        "ERR_NO_OPEN_DOCS":     "Ingen lagrede dokumenter er \u00e5pne.",
    },
    "cs": {
        "RENAME":               "P\u0159ejmenovat\u2026",
        "SAVE":                 "Ulo\u017eit",
        "SAVE_AS":              "Ulo\u017eit jako\u2026",
        "MOVE_LEFT":            "P\u0159esunout doleva",
        "MOVE_RIGHT":           "P\u0159esunout doprava",
        "NEW_DOCUMENT":         "Nov\u00fd dokument",
        "CLOSE":                "Zav\u0159\u00edt",
        "CLOSE_ALL_OTHERS":     "Zav\u0159\u00edt ostatn\u00ed",
        "CLOSE_ALL":            "Zav\u0159\u00edt v\u0161e",
        "SAVE_CURRENT_SET":     "Ulo\u017eit aktu\u00e1ln\u00ed sadu\u2026",
        "UPDATE_A_SET":         "Aktualizovat sadu\u2026",
        "RENAME_A_SET":         "P\u0159ejmenovat sadu\u2026",
        "DELETE_A_SET":         "Odstranit sadu\u2026",
        "RESTORE_LAST_SESSION": "Obnovit posledn\u00ed relaci",
        "TAB_KEY_SWITCHING":    "P\u0159ep\u00ednat kl\u00e1vesou Tab",
        "SETS_BUTTON":          "\u2630 Sady",
        "DLG_SAVE_SET":         "Ulo\u017eit sadu karet",
        "DLG_RENAME_SET":       "P\u0159ejmenovat sadu",
        "DLG_UPDATE_SET":       "Aktualizovat sadu",
        "DLG_DELETE_SET":       "Odstranit sadu",
        "DLG_RENAME_DOC":       "P\u0159ejmenovat dokument",
        "PROMPT_SET_NAME":      "N\u00e1zev sady:",
        "PROMPT_RENAME_SET":    "Nov\u00fd n\u00e1zev pro \u2018{old}\u2019:",
        "PROMPT_OVERWRITE":     "Vyberte sadu k p\u0159eps\u00e1n\u00ed:",
        "PROMPT_RENAME_WHICH":  "Vyberte sadu k p\u0159ejmenov\u00e1n\u00ed:",
        "PROMPT_DELETE_WHICH":  "Vyberte sadu k odstran\u011bn\u00ed:",
        "PROMPT_RENAME_DOC":    "Nov\u00fd n\u00e1zev pro \u2018{name}\u2019:",
        "ERR_NO_DOCS":          "\u017d\u00e1dn\u00e9 ulo\u017een\u00e9 dokumenty nejsou otev\u0159en\u00e9.\nNejprve ulo\u017ete sv\u00e9 dokumenty.",
        "ERR_NO_SETS":          "\u017d\u00e1dn\u00e9 ulo\u017een\u00e9 sady.",
        "ERR_NO_SETS_UPDATE":   "\u017d\u00e1dn\u00e9 sady k aktualizaci.",
        "ERR_NO_OPEN_DOCS":     "\u017d\u00e1dn\u00e9 ulo\u017een\u00e9 dokumenty nejsou otev\u0159en\u00e9.",
    },
    "hu": {
        "RENAME":               "\u00c1tnevez\u00e9s\u2026",
        "SAVE":                 "Ment\u00e9s",
        "SAVE_AS":              "Ment\u00e9s m\u00e1sk\u00e9nt\u2026",
        "MOVE_LEFT":            "Mozgat\u00e1s balra",
        "MOVE_RIGHT":           "Mozgat\u00e1s jobbra",
        "NEW_DOCUMENT":         "\u00daj dokumentum",
        "CLOSE":                "Bez\u00e1r\u00e1s",
        "CLOSE_ALL_OTHERS":     "T\u00f6bbi bez\u00e1r\u00e1sa",
        "CLOSE_ALL":            "Minden bez\u00e1r\u00e1sa",
        "SAVE_CURRENT_SET":     "Aktu\u00e1lis k\u00e9szlet ment\u00e9se\u2026",
        "UPDATE_A_SET":         "K\u00e9szlet friss\u00edt\u00e9se\u2026",
        "RENAME_A_SET":         "K\u00e9szlet \u00e1tnevez\u00e9se\u2026",
        "DELETE_A_SET":         "K\u00e9szlet t\u00f6rl\u00e9se\u2026",
        "RESTORE_LAST_SESSION": "Utols\u00f3 munkamenet visszajelz\u00e9se",
        "TAB_KEY_SWITCHING":    "V\u00e1lt\u00e1s Tab bill\u00enty\u0171vel",
        "SETS_BUTTON":          "\u2630 K\u00e9szletek",
        "DLG_SAVE_SET":         "F\u00fclek\u00e9szlet ment\u00e9se",
        "DLG_RENAME_SET":       "K\u00e9szlet \u00e1tnevez\u00e9se",
        "DLG_UPDATE_SET":       "K\u00e9szlet friss\u00edt\u00e9se",
        "DLG_DELETE_SET":       "K\u00e9szlet t\u00f6rl\u00e9se",
        "DLG_RENAME_DOC":       "Dokumentum \u00e1tnevez\u00e9se",
        "PROMPT_SET_NAME":      "K\u00e9szlet neve:",
        "PROMPT_RENAME_SET":    "\u2018{old}\u2019 \u00faj neve:",
        "PROMPT_OVERWRITE":     "V\u00e1lassza ki a fel\u00fcl\u00edrand\u00f3 k\u00e9szletet:",
        "PROMPT_RENAME_WHICH":  "V\u00e1lassza ki az \u00e1tnevezend\u0151 k\u00e9szletet:",
        "PROMPT_DELETE_WHICH":  "V\u00e1lassza ki a t\u00f6rlend\u0151 k\u00e9szletet:",
        "PROMPT_RENAME_DOC":    "\u2018{name}\u2019 \u00faj neve:",
        "ERR_NO_DOCS":          "Nincs megnyitott mentett dokumentum.\nEl\u0151sz\u00f6r mentse el a dokumentumokat.",
        "ERR_NO_SETS":          "Nincsenek mentett k\u00e9szletek.",
        "ERR_NO_SETS_UPDATE":   "Nincsenek friss\u00edtend\u0151 k\u00e9szletek.",
        "ERR_NO_OPEN_DOCS":     "Nincs megnyitott mentett dokumentum.",
    },
    "tr": {
        "RENAME":               "Yeniden adland\u0131r\u2026",
        "SAVE":                 "Kaydet",
        "SAVE_AS":              "Farkl\u0131 kaydet\u2026",
        "MOVE_LEFT":            "Sola ta\u015f\u0131",
        "MOVE_RIGHT":           "Sa\u011fa ta\u015f\u0131",
        "NEW_DOCUMENT":         "Yeni belge",
        "CLOSE":                "Kapat",
        "CLOSE_ALL_OTHERS":     "Di\u011ferlerini kapat",
        "CLOSE_ALL":            "T\u00fcm\u00fcn\u00fc kapat",
        "SAVE_CURRENT_SET":     "Ge\u00e7erli k\u00fcmeyi kaydet\u2026",
        "UPDATE_A_SET":         "K\u00fcmeyi g\u00fcncelle\u2026",
        "RENAME_A_SET":         "K\u00fcmeyi yeniden adland\u0131r\u2026",
        "DELETE_A_SET":         "K\u00fcmeyi sil\u2026",
        "RESTORE_LAST_SESSION": "Son oturumu geri y\u00fckle",
        "TAB_KEY_SWITCHING":    "Tab ile ge\u00e7i\u015f",
        "SETS_BUTTON":          "\u2630 K\u00fcmeler",
        "DLG_SAVE_SET":         "Sekme k\u00fcmesini kaydet",
        "DLG_RENAME_SET":       "K\u00fcmeyi yeniden adland\u0131r",
        "DLG_UPDATE_SET":       "K\u00fcmeyi g\u00fcncelle",
        "DLG_DELETE_SET":       "K\u00fcmeyi sil",
        "DLG_RENAME_DOC":       "Belgeyi yeniden adland\u0131r",
        "PROMPT_SET_NAME":      "K\u00fcme ad\u0131:",
        "PROMPT_RENAME_SET":    "\u2018{old}\u2019 i\u00e7in yeni ad:",
        "PROMPT_OVERWRITE":     "\u00dczerine yazd\u0131rmak i\u00e7in k\u00fcme se\u00e7in:",
        "PROMPT_RENAME_WHICH":  "Yeniden adland\u0131r\u0131lacak k\u00fcmeyi se\u00e7in:",
        "PROMPT_DELETE_WHICH":  "Silinecek k\u00fcmeyi se\u00e7in:",
        "PROMPT_RENAME_DOC":    "\u2018{name}\u2019 i\u00e7in yeni ad:",
        "ERR_NO_DOCS":          "A\u00e7\u0131k kaydedilmi\u015f belge yok.\n\u00d6nce belgelerinizi kaydedin.",
        "ERR_NO_SETS":          "Kaydedilmi\u015f k\u00fcme yok.",
        "ERR_NO_SETS_UPDATE":   "G\u00fcncellenecek k\u00fcme yok.",
        "ERR_NO_OPEN_DOCS":     "A\u00e7\u0131k kaydedilmi\u015f belge yok.",
    },
    "zh-cn": {
        "RENAME":               "\u91cd\u547d\u540d\u2026",
        "SAVE":                 "\u4fdd\u5b58",
        "SAVE_AS":              "\u53e6\u5b58\u4e3a\u2026",
        "MOVE_LEFT":            "\u5411\u5de6\u79fb\u52a8",
        "MOVE_RIGHT":           "\u5411\u53f3\u79fb\u52a8",
        "NEW_DOCUMENT":         "\u65b0\u5efa\u6587\u6863",
        "CLOSE":                "\u5173\u95ed",
        "CLOSE_ALL_OTHERS":     "\u5173\u95ed\u5176\u4ed6",
        "CLOSE_ALL":            "\u5168\u90e8\u5173\u95ed",
        "SAVE_CURRENT_SET":     "\u4fdd\u5b58\u5f53\u524d\u7ec4\u2026",
        "UPDATE_A_SET":         "\u66f4\u65b0\u7ec4\u2026",
        "RENAME_A_SET":         "\u91cd\u547d\u540d\u7ec4\u2026",
        "DELETE_A_SET":         "\u5220\u9664\u7ec4\u2026",
        "RESTORE_LAST_SESSION": "\u6062\u590d\u4e0a\u6b21\u4f1a\u8bdd",
        "TAB_KEY_SWITCHING":    "Tab \u952e\u5207\u6362",
        "SETS_BUTTON":          "\u2630 \u7ec4",
        "DLG_SAVE_SET":         "\u4fdd\u5b58\u6807\u7b7e\u7ec4",
        "DLG_RENAME_SET":       "\u91cd\u547d\u540d\u7ec4",
        "DLG_UPDATE_SET":       "\u66f4\u65b0\u7ec4",
        "DLG_DELETE_SET":       "\u5220\u9664\u7ec4",
        "DLG_RENAME_DOC":       "\u91cd\u547d\u540d\u6587\u6863",
        "PROMPT_SET_NAME":      "\u7ec4\u540d\u79f0\uff1a",
        "PROMPT_RENAME_SET":    "\u2018{old}\u2019 \u7684\u65b0\u540d\u79f0\uff1a",
        "PROMPT_OVERWRITE":     "\u9009\u62e9\u8986\u76d6\u7684\u7ec4\uff1a",
        "PROMPT_RENAME_WHICH":  "\u9009\u62e9\u8981\u91cd\u547d\u540d\u7684\u7ec4\uff1a",
        "PROMPT_DELETE_WHICH":  "\u9009\u62e9\u8981\u5220\u9664\u7684\u7ec4\uff1a",
        "PROMPT_RENAME_DOC":    "\u2018{name}\u2019 \u7684\u65b0\u540d\u79f0\uff1a",
        "ERR_NO_DOCS":          "\u6ca1\u6709\u5df2\u4fdd\u5b58\u7684\u6587\u6863\u3002\n\u8bf7\u5148\u4fdd\u5b58\u60a8\u7684\u6587\u6863\u3002",
        "ERR_NO_SETS":          "\u6ca1\u6709\u5df2\u4fdd\u5b58\u7684\u7ec4\u3002",
        "ERR_NO_SETS_UPDATE":   "\u6ca1\u6709\u53ef\u66f4\u65b0\u7684\u7ec4\u3002",
        "ERR_NO_OPEN_DOCS":     "\u6ca1\u6709\u5df2\u4fdd\u5b58\u7684\u6587\u6863\u3002",
    },
    "zh-tw": {
        "RENAME":               "\u91cd\u65b0\u547d\u540d\u2026",
        "SAVE":                 "\u5132\u5b58",
        "SAVE_AS":              "\u53e6\u5b58\u65b0\u6a94\u2026",
        "MOVE_LEFT":            "\u5411\u5de6\u79fb\u52d5",
        "MOVE_RIGHT":           "\u5411\u53f3\u79fb\u52d5",
        "NEW_DOCUMENT":         "\u65b0\u5efa\u6587\u4ef6",
        "CLOSE":                "\u95dc\u9589",
        "CLOSE_ALL_OTHERS":     "\u95dc\u9589\u5176\u4ed6",
        "CLOSE_ALL":            "\u5168\u90e8\u95dc\u9589",
        "SAVE_CURRENT_SET":     "\u5132\u5b58\u76ee\u524d\u7d44\u2026",
        "UPDATE_A_SET":         "\u66f4\u65b0\u7d44\u2026",
        "RENAME_A_SET":         "\u91cd\u65b0\u547d\u540d\u7d44\u2026",
        "DELETE_A_SET":         "\u522a\u9664\u7d44\u2026",
        "RESTORE_LAST_SESSION": "\u6062\u5fa9\u4e0a\u6b21\u4f5c\u696d\u968e\u6bb5",
        "TAB_KEY_SWITCHING":    "Tab \u9375\u5207\u63db",
        "SETS_BUTTON":          "\u2630 \u7d44",
        "DLG_SAVE_SET":         "\u5132\u5b58\u6a19\u7c64\u7d44",
        "DLG_RENAME_SET":       "\u91cd\u65b0\u547d\u540d\u7d44",
        "DLG_UPDATE_SET":       "\u66f4\u65b0\u7d44",
        "DLG_DELETE_SET":       "\u522a\u9664\u7d44",
        "DLG_RENAME_DOC":       "\u91cd\u65b0\u547d\u540d\u6587\u4ef6",
        "PROMPT_SET_NAME":      "\u7d44\u540d\u7a31\uff1a",
        "PROMPT_RENAME_SET":    "\u2018{old}\u2019 \u7684\u65b0\u540d\u7a31\uff1a",
        "PROMPT_OVERWRITE":     "\u9078\u64c7\u8981\u8986\u84cb\u7684\u7d44\uff1a",
        "PROMPT_RENAME_WHICH":  "\u9078\u64c7\u8981\u91cd\u65b0\u547d\u540d\u7684\u7d44\uff1a",
        "PROMPT_DELETE_WHICH":  "\u9078\u64c7\u8981\u522a\u9664\u7684\u7d44\uff1a",
        "PROMPT_RENAME_DOC":    "\u2018{name}\u2019 \u7684\u65b0\u540d\u7a31\uff1a",
        "ERR_NO_DOCS":          "\u6c92\u6709\u5df2\u5132\u5b58\u7684\u6587\u4ef6\u3002\n\u8acb\u5148\u5132\u5b58\u60a8\u7684\u6587\u4ef6\u3002",
        "ERR_NO_SETS":          "\u6c92\u6709\u5df2\u5132\u5b58\u7684\u7d44\u3002",
        "ERR_NO_SETS_UPDATE":   "\u6c92\u6709\u53ef\u66f4\u65b0\u7684\u7d44\u3002",
        "ERR_NO_OPEN_DOCS":     "\u6c92\u6709\u5df2\u5132\u5b58\u7684\u6587\u4ef6\u3002",
    },
    "ja": {
        "RENAME":               "\u540d\u524d\u5909\u66f4\u2026",
        "SAVE":                 "\u4fdd\u5b58",
        "SAVE_AS":              "\u540d\u524d\u3092\u4ed8\u3051\u3066\u4fdd\u5b58\u2026",
        "MOVE_LEFT":            "\u5de6\u3078\u79fb\u52d5",
        "MOVE_RIGHT":           "\u53f3\u3078\u79fb\u52d5",
        "NEW_DOCUMENT":         "\u65b0\u898f\u30c9\u30ad\u30e5\u30e1\u30f3\u30c8",
        "CLOSE":                "\u9589\u3058\u308b",
        "CLOSE_ALL_OTHERS":     "\u4ed6\u3092\u9589\u3058\u308b",
        "CLOSE_ALL":            "\u3059\u3079\u3066\u9589\u3058\u308b",
        "SAVE_CURRENT_SET":     "\u73fe\u5728\u306e\u30bb\u30c3\u30c8\u3092\u4fdd\u5b58\u2026",
        "UPDATE_A_SET":         "\u30bb\u30c3\u30c8\u3092\u66f4\u65b0\u2026",
        "RENAME_A_SET":         "\u30bb\u30c3\u30c8\u3092\u5909\u540d\u2026",
        "DELETE_A_SET":         "\u30bb\u30c3\u30c8\u3092\u524a\u9664\u2026",
        "RESTORE_LAST_SESSION": "\u524d\u56de\u306e\u30bb\u30c3\u30b7\u30e7\u30f3\u3092\u5fa9\u5143",
        "TAB_KEY_SWITCHING":    "Tab\u30ad\u30fc\u3067\u5207\u308a\u66ff\u3048",
        "SETS_BUTTON":          "\u2630 \u30bb\u30c3\u30c8",
        "DLG_SAVE_SET":         "\u30bf\u30d6\u30bb\u30c3\u30c8\u3092\u4fdd\u5b58",
        "DLG_RENAME_SET":       "\u30bb\u30c3\u30c8\u3092\u5909\u540d",
        "DLG_UPDATE_SET":       "\u30bb\u30c3\u30c8\u3092\u66f4\u65b0",
        "DLG_DELETE_SET":       "\u30bb\u30c3\u30c8\u3092\u524a\u9664",
        "DLG_RENAME_DOC":       "\u30c9\u30ad\u30e5\u30e1\u30f3\u30c8\u3092\u5909\u540d",
        "PROMPT_SET_NAME":      "\u30bb\u30c3\u30c8\u540d\uff1a",
        "PROMPT_RENAME_SET":    "\u2018{old}\u2019 \u306e\u65b0\u3057\u3044\u540d\u524d\uff1a",
        "PROMPT_OVERWRITE":     "\u4e0a\u66f8\u304d\u3059\u308b\u30bb\u30c3\u30c8\u3092\u9078\u629e\uff1a",
        "PROMPT_RENAME_WHICH":  "\u5909\u540d\u3059\u308b\u30bb\u30c3\u30c8\u3092\u9078\u629e\uff1a",
        "PROMPT_DELETE_WHICH":  "\u524a\u9664\u3059\u308b\u30bb\u30c3\u30c8\u3092\u9078\u629e\uff1a",
        "PROMPT_RENAME_DOC":    "\u2018{name}\u2019 \u306e\u65b0\u3057\u3044\u540d\u524d\uff1a",
        "ERR_NO_DOCS":          "\u4fdd\u5b58\u6e08\u307f\u306e\u30c9\u30ad\u30e5\u30e1\u30f3\u30c8\u304c\u3042\u308a\u307e\u305b\u3093\u3002\n\u307e\u305a\u30c9\u30ad\u30e5\u30e1\u30f3\u30c8\u3092\u4fdd\u5b58\u3057\u3066\u304f\u3060\u3055\u3044\u3002",
        "ERR_NO_SETS":          "\u4fdd\u5b58\u6e08\u307f\u306e\u30bb\u30c3\u30c8\u304c\u3042\u308a\u307e\u305b\u3093\u3002",
        "ERR_NO_SETS_UPDATE":   "\u66f4\u65b0\u3059\u308b\u30bb\u30c3\u30c8\u304c\u3042\u308a\u307e\u305b\u3093\u3002",
        "ERR_NO_OPEN_DOCS":     "\u4fdd\u5b58\u6e08\u307f\u306e\u30c9\u30ad\u30e5\u30e1\u30f3\u30c8\u304c\u3042\u308a\u307e\u305b\u3093\u3002",
    },
    "ko": {
        "RENAME":               "\u0438\ub984 \ubc14\uafb8\uae30\u2026",
        "SAVE":                 "\uc800\uc7a5",
        "SAVE_AS":              "\ub2e4\ub978 \uc774\ub984\uc73c\ub85c \uc800\uc7a5\u2026",
        "MOVE_LEFT":            "\uc67c\ucabd\uc73c\ub85c \uc774\ub3d9",
        "MOVE_RIGHT":           "\uc624\ub978\ucabd\uc73c\ub85c \uc774\ub3d9",
        "NEW_DOCUMENT":         "\uc0c8 \ubb38\uc11c",
        "CLOSE":                "\ub2eb\uae30",
        "CLOSE_ALL_OTHERS":     "\ub2e4\ub978 \uac83 \ub2eb\uae30",
        "CLOSE_ALL":            "\ubaa8\ub450 \ub2eb\uae30",
        "SAVE_CURRENT_SET":     "\ud604\uc7ac \uc138\ud2b8 \uc800\uc7a5\u2026",
        "UPDATE_A_SET":         "\uc138\ud2b8 \uc5c5\ub370\uc774\ud2b8\u2026",
        "RENAME_A_SET":         "\uc138\ud2b8 \uc774\ub984 \ubc14\uafb8\uae30\u2026",
        "DELETE_A_SET":         "\uc138\ud2b8 \uc0ad\uc81c\u2026",
        "RESTORE_LAST_SESSION": "\ub9c8\uc9c0\ub9c9 \uc138\uc158 \ubcf5\uc6d0",
        "TAB_KEY_SWITCHING":    "Tab \ud0a4\ub85c \uc804\ud658",
        "SETS_BUTTON":          "\u2630 \uc138\ud2b8",
        "DLG_SAVE_SET":         "\ud0ed \uc138\ud2b8 \uc800\uc7a5",
        "DLG_RENAME_SET":       "\uc138\ud2b8 \uc774\ub984 \ubc14\uafb8\uae30",
        "DLG_UPDATE_SET":       "\uc138\ud2b8 \uc5c5\ub370\uc774\ud2b8",
        "DLG_DELETE_SET":       "\uc138\ud2b8 \uc0ad\uc81c",
        "DLG_RENAME_DOC":       "\ubb38\uc11c \uc774\ub984 \ubc14\uafb8\uae30",
        "PROMPT_SET_NAME":      "\uc138\ud2b8 \uc774\ub984:",
        "PROMPT_RENAME_SET":    "\u2018{old}\u2019\uc758 \uc0c8 \uc774\ub984:",
        "PROMPT_OVERWRITE":     "\uacb9\uccd0 \uc4f8 \uc138\ud2b8 \uc120\ud0dd:",
        "PROMPT_RENAME_WHICH":  "\uc774\ub984\uc744 \ubc14\uafb8\uc635 \uc138\ud2b8 \uc120\ud0dd:",
        "PROMPT_DELETE_WHICH":  "\uc0ad\uc81c\ud560 \uc138\ud2b8 \uc120\ud0dd:",
        "PROMPT_RENAME_DOC":    "\u2018{name}\u2019\uc758 \uc0c8 \uc774\ub984:",
        "ERR_NO_DOCS":          "\uc800\uc7a5\ub41c \ubb38\uc11c\uac00 \uc5c6\uc2b5\ub2c8\ub2e4.\n\uba3c\uc800 \ubb38\uc11c\ub97c \uc800\uc7a5\ud558\uc138\uc694.",
        "ERR_NO_SETS":          "\uc800\uc7a5\ub41c \uc138\ud2b8\uac00 \uc5c6\uc2b5\ub2c8\ub2e4.",
        "ERR_NO_SETS_UPDATE":   "\uc5c5\ub370\uc774\ud2b8\ud560 \uc138\ud2b8\uac00 \uc5c6\uc2b5\ub2c8\ub2e4.",
        "ERR_NO_OPEN_DOCS":     "\uc800\uc7a5\ub41c \ubb38\uc11c\uac00 \uc5c6\uc2b5\ub2c8\ub2e4.",
    },
    "sk": {
        "RENAME":               "Premenova\u0165\u2026",
        "SAVE":                 "Ulo\u017ei\u0165",
        "SAVE_AS":              "Ulo\u017ei\u0165 ako\u2026",
        "MOVE_LEFT":            "Presun\u00fa\u0165 do\u013eava",
        "MOVE_RIGHT":           "Presun\u00fa\u0165 doprava",
        "NEW_DOCUMENT":         "Nov\u00fd dokument",
        "CLOSE":                "Zavrie\u0165",
        "CLOSE_ALL_OTHERS":     "Zavrie\u0165 ostatn\u00e9",
        "CLOSE_ALL":            "Zavrie\u0165 v\u0161etky",
        "SAVE_CURRENT_SET":     "Ulo\u017ei\u0165 aktu\u00e1lnu sadu\u2026",
        "UPDATE_A_SET":         "Aktualizova\u0165 sadu\u2026",
        "RENAME_A_SET":         "Premenova\u0165 sadu\u2026",
        "DELETE_A_SET":         "Odstr\u00e1ni\u0165 sadu\u2026",
        "RESTORE_LAST_SESSION": "Obnovi\u0165 posledn\u00fa rel\u00e1ciu",
        "TAB_KEY_SWITCHING":    "Prep\u00ednanie kl\u00e1vesom Tab",
        "SETS_BUTTON":          "\u2630 Sady",
        "DLG_SAVE_SET":         "Ulo\u017ei\u0165 sadu kariet",
        "DLG_RENAME_SET":       "Premenova\u0165 sadu",
        "DLG_UPDATE_SET":       "Aktualizova\u0165 sadu",
        "DLG_DELETE_SET":       "Odstr\u00e1ni\u0165 sadu",
        "DLG_RENAME_DOC":       "Premenova\u0165 dokument",
        "PROMPT_SET_NAME":      "N\u00e1zov sady:",
        "PROMPT_RENAME_SET":    "Nov\u00fd n\u00e1zov pre \u2018{old}\u2019:",
        "PROMPT_OVERWRITE":     "Vybra\u0165 sadu na prep\u00edsanie:",
        "PROMPT_RENAME_WHICH":  "Vybra\u0165 sadu na premenovanie:",
        "PROMPT_DELETE_WHICH":  "Vybra\u0165 sadu na odstr\u00e1nenie:",
        "PROMPT_RENAME_DOC":    "Nov\u00fd n\u00e1zov pre \u2018{name}\u2019:",
        "ERR_NO_DOCS":          "\u017diadne ulo\u017een\u00e9 dokumenty nie s\u00fa otvor\u00e9n\u00e9.\nNajprv ulo\u017ete svoje dokumenty.",
        "ERR_NO_SETS":          "\u017diadne ulo\u017een\u00e9 sady.",
        "ERR_NO_SETS_UPDATE":   "\u017diadne sady na aktualiz\u00e1ciu.",
        "ERR_NO_OPEN_DOCS":     "\u017diadne ulo\u017een\u00e9 dokumenty nie s\u00fa otvor\u00e9n\u00e9.",
    },
    "ro": {
        "RENAME":               "Redenumire\u2026",
        "SAVE":                 "Salvare",
        "SAVE_AS":              "Salvare ca\u2026",
        "MOVE_LEFT":            "Mut\u0103 la st\u00e2nga",
        "MOVE_RIGHT":           "Mut\u0103 la dreapta",
        "NEW_DOCUMENT":         "Document nou",
        "CLOSE":                "\u00cenchide",
        "CLOSE_ALL_OTHERS":     "\u00cenchide celelalte",
        "CLOSE_ALL":            "\u00cenchide tot",
        "SAVE_CURRENT_SET":     "Salveaz\u0103 setul curent\u2026",
        "UPDATE_A_SET":         "Actualizeaz\u0103 un set\u2026",
        "RENAME_A_SET":         "Redenume\u0219te un set\u2026",
        "DELETE_A_SET":         "\u015eterge un set\u2026",
        "RESTORE_LAST_SESSION": "Restaureaz\u0103 ultima sesiune",
        "TAB_KEY_SWITCHING":    "Comutare cu Tab",
        "SETS_BUTTON":          "\u2630 Seturi",
        "DLG_SAVE_SET":         "Salvare set file",
        "DLG_RENAME_SET":       "Redenumire set",
        "DLG_UPDATE_SET":       "Actualizare set",
        "DLG_DELETE_SET":       "\u015etergere set",
        "DLG_RENAME_DOC":       "Redenumire document",
        "PROMPT_SET_NAME":      "Numele setului:",
        "PROMPT_RENAME_SET":    "Nume nou pentru \u2018{old}\u2019:",
        "PROMPT_OVERWRITE":     "Alege\u021bi setul de suprascris:",
        "PROMPT_RENAME_WHICH":  "Alege\u021bi setul de redenumit:",
        "PROMPT_DELETE_WHICH":  "Alege\u021bi setul de \u015ters:",
        "PROMPT_RENAME_DOC":    "Nume nou pentru \u2018{name}\u2019:",
        "ERR_NO_DOCS":          "Niciun document salvat nu este deschis.\nSalva\u021bi mai \u00ent\u00e2i documentele.",
        "ERR_NO_SETS":          "Niciun set salvat.",
        "ERR_NO_SETS_UPDATE":   "Niciun set de actualizat.",
        "ERR_NO_OPEN_DOCS":     "Niciun document salvat nu este deschis.",
    },
    "el": {
        "RENAME":               "\u039c\u03b5\u03c4\u03bf\u03bd\u03bf\u03bc\u03b1\u03c3\u03af\u03b1\u2026",
        "SAVE":                 "\u0391\u03c0\u03bf\u03b8\u03ae\u03ba\u03b5\u03c5\u03c3\u03b7",
        "SAVE_AS":              "\u0391\u03c0\u03bf\u03b8\u03ae\u03ba\u03b5\u03c5\u03c3\u03b7 \u03c9\u03c2\u2026",
        "MOVE_LEFT":            "\u039c\u03b5\u03c4\u03b1\u03ba\u03af\u03bd\u03b7\u03c3\u03b7 \u03b1\u03c1\u03b9\u03c3\u03c4\u03b5\u03c1\u03ac",
        "MOVE_RIGHT":           "\u039c\u03b5\u03c4\u03b1\u03ba\u03af\u03bd\u03b7\u03c3\u03b7 \u03b4\u03b5\u03be\u03b9\u03ac",
        "NEW_DOCUMENT":         "\u039d\u03ad\u03bf \u03ad\u03b3\u03b3\u03c1\u03b1\u03c6\u03bf",
        "CLOSE":                "\u039a\u03bb\u03b5\u03af\u03c3\u03b9\u03bc\u03bf",
        "CLOSE_ALL_OTHERS":     "\u039a\u03bb\u03b5\u03af\u03c3\u03b9\u03bc\u03bf \u03ac\u03bb\u03bb\u03c9\u03bd",
        "CLOSE_ALL":            "\u039a\u03bb\u03b5\u03af\u03c3\u03b9\u03bc\u03bf \u03cc\u03bb\u03c9\u03bd",
        "SAVE_CURRENT_SET":     "\u0391\u03c0\u03bf\u03b8\u03ae\u03ba\u03b5\u03c5\u03c3\u03b7 \u03c4\u03c1\u03ad\u03c7\u03bf\u03c5\u03c3\u03b1\u03c2 \u03bf\u03bc\u03ac\u03b4\u03b1\u03c2\u2026",
        "UPDATE_A_SET":         "\u0395\u03bd\u03b7\u03bc\u03ad\u03c1\u03c9\u03c3\u03b7 \u03bf\u03bc\u03ac\u03b4\u03b1\u03c2\u2026",
        "RENAME_A_SET":         "\u039c\u03b5\u03c4\u03bf\u03bd\u03bf\u03bc\u03b1\u03c3\u03af\u03b1 \u03bf\u03bc\u03ac\u03b4\u03b1\u03c2\u2026",
        "DELETE_A_SET":         "\u0394\u03b9\u03b1\u03b3\u03c1\u03b1\u03c6\u03ae \u03bf\u03bc\u03ac\u03b4\u03b1\u03c2\u2026",
        "RESTORE_LAST_SESSION": "\u0395\u03c0\u03b1\u03bd\u03b1\u03c6\u03bf\u03c1\u03ac \u03c4\u03b5\u03bb\u03b5\u03c5\u03c4\u03b1\u03af\u03b1\u03c2 \u03c3\u03c5\u03bd\u03b5\u03b4\u03c1\u03af\u03b1\u03c2",
        "TAB_KEY_SWITCHING":    "\u0395\u03bd\u03b1\u03bb\u03bb\u03b1\u03b3\u03ae \u03bc\u03b5 Tab",
        "SETS_BUTTON":          "\u2630 \u039f\u03bc\u03ac\u03b4\u03b5\u03c2",
        "DLG_SAVE_SET":         "\u0391\u03c0\u03bf\u03b8\u03ae\u03ba\u03b5\u03c5\u03c3\u03b7 \u03bf\u03bc\u03ac\u03b4\u03b1\u03c2 \u03ba\u03b1\u03c1\u03c4\u03b5\u03bb\u03ce\u03bd",
        "DLG_RENAME_SET":       "\u039c\u03b5\u03c4\u03bf\u03bd\u03bf\u03bc\u03b1\u03c3\u03af\u03b1 \u03bf\u03bc\u03ac\u03b4\u03b1\u03c2",
        "DLG_UPDATE_SET":       "\u0395\u03bd\u03b7\u03bc\u03ad\u03c1\u03c9\u03c3\u03b7 \u03bf\u03bc\u03ac\u03b4\u03b1\u03c2",
        "DLG_DELETE_SET":       "\u0394\u03b9\u03b1\u03b3\u03c1\u03b1\u03c6\u03ae \u03bf\u03bc\u03ac\u03b4\u03b1\u03c2",
        "DLG_RENAME_DOC":       "\u039c\u03b5\u03c4\u03bf\u03bd\u03bf\u03bc\u03b1\u03c3\u03af\u03b1 \u03b5\u03b3\u03b3\u03c1\u03ac\u03c6\u03bf\u03c5",
        "PROMPT_SET_NAME":      "\u038c\u03bd\u03bf\u03bc\u03b1 \u03bf\u03bc\u03ac\u03b4\u03b1\u03c2:",
        "PROMPT_RENAME_SET":    "\u039d\u03ad\u03bf \u03cc\u03bd\u03bf\u03bc\u03b1 \u03b3\u03b9\u03b1 \u03c4\u03bf \u2018{old}\u2019:",
        "PROMPT_OVERWRITE":     "\u0395\u03c0\u03b9\u03bb\u03bf\u03b3\u03ae \u03bf\u03bc\u03ac\u03b4\u03b1\u03c2 \u03b3\u03b9\u03b1 \u03b1\u03bd\u03c4\u03b9\u03ba\u03b1\u03c4\u03ac\u03c3\u03c4\u03b1\u03c3\u03b7:",
        "PROMPT_RENAME_WHICH":  "\u0395\u03c0\u03b9\u03bb\u03bf\u03b3\u03ae \u03bf\u03bc\u03ac\u03b4\u03b1\u03c2 \u03b3\u03b9\u03b1 \u03bc\u03b5\u03c4\u03bf\u03bd\u03bf\u03bc\u03b1\u03c3\u03af\u03b1:",
        "PROMPT_DELETE_WHICH":  "\u0395\u03c0\u03b9\u03bb\u03bf\u03b3\u03ae \u03bf\u03bc\u03ac\u03b4\u03b1\u03c2 \u03b3\u03b9\u03b1 \u03b4\u03b9\u03b1\u03b3\u03c1\u03b1\u03c6\u03ae:",
        "PROMPT_RENAME_DOC":    "\u039d\u03ad\u03bf \u03cc\u03bd\u03bf\u03bc\u03b1 \u03b3\u03b9\u03b1 \u03c4\u03bf \u2018{name}\u2019:",
        "ERR_NO_DOCS":          "\u0394\u03b5\u03bd \u03c5\u03c0\u03ac\u03c1\u03c7\u03bf\u03c5\u03bd \u03b1\u03bd\u03bf\u03b9\u03c7\u03c4\u03ac \u03b1\u03c0\u03bf\u03b8\u03b7\u03ba\u03b5\u03c5\u03bc\u03ad\u03bd\u03b1 \u03ad\u03b3\u03b3\u03c1\u03b1\u03c6\u03b1.\n\u0391\u03c0\u03bf\u03b8\u03b7\u03ba\u03b5\u03cd\u03c3\u03c4\u03b5 \u03c0\u03c1\u03ce\u03c4\u03b1 \u03c4\u03b1 \u03ad\u03b3\u03b3\u03c1\u03b1\u03c6\u03ac \u03c3\u03b1\u03c2.",
        "ERR_NO_SETS":          "\u0394\u03b5\u03bd \u03c5\u03c0\u03ac\u03c1\u03c7\u03bf\u03c5\u03bd \u03b1\u03c0\u03bf\u03b8\u03b7\u03ba\u03b5\u03c5\u03bc\u03ad\u03bd\u03b5\u03c2 \u03bf\u03bc\u03ac\u03b4\u03b5\u03c2.",
        "ERR_NO_SETS_UPDATE":   "\u0394\u03b5\u03bd \u03c5\u03c0\u03ac\u03c1\u03c7\u03bf\u03c5\u03bd \u03bf\u03bc\u03ac\u03b4\u03b5\u03c2 \u03b3\u03b9\u03b1 \u03b5\u03bd\u03b7\u03bc\u03ad\u03c1\u03c9\u03c3\u03b7.",
        "ERR_NO_OPEN_DOCS":     "\u0394\u03b5\u03bd \u03c5\u03c0\u03ac\u03c1\u03c7\u03bf\u03c5\u03bd \u03b1\u03bd\u03bf\u03b9\u03c7\u03c4\u03ac \u03b1\u03c0\u03bf\u03b8\u03b7\u03ba\u03b5\u03c5\u03bc\u03ad\u03bd\u03b1 \u03ad\u03b3\u03b3\u03c1\u03b1\u03c6\u03b1.",
    },
    "uk": {
        "RENAME":               "\u041f\u0435\u0440\u0435\u0439\u043c\u0435\u043d\u0443\u0432\u0430\u0442\u0438\u2026",
        "SAVE":                 "\u0417\u0431\u0435\u0440\u0435\u0433\u0442\u0438",
        "SAVE_AS":              "\u0417\u0431\u0435\u0440\u0435\u0433\u0442\u0438 \u044f\u043a\u2026",
        "MOVE_LEFT":            "\u041f\u0435\u0440\u0435\u043c\u0456\u0441\u0442\u0438\u0442\u0438 \u043b\u0456\u0432\u043e\u0440\u0443\u0447",
        "MOVE_RIGHT":           "\u041f\u0435\u0440\u0435\u043c\u0456\u0441\u0442\u0438\u0442\u0438 \u043f\u0440\u0430\u0432\u043e\u0440\u0443\u0447",
        "NEW_DOCUMENT":         "\u041d\u043e\u0432\u0438\u0439 \u0434\u043e\u043a\u0443\u043c\u0435\u043d\u0442",
        "CLOSE":                "\u0417\u0430\u043a\u0440\u0438\u0442\u0438",
        "CLOSE_ALL_OTHERS":     "\u0417\u0430\u043a\u0440\u0438\u0442\u0438 \u0456\u043d\u0448\u0456",
        "CLOSE_ALL":            "\u0417\u0430\u043a\u0440\u0438\u0442\u0438 \u0432\u0441\u0456",
        "SAVE_CURRENT_SET":     "\u0417\u0431\u0435\u0440\u0435\u0433\u0442\u0438 \u043f\u043e\u0442\u043e\u0447\u043d\u0438\u0439 \u043d\u0430\u0431\u0456\u0440\u2026",
        "UPDATE_A_SET":         "\u041e\u043d\u043e\u0432\u0438\u0442\u0438 \u043d\u0430\u0431\u0456\u0440\u2026",
        "RENAME_A_SET":         "\u041f\u0435\u0440\u0435\u0439\u043c\u0435\u043d\u0443\u0432\u0430\u0442\u0438 \u043d\u0430\u0431\u0456\u0440\u2026",
        "DELETE_A_SET":         "\u0412\u0438\u0434\u0430\u043b\u0438\u0442\u0438 \u043d\u0430\u0431\u0456\u0440\u2026",
        "RESTORE_LAST_SESSION": "\u0412\u0456\u0434\u043d\u043e\u0432\u0438\u0442\u0438 \u043e\u0441\u0442\u0430\u043d\u043d\u044e \u0441\u0435\u0441\u0456\u044e",
        "TAB_KEY_SWITCHING":    "\u041f\u0435\u0440\u0435\u043c\u0438\u043a\u0430\u043d\u043d\u044f Tab",
        "SETS_BUTTON":          "\u2630 \u041d\u0430\u0431\u043e\u0440\u0438",
        "DLG_SAVE_SET":         "\u0417\u0431\u0435\u0440\u0435\u0433\u0442\u0438 \u043d\u0430\u0431\u0456\u0440 \u0432\u043a\u043b\u0430\u0434\u043e\u043a",
        "DLG_RENAME_SET":       "\u041f\u0435\u0440\u0435\u0439\u043c\u0435\u043d\u0443\u0432\u0430\u0442\u0438 \u043d\u0430\u0431\u0456\u0440",
        "DLG_UPDATE_SET":       "\u041e\u043d\u043e\u0432\u0438\u0442\u0438 \u043d\u0430\u0431\u0456\u0440",
        "DLG_DELETE_SET":       "\u0412\u0438\u0434\u0430\u043b\u0438\u0442\u0438 \u043d\u0430\u0431\u0456\u0440",
        "DLG_RENAME_DOC":       "\u041f\u0435\u0440\u0435\u0439\u043c\u0435\u043d\u0443\u0432\u0430\u0442\u0438 \u0434\u043e\u043a\u0443\u043c\u0435\u043d\u0442",
        "PROMPT_SET_NAME":      "\u041d\u0430\u0437\u0432\u0430 \u043d\u0430\u0431\u043e\u0440\u0443:",
        "PROMPT_RENAME_SET":    "\u041d\u043e\u0432\u0430 \u043d\u0430\u0437\u0432\u0430 \u0434\u043b\u044f \u00ab{old}\u00bb:",
        "PROMPT_OVERWRITE":     "\u0412\u0438\u0431\u0435\u0440\u0456\u0442\u044c \u043d\u0430\u0431\u0456\u0440 \u0434\u043b\u044f \u043f\u0435\u0440\u0435\u0437\u0430\u043f\u0438\u0441\u0443:",
        "PROMPT_RENAME_WHICH":  "\u0412\u0438\u0431\u0435\u0440\u0456\u0442\u044c \u043d\u0430\u0431\u0456\u0440 \u0434\u043b\u044f \u043f\u0435\u0440\u0435\u0439\u043c\u0435\u043d\u0443\u0432\u0430\u043d\u043d\u044f:",
        "PROMPT_DELETE_WHICH":  "\u0412\u0438\u0431\u0435\u0440\u0456\u0442\u044c \u043d\u0430\u0431\u0456\u0440 \u0434\u043b\u044f \u0432\u0438\u0434\u0430\u043b\u0435\u043d\u043d\u044f:",
        "PROMPT_RENAME_DOC":    "\u041d\u043e\u0432\u0430 \u043d\u0430\u0437\u0432\u0430 \u0434\u043b\u044f \u00ab{name}\u00bb:",
        "ERR_NO_DOCS":          "\u041d\u0435\u043c\u0430\u0454 \u0432\u0456\u0434\u043a\u0440\u0438\u0442\u0438\u0445 \u0437\u0431\u0435\u0440\u0435\u0436\u0435\u043d\u0438\u0445 \u0434\u043e\u043a\u0443\u043c\u0435\u043d\u0442\u0456\u0432.\n\u0421\u043f\u043e\u0447\u0430\u0442\u043a\u0443 \u0437\u0431\u0435\u0440\u0435\u0436\u0456\u0442\u044c \u0434\u043e\u043a\u0443\u043c\u0435\u043d\u0442\u0438.",
        "ERR_NO_SETS":          "\u041d\u0435\u043c\u0430\u0454 \u0437\u0431\u0435\u0440\u0435\u0436\u0435\u043d\u0438\u0445 \u043d\u0430\u0431\u043e\u0440\u0456\u0432.",
        "ERR_NO_SETS_UPDATE":   "\u041d\u0435\u043c\u0430\u0454 \u043d\u0430\u0431\u043e\u0440\u0456\u0432 \u0434\u043b\u044f \u043e\u043d\u043e\u0432\u043b\u0435\u043d\u043d\u044f.",
        "ERR_NO_OPEN_DOCS":     "\u041d\u0435\u043c\u0430\u0454 \u0432\u0456\u0434\u043a\u0440\u0438\u0442\u0438\u0445 \u0437\u0431\u0435\u0440\u0435\u0436\u0435\u043d\u0438\u0445 \u0434\u043e\u043a\u0443\u043c\u0435\u043d\u0442\u0456\u0432.",
    },
    "bg": {
        "RENAME":               "\u041f\u0440\u0435\u0438\u043c\u0435\u043d\u0443\u0432\u0430\u043d\u0435\u2026",
        "SAVE":                 "\u0417\u0430\u043f\u0438\u0441\u0432\u0430\u043d\u0435",
        "SAVE_AS":              "\u0417\u0430\u043f\u0438\u0441\u0432\u0430\u043d\u0435 \u043a\u0430\u0442\u043e\u2026",
        "MOVE_LEFT":            "\u041f\u0440\u0435\u043c\u0435\u0441\u0442\u0432\u0430\u043d\u0435 \u043d\u0430\u043b\u044f\u0432\u043e",
        "MOVE_RIGHT":           "\u041f\u0440\u0435\u043c\u0435\u0441\u0442\u0432\u0430\u043d\u0435 \u043d\u0430\u0434\u044f\u0441\u043d\u043e",
        "NEW_DOCUMENT":         "\u041d\u043e\u0432 \u0434\u043e\u043a\u0443\u043c\u0435\u043d\u0442",
        "CLOSE":                "\u0417\u0430\u0442\u0432\u0430\u0440\u044f\u043d\u0435",
        "CLOSE_ALL_OTHERS":     "\u0417\u0430\u0442\u0432\u0430\u0440\u044f\u043d\u0435 \u043d\u0430 \u043e\u0441\u0442\u0430\u043d\u0430\u043b\u0438\u0442\u0435",
        "CLOSE_ALL":            "\u0417\u0430\u0442\u0432\u0430\u0440\u044f\u043d\u0435 \u043d\u0430 \u0432\u0441\u0438\u0447\u043a\u0438",
        "SAVE_CURRENT_SET":     "\u0417\u0430\u043f\u0430\u0437\u0432\u0430\u043d\u0435 \u043d\u0430 \u0442\u0435\u043a\u0443\u0449\u0438\u044f \u043d\u0430\u0431\u043e\u0440\u2026",
        "UPDATE_A_SET":         "\u0410\u043a\u0442\u0443\u0430\u043b\u0438\u0437\u0438\u0440\u0430\u043d\u0435 \u043d\u0430 \u043d\u0430\u0431\u043e\u0440\u2026",
        "RENAME_A_SET":         "\u041f\u0440\u0435\u0438\u043c\u0435\u043d\u0443\u0432\u0430\u043d\u0435 \u043d\u0430 \u043d\u0430\u0431\u043e\u0440\u2026",
        "DELETE_A_SET":         "\u0418\u0437\u0442\u0440\u0438\u0432\u0430\u043d\u0435 \u043d\u0430 \u043d\u0430\u0431\u043e\u0440\u2026",
        "RESTORE_LAST_SESSION": "\u0412\u044a\u0437\u0441\u0442\u0430\u043d\u043e\u0432\u044f\u0432\u0430\u043d\u0435 \u043d\u0430 \u043f\u043e\u0441\u043b\u0435\u0434\u043d\u0430\u0442\u0430 \u0441\u0435\u0441\u0438\u044f",
        "TAB_KEY_SWITCHING":    "\u041f\u0440\u0435\u0432\u043a\u043b\u044e\u0447\u0432\u0430\u043d\u0435 \u0441 Tab",
        "SETS_BUTTON":          "\u2630 \u041d\u0430\u0431\u043e\u0440\u0438",
        "DLG_SAVE_SET":         "\u0417\u0430\u043f\u0430\u0437\u0432\u0430\u043d\u0435 \u043d\u0430 \u043d\u0430\u0431\u043e\u0440 \u0441 \u0440\u0430\u0437\u0434\u0435\u043b\u0438",
        "DLG_RENAME_SET":       "\u041f\u0440\u0435\u0438\u043c\u0435\u043d\u0443\u0432\u0430\u043d\u0435 \u043d\u0430 \u043d\u0430\u0431\u043e\u0440",
        "DLG_UPDATE_SET":       "\u0410\u043a\u0442\u0443\u0430\u043b\u0438\u0437\u0438\u0440\u0430\u043d\u0435 \u043d\u0430 \u043d\u0430\u0431\u043e\u0440",
        "DLG_DELETE_SET":       "\u0418\u0437\u0442\u0440\u0438\u0432\u0430\u043d\u0435 \u043d\u0430 \u043d\u0430\u0431\u043e\u0440",
        "DLG_RENAME_DOC":       "\u041f\u0440\u0435\u0438\u043c\u0435\u043d\u0443\u0432\u0430\u043d\u0435 \u043d\u0430 \u0434\u043e\u043a\u0443\u043c\u0435\u043d\u0442",
        "PROMPT_SET_NAME":      "\u0418\u043c\u0435 \u043d\u0430 \u043d\u0430\u0431\u043e\u0440\u0430:",
        "PROMPT_RENAME_SET":    "\u041d\u043e\u0432\u043e \u0438\u043c\u0435 \u0437\u0430 \u00ab{old}\u00bb:",
        "PROMPT_OVERWRITE":     "\u0418\u0437\u0431\u0435\u0440\u0435\u0442\u0435 \u043d\u0430\u0431\u043e\u0440 \u0437\u0430 \u043f\u0440\u0435\u0437\u0430\u043f\u0438\u0441\u0432\u0430\u043d\u0435:",
        "PROMPT_RENAME_WHICH":  "\u0418\u0437\u0431\u0435\u0440\u0435\u0442\u0435 \u043d\u0430\u0431\u043e\u0440 \u0437\u0430 \u043f\u0440\u0435\u0438\u043c\u0435\u043d\u0443\u0432\u0430\u043d\u0435:",
        "PROMPT_DELETE_WHICH":  "\u0418\u0437\u0431\u0435\u0440\u0435\u0442\u0435 \u043d\u0430\u0431\u043e\u0440 \u0437\u0430 \u0438\u0437\u0442\u0440\u0438\u0432\u0430\u043d\u0435:",
        "PROMPT_RENAME_DOC":    "\u041d\u043e\u0432\u043e \u0438\u043c\u0435 \u0437\u0430 \u00ab{name}\u00bb:",
        "ERR_NO_DOCS":          "\u041d\u044f\u043c\u0430 \u043e\u0442\u0432\u043e\u0440\u0435\u043d\u0438 \u0437\u0430\u043f\u0430\u0437\u0435\u043d\u0438 \u0434\u043e\u043a\u0443\u043c\u0435\u043d\u0442\u0438.\n\u041d\u0430\u0439-\u043d\u0430\u043f\u0440\u0435\u0434 \u0437\u0430\u043f\u0430\u0437\u0435\u0442\u0435 \u0434\u043e\u043a\u0443\u043c\u0435\u043d\u0442\u0438\u0442\u0435 \u0441\u0438.",
        "ERR_NO_SETS":          "\u041d\u044f\u043c\u0430 \u0437\u0430\u043f\u0430\u0437\u0435\u043d\u0438 \u043d\u0430\u0431\u043e\u0440\u0438.",
        "ERR_NO_SETS_UPDATE":   "\u041d\u044f\u043c\u0430 \u043d\u0430\u0431\u043e\u0440\u0438 \u0437\u0430 \u0430\u043a\u0442\u0443\u0430\u043b\u0438\u0437\u0438\u0440\u0430\u043d\u0435.",
        "ERR_NO_OPEN_DOCS":     "\u041d\u044f\u043c\u0430 \u043e\u0442\u0432\u043e\u0440\u0435\u043d\u0438 \u0437\u0430\u043f\u0430\u0437\u0435\u043d\u0438 \u0434\u043e\u043a\u0443\u043c\u0435\u043d\u0442\u0438.",
    },
    "hr": {
        "RENAME":               "Preimenu\u0161 \u2026",
        "SAVE":                 "Spremi",
        "SAVE_AS":              "Spremi kao\u2026",
        "MOVE_LEFT":            "Pomakni lijevo",
        "MOVE_RIGHT":           "Pomakni desno",
        "NEW_DOCUMENT":         "Novi dokument",
        "CLOSE":                "Zatvori",
        "CLOSE_ALL_OTHERS":     "Zatvori ostale",
        "CLOSE_ALL":            "Zatvori sve",
        "SAVE_CURRENT_SET":     "Spremi trenutni skup\u2026",
        "UPDATE_A_SET":         "A\u017euriraj skup\u2026",
        "RENAME_A_SET":         "Preimenuj skup\u2026",
        "DELETE_A_SET":         "Izbri\u0161i skup\u2026",
        "RESTORE_LAST_SESSION": "Vrati posljednju sesiju",
        "TAB_KEY_SWITCHING":    "Preklop tipkom Tab",
        "SETS_BUTTON":          "\u2630 Skupovi",
        "DLG_SAVE_SET":         "Spremi skup kartica",
        "DLG_RENAME_SET":       "Preimenuj skup",
        "DLG_UPDATE_SET":       "A\u017euriraj skup",
        "DLG_DELETE_SET":       "Izbri\u0161i skup",
        "DLG_RENAME_DOC":       "Preimenuj dokument",
        "PROMPT_SET_NAME":      "Naziv skupa:",
        "PROMPT_RENAME_SET":    "Novi naziv za \u2018{old}\u2019:",
        "PROMPT_OVERWRITE":     "Odaberite skup za prepisivanje:",
        "PROMPT_RENAME_WHICH":  "Odaberite skup za preimenovanje:",
        "PROMPT_DELETE_WHICH":  "Odaberite skup za brisanje:",
        "PROMPT_RENAME_DOC":    "Novi naziv za \u2018{name}\u2019:",
        "ERR_NO_DOCS":          "Nema otvorenih spremljenih dokumenata.\nNajprije spremite svoje dokumente.",
        "ERR_NO_SETS":          "Nema spremljenih skupova.",
        "ERR_NO_SETS_UPDATE":   "Nema skupova za a\u017euriranje.",
        "ERR_NO_OPEN_DOCS":     "Nema otvorenih spremljenih dokumenata.",
    },
    "ca": {
        "RENAME":               "Canvia el nom\u2026",
        "SAVE":                 "Desa",
        "SAVE_AS":              "Desa com a\u2026",
        "MOVE_LEFT":            "Mou a l\u2019esquerra",
        "MOVE_RIGHT":           "Mou a la dreta",
        "NEW_DOCUMENT":         "Document nou",
        "CLOSE":                "Tanca",
        "CLOSE_ALL_OTHERS":     "Tanca els altres",
        "CLOSE_ALL":            "Tanca-ho tot",
        "SAVE_CURRENT_SET":     "Desa el conjunt actual\u2026",
        "UPDATE_A_SET":         "Actualitza un conjunt\u2026",
        "RENAME_A_SET":         "Canvia el nom d\u2019un conjunt\u2026",
        "DELETE_A_SET":         "Suprimeix un conjunt\u2026",
        "RESTORE_LAST_SESSION": "Restaura l\u2019\u00faltima sessi\u00f3",
        "TAB_KEY_SWITCHING":    "Commutaci\u00f3 amb Tab",
        "SETS_BUTTON":          "\u2630 Conjunts",
        "DLG_SAVE_SET":         "Desa el conjunt de pestanyes",
        "DLG_RENAME_SET":       "Canvia el nom del conjunt",
        "DLG_UPDATE_SET":       "Actualitza el conjunt",
        "DLG_DELETE_SET":       "Suprimeix el conjunt",
        "DLG_RENAME_DOC":       "Canvia el nom del document",
        "PROMPT_SET_NAME":      "Nom del conjunt:",
        "PROMPT_RENAME_SET":    "Nom nou per a \u2018{old}\u2019:",
        "PROMPT_OVERWRITE":     "Trieu el conjunt a sobreescriure:",
        "PROMPT_RENAME_WHICH":  "Trieu el conjunt a canviar de nom:",
        "PROMPT_DELETE_WHICH":  "Trieu el conjunt a suprimir:",
        "PROMPT_RENAME_DOC":    "Nom nou per a \u2018{name}\u2019:",
        "ERR_NO_DOCS":          "No hi ha documents desats oberts.\nDeseu primer els vostres documents.",
        "ERR_NO_SETS":          "No hi ha conjunts desats.",
        "ERR_NO_SETS_UPDATE":   "No hi ha conjunts per actualitzar.",
        "ERR_NO_OPEN_DOCS":     "No hi ha documents desats oberts.",
    },
    "eu": {
        "RENAME":               "Aldatu izena\u2026",
        "SAVE":                 "Gorde",
        "SAVE_AS":              "Gorde honela\u2026",
        "MOVE_LEFT":            "Mugitu ezkerrera",
        "MOVE_RIGHT":           "Mugitu eskuinera",
        "NEW_DOCUMENT":         "Dokumentu berria",
        "CLOSE":                "Itxi",
        "CLOSE_ALL_OTHERS":     "Itxi besteak",
        "CLOSE_ALL":            "Itxi guztiak",
        "SAVE_CURRENT_SET":     "Gorde uneko multzoa\u2026",
        "UPDATE_A_SET":         "Eguneratu multzo bat\u2026",
        "RENAME_A_SET":         "Aldatu multzoaren izena\u2026",
        "DELETE_A_SET":         "Ezabatu multzo bat\u2026",
        "RESTORE_LAST_SESSION": "Leheneratu azken saioa",
        "TAB_KEY_SWITCHING":    "Tab teklaz aldatzea",
        "SETS_BUTTON":          "\u2630 Multzoak",
        "DLG_SAVE_SET":         "Gorde fitxen multzoa",
        "DLG_RENAME_SET":       "Aldatu multzoaren izena",
        "DLG_UPDATE_SET":       "Eguneratu multzoa",
        "DLG_DELETE_SET":       "Ezabatu multzoa",
        "DLG_RENAME_DOC":       "Aldatu dokumentuaren izena",
        "PROMPT_SET_NAME":      "Multzoaren izena:",
        "PROMPT_RENAME_SET":    "\u2018{old}\u2019-ren izen berria:",
        "PROMPT_OVERWRITE":     "Hautatu gainidazteko multzoa:",
        "PROMPT_RENAME_WHICH":  "Hautatu izena aldatzeko multzoa:",
        "PROMPT_DELETE_WHICH":  "Hautatu ezabatzeko multzoa:",
        "PROMPT_RENAME_DOC":    "\u2018{name}\u2019-ren izen berria:",
        "ERR_NO_DOCS":          "Ez dago irekitako dokumentu gordetzerik.\nLehenik gorde zure dokumentuak.",
        "ERR_NO_SETS":          "Ez dago multzo gordetzerik.",
        "ERR_NO_SETS_UPDATE":   "Ez dago eguneratzeko multzorik.",
        "ERR_NO_OPEN_DOCS":     "Ez dago irekitako dokumentu gordetzerik.",
    },
    "gl": {
        "RENAME":               "Cambiar nome\u2026",
        "SAVE":                 "Gardar",
        "SAVE_AS":              "Gardar como\u2026",
        "MOVE_LEFT":            "Mover \u00e1 esquerda",
        "MOVE_RIGHT":           "Mover \u00e1 dereita",
        "NEW_DOCUMENT":         "Novo documento",
        "CLOSE":                "Pechar",
        "CLOSE_ALL_OTHERS":     "Pechar os outros",
        "CLOSE_ALL":            "Pechar todo",
        "SAVE_CURRENT_SET":     "Gardar conxunto actual\u2026",
        "UPDATE_A_SET":         "Actualizar un conxunto\u2026",
        "RENAME_A_SET":         "Cambiar nome dun conxunto\u2026",
        "DELETE_A_SET":         "Eliminar un conxunto\u2026",
        "RESTORE_LAST_SESSION": "Restaurar \u00faltima sesi\u00f3n",
        "TAB_KEY_SWITCHING":    "Cambiar con Tab",
        "SETS_BUTTON":          "\u2630 Conxuntos",
        "DLG_SAVE_SET":         "Gardar conxunto de lapelas",
        "DLG_RENAME_SET":       "Cambiar nome do conxunto",
        "DLG_UPDATE_SET":       "Actualizar conxunto",
        "DLG_DELETE_SET":       "Eliminar conxunto",
        "DLG_RENAME_DOC":       "Cambiar nome do documento",
        "PROMPT_SET_NAME":      "Nome do conxunto:",
        "PROMPT_RENAME_SET":    "Novo nome para \u2018{old}\u2019:",
        "PROMPT_OVERWRITE":     "Escoller conxunto a sobreescribir:",
        "PROMPT_RENAME_WHICH":  "Escoller conxunto a renomear:",
        "PROMPT_DELETE_WHICH":  "Escoller conxunto a eliminar:",
        "PROMPT_RENAME_DOC":    "Novo nome para \u2018{name}\u2019:",
        "ERR_NO_DOCS":          "Non hai documentos gardados abertos.\nGarde primeiro os seus documentos.",
        "ERR_NO_SETS":          "Non hai conxuntos gardados.",
        "ERR_NO_SETS_UPDATE":   "Non hai conxuntos para actualizar.",
        "ERR_NO_OPEN_DOCS":     "Non hai documentos gardados abertos.",
    },
    "vi": {
        "RENAME":               "\u0110\u1ed5i t\u00ean\u2026",
        "SAVE":                 "L\u01b0u",
        "SAVE_AS":              "L\u01b0u d\u01b0\u1edbi t\u00ean\u2026",
        "MOVE_LEFT":            "Di chuy\u1ec3n sang tr\u00e1i",
        "MOVE_RIGHT":           "Di chuy\u1ec3n sang ph\u1ea3i",
        "NEW_DOCUMENT":         "T\u00e0i li\u1ec7u m\u1edbi",
        "CLOSE":                "\u0110\u00f3ng",
        "CLOSE_ALL_OTHERS":     "\u0110\u00f3ng c\u00e1c t\u00e0i li\u1ec7u kh\u00e1c",
        "CLOSE_ALL":            "\u0110\u00f3ng t\u1ea5t c\u1ea3",
        "SAVE_CURRENT_SET":     "L\u01b0u b\u1ed9 hi\u1ec7n t\u1ea1i\u2026",
        "UPDATE_A_SET":         "C\u1eadp nh\u1eadt b\u1ed9\u2026",
        "RENAME_A_SET":         "\u0110\u1ed5i t\u00ean b\u1ed9\u2026",
        "DELETE_A_SET":         "X\u00f3a b\u1ed9\u2026",
        "RESTORE_LAST_SESSION": "Kh\u00f4i ph\u1ee5c phi\u00ean l\u00e0m vi\u1ec7c tr\u01b0\u1edbc",
        "TAB_KEY_SWITCHING":    "Chuy\u1ec3n b\u1eb1ng ph\u00edm Tab",
        "SETS_BUTTON":          "\u2630 B\u1ed9",
        "DLG_SAVE_SET":         "L\u01b0u b\u1ed9 th\u1ebb",
        "DLG_RENAME_SET":       "\u0110\u1ed5i t\u00ean b\u1ed9",
        "DLG_UPDATE_SET":       "C\u1eadp nh\u1eadt b\u1ed9",
        "DLG_DELETE_SET":       "X\u00f3a b\u1ed9",
        "DLG_RENAME_DOC":       "\u0110\u1ed5i t\u00ean t\u00e0i li\u1ec7u",
        "PROMPT_SET_NAME":      "T\u00ean b\u1ed9:",
        "PROMPT_RENAME_SET":    "T\u00ean m\u1edbi cho \u2018{old}\u2019:",
        "PROMPT_OVERWRITE":     "Ch\u1ecdn b\u1ed9 \u0111\u1ec3 ghi \u0111\u00e8:",
        "PROMPT_RENAME_WHICH":  "Ch\u1ecdn b\u1ed9 \u0111\u1ec3 \u0111\u1ed5i t\u00ean:",
        "PROMPT_DELETE_WHICH":  "Ch\u1ecdn b\u1ed9 \u0111\u1ec3 x\u00f3a:",
        "PROMPT_RENAME_DOC":    "T\u00ean m\u1edbi cho \u2018{name}\u2019:",
        "ERR_NO_DOCS":          "Kh\u00f4ng c\u00f3 t\u00e0i li\u1ec7u \u0111\u00e3 l\u01b0u n\u00e0o \u0111ang m\u1edf.\nH\u00e3y l\u01b0u t\u00e0i li\u1ec7u tr\u01b0\u1edbc.",
        "ERR_NO_SETS":          "Kh\u00f4ng c\u00f3 b\u1ed9 \u0111\u00e3 l\u01b0u.",
        "ERR_NO_SETS_UPDATE":   "Kh\u00f4ng c\u00f3 b\u1ed9 n\u00e0o \u0111\u1ec3 c\u1eadp nh\u1eadt.",
        "ERR_NO_OPEN_DOCS":     "Kh\u00f4ng c\u00f3 t\u00e0i li\u1ec7u \u0111\u00e3 l\u01b0u n\u00e0o \u0111ang m\u1edf.",
    },
    "id": {
        "RENAME":               "Ganti nama\u2026",
        "SAVE":                 "Simpan",
        "SAVE_AS":              "Simpan sebagai\u2026",
        "MOVE_LEFT":            "Pindah ke kiri",
        "MOVE_RIGHT":           "Pindah ke kanan",
        "NEW_DOCUMENT":         "Dokumen baru",
        "CLOSE":                "Tutup",
        "CLOSE_ALL_OTHERS":     "Tutup yang lain",
        "CLOSE_ALL":            "Tutup semua",
        "SAVE_CURRENT_SET":     "Simpan set saat ini\u2026",
        "UPDATE_A_SET":         "Perbarui set\u2026",
        "RENAME_A_SET":         "Ganti nama set\u2026",
        "DELETE_A_SET":         "Hapus set\u2026",
        "RESTORE_LAST_SESSION": "Pulihkan sesi terakhir",
        "TAB_KEY_SWITCHING":    "Beralih dengan Tab",
        "SETS_BUTTON":          "\u2630 Set",
        "DLG_SAVE_SET":         "Simpan set tab",
        "DLG_RENAME_SET":       "Ganti nama set",
        "DLG_UPDATE_SET":       "Perbarui set",
        "DLG_DELETE_SET":       "Hapus set",
        "DLG_RENAME_DOC":       "Ganti nama dokumen",
        "PROMPT_SET_NAME":      "Nama set:",
        "PROMPT_RENAME_SET":    "Nama baru untuk \u2018{old}\u2019:",
        "PROMPT_OVERWRITE":     "Pilih set untuk ditimpa:",
        "PROMPT_RENAME_WHICH":  "Pilih set untuk diganti namanya:",
        "PROMPT_DELETE_WHICH":  "Pilih set untuk dihapus:",
        "PROMPT_RENAME_DOC":    "Nama baru untuk \u2018{name}\u2019:",
        "ERR_NO_DOCS":          "Tidak ada dokumen tersimpan yang terbuka.\nSimpan dokumen Anda terlebih dahulu.",
        "ERR_NO_SETS":          "Tidak ada set tersimpan.",
        "ERR_NO_SETS_UPDATE":   "Tidak ada set untuk diperbarui.",
        "ERR_NO_OPEN_DOCS":     "Tidak ada dokumen tersimpan yang terbuka.",
    },
    "ar": {
        "RENAME":               "\u0625\u0639\u0627\u062f\u0629 \u062a\u0633\u0645\u064a\u0629\u2026",
        "SAVE":                 "\u062d\u0641\u0638",
        "SAVE_AS":              "\u062d\u0641\u0638 \u0628\u0627\u0633\u0645\u2026",
        "MOVE_LEFT":            "\u062a\u062d\u0631\u064a\u0643 \u064a\u0633\u0627\u0631\u0627\u064b",
        "MOVE_RIGHT":           "\u062a\u062d\u0631\u064a\u0643 \u064a\u0645\u064a\u0646\u0627\u064b",
        "NEW_DOCUMENT":         "\u0645\u0633\u062a\u0646\u062f \u062c\u062f\u064a\u062f",
        "CLOSE":                "\u0625\u063a\u0644\u0627\u0642",
        "CLOSE_ALL_OTHERS":     "\u0625\u063a\u0644\u0627\u0642 \u0627\u0644\u0628\u0627\u0642\u064a",
        "CLOSE_ALL":            "\u0625\u063a\u0644\u0627\u0642 \u0627\u0644\u0643\u0644",
        "SAVE_CURRENT_SET":     "\u062d\u0641\u0638 \u0627\u0644\u0645\u062c\u0645\u0648\u0639\u0629 \u0627\u0644\u062d\u0627\u0644\u064a\u0629\u2026",
        "UPDATE_A_SET":         "\u062a\u062d\u062f\u064a\u062b \u0645\u062c\u0645\u0648\u0639\u0629\u2026",
        "RENAME_A_SET":         "\u0625\u0639\u0627\u062f\u0629 \u062a\u0633\u0645\u064a\u0629 \u0645\u062c\u0645\u0648\u0639\u0629\u2026",
        "DELETE_A_SET":         "\u062d\u0630\u0641 \u0645\u062c\u0645\u0648\u0639\u0629\u2026",
        "RESTORE_LAST_SESSION": "\u0627\u0633\u062a\u0639\u0627\u062f\u0629 \u0622\u062e\u0631 \u062c\u0644\u0633\u0629",
        "TAB_KEY_SWITCHING":    "\u0627\u0644\u062a\u0628\u062f\u064a\u0644 \u0628\u0645\u0641\u062a\u0627\u062d Tab",
        "SETS_BUTTON":          "\u2630 \u0627\u0644\u0645\u062c\u0645\u0648\u0639\u0627\u062a",
        "DLG_SAVE_SET":         "\u062d\u0641\u0638 \u0645\u062c\u0645\u0648\u0639\u0629 \u0627\u0644\u062a\u0628\u0648\u064a\u0628\u0627\u062a",
        "DLG_RENAME_SET":       "\u0625\u0639\u0627\u062f\u0629 \u062a\u0633\u0645\u064a\u0629 \u0627\u0644\u0645\u062c\u0645\u0648\u0639\u0629",
        "DLG_UPDATE_SET":       "\u062a\u062d\u062f\u064a\u062b \u0627\u0644\u0645\u062c\u0645\u0648\u0639\u0629",
        "DLG_DELETE_SET":       "\u062d\u0630\u0641 \u0627\u0644\u0645\u062c\u0645\u0648\u0639\u0629",
        "DLG_RENAME_DOC":       "\u0625\u0639\u0627\u062f\u0629 \u062a\u0633\u0645\u064a\u0629 \u0627\u0644\u0645\u0633\u062a\u0646\u062f",
        "PROMPT_SET_NAME":      "\u0627\u0633\u0645 \u0627\u0644\u0645\u062c\u0645\u0648\u0639\u0629:",
        "PROMPT_RENAME_SET":    "\u0627\u0633\u0645 \u062c\u062f\u064a\u062f \u0644\u0640 \u2018{old}\u2019:",
        "PROMPT_OVERWRITE":     "\u0627\u062e\u062a\u0631 \u0645\u062c\u0645\u0648\u0639\u0629 \u0644\u0644\u0643\u062a\u0627\u0628\u0629 \u0641\u0648\u0642\u0647\u0627:",
        "PROMPT_RENAME_WHICH":  "\u0627\u062e\u062a\u0631 \u0645\u062c\u0645\u0648\u0639\u0629 \u0644\u0625\u0639\u0627\u062f\u0629 \u0627\u0644\u062a\u0633\u0645\u064a\u0629:",
        "PROMPT_DELETE_WHICH":  "\u0627\u062e\u062a\u0631 \u0645\u062c\u0645\u0648\u0639\u0629 \u0644\u0644\u062d\u0630\u0641:",
        "PROMPT_RENAME_DOC":    "\u0627\u0633\u0645 \u062c\u062f\u064a\u062f \u0644\u0640 \u2018{name}\u2019:",
        "ERR_NO_DOCS":          "\u0644\u0627 \u062a\u0648\u062c\u062f \u0645\u0633\u062a\u0646\u062f\u0627\u062a \u0645\u062d\u0641\u0648\u0638\u0629 \u0645\u0641\u062a\u0648\u062d\u0629.\n\u0627\u062d\u0641\u0638 \u0645\u0633\u062a\u0646\u062f\u0627\u062a\u0643 \u0623\u0648\u0644\u0627\u064b.",
        "ERR_NO_SETS":          "\u0644\u0627 \u062a\u0648\u062c\u062f \u0645\u062c\u0645\u0648\u0639\u0627\u062a \u0645\u062d\u0641\u0648\u0638\u0629.",
        "ERR_NO_SETS_UPDATE":   "\u0644\u0627 \u062a\u0648\u062c\u062f \u0645\u062c\u0645\u0648\u0639\u0627\u062a \u0644\u0644\u062a\u062d\u062f\u064a\u062b.",
        "ERR_NO_OPEN_DOCS":     "\u0644\u0627 \u062a\u0648\u062c\u062f \u0645\u0633\u062a\u0646\u062f\u0627\u062a \u0645\u062d\u0641\u0648\u0638\u0629 \u0645\u0641\u062a\u0648\u062d\u0629.",
    },
    "he": {
        "RENAME":               "\u05e9\u05d9\u05e0\u05d5\u05d9 \u05e9\u05dd\u2026",
        "SAVE":                 "\u05e9\u05de\u05d9\u05e8\u05d4",
        "SAVE_AS":              "\u05e9\u05de\u05d9\u05e8\u05d4 \u05d1\u05e9\u05dd\u2026",
        "MOVE_LEFT":            "\u05d4\u05d6\u05d6 \u05e9\u05de\u05d0\u05dc\u05d4",
        "MOVE_RIGHT":           "\u05d4\u05d6\u05d6 \u05d9\u05de\u05d9\u05e0\u05d4",
        "NEW_DOCUMENT":         "\u05de\u05e1\u05de\u05da \u05d7\u05d3\u05e9",
        "CLOSE":                "\u05e1\u05d2\u05d9\u05e8\u05d4",
        "CLOSE_ALL_OTHERS":     "\u05e1\u05d2\u05d5\u05e8 \u05d0\u05ea \u05d4\u05e9\u05d0\u05e8",
        "CLOSE_ALL":            "\u05e1\u05d2\u05d5\u05e8 \u05d4\u05db\u05dc",
        "SAVE_CURRENT_SET":     "\u05e9\u05de\u05d5\u05e8 \u05e1\u05d8 \u05e0\u05d5\u05db\u05d7\u05d9\u2026",
        "UPDATE_A_SET":         "\u05e2\u05d3\u05db\u05d5\u05df \u05e1\u05d8\u2026",
        "RENAME_A_SET":         "\u05e9\u05d9\u05e0\u05d5\u05d9 \u05e9\u05dd \u05e1\u05d8\u2026",
        "DELETE_A_SET":         "\u05de\u05d7\u05d9\u05e7\u05ea \u05e1\u05d8\u2026",
        "RESTORE_LAST_SESSION": "\u05e9\u05d7\u05d6\u05d5\u05e8 \u05dc\u05e1\u05e9\u05d9\u05d0\u05d4 \u05d4\u05d0\u05d7\u05e8\u05d5\u05e0\u05d4",
        "TAB_KEY_SWITCHING":    "\u05de\u05d9\u05ea\u05d5\u05d2 \u05d1\u05de\u05e7\u05e9 Tab",
        "SETS_BUTTON":          "\u2630 \u05e1\u05d8\u05d9\u05dd",
        "DLG_SAVE_SET":         "\u05e9\u05de\u05d9\u05e8\u05ea \u05e1\u05d8 \u05d4\u05dc\u05e9\u05d5\u05e0\u05d9\u05ea",
        "DLG_RENAME_SET":       "\u05e9\u05d9\u05e0\u05d5\u05d9 \u05e9\u05dd \u05e1\u05d8",
        "DLG_UPDATE_SET":       "\u05e2\u05d3\u05db\u05d5\u05df \u05e1\u05d8",
        "DLG_DELETE_SET":       "\u05de\u05d7\u05d9\u05e7\u05ea \u05e1\u05d8",
        "DLG_RENAME_DOC":       "\u05e9\u05d9\u05e0\u05d5\u05d9 \u05e9\u05dd \u05de\u05e1\u05de\u05da",
        "PROMPT_SET_NAME":      "\u05e9\u05dd \u05d4\u05e1\u05d8:",
        "PROMPT_RENAME_SET":    "\u05e9\u05dd \u05d7\u05d3\u05e9 \u05e2\u05d1\u05d5\u05e8 \u2018{old}\u2019:",
        "PROMPT_OVERWRITE":     "\u05d1\u05d7\u05e8 \u05e1\u05d8 \u05dc\u05d3\u05e8\u05d9\u05e1\u05d4:",
        "PROMPT_RENAME_WHICH":  "\u05d1\u05d7\u05e8 \u05e1\u05d8 \u05dc\u05e9\u05d9\u05e0\u05d5\u05d9 \u05e9\u05dd:",
        "PROMPT_DELETE_WHICH":  "\u05d1\u05d7\u05e8 \u05e1\u05d8 \u05dc\u05de\u05d7\u05d9\u05e7\u05d4:",
        "PROMPT_RENAME_DOC":    "\u05e9\u05dd \u05d7\u05d3\u05e9 \u05e2\u05d1\u05d5\u05e8 \u2018{name}\u2019:",
        "ERR_NO_DOCS":          "\u05d0\u05d9\u05df \u05de\u05e1\u05de\u05db\u05d9\u05dd \u05e9\u05de\u05d5\u05e8\u05d9\u05dd \u05e4\u05ea\u05d5\u05d7\u05d9\u05dd.\n\u05e9\u05de\u05d5\u05e8 \u05d0\u05ea \u05d4\u05de\u05e1\u05de\u05db\u05d9\u05dd \u05ea\u05d7\u05d9\u05dc\u05d4.",
        "ERR_NO_SETS":          "\u05d0\u05d9\u05df \u05e1\u05d8\u05d9\u05dd \u05e9\u05de\u05d5\u05e8\u05d9\u05dd.",
        "ERR_NO_SETS_UPDATE":   "\u05d0\u05d9\u05df \u05e1\u05d8\u05d9\u05dd \u05dc\u05e2\u05d3\u05db\u05d5\u05df.",
        "ERR_NO_OPEN_DOCS":     "\u05d0\u05d9\u05df \u05de\u05e1\u05de\u05db\u05d9\u05dd \u05e9\u05de\u05d5\u05e8\u05d9\u05dd \u05e4\u05ea\u05d5\u05d7\u05d9\u05dd.",
    },
    "hi": {
        "RENAME":               "\u0928\u093e\u092e \u092c\u0926\u0932\u0947\u0902\u2026",
        "SAVE":                 "\u0938\u0939\u0947\u091c\u0947\u0902",
        "SAVE_AS":              "\u0907\u0938 \u0928\u093e\u092e \u0938\u0947 \u0938\u0939\u0947\u091c\u0947\u0902\u2026",
        "MOVE_LEFT":            "\u092c\u093e\u090f\u0902 \u0932\u0947 \u091c\u093e\u090f\u0902",
        "MOVE_RIGHT":           "\u0926\u093e\u090f\u0902 \u0932\u0947 \u091c\u093e\u090f\u0902",
        "NEW_DOCUMENT":         "\u0928\u092f\u093e \u0926\u0938\u094d\u0924\u093e\u0935\u0947\u091c\u093c",
        "CLOSE":                "\u092c\u0902\u0926 \u0915\u0930\u0947\u0902",
        "CLOSE_ALL_OTHERS":     "\u0905\u0928\u094d\u092f \u0938\u092d\u0940 \u092c\u0902\u0926 \u0915\u0930\u0947\u0902",
        "CLOSE_ALL":            "\u0938\u092d\u0940 \u092c\u0902\u0926 \u0915\u0930\u0947\u0902",
        "SAVE_CURRENT_SET":     "\u0935\u0930\u094d\u0924\u092e\u093e\u0928 \u0938\u0947\u091f \u0938\u0939\u0947\u091c\u0947\u0902\u2026",
        "UPDATE_A_SET":         "\u0938\u0947\u091f \u0905\u092a\u0921\u0947\u091f \u0915\u0930\u0947\u0902\u2026",
        "RENAME_A_SET":         "\u0938\u0947\u091f \u0915\u093e \u0928\u093e\u092e \u092c\u0926\u0932\u0947\u0902\u2026",
        "DELETE_A_SET":         "\u0938\u0947\u091f \u0939\u091f\u093e\u090f\u0902\u2026",
        "RESTORE_LAST_SESSION": "\u0905\u0902\u0924\u093f\u092e \u0938\u0924\u094d\u0930 \u092a\u0941\u0928\u0903\u0938\u094d\u0925\u093e\u092a\u093f\u0924 \u0915\u0930\u0947\u0902",
        "TAB_KEY_SWITCHING":    "Tab \u0938\u0947 \u0938\u094d\u0935\u093f\u091a \u0915\u0930\u0947\u0902",
        "SETS_BUTTON":          "\u2630 \u0938\u0947\u091f",
        "DLG_SAVE_SET":         "\u091f\u0948\u092c \u0938\u0947\u091f \u0938\u0939\u0947\u091c\u0947\u0902",
        "DLG_RENAME_SET":       "\u0938\u0947\u091f \u0915\u093e \u0928\u093e\u092e \u092c\u0926\u0932\u0947\u0902",
        "DLG_UPDATE_SET":       "\u0938\u0947\u091f \u0905\u092a\u0921\u0947\u091f \u0915\u0930\u0947\u0902",
        "DLG_DELETE_SET":       "\u0938\u0947\u091f \u0939\u091f\u093e\u090f\u0902",
        "DLG_RENAME_DOC":       "\u0926\u0938\u094d\u0924\u093e\u0935\u0947\u091c\u093c \u0915\u093e \u0928\u093e\u092e \u092c\u0926\u0932\u0947\u0902",
        "PROMPT_SET_NAME":      "\u0938\u0947\u091f \u0915\u093e \u0928\u093e\u092e:",
        "PROMPT_RENAME_SET":    "\u2018{old}\u2019 \u0915\u093e \u0928\u092f\u093e \u0928\u093e\u092e:",
        "PROMPT_OVERWRITE":     "\u0913\u0935\u0930\u0930\u093e\u0907\u091f \u0915\u0930\u0928\u0947 \u0939\u0947\u0924\u0941 \u0938\u0947\u091f \u091a\u0941\u0928\u0947\u0902:",
        "PROMPT_RENAME_WHICH":  "\u0928\u093e\u092e \u092c\u0926\u0932\u0928\u0947 \u0939\u0947\u0924\u0941 \u0938\u0947\u091f \u091a\u0941\u0928\u0947\u0902:",
        "PROMPT_DELETE_WHICH":  "\u0939\u091f\u093e\u0928\u0947 \u0939\u0947\u0924\u0941 \u0938\u0947\u091f \u091a\u0941\u0928\u0947\u0902:",
        "PROMPT_RENAME_DOC":    "\u2018{name}\u2019 \u0915\u093e \u0928\u092f\u093e \u0928\u093e\u092e:",
        "ERR_NO_DOCS":          "\u0915\u094b\u0908 \u0938\u0939\u0947\u091c\u093e \u0926\u0938\u094d\u0924\u093e\u0935\u0947\u091c\u093c \u0916\u0941\u0932\u093e \u0928\u0939\u0940\u0902 \u0939\u0948\u0964\n\u092a\u0939\u0932\u0947 \u0905\u092a\u0928\u0947 \u0926\u0938\u094d\u0924\u093e\u0935\u0947\u091c\u093c \u0938\u0939\u0947\u091c\u0947\u0902\u0964",
        "ERR_NO_SETS":          "\u0915\u094b\u0908 \u0938\u0939\u0947\u091c\u093e \u0938\u0947\u091f \u0928\u0939\u0940\u0902 \u0939\u0948\u0964",
        "ERR_NO_SETS_UPDATE":   "\u0905\u092a\u0921\u0947\u091f \u0915\u0930\u0928\u0947 \u0939\u0947\u0924\u0941 \u0915\u094b\u0908 \u0938\u0947\u091f \u0928\u0939\u0940\u0902 \u0939\u0948\u0964",
        "ERR_NO_OPEN_DOCS":     "\u0915\u094b\u0908 \u0938\u0939\u0947\u091c\u093e \u0926\u0938\u094d\u0924\u093e\u0935\u0947\u091c\u093c \u0916\u0941\u0932\u093e \u0928\u0939\u0940\u0902 \u0939\u0948\u0964",
    },
    "af": {
        "RENAME":               "Hernoem\u2026",
        "SAVE":                 "Stoor",
        "SAVE_AS":              "Stoor as\u2026",
        "MOVE_LEFT":            "Skuif links",
        "MOVE_RIGHT":           "Skuif regs",
        "NEW_DOCUMENT":         "Nuwe dokument",
        "CLOSE":                "Sluit",
        "CLOSE_ALL_OTHERS":     "Sluit ander",
        "CLOSE_ALL":            "Sluit alles",
        "SAVE_CURRENT_SET":     "Stoor huidige stel\u2026",
        "UPDATE_A_SET":         "Werk \u2018n stel op\u2026",
        "RENAME_A_SET":         "Hernoem \u2018n stel\u2026",
        "DELETE_A_SET":         "Vee \u2018n stel uit\u2026",
        "RESTORE_LAST_SESSION": "Herstel laaste sessie",
        "TAB_KEY_SWITCHING":    "Skakel met Tab",
        "SETS_BUTTON":          "\u2630 Stelle",
        "DLG_SAVE_SET":         "Stoor oortjiestel",
        "DLG_RENAME_SET":       "Hernoem stel",
        "DLG_UPDATE_SET":       "Werk stel op",
        "DLG_DELETE_SET":       "Vee stel uit",
        "DLG_RENAME_DOC":       "Hernoem dokument",
        "PROMPT_SET_NAME":      "Stelnaam:",
        "PROMPT_RENAME_SET":    "Nuwe naam vir \u2018{old}\u2019:",
        "PROMPT_OVERWRITE":     "Kies stel om oor te skryf:",
        "PROMPT_RENAME_WHICH":  "Kies stel om te hernoem:",
        "PROMPT_DELETE_WHICH":  "Kies stel om uit te vee:",
        "PROMPT_RENAME_DOC":    "Nuwe naam vir \u2018{name}\u2019:",
        "ERR_NO_DOCS":          "Geen gestoorde dokumente oop nie.\nStoor eers u dokumente.",
        "ERR_NO_SETS":          "Geen gestoorde stelle nie.",
        "ERR_NO_SETS_UPDATE":   "Geen stelle om op te dateer nie.",
        "ERR_NO_OPEN_DOCS":     "Geen gestoorde dokumente oop nie.",
    },
    "fa": {
        "RENAME":               "\u062a\u063a\u06cc\u06cc\u0631 \u0646\u0627\u0645\u2026",
        "SAVE":                 "\u0630\u062e\u06cc\u0631\u0647",
        "SAVE_AS":              "\u0630\u062e\u06cc\u0631\u0647 \u0628\u0627 \u0646\u0627\u0645\u2026",
        "MOVE_LEFT":            "\u0627\u0646\u062a\u0642\u0627\u0644 \u0628\u0647 \u0686\u067e",
        "MOVE_RIGHT":           "\u0627\u0646\u062a\u0642\u0627\u0644 \u0628\u0647 \u0631\u0627\u0633\u062a",
        "NEW_DOCUMENT":         "\u0633\u0646\u062f \u062c\u062f\u06cc\u062f",
        "CLOSE":                "\u0628\u0633\u062a\u0646",
        "CLOSE_ALL_OTHERS":     "\u0628\u0633\u062a\u0646 \u0628\u0642\u06cc\u0647",
        "CLOSE_ALL":            "\u0628\u0633\u062a\u0646 \u0647\u0645\u0647",
        "SAVE_CURRENT_SET":     "\u0630\u062e\u06cc\u0631\u0647 \u0645\u062c\u0645\u0648\u0639\u0647 \u0641\u0639\u0644\u06cc\u2026",
        "UPDATE_A_SET":         "\u0628\u0647\u200c\u0631\u0648\u0632\u0631\u0633\u0627\u0646\u06cc \u0645\u062c\u0645\u0648\u0639\u0647\u2026",
        "RENAME_A_SET":         "\u062a\u063a\u06cc\u06cc\u0631 \u0646\u0627\u0645 \u0645\u062c\u0645\u0648\u0639\u0647\u2026",
        "DELETE_A_SET":         "\u062d\u0630\u0641 \u0645\u062c\u0645\u0648\u0639\u0647\u2026",
        "RESTORE_LAST_SESSION": "\u0628\u0627\u0632\u06af\u0631\u062f\u0627\u0646\u06cc \u0622\u062e\u0631\u06cc\u0646 \u0646\u0634\u0633\u062a",
        "TAB_KEY_SWITCHING":    "\u062c\u0627\u0628\u062c\u0627\u06cc\u06cc \u0628\u0627 Tab",
        "SETS_BUTTON":          "\u2630 \u0645\u062c\u0645\u0648\u0639\u0647\u200c\u0647\u0627",
        "DLG_SAVE_SET":         "\u0630\u062e\u06cc\u0631\u0647 \u0645\u062c\u0645\u0648\u0639\u0647 \u0632\u0628\u0627\u0646\u0647",
        "DLG_RENAME_SET":       "\u062a\u063a\u06cc\u06cc\u0631 \u0646\u0627\u0645 \u0645\u062c\u0645\u0648\u0639\u0647",
        "DLG_UPDATE_SET":       "\u0628\u0647\u200c\u0631\u0648\u0632\u0631\u0633\u0627\u0646\u06cc \u0645\u062c\u0645\u0648\u0639\u0647",
        "DLG_DELETE_SET":       "\u062d\u0630\u0641 \u0645\u062c\u0645\u0648\u0639\u0647",
        "DLG_RENAME_DOC":       "\u062a\u063a\u06cc\u06cc\u0631 \u0646\u0627\u0645 \u0633\u0646\u062f",
        "PROMPT_SET_NAME":      "\u0646\u0627\u0645 \u0645\u062c\u0645\u0648\u0639\u0647:",
        "PROMPT_RENAME_SET":    "\u0646\u0627\u0645 \u062c\u062f\u06cc\u062f \u0628\u0631\u0627\u06cc \u2018{old}\u2019:",
        "PROMPT_OVERWRITE":     "\u0645\u062c\u0645\u0648\u0639\u0647\u200c\u0627\u06cc \u0631\u0627 \u0628\u0631\u0627\u06cc \u0628\u0627\u0632\u0646\u0648\u06cc\u0633\u06cc \u0627\u0646\u062a\u062e\u0627\u0628 \u06a9\u0646\u06cc\u062f:",
        "PROMPT_RENAME_WHICH":  "\u0645\u062c\u0645\u0648\u0639\u0647\u200c\u0627\u06cc \u0631\u0627 \u0628\u0631\u0627\u06cc \u062a\u063a\u06cc\u06cc\u0631 \u0646\u0627\u0645 \u0627\u0646\u062a\u062e\u0627\u0628 \u06a9\u0646\u06cc\u062f:",
        "PROMPT_DELETE_WHICH":  "\u0645\u062c\u0645\u0648\u0639\u0647\u200c\u0627\u06cc \u0631\u0627 \u0628\u0631\u0627\u06cc \u062d\u0630\u0641 \u0627\u0646\u062a\u062e\u0627\u0628 \u06a9\u0646\u06cc\u062f:",
        "PROMPT_RENAME_DOC":    "\u0646\u0627\u0645 \u062c\u062f\u06cc\u062f \u0628\u0631\u0627\u06cc \u2018{name}\u2019:",
        "ERR_NO_DOCS":          "\u0647\u06cc\u0686 \u0633\u0646\u062f \u0630\u062e\u06cc\u0631\u0647\u200c\u0634\u062f\u0647\u200c\u0627\u06cc \u0628\u0627\u0632 \u0646\u06cc\u0633\u062a.\n\u0627\u0628\u062a\u062f\u0627 \u0627\u0633\u0646\u0627\u062f \u062e\u0648\u062f \u0631\u0627 \u0630\u062e\u06cc\u0631\u0647 \u06a9\u0646\u06cc\u062f.",
        "ERR_NO_SETS":          "\u0647\u06cc\u0686 \u0645\u062c\u0645\u0648\u0639\u0647 \u0632\u0628\u0627\u0646\u0647\u200c\u0627\u06cc \u0630\u062e\u06cc\u0631\u0647 \u0646\u0634\u062f\u0647.",
        "ERR_NO_SETS_UPDATE":   "\u0647\u06cc\u0686 \u0645\u062c\u0645\u0648\u0639\u0647\u200c\u0627\u06cc \u0628\u0631\u0627\u06cc \u0628\u0647\u200c\u0631\u0648\u0632\u0631\u0633\u0627\u0646\u06cc \u0648\u062c\u0648\u062f \u0646\u062f\u0627\u0631\u062f.",
        "ERR_NO_OPEN_DOCS":     "\u0647\u06cc\u0686 \u0633\u0646\u062f \u0630\u062e\u06cc\u0631\u0647\u200c\u0634\u062f\u0647\u200c\u0627\u06cc \u0628\u0627\u0632 \u0646\u06cc\u0633\u062a.",
    },
    "sw": {
        "RENAME":               "Badilisha jina\u2026",
        "SAVE":                 "Hifadhi",
        "SAVE_AS":              "Hifadhi kama\u2026",
        "MOVE_LEFT":            "Hamisha kushoto",
        "MOVE_RIGHT":           "Hamisha kulia",
        "NEW_DOCUMENT":         "Hati mpya",
        "CLOSE":                "Funga",
        "CLOSE_ALL_OTHERS":     "Funga zingine",
        "CLOSE_ALL":            "Funga zote",
        "SAVE_CURRENT_SET":     "Hifadhi seti ya sasa\u2026",
        "UPDATE_A_SET":         "Sasisha seti\u2026",
        "RENAME_A_SET":         "Badilisha jina la seti\u2026",
        "DELETE_A_SET":         "Futa seti\u2026",
        "RESTORE_LAST_SESSION": "Rejesha kikao cha mwisho",
        "TAB_KEY_SWITCHING":    "Kubadilisha kwa Tab",
        "SETS_BUTTON":          "\u2630 Seti",
        "DLG_SAVE_SET":         "Hifadhi seti ya vichupo",
        "DLG_RENAME_SET":       "Badilisha jina la seti",
        "DLG_UPDATE_SET":       "Sasisha seti",
        "DLG_DELETE_SET":       "Futa seti",
        "DLG_RENAME_DOC":       "Badilisha jina la hati",
        "PROMPT_SET_NAME":      "Jina la seti:",
        "PROMPT_RENAME_SET":    "Jina jipya la \u2018{old}\u2019:",
        "PROMPT_OVERWRITE":     "Chagua seti ya kufunika:",
        "PROMPT_RENAME_WHICH":  "Chagua seti ya kubadilisha jina:",
        "PROMPT_DELETE_WHICH":  "Chagua seti ya kufuta:",
        "PROMPT_RENAME_DOC":    "Jina jipya la \u2018{name}\u2019:",
        "ERR_NO_DOCS":          "Hakuna hati zilizohifadhiwa zilizo wazi.\nHifadhi hati zako kwanza.",
        "ERR_NO_SETS":          "Hakuna seti zilizohifadhiwa.",
        "ERR_NO_SETS_UPDATE":   "Hakuna seti za kusasisha.",
        "ERR_NO_OPEN_DOCS":     "Hakuna hati zilizohifadhiwa zilizo wazi.",
    },
}

_locale = "en"


def _detect_locale(ctx):
    global _locale
    try:
        cp = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.configuration.ConfigurationProvider", ctx)
        args = (uno.createUnoStruct("com.sun.star.beans.NamedValue"),)
        args[0].Name  = "nodepath"
        args[0].Value = "/org.openoffice.Setup/L10N"
        access  = cp.createInstanceWithArguments(
            "com.sun.star.configuration.ConfigurationUpdateAccess", args)
        tag      = access.getByName("ooLocale")
        tag_norm = tag.lower().replace("_", "-")
        for full in ("zh-cn", "zh-tw", "pt-br"):
            if tag_norm.startswith(full):
                if full in _STRINGS:
                    _locale = full
                return
        lang = tag_norm.split("-")[0]
        if lang in _STRINGS:
            _locale = lang
    except Exception:
        pass


def _t(key):
    """Return the localised string for key, falling back to English."""
    return _STRINGS.get(_locale, _STRINGS["en"]).get(key, _STRINGS["en"].get(key, key))


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
        _show_message(ctx, win, _t("DLG_SAVE_SET"), _t("ERR_NO_DOCS"))
        return

    name = _get_input(ctx, win, _t("DLG_SAVE_SET"), _t("PROMPT_SET_NAME"), "")
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
        _show_message(ctx, win, _t("DLG_RENAME_SET"), _t("ERR_NO_SETS"))
        return

    old = _pick_from_list(ctx, win, _t("DLG_RENAME_SET"), _t("PROMPT_RENAME_WHICH"), names)
    if old is None:
        return

    new = _get_input(ctx, win, _t("DLG_RENAME_SET"), _t("PROMPT_RENAME_SET").format(old=old), old)
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
        _show_message(ctx, win, _t("DLG_UPDATE_SET"), _t("ERR_NO_SETS_UPDATE"))
        return
    name = _pick_from_list(ctx, win, _t("DLG_UPDATE_SET"), _t("PROMPT_OVERWRITE"), names)
    if name is None:
        return
    urls = _current_saved_urls()
    if not urls:
        _show_message(ctx, win, _t("DLG_UPDATE_SET"), _t("ERR_NO_OPEN_DOCS"))
        return
    sets[name] = urls
    _save_sets(sets)
    _log(f"updated tab set {name!r}: {len(urls)} URL(s)")


def _delete_set_dialog(ctx, win):
    """Pick a saved set by name, then delete it."""
    sets  = _load_sets()
    names = sorted(sets)
    if not names:
        _show_message(ctx, win, _t("DLG_DELETE_SET"), _t("ERR_NO_SETS"))
        return

    name = _pick_from_list(ctx, win, _t("DLG_DELETE_SET"), _t("PROMPT_DELETE_WHICH"), names)
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
        popup.insertItem(1, _t("SAVE_CURRENT_SET"), 0, 0)
        popup.insertItem(4, _t("UPDATE_A_SET"),     0, 1)
        if not names:
            popup.enableItem(4, False)

        # ── Middle: named sets ───────────────────────────────────────────────
        if names:
            popup.insertSeparator(2)
            for i, name in enumerate(names):
                popup.insertItem(100 + i, name, 0, 3 + i)
            base = 3 + len(names)
            popup.insertSeparator(base)
            popup.insertItem(2, _t("RENAME_A_SET"), 0, base + 1)
            popup.insertItem(3, _t("DELETE_A_SET"), 0, base + 2)
            foot = base + 3
        else:
            foot = 2

        # ── Bottom: session + keyboard toggle ────────────────────────────────
        popup.insertSeparator(foot)
        popup.insertItem(5, _t("RESTORE_LAST_SESSION"), 0, foot + 1)
        if not has_last:
            popup.enableItem(5, False)

        popup.insertItem(6, _t("TAB_KEY_SWITCHING"), 0, foot + 2)
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
        _pv("Label",      _t("SETS_BUTTON")),
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
                              _t("DLG_RENAME_DOC"),
                              _t("PROMPT_RENAME_DOC").format(name=basename),
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
        popup.insertItem(1, _t("RENAME"),          0, 0)
        popup.insertSeparator(1)
        popup.insertItem(2, _t("SAVE"),            0, 2)
        popup.insertItem(3, _t("SAVE_AS"),         0, 3)
        popup.insertSeparator(4)
        popup.insertItem(5, _t("MOVE_LEFT"),       0, 5)
        popup.insertItem(6, _t("MOVE_RIGHT"),      0, 6)
        popup.insertSeparator(7)
        popup.insertItem(4, _t("NEW_DOCUMENT"),    0, 8)
        popup.insertSeparator(9)
        popup.insertItem(7, _t("CLOSE"),           0, 10)
        popup.insertItem(8, _t("CLOSE_ALL_OTHERS"),0, 11)
        popup.insertItem(9, _t("CLOSE_ALL"),       0, 12)

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

    # Detect UI locale once at startup
    _detect_locale(ctx)
    _log(f"bootstrap: locale={_locale}")

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
