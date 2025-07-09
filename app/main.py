import sys
import importlib
import pkgutil
from pathlib import Path

import yaml
import ast
import textwrap
import traceback
from PySide6.QtCore import QThread, Signal
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QTextBrowser,
    QHBoxLayout,
    QVBoxLayout,
    QWidget,
)

# Determine where the bundled scripts live:
if getattr(sys, "frozen", False):  # running from a PyInstaller EXE
    BASE_DIR = Path(sys._MEIPASS) / "app" / "scripts"
else:                              # running from source
    BASE_DIR = Path(__file__).parent / "scripts"

SCRIPTS_DIR = BASE_DIR


def discover_plugins():
    """
    Yield ``(meta_dict, path_to_file)`` pairs for every plug‑in script
    without importing the module code.

    A plug‑in is any ``*.py`` file in *SCRIPTS_DIR* whose module
    doc‑string contains a YAML header with at least a ``name`` field.
    """
    for file in SCRIPTS_DIR.glob("*.py"):
        if file.name in ("__init__.py",) or file.name.startswith("_"):
            continue
        try:
            # extract the module‑level doc‑string w/o executing code
            with file.open("r", encoding="utf-8") as fh:
                tree = ast.parse(fh.read(), filename=str(file))
            doc = ast.get_docstring(tree) or ""
            meta = yaml.safe_load(textwrap.dedent(doc))
            if isinstance(meta, dict) and "name" in meta:
                yield meta, file
        except Exception as exc:
            print(f"[WARN] Skipping {file.name}: {exc}")


class Runner(QThread):
    finished = Signal(object)        # emits script result (e.g., Path)

    def __init__(self, func, kwargs):
        super().__init__()
        self.func, self.kwargs = func, kwargs

    def run(self):
        try:
            result = self.func(**self.kwargs)
            # success: pack into a tuple so the handler can distinguish
            self.finished.emit(("ok", result))
        except Exception:  # catch *everything* so GUI won't silently die
            tb = traceback.format_exc()
            self.finished.emit(("error", tb))


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Data-Ops Toolbox")

        # widgets
        self.listbox = QListWidget()
        self.output  = QTextBrowser()
        self.run_btn = QPushButton("Run")

        # layout
        right = QVBoxLayout()
        right.addWidget(self.output, 1)
        right.addWidget(self.run_btn, 0)

        root = QHBoxLayout(self)
        root.addWidget(self.listbox, 1)
        root.addLayout(right, 2)

        # discover plug-ins
        self.plugins = {m["name"]: (m, path) for m, path in discover_plugins()}
        for name in self.plugins:
            self.listbox.addItem(QListWidgetItem(name))

        # wire signals
        self.listbox.currentItemChanged.connect(self._choose)
        self.run_btn.clicked.connect(self._run)

        # state
        self._current_name   = None
        self._selected_files = {}
        self._workers        = []   # keep QThread objects alive while they run

    def _choose(self, item):
        if not item:
            return
        name = item.text()
        self._current_name = name
        meta, _ = self.plugins[name]

        self.output.clear()
        self.output.append(f"<b>{meta['description']}</b><br/>")

        # list the files that will be requested
        self.output.append("<u>Files required:</u>")
        for label in meta["required_files"].values():
            self.output.append(f"• {label}")
        self.output.append("<i>Click Run to choose them</i><br/>")

        # reset any previous selections
        self._selected_files = {}

    def _run(self):
        if not self._current_name:
            return
        meta, path = self.plugins[self._current_name]
        loader = importlib.machinery.SourceFileLoader(path.stem, str(path))  # pass str, not Path
        spec   = importlib.util.spec_from_loader(path.stem, loader)
        mod    = importlib.util.module_from_spec(spec)
        loader.exec_module(mod)
        fn     = mod.main

        # pick the required files now
        self._selected_files = {}
        for key, label in meta["required_files"].items():
            fname, _ = QFileDialog.getOpenFileName(self, f"Select {label}")
            if not fname:
                self.output.append(f"<i>Cancelled – missing {label}</i>")
                return
            self._selected_files[key] = Path(fname)
            self.output.append(f"{label}: {fname}")

        self.run_btn.setEnabled(False)
        self.output.append("\nRunning …")

        worker = Runner(fn, self._selected_files)
        self._workers.append(worker)            # <‑‑ keep reference
        worker.finished.connect(lambda res, w=worker: self._done(res, w))
        worker.start()

    def _done(self, result, worker):
        status, payload = result
        if status == "ok":
            self.output.append(f"<br/><b>✓ Done → {payload}</b>")
        else:  # ("error", traceback)
            self.output.append(
                "<br/><span style='color:red'><b>✗ Script crashed</b></span>"
                f"<pre>{payload}</pre>"
            )
        self.run_btn.setEnabled(True)
        if worker in self._workers:
            self._workers.remove(worker)
        worker.deleteLater()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    mw  = MainWindow()
    mw.resize(900, 500)
    mw.show()
    sys.exit(app.exec())