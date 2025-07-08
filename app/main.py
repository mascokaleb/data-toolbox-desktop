# app/main.py
from pathlib import Path
import importlib, sys, yaml

from PySide6.QtWidgets import (
    QApplication, QWidget, QListWidget, QTextBrowser, QPushButton,
    QFileDialog, QHBoxLayout, QVBoxLayout, QListWidgetItem
)
from PySide6.QtCore import QThread, Signal

SCRIPTS_DIR = Path(__file__).parent / "scripts"


def discover_plugins():
    """Yield (meta_dict, main_func) for every *.py under app/scripts.

    Malformed YAML headers are skipped but logged to stdout.
    """
    for src in SCRIPTS_DIR.glob("*.py"):
        try:
            header = src.read_text().split('"""')[1]
            meta   = yaml.safe_load(header)
            mod    = importlib.import_module(f"app.scripts.{src.stem}")
            yield meta, mod.main
        except (IndexError, yaml.YAMLError) as e:
            print(f"[WARN] Skipping {src.name}: bad YAML header → {e}")
        except Exception as e:
            print(f"[WARN] Skipping {src.name}: {e}")


class Runner(QThread):
    finished = Signal(object)        # emits script result (e.g., Path)

    def __init__(self, func, kwargs):
        super().__init__()
        self.func, self.kwargs = func, kwargs

    def run(self):
        self.finished.emit(self.func(**self.kwargs))


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
        self.plugins = {m["name"]: (m, fn) for m, fn in discover_plugins()}
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
        meta, fn = self.plugins[self._current_name]

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
        self.output.append(f"<br/><b>✓ Done → {result}</b>")
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