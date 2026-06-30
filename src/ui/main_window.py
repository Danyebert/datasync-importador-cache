from __future__ import annotations

import sys
import time
import winreg
from pathlib import Path

from PySide6.QtCore import Qt, QThread, Signal, QTimer
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFileDialog, QLabel, QPushButton,
    QVBoxLayout, QHBoxLayout, QFrame, QProgressBar, QComboBox, QMessageBox,
    QGridLayout
)

from src.core.import_service import ImportCacheService


APP_VERSION = "1.4.1"
BASE_DIR = Path(__file__).resolve().parents[2]
CONFIG_PATH = BASE_DIR / "config" / "import_cache_mapping.json"


def windows_dark_mode() -> bool:
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        )
        value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
        winreg.CloseKey(key)
        return value == 0
    except Exception:
        return False


class ImportWorker(QThread):
    log = Signal(str)
    finished_ok = Signal(dict)
    failed = Signal(str)

    def __init__(self, sqlite_path: str, mdb_path: str):
        super().__init__()
        self.sqlite_path = sqlite_path
        self.mdb_path = mdb_path

    def run(self):
        try:
            service = ImportCacheService(BASE_DIR, progress=self.log.emit)
            result = service.executar(
                sqlite_path=self.sqlite_path,
                mdb_path=self.mdb_path,
                config_path=str(CONFIG_PATH)
            )
            self.finished_ok.emit(result)
        except Exception as e:
            self.failed.emit(str(e))


class Card(QFrame):
    def __init__(self, titulo: str, texto_botao: str, callback):
        super().__init__()
        self.setObjectName("card")

        self.title = QLabel(titulo)
        self.title.setObjectName("cardTitle")

        self.status = QLabel("Não selecionado")
        self.status.setObjectName("cardStatus")

        self.path = QLabel("-")
        self.path.setObjectName("cardPath")
        self.path.setWordWrap(True)

        self.button = QPushButton(texto_botao)
        self.button.clicked.connect(callback)
        self.button.setObjectName("cardButton")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(12)
        layout.addWidget(self.title)
        layout.addStretch()
        layout.addWidget(self.status)
        layout.addWidget(self.path)
        layout.addStretch()
        layout.addWidget(self.button)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.sqlite_path = ""
        self.mdb_path = ""
        self.total_consultas = 0
        self.consultas_executadas = 0
        self.worker = None
        self.started_at: float | None = None

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_elapsed_time)

        self.setWindowTitle("DataSync - Dashboard")
        self.resize(760, 780)

        self.build_ui()
        self.apply_theme("Sistema")
        self.reset_steps()

    def build_ui(self):
        root = QWidget()
        self.setCentralWidget(root)

        main = QVBoxLayout(root)
        main.setContentsMargins(24, 18, 24, 14)
        main.setSpacing(16)

        top = QHBoxLayout()
        title = QLabel("MODELO 2 - DASHBOARD")
        title.setObjectName("title")

        self.theme_combo = QComboBox()
        self.theme_combo.addItems(["Sistema", "Claro", "Escuro"])
        self.theme_combo.currentTextChanged.connect(self.apply_theme)

        top.addStretch()
        top.addWidget(title)
        top.addStretch()
        top.addWidget(self.theme_combo)
        main.addLayout(top)

        cards = QHBoxLayout()
        cards.setSpacing(20)

        self.sqlite_card = Card("SQLite", "Selecionar SQLite", self.select_sqlite)
        self.mdb_card = Card("Access MDB", "Selecionar MDB", self.select_mdb)

        arrow = QLabel("→")
        arrow.setObjectName("arrow")
        arrow.setAlignment(Qt.AlignCenter)

        cards.addWidget(self.sqlite_card)
        cards.addWidget(arrow)
        cards.addWidget(self.mdb_card)
        main.addLayout(cards)

        self.btn_import = QPushButton("▶  IMPORTAR AGORA")
        self.btn_import.setObjectName("primaryButton")
        self.btn_import.clicked.connect(self.start_import)
        main.addWidget(self.btn_import)

        progress_header = QHBoxLayout()
        progress_label = QLabel("Progresso da Importação")
        progress_label.setObjectName("sectionTitle")
        self.percent_label = QLabel("0%")
        self.percent_label.setObjectName("percent")
        progress_header.addWidget(progress_label)
        progress_header.addStretch()
        progress_header.addWidget(self.percent_label)
        main.addLayout(progress_header)

        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.setTextVisible(False)
        main.addWidget(self.progress)

        info = QHBoxLayout()
        self.query_label = QLabel("Consulta 0 de 0")
        self.current_label = QLabel("Aguardando importação...")
        self.time_label = QLabel("Tempo decorrido: 00:00:00")
        self.time_label.setAlignment(Qt.AlignRight)
        info.addWidget(self.query_label, 1)
        info.addWidget(self.current_label, 2)
        info.addWidget(self.time_label, 1)
        main.addLayout(info)

        self.steps_box = QFrame()
        self.steps_box.setObjectName("stepsBox")
        steps_layout = QVBoxLayout(self.steps_box)

        self.steps: dict[str, tuple[QLabel, QLabel]] = {}
        for key, nome in [
            ("clientes", "Clientes"),
            ("produtos", "Produtos"),
            ("tributacao", "Tributação"),
            ("lojas", "Lojas e Preços"),
            ("finalizacao", "Finalização"),
        ]:
            row = QFrame()
            row.setObjectName("stepRow")
            row_layout = QHBoxLayout(row)
            row_layout.setContentsMargins(8, 6, 8, 6)

            left = QLabel(f"○  {nome}")
            left.setObjectName("step")

            right = QLabel("Aguardando")
            right.setObjectName("stepStatus")
            right.setAlignment(Qt.AlignRight | Qt.AlignVCenter)

            row_layout.addWidget(left, 1)
            row_layout.addWidget(right, 0)
            steps_layout.addWidget(row)
            self.steps[key] = (left, right)

        main.addWidget(self.steps_box)

        self.log_box = QLabel("")
        self.log_box.setObjectName("logBox")
        self.log_box.setWordWrap(True)
        self.log_box.setMinimumHeight(120)
        main.addWidget(self.log_box)

        self.btn_cancel = QPushButton("■  Cancelar Importação")
        self.btn_cancel.setObjectName("cancelButton")
        self.btn_cancel.setEnabled(False)
        main.addWidget(self.btn_cancel)

        footer = QHBoxLayout()
        self.version_label = QLabel(f"DataSync v{APP_VERSION}")
        self.status_label = QLabel("Pronto")
        footer.addWidget(self.version_label)
        footer.addStretch()
        footer.addWidget(self.status_label)
        main.addLayout(footer)

    def reset_steps(self):
        for key, (left, right) in self.steps.items():
            name = left.text().replace("✔  ", "").replace("◉  ", "").replace("○  ", "")
            left.setText(f"○  {name}")
            right.setText("Aguardando")

    def set_step(self, key: str, status: str):
        if key not in self.steps:
            return

        left, right = self.steps[key]
        name = left.text().replace("✔  ", "").replace("◉  ", "").replace("○  ", "")

        if status == "concluido":
            left.setText(f"✔  {name}")
            right.setText("Concluído")
        elif status == "executando":
            left.setText(f"◉  {name}")
            right.setText("Executando")
        else:
            left.setText(f"○  {name}")
            right.setText("Aguardando")

    def select_sqlite(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Selecionar banco SQLite",
            "",
            "SQLite (*.db *.sqlite *.sqlite3)"
        )
        if path:
            self.sqlite_path = path
            self.sqlite_card.status.setText("✔ Conectado")
            self.sqlite_card.path.setText(path)

    def select_mdb(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Selecionar MDB",
            "",
            "Access MDB (*.mdb)"
        )
        if path:
            self.mdb_path = path
            self.mdb_card.status.setText("✔ Conectado")
            self.mdb_card.path.setText(path)

    def start_import(self):
        if not self.sqlite_path or not self.mdb_path:
            QMessageBox.warning(self, "Atenção", "Selecione o SQLite e o MDB antes de importar.")
            return

        self.total_consultas = 0
        self.consultas_executadas = 0
        self.started_at = time.time()
        self.timer.start(1000)
        self.reset_steps()

        self.progress.setValue(0)
        self.percent_label.setText("0%")
        self.query_label.setText("Consulta 0 de 0")
        self.current_label.setText("Iniciando...")
        self.time_label.setText("Tempo decorrido: 00:00:00")
        self.status_label.setText("Importando...")
        self.log_box.setText("Iniciando importação...")

        self.btn_import.setEnabled(False)
        self.btn_cancel.setEnabled(True)

        self.worker = ImportWorker(self.sqlite_path, self.mdb_path)
        self.worker.log.connect(self.handle_log)
        self.worker.finished_ok.connect(self.import_finished)
        self.worker.failed.connect(self.import_failed)
        self.worker.start()

    def handle_log(self, msg: str):
        self.log_box.setText(msg)
        msg_lower = msg.lower()

        if "consultas encontradas" in msg_lower:
            try:
                self.total_consultas = int(msg.split()[0])
                self.update_progress()
            except Exception:
                pass

        if "executando consulta" in msg_lower:
            self.consultas_executadas += 1
            self.current_label.setText(msg)
            self.set_step("lojas", "executando")
            self.update_progress()

        if "cliente" in msg_lower and ("finalizado" in msg_lower or "importado" in msg_lower):
            self.set_step("clientes", "concluido")

        if ("mercadorias" in msg_lower or "produtos" in msg_lower) and ("finalizado" in msg_lower or "importado" in msg_lower):
            self.set_step("produtos", "concluido")

        if ("tributacao" in msg_lower or "tributação" in msg_lower) and ("finalizado" in msg_lower or "importado" in msg_lower):
            self.set_step("tributacao", "concluido")

        if "lojas" in msg_lower and ("finalizado" in msg_lower or "importado" in msg_lower):
            self.set_step("lojas", "concluido")

    def update_progress(self):
        if self.total_consultas <= 0:
            return

        percent = int((self.consultas_executadas / self.total_consultas) * 100)
        percent = max(0, min(percent, 100))
        self.progress.setValue(percent)
        self.percent_label.setText(f"{percent}%")
        self.query_label.setText(f"Consulta {self.consultas_executadas} de {self.total_consultas}")

    def update_elapsed_time(self):
        if not self.started_at:
            return
        elapsed = int(time.time() - self.started_at)
        h = elapsed // 3600
        m = (elapsed % 3600) // 60
        s = elapsed % 60
        self.time_label.setText(f"Tempo decorrido: {h:02d}:{m:02d}:{s:02d}")

    def import_finished(self, result: dict):
        self.timer.stop()
        self.update_elapsed_time()

        if self.total_consultas <= 0:
            self.total_consultas = int(result.get("consultas_importcache_encontradas", 0) or 0)

        executadas = int(result.get("consultas_importcache_executadas", self.total_consultas) or self.total_consultas)
        self.consultas_executadas = max(self.consultas_executadas, executadas)

        if self.total_consultas > 0:
            self.consultas_executadas = self.total_consultas
            self.query_label.setText(f"Consulta {self.total_consultas} de {self.total_consultas}")

        self.progress.setValue(100)
        self.percent_label.setText("100%")

        for key in ["clientes", "produtos", "tributacao", "lojas", "finalizacao"]:
            self.set_step(key, "concluido")

        self.current_label.setText("Importação finalizada")
        self.status_label.setText("Pronto")
        self.btn_import.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        self.log_box.setText("Importação concluída com sucesso.")
        QMessageBox.information(self, "Finalizado", "Importação concluída com sucesso.")

    def import_failed(self, error: str):
        self.timer.stop()
        self.update_elapsed_time()
        self.status_label.setText("Erro")
        self.btn_import.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        QMessageBox.critical(self, "Erro", error)

    def apply_theme(self, theme: str):
        if theme == "Sistema":
            dark = windows_dark_mode()
        else:
            dark = theme == "Escuro"

        self.setStyleSheet(DARK_STYLE if dark else LIGHT_STYLE)


LIGHT_STYLE = """
QMainWindow { background: #f6f7fb; }
QWidget { font-family: Segoe UI; font-size: 14px; color: #101828; }
#title { font-size: 26px; font-weight: 700; letter-spacing: 1px; }
#card { background: white; border: 1px solid #d9dee8; border-radius: 12px; min-height: 170px; }
#cardTitle { font-size: 19px; font-weight: 700; }
#cardStatus { color: #0a8f28; font-size: 17px; font-weight: 700; }
#cardPath { color: #344054; }
#arrow { font-size: 48px; color: #101828; }
QPushButton { background: white; border: 1px solid #d0d5dd; border-radius: 8px; padding: 10px 16px; }
QPushButton:hover { background: #eef4ff; }
QPushButton:disabled { color: #98a2b3; background: #eef2f7; }
#cardButton { background:#ffffff; border:1px solid #d0d5dd; border-radius:8px; padding:10px; font-size:14px; }
#cardButton:hover { background:#eef4ff; }
#primaryButton { background: #0d6efd; color: white; font-size: 18px; font-weight: 700; border: none; padding: 16px; }
#cancelButton { color: #c1121f; border: 1px solid #ef233c; background: white; font-weight: 600; padding: 14px; }
#sectionTitle { font-size: 17px; font-weight: 700; }
#percent { font-size: 24px; font-weight: 800; }
QProgressBar { background: #e9edf5; border-radius: 7px; height: 18px; }
QProgressBar::chunk { background: #0d6efd; border-radius: 7px; }
#stepsBox { background: white; border: 1px solid #d9dee8; border-radius: 12px; padding: 10px; }
#step, #stepStatus { font-size: 16px; padding: 4px; }
#stepStatus { font-weight: 600; }
#logBox { background: white; border: 1px solid #d9dee8; border-radius: 10px; padding: 12px; color: #344054; }
"""


DARK_STYLE = """
QMainWindow { background: #0f172a; }
QWidget { font-family: Segoe UI; font-size: 14px; color: #e5e7eb; }
#title { font-size: 26px; font-weight: 700; letter-spacing: 1px; color: #f8fafc; }
#card { background: #111827; border: 1px solid #334155; border-radius: 12px; min-height: 170px; }
#cardTitle { font-size: 19px; font-weight: 700; color: #f8fafc; }
#cardStatus { color: #22c55e; font-size: 17px; font-weight: 700; }
#cardPath { color: #cbd5e1; }
#arrow { font-size: 48px; color: #f8fafc; }
QPushButton { background: #1e293b; border: 1px solid #475569; border-radius: 8px; padding: 10px 16px; color: #f8fafc; }
QPushButton:hover { background: #334155; }
QPushButton:disabled { color: #64748b; background: #1e293b; }
#cardButton { background:#1e293b; border:1px solid #475569; border-radius:8px; padding:10px; color:white; }
#cardButton:hover { background:#334155; }
#primaryButton { background: #2563eb; color: white; font-size: 18px; font-weight: 700; border: none; padding: 16px; }
#cancelButton { color: #f87171; border: 1px solid #ef4444; background: #1e293b; font-weight: 600; padding: 14px; }
#sectionTitle { font-size: 17px; font-weight: 700; }
#percent { font-size: 24px; font-weight: 800; }
QProgressBar { background: #020617; border-radius: 7px; height: 18px; }
QProgressBar::chunk { background: #2563eb; border-radius: 7px; }
#stepsBox { background: #111827; border: 1px solid #334155; border-radius: 12px; padding: 10px; }
#step, #stepStatus { font-size: 16px; padding: 4px; }
#stepStatus { font-weight: 600; }
#logBox { background: #111827; border: 1px solid #334155; border-radius: 10px; padding: 12px; color: #cbd5e1; }
"""


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
