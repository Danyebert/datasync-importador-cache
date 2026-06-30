from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import Any, Iterable


class SQLiteDB:
    def __init__(self, path: str | Path):
        self.path = Path(path)
        self.conn: sqlite3.Connection | None = None

    def conectar(self) -> None:
        if not self.path.exists():
            raise FileNotFoundError(f"Banco SQLite não encontrado: {self.path}")
        self.conn = sqlite3.connect(str(self.path))
        self.conn.row_factory = sqlite3.Row

    def fechar(self) -> None:
        if self.conn:
            self.conn.close()
            self.conn = None

    def listar_tabelas(self) -> list[str]:
        assert self.conn is not None
        cur = self.conn.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
        return [r[0] for r in cur.fetchall()]

    def consultar(self, query: str) -> Iterable[dict[str, Any]]:
        assert self.conn is not None
        cur = self.conn.execute(query)
        for row in cur:
            yield dict(row)
