from __future__ import annotations

from pathlib import Path
from typing import Any
import re
import pyodbc


class AccessDB:
    def __init__(self, path: str | Path):
        self.path = Path(path)
        self.conn: pyodbc.Connection | None = None

    @staticmethod
    def drivers_access() -> list[str]:
        return [d for d in pyodbc.drivers() if "Access" in d or "Microsoft Access" in d]

    @classmethod
    def validar_driver(cls) -> str:
        drivers = cls.drivers_access()
        preferidos = [
            "Microsoft Access Driver (*.mdb, *.accdb)",
            "Microsoft Access Driver (*.mdb)",
        ]
        for driver in preferidos:
            if driver in drivers:
                return driver
        if drivers:
            return drivers[0]
        raise RuntimeError(
            "Driver ODBC do Microsoft Access não encontrado. Instale o Microsoft Access Database Engine/Runtime "
            "com a mesma arquitetura do Python: 32 bits com 32 bits ou 64 bits com 64 bits."
        )

    def conectar(self) -> None:
        if not self.path.exists():
            raise FileNotFoundError(f"Arquivo MDB não encontrado: {self.path}")
        driver = self.validar_driver()
        conn_str = f"DRIVER={{{driver}}};DBQ={self.path};"
        self.conn = pyodbc.connect(conn_str, autocommit=False)

    def fechar(self) -> None:
        if self.conn:
            self.conn.close()
            self.conn = None

    def commit(self) -> None:
        assert self.conn is not None
        self.conn.commit()

    def rollback(self) -> None:
        assert self.conn is not None
        self.conn.rollback()

    def tabela_existe(self, tabela: str) -> bool:
        assert self.conn is not None
        cur = self.conn.cursor()
        for row in cur.tables(table=tabela, tableType="TABLE"):
            if row.table_name.lower() == tabela.lower():
                return True
        return False

    def listar_colunas(self, tabela: str) -> list[str]:
        assert self.conn is not None
        cur = self.conn.cursor()
        return [row.column_name for row in cur.columns(table=tabela)]

    def executar(self, sql: str, params: tuple[Any, ...] | None = None) -> None:
        assert self.conn is not None
        cur = self.conn.cursor()
        cur.execute(sql, params or ())

    def executar_sql(self, sql: str) -> int:
        """Executa uma consulta/action query do Access e retorna rowcount quando disponível."""
        assert self.conn is not None
        cur = self.conn.cursor()
        cur.execute(sql)
        try:
            return int(cur.rowcount or 0)
        except Exception:
            return 0

    def criar_tabela_se_nao_existe(self, tabela: str, campos: dict[str, str]) -> None:
        if self.tabela_existe(tabela):
            return
        definicoes = ", ".join(f"[{campo}] {tipo}" for campo, tipo in campos.items())
        self.executar(f"CREATE TABLE [{tabela}] ({definicoes})")

    def limpar_tabela(self, tabela: str) -> None:
        self.executar(f"DELETE FROM [{tabela}]")

    def inserir_lote(self, tabela: str, campos: list[str], rows: list[dict[str, Any]]) -> int:
        if not rows:
            return 0
        placeholders = ", ".join("?" for _ in campos)
        campos_sql = ", ".join(f"[{c}]" for c in campos)
        sql = f"INSERT INTO [{tabela}] ({campos_sql}) VALUES ({placeholders})"
        valores = [tuple(row.get(c) for c in campos) for row in rows]
        cur = self.conn.cursor()  # type: ignore[union-attr]
        cur.fast_executemany = False
        cur.executemany(sql, valores)
        return len(rows)

    def buscar_consultas_lista_tabelas(self, modulo: str = "ImportCache", somente_marcadas: bool = False) -> list[dict[str, Any]]:
        """
        Lê a tabela ListaTabelas do MDB e retorna as consultas do módulo informado.
        Campos esperados no MDB: Ordem, NomeTabela, Descri, Grupo, Consulta, Executar, Modulo.
        """
        if not self.tabela_existe("ListaTabelas"):
            raise RuntimeError("A tabela ListaTabelas não foi encontrada no MDB.")

        colunas = {c.lower(): c for c in self.listar_colunas("ListaTabelas")}
        if "modulo" not in colunas:
            raise RuntimeError("A tabela ListaTabelas não possui o campo Modulo.")
        if "consulta" not in colunas:
            raise RuntimeError("A tabela ListaTabelas não possui o campo Consulta.")

        campos_select = []
        for campo in ["Ordem", "NomeTabela", "Descri", "Grupo", "Consulta", "Executar", "Executado", "Modulo"]:
            real = colunas.get(campo.lower())
            if real:
                campos_select.append(f"[{real}]")

        where = f"WHERE [{colunas['modulo']}] = ? AND [{colunas['consulta']}] IS NOT NULL"
        params: list[Any] = [modulo]

        if somente_marcadas and "executar" in colunas:
            where += f" AND [{colunas['executar']}] = True"

        order_by = ""
        if "ordem" in colunas:
            order_by = f" ORDER BY [{colunas['ordem']}]"

        sql = f"SELECT {', '.join(campos_select)} FROM [ListaTabelas] {where}{order_by}"
        cur = self.conn.cursor()  # type: ignore[union-attr]
        rows = cur.execute(sql, tuple(params)).fetchall()
        columns = [c[0] for c in cur.description]

        consultas: list[dict[str, Any]] = []
        for row in rows:
            item = dict(zip(columns, row))
            consulta = item.get("Consulta")
            if consulta and str(consulta).strip():
                consultas.append(item)
        return consultas

    @staticmethod
    def normalizar_sql_access(sql: str) -> str:
        """Remove caracteres finais comuns que podem atrapalhar o pyodbc."""
        texto = str(sql).strip()
        texto = re.sub(r"\s+WITH\s+OWNERACCESS\s+OPTION\s*;?\s*$", "", texto, flags=re.IGNORECASE)
        texto = texto.rstrip("; \r\n\t")
        return texto

    def marcar_lista_tabelas_executado(self, nome_tabela: str | None, executado: bool = True) -> None:
        """Marca o registro como executado quando o MDB possuir os campos necessários."""
        if not nome_tabela or not self.tabela_existe("ListaTabelas"):
            return
        colunas = {c.lower(): c for c in self.listar_colunas("ListaTabelas")}
        if "executado" not in colunas or "nometabela" not in colunas:
            return
        valor = -1 if executado else 0
        self.executar(
            f"UPDATE [ListaTabelas] SET [{colunas['executado']}] = ? WHERE [{colunas['nometabela']}] = ?",
            (valor, nome_tabela),
        )
