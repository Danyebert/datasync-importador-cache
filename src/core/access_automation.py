from __future__ import annotations

import re
import time
from pathlib import Path
from typing import Any, Callable

ProgressCallback = Callable[[str], None]


def _safe_query_name(nome: str, idx: int) -> str:
    base = str(nome or f"Consulta_{idx}").strip()
    base = re.sub(r"[^A-Za-z0-9_]+", "_", base)
    base = base.strip("_") or f"Consulta_{idx}"
    if not base.lower().startswith("imp_cache_"):
        base = f"Imp_Cache_{idx:02d}_{base}"
    return base[:60]


class AccessAutomation:
    """
    Executa consultas dentro do Microsoft Access usando automacao COM.

    Importante:
    - Funciona somente no Windows.
    - Exige Microsoft Access instalado.
    - Nao funciona apenas com o Access Database Engine/ODBC.
    """

    def __init__(self, visible: bool = False, progress: ProgressCallback | None = None):
        self.visible = visible
        self.progress = progress or (lambda msg: None)

    def _emit(self, msg: str) -> None:
        self.progress(msg)

    def _abrir_access(self):
        try:
            import win32com.client  # type: ignore
        except ImportError as exc:
            raise RuntimeError(
                "Biblioteca pywin32 nao encontrada. Instale com: pip install pywin32"
            ) from exc

        try:
            app = win32com.client.Dispatch("Access.Application")
        except Exception as exc:
            raise RuntimeError(
                "Nao foi possivel abrir o Microsoft Access via COM. "
                "Verifique se o Microsoft Access esta instalado neste computador."
            ) from exc

        app.Visible = bool(self.visible)
        return app

    @staticmethod
    def _apagar_query_se_existir(db: Any, nome_query: str) -> None:
        try:
            db.QueryDefs.Delete(nome_query)
        except Exception:
            pass

    def salvar_e_executar_consulta_aberta(self, app: Any, nome_query: str, sql: str) -> None:
        """Cria, executa e remove uma consulta usando um Access ja aberto."""
        db = app.CurrentDb()
        self._apagar_query_se_existir(db, nome_query)
        db.CreateQueryDef(nome_query, sql)
        db.QueryDefs.Refresh()

        self._emit(f"Executando dentro do Access: {nome_query}...")
        app.DoCmd.OpenQuery(nome_query)

        self._apagar_query_se_existir(db, nome_query)

    def salvar_e_executar_consulta(
        self,
        mdb_path: str | Path,
        nome_query: str,
        sql: str,
        fechar_access: bool = True,
    ) -> None:
        mdb = str(Path(mdb_path).resolve())
        app = None
        try:
            app = self._abrir_access()
            app.OpenCurrentDatabase(mdb)
            try:
                app.DoCmd.SetWarnings(False)
            except Exception:
                pass

            self.salvar_e_executar_consulta_aberta(app, nome_query, sql)

            try:
                app.DoCmd.SetWarnings(True)
            except Exception:
                pass

            if fechar_access:
                try:
                    app.CloseCurrentDatabase()
                except Exception:
                    pass
                app.Quit()
                app = None

            time.sleep(0.5)

        except Exception as exc:
            raise RuntimeError(f"Erro ao executar consulta dentro do Access ({nome_query}): {exc}") from exc
        finally:
            if app is not None:
                try:
                    app.DoCmd.SetWarnings(True)
                except Exception:
                    pass
                try:
                    app.CloseCurrentDatabase()
                except Exception:
                    pass
                try:
                    app.Quit()
                except Exception:
                    pass

    def executar_consultas_reabrindo_access(
        self,
        mdb_path: str | Path,
        consultas: list[dict[str, Any]],
    ) -> int:
        """Modo mais seguro: abre e fecha o Access a cada consulta."""
        executadas = 0
        total = len(consultas)

        for idx, item in enumerate(consultas, start=1):
            nome_original = item.get("NomeTabela") or item.get("Descri") or f"Consulta_{idx}"
            sql = str(item.get("Consulta") or "").strip()
            if not sql:
                continue

            nome_query = _safe_query_name(nome_original, idx)
            self._emit(f"PROGRESS_CONSULTA_INICIO|{idx}|{total}|{nome_original}")
            self._emit(f"Abrindo Access para consulta {idx}/{total}: {nome_original}...")
            inicio = time.time()
            self.salvar_e_executar_consulta(mdb_path, nome_query, sql, fechar_access=True)
            segundos = round(time.time() - inicio, 2)
            executadas += 1
            self._emit(f"PROGRESS_CONSULTA_OK|{executadas}|{total}|{nome_original}")
            self._emit(f"Consulta {idx}/{total} finalizada em {segundos}s. Access fechado.")

        return executadas

    def executar_consultas_unica_sessao(
        self,
        mdb_path: str | Path,
        consultas: list[dict[str, Any]],
        continuar_apos_erro: bool = False,
    ) -> int:
        """Modo desempenho: abre o Access uma vez e executa todas as consultas.

        Use este modo para testar performance. Se aparecer erro nativo do Access
        ou duplicidade de chave primaria, volte para o modo reabrindo_access.
        """
        mdb = str(Path(mdb_path).resolve())
        app = None
        executadas = 0
        erros: list[str] = []
        total = len(consultas)

        try:
            self._emit("Abrindo Microsoft Access uma unica vez...")
            app = self._abrir_access()
            app.OpenCurrentDatabase(mdb)
            try:
                app.DoCmd.SetWarnings(False)
            except Exception:
                pass

            for idx, item in enumerate(consultas, start=1):
                nome_original = item.get("NomeTabela") or item.get("Descri") or f"Consulta_{idx}"
                sql = str(item.get("Consulta") or "").strip()
                if not sql:
                    continue

                nome_query = _safe_query_name(nome_original, idx)
                inicio = time.time()
                self._emit(f"PROGRESS_CONSULTA_INICIO|{idx}|{total}|{nome_original}")
                self._emit(f"Executando consulta {idx}/{total}: {nome_original}...")

                try:
                    self.salvar_e_executar_consulta_aberta(app, nome_query, sql)
                    segundos = round(time.time() - inicio, 2)
                    executadas += 1
                    self._emit(f"PROGRESS_CONSULTA_OK|{executadas}|{total}|{nome_original}")
                    self._emit(f"Consulta {idx}/{total} concluida em {segundos}s.")
                except Exception as exc:
                    msg = f"Erro na consulta {idx}/{total} ({nome_original}): {exc}"
                    self._emit(msg)
                    erros.append(msg)
                    if not continuar_apos_erro:
                        raise RuntimeError(msg) from exc

            if erros:
                self._emit(f"Finalizado com {len(erros)} erro(s). Verifique o log.")
            else:
                self._emit("Todas as consultas foram executadas na mesma sessao do Access.")

            return executadas

        finally:
            if app is not None:
                try:
                    app.DoCmd.SetWarnings(True)
                except Exception:
                    pass
                try:
                    app.CloseCurrentDatabase()
                except Exception:
                    pass
                try:
                    app.Quit()
                except Exception:
                    pass
                self._emit("Microsoft Access fechado.")
