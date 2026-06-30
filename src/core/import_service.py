from __future__ import annotations

import json
from pathlib import Path
from typing import Callable

from .sqlite_db import SQLiteDB
from .access_db import AccessDB
from .access_automation import AccessAutomation

ProgressCallback = Callable[[str], None]


class ImportCacheService:
    def __init__(self, base_dir: Path, progress: ProgressCallback | None = None):
        self.base_dir = base_dir
        self.progress = progress or (lambda msg: None)

    def _emit(self, msg: str) -> None:
        self.progress(msg)

    def executar(self, sqlite_path: str, mdb_path: str, config_path: str) -> dict[str, int | str]:
        config_file = Path(config_path)
        if not config_file.exists():
            raise FileNotFoundError(f"Configuração não encontrada: {config_file}")

        config = json.loads(config_file.read_text(encoding="utf-8"))
        batch_size = int(config.get("batch_size", 500))
        modo = config.get("modo_importacao", "replace")

        sqlite_db = SQLiteDB(sqlite_path)
        access_db = AccessDB(mdb_path)
        total_por_tabela: dict[str, int | str] = {}

        try:
            self._emit("Conectando no SQLite...")
            sqlite_db.conectar()

            tabelas_sqlite = set(sqlite_db.listar_tabelas())
            obrigatorias = config.get("validar_tabelas_sqlite", [])
            faltando = [t for t in obrigatorias if t not in tabelas_sqlite]
            if faltando:
                raise RuntimeError(
                    "Banco SQLite inválido. Tabelas obrigatórias não encontradas: "
                    + ", ".join(faltando)
                )

            self._emit("Conectando no Access MDB...")
            access_db.conectar()

            for tabela_cfg in config["tabelas"]:
                tabela = tabela_cfg["destino_access"]
                query = tabela_cfg["origem_sqlite_query"]
                campos: dict[str, str] = tabela_cfg["campos"]
                campos_lista = list(campos.keys())

                self._emit(f"Preparando tabela {tabela}...")
                access_db.criar_tabela_se_nao_existe(tabela, campos)
                if modo == "replace":
                    access_db.limpar_tabela(tabela)

                self._emit(f"Lendo SQLite e inserindo em {tabela}...")
                lote = []
                total = 0
                for row in sqlite_db.consultar(query):
                    lote.append(row)
                    if len(lote) >= batch_size:
                        total += access_db.inserir_lote(tabela, campos_lista, lote)
                        self._emit(f"{tabela}: {total} registros importados...")
                        lote.clear()

                if lote:
                    total += access_db.inserir_lote(tabela, campos_lista, lote)

                total_por_tabela[tabela] = total
                self._emit(f"{tabela}: finalizado com {total} registros.")

            if config.get("executar_consultas_importcache", True):
                modulo = config.get("modulo_lista_tabelas", "ImportCache")
                somente_marcadas = bool(config.get("somente_consultas_marcadas_para_executar", False))
                fonte_consultas = config.get("fonte_consultas_importcache", "template")
                template_mdb = self.base_dir / config.get("template_mdb_consultas", "templates/DataSync - Recovery.mdb")

                consultas_db = access_db
                template_db: AccessDB | None = None

                if fonte_consultas == "template":
                    if not template_mdb.exists():
                        raise FileNotFoundError(f"MDB modelo com as consultas não encontrado: {template_mdb}")
                    self._emit(f"Lendo consultas fixas do projeto: {template_mdb}")
                    template_db = AccessDB(template_mdb)
                    template_db.conectar()
                    consultas_db = template_db
                else:
                    self._emit("Lendo consultas do próprio MDB selecionado.")

                try:
                    self._emit(f"Buscando consultas da ListaTabelas para Modulo={modulo}...")
                    consultas = consultas_db.buscar_consultas_lista_tabelas(
                        modulo=modulo,
                        somente_marcadas=somente_marcadas,
                    )
                finally:
                    if template_db:
                        template_db.fechar()

                total_por_tabela["consultas_importcache_encontradas"] = len(consultas)
                self._emit(f"{len(consultas)} consultas encontradas.")

                # As consultas do ImportCache podem depender de funcoes VBA do Access
                # como Util_nz. Por isso elas NAO sao executadas via pyodbc.
                # Primeiro gravamos e fechamos o MDB; depois abrimos o Microsoft Access
                # por automacao COM, salvamos a consulta dentro do proprio Access,
                # executamos, fechamos o Access e repetimos ate terminar.
                self._emit("Gravando cache no MDB antes de executar consultas dentro do Access...")
                access_db.commit()
                access_db.fechar()

                executar_dentro_access = bool(config.get("executar_consultas_via_access_com", True))
                if not executar_dentro_access:
                    raise RuntimeError(
                        "A execucao via pyodbc foi desativada porque as consultas podem usar funcoes VBA. "
                        "Ative executar_consultas_via_access_com no JSON."
                    )

                access_visible = bool(config.get("access_visivel_durante_execucao", False))
                modo_execucao_access = str(config.get("modo_execucao_access", "unica_sessao")).strip().lower()
                continuar_apos_erro = bool(config.get("continuar_apos_erro_consulta", False))

                automation = AccessAutomation(visible=access_visible, progress=self._emit)

                if modo_execucao_access == "reabrindo_access":
                    self._emit("Modo de execucao Access: seguro, abrindo e fechando a cada consulta.")
                    executadas = automation.executar_consultas_reabrindo_access(mdb_path, consultas)
                else:
                    self._emit("Modo de execucao Access: desempenho, uma unica sessao do Access.")
                    executadas = automation.executar_consultas_unica_sessao(
                        mdb_path,
                        consultas,
                        continuar_apos_erro=continuar_apos_erro,
                    )

                total_por_tabela["consultas_importcache_executadas"] = executadas
                self._emit("Consultas ImportCache executadas dentro do Microsoft Access.")
                self._emit("Importação concluída com sucesso.")
                return total_por_tabela

            access_db.commit()
            self._emit("Importação concluída com sucesso.")
            return total_por_tabela

        except Exception:
            try:
                access_db.rollback()
            except Exception:
                pass
            raise
        finally:
            sqlite_db.fechar()
            access_db.fechar()
