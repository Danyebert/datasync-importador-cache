# DataSync - Importar Cache SQLite para Access MDB

Projeto desktop em Python para importar dados de um banco **SQLite** para o **MDB do sistema Access 2003**.

## O que mudou nesta versão

Esta versão já leva dentro do projeto o arquivo modelo:

```text
templates/DataSync - Recovery.mdb
```

Esse MDB modelo é usado apenas como fonte das consultas da tabela:

```text
ListaTabelas
```

O sistema busca automaticamente tudo que estiver com:

```sql
Modulo = 'ImportCache'
```

Depois executa essas consultas no MDB selecionado pelo usuário.

## Fluxo da importação

```text
1. Usuário seleciona o banco SQLite
2. Usuário seleciona o MDB do sistema
3. O sistema inicia automaticamente
4. Copia o SQLite para as tabelas temporárias dbo_A_CAIXA_*
5. Lê as consultas ImportCache do MDB modelo interno
6. Executa as consultas no MDB do sistema pelo Microsoft Access
7. Atualiza a barra de progresso conforme as consultas são executadas
8. Mostra o resumo no log
```

## Tabelas importadas do SQLite para o MDB

```text
SQLite.A_CAIXA_CLIENTE                    -> MDB.dbo_A_CAIXA_CLIENTE
SQLite.A_CAIXA_MERCADORIAS                -> MDB.dbo_A_CAIXA_MERCADORIAS
SQLite.A_CAIXA_MERCADORIAS_TRIBUTACAO     -> MDB.dbo_A_CAIXA_MERCADORIAS_TRIBUTACAO
SQLite.A_CAIXA_MERCADORIAS_LOJAS          -> MDB.dbo_A_CAIXA_MERCADORIAS_LOJAS
```

## Requisitos

- Windows
- Python 3.11 ou 3.12
- Driver ODBC do Microsoft Access instalado
- A arquitetura precisa bater:
  - Python 32 bits + Access Driver 32 bits
  - Python 64 bits + Access Driver 64 bits

## Instalação

```bat
python -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
```

## Executar

```bat
python app.py
```

## Como usar

1. Abra o sistema.
2. Selecione o banco SQLite, por exemplo `dados.db`.
3. Selecione o arquivo `.mdb` do sistema.
4. A importação começa automaticamente.

A importação inicia automaticamente quando os dois arquivos são selecionados.

## Progresso real

A barra de progresso acompanha a execução real das consultas ImportCache.
O sistema atualiza o contador no formato:

```text
Consulta atual / Total de consultas
```

## Configuração técnica

Arquivo:

```text
config/import_cache_mapping.json
```

As opções principais são:

```json
{
  "executar_consultas_importcache": true,
  "modulo_lista_tabelas": "ImportCache",
  "fonte_consultas_importcache": "template",
  "template_mdb_consultas": "templates/DataSync - Recovery.mdb"
}
```

Com isso, as consultas não dependem mais de o MDB selecionado ter a tabela `ListaTabelas` preenchida. Elas são lidas do MDB modelo que vai dentro do projeto.

## Gerar EXE

Execute:

```bat
build_exe.bat
```

Depois use a pasta `dist` completa:

```text
dist/
  DataSyncImportCache.exe
  config/
    import_cache_mapping.json
  templates/
    DataSync - Recovery.mdb
```

## Execucao das consultas ImportCache dentro do Microsoft Access

Nesta versao, as consultas da tabela `ListaTabelas` com `Modulo = ImportCache` nao sao executadas via ODBC/pyodbc. Elas sao executadas pelo proprio Microsoft Access usando automacao COM.

Fluxo usado:

1. Importa os dados do SQLite para as tabelas temporarias `dbo_A_CAIXA_*` no MDB selecionado.
2. Le as consultas ImportCache do MDB modelo em `templates/DataSync - Recovery.mdb`.
3. Fecha a conexao ODBC com o MDB do sistema.
4. Abre o Microsoft Access.
5. Cria temporariamente uma consulta salva no MDB selecionado.
6. Executa essa consulta dentro do Access.
7. Fecha o Access.
8. Repete o processo ate terminar todas as consultas.

Esse modo permite executar consultas que usam funcoes VBA do Access, por exemplo `Util_nz`.

Requisitos adicionais:

- Microsoft Access instalado no Windows.
- `pywin32` instalado no ambiente virtual.

Instalacao:

```powershell
pip install -r requirements.txt
```

No arquivo `config/import_cache_mapping.json`, estes campos controlam o comportamento:

```json
"executar_consultas_importcache": true,
"executar_consultas_via_access_com": true,
"access_visivel_durante_execucao": false
```

Se quiser ver o Access abrindo durante a execucao, altere:

```json
"access_visivel_durante_execucao": true
```

## v1.1.0 - melhoria de desempenho no Access

Esta versão adiciona o modo de execução `unica_sessao`, onde o Microsoft Access é aberto uma única vez, executa todas as consultas `ImportCache` e fecha somente no final.

No arquivo `config/import_cache_mapping.json`:

```json
"modo_execucao_access": "unica_sessao",
"continuar_apos_erro_consulta": false
```

Para voltar ao modo mais seguro antigo, altere para:

```json
"modo_execucao_access": "reabrindo_access"
```


## Versão 1.3.0

Alterações desta versão:

- Interface moderna em PySide6 mantida como interface principal.
- Removido backup automático do MDB para reduzir tempo de execução.
- Barra de progresso real baseada nas consultas executadas pelo Microsoft Access.
- Eventos internos `PROGRESS_CONSULTA_INICIO` e `PROGRESS_CONSULTA_OK` para atualizar a tela em tempo real.
- Versão atualizada para 1.2.0.
