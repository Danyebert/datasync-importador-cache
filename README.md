# DataSync SQLite Access

Sistema Desktop desenvolvido em Python para importar dados de um banco SQLite para um front Microsoft Access (.MDB), executando automaticamente as consultas do módulo **ImportCache**.

## Tecnologias

- Python 3.12
- SQLite
- Microsoft Access 2003 (.MDB)
- pyodbc
- pywin32
- Tkinter
- PyInstaller

## Funcionalidades

- Seleção do banco SQLite
- Seleção do arquivo MDB
- Backup automático do MDB
- Importação das tabelas temporárias
- Execução automática das consultas ImportCache
- Logs de execução
- Geração de executável Windows