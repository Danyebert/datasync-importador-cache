@echo off
setlocal
python -m venv .venv
call .venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
pyinstaller --noconsole --onefile --name DataSyncImportCache app.py
if not exist dist\config mkdir dist\config
if not exist dist\templates mkdir dist\templates
copy config\import_cache_mapping.json dist\config\import_cache_mapping.json /Y
copy "templates\DataSync - Recovery.mdb" "dist\templates\DataSync - Recovery.mdb" /Y
echo.
echo EXE gerado em dist\DataSyncImportCache.exe
echo Copie junto as pastas dist\config e dist\templates.
pause
