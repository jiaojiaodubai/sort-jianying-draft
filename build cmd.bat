cd /d %~dp0
cd .\venv\Scripts
activate
pyinstaller -F -w --uac-admin -i .\icon1.ico .\src\JianyingAssistant.py