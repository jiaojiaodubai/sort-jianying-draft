cd /d %~dp0
python -m pip install --upgrade pip
pip install -r .\requirements.txt
pyinstaller -F -w --uac-admin -i .\icon1.ico .\src\JianyingAssistant.py --upx-dir .\upx-4.0.2-win64 -w --version-file .\VersionInfo.txt