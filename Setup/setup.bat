cd "%cd%\\..\\"
python -m venv "automation"
cd automation\\Scripts
pip.exe install pyinstaller
pip.exe install -r "..\\..\\requirements.txt"
cmd /k