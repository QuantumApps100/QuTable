pyinstaller --onefile --noconsole --name QuTable --icon="!/icons/QTableIcon.ico" QuTable.py
REM pyinstaller --onefile --name QuTable --icon="!/icons/QTableIcon.ico" QuTable.py
cd dist
del "../!/QuTable.exe"
ren "QuTable.exe" "../!/QuTable.exe"
del "../QuTable.spec"
Set-ItemProperty -Path "../!/QuTable.exe" -Name Publisher -Value "Quantum Apps"
cd ..