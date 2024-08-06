@echo off
REM Install dependencies from cwd/dependencies.txt
pip install -r "%cd%\dependencies.txt" || pause

REM Create a shortcut to run.bat in the same directory as run.bat
powershell -Command "$WshShell = New-Object -ComObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('%cd%\Report Generator.lnk'); $Shortcut.TargetPath = '%cd%\run.bat'; $Shortcut.IconLocation = '%cd%\icon.ico'; $Shortcut.Save()"
