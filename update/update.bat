pip install pyinstaller
pyinstaller --onefile "%CD%\..\main.py" --upx-dir "%CD%"
@RD /S /Q "%CD%\build"
@RD /S /Q "%CD%\..\__pycache__"
del main.spec
ren "dist\main.exe" "send_emails_new.exe"
move "dist\send_emails_new.exe" "%CD%"
@RD /S /Q "%CD%\dist"
move "%CD%\send_emails_new.exe" "%CD%\.."
echo|set/p="Press <ENTER> to exit.."&runas/u: "">NUL
PAUSE
