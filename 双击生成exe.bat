rmdir /s /q dist
rmdir /s /q build
pyinstaller -F -w .\zlgcan_demo.py
copy .\*dll dist
copy .\dev_info.json dist
xcopy /E /I kerneldlls dist\kerneldlls 