rmdir /s /q dist
rmdir /s /q build
pyinstaller -F -w .\dh_can_test.py
copy .\*dll dist
copy .\dev_info.json dist
xcopy /E /I kerneldlls dist\kerneldlls 