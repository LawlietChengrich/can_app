rmdir /s /q dist
rmdir /s /q build
pyinstaller -F -w .\dh_can_test.py
md dist\dll\
copy .\dll\*dll dist\dll\
copy .\dev_info.json dist
xcopy /E /I kerneldlls dist\kerneldlls 