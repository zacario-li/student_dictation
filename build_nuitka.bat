@echo off
REM ===== Nuitka 打包脚本 =====
REM 需先激活好 Python 环境并安装好 Nuitka、PySide6、edge-tts、pyttsx3 等依赖
REM 自动生成的ico文件名为 app_icon_auto.ico

python -m nuitka ^
  --onefile ^
  --windows-console-mode=disable ^
  --windows-icon-from-ico=app_icon_auto.ico ^
  --enable-plugin=pyside6 ^
  --include-qt-plugins=all ^
  word_announcer.py

pause 