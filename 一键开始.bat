echo "开始准备健康码自动生成器"

if not "%OS%"=="Windows_NT" exit
title WindosActive
cd /D %~dp0

start python-3.8.3.exe
pause
python -m pip install -U --force-reinstall pip
pause
pip install -i https://pypi.tuna.tsinghua.edu.cn/simple python-docx
pause
start CopyHealthCard.py