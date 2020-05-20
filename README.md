# HealthCard_Creator
自动复制并修改健康卡中的日期，文件修改时间也会同步发生修改，使用python和docx库。

因为修改系统时间需要管理员权限，所以请以管理员身份运行程序。

## 一、前置安装

运行python的安装包（已在文件夹中）安装python，注意pip的钩勾上（默认勾上）；

打开cmd，输入pip install python-docx -i https://pypi.tuna.tsinghua.edu.cn/simple安装docx库；

进入C:\Windows找到py.exe，右键属性->兼容性勾选以管理员身份运行；

## 二、使用步骤

运行前请先确认：

- 源文件为docx格式

- 源文件的最后一行日期处应该为空（2020年 月 日）

操作步骤：

- 双击运行CopyHealthCard.py，根据中文提示输入文件的所在位置，回车继续；
- 输入文件保存路径，回车继续；输入保存文件名的前缀，回车继续；
- 程序会打印word文档中的内容，之后提示输入日期所在的段落，根据打印的word内容确定所在段落，回车继续；
- 稍等片刻即可得到复制得到的word文档

