import os
from datetime import date
from datetime import datetime
from datetime import timedelta
from datetime import timedelta
from docx import Document
from docx.shared import Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import re
import os


#str="D:\VSCodeProject\PythonProject\copyhealthycard\190110419李怡凯.docx"
addressIn=input("请输入文件绝对地址，注意后缀应该为docx")
addressOut=input("请输入保存地址")
document=Document(r"%s"%addressIn)
begin=datetime.strptime('2020-2-24','%Y-%m-%d')
end=datetime.now()
b=begin.date()
e=end.date()
for i in range(len(document.paragraphs)):
        if len(document.paragraphs[i].text.replace(' ',''))>4:
            print("第"+str(i)+"段的内容是："+document.paragraphs[i].text)
par=input("请输入日期所在的段落号，注意请先确认日期处为空“ 月 日"，否则日期无法自动修改")
par=int(par)
for i in range((e - b).days+1): 
    myday = b + timedelta(days=i) 
    os.system("date 2020/%d/%d" %(myday.month, myday.day))
    document=Document(r"%s"%addressIn)
    paragraphs = document.paragraphs
    text = re.sub(" 月 号", "%d月%d号"%(myday.month,myday.day), paragraphs[par].text)
    paragraphs[par].text = text
    if myday.day <10:
        saveadd=addressOut+"\healthcard0%d0%d.docx"%(myday.month, myday.day)
        document.save("%s"%saveadd)
    else:
        saveadd = addressOut + "\healthcard0%d%d.docx"%(myday.month, myday.day)
        document.save("%s"%saveadd)

    