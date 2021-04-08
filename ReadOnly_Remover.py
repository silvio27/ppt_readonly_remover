import os
import shutil
import zipfile
import re
import sys

# 制作副本

def change_suffix(filename):
    shutil.copy(filename, '1.zip')

def deal_zip(filename):
    change_suffix(filename)
    aa = zipfile.ZipFile(filename)
    aa.extractall('./1')
    deal_xml('./1./ppt/presentation.xml') # 处理xml文件
    shutil.make_archive('1','zip','1') # 重新打包
    shutil.rmtree('./1')    # 删除临时文件
    os.rename('1.zip', os.path.splitext(filename)[0] + ' 已破解'+ '.pptx') # 文件重命名

# 处理xml
def deal_xml(xmlfile):
    with open(xmlfile, 'r') as text:
        startext = '<p:modifyVerifier'
        endtext = '=="/>'
        pattern = startext + '.+?' + endtext
        bb = re.sub(pattern, '', text.read())
    with open(xmlfile, 'w') as text:
        text.write(bb)


if __name__ == '__main__':
    files = sys.argv
    for i in sys.argv[1:]:
        if 'ppt' in os.path.splitext(i)[1]:
            deal_zip(i)