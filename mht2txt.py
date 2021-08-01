#-*- coding: UTF-8 -*-

import sys
import os
import os
import zipfile
from win32com import client as wc
import xlrd
from bs4 import BeautifulSoup
from pydocx import PyDocX
from lxml import html

#reload(sys)

#sys.setdefaultencoding('gbk')
#sys.setdefaultencoding('utf-8')

def deal_1file(filenamein, f_out):
    f = open(filenamein, "r")
    lines = f.readlines()
    start=0
    for line in lines:
        if(start==0):
            #if(u'</td>&nbsp; &nbsp;&nbsp; &nbsp;&nbsp;&nbsp;</tr></table>' in line or u'count_add_one' in line):
            if('</tr></table><br />' in line):
                start = 1
                f_out.write(line)

        else:
            #if(u'谢谢点红心支持' in line or u'版主提醒：' in line):
            if ('thanks.gif' in line or 'fieldset' in line):
                f_out.write('end of chapter' + '\n')
                start=2
                break
                #print(line)
            else:
                f_out.write(line)
    if(start != 2):
        print("not find end in %s" %(filenamein))

def filter_content(content_str):
    #idx = content_str.rfind('* 修改这里')
    idx = -1
    content_str_rt = content_str[:idx]
    return content_str_rt


def deal_1file_mht0(filenamein, f_out):
    '''
    把mht改后缀为doc，再转成docx，再改后缀成zip，读zip里面的xml后变成txt
    这个方法繁琐，而且出来的txt没有回车换行
    :param filenamein:
    :param f_out:
    :return:
    '''
    doc_filename = filenamein.replace("mht", "doc")
    docx_filename = filenamein.replace("mht", "docx")
    zip_filename = filenamein.replace("mht", "zip")

    #把mht改后缀为doc，再转成docx
    os.rename(filenamein, doc_filename)
    word = wc.Dispatch('Word.Application')
    doc = word.Documents.Open(doc_filename)
    doc.SaveAs(docx_filename, 12, False, "", True, "", False, False, False, False)  # 转化后路径下的文件
    doc.Close()
    word.Quit()

    #把docx改后缀成zip，读zip里面的xml后变成txt
    os.rename(docx_filename, zip_filename)
    f = zipfile.ZipFile(zip_filename, 'r')
    xml = f.read("word/document.xml")
    wordObj = BeautifulSoup(xml.decode("utf-8"))
    #wordObj = BeautifulSoup(wordObj.text.replace('&nbsp;', ' '), "lxml")
    # print(wordObj)
    texts = wordObj.findAll("w:t")
    content = []
    for text in texts:
        content.append(text.text)
    content_str = "".join(content)

    把单个txt(每个mht对应一个)写入到汇总的txt（所有的mht写入到同一个txt）
    content_str_filted = filter_content(content_str)
    # content_str_filted = content_str_filted.replace(u'\xa0', u' ')
    f_out.write(content_str_filted)



def deal_1file_mht1(filenamein, f_out):
    '''
        把mht改后缀为doc，再转成docx，读docx后变成txt
        这个方法繁琐，而且不知道为什么在shell里面好好地，变成.py脚本运行的时候读docx的而且出来的wordapp.Documents.Open总失败
        :param filenamein:
        :param f_out:
        :return:
    '''
    doc_filename = filenamein.replace("mht", "doc")
    docx_filename = filenamein.replace("mht", "docx")
    zip_filename = filenamein.replace("mht", "zip")
    txt_filename = filenamein.replace("mht", "txt")
    os.rename(filenamein, doc_filename)
    word = wc.Dispatch('Word.Application')
    doc = word.Documents.Open(doc_filename)
    doc.SaveAs(docx_filename, 12, False, "", True, "", False, False, False, False)  # 转化后路径下的文件
    doc.Close()
    word.Quit()


    docx_filename2 = '%s.docx' %(docx_filename[:29])
    print(docx_filename2)
    os.rename(docx_filename, docx_filename2)
    try:
        wordapp = wc.Dispatch('Word.Application')
        doc = wordapp.Documents.Open(docx_filename2)
        # 为了让python可以在后续操作中r方式读取txt和不产生乱码，参数为4
        doc.SaveAs(txt_filename, 4)
        doc.Close()

        with open(txt_filename, "r", encoding='utf-8') as ftxt:
            content_str = ftxt.read()
        content_str_filted = filter_content(content_str)
        # content_str_filted = content_str_filted.replace(u'\xa0', u' ')
        f_out.write(content_str_filted)
    except:
        print('fail')


def deal_1file_mht(filenamein, f_out):
    '''
    直接读mht，存成txt
    :param filenamein:
    :param f_out:
    :return:
    '''
    doc_filename = filenamein.replace("mht", "doc")
    docx_filename = filenamein.replace("mht", "docx")
    zip_filename = filenamein.replace("mht", "zip")
    txt_filename = filenamein.replace("mht", "txt")

    try:
        wordapp = wc.Dispatch('Word.Application')
        doc = wordapp.Documents.Open(filenamein)
        # 为了让python可以在后续操作中r方式读取txt和不产生乱码，参数为4
        doc.SaveAs(txt_filename, 4)
        doc.Close()

        with open(txt_filename, "r") as ftxt:
            # with open(txt_filename, "r", encoding='utf-8') as ftxt:
            content_str = ftxt.read()
        content_str_filted = filter_content(content_str)
        # content_str_filted = content_str_filted.replace(u'\xa0', u' ')
        f_out.write(content_str_filted)
    except:
        print('fail '*10)

def get_filelist(filepath):
    filelist=list()
    g = os.walk(filepath)
    for path, dir_list, file_list in g:
        for file_name in file_list:
            #print(os.path.join(path, file_name))
            filelist.append(os.path.join(path, file_name))
    filelist.sort()
    return filelist

#把多篇文章的mht，汇总成一个txt
if __name__ == '__main__':

    f_out = open('out.txt', 'w',encoding='utf-8')
    filelist = get_filelist('K:\\mycode\\search_ss\\tmp3')
    for file in filelist:
        print('dealing %s' %(file))
        deal_1file_mht(file, f_out)
    f_out.close()
    print(len(filelist))

'''
import os
import zipfile
from win32com import client as wc
import xlrd
from bs4 import BeautifulSoup
from pydocx import PyDocX
from lxml import html
doc_filename='K:\\mycode\\search_ss\\tmp2\\777.doc'
docx_filename = doc_filename.replace("doc", "docx")
word = wc.Dispatch('Word.Application')
doc = word.Documents.Open(doc_path)
doc.SaveAs(docx_filename, 12, False, "", True, "", False, False, False, False)  # 转化后路径下的文件
doc.Close()
word.Quit()

zipfilename = 'K:\\mycode\\search_ss\\tmp2\\777.zip'
f = zipfile.ZipFile(zipfilename, 'r')

xml = f.read("word/document.xml")

wordObj = BeautifulSoup(xml.decode("utf-8"))
# print(wordObj)
texts = wordObj.findAll("w:t")
content = []
for text in texts:
    content.append(text.text)
content_str = "".join(content)
'''