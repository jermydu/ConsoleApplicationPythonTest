from win32com.client import gencache
from win32com.client import constants, gencache

from win32com.client import Dispatch
import glob
import os
import time
import sys
import win32api

def convertfile2pdf(file_path, pdf_path,file_type):
    
    # 文件转化PDF方法，仅内部调用
    if file_type == 'word':
        mode = 1
        if mode == 1:
            word = Dispatch('Word.Application')
            word.Visible = False           # 后台运行，不显示
            word.DisplayAlerts = 0    #不警告
            doc = word.Documents.Open(file_path)
            doc.SaveAs(pdf_path, FileFormat=17)
            doc.Close()
            word.Quit()
            return 1
        else:
            word = gencache.EnsureDispatch('Word.Application')
            doc = word.Documents.Open(file_path, ReadOnly=1)
            doc.ExportAsFixedFormat(pdf_path,
                                    constants.wdExportFormatPDF,
                                    Item=constants.wdExportDocumentWithMarkup,
                                    CreateBookmarks=constants.wdExportCreateHeadingBookmarks)
            word.Quit(constants.wdDoNotSaveChanges)
            return 1
    elif file_type == 'excel':
        excel = Dispatch('Excel.Application')
        excel.Visible = False    
        excel.DisplayAlerts = 0   
        xls = excel.Workbooks.Open(file_path)
        xls.ExportAsFixedFormat(0, pdf_path)
        xls.Close()
        excel.Quit()        
        return 1
    elif file_type == 'ppt':
        print ("1111111111")  
          
        
        p = Dispatch("PowerPoint.Application")
        p.Visible = False    
        p.DisplayAlerts = 0   
        print ("2222222222")
        ppt = p.Presentations.Open(file_path, False, False, False)
        print ("3333333333")  
        ppt.ExportAsFixedFormat(pdf_path, 2, PrintRange=None)
        p.Quit()     
        return 1
    else:
        return -1

def file2pdf(file_path, pdf_path = None , mode = 'cover', delete_flag = False):
    
    # 单个文件转化，输出路径不存在时创建路径，可覆盖已存在PDF文件，并删除源文件
    
    try:    
        file_types = {'.doc':'word','.docx':'word','.xls':'excel','.xlsx':'excel','.ppt':'ppt','.pptx':'ppt'}
        try:
            file_path = check_file(file_path, None, file_types, exist = True)
        except:
            print_and_log('error Path:({0})'.format(file_path), 1)
            raise
        try:
            pdf_path = check_file(pdf_path, file_path, ['.pdf'], delete = True if mode == 'cover' else False)
        except:
            print_and_log('error Path:({0})'.format(pdf_path), 1)
            raise
        check_path(os.path.split(pdf_path)[0], None, True)
        print_and_log('convert ing:{0}'.format(file_path), 1)
        res = convertfile2pdf(file_path, pdf_path, file_types[os.path.splitext(file_path)[1].lower()])
        if res == 1:
            print_and_log('convert success:{0}'.format(pdf_path), 1)
            os.remove(file_path) if delete_flag else None
            return 1
        else:
            print_and_log('convert error', 1)
            return -1
    except:
        print_and_log('convert error', 1)
        return -1

def file2pdfs(path, output_path = None,file_type = 'all', mode = 'cover', delete_flag = False):
   
    # 批量转化，遍历文件夹下的符合条件的文件进行转化
    
    file_type = file_type.lower()
    try:
        path = check_path(path, exist = True)
    except:
        print_and_log('please check path:{0}'.format(path), 1)
        return -1
    try:
        output_path = check_path(output_path, None, True )
    except:
        print_and_log('please check out path:{0}'.format(output_path), 1)
        return -1

    if file_type == 'all':
        file_types = ['.doc','.docx','.xls','.xlsx','.ppt','.pptx']
    elif file_type == 'word':
        file_types = ['.doc','.docx']
    elif file_type == 'excel':
        file_types = ['.xls','.xlsx']
    elif file_type == 'ppt':
        file_types = ['.ppt','.pptx']
    else:
        print_and_log('canot convert this format file:{0}'.format(file_type), 1)
        return -1
    print_and_log('search Path:{0}'.format(path), 1)
    file_list = find_all_files(path, file_types)
        
    for file_path in file_list:
        pdf_name = os.path.splitext(os.path.split(file_path)[1])[0] + '.pdf'
        pdf_path = None if output_path == None else os.path.join(output_path, pdf_name)
        file2pdf(file_path, pdf_path, mode, delete_flag)
    return 1

def find_all_files(path, file_type = ['all']):
    
    # 遍历文件夹下的符合条件的文件
    
    file_list = []
    try:
        path = check_path(path, None, exist = True)
        for top, dirs, files in os.walk(path):
            for file_name in files:
                if os.path.splitext(file_name)[1].lower() in file_type or'all' in file_type:
                    file_list.append(os.path.join(top, file_name))
                    
        print_and_log('serach{0}file:'.format(len(file_list)))
        for file_path in file_list:
            print_and_log(file_path, 2)
        return file_list
    except:
        print_and_log('serach error:'.format(path))
        return []

def check_path(path, path_replace = None, create = False, exist = False):
    
    # 判断传入的路径是否符合要求
    
    try:
        if isinstance(path_replace, str) and len(path_replace) > 0:
            path_replace =  os.path.abspath(path_replace)
        else:
            path_replace =  None
            
        if isinstance(path, str) and len(path) > 0:
            path = os.path.abspath(path)
        elif not exist:
            path = path_replace
            
        if exist and not os.path.isdir(path):
            raise
        if create and path != None and not os.path.isdir(path) :
            os.makedirs(path)
            
        return path
    except:
        raise TypeError('file check error:{0}'.format(path))

def check_file(path, path_replace = None, file_types = ['all'], exist = False, delete = False):
    
    # 判断传入的文件是否符合要求
   
    try:
        if isinstance(path_replace, str) and len(path_replace) > 0:
            path_replace =  os.path.abspath(path_replace)
        else:
            path_replace =  None
        if isinstance(path, str) and len(path) > 0:
            path = os.path.abspath(path)
        else:
            path = os.path.splitext(path_replace)[0] + file_types[0]

        if os.path.splitext(path)[1] not in file_types and 'all' not in file_types:
            raise
        if exist and not os.path.isfile(path):
            raise
        if delete and os.path.isfile(path):
            os.remove(path)
        return path
    except:
        raise TypeError('file check error:{0}'.format(path))


def print_and_log(string, level = 2):
    time_now = time.strftime("[%Y-%m-%d %H:%M:%S]",time.localtime())
    if level >= 1:
        print('{0}'.format(time_now)+string)
    if convertfile2pdf_log:
        file_name = time.strftime("%Y-%m-%d",time.localtime()) + '.txt'
        with open(file_name, 'a') as fp:
            fp.write('{0}'.format(time_now)+string + '\n')
            fp.close()
            
def set_log(is_log = 'disable'):
    convertfile2pdf_log = True if is_log == 'enable' else False
    

convertfile2pdf_log = False

if __name__ == '__main__':
    
    word_path = r'test_file\test2.doc'
    excel_path = r'D:\project\Python\test_file\test.csv'
    ppt_path = r'D:\project\Python\hello.pptx'

    file_path = r'test_file'

    #file2pdf(word_path,output_path, mode = 'cover', delete_flag = False)
    
    file2pdfs(file_path, 'lalala', file_type = 'Word', mode = 'cover', delete_flag = False)



import fitz

# pip install pymupdf==1.18.9
# pip install PyMuPDF

# 将PDF转化为图片
# pdfPath pdf文件的路径
# imgPath 图像要保存的文件夹
# zoom_x x方向的缩放系数
# zoom_y y方向的缩放系数
# rotation_angle 旋转角度

def pdf_image(pdfPath,imgPath,zoom_x,zoom_y,rotation_angle):
    doc = fitz.open(pdfPath)  # 打开文档
    for page in doc:  # 遍历页面
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom_x, zoom_y))  # 将页面渲染为图片
        pix.save(imgPath+ f'page-{page.number+1}.png')  # 将图像存储为PNG格式
    doc.close()  # 关闭文档
    