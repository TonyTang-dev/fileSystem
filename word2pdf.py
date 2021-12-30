# -*- encoding: utf-8 -*-
import  os
from win32com import client
#pip install win32com
def doc2pdf(doc_name):
    """
    :word文件转pdf
    :param doc_name word文件名称
    :param pdf_name 转换后pdf文件名称
    """
    if doc_name[-3:]=="doc" or doc_name[-4:]=="docx":
        try:
            pdfName=""
            if doc_name[-3:]=="doc":
                pdfName=doc_name[:-4]
            else:
                pdfName=doc_name[:-5]
            word = client.DispatchEx("Word.Application")
            if os.path.exists(pdfName+".pdf"):
                #print("当前路径下有该PDF文件了哦")
                return 1
                os.remove(pdf_name+"pdf")
            worddoc = word.Documents.Open(doc_name,ReadOnly = 1)
            worddoc.SaveAs(pdfName+".pdf", FileFormat = 17)
            worddoc.Close()
            return 0    #pdf_name
        except:
            #print("转换发生错误")
            return 1
    
    elif doc_name[-3:]=="ppt" or doc_name[-4:]=="pptx":
        try:
            pdfName=""
            if doc_name[-3:]=="ppt":
                pdfName=doc_name[:-4]
            else:
                pdfName=doc_name[:-5]
            # 2). 打开PPT程序
            ppt_app = client.Dispatch('PowerPoint.Application')
            if os.path.exists(pdfName+".pdf"):
                #print("当前路径下有该PDF文件了哦")
                return 1
                os.remove(pdfName+".pdf")
            # ppt_app.Visible = True  # 程序操作应用程序的过程是否可视化

            # 3). 通过PPT的应用程序打开指定的PPT文件
            # filename = "C:/Users/Administrator/Desktop/PPT办公自动化/ppt/PPT素材1.pptx"
            # output_filename = "C:/Users/Administrator/Desktop/PPT办公自动化/ppt/PPT素材1.pdf"
            ppt = ppt_app.Presentations.Open(doc_name,ReadOnly = 1)

            # 4). 打开的PPT另存为pdf文件。17数字是ppt转图片，32数字是ppt转pdf。
            ppt.SaveAs(pdfName+".pdf", 32)
            # print("导出成pdf格式成功!!!")
            # 退出PPT程序
            ppt_app.Quit()
            return 0
        except:
            # print(e)
            return 1
            
 
def main():
    # input = r'c:\\Users\\Administrator\\Desktop\\cs21106_proj_v1.0.pptx'
    # print(input)
    # output = r'c:\\Users\\Administrator\\Desktop\\2.pdf'
    # print(output)
    rc = doc2pdf(input)
    print(rc)
    # rc = doc2html(input, output)
    # rc = pdf2doc(input, output)
    if rc:
        print('转换成功')
    else:
        print('转换失败')
 
if __name__ == '__main__':
    main()