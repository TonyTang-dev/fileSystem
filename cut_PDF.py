# 导入读写pdf模块
from PyPDF2 import PdfFileReader, PdfFileWriter
'''
注意：
页数从0开始索引
range()是左闭右开区间
'''
 
def split_pdf(file_name, start_page, end_page, output_pdf):
    '''
    :param file_name:待分割的pdf文件名
    :param start_page: 执行分割的开始页数
    :param end_page: 执行分割的结束位页数
    :param output_pdf: 保存切割后的文件名
    '''
    # 读取待分割的pdf文件
    input_file = PdfFileReader(open(file_name, 'rb'))
    # 实例一个 PDF文件编写器
    output_file = PdfFileWriter()
    # 把分割的文件添加在一起
    for i in range(start_page, end_page):
        output_file.addPage(input_file.getPage(i))
    # 将分割的文件输出保存
    with open(output_pdf, 'wb') as f:
        output_file.write(f)
 
def merge_pdf(merge_list, output_pdf):
    """
    merge_list: 需要合并的pdf列表
    output_pdf：合并之后的pdf名
    """
    # 实例一个 PDF文件编写器
    output = PdfFileWriter()
    for ml in merge_list:
        pdf_input = PdfFileReader(open(ml, 'rb'))
        page_count = pdf_input.getNumPages()
        for i in range(page_count):
            output.addPage(pdf_input.getPage(i))
 
    output.write(open(output_pdf, 'wb'))
 
 
if __name__ == '__main__':
    # 分割pdf
    # split_pdf("c:\\Users\\Administrator\\Desktop\\test.pdf", 7, 26, "c:\\Users\\Administrator\\Desktop\\第一章.pdf")
    # split_pdf("c:\\Users\\Administrator\\Desktop\\test.pdf", 29, 51, "c:\\Users\\Administrator\\Desktop\\第二章.pdf")
    # split_pdf("c:\\Users\\Administrator\\Desktop\\test.pdf", 61, 85, "c:\\Users\\Administrator\\Desktop\\第三章.pdf")
    # split_pdf("c:\\Users\\Administrator\\Desktop\\test.pdf", 91, 110, "c:\\Users\\Administrator\\Desktop\\第四章.pdf")
    # split_pdf("c:\\Users\\Administrator\\Desktop\\test.pdf", 116, 139, "c:\\Users\\Administrator\\Desktop\\第五章.pdf")
    # split_pdf("c:\\Users\\Administrator\\Desktop\\test.pdf", 144, 165, "c:\\Users\\Administrator\\Desktop\\第六章.pdf")
    
    split_pdf("信号处理答案-解析部分.pdf", 9, 128, "信号处理答案-解析部分-新.pdf")

    
    # 合并pdf
    # 合并的pdf列表
    pdf_list = ["第一章.pdf", "第二章.pdf", "第三章.pdf", "第四章.pdf", "第五章.pdf", "第六章.pdf"]
    merge_pdf(pdf_list, "信号处理答案-解析部分.pdf")
    
    print("dealing finishing")