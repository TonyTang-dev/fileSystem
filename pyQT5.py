 
import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
 
 
class example(QWidget):
    def __init__(self):
        super(example, self).__init__()
        # 窗口标题
        self.setWindowTitle('拖拽获取文件路径')
        # 定义窗口大小
        self.resize(500, 400)
        self.QLabl = QLabel(self)
        self.QLabl.setGeometry(0, 100, 4000, 100)
        # 调用Drops方法
        self.setAcceptDrops(True)
 
    # 鼠标拖入事件
    def dragEnterEvent(self, evn):
        self.setWindowTitle('鼠标拖入窗口了')
        self.QLabl.setText('文件路径：\n' + evn.mimeData().text())
        # 鼠标放开函数事件
        evn.accept()
 
    # 鼠标放开执行
    def dropEvent(self, evn):
        self.setWindowTitle('鼠标放开了')
 
    def dragMoveEvent(self, evn):
        print('鼠标移入')
 
 
if __name__ == "__main__":
    app = QApplication(sys.argv)
    e = example()
    e.show()
    sys.exit(app.exec_())