import sys
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
def window():
    app = QApplication(sys.argv)
    win = QDialog()
    win.setGeometry(100,100,400,500)
    win.show()
    sys.exit(app.exec_())
    
if __name__ == '__main__':
   window()    