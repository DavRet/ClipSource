import sys
import urllib2

import exifread
import win32con
from bs4 import BeautifulSoup

import win32clipboard
import HTMLClipboard





from PySide import QtGui, QtCore

app = QtGui.QApplication(sys.argv)
clipboard = app.clipboard()

clips = []


try:
    import tkinter as tk  # for python 3
except:
    import Tkinter as tk  # for python 2
import pygubu
from PyQt4 import QtGui, uic




class Application:
    def __init__(self, master):

        #1: Create a builder
        self.builder = builder = pygubu.Builder()

        #2: Load an ui file
        builder.add_from_file('ClipSourceUI.ui')

        #3: Create the widget using a master as parent
        self.mainwindow = builder.get_object('LabelFrame_3', master)



class Window(QtGui.QMainWindow):
    def __init__(self):
        super(Window, self).__init__()
        uic.loadUi('clipui.ui', self)


        self.show()

    def test(self):
        print "test"
        # def initUI(self):
    #     QtGui.QToolTip.setFont(QtGui.QFont('SansSerif', 10))
    #
    #     self.setToolTip('This is a <b>QWidget</b> widget')
    #
    #     btn = QtGui.QPushButton('Button', self)
    #     btn.setToolTip('This is a <b>QPushButton</b> widget')
    #     btn.resize(btn.sizeHint())
    #     btn.move(50, 50)
    #
    #     self.setGeometry(1100, 200, 500, 800)
    #     self.setWindowTitle('Tooltips')
    #     self.show()
    #
    #
    # def _populate(self):
    #
    #     self.clear()
    #
    #     for clip in clips:
    #         item = QtGui.QListWidgetItem(self)
    #         item.setText(clip)

window = Window()

content_list = window.content_list
content_model = QtGui.QStandardItemModel(content_list)

source_list = window.source_list
source_model = QtGui.QStandardItemModel(source_list)


def main():

    clipboard.dataChanged.connect(clipboardChanged)

    #window = Window()


    content_list.setModel(content_model)
    source_list.setModel(source_model)


    # We set the form to be our ExampleApp (design)

    #form.content_list.setModel(model)
    window.show()

    #list.setWindowTitle('ClipSource')
    #list.setMinimumSize(600, 400)


    #list.show()

    setModelForList()

    #root = tk.Tk()
    #app = Application(root)
    #root.mainloop()



    sys.exit(app.exec_())


def setModelForList():


    for clip in clips:
        clipText = QtGui.QStandardItem(clip[0])
        source = QtGui.QStandardItem(clip[1])

        content_model.appendRow(clipText)
        source_model.appendRow(source)

    def on_item_changed(item):
        if not item.checkState():
            return item
        print item.text()

        clipboard.setText(item.text())


        i = 0
        while content_model.item(i):
            if not content_model.item(i).checkState():
                return
            i += 1

    content_model.itemChanged.connect(on_item_changed)

    content_list.setModel(content_model)
    source_list.setModel(source_model)


def buttonClick():
    print "click"


def clipboardChanged():

    mimeData = clipboard.mimeData()

    image = clipboard.image()

    print image.height()

    image.save("C:\\Users\\David\\Documents\\Images\\aaa.png", "PNG", -1)

    f = open("tiger.jpg", 'rb')

    tags = exifread.process_file(f)

    print tags

    #print image.height()


    if HTMLClipboard.HasHtml():
        # print('there is HTML!!')
        source = HTMLClipboard.GetSource()
        print(source)

        #url = QtCore.QUrl(source)
        #mimeData.setUrls(url)
        #print clipboard.mimeData().urls()

        originalText = clipboard.text()

        clips.append(tuple((originalText, source)))

        setModelForList()

        #clipboard.setText(originalText + "\n" + "Copied from:\n"  + source)

        print clipboard.text()
    else:
        print('no html')

    soup = BeautifulSoup(mimeData.html(), "html.parser")

    links = soup.find_all('a')






    #print links
    for tag in links:
        link = tag.get('href', None)
        if link is not None:
            print link



if __name__ == '__main__':
    main()