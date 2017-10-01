import sys
import urllib2

import exifread
import rdflib
import win32con
from bs4 import BeautifulSoup

import HTMLClipboard
import os
import os.path
import sys

from libcredit import Credit, HTMLCreditFormatter, TextCreditFormatter

from PySide import QtGui, QtCore

app = QtGui.QApplication(sys.argv)
clipboard = app.clipboard()

clips = []

currentIndex = []


import pygubu
from PyQt4 import QtGui, uic

import win32clipboard as clp, win32api

# clp.OpenClipboard(0)
# clp.EmptyClipboard()
# clp.SetClipboardData(0, None)
# clp.CloseClipboard()

# clp.OpenClipboard(None)
# rc = clp.EnumClipboardFormats(0)
# while rc:
#     try:
#         format_name = clp.GetClipboardFormatName(rc)
#     except win32api.error:
#         format_name = "?"
#
#     print format_name
#     try:
#         format = clp.GetClipboardData(rc)
#     except win32api.error:
#         format = "?"
#     print format
#     rc = clp.EnumClipboardFormats(rc)
#
# clp.CloseClipboard()


class Application:
    def __init__(self, master):
        # 1: Create a builder
        self.builder = builder = pygubu.Builder()

        # 2: Load an ui file
        builder.add_from_file('ClipSourceUI.ui')

        # 3: Create the widget using a master as parent
        self.mainwindow = builder.get_object('LabelFrame_3', master)


class Window(QtGui.QMainWindow):
    def __init__(self):
        super(Window, self).__init__()
        uic.loadUi('clipui.ui', self)
        self.show()


window = Window()

content_list = window.content_list
content_model = QtGui.QStandardItemModel(content_list)

source_list = window.source_list
source_model = QtGui.QStandardItemModel(source_list)

source_radio = window.source_radio

clip_edit = window.clip_edit

auto_source_enabled = False


def main():
    clipboard.dataChanged.connect(clipboardChanged)

    content_list.setModel(content_model)
    source_list.setModel(source_model)

    source_radio.toggled.connect(sourceRadioClicked)

    window.show()

    setModelForList()

    content_list.clicked.connect(contentClick)
    source_list.clicked.connect(sourceClick)

    clip_edit.textChanged.connect(editTextChanged)

    window.hash_comment_radio.clicked.connect(hashCommentClicked)
    window.slash_comment_radio.clicked.connect(slashCommentClicked)





    # root = tk.Tk()
    # app = Application(root)
    # root.mainloop()



    sys.exit(app.exec_())


def slashCommentClicked():
    print "clicked"
    if(window.slash_comment_radio.isChecked()):
        window.hash_comment_radio.setChecked(False)

def hashCommentClicked():
    print "clicked"
    if (window.hash_comment_radio.isChecked()):
        window.slash_comment_radio.setChecked(False)

def editTextChanged():
    content_model.item(currentIndex[-1].row(), currentIndex[-1].column()).setText(clip_edit.toPlainText())

def setTextForEdit(clip):
    window.clip_edit.setText(clip)

def contentClick(index):
    currentIndex.append(index)

    item = content_model.item(index.row(), index.column())
    unicode_item = unicode(item.text(), "utf-8")
    clipboard.setText(unicode_item)
    setTextForEdit(unicode_item)


def sourceClick(index):
    item = source_model.item(index.row(), index.column())
    unicode_item = unicode(item.text(), "utf-8")
    clipboard.setText(unicode_item)
    setTextForEdit(unicode_item)



def sourceRadioClicked(enabled):
    print "clicked"


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

    print "copied"

    print clipboard.text()

    clp.OpenClipboard(None)

    rdf = ''

    rc = clp.EnumClipboardFormats(0)
    # while rc:
    #     try:
    #         format_name = clp.GetClipboardFormatName(rc)
    #     except win32api.error:
    #         format_name = "?"
    #     # print "format", rc, format_name
    #     try:
    #         format = clp.GetClipboardData(rc)
    #     except win32api.error:
    #         format = "?"
    #
    #     if (format_name == 'application/rdf+xml'):
    #         rdf = format
    #     print format_name
    #     print format
    #
    #     rc = clp.EnumClipboardFormats(rc)

    clp.CloseClipboard()

    print rdf


    #credit = Credit(rdf, "https://labs.creativecommons.org/2011/ccrel-guide/examples/image.html")
    # #
    # print str(credit)
    # formattter = TextCreditFormatter()
    # credit.format(formattter)
    # #
    # print formattter.get_text()
    if HTMLClipboard.HasHtml():
        print "HAS HTML"
        source = HTMLClipboard.GetSource()
        print(source)

        if (source == None):
            html = ""
            clp.OpenClipboard(None)
            rc = clp.EnumClipboardFormats(0)
            while rc:
                try:
                    format_name = clp.GetClipboardFormatName(rc)
                except win32api.error:
                    format_name = "?"
                try:
                    format = clp.GetClipboardData(rc)
                except win32api.error:
                    format = "?"
                if (format_name == "HTML Format"):
                    html = format
                rc = clp.EnumClipboardFormats(rc)

            clp.CloseClipboard()
            soup = BeautifulSoup(html, "html.parser")

            link = soup.find('img')['src']
            print link
            source = link

        originalText = clipboard.text()

        if originalText == None or originalText == "":
            originalText = "Image"

        clips.append(tuple((originalText, source)))

        setModelForList()

        if (source_radio.isChecked()):
            wrap_with_comment = "none"
            if window.hash_comment_radio.isChecked():
                wrap_with_comment = "hash"
            elif  window.slash_comment_radio.isChecked():
                wrap_with_comment = "slash"


            if (wrap_with_comment == "hash"):
                clipboard.setText(originalText + "\n" + "# Copied from:\n" + "# " + source)
            elif (wrap_with_comment == "slash"):
                clipboard.setText(originalText + "\n" + "// Copied from:\n" + "// " + source)
            else:
                clipboard.setText(originalText + "\n" + "Copied from:\n" + source)

        print clipboard.text()
    else:
        clp.OpenClipboard(None)

        rc = clp.EnumClipboardFormats(0)
        while rc:
            try:
                format_name = clp.GetClipboardFormatName(rc)
            except win32api.error:
                format_name = "?"
            # print "format", rc, format_name
            try:
                format = clp.GetClipboardData(rc)
            except win32api.error:
                format = "?"
            if (format_name == 'FileNameW'):
                print "File Source Path: \n", format
            rc = clp.EnumClipboardFormats(rc)

        clp.CloseClipboard()

    soup = BeautifulSoup(mimeData.html(), "html.parser")

    links = soup.find_all('a')

    # print links
    for tag in links:
        link = tag.get('href', None)
        if link is not None:
            print link


if __name__ == '__main__':
    main()
