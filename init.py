import sys
import urllib2

import datetime
import exifread
import rdflib
import win32con
import win32gui

import win32process

import wikipedia
import bibgen

from PyQt4.QtCore import QTimer
from bs4 import BeautifulSoup

import HTMLClipboard
import os
import os.path
import sys
import pyapa
import wmi
import scandir
import urllib

import citeproc
from citeproc import CitationStylesStyle, CitationStylesBibliography
from citeproc import Citation, CitationItem
from citeproc import formatter
from citeproc.source.json import CiteProcJSON


from lxml import html

import pdfminer
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument

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

c = wmi.WMI()


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

last_clipboard_content_for_pdf = []

wikipedia_base_url = "https://en.wikipedia.org"


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

    sys.exit(app.exec_())



def get_app_path(hwnd):
    """Get applicatin path given hwnd."""
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        for p in c.query('SELECT ExecutablePath FROM Win32_Process WHERE ProcessId = %s' % str(pid)):
            exe = p.ExecutablePath
            break
    except:
        return None
    else:
        return exe


def get_app_name(hwnd):
    """Get applicatin filename given hwnd."""
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        for p in c.query('SELECT Name FROM Win32_Process WHERE ProcessId = %s' % str(pid)):
            exe = p.Name
            break
    except:
        return None
    else:
        return exe

def getLinkForWikiCitation(url):
    f = urllib.urlopen(url)
    soup = BeautifulSoup(f, "lxml")

    citeUrl = soup.find('li', attrs={'id': 't-cite'})

    for tag in citeUrl:
        link = tag.get('href', None)
        if link is not None:
            print link
            return link
        else:
            return "no link"

def getWikiCitation(url):

    f = urllib.urlopen(url)

    soup = BeautifulSoup(f, "lxml")

    div = soup.findAll("div", {"class": "plainlinks"})

    citations = {}

    for item in div:
        bibTexLaTex = item.findAll("pre")

        for i in bibTexLaTex:
            citations["BibLatex"] = i

        all_text = item.findAll(text=True)

        for i in range(len(all_text)):

            if "APA" in all_text[i]:

                apa_style = all_text[i+1] + all_text[i+2] + all_text[i+3] + all_text[i+4] + all_text[i+5] + all_text[i+6]

                citations["APA"] = apa_style

            if "AMA" in all_text[i]:
                ama_style = all_text[i+1] + all_text[i+2] + all_text[i+3] + all_text[i+4] + all_text[i+5]
                citations["AMA"] = ama_style
    return citations


def formatToAPA(author, created, publisher, retrieved, source):
    print "test"

def slashCommentClicked():
    if (window.slash_comment_radio.isChecked()):
        window.hash_comment_radio.setChecked(False)


def hashCommentClicked():
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


def getGettyImageMetadata(source):
    meta_data = []
    if "media" in source:
        print "MEDIA"+source
        first_link_fragment = "http://www.gettyimages.de/detail/foto/"
        name_and_id = source.split("/", 4)[4]
        print name_and_id.split("-picture-", 2)
        name = name_and_id.split("-picture-", 2)[0]
        id = name_and_id.split("-picture-", 2)[1].split("id", 2)[1]
        url = first_link_fragment + name + "-lizenzfreies-bild/" +  id

        print url

        f = urllib.urlopen(url)
        soup = BeautifulSoup(f, "lxml")

        description = soup.find("meta", itemprop="description")

        for meta_data_item in soup.findAll("meta"):
            meta_data.append(meta_data_item)
    else:
        f = urllib.urlopen(source)
        soup = BeautifulSoup(f, "lxml")

        for meta_data_item in soup.findAll("meta"):
            meta_data.append(meta_data_item)

    return meta_data


def clipboardChanged():
    window = win32gui.GetWindowText(win32gui.GetForegroundWindow())

    app_name = get_app_name(win32gui.GetForegroundWindow())
    mimeData = clipboard.mimeData()

    print clipboard.text()



    if ".pdf" in window:
        last_clipboard_content_for_pdf.append(clipboard.text())
        if (len(last_clipboard_content_for_pdf) % 5 == 0):
            print "IS PDF"
            pdf_title = window.split(" - ", 1)[0]
            pdf_path = find(pdf_title, "C:\Users\David\Desktop")
            fp = open(pdf_path, 'rb')
            parser = PDFParser(fp)
            doc = PDFDocument(parser)

            if 'Author' in doc.info[0]:
                author = doc.info[0]['Author']
            else:
                author = "No Author found"
            if 'Title' in doc.info[0]:
                title = doc.info[0]['Title']
            else:
                title = "No Title found"
            if 'CreationDate' in doc.info[0]:
                created = doc.info[0]['CreationDate']
            else:
                created = "No Date found"
            if 'Keywords' in doc.info[0]:
                keywords = doc.info[0]['Keywords']
            else:
                keywords = "No Keywords found"

            new_date = created.split(":", 1)[1]
            year = new_date[:4]
            month = new_date[4:][:2]
            day = new_date[6:][:2]

            date = datetime.date(int(year), int(month), int(day))


            clipboard.setText(clipboard.text() + "\n" + "Title: " + title
                                 + "\n" + "Authors: " + author + "\n" + "Date: " + str(date) + "\n" +
                                 "Keywords: " + keywords)

    else:

        # clp.OpenClipboard(None)
        #
        # rdf = ''
        #
        # rc = clp.EnumClipboardFormats(0)
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
        #
        # clp.CloseClipboard()
        #
        # print rdf


        # credit = Credit(rdf, "https://labs.creativecommons.org/2011/ccrel-guide/examples/image.html")
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
            if "wikipedia" in window.lower():
                print getWikiCitation(wikipedia_base_url + getLinkForWikiCitation(source))




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

            if "getty" in window.lower():
                print getGettyImageMetadata(source)


            originalText = clipboard.text()

            if originalText == None or originalText == "":
                originalText = "Image"

            clips.append(tuple((originalText, source)))

            setModelForList()

            if (source_radio.isChecked()):
                wrap_with_comment = "none"
                if window.hash_comment_radio.isChecked():
                    wrap_with_comment = "hash"
                elif window.slash_comment_radio.isChecked():
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
                if (format_name == '?'):
                    if (len(format) > 0):
                        print "File Source Path: \n", format[0]

                        path = format[0]

                        if (path[-4:] == '.pdf'):
                            fp = open(path, 'rb')
                            parser = PDFParser(fp)
                            doc = PDFDocument(parser)

                            print doc.info

                rc = clp.EnumClipboardFormats(rc)

            clp.CloseClipboard()

        # soup = BeautifulSoup(mimeData.html(), "html.parser")
        #
        # links = soup.find_all('a')
        #
        # # print links
        # for tag in links:
        #     link = tag.get('href', None)
        #     if link is not None:
        #         print link


def find(name, path):
    for root, dirs, files in scandir.walk(path):
        if name in files:
            return os.path.join(root, name)


if __name__ == '__main__':
    main()
