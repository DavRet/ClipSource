## Imports
import argparse
import ctypes
import socket
import sys
import urllib2
from threading import Thread
from time import sleep
import datetime
from ctypes import byref
import PyPDF2 as PyPDF2
import exifread
import psutil as psutil
import pythoncom
import rdflib
import win32con
import win32gui
import win32process
import wikipedia
import bibgen
import win32event
import json
import win32ui
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
# import textract
from crossref.restful import Works
from habanero import Crossref
from habanero import cn
import timeout_decorator
import citeproc
from citeproc import CitationStylesStyle, CitationStylesBibliography
from citeproc import Citation, CitationItem
from citeproc import formatter
from citeproc.source.json import CiteProcJSON
from lxml import html
import pdfminer
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
import scraperwiki
#from libcredit import Credit, HTMLCreditFormatter, TextCreditFormatter
from PySide import QtGui, QtCore
import pygubu
from PyQt4 import QtGui, uic
import win32clipboard as clp, win32api
import pyHook
from flask import Flask

import isbnlib

import re

import base64

import requests


# Application Window (Not used anymore)
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
        # self.show()


# Setup QT for listening to clipboard changes
app = QtGui.QApplication(sys.argv)
clipboard = app.clipboard()

clips = []
currentIndex = []
c = wmi.WMI()

# UI Variables (not used anymore)
window = Window()
content_list = window.content_list
content_model = QtGui.QStandardItemModel(content_list)
source_list = window.source_list
source_model = QtGui.QStandardItemModel(source_list)
source_radio = window.source_radio
clip_edit = window.clip_edit
auto_source_enabled = False

# Copying in PDFs fires multiple events, so the contents are stored here and then checked if they have changed
last_clipboard_content_for_pdf = []

# Wiki Base URLs for retrieving metadata
wikipedia_base_url_german = "https://de.wikipedia.org"
wikipedia_base_url = "https://en.wikipedia.org"
wikimedia_base_url = 'https://commons.wikimedia.org'

# Register custom clipboard formats
citation_format = clp.RegisterClipboardFormat("CITATIONS")
src_format = clp.RegisterClipboardFormat("SOURCE")
metadata_format = clp.RegisterClipboardFormat("METADATA")

# Setup local flask server
server = Flask(__name__)


# Returns current clipboard citations as JSON (https://localhost:5000/citations.py)
@server.route("/citations.py")
def get_citations():
    try:
        clp.OpenClipboard(None)

        try:
            text = unicode(clp.GetClipboardData(clp.CF_TEXT), errors='replace')
        except:
            text = "No Text available"

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
            if format_name == "CITATIONS":
                clp.CloseClipboard()
                return format

            rc = clp.EnumClipboardFormats(rc)

        clp.CloseClipboard()
        data = {}

        data['APA'] = "no citations in clipboard"
        data['AMA'] = "no citations in clipboard"
        data['content'] = text

        json_data = json.dumps(data)
        return json_data
    except:
        try:
            clp.CloseClipboard()
        except:
            print "clipboard already closed"

        data = {}

        data['APA'] = "no citations in clipboard"
        data['AMA'] = "no citations in clipboard"
        data['content'] = "no text in clipboard"

        json_data = json.dumps(data)
        return json_data



# Returns current clipboard source as JSON (https://localhost:5000/source.py)
@server.route("/source.py")
def get_source():
    try:
        clp.OpenClipboard(None)

        try:
            text = unicode(clp.GetClipboardData(clp.CF_TEXT), errors='replace')
        except:
            text = "No Text available"

        formats = {}

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
            if format_name == "SOURCE":
                clp.CloseClipboard()
                return format

            formats[format_name] = format

            rc = clp.EnumClipboardFormats(rc)

        clp.CloseClipboard()

        data = {}

        data['source'] = "no source in clipboard"
        data['content'] = text

        json_data = json.dumps(data)
        return json_data
    except:
        try:
            clp.CloseClipboard()
        except:
            print "clipboard already closed"

        data = {}

        data['source'] = "no source in clipboard"
        data['content'] = "no text in clipboard"

        json_data = json.dumps(data)
        return json_data


# Starts local flask server
def start_server():
    server.run(ssl_context='adhoc')


def reverseImageSearch():
    imageUrl = 'http://cdn.sstatic.net/Sites/stackoverflow/company/img/logos/so/so-icon.png'
    searchUrl = 'https://www.google.com/searchbyimage?site=search&sa=X&image_url='

    url = searchUrl + imageUrl
    f = urllib.urlopen(url)

    soup = BeautifulSoup(f, "lxml")

    print soup


def testEbookExtraction():
    isbn = isbnlib.isbn_from_words("game of thrones")
    # isbn = '3493589182'

    print isbnlib.meta(isbn, service='default', cache='default')


def start_clipboard_watcher():
    clipboard_thread = Thread(target=clipboardChanged)
    clipboard_thread.start()


def main():
    #testEbookExtraction()
    # reverseImageSearch()
    # Generate new thread for local flask server
    #server_thread = Thread(target=start_server)
    #server_thread.start()

    # Connects the clipboard changed event to clipboardChanged function
    clipboard.dataChanged.connect(clipboardChanged)

    # Fixes false decoding errors
    reload(sys)
    sys.setdefaultencoding('Cp1252')



    # UI Setup (not used anymore)
    # UI Model Setup
    content_list.setModel(content_model)
    source_list.setModel(source_model)
    setModelForList()

    # UI Click Listeners
    source_radio.toggled.connect(sourceRadioClicked)
    # window.show()
    content_list.clicked.connect(contentClick)
    source_list.clicked.connect(sourceClick)
    clip_edit.textChanged.connect(editTextChanged)
    window.hash_comment_radio.clicked.connect(hashCommentClicked)
    window.slash_comment_radio.clicked.connect(slashCommentClicked)

    # Starts the QT Application loop which listens to clipboard changes
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
    """Gets application filename given hwnd."""
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
                apa_style = all_text[i + 1] + all_text[i + 2] + all_text[i + 3] + all_text[i + 4] + all_text[i + 5] + \
                            all_text[i + 6]

                citations["APA"] = apa_style.strip('\n')

            if "AMA" in all_text[i]:
                ama_style = all_text[i + 2] + all_text[i + 3] + all_text[i + 4] + all_text[i + 5]
                citations["AMA"] = ama_style.strip('\n')
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
    meta_data_dict = {}

    if "media" in source:
        print "MEDIA" + source
        first_link_fragment = "http://www.gettyimages.de/detail/foto/"
        name_and_id = source.split("/", 4)[4]
        print name_and_id.split("-picture-", 2)
        name = name_and_id.split("-picture-", 2)[0]
        id = name_and_id.split("-picture-", 2)[1].split("id", 2)[1]
        url = first_link_fragment + name + "-lizenzfreies-bild/" + id

        print url

        f = urllib.urlopen(url)
        soup = BeautifulSoup(f, "lxml")

        description = soup.find("meta", itemprop="description")
        author = soup.find("meta", itemprop="author")
        copyright = soup.find("meta", itemprop="copyrightHolder")
        name = soup.find("meta", itemprop="name")

        meta_data_dict['author'] = author['content']
        meta_data_dict['description'] = description['content']
        meta_data_dict['copyright'] = copyright['content']
        meta_data_dict['name'] = name['content']







        # for meta_data_item in soup.findAll("meta"):
        #     meta_data.append(meta_data_item)
    else:
        f = urllib.urlopen(source)
        soup = BeautifulSoup(f, "lxml")

        description = soup.find("meta", itemprop="description")
        author = soup.find("meta", itemprop="author")
        copyright = soup.find("meta", itemprop="copyrightHolder")
        name = soup.find("meta", itemprop="name")

        meta_data_dict['author'] = author['content']
        meta_data_dict['description'] = description['content']
        meta_data_dict['copyright'] = copyright['content']
        meta_data_dict['name'] = name['content']

        for meta_data_item in soup.findAll("meta"):
            meta_data.append(meta_data_item)

    return meta_data_dict


def getAllClipboardFormats():
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
        print format_name
        print format
        rc = clp.EnumClipboardFormats(rc)
    clp.CloseClipboard()


def getWikiMediaMetaData(source):
    url_to_metadata = source.rsplit('/', 1)

    url_to_metadata = url_to_metadata[0].rsplit('/', 1)[1]

    url_to_metadata = 'https://commons.wikimedia.org/wiki/File:' + url_to_metadata

    print url_to_metadata

    citation_url = wikimedia_base_url + getLinkForWikiCitation(url_to_metadata)
    return getWikiCitation(citation_url)

    # f = urllib.urlopen(url_to_metadata)
    # soup = BeautifulSoup(f, "lxml")

    # metadata_items = []
    # for item in soup.findAll(id="stockphoto_dialog"):
    # metadata_items.append(item)
    # print item
    # return metadata_items

def printAllFormats():
    # Opens the Clipboard
    clp.OpenClipboard()
    # Enumerates Clipboard Formats
    clpEnum = clp.EnumClipboardFormats(0)
    # Loops over Format Enumerations
    while clpEnum:
        try:
            # Gets the Format Name
            format_name = clp.GetClipboardFormatName(clpEnum)
        except win32api.error:
            format_name = "not defined"
        try:
            # Gets the Format
            format = clp.GetClipboardData(clpEnum)
        except win32api.error:
            format = "not defined"

        print format_name
        print format
        clpEnum = clp.EnumClipboardFormats(clpEnum)


    # Closes the Clipboard
    clp.CloseClipboard()


def getAsBase64(url):
    return base64.b64encode(requests.get(url).content)

def clipboardChanged():

    #printAllFormats()


    try:
        print "CLIPBOARD CHANGED"
        current_window = win32gui.GetWindowText(win32gui.GetForegroundWindow())

        app_name = get_app_name(win32gui.GetForegroundWindow())

        print app_name
        mimeData = clipboard.mimeData()

        if app_name == 'WINWORD.EXE':
            return

        if "AcroRd32.exe" in app_name or "AcroRd64.exe" in app_name:
            last_clipboard_content_for_pdf.append(clipboard.text())
            if (len(last_clipboard_content_for_pdf) % 5 == 0):
                getPdfMetaData(current_window)
                return
        else:
            if HTMLClipboard.HasHtml():
                source = HTMLClipboard.GetSource()
                print source

                if(source != None):
                    if source != 'about:blank':
                        getMetaDataFromUrl(source)

                    if wikipedia_base_url in source:
                        print "is english wiki"
                        putWikiCitationToClipboard(source)
                        return

                    if wikipedia_base_url_german in source:
                        print "is german wiki"
                        putGermanWikiCitationToClipboard(source)
                        return

                if (source == None):
                    try:
                        print "is image"
                        try:
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
                                    print format
                                    html = format
                                    break
                                rc = clp.EnumClipboardFormats(rc)

                            clp.CloseClipboard()

                            soup = BeautifulSoup(html, "html.parser")

                            link = soup.find('img')['src']
                            print link
                            source = link


                        except:
                            print "could not find source"
                            return

                        if "wikimedia" in current_window.lower():
                            print "is wikimedia"
                            meta_data = getWikiMediaMetaData(source)

                            clp.OpenClipboard(None)

                            sources = {}

                            meta = {}

                            sources['content'] = "image with citation"
                            sources['source'] = source

                            meta['citations'] = meta_data['APA']

                            print meta_data
                            clp.SetClipboardData(src_format, json.dumps(sources))
                            clp.SetClipboardData(citation_format, json.dumps(meta))

                            clp.CloseClipboard()

                            return
                        if "getty" in current_window.lower():

                            meta_data = getGettyImageMetadata(source)

                            clp.OpenClipboard(None)

                            sources = {}

                            meta = {}

                            sources['content'] = "image with metadata"
                            sources['source'] = source

                            meta['citations'] = meta_data

                            print meta_data
                            clp.SetClipboardData(src_format, json.dumps(sources))
                            clp.SetClipboardData(citation_format, json.dumps(meta))

                            print clp.GetClipboardData(citation_format)

                            clp.CloseClipboard()
                            return
                        else:

                            print "is normal image"
                            #
                            # try:
                            #     base64image = getAsBase64(source)
                            # except:
                            #     print "could not find base64"
                            #     base64image = 'no base64'
                            #
                            # print base64image
                            #
                            # # try:
                            # #     thisSrc = clp.GetClipboardData(src_format)
                            # #     clp.CloseClipboard()
                            # #
                            # #
                            # # except:
                            # #     print "exception here"
                            # #     clp.CloseClipboard()
                            #
                            try:
                                clp.OpenClipboard(None)

                                try:
                                    src = clp.GetClipboardData(src_format)
                                except:
                                    src = 'none'


                                print "SOURCE ", src
                                if "image" in src:
                                    clp.CloseClipboard()
                                    return
                                else:
                                    sources = {}

                                    sources['content'] = "image"
                                    sources['source'] = source

                                    clp.OpenClipboard(None)
                                    clp.SetClipboardData(src_format, json.dumps(sources))
                                    print clp.GetClipboardData(src_format)

                                    clp.CloseClipboard()
                                    return
                            except:
                                sources = {}

                                sources['content'] = "image"
                                sources['source'] = source

                                clp.OpenClipboard(None)
                                clp.SetClipboardData(src_format, json.dumps(sources))
                                print clp.GetClipboardData(src_format)

                                clp.CloseClipboard()

                                return
                    except:
                        print "image exception"



                clp.OpenClipboard(None)
                sources = {}
                sources['source'] = source
                sources['content'] = unicode(clp.GetClipboardData(clp.CF_TEXT), errors='replace')

                clp.SetClipboardData(src_format, json.dumps(sources))

                clp.CloseClipboard()

                originalText = clipboard.text()

                if originalText == None or originalText == "":
                    originalText = "Image"

                # clips.append(tuple((originalText, source)))
                #
                # setModelForList()
                #
                # if (source_radio.isChecked()):
                #     wrap_with_comment = "none"
                #     if window.hash_comment_radio.isChecked():
                #         wrap_with_comment = "hash"
                #     elif window.slash_comment_radio.isChecked():
                #         wrap_with_comment = "slash"
                #
                #     if (wrap_with_comment == "hash"):
                #         clipboard.setText(originalText + "\n" + "# Copied from:\n" + "# " + source)
                #     elif (wrap_with_comment == "slash"):
                #         clipboard.setText(originalText + "\n" + "// Copied from:\n" + "// " + source)
                #     else:
                #         clipboard.setText(originalText + "\n" + "Copied from:\n" + source)
                #
                # print clipboard.text()
            else:
                checkForFile()
    except:
        try:
            clp.CloseClipboard()
        except:
            print "Clipboard was not open"
        print "exception"


def getMetaDataFromUrl(url):
    f = urllib.urlopen(url)
    soup = BeautifulSoup(f, "lxml")

    metadata_items = []
    for item in soup.findAll("meta"):
        metadata_items.append(item)


def getGermanWikiCitation(url):
    f = urllib.urlopen(url)

    soup = BeautifulSoup(f, "lxml")

    first_paragraph = soup.find("p")
    citation = ''
    for item in first_paragraph.findAll(text=True):
        citation = citation + item

    return citation


def putGermanWikiCitationToClipboard(source):
    citation_link = getLinkForWikiCitation(source)
    wiki_citation = getGermanWikiCitation(wikipedia_base_url_german + citation_link)

    clp.OpenClipboard(None)

    citations = {}

    citations['APA'] = wiki_citation
    citations['content'] = unicode(clp.GetClipboardData(clp.CF_TEXT), errors='replace')

    clp.SetClipboardData(citation_format, json.dumps(citations))

    sources = {}

    sources['source'] = source
    sources['content'] = unicode(clp.GetClipboardData(clp.CF_TEXT), errors='replace')

    clp.SetClipboardData(src_format, json.dumps(sources))

    print clp.GetClipboardData(citation_format)

    clp.CloseClipboard()


def putWikiCitationToClipboard(source):
    citation_link = getLinkForWikiCitation(source)
    wiki_citation = getWikiCitation(wikipedia_base_url + citation_link)

    id = citation_link.split("id=", 1)[1]

    name = source.split("/wiki/", 1)[1]

    # wiki_page = wikipedia.page(wikipedia.suggest(name))

    # references_in_page = wiki_page.references
    print wiki_citation
    clp.OpenClipboard(None)

    citations = {}

    citations['AMA'] = wiki_citation['AMA']
    citations['APA'] = wiki_citation['APA']
    citations['content'] = unicode(clp.GetClipboardData(clp.CF_TEXT), errors='replace')

    clp.SetClipboardData(citation_format, json.dumps(citations))

    sources = {}

    sources['source'] = source
    sources['content'] = unicode(clp.GetClipboardData(clp.CF_TEXT), errors='replace')

    clp.SetClipboardData(src_format, json.dumps(sources))

    print clp.GetClipboardData(src_format)

    clp.CloseClipboard()


def getPdfMetaData(current_window):
    pdf_title = current_window.split(" - ", 1)[0]
    pdf_path = ''

    process_id = win32process.GetWindowThreadProcessId(win32gui.GetForegroundWindow())

    p = psutil.Process(process_id[1])
    files = p.open_files()

    for file in files:
        if pdf_title in str(file):
            print file[0]
            pdf_path = file[0]

    if pdf_path == '':
        pdf_path = find(pdf_title, "C:\Users\David\Desktop")

    fp = open(pdf_path, 'rb')
    parser = PDFParser(fp)
    doc = PDFDocument(parser)

    pdf_file = open(pdf_path, 'rb')

    # Extracting Text from PDF
    # read_pdf = PyPDF2.PdfFileReader(pdf_file)
    # number_of_pages = read_pdf.getNumPages()
    # page = read_pdf.getPage(0)
    # page_content = page.extractText()
    # print page_content

    # text = textract.process(pdf_path, method='pdfminer')
    # print text



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

    data_for_clipboard = {"Title": title, "Author": author, "Date": str(date), "Keywords": keywords}

    getCrossRefMetaData(title, pdf_path)


def prettifyUTF8Strings(string):
    return re.sub(ur'[\xc2-\xf4][\x80-\xbf]+', lambda m: m.group(0).encode('latin1').decode('utf8'), string)

def getCrossRefMetaData(title, path):
    print title

    print "getting crossref"

    cr = Crossref()
    query = cr.works(query=title, limit=1)

    doi = ''

    for item in query['message']['items']:
        doi = item['DOI']

    #print isbnlib.doi2tex(doi)

    apa_citation = cn.content_negotiation(ids=doi, format="text", style="apa")
    #rdf_citation = cn.content_negotiation(ids=doi, format="rdf-xml")

    #json_citation = cn.content_negotiation(ids=doi, format="citeproc-json")

    #bib_entry = cn.content_negotiation(ids=doi, format="bibentry")

    apa_citation = prettifyUTF8Strings(apa_citation).strip('\n')

    clp.OpenClipboard(None)

    citations = {}

    citations['APA'] = apa_citation
    try:
        citations['content'] = unicode(clp.GetClipboardData(clp.CF_TEXT), errors='replace')
    except:
        citations['content'] = 'no text content available'

    clp.SetClipboardData(citation_format, json.dumps(citations))

    sources = {}

    sources['source'] = path
    try:
        sources['content'] = unicode(clp.GetClipboardData(clp.CF_TEXT), errors='replace')
    except:
        sources['content'] = 'no text content available'

    clp.SetClipboardData(src_format, json.dumps(sources))

    print clp.GetClipboardData(citation_format)

    clp.CloseClipboard()


def checkForFile():
    try:
        clp.OpenClipboard()

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
            if (format_name == '?'):
                print format
                if (len(format) > 0):
                    print "File Source Path: \n", format[0]

                    path = format[0]

                    if (path[-4:] == '.pdf'):
                        fp = open(path, 'rb')
                        parser = PDFParser(fp)
                        doc = PDFDocument(parser)

                        if 'Title' in doc.info[0]:
                            title = doc.info[0]['Title']
                        else:
                            title = "No Title found"

                        getCrossRefMetaData(title, path)

            rc = clp.EnumClipboardFormats(rc)

        clp.CloseClipboard()

    except:
        print "error"


def find(name, path):
    for root, dirs, files in scandir.walk(path):
        if name in files:
            return os.path.join(root, name)


if __name__ == '__main__':
    main()
