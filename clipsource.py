## Imports
from threading import Thread
import datetime
import psutil as psutil
import win32gui
import win32process
import json
from bs4 import BeautifulSoup
import HTMLClipboard
import os
import os.path
import sys
import wmi
import scandir
import urllib
from habanero import Crossref
from habanero import cn
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from PySide import QtGui, QtCore
from PyQt4 import QtGui, uic
import win32clipboard as clp, win32api
import isbnlib
import re
import base64
import requests

# Setup QT for listening to clipboard changes
app = QtGui.QApplication(sys.argv)
clipboard = app.clipboard()

clips = []
currentIndex = []

# Used for getting application name and path
c = wmi.WMI()

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

# Array to check window ID when copying out of PDF
pdf_window_id = []


def main():
    """
    Main method, sets up clipboard connection and PyQt Loop
    """

    # Connects the clipboard changed event to clipboardChanged function
    clipboard.dataChanged.connect(clipboardChanged)

    # Fixes false decoding errors
    reload(sys)
    sys.setdefaultencoding('Cp1252')

    # Starts the QT Application loop which listens to clipboard changes
    sys.exit(app.exec_())


def testIsbnExtraction():
    """
    For further usage, not used at the moment
    """

    isbn = isbnlib.isbn_from_words("to kill a mockingbird")
    # isbn = '3493589182'

    print isbnlib.meta(isbn, service='default', cache='default')


def printAllFormats():
    """
    Prints all current clipboard formats and their names, for testing purposes only
    """

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

def get_app_path(hwnd):
    """
    Gets an application's path, given it's window
    :param hwnd: Window handle
    :return: Application path
    """
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
    """
    Gets an application's name, given it's window
    :param hwnd: Window handle
    :return: Application name
    """
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
    """
    Builds link to Wiki Citation page
    :param url: URL of Wiki article
    :return: Link to Citation page
    """
    f = urllib.urlopen(url)
    soup = BeautifulSoup(f, "lxml")

    citeUrl = soup.find('li', attrs={'id': 't-cite'})

    for tag in citeUrl:
        link = tag.get('href', None)
        if link is not None:
            return link
        else:
            return "no link"


def getWikiCitation(url):
    """
    Gets the contents of Wikipedia "Cite this page"
    :param url: URL of Wiki article
    :return: Array with the sources of "Cite this page"
    """
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



def getGettyImageMetadata(source):
    """
    Gets metadata of images copied on "Getty Images"
    :param source: Source URL of the copied image
    :return: Metadata dictionary
    """

    # Metadata array, for testing purposes
    meta_data = []

    # Metadata dictionary, will be filled and returned
    meta_data_dict = {}

    if "media" in source:
        # If URL cointains "media", we have to reconstruct the URL of the gallery, containing the image

        first_link_fragment = "http://www.gettyimages.de/detail/foto/"
        name_and_id = source.split("/", 4)[4]
        print name_and_id.split("-picture-", 2)
        name = name_and_id.split("-picture-", 2)[0]
        id = name_and_id.split("-picture-", 2)[1].split("id", 2)[1]

        # Reconstructed URL
        url = first_link_fragment + name + "-lizenzfreies-bild/" + id

        f = urllib.urlopen(url)
        soup = BeautifulSoup(f, "lxml")

        # Extract metadata
        description = soup.find("meta", itemprop="description")
        author = soup.find("meta", itemprop="author")
        copyright = soup.find("meta", itemprop="copyrightHolder")
        name = soup.find("meta", itemprop="name")

        # Add metadata to dictionary
        meta_data_dict['author'] = author['content']
        meta_data_dict['description'] = description['content']
        meta_data_dict['copyright'] = copyright['content']
        meta_data_dict['name'] = name['content']

        # Gets all metadata, not used at the moment
        # for meta_data_item in soup.findAll("meta"):
        #     meta_data.append(meta_data_item)
    else:
        # Here, the URL is fine

        f = urllib.urlopen(source)
        soup = BeautifulSoup(f, "lxml")

        # Extracts metadata
        description = soup.find("meta", itemprop="description")
        author = soup.find("meta", itemprop="author")
        copyright = soup.find("meta", itemprop="copyrightHolder")
        name = soup.find("meta", itemprop="name")

        # Add metadata to dictionary
        meta_data_dict['author'] = author['content']
        meta_data_dict['description'] = description['content']
        meta_data_dict['copyright'] = copyright['content']
        meta_data_dict['name'] = name['content']

        # Gets all metadata, not used at the moment
        for meta_data_item in soup.findAll("meta"):
            meta_data.append(meta_data_item)

    return meta_data_dict



def getWikiMediaMetaData(source):
    """
    Gets "Cite this page" metadata form Wikimedia
    :param source: Source URL of copied object
    :return: List of "Cite this page" sources
    """

    # First we have to build the right URL
    url_to_metadata = source.rsplit('/', 1)
    url_to_metadata = url_to_metadata[0].rsplit('/', 1)[1]
    url_to_metadata = 'https://commons.wikimedia.org/wiki/File:' + url_to_metadata

    # Build the URL to "Cite this page"
    citation_url = wikimedia_base_url + getLinkForWikiCitation(url_to_metadata)

    # Gets the contents of "Cite this page" und returns them
    return getWikiCitation(citation_url)



def getAsBase64(url):
    """
    Helper to convert images by their URL to Base64
    :param url: Image url
    :return: Base64 image
    """
    return base64.b64encode(requests.get(url).content)

def prettifyUTF8Strings(string):
    """
    Helper to prettify UTF8 strings
    :param string: "ugly" string
    :return: pretty string
    """
    return re.sub(ur'[\xc2-\xf4][\x80-\xbf]+', lambda m: m.group(0).encode('latin1').decode('utf8'), string)

def find(name, path):
    """
    Recursive search for absolute path of file, not used anymore. Was used when trying to find out the absolute paths of PDF documents
    :param name: Name to search for
    :param path: Folder to search in
    :return:
    """
    for root, dirs, files in scandir.walk(path):
        if name in files:
            return os.path.join(root, name)



def clipboardChanged():
    """
    Gets called everytime the clioboard data is changed. Source extraction always begins here.
    :return: Returns in this method are just used to prevent code from continueing. Actually the extracted sources are then "returned" to the clipboard
    """
    try:
        print "CLIPBOARD CHANGED"
        current_window = win32gui.GetWindowText(win32gui.GetForegroundWindow())

        app_name = get_app_name(win32gui.GetForegroundWindow())

        mimeData = clipboard.mimeData()

        if app_name == 'WINWORD.EXE':
            return

        if "AcroRd32.exe" in app_name or "AcroRd64.exe" in app_name:
            last_clipboard_content_for_pdf.append(clipboard.text())

            if (".pdf" in win32gui.GetWindowText(win32gui.GetForegroundWindow())):
                pdf_window_id.append(win32gui.GetForegroundWindow())

            if (len(last_clipboard_content_for_pdf) % 5 == 0):
                print pdf_window_id[-1]

                process_id = win32process.GetWindowThreadProcessId(pdf_window_id[-1])
                getPdfMetaData(current_window, process_id)
                return
        else:
            if HTMLClipboard.HasHtml():
                source = HTMLClipboard.GetSource()
                print source

                if (source != None):
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

                            clp.CloseClipboard()
                            return
                        else:

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

            else:
                # If clipboard has no HTML-content, check if it's a file
                checkForFile()
    except:
        # Close the clipboard, if there was an exception
        try:
            clp.CloseClipboard()
        except:
            print "Clipboard was not open"
        print "exception"


def getMetaDataFromUrl(url):
    """
    Gets of available metadata from an URL
    :param url: URL of copied object
    :return: returns list of metadata items
    """
    f = urllib.urlopen(url)
    soup = BeautifulSoup(f, "lxml")

    metadata_items = []
    for item in soup.findAll("meta"):
        metadata_items.append(item)
    return metadata_items


def getGermanWikiCitation(url):
    """
    Gets contents of german wiki "Cite this page" page
    :param url: Wiki article URL
    :return: APA citation
    """
    f = urllib.urlopen(url)

    soup = BeautifulSoup(f, "lxml")

    first_paragraph = soup.find("p")
    citation = ''
    for item in first_paragraph.findAll(text=True):
        citation = citation + item

    return citation


def putGermanWikiCitationToClipboard(source):
    """
    Puts sources of a german Wiki page to the clipboard
    :param source:
    """

    # First we have to get the link to "Cite this Page"
    citation_link = getLinkForWikiCitation(source)

    # Then we get the contents from this link
    wiki_citation = getGermanWikiCitation(wikipedia_base_url_german + citation_link)

    clp.OpenClipboard(None)
    citations = {}
    citations['APA'] = wiki_citation
    citations['content'] = unicode(clp.GetClipboardData(clp.CF_TEXT), errors='replace')
    # Puts source data on the clipboard
    clp.SetClipboardData(citation_format, json.dumps(citations))

    sources = {}
    sources['source'] = source
    sources['content'] = unicode(clp.GetClipboardData(clp.CF_TEXT), errors='replace')
    # Puts the citation data on the clipboard
    clp.SetClipboardData(src_format, json.dumps(sources))

    clp.CloseClipboard()


def putWikiCitationToClipboard(source):
    """
    Puts the sources of a english Wiki page on the clipboard
    :param source: Source URL of Wiki article
    """

    # Gets the link to "Cite this Page"
    citation_link = getLinkForWikiCitation(source)

    # Gets the sources of "Cite this Page"
    wiki_citation = getWikiCitation(wikipedia_base_url + citation_link)

    # Testing the wikipedia python module, not used
    #id = citation_link.split("id=", 1)[1]
    #name = source.split("/wiki/", 1)[1]
    # wiki_page = wikipedia.page(wikipedia.suggest(name))
    # references_in_page = wiki_page.references

    clp.OpenClipboard(None)
    citations = {}
    citations['AMA'] = wiki_citation['AMA']
    citations['APA'] = wiki_citation['APA']
    citations['content'] = unicode(clp.GetClipboardData(clp.CF_TEXT), errors='replace')
    # Put the citations on the clipboard
    clp.SetClipboardData(citation_format, json.dumps(citations))

    sources = {}
    sources['source'] = source
    sources['content'] = unicode(clp.GetClipboardData(clp.CF_TEXT), errors='replace')
    # Put the sources on the clipboard
    clp.SetClipboardData(src_format, json.dumps(sources))

    clp.CloseClipboard()


def getPdfMetaData(current_window, process_id):
    """
    Extracts PDF metadata when copying in a PDF document
    :param current_window: Window of PDF reader
    :param process_id: Process ID of PDF reader
    """

    # Get the PDF title
    pdf_title = current_window.split(" - ", 1)[0]
    pdf_path = ''

    # Gets the filename by looking up which files are openend by process ID
    p = psutil.Process(process_id[1])
    files = p.open_files()

    for file in files:
        # Checks if title is in path (we need this in case multiple documents are openend in the reader
        if pdf_title in str(file):
            pdf_path = file[0]

    # Open the file by it's path
    fp = open(pdf_path, 'rb')

    # Parses the PDF document and extracts it's metadata
    parser = PDFParser(fp)
    doc = PDFDocument(parser)


    # Testing PDF text extraction, not used anymore, left for testing purposes only
    # pdf_file = open(pdf_path, 'rb')
    # Extracting Text from PDF
    # read_pdf = PyPDF2.PdfFileReader(pdf_file)
    # number_of_pages = read_pdf.getNumPages()
    # page = read_pdf.getPage(0)
    # page_content = page.extractText()
    # print page_content
    # text = textract.process(pdf_path, method='pdfminer')
    # print text

    # Extracts Author, Title, Date and Keyword from PDf document
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

    # We have to reconstruct the date because it's not formatted well
    new_date = created.split(":", 1)[1]
    year = new_date[:4]
    month = new_date[4:][:2]
    day = new_date[6:][:2]
    date = datetime.date(int(year), int(month), int(day))

    # Before getting the Crossref data, this data was set to the clipboard
    data_for_clipboard = {"Title": title, "Author": author, "Date": str(date), "Keywords": keywords}

    # Gets the Crossref metadata by the title of the document. Also contains the path as parameter, because it's later saved as "Source" of the document
    getCrossRefMetaData(title, pdf_path)




def getCrossRefMetaData(title, path):
    """
    Gets Crossref metadata, given an article's title. Then puts the metadata on the clipboard
    :param title: Title to search for
    :param path: PDF-Path, not necessary
    """

    print "getting crossref"


    # Searches the Crossref API for the given title, gets best result
    cr = Crossref()
    query = cr.works(query=title, limit=1)

    doi = ''

    # Extract DOI out of Crossref answer
    for item in query['message']['items']:
        doi = item['DOI']


    # Not used, but useful. Gets metadata from isbnlib, given DOI
    # print isbnlib.doi2tex(doi)

    # Gets APA citation, given DOI
    apa_citation = cn.content_negotiation(ids=doi, format="text", style="apa")

    # We could get more formats this way, but this is not used at the moment, better performance without getting these formats
    # rdf_citation = cn.content_negotiation(ids=doi, format="rdf-xml")
    # json_citation = cn.content_negotiation(ids=doi, format="citeproc-json")
    # bib_entry = cn.content_negotiation(ids=doi, format="bibentry")


    # Prettify APA citation
    apa_citation = prettifyUTF8Strings(apa_citation).strip('\n')

    clp.OpenClipboard(None)
    citations = {}
    citations['APA'] = apa_citation
    try:
        citations['content'] = unicode(clp.GetClipboardData(clp.CF_TEXT), errors='replace')
    except:
        citations['content'] = 'no text content available'
    # Puts the citations on the clipboard
    clp.SetClipboardData(citation_format, json.dumps(citations))

    sources = {}
    sources['source'] = path
    try:
        sources['content'] = unicode(clp.GetClipboardData(clp.CF_TEXT), errors='replace')
    except:
        sources['content'] = 'no text content available'
    # Puts the sources on the clipboard
    clp.SetClipboardData(src_format, json.dumps(sources))
    clp.CloseClipboard()

def checkForFile():
    """
    Checks if copied object was a file and then tries to get it's path
    """
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

                    # Path of the copied file
                    path = format[0]

                    # If path contains .pdf, extract it's metadata
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


if __name__ == '__main__':
    main()
