from flask import Flask

import win32clipboard as clp
import win32api
import json
from OpenSSL import SSL



app = Flask(__name__)

# Returns current clipboard citations as JSON (https://localhost:5000/citations.py)
@app.route("/citations.py")
def get_citations():
    try:
        clp.OpenClipboard(None)

        try:
            text = unicode(clp.GetClipboardData(clp.CF_TEXT), errors='replace')
        except:
            text = "No Text available"

        citation_format = clp.RegisterClipboardFormat("CITATIONS")
        if clp.IsClipboardFormatAvailable(citation_format):
            data = clp.GetClipboardData(citation_format)
            clp.CloseClipboard()
            return data

        else:
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
@app.route("/source.py")
def get_source():
    try:
        clp.OpenClipboard(None)

        try:
            text = unicode(clp.GetClipboardData(clp.CF_TEXT), errors='replace')
        except:
            text = "No Text available"

        source_format = clp.RegisterClipboardFormat("SOURCE")
        if clp.IsClipboardFormatAvailable(source_format):
            data = clp.GetClipboardData(source_format)
            clp.CloseClipboard()
            return data

        else:
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

if __name__ == "__main__":
    app.run(ssl_context='adhoc')
