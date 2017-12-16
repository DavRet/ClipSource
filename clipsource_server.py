from flask import Flask
import win32clipboard as clp
import json


# Setup server
app = Flask(__name__)

@app.route("/citations.py")
def get_citations():
    """
    Returns the contents of the specified CITATIONS format when opening following URL: https://localhost:5000/source.py

    :return: Current clipboard citations in JSON format
    """
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
    """
    Returns the contents of the specified SOURCE format when opening following URL: https://localhost:5000/source.py

    :return: Current clipboard source in JSON format
    """
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
    # Run flask server without certificate
    app.run(ssl_context='adhoc')
