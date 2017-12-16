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
        # Try to get the clipboard's text. If no text is available, set "No Text available" as text
        clp.OpenClipboard(None)
        try:
            text = unicode(clp.GetClipboardData(clp.CF_TEXT), errors='replace')
        except:
            text = "No Text available"

        # Register citation format, so we can check if this format is available
        citation_format = clp.RegisterClipboardFormat("CITATIONS")
        # Checks if citations are available, then returns them as JSON
        if clp.IsClipboardFormatAvailable(citation_format):
            data = clp.GetClipboardData(citation_format)
            clp.CloseClipboard()
            return data

        else:
            # If format is unavailable, set placeholders as data and return JSON object
            clp.CloseClipboard()
            data = {}
            data['APA'] = "no citations in clipboard"
            data['AMA'] = "no citations in clipboard"
            data['content'] = text
            json_data = json.dumps(data)
            return json_data
    except:
        # If anything above fails, close the clipboard und return placeholders as JSON object
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
        # Try to get the clipboard's text. If no text is available, set "No Text available" as text
        clp.OpenClipboard(None)
        try:
            text = unicode(clp.GetClipboardData(clp.CF_TEXT), errors='replace')
        except:
            text = "No Text available"

        # Register source format, so we can check if this format is available
        source_format = clp.RegisterClipboardFormat("SOURCE")
        if clp.IsClipboardFormatAvailable(source_format):
            # Checks if sources are available, then returns them as JSON
            data = clp.GetClipboardData(source_format)
            clp.CloseClipboard()
            return data

        else:
            # If format is unavailable, set placeholders as data and return JSON object
            clp.CloseClipboard()
            data = {}
            data['source'] = "no source in clipboard"
            data['content'] = text
            json_data = json.dumps(data)
            return json_data
    except:
        # If anything above fails, close the clipboard und return placeholders as JSON object
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
