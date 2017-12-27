# ClipSource Python Scripts

clipsource.py:

Extracts the sources of copied objects and puts them on the clipboard.
Available source extractions:
- PDF metadata and citation extraction (only works with Acrobat Reader for now)
- URL extraction from every object copied around the web
- URL extraction from every image copied around the web
- Additional metadata extraction from images copied from www.gettyimages.com
- "Cite this Page" Wikipedia metadata extraction when copying something out of a Wikipedia article

clipsource_server.py:

Supplies the self defined clipboard formats "SOURCE" and "CITATIONS" through a python flask server.
When running, you can access these formats through https://localhost:5000/source.py and https://localhost:5000/citations.py.


How to run:
- Run clipsource.py
- Run clipsource_server.py
- Open ClipSource Word Add-In (https://github.com/DavRet/ClipSourceWordAddIn) Visual Studio solution (ClipSource.sln) and run the Add-In there
