import win32com.client
import re

FILE_TYPE_HTML = 1
FILE_TYPE_PHP = 2
FILE_TYPE_JS = 3
FILE_TYPE_UNDEFINED = -1


def initAccessDatabase():
    conn = win32com.client.Dispatch(r'ADODB.Connection')
    DSN = 'PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=c:\\dev\\LangPackExtractor\\data.mdb;'
    conn.Open(DSN)
    return conn


def getFileType(path):
    ext = re.compile(r'[^.]+$').findall(path)
    if len(ext) == 0:
        return FILE_TYPE_UNDEFINED
    else:
        ext = ext[0].lower()
        if ext == 'html':
            return FILE_TYPE_HTML
        elif ext == 'js':
            return FILE_TYPE_JS
        else:
            return FILE_TYPE_UNDEFINED
