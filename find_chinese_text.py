import os
import re
import win32com.client
from common import initAccessDatabase
from common import *

result = []


def exportToAccess():
    conn = initAccessDatabase()
    conn.execute('delete from data')
    for item in result:
        lineNum = str(item["line_num"])
        path = item["path"]
        txt = item["txt"]
        # txt = re.compile("'").sub("''",txt)
        boundary = item['boundary']
        # boundary =  re.compile("'").sub("''",boundary)
        fileType = item['file_type']
        # conn.execute("insert into data(line_num, path, txt, boundary, file_type) values ("+ lineNum +",'"+ path +"','"+ txt +"','"+ boundary +"', "+ str(fileType) +")")
        cmd = win32com.client.Dispatch(r'ADODB.Command')
        cmd.ActiveConnection = conn
        cmd.CommandType = 1
        cmd.CommandText = "insert into data(line_num, path, txt, boundary, file_type) values (?,?,?,?,?)"
        cmd.Parameters.Append(cmd.CreateParameter("@line_num", 3, 1, 10, lineNum))
        cmd.Parameters.Append(cmd.CreateParameter("@path", 200, 1, 255, path))
        cmd.Parameters.Append(cmd.CreateParameter("@txt", 200, 1, 4000, txt))
        cmd.Parameters.Append(cmd.CreateParameter("@boundary", 200, 1, 255, boundary))
        cmd.Parameters.Append(cmd.CreateParameter("@file_type", 3, 1, 10, fileType))
        cmd.execute()

    conn.execute('delete from [translate]')
    sql = "insert into [translate](chinese) select distinct(txt) from data"
    conn.execute(sql)


def listDir(rootDir):
    for lists in os.listdir(rootDir):
        path = os.path.join(rootDir, lists)
        if os.path.isfile(path):
            fileType = getFileType(path)
            if fileType == FILE_TYPE_HTML:
                handleHtmlFile(path)
            elif fileType == FILE_TYPE_JS:
                handleJS(path)
        if os.path.isdir(path):
            listDir(path)



def handleFile(path, findTextPattern, onMatch):
    replaceCommentPattern = re.compile(r'//.+')
    commentStartPattern = re.compile(r'/\*')
    commentEndPattern = re.compile(r'\*/')
    phpStartPattern = re.compile(r'^ *<script')
    phpEndPattern = re.compile(r'</script>')

    lineNum = 1
    f = open(path, 'r', encoding='utf8')
    line = f.readline()
    isComment = False
    isScriptInHtmlFile = False
    fileType = getFileType(path)
    while line:
        if not isComment and commentStartPattern.search(line):
            isComment = True

        if fileType == FILE_TYPE_HTML and not isScriptInHtmlFile and phpStartPattern.search(line):
            isScriptInHtmlFile = True

        if not isComment and not isScriptInHtmlFile:
            line = replaceCommentPattern.sub('', line)
            matches = findTextPattern.finditer(line)
            for match in matches:
                if len(match.group()) > 1000:
                    continue
                onMatch(path, lineNum, match)

        if isComment and commentEndPattern.search(line):
            isComment = False

        if fileType == FILE_TYPE_HTML and isScriptInHtmlFile and phpEndPattern.search(line):
            isScriptInHtmlFile = False

        line = f.readline()
        lineNum += 1
    f.close()


def handleHtmlFileOnMatch(path, lineNum, match):
    result.append({"txt": match.group(), "line_num": lineNum, "path": path, 'boundary': '', 'file_type': FILE_TYPE_HTML})

def handleHtmlFile(path):
    findTextPattern = re.compile(r'([\u4e00-\u9fa5]+[^<\'\"\|]+)*[\u4e00-\u9fa5]+')
    handleFile(path, findTextPattern, handleHtmlFileOnMatch)

def handleJS(path):
    return

# def handlePhpFileOnMatch(path, lineNum, match,isScriptInHtmlFile):
#     txt = match.group()
#     boundary = txt[0]
#     txt = re.sub("^('|\")|('|\")$", '', txt)
#     result.append({"txt": txt, 'boundary': boundary, "line_num": lineNum, "path": path, 'file_type': FILE_TYPE_PHP})

# def handlePhpFile(path):
#     regPart1 = r'"[^"\n\r\$]*?(?=[\u4e00-\u9fa5])[^"\n\r\$]*?(?<!\\)"'
#     regPart2 = r"'[^\'\n\r\$]*?(?=[\u4e00-\u9fa5])[^'\n\r\$]*?(?<!\\)'"
#     reg = regPart1 + '|' + regPart2
#     findTextPattern = re.compile(reg)
#     handleFile(path, findTextPattern, handlePhpFileOnMatch)


def printResult():
    for item in result:
        print(item)

PROJECT_PATH = 'C:\\dev\\LangPackExtractor\\community_for_ios'
listDir(PROJECT_PATH)
printResult()


'''
print('开始遍历项目，寻找中文...')
# listDir(PROJECT_PATH)

for path in pathList:
    listDir(path)

print('共有' + str(len(result)) + '处中文')
print('开始将数据写入数据库')
exportToAccess()
print('写入数据库完成')
'''
