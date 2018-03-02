import win32com.client
from common import *
import fileinput

translated = [];

def initTranslateData():
    conn = initAccessDatabase()
    conn.execute("update [translate] set english = chinese")

    rs = win32com.client.Dispatch(r'ADODB.Recordset')
    tableName = 'translate'
    rs.Open('[' + tableName + ']', conn, 1, 3)
    count = 0
    while not rs.EOF:
        count += 1
        id  = rs.Fields.item('id').value
        english = rs.Fields.item('english').value
        chinese = rs.Fields.item('chinese').value
        translated.append({'id': id, 'english': english, 'chinese': chinese, 'variable': 'LANGUAGE_PACKAGE_ITEM_' + str(count)})
        rs.MoveNext()


def fillToDataBase():
    conn = initAccessDatabase()
    for item in translated:
        id = item['id']
        variable = item['variable']
        english = item['english']
        chinese = item['chinese']
        conn.execute( "update [translate] set variable = '" + variable + "' where id = "+ str(id) )

        cmd = win32com.client.Dispatch(r'ADODB.Command')
        cmd.ActiveConnection = conn
        cmd.CommandType = 1
        cmd.CommandText = "update data set translated_txt = ?, variable = ? where txt = ?"
        cmd.Parameters.Append(cmd.CreateParameter("@translated_txt", 200, 1, 4000, english));
        cmd.Parameters.Append(cmd.CreateParameter("@variable", 200, 1, 255, variable));
        cmd.Parameters.Append(cmd.CreateParameter("@txt", 200, 1, 4000, chinese));
        cmd.execute();

def convertProjectLanguage():
    conn = initAccessDatabase()
    rs = win32com.client.Dispatch(r'ADODB.Recordset')
    rs.Open('select distinct(path) as path from data', conn, 1, 3)
    while not rs.EOF:
        path = rs.Fields.item('path').value
        covnertFileLanguage(path)
        rs.MoveNext()

def covnertFileLanguage( path ):
    conn = initAccessDatabase()
    rs = win32com.client.Dispatch(r'ADODB.Recordset')
    sql = "select * from data where path = '"+ path +"'"
    rs.Open(sql, conn, 1, 3)

    fp = open(path,'r', encoding='utf8')
    lines = fp.readlines()
    fp.close()

    while not rs.EOF:
        path = rs.Fields.item('path').value
        lineNum = rs.Fields.item('line_num').value
        chinese = rs.Fields.item('txt').value
        boundary = rs.Fields.item('boundary').value
        variable = rs.Fields.item('variable').value
        line = lines[lineNum - 1]
        fileType = getFileType(path)
        if fileType == FILE_TYPE_HTML:
            line = line.replace(chinese, "<?php echo "+ variable +"; ?>")
        elif fileType == FILE_TYPE_PHP:
            old = boundary + chinese + boundary
            line = line.replace( old, variable)
        lines[lineNum - 1] = line
        rs.MoveNext()

    fp = open(path,'w', encoding='utf8')
    fp.writelines(lines)
    fp.close()

print('开始计算语言变量...')
initTranslateData()
print('初始化数据库数据...')
fillToDataBase()
print('初始化完成...')
print('开始将项目语言转为英文...')
convertProjectLanguage();
print('转换完成')

