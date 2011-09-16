import types

__author__ = 'Mark.Wei'

from xlrd import open_workbook



def generateButton(filename=None):
    wb = open_workbook("g:/%s.xls" % filename,'rb')

    sheet = wb.sheet_by_index(0)
    title = sheet.row_values(0)

    print title

    rownum = range(1,sheet.nrows)
    closnum = range(1,sheet.ncols)
    flag = 0
    info =['<?xml version="1.0" encoding="UTF-8"?>']
    info.append("<DynamicPageInfo>")
    for row in rownum:
        print "line %s sheet content %s"% (row,sheet.cell(row,0).value)
        if sheet.cell(row,0).value != '':
            info.append('<pageEntry> ')
            info.append('  <key>%s</key>'%sheet.cell(row,0).value)
            info.append('  <pageInfo>')
        else:
            flag = row
        info.append("   <button>")
        for col in closnum:
            info.append('     <%s>%s</%s>'%(title[col],convertCode(sheet.cell(row,col).value,col),title[col]))
#            info.append('</'+title[col]+'>')
        info.append("  </button>")
        if flag == row :
           info.append('</pageInfo>')
           info.append('</pageEntry>')
    info.append("</DynamicPageInfo>")

    xmlstr =  ''.join(info)

    for line in info:
        print(line)

    wf = open('g:/%s.xml'%filename,'w')
    wf.write(xmlstr)

def convertCode(obj,col):
    if col in (4,5,7):
        if obj == 1:
            return 'True'
        else:
            return 'False'
    elif col == 2:
        return int(obj)
    else:
        return obj
generateButton("dynaButtons")


        



