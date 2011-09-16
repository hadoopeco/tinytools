__author__ = 'Mark.Wei'

from xlrd import open_workbook



def generateButton(filename=None):
    wb = open_workbook("g:/%s.xls" % filename,'rb')

    sheet = wb.sheet_by_index(0)
    title = sheet.row_values(0)

    print title

    rownum = range(1,sheet.nrows)
    closnum = range(1,sheet.ncols)

    print "rows= %s , cols %s"%(rownum,closnum)
    flag = 0
    info =['<?xml version="1.0" encoding="UTF-8"?>']
    info.append("<DynamicPageInfo>")
    for row in rownum:
        if sheet.cell(row,0).value != '':
            info.append('<pageEntry> <key>')
            info.append(sheet.cell(row,0).value)
            info.append('</key>')
            info.append('<pageInfo>')
        else:
            flag = row
        info.append("<button>")
        for col in closnum:
            info.append('<'+title[col]+'>')
            info.append(str(sheet.cell(row,col).value))
            info.append('</'+title[col]+'>')
        info.append("</button>")
        if flag != row :
           info.append('</pageInfo>')
           info.append('</pageEntry>')
    info.append('</pageInfo>')
    info.append('</pageEntry>')
    info.append("</DynamicPageInfo>")
    xmlstr =  ''.join(info)
    
    wf = open('g:/%s.xml'%filename,'w')
    wf.write(xmlstr)


generateButton("dynaButtons")


        



