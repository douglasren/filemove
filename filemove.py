#!/usr/bin/python

'''
将文件A中的文件复制到文件B中
excel中是文件A要复制到文件B的时候文件名字的对应情况
'''
import os,sys
import re
import xlrd

g_destexcelfile = os.getcwd() + '/test.xls'
g_exceltabel = 'table1'
g_sourcefile = os.getcwd() + '/a'
g_destfile = os.getcwd() + '/b'

g_sourcelist = []

def getsourcefilelists() :
    global g_sourcelist
    g_sourcelist = os.listdir(g_sourcefile)

def open_excel(filename) :
    data = None
    try:
        data = xlrd.open_workbook(filename)
    except Exception, e:
        raise e
    return data

def movefile(src, dest) :

    for v in g_sourcelist :
        dex = v.rfind('.')
        if dex < 0 :
            continue
        tv = v[0 : dex]
        if tv == src :
            dest = dest + v[dex:]
            # linux cp, windows copy
            cmd = 'cp ' + g_sourcefile + '/' + v + '  ' + g_destfile + '/' + dest
            print cmd
            os.system(cmd)
            print 'copy ', v, ' finish'

def solve() :
    getsourcefilelists()
    data = open_excel(g_destexcelfile)
    if not data :
        print 'open excell error'
        return
    table = data.sheet_by_name(g_exceltabel)

    if not table :
        print 'get table error'
        return 
    rowlen = table.nrows
    print 'rowlen:', rowlen

    for i in range(rowlen) :
        if 0 == i :
            continue
        rowvals = table.row_values(i)
        if rowvals:
            movefile(rowvals[1], rowvals[2])


if __name__ == '__main__' :
    solve()
