# file to fix problem with bank name

import xlrd
import psycopg2
import codecs

wb = xlrd.open_workbook('atms2.xls')
sh = wb.sheet_by_index(0)
conn = psycopg2.connect("dbname=dmytro_atms user=dgolub")
cur = conn.cursor()

for rownum in range(sh.nrows):
 searchQuery = "SELECT * from atms where city='%s' and streetname='%s' and bankname='%s'"%(sh.cell(rownum,2).value,sh.cell(rownum,3).value,sh.cell(rownum,0).value)
 cur.execute(searchQuery)
 rows = cur.fetchall()
 f = open('/home/dgolub/Downloads/atms/atms_dima/diff2.txt', 'a+')
 f.write( codecs.BOM_UTF8 )
 if len(rows) == 0:
  logStr = "%s;%s;%s;%s;%s;%s\n"%(sh.cell(rownum,0).value,sh.cell(rownum,1).value,sh.cell(rownum,2).value,sh.cell(rownum,3).value,sh.cell(rownum,4).value,sh.cell(rownum,5).value)
  f.write(logStr.encode( "utf-8" ))
f.close()


