# file to fix problem with bank name

import xlrd
import psycopg2

wb = xlrd.open_workbook('atms2.xls')
sh = wb.sheet_by_index(0)
conn = psycopg2.connect("dbname=dmytro_atms user=dgolub")
cur = conn.cursor()

for rownum in range(sh.nrows):
 bankName = "%s"%sh.cell(rownum,0).value
 searchQuery = "UPDATE atms SET bankname='%s' Where city='%s' and streetname='%s'"%(bankName,sh.cell(rownum,2).value,sh.cell(rownum,3).value)
 cur.execute(searchQuery)
conn.commit()




