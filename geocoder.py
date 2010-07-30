import xlrd
import urllib2
import sys
import urllib 
import httplib
import time
import csv
import psycopg2


def getLocation(url,rowid=1):
 r = urllib.urlopen(url)
 lst = [] 
 if(r.getcode()==200):
  csvResponse = r.read()
  lst = csvResponse.split(',')
 return lst

def makeRequest(strRequest):
  urlencode = lambda s: urllib.urlencode({'x': s})[2:]
  return urlencode(strRequest.encode('utf-8'))


wb = xlrd.open_workbook('diff.xls')
sh = wb.sheet_by_index(0)
conn = psycopg2.connect("dbname=dmytro_atms user=dgolub")

for rownum in range(2500): #range(sh.nrows):
 #print sh.row_values(rownum)
 print rownum
 strR = "%s,%s"%(sh.cell(rownum,2).value,sh.cell(rownum,3).value)
 url = makeRequest(strR)
 strRequest = "http://maps.google.com/maps/geo?q=%s&output=csv"%(url)
 #print strRequest
 lst = getLocation(strRequest)
 if lst and int(lst[0]) == 200:
  # Connect to an existing database
  # Open a cursor to perform database operations
  cur = conn.cursor()
  # Pass data to fill a query placeholders and let Psycopg perform
  # the correct conversion (no more SQL injections!)
  cur.execute("""INSERT INTO atms (bankname,city,streetname,workinghours,additionalinfo,phonenumber)  VALUES (%s, %s,%s, %s,%s, %s) returning atm_id;""",(sh.cell(rownum,0).value,sh.cell(rownum,2).value,sh.cell(rownum,3).value,sh.cell(rownum,4).value,sh.cell(rownum,5).value,sh.cell(rownum,1).value,))
  atm_id = cur.fetchone()
 # 50.4500000,30.5233333  
  cur.execute("""INSERT INTO atms_position (atm_id,geom) values(%s, ST_GeomFromText('POINT(%s %s)',4326))""",(atm_id[0],float(lst[3]),float(lst[2]),))
  time.sleep(2)
conn.commit()

