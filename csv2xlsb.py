#USAGE: python csv2xlsb.py C:\wamp\www\python\excel\test.csv
import win32com.client
import argparse
parser = argparse.ArgumentParser()
parser.add_argument("csvfile", help="full path of CSV file to convert")
args = parser.parse_args()
excel = win32com.client.Dispatch("Excel.Application")
doc = excel.Workbooks.Open(args.csvfile)
doc.SaveAs( args.csvfile.replace('csv','xlsb'), 50 )
doc.Close(False)
