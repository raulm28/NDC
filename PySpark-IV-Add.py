import pyspark
import sys
import os
import pymssql
import pandas as pd
from pyspark import SparkContext
from pyspark.sql import SQLContext
from pyspark.sql import SparkSession, Window
import pyspark.sql.functions as func
import matplotlib.pyplot as plt
from pyspark.sql.types import Row, StringType, LongType, DoubleType, StructField, StructType, TimestampType, IntegerType
from openpyxl import Workbook
from datetime import datetime

pd.set_option('display.max_rows',150000)
pd.set_option('display.max_columns',10)
pd.set_option('display.width', 100000)

os.environ['HADOOP_HOME'] = 'C:\Hadoop'
sys.path.append("C:\Hadoop\\bin")
#writer = ExcelWriter()
cnx = pymssql.connect(server='',user='',password='',database='')

cursor = cnx.cursor()
#f1 = open('G:\My Documents\SQL Server Management Studio\VerificationHeatMapping.sql','r')  
f1 = open('/Users/molinar1/Documents/RMQs/SQL Server Management Studio/SCM-LVPOrders.sql','r',encoding='utf-16')
query = pd.read_sql_query(f1.read(),cnx)
df = pd.DataFrame(query)#.sort_values(by=['PharmacyName','Unit','VerifyDate'])
#print(df)
sc = SparkContext()
sql = SQLContext(sc)
spark = SparkSession.builder.appName('Query').getOrCreate()

iva = sql.createDataFrame(df)
iva_pivot = iva.groupBy('VerifyDate','VerifyHour','MRN','VisitID','CurrentLocation','Service','OrderNum','Name','SummaryLine','PrimaryGI').pivot('add#').agg(func.first('Additives').alias('Additive'))
iva_data = iva_pivot.coalesce(1)
iva_data.toPandas().to_excel("iva_data.xlsx",sheet_name='Data',index=False)

