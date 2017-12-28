#!/usr/bin/env python
import os
import sys
import pyodbc

con = pyodbc.connect('Trusted_Connection=yes', driver = '{SQL Server}', server = 'AMSDC1-S-41720.europe.shell.com\fps,51001', database = 'FPS_US_WSP')
cur = con.cursor()

querystring = "SELECT TOP 5 * from int_fle_hst where ifh_fle_name like '%US_LVP_OPS_FPS%' order by 1 desc"

cur.execute(querystring)
con.commit()
