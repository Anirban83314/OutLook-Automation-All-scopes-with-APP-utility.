#!/usr/bin/env python
import os
import sys
import pyodbc

con = pyodbc.connect('Trusted_Connection=yes', driver = '{SQL Server}', server = '##DB_Name/other_details', database = 'DB SUB version operation')
cur = con.cursor()

querystring = "Query to be executed here. "

cur.execute(querystring)
con.commit()
