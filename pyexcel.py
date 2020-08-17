import pandas as pd
import pyodbc

class Report:
    def __init__(self, startrow, startcol, name, sql_query):
        self.startrow = startrow
        self.startcol = startcol
        self.name = name
        self.sql_query = sql_query

def print_start_creating(filename):
    print("Creating " + filename + " ...")

def print_done_creating(filename):
    print(filename + " created")

def create_excel_from_sql(report, conn):
    result = pd.read_sql_query(report.sql_query, conn)
    max_rows = 30000
    rows = len(result.index)
    num_max_files =int(rows / max_rows)

    print("Number of rows exceeds the max of " + max_rows.__str__() + "\n Splitting into " + (num_max_files+1).__str__() + " files")

    if rows > max_rows:
        
        for i in range(1, num_max_files+1):
            filename = report.name + i.__str__() + ".xlsx"
            print_start_creating(filename)

            if i == 1:
                df = result.iloc[:max_rows, :]
            else:
                start = (i-1)*max_rows
                end = i*max_rows
                df = result.iloc[start:end, :]

            
            df.to_excel(filename, index=False, startrow=report.startrow, startcol=report.startcol)
            print_done_creating(filename)

            if i == num_max_files:
                filename = report.name + (i+1).__str__() + ".xlsx"
                print_start_creating(filename)
                start = i*max_rows
                df = result.iloc[start:, :]
                df.to_excel(filename, index=False, startrow=report.startrow, startcol=report.startcol)
                print_done_creating(filename)
        
    else:
        filename =  report.name + ".xlsx"
        print_start_creating(filename)
        result.to_excel(filename, index=False, startrow=report.startrow, startcol=report.startcol)
        print_done_creating(filename)


conn_str = (
    r'Driver={SQL Server};'
    r'Server=localhost;'
    r'Database=AdventureWorksDW2017;'
    r'Trusted_Connection=yes;'
    )

conn = pyodbc.connect(conn_str)

reports = []

reports.append(Report(7,0,'FactResellerSales', 'SELECT * FROM [AdventureWorksDW2017].[dbo].[FactResellerSales]'))
reports.append(Report(3,0,'FactInternetSalesReason', 'SELECT * FROM [AdventureWorksDW2017].[dbo].[FactInternetSalesReason]'))

for report in reports:
    message = "Executing query for " + report.name + " , it might take a few minutes"
    print(message)
    create_excel_from_sql(report, conn)


