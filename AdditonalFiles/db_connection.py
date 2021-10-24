
import pyodbc as db

m_reporting = db.connect('DRIVER={SQL Server};'
                         'SERVER=13.76.190.123;'
                         'DATABASE=DCR_MREPORTING;'
                         'UID=sa;PWD=erp@123;')

azure = db.connect('DRIVER={SQL Server};'
                   'SERVER=137.116.139.217;'
                   'DATABASE=ARCHIVESKF;'
                   'UID=sa;PWD=erp@123;')