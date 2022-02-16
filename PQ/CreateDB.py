import pyodbc
import sqlalchemy as sq

conn = pyodbc.connect('Driver={SQL Server Native Client 10.0};'
                      'Server=GBWADJ8HJ293;'
                      'Database=ICP-OES;'
                      'Trusted_Connection=yes;')

cursor = conn.cursor()

cursor.execute('''

                 CREATE TABLE TM21
                 (
                 Solution_Label nvarchar (30),
                 Date date,
                 Al_TM21 decimal (6,3),
                 As_TM21 decimal (6,3),
                 Ca_TM21 decimal (6,3),
                 Cd_TM21 decimal (6,3),
                 Co_TM21 decimal (6,3),
                 Cr_TM21 decimal (6,3),
                 Cu_TM21 decimal (6,3),
                 Fe_TM21 decimal (6,3),
                 Mg_TM21 decimal (6,3),
                 Ni_TM21 decimal (6,3),
                 Pb_TM21 decimal (6,3),
                 Sb_TM21 decimal (6,3),
                 Ti_TM21 decimal (6,3),
                 V_TM21 decimal (6,3),
                 Zn_TM21 decimal (6,3),
                 Product nvarchar (50),
                 Batch nvarchar (50)
                 )
                 ''')

conn.commit()
