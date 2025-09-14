import os
import pyodbc
from datetime import datetime, timedelta

# Establish a connection
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Tools\ONAIR AUto\OnAirFx1.mdb;'
    r'PWD=Neuro2010J138;'
)

conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

# Execute a SELECT query to fetch specific columns from tblTrackInfo where TypeID = 4
cursor.execute("SELECT tblRow, TypeID, GpID, FilePath FROM tblTrackInfo WHERE TypeID = ? AND FilePath LIKE ?", (4, '%2025%'))

rows = cursor.fetchall()


# Print results
#for row in rows:
    #print(f"tblRow: {row.tblRow}, TypeID: {row.TypeID}, GpID: {row.GpID}, FilePath: {row.FilePath}")

# Fetch the results and store them in a list with GpID values without .0
result_list = [(row.tblRow, row.TypeID, int(row.GpID)) for row in cursor.fetchall()]
query = f"SELECT MAX(tblRow) FROM tblTrackSchedule;"
cursor.execute(query)
row_count = cursor.fetchone()[0]
enddate = '02-02-2222'
value = 0
date_input = input("Enter a start date (e.g., 22-01-2024) ")
# Convert the user input to a datetime object
try:
    input_date = datetime.strptime(date_input, "%d-%m-%Y")
    end_date = datetime.strptime(enddate, "%m-%d-%Y")
except ValueError:
    print("Invalid date format. Please use the format '10 Dec 2023'.")
    exit()
    
start_slot = "01:00:00"
# Print the stored list
print(len(result_list))
print(row_count)
slot_count = 5

for row in rows:
    formatted_date = input_date.strftime("%m-%d-%Y")
    row_count += 1
    movie_rowid = int(row.tblRow)
    movie_id =  int(row.GpID)
    prg_type = int(row.TypeID)
    print(movie_id)
    sql_query = "INSERT INTO tblTrackSchedule VALUES (?, ?, ?, #{}#, #{}#, ?, ?, ?, 0, 0, ?, 0, ?, 0, ?)".format(formatted_date, end_date)
    print(sql_query)
    cursor.execute(sql_query, (row_count, prg_type, movie_id, start_slot, start_slot, None, slot_count, movie_rowid, False))
    input_date += timedelta(days=1)
    
    
cursor.close()
conn.commit()
conn.close()