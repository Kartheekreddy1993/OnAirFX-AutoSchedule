import os
import pyodbc
from datetime import datetime, timedelta

# Establish a connection
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\karth\Documents\OnAirFx1.mdb;'
    r'PWD=Neuro2010J138;'
)

conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

# Execute a SELECT query to fetch specific columns from tblTrackInfo where TypeID = 4
cursor.execute("SELECT tblRow, TypeID, GpID FROM tblTrackInfo WHERE TypeID = ?", 4)

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
    end_date = datetime.strptime(enddate, "%d-%m-%Y")
except ValueError:
    print("Invalid date format. Please use the format '10 Dec 2023'.")
    exit()
start_slots = ["21:00:00", "09:00:00", "12:00:00", "15:00:00", "18:00:00"]
# Print the stored list
print(len(result_list))
print(row_count)
while value < len(result_list):
    slot_count = 0
    for item in start_slots:
        if value >= len(result_list):
            break
        row_count += 1
        #print(item)
        formatted_date = input_date.strftime("%m-%d-%Y")
        if result_list[value][0] is not None:
            movie_rowid = int(result_list[value][0])
            movie_id =  int(result_list[value][2])
            prg_type = int(result_list[value][1])
            print(formatted_date,item)
            sql_query = "INSERT INTO tblTrackSchedule VALUES (?, ?, ?, #{}#, #{}#, ?, ?, #{}#, 0, 0, ?, 0, ?, 0, ?)".format(formatted_date, end_date, formatted_date)
            #sql_query = "INSERT INTO tblTrackSchedule VALUES (?, ?, ?, #{}, #{}, ?, ?, {}, 0, 0, ?, 0, ?, 0, ?)".format(formatted_date, end_date, empty)

            #print(sql_query)
            cursor.execute(sql_query, (row_count, prg_type, movie_id, item, item, slot_count, movie_rowid, False))
            slot_count += 1
            value += 1
    input_date += timedelta(days=1)
#print(result_list[2])
# Close the cursor and connection
cursor.close()
conn.commit()
conn.close()
