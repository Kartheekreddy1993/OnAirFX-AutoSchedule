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

# Read notepad file
with open("C:\\Tools\\ONAIR AUto\\list.txt", "r") as f:   # <-- change path to your notepad file
    config_lines = [line.strip() for line in f if line.strip()]

# Get the current max tblRow
cursor.execute("SELECT MAX(tblRow) FROM tblTrackSchedule;")
row_count = cursor.fetchone()[0] or 0  # handle None if table is empty

enddate = '02-02-2222'
end_date = datetime.strptime(enddate, "%m-%d-%Y")

for line in config_lines:
    filepath_filter, date_str, start_slot, slot_count = line.split(",")
    date_input = datetime.strptime(date_str.strip(), "%d-%m-%Y")
    slot_count = int(slot_count.strip())

    # Fetch rows matching the FilePath filter
    cursor.execute(
        "SELECT tblRow, TypeID, GpID, FilePath "
        "FROM tblTrackInfo WHERE TypeID = ? AND FilePath LIKE ? ORDER BY tblRow",
        (4, f"%{filepath_filter.strip()}%")
    )
    rows = cursor.fetchall()

    print(f"\nProcessing filter {filepath_filter} -> Found {len(rows)} rows")

    for row in rows:
        formatted_date = date_input.strftime("%m-%d-%Y")
        row_count += 1
        movie_rowid = int(row.tblRow)
        movie_id = int(row.GpID)
        prg_type = int(row.TypeID)

        sql_query = (
            "INSERT INTO tblTrackSchedule VALUES "
            "(?, ?, ?, #{}#, #{}#, ?, ?, ?, 0, 0, ?, 0, ?, 0, ?)"
            .format(formatted_date, end_date.strftime("%m-%d-%Y"))
        )

        print(sql_query)  # debug
        cursor.execute(
            sql_query,
            (row_count, prg_type, movie_id, start_slot, start_slot, None,
             slot_count, movie_rowid, False)
        )

        # Increment date by 1 day for next row
        date_input += timedelta(days=1)

# Finalize
conn.commit()
cursor.close()
conn.close()
print("✅ Import completed")
