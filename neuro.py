import os
from datetime import datetime, timedelta
import pandas as pd

value = 0

# Load the Excel file into a DataFrame
file_path = r'C:\Users\karth\Documents\MovieDetail.xlsx'  # Replace with the actual path to your Excel file
df = pd.read_excel(file_path)

# Extract the "moviepath" column
movie_paths = df['moviepath']

# Display the extracted column
# print(movie_paths)
extracted_text = []
extracted_text1 = []

file_data = []

for path in movie_paths:
    # Find the last two backslashes and extract text in between
    first_backslash = path.rfind('\\')
    second_backslash = path.rfind('\\', 0, first_backslash)

    if first_backslash != -1 and second_backslash != -1:
        text_between = path[second_backslash + 1:first_backslash]
        extracted_text1.append(text_between)

# print(extracted_text)
extracted_text = sorted(extracted_text1)
item_count = len(extracted_text)
print("Number of items in the list:", item_count)
start_slots = ["09:00:00 PM", "06:00:00 AM", "09:00:00 AM", "12:00:00 PM", "03:00:00 PM", "06:00:00 PM"]
date_input = input("Enter a start date (e.g., 22/01/2024): ")

try:
    input_date = datetime.strptime(date_input, '%d/%m/%Y')
except ValueError:
    print("Invalid date format. Please use the format '22/09/2023'.")
    exit()

print("Successfully parsed date:", input_date)

while value < len(extracted_text):
    for item in start_slots:
        if value >= len(extracted_text):
            break

        movie = extracted_text[value]
        if extracted_text[value] is not None:
            file_data.append({
                'moviedate': input_date.strftime("%d/%m/%Y"),
                'movietime': item,  # Use 'item' instead of 'start_slots'
                'moviename': movie,
                'ScrollAds': 'PlayAll',  # Use 'PlayAll' instead of 'scroll'
                'ScrollAds2': 'PlayAll',  # Use 'PlayAll' instead of 'scroll'
                'SponsorGroup': 'Default',  # Use 'Default' instead of 'sponsor'
                'ChLogo': 'Logo1',
            })

            value += 1
    input_date += timedelta(days=1)

# Create a DataFrame from the collected data
result_df = pd.DataFrame(file_data)

result_file_name = "MovieScheduleDetail.xlsx"
result_df.to_excel(os.path.join(os.path.dirname(__file__), result_file_name), index=False)

print(f"Result saved to {result_file_name}")
