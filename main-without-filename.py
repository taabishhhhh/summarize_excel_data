
import pandas as pd
import os

# Paths to the folders
data_folder = 'data_folder'
result_folder = 'result_folder'

# List all Excel files in the data_folder
data_files = [f for f in os.listdir(data_folder) if f.endswith('.xlsx')]

# Initialize a list to store the data
data_to_write = []

# Loop through each Excel file in the data_folder
for file in data_files:
    # Construct the full path to the file
    file_path = os.path.join(data_folder, file)
    # Read the content of cell C9
    df = pd.read_excel(file_path, header=None, usecols="C", skiprows=8, nrows=1)
    # Append the content of cell C9 to the list
    data_to_write.append(df.iloc[0, 0])

# Assume there is only one Excel file in the result_folder and find its name
result_file_name = [f for f in os.listdir(result_folder) if f.endswith('.xlsx')][0]
result_file_path = os.path.join(result_folder, result_file_name)

# Convert the list to a DataFrame
df_to_write = pd.DataFrame(data_to_write, columns=['Content from C9'])

# Write the DataFrame to the Excel file in the result_folder
# If you want to add the data to an existing sheet without overwriting other data,
# you might need to read the sheet first, find the last row with data, and then append.
with pd.ExcelWriter(result_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df_to_write.to_excel(writer, index=False, startrow=0)

print('Data has been written to the result file.')
