# import pandas as pd
# import os

# # Directory containing your Excel files
# directory = r'C:\Users\Lenovo\Documents\Mbessa Pride\Member Data'

# # List to store extracted data
# extracted_data = []

# # Loop through all Excel files in the directory
# for filename in os.listdir(directory):
#     if filename.endswith(".xlsx"):
#         file_path = os.path.join(directory, filename)
#         # Read the Excel file
#         df = pd.read_excel(file_path, engine='openpyxl', header=None)
#         # Extract data for the specified date
#         filtered_data = df[df.iloc[:, 0] == '18/8/2023']
#         if not filtered_data.empty:
#             # Extract balance from column E
#             balance = filtered_data.iloc[0, 4] if len(filtered_data.columns) > 4 else None
#             # Add data to the list
#             extracted_data.append((filename, balance))

# # Create a DataFrame from the extracted data
# result_df = pd.DataFrame(extracted_data, columns=['File', 'Balance'])

# # Save the result to a new Excel file
# result_df.to_excel('extracted_data.xlsx', index=False)
# import pandas as pd
# import os

# # Directory containing your Excel files
# directory = r'C:\Users\Lenovo\Documents\Mbessa Pride\Member Data'

# # List to store extracted data
# extracted_data = []

# # Loop through all Excel files in the directory
# for filename in os.listdir(directory):
#     if filename.endswith(".xlsx"):
#         file_path = os.path.join(directory, filename)
#         # Read the Excel file without header
#         df = pd.read_excel(file_path, engine='openpyxl', header=None)
#         # Iterate through column A and find rows with the specified date
#         for index, row in df.iterrows():
#             if row[0] == '18/8/2023':
#                 # Get the balance from the last column
#                 balance = row.iloc[-1]
#                 extracted_data.append((filename, balance))
#                 break  # Break the loop if date is found to avoid duplicates

# # Create a DataFrame from the extracted data
# result_df = pd.DataFrame(extracted_data, columns=['File', 'Balance'])

# # Save the result to a new Excel file
# result_df.to_excel('extracted_data.xlsx', index=False)

# import pandas as pd
# import os

# # Directory containing your Excel files
# directory = r'C:\Users\Lenovo\Documents\Mbessa Pride\Member Data'

# # List to store extracted data
# extracted_data = []

# # Counter to keep track of processed sheets
# processed_sheets = 0

# # Loop through all Excel files in the directory
# for filename in os.listdir(directory):
#     if filename.endswith(".xlsx") and not filename.startswith("~$"):
#         file_path = os.path.join(directory, filename)
#         # Read the Excel file without header
#         df = pd.read_excel(file_path, engine='openpyxl', header=None)
#         # Iterate through column A and find rows with the specified date
#         for index, row in df.iterrows():
#             if row[0] == '18/8/2023':
#                 # Print the matching rows for verification
#                 print(f'Filename: {filename}, Row Index: {index}, Data: {row.tolist()}')
#                 # Increment the processed_sheets counter
#                 processed_sheets += 1
#                 break  # Break the loop if date is found to avoid duplicates

# # Display the number of sheets processed
# print(f'Number of sheets processed: {processed_sheets}')

# import pandas as pd
# import os

# # Directory containing your Excel files
# directory = r'C:\Users\Lenovo\Documents\Mbessa Pride\Member Data'

# # List to store extracted data
# extracted_data = []

# # Counter to keep track of processed sheets
# processed_sheets = 0

# # Get a sorted list of Excel files based on the numeric prefix in the file names
# excel_files = sorted([filename for filename in os.listdir(directory) if filename.endswith(".xlsx") and not filename.startswith("~$")],
#                      key=lambda x: int(x.split('.')[0]))

# # Loop through sorted Excel files
# for filename in excel_files:
#     file_path = os.path.join(directory, filename)
#     # Read the Excel file without header
#     df = pd.read_excel(file_path, engine='openpyxl', header=None)
#     # Iterate through column A and find rows with the specified date
#     for index, row in df.iterrows():
#         if row[0] == '18/8/2023':
#             # Print the matching rows for verification
#             print(f'Filename: {filename}, Row Index: {index}, Data: {row.tolist()}')
#             # Increment the processed_sheets counter
#             processed_sheets += 1
#             break  # Break the loop if date is found to avoid duplicates

# # Display the number of sheets processed
# print(f'Number of sheets processed: {processed_sheets}')

# import pandas as pd
# import os

# # Directory containing your Excel files
# directory = r'C:\Users\Lenovo\Documents\Mbessa Pride\Member Data'

# # List to store extracted data
# extracted_data = []

# # Get a sorted list of Excel files based on the numeric prefix in the file names
# excel_files = sorted([filename for filename in os.listdir(directory) if filename.endswith(".xlsx") and not filename.startswith("~$")],
#                      key=lambda x: int(x.split('.')[0]))

# # Counter to keep track of processed sheets
# processed_sheets = 0

# # Loop through sorted Excel files
# for filename in excel_files:
#     file_path = os.path.join(directory, filename)
#     # Read the Excel file without header
#     df = pd.read_excel(file_path, engine='openpyxl', header=None)
#     # Iterate through column A and find rows with the specified date
#     for index, row in df.iterrows():
#         if row[0] == '18/8/2023':
#             # Add data to the extracted_data list
#             extracted_data.append((filename, row.iloc[-1]))
#             break  # Break the loop if date is found to avoid duplicates
#     # Increment the processed_sheets counter
#     processed_sheets += 1
# # Create a DataFrame from the extracted data
# result_df = pd.DataFrame(extracted_data, columns=['File', 'Row Index', 'Data'])

# # Save the result to a new Excel file in the order 1-390
# result_df.to_excel('extracted_data_sorted.xlsx', index=False)

# # Display the number of sheets processed
# print(f'Number of sheets processed: {processed_sheets}')

# import pandas as pd
# import os

# # Directory containing your Excel files
# directory = r'C:\Users\Lenovo\Documents\Mbessa Pride\Member Data'

# # List to store extracted data
# extracted_data = []

# # Get a sorted list of Excel files based on the numeric prefix in the file names
# excel_files = sorted([filename for filename in os.listdir(directory) if filename.endswith(".xlsx") and not filename.startswith("~$")],
#                      key=lambda x: int(x.split('.')[0]))

# # Counter to keep track of processed sheets
# processed_sheets = 0

# # Loop through sorted Excel files
# for filename in excel_files:
#     file_path = os.path.join(directory, filename)
#     # Read the Excel file without header
#     df = pd.read_excel(file_path, engine='openpyxl', header=None)
#     # Iterate through column A and find rows with the specified date
#     for index, row in df.iterrows():
#         if row[0] == '18/8/2023':
#             # Add data to the extracted_data list as a dictionary
#             extracted_data.append({'File': filename, 'Row Index': index, 'Data': row.tolist()})
#             break  # Break the loop if date is found to avoid duplicates
#     # Increment the processed_sheets counter
#     processed_sheets += 1
# # Create a DataFrame from the extracted data
# result_df = pd.DataFrame(extracted_data)

# # Display the number of sheets processed
# print(f'Number of sheets processed: {processed_sheets}')

# # Save the DataFrame to a new Excel file
# result_df.to_excel('extracted_data_1.xlsx', index=False) #all data in one cell

# import pandas as pd
# import os

# # Directory containing your Excel files
# directory = r'C:\Users\Lenovo\Documents\Mbessa Pride\Member Data'

# # List to store extracted data
# extracted_data = []

# # Get a sorted list of Excel files based on the numeric prefix in the file names
# excel_files = sorted([filename for filename in os.listdir(directory) if filename.endswith(".xlsx") and not filename.startswith("~$")],
#                      key=lambda x: int(x.split('.')[0]))

# # Counter to keep track of processed sheets
# processed_sheets = 0

# # Loop through sorted Excel files
# for filename in excel_files:
#     file_path = os.path.join(directory, filename)
#     # Read the Excel file without header
#     df = pd.read_excel(file_path, engine='openpyxl', header=None)
#     # Iterate through column A and find rows with the specified date
#     for index, row in df.iterrows():
#         if row[0] == '18/8/2023':
#             # Create a dictionary with keys as column names and values as row elements
#             data_dict = {'Column_' + str(i + 1): row[i] for i in range(len(row))}
#             # Add data to the extracted_data list as a dictionary
#             extracted_data.append({'File': filename, 'Row Index': index, 'Data': data_dict})
#             break  # Break the loop if date is found to avoid duplicates
#     # Increment the processed_sheets counter
#     processed_sheets += 1
# # Create a DataFrame from the extracted data
# result_df = pd.DataFrame(extracted_data)

# # Display the number of sheets processed
# print(f'Number of sheets processed: {processed_sheets}')

# # Save the DataFrame to a new Excel file
# result_df.to_excel('extracted_data_2.xlsx', index=False) #still data in one cell

# import pandas as pd
# import os

# # Directory containing your Excel files
# directory = r'C:\Users\Lenovo\Documents\Mbessa Pride\Member Data'

# # List to store extracted data
# extracted_data = []

# # Get a sorted list of Excel files based on the numeric prefix in the file names
# excel_files = sorted([filename for filename in os.listdir(directory) if filename.endswith(".xlsx") and not filename.startswith("~$")],
#                      key=lambda x: int(x.split('.')[0]))

# # Counter to keep track of processed sheets
# processed_sheets = 0

# # Loop through sorted Excel files
# for filename in excel_files:
#     file_path = os.path.join(directory, filename)
#     # Read the Excel file without header
#     df = pd.read_excel(file_path, engine='openpyxl', header=None)
#     # Iterate through column A and find rows with the specified date
#     for index, row in df.iterrows():
#         if row[0] == '18/8/2023':
#             # Create a dictionary with keys as column names and values as row elements
#             data_dict = {'Column_' + str(i + 1): row[i] for i in range(len(row))}
#             # Add data to the extracted_data list as a dictionary
#             extracted_data.append({'File': filename, 'Date': row[0], 'Balance': row[4]})
#             break  # Break the loop if date is found to avoid duplicates
#     # Increment the processed_sheets counter
#     processed_sheets += 1
# # Create a DataFrame from the extracted data
# result_df = pd.DataFrame(extracted_data)

# # Display the number of sheets processed
# print(f'Number of sheets processed: {processed_sheets}')

# # Save the DataFrame to a new Excel file
# result_df.to_excel('extracted_data_4.xlsx', index=False)

# import pandas as pd
# import os

# # Directory containing your Excel files
# directory = r'C:\Users\Lenovo\Documents\Mbessa Pride\Member Data'

# # List of target dates and their corresponding columns (0-based index)
# target_dates = {'09/01/2023': 3, '15/9/2023': 3, '10/06/2023': 3, '18/8/2023': 4}

# # List to store extracted data
# extracted_data = []

# # Get a sorted list of Excel files based on the numeric prefix in the file names
# excel_files = sorted([filename for filename in os.listdir(directory) if filename.endswith(".xlsx") and not filename.startswith("~$")],
#                      key=lambda x: int(x.split('.')[0]))

# # Loop through sorted Excel files
# for filename in excel_files:
#     file_path = os.path.join(directory, filename)
#     # Read the Excel file without header
#     df = pd.read_excel(file_path, engine='openpyxl', header=None)
#     # Iterate through column A and find rows with the specified dates
#     # for index, row in df.iterrows():
#     #     date = row[0]
#     #     if date in target_dates:
#     #         # Extract data from the specified column based on the date
#     #         column_index = target_dates[date]
#     #         data_value = row[column_index]
#     #         # Add data to the extracted_data list as a dictionary
#     #         extracted_data.append({'File': filename, 'Date': date, 'Data': data_value})
#     for index, row in df.iterrows():
#         date = str(row[0])  # Convert the value to string, handling NaN automatically
#         date = date.strip()  # Remove leading and trailing spaces from the date
#         if date in target_dates:
#             # Extract data from the specified column based on the date
#             column_index = target_dates[date]
#             data_value = row[column_index]
#             # Add data to the extracted_data list as a dictionary
#             extracted_data.append({'File': filename, 'Date': date, 'Data': data_value})

# # Create a DataFrame from the extracted data
# result_df = pd.DataFrame(extracted_data)

# # Save the DataFrame to a new Excel file
# result_df.to_excel('extracted_data_7.xlsx', index=False)

import pandas as pd
import os

# Directory containing your Excel files
directory = r'C:\Users\Lenovo\Documents\Mbessa Pride\Member Data'

# List of target dates and their corresponding columns (0-based index)
target_dates = {'09/01/2023': 3, '15/9/2023': 3, '10/06/2023': 3, '18/8/2023': 4}

# List to store extracted data
extracted_data = []

# Get a sorted list of Excel files based on the numeric prefix in the file names
excel_files = sorted([filename for filename in os.listdir(directory) if filename.endswith(".xlsx") and not filename.startswith("~$")],
                     key=lambda x: int(x.split('.')[0]))

# Loop through sorted Excel files
for filename in excel_files:
    file_path = os.path.join(directory, filename)
    # Read the Excel file without header
    df = pd.read_excel(file_path, engine='openpyxl', header=None)
    # Iterate through column A and find rows with the specified dates
    for index, row in df.iterrows():
        date = str(row[0]).strip()  # Convert the value to string, handling NaN automatically
        try:
            # Convert date to datetime object
            date = pd.to_datetime(date, format='%d/%m/%Y').strftime('%d/%m/%Y')
        except ValueError:
            # Handle invalid date formats or NaN values
            continue
        if date in target_dates:
            # Extract data from the specified column based on the date
            column_index = target_dates[date]
            data_value = row[column_index]
            # Add data to the extracted_data list as a dictionary
            extracted_data.append({'File': filename, 'Date': date, 'Data': data_value})

# Create a DataFrame from the extracted data
result_df = pd.DataFrame(extracted_data)

# Save the DataFrame to a new Excel file
result_df.to_excel('extracted_data_8.xlsx', index=False)


