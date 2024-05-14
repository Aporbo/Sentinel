import pandas as pd
import os

# Define the directory containing your Excel files (replace with your actual path)
directory = "C:/Users/aporb/Desktop/New folder/files"

# Create an empty list to store the DataFrames
all_dataframes = []

# Loop through each Excel file in the directory
for filename in os.listdir(directory):
    if filename.endswith(".xlsx") or filename.endswith(".xls"):
        filepath = os.path.join(directory, filename)

        # Handle potential errors (file not found, etc.)
        try:
            # Read the Excel file into a DataFrame
            df = pd.read_excel(filepath)
            # Append the DataFrame to the list
            all_dataframes.append(df)
        except FileNotFoundError as e:
            print(f"Error: File not found - {filename}")

# Check if any DataFrames were successfully read
if not all_dataframes:
    print("Error: No Excel files found in the specified directory.")
else:
    # Concatenate all DataFrames vertically (stacking rows)
    try:
        merged_df = pd.concat(all_dataframes, ignore_index=True)  # Avoid duplicate indexing
    except ValueError as e:
        print(f"Error: Concatenation failed - {e}")
        # Handle the error based on the specific reason (e.g., empty list, non-iterable objects)
    else:
        # Define the output filename (replace with your desired name)
        output_filename = "merged_data.xlsx"

        # Save the merged DataFrame to a new Excel file
        merged_df.to_excel(output_filename, index=False)

        print(f"Excel files merged successfully! Output: {output_filename}")
