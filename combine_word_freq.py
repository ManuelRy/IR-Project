import os
import pandas as pd

# Function to combine all individual word frequency Excel files into one combined Excel file
def combine_word_frequency_excel(input_folder, output_filename):
    all_dfs = []
    for filename in os.listdir(input_folder):
        if filename.endswith(".xlsx"):
            excel_path = os.path.join(input_folder, filename)
            try:
                df = pd.read_excel(excel_path)
                all_dfs.append(df)
            except Exception as e:
                print(f"Error reading file {excel_path}: {e}")
    combined_df = pd.concat(all_dfs)
    
    # Combine word counts by summing up frequencies for each word
    combined_df = combined_df.groupby('Word', as_index=False)['Frequency'].sum()
    
    # Check if the output file already exists, replace if it does
    if os.path.exists(output_filename):
        os.remove(output_filename)
    
    # Export the combined DataFrame to Excel
    combined_df.to_excel(output_filename, index=False)

# Folder paths
input_folder = "Word Frequency"
output_folder = "Statistics"  # Adjusted output folder name
combined_output_filename = os.path.join(output_folder, "combined_word_frequency.xlsx")

# Create output folder if it doesn't exist
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Combine all individual word frequency Excel files into one combined Excel file
combine_word_frequency_excel(input_folder, combined_output_filename)

print("Combined word frequency analysis completed and exported.")
