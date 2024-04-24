import os
import pandas as pd
from collections import Counter
from docx import Document
import re

# Function to count word frequency in a document
def count_word_frequency(docx_path):
    doc = Document(docx_path)
    word_frequency = Counter()
    for paragraph in doc.paragraphs:
        # Remove symbols and split paragraph into words
        words = re.findall(r'\b\w+\b', re.sub(r'[^\w\s]', '', paragraph.text.lower()))
        word_frequency.update(words)
    return word_frequency

# Function to export word frequency to an Excel file with adjusted column widths
def export_word_frequency_to_excel(word_frequency, output_filename):
    # Create DataFrame from word frequency dictionary
    df = pd.DataFrame(list(word_frequency.items()), columns=["Word", "Frequency"])
    # Create ExcelWriter object
    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        # Write DataFrame to Excel file
        df.to_excel(writer, index=False)
        # Get workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        # Adjust column widths based on the length of the longest word
        max_word_length = max(df['Word'].apply(len))
        worksheet.column_dimensions['A'].width = max_word_length * 1.2  # Adjust width of the 'Word' column

# Folder paths
input_folder = "Mutation of the Apocalypse"
output_folder = "Word Frequency"  # Adjusted output folder name

# Create output folder if it doesn't exist
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Process each .docx file in the input folder
for filename in os.listdir(input_folder):
    if filename.endswith(".docx"):
        docx_path = os.path.join(input_folder, filename)
        word_frequency = count_word_frequency(docx_path)
        output_filename = os.path.join(output_folder, f"{filename.split('.')[0]}_word_frequency.xlsx")
        
        # Export word frequency to Excel file with adjusted column widths
        export_word_frequency_to_excel(word_frequency, output_filename)

print("Individual word frequency analysis completed.")
