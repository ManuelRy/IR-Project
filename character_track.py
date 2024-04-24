import os
import pandas as pd
from collections import Counter
from docx import Document
import re
from nltk.corpus import stopwords

# Download NLTK stopwords if not already downloaded
import nltk
nltk.download('stopwords')

# Function to extract names from a document
def extract_names(docx_path):
    doc = Document(docx_path)
    names = []
    stop_words = set(stopwords.words('english'))  # Set of English stopwords
    for paragraph in doc.paragraphs:
        # Use regular expression to find names excluding personal pronouns and stop words
        found_names = re.findall(r'\b(?:[A-Z][a-z]+)\b', paragraph.text)
        found_names = [name for name in found_names if name.lower() not in stop_words]
        names.extend(found_names)
    return names

# Function to extract chapter name from filename
def extract_chapter_name(filename):
    # Extract chapter name from the filename
    chapter_info = filename.split("_")[2:-1]
    chapter_name = "_".join(chapter_info).replace(".docx", "")
    return chapter_name

# Folder paths
input_folder = "Mutation of the Apocalypse"
output_folder = "Statistics"
output_file = "character_track.xlsx"
output_path = os.path.join(output_folder, output_file)

# Create output folder if it doesn't exist
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Create an empty DataFrame to store character appearances
character_appearances = pd.DataFrame(columns=["Chapter Number", "Chapter Name", "Name"])

# Process each Word document in the input folder
for filename in os.listdir(input_folder):
    if filename.endswith(".docx"):
        # Extract chapter number and name from the filename
        chapter_number = int(filename.split("_")[1])
        chapter_name = extract_chapter_name(filename)

        # Extract names from the document
        docx_path = os.path.join(input_folder, filename)
        names = extract_names(docx_path)

        # Count the frequency of each name within this document
        name_frequency = Counter(names)

        # Create DataFrame from the name frequency
        df = pd.DataFrame(name_frequency.items(), columns=["Name", "Frequency"])

        # Add Chapter Number and Chapter Name columns
        df["Chapter Number"] = chapter_number
        df["Chapter Name"] = chapter_name

        # Append the DataFrame to character_appearances
        character_appearances = pd.concat([character_appearances, df], ignore_index=True)

# Check if the output file exists
if os.path.exists(output_path):
    # Replace the existing file
    os.remove(output_path)

# Export character appearances to Excel
character_appearances.to_excel(output_path, index=False)
