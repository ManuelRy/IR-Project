import os
from docx import Document
from textblob import TextBlob
import pandas as pd

# Function to analyze sentiment of a document
def analyze_sentiment(docx_path):
    doc = Document(docx_path)
    full_text = ""
    for paragraph in doc.paragraphs:
        full_text += paragraph.text
    blob = TextBlob(full_text)
    return blob.sentiment.polarity

# Function to map sentiment to emotion
def map_sentiment_to_emotion(sentiment):
    if sentiment < -0.16:
        return "Desperate"
    elif -0.16 <= sentiment < -0.12:
        return "Miserable"
    elif -0.12 <= sentiment < -0.08:
        return "Depressed"
    elif -0.08 <= sentiment < -0.04:
        return "Sad"
    elif -0.04 <= sentiment < -0.00:
        return "Displeased"
    elif 0.00 <= sentiment < 0.04:
        return "Neutral"
    elif 0.04 <= sentiment < 0.08:
        return "Satisfied"
    elif 0.08 <= sentiment < 0.12:
        return "Content"
    elif 0.12 <= sentiment < 0.16:
        return "Happy"
    else:
        return "Joyful"

# Folder paths
input_folder = "Mutation of the Apocalypse"
output_folder = "Statistics"
output_file = os.path.join(output_folder, "chapter_sentiment.xlsx")

# Analyze sentiment and map to emotion for each chapter
chapter_sentiments = []
for filename in os.listdir(input_folder):
    if filename.endswith(".docx"):
        chapter_number = int(filename.split("_")[1])
        docx_path = os.path.join(input_folder, filename)
        sentiment = analyze_sentiment(docx_path)
        emotion = map_sentiment_to_emotion(sentiment)
        chapter_sentiments.append((chapter_number, sentiment, emotion))

# Create DataFrame
df = pd.DataFrame(chapter_sentiments, columns=["Chapter", "Sentiment", "Feeling"])

# Sort DataFrame by Chapter Number
df = df.sort_values(by="Chapter")

# Check if the output file already exists
if os.path.exists(output_file):
    os.remove(output_file)  # Remove the existing file

# Export DataFrame to Excel
df.to_excel(output_file, index=False)

print("Chapter sentiment analysis completed and exported.")
