import os
from pptx import Presentation
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import pandas as pd

# Folder containing PPTX files
FOLDER = 'Ch2'

# Get all pptx files
pptx_files = [f for f in os.listdir(FOLDER) if f.lower().endswith('.pptx')]
print(f"Found {len(pptx_files)} PPTX files: {pptx_files}")

# Function to extract all text from a pptx file
def extract_text_from_pptx(file_path):
    prs = Presentation(file_path)
    text = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text.append(shape.text)
    return '\n'.join(text)

# Extract text from all files
texts = []
for file in pptx_files:
    try:
        text = extract_text_from_pptx(os.path.join(FOLDER, file))
        texts.append(text)
        print(f"Extracted {len(text)} chars from {file}")
    except Exception as e:
        print(f"Error extracting from {file}: {e}")
        texts.append("")

print(f"Total texts: {len(texts)}")

if len(texts) > 1 and any(len(t) > 0 for t in texts):
    try:
        vectorizer = TfidfVectorizer().fit_transform(texts)
        similarity_matrix = cosine_similarity(vectorizer)
        similarity_df = pd.DataFrame(similarity_matrix, index=pptx_files, columns=pptx_files)
        similarity_df = (similarity_df * 100).round(2)
        similarity_df.to_excel('pptx_similarity.xlsx')
        similarity_df.to_csv('pptx_similarity.csv')
        print('Similarity comparison table saved as pptx_similarity.xlsx and pptx_similarity.csv')
    except Exception as e:
        print(f"Error in similarity calculation: {e}")
else:
    print('Not enough valid PPTX files to compare.')

