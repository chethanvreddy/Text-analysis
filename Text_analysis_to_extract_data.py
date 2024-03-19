import pandas as pd
import requests
from bs4 import BeautifulSoup

# Load the input.xlsx file containing URLs and URL_IDs
df = pd.read_excel('Input.xlsx')

# Function to extract article text from URL
def extract_article_text(url):
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # Find article title
        article_title = soup.find('title').text.strip()
        
        # Find article text
        article_text = ''
        for paragraph in soup.find_all('p'):
            article_text += paragraph.text.strip() + '\n'
        
        return article_title, article_text
    except Exception as e:
        print(f"Error extracting article from {url}: {e}")
        return None, None

# Iterate over each row in the DataFrame and extract article text
for index, row in df.iterrows():
    url_id = row['URL_ID']
    url = row['URL']
    
    article_title, article_text = extract_article_text(url)
    
    if article_title and article_text:
        # Save the extracted article text to a text file
        with open(f"{url_id}.txt", 'w', encoding='utf-8') as file:
            file.write(article_title + '\n\n')
            file.write(article_text)

print("Extraction completed and files saved successfully.")
