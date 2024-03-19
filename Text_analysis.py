import os
import pandas as pd
import nltk
from nltk.tokenize import word_tokenize, sent_tokenize
from nltk.corpus import cmudict

# Download necessary resources
nltk.download('punkt')
nltk.download('cmudict')

# Function to read stopwords from files in a folder
def read_stopwords_from_folder(folder_path):
    stopwords = []
    for root, _, files in os.walk(folder_path):
        for file_name in files:
            file_path = os.path.join(root, file_name)
            with open(file_path, 'r') as file:
                stopwords.extend(file.read().splitlines())
    return stopwords

# Function to read positive and negative words from files
def read_words_from_file(file_path):
    with open(file_path, 'r') as file:
        words = file.readlines()
    return [word.strip() for word in words]

# Function to clean text using stop words list
def clean_text(text, stopwords):
    words = word_tokenize(text.lower())
    cleaned_words = [word for word in words if word.isalnum() and word not in stopwords]
    return cleaned_words

# Function to count syllables in a word
def count_syllables(word):
    d = cmudict.dict()
    if word.lower() in d:
        return max([len(list(y for y in x if y[-1].isdigit())) for x in d[word.lower()]])
    else:
        return 0

# Function to compute variables for the text
def compute_variables(text):
    cleaned_words = clean_text(text, stopwords)
    sentences = sent_tokenize(text)
    total_words = len(cleaned_words)
    total_sentences = len(sentences)
    
    # Count positive and negative words
    positive_score = sum(1 for word in cleaned_words if word in positive_words)
    negative_score = sum(1 for word in cleaned_words if word in negative_words)
    
    # Polarity score
    polarity_score = (positive_score - negative_score) / (positive_score + negative_score + 0.000001)
    
    # Subjectivity score
    subjectivity_score = (positive_score + negative_score) / (total_words + 0.000001)
    
    # Average sentence length
    avg_sentence_length = total_words / total_sentences
    
    # Percentage of complex words
    complex_words = [word for word in cleaned_words if count_syllables(word) > 2]
    percentage_complex_words = (len(complex_words) / total_words) * 100
    
    # Fog Index
    fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)
    
    # Average number of words per sentence
    avg_words_per_sentence = total_words / total_sentences
    
    # Complex word count
    complex_word_count = len(complex_words)
    
    # Word count
    word_count = total_words
    
    # Syllable per word
    syllable_per_word = sum(count_syllables(word) for word in cleaned_words) / total_words
    
    # Personal pronouns count
    personal_pronouns_count = sum(1 for word in cleaned_words if word.lower() in ['i', 'me', 'my', 'mine', 'myself', 'we', 'us', 'our', 'ours', 'ourselves'])
    
    # Average word length
    avg_word_length = sum(len(word) for word in cleaned_words) / total_words
    
    return [positive_score, negative_score, polarity_score, subjectivity_score, avg_sentence_length, percentage_complex_words, fog_index, avg_words_per_sentence, complex_word_count, word_count, syllable_per_word, personal_pronouns_count, avg_word_length]

# Read positive and negative words
positive_words = read_words_from_file('MasterDictionary/positive-words.txt')
negative_words = read_words_from_file('MasterDictionary/negative-words.txt')

# Read stopwords
stopwords_folder_path = 'StopWords'
stopwords = read_stopwords_from_folder(stopwords_folder_path)

# Path to the folder containing extracted article text files
folder_path = 'extract_article_text'

# Initialize results list
results_list = []

# Iterate through each file in the folder
for file_name in os.listdir(folder_path):
    file_path = os.path.join(folder_path, file_name)
    
    # Read the text from the file
    with open(file_path, 'r', encoding='utf-8') as file:
        text = file.read()
    
    # Compute variables for the text
    results = compute_variables(text)
    
    # Append results to the results list
    results_list.append(results)

# Read the existing Excel file
excel_file_path = 'Output_text_analysis.xlsx'
df = pd.read_excel(excel_file_path)

# Add the results to the DataFrame starting from the 3rd column
for i, results in enumerate(results_list, start=1):
    df = df.append(pd.DataFrame([results], columns=df.columns[2:]), ignore_index=True)

# Save the DataFrame back to the Excel file starting from the 2nd row
df.to_excel(excel_file_path, index=False, startrow=1)

print("done saving")
