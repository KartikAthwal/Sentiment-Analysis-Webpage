# %%
import pandas as pd
import os
import requests
from bs4 import BeautifulSoup
import nltk
from nltk.tokenize import sent_tokenize, word_tokenize
from nltk.corpus import stopwords as nltk_stopwords
import re
nltk.download('stopwords')
nltk.download('punkt')

# %%
# Read the Excel file and print first 5 entries
df = pd.read_excel("Input.xlsx")
df.head()

# %%
# Set the StopWords path and make a set for StopWord
folder_path=r"StopWords"
stopwords = set()
# Loop over each text file and add the stop words to the stopwords set
for file_name in os.listdir(folder_path):
    file_path = os.path.join(folder_path, file_name)
    with open(file_path, 'r') as file:
        stopwords.update(file.read().split())

# %%
#Make a master Dictionary with Positive and Negative words 
Master_dict = {'Positive': set(), 'Negative': set()}

#set the files
positive_words_file = "positive-words.txt"
negative_words_file = "negative-words.txt"

#loop over each file and update the Master_dict
for sentiment, word_file in [('Positive', positive_words_file), ('Negative', negative_words_file)]:
    with open(word_file, 'r') as file:
        Master_dict[sentiment].update(word for word in file.read().split() if word not in stopwords)

# %%
#Sentiment Analysis, formulas used from Test_analysis file
def sentiment_analysis(words_cleaned, Master_dict):
    
    positive_score = sum(1 for word in words_cleaned if word in Master_dict['Positive'])
    negative_score = sum(1 for word in words_cleaned if word in Master_dict['Negative'])
    polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    subjectivity_score = (positive_score + negative_score) / (len(words_cleaned) + 0.000001)
    return(positive_score, negative_score,polarity_score, subjectivity_score)

# %%
#Function to count Syllable as per provided constraints

def syllable_count(word):
    word = word.lower()
    count = 0
    vowels = "aeiouy"
    if word[0] in vowels:
        count += 1
    for index in range(1, len(word)):
        if word[index] in vowels and word[index - 1] not in vowels:
            count += 1
    if word.endswith("es") or word.endswith("ed"):
        # Avoid miscounting syllables
        if len(word) > 2 and word[-3] not in vowels:
            count -= 1
    if count == 0:
        count += 1
    return count

# %%
#to store results
results=[]


page_not_found=[] #to store entries where wew encouncter errors

#looping over each entry in Input.xlsx
for index, row in df.iterrows():
    url_id = row['URL_ID']
    url = row['URL']
    try:
        response = requests.get(url)
        # Check if the page exists
        if response.status_code == 404:
            print(f"Page not found for URL: {url}")
        response = requests.get(url)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')

        #to get title we are going to assume it uses the html <H1> tag 
        title_tag = soup.find('h1')
        title = title_tag.get_text() if title_tag else "No Title"

        #to get the Article we are going to assume it uses the html <p> tag
        article_tags = soup.find_all('p')
        article_text = ' '.join([p.get_text() for p in article_tags]) if article_tags else "No Content"
        
        #Saving the file with its USER_ID as its name
        with open(f'{url_id}.txt', 'w', encoding='utf-8') as file:
            file.write(title + "\n" + article_text)
            
        #tokenizing and cleaning the words
        words = word_tokenize(article_text)
        words_cleaned = [word.lower() for word in words if word.lower() not in stopwords and word.isalpha()]

        #doing sentiment analysis
        positive_score, negative_score,polarity_score, subjectivity_score = sentiment_analysis(words_cleaned, Master_dict)
        
        #calculating other variables
        sentences= sent_tokenize(article_text)
        avg_sentence_length = len(words_cleaned) / len(sentences)

        complex_words_count=sum(1 for word in words_cleaned if syllable_count(word) > 2)

        percentage_complex_words = complex_words_count / len(words_cleaned)

        fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)

        avg_words_per_sentence = len(words_cleaned) / len(sentences)

        
        #Using nltk stopwords module for word count
        nltk_stopwords_set = set(nltk_stopwords.words('english'))
        words_without_nltk_stopwords = [word for word in words_cleaned if word not in nltk_stopwords_set]

        
        #Using Regex Module to remove punctuations
        words_without_punctuation = [re.sub(r'[^\w\s]', '', word) for word in words_without_nltk_stopwords]
        words_without_punctuation = [word for word in words_without_punctuation if word]

        word_count = len(words_without_punctuation)

        syllables_per_word = sum(syllable_count(word) for word in words_without_punctuation) / word_count

        personal_pronouns_count = len(re.findall(r'\bI\b|\bwe\b|\bmy\b|\bours\b|\bus\b', article_text, re.IGNORECASE))

        avg_word_length = sum(len(word) for word in words_cleaned) / len(words_cleaned)

        results.append({
            'URL_ID': url_id,
            'URL': url,
            'POSITIVE SCORE': positive_score,
            'NEGATIVE SCORE': negative_score,
            'POLARITY SCORE': polarity_score,
            'SUBJECTIVITY SCORE': subjectivity_score,
            'AVG SENTENCE LENGTH': avg_sentence_length,
            'PERCENTAGE OF COMPLEX WORDS': percentage_complex_words,
            'FOG INDEX': fog_index,
            'AVG NUMBER OF WORDS PER SENTENCE': avg_words_per_sentence,
            'COMPLEX WORD COUNT': complex_words_count,
            'WORD COUNT': word_count,
            'SYLLABLE PER WORD': syllables_per_word,
            'PERSONAL PRONOUNS': personal_pronouns_count,
            'AVG WORD LENGTH': avg_word_length
        })
    except requests.exceptions.HTTPError as e:
        page_not_found.append(url_id)
         
         
# Convert results to DataFrame and save as Excel
results_df = pd.DataFrame(results)
results_df.to_excel('Output.xlsx', index=False)

print(f' Page not found for URL_ID: {page_not_found}')
print("The output file has been made")
        


# %%


# %%
