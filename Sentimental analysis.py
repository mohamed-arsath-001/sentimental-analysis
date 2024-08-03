import pandas as pd
from selectolax.parser import HTMLParser
import requests
import re
import os
from IPython.display import FileLink


import nltk
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize, sent_tokenize
from bs4 import BeautifulSoup, Comment

import warnings
warnings.filterwarnings("ignore")


nltk.download('punkt')
nltk.download('stopwords')

#==============================================================================================================================

def stop_words(file_path):
    word=[]
    with open(file_path, 'r',encoding='utf-8', errors='ignore') as file:
        lines = file.readlines()
        
    for line in lines:
        word.append(line.strip())
    return word
        
        
word_path = os.listdir(r'analy.1\\test_assign\\cleaning-words')   # folder directory
stop_word=[]
print("\n=================================================================\n")
for i in word_path:
    print(f"File loaded successfully  Name : {i}")
    stop_word.extend(stop_words(r"analy.1\\test_assign\\cleaning-words"+i))
    
    
#--------------------------------------------------------------------------------------------------------------------------------    

positive_word=[]
with open(r"analy.1\test_assign\masterdictionary\positive-words.txt", 'r',encoding='utf-8', errors='ignore') as file:
        lines = file.readlines()
        
for line in lines:
        positive_word.append(line.strip())
    
negative_word=[]
with open(r"analy.1\test_assign\masterdictionary\negative-words.txt", 'r',encoding='utf-8', errors='ignore') as file:
        lines = file.readlines()
        
for line in lines:
        negative_word.append(line.strip())
        
#--------------------------------------------------------------------------------------------------------------------------------------
# Function to count the number of syllables in a word
def count_syllables(word):
    word = word.lower()
    syllables = re.findall(r'[aeiouy]{1,2}', word)
    syllable_count = len(syllables)
    if word.endswith(('es', 'ed')) and not word.endswith(('les', 'ted', 'ded')):
        syllable_count -= 1
    return max(syllable_count, 1)  # Ensure at least one syllable

# Function to determine if a word is complex (3 or more syllables)
def is_complex(word):
    return count_syllables(word) >= 3

# Function to clean words (remove stopwords and punctuations)
def clean_words(words, stopword):
    stop_words = set(stopword)
    cleaned = [word for word in words if word.lower() not in stop_words and word.isalpha()]
    return cleaned

def sentiment_scores(words, positive_words, negative_words):
    positive_score = sum(1 for word in words if word in positive_words)
    negative_score = sum(1 for word in words if word in negative_words)
    polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    subjectivity_score = (positive_score + negative_score) / (len(words) + 0.000001)
    return positive_score, negative_score, polarity_score, subjectivity_score


# Function to calculate the Gunning Fox Index
def gunning_fox_index(text, stopword):
    # Split text into sentences
    sentences = sent_tokenize(text)
    total_sentences = len(sentences)
    
    # Split text into words
    words = word_tokenize(text)
    cleaned_words = clean_words(words, stopword)
    total_words = len(cleaned_words)
    
    # Count complex words
    complex_words = [word for word in cleaned_words if is_complex(word)]
    total_complex_words = len(complex_words)
    
    # Calculate metrics
    avg_sentence_length = total_words / total_sentences if total_sentences > 0 else 0
    percentage_complex_words = total_complex_words / total_words if total_words > 0 else 0
    fog_index = 0.4 * (avg_sentence_length + 100 * percentage_complex_words)
    
    return {
        'total_words': total_words,
        'total_sentences': total_sentences,
        'total_complex_words': total_complex_words,
        'avg_sentence_length': avg_sentence_length,
        'percentage_complex_words': percentage_complex_words,
        'fog_index': fog_index
    }

# Function to count personal pronouns
def count_personal_pronouns(text):
    pronouns = re.findall(r'\b(I|we|my|ours|us)\b', text, re.I)
    return len(pronouns)

def average_syllables_per_word(paragraph):
    
    words = re.findall(r'\b\w+\b', paragraph) 
    total_syllables = sum(count_syllables(word) for word in words)
    total_words = len(words)
    average_syllables = total_syllables / total_words if total_words > 0 else 0
    return average_syllables

# Function to calculate average word length
def average_word_length(words):
    total_characters = sum(len(word) for word in words)
    avg_word_length = total_characters / len(words) if words else 0
    return avg_word_length



def full_analysis(text, stopword, positive_words, negative_words):

    # Tokenize words
    
    
    words = word_tokenize(text)
    cleaned_words = clean_words(words, stopword)
    
    # Perform readability analysis
    readability = gunning_fox_index(text, stopword)
    
    # Perform sentiment analysis
    positive_score, negative_score, polarity_score, subjectivity_score = sentiment_scores(cleaned_words, positive_words, negative_words)
    
    # Count personal pronouns
    personal_pronouns_count = count_personal_pronouns(text)
    
    # Calculate average word length
    avg_word_length = average_word_length(cleaned_words)
    
    syllable_per_word = average_syllables_per_word(text)
    
    result= {
        "Url_id":'',
        "POSITIVE SCORE": positive_score,
        "NEGATIVE SCORE": negative_score,
        "POLARITY SCORE": round(polarity_score, 2),
        "SUBJECTIVITY SCORE": round(subjectivity_score, 2),
        "AVG SENTENCE LENGTH": round(readability["avg_sentence_length"], 2),
        "PERCENTAGE OF COMPLEX WORDS": round(readability["percentage_complex_words"] * 100, 2),
        "FOG INDEX": round(readability["fog_index"], 2),
        "AVG NUMBER OF WORDS PER SENTENCE": round(readability["avg_sentence_length"], 2),
        "COMPLEX WORD COUNT": readability["total_complex_words"],
        "WORD COUNT": readability["total_words"],
        "SYLLABLE PER WORD": round(syllable_per_word,1),
        "PERSONAL PRONOUNS": round(personal_pronouns_count,1),
        "AVG WORD LENGTH": round(avg_word_length,1)
            }
    
    
    return result




def read_xl():
    # Define the path to the XLSX file
    file_path = 'analy.1\Input.xlsx'

    # Read the XLSX file into a DataFrame
    df = pd.read_excel(file_path)

    # Create a dictionary to store the URL_ID and URL
    url_dict = {}

    # Iterate through the rows of the DataFrame
    for index, row in df.iterrows():
        url_id = row['URL_ID']
        url = row['URL']
        url_dict[url_id] = url

    return url_dict

def scrape_n_store_txt(url_dict):
    
    try:
        # Adjust sheet name as necessary
        output_data = pd.DataFrame(columns=["Url_id", "POSITIVE SCORE", "NEGATIVE SCORE", "POLARITY SCORE",
                                    "SUBJECTIVITY SCORE", "AVG SENTENCE LENGTH", "PERCENTAGE OF COMPLEX WORDS",
                                    "FOG INDEX", "AVG NUMBER OF WORDS PER SENTENCE", "COMPLEX WORD COUNT",
                                    "WORD COUNT", "SYLLABLE PER WORD", "PERSONAL PRONOUNS", "AVG WORD LENGTH"])
        
        skip=0
        i=1
        for url_id,url in url_dict.items():
            # Fetch the HTML content of the webpage
            
            
            print(f"\n{i} calulating requirements....")   
            i+=1
            response = requests.get(url)
        
            # if response.status_code != 200:
            # raise Exception(f"Failed to fetch the webpage. Status code: {response.status_code}")
        
            html_content = response.text
        
            # Parse the HTML content using Selectolax
            tree = HTMLParser(html_content)
        
            article_tags = tree.css('div.td-post-content.tagdiv-type')
            
            if article_tags == []:
                skip +=1
                continue
            
            l = []
            l.append(article_tags[0].text())
        
            # Replace single newlines with a space
            text = re.sub(r'\n', ' ', l[0])
            text = re.sub(r'\n\n\n', '    ', l[0])
        
        
            # print(text,"\n\n\n")
            
            # checking the artcile tag present or not
            if article_tags == []:
                pass 
            else:
                data = article_tags[0].text()
           
                val = full_analysis(data, stop_word, positive_word, negative_word)
            
                val["Url_id"] = url
                # Append a row for the analysis results
                output_data = pd.concat([output_data, pd.DataFrame([val])], ignore_index=True)
                
                
           
        print(f"\nProcess Done...{i-(skip+1)} websites data \nskipped{skip} website data  ")
        return  output_data

    except Exception as e:
        print(f"Error occurred: {str(e)}") 
        
def main():
    url_dict = read_xl()
    return scrape_n_store_txt(url_dict)
    
    


final_result = main()
final_result.to_excel('Output_data.xlsx', index=False)

    # Provide a link to download the file
FileLink(r'Output_data.xlsx')
