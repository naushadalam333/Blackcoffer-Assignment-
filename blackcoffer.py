import nltk
from nltk.tokenize import word_tokenize, sent_tokenize
from nltk.corpus import stopwords
from nltk.probability import FreqDist
from nltk.sentiment import SentimentIntensityAnalyzer
from nltk import pos_tag
from textblob import TextBlob
from textstat import syllable_count
from goose3 import Goose
import os
import string
import pandas as pd

# Download necessary NLTK resources
nltk.download('punkt')
nltk.download('stopwords')

def read_stopwords_from_file(file_path):
    with open(file_path, 'r') as file:
        return file.read().splitlines()

def create_word_dictionary(file_path):
    with open(file_path, 'r') as file:
        return {word.lower(): True for word in file.read().splitlines()}

def read_urls_from_excel(file_path, sheet_name='Sheet1', column_name='URL'):
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        return df[column_name].tolist()
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return []

# Specify the path to the Excel file containing URLs
excel_file_path = r"C:\Users\lenovo\Downloads\Input.xlsx"

# Check if the file exists before proceeding
if os.path.exists(excel_file_path):
    # Read URLs from the Excel file
    urls = read_urls_from_excel(excel_file_path)

    # Create an empty list to store results
    results = []

    # Loop through each URL
    for url in urls:
        try:
            # Extracting text from the article
            g = Goose()
            article = g.extract(url=url)

            # Check if the article is not empty
            if not article.cleaned_text:
                print(f"Empty article for URL: {url}")
                continue

            # Read stopwords from a folder
            stopwords_folder_path = r"C:\Users\lenovo\Downloads\StopWords-20231212T124238Z-001\StopWords"
            all_custom_stopwords = [word.lower() for file_path in [os.path.join(stopwords_folder_path, file) for file in os.listdir(stopwords_folder_path) if file.endswith('.txt')] for word in read_stopwords_from_file(file_path)]

            # Remove custom stopwords from the text
            filtered_words = [word.lower() for word in word_tokenize(article.cleaned_text) if word.lower() not in all_custom_stopwords and word.isalnum()]
            filtered_text = ' '.join(filtered_words)

            # Display the result
            print(f"\nAnalysis for URL: {url}")
            print("Original Text:")
            print(article.cleaned_text)
            print("\nText without Custom Stopwords:")
            print(filtered_text)

            # Read positive and negative words from files and create dictionaries
            positive_words_dict = create_word_dictionary(r"C:\Users\lenovo\Downloads\MasterDictionary-20231212T124239Z-001\MasterDictionary\positive-words.txt")
            negative_words_dict = create_word_dictionary(r"C:\Users\lenovo\Downloads\MasterDictionary-20231212T124239Z-001\MasterDictionary\negative-words.txt")

            # Sentiment analysis
            sia = SentimentIntensityAnalyzer()
            sentiment_score = sia.polarity_scores(article.cleaned_text)

            # Count positive and negative words
            positive_count = sum(1 for word in filtered_words if word in positive_words_dict)
            negative_count = sum(1 for word in filtered_words if word in negative_words_dict)

            # Display the result
            print("Sentiment Analysis Score:")
            print(sentiment_score)
            print("\nPositive Word Count:", positive_count)
            print("Negative Word Count:", negative_count)

            # Frequency distribution of words
            fdist = FreqDist(filtered_words)
            print("Word Frequency Distribution:")
            print(fdist.most_common(10))

            # Additional analysis using TextBlob, textstat, and NLTK functions
            blob = TextBlob(article.cleaned_text)
            subjectivity_score = blob.sentiment.subjectivity

            sentences = sent_tokenize(article.cleaned_text)
            sentence_lengths = [len(word_tokenize(sentence)) for sentence in sentences]
            avg_sentence_length = sum(sentence_lengths) / len(sentence_lengths)

            words = word_tokenize(article.cleaned_text)
            complex_word_count = sum(1 for word in words if syllable_count(word) > 2)
            total_word_count = len(words)
            percentage_complex_words = (complex_word_count / total_word_count) * 100

            fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)

            sentences = sent_tokenize(article.cleaned_text)
            words_per_sentence = [len(word_tokenize(sentence)) for sentence in sentences]
            avg_words_per_sentence = sum(words_per_sentence) / len(words_per_sentence)

            complex_words = [word for word in words if syllable_count(word) > 2]
            complex_word_count = len(complex_words)

            word_count = len(words)

            total_syllables = sum(syllable_count(word) for word in words)
            average_syllables_per_word = total_syllables / total_word_count

            pos_tags = pos_tag(words)
            personal_pronoun_tags = ['PRP', 'PRP$', 'WP', 'WP$']
            personal_pronoun_count = sum(1 for word, pos in pos_tags if pos in personal_pronoun_tags)

            total_characters = sum(len(word) for word in words)
            average_word_length = total_characters / total_word_count

            # Append the results to the list
            results.append({
                "URL": url,
                "Title": article.title,
                "Sentiment Score": sentiment_score["compound"],
                "Positive Word Count": positive_count,
                "Negative Word Count": negative_count,
                "Top Words and Frequencies": fdist.most_common(10),
                "Subjectivity Score": subjectivity_score,
                "Average Sentence Length": avg_sentence_length,
                "Percentage of Complex Words": percentage_complex_words,
                "Fog Index": fog_index,
                "Word count": total_word_count,
                "Complex word count": complex_word_count,
                "Average Number of Words Per Sentence": avg_words_per_sentence
                # Add more fields as needed
            })

        except Exception as e:
            print(f"Error processing URL {url}: {e}")

    # Specify the Excel file path
    output_excel_file = r"C:\Users\lenovo\OneDrive\Desktop\resultbook.xlsx"

    # Write the DataFrame to an Excel file
    results_df = pd.DataFrame(results)
    results_df.to_excel(output_excel_file, index=False)

    print(f"Results have been exported to '{output_excel_file}'.")
else:
    print(f"Excel file '{excel_file_path}' not found.")