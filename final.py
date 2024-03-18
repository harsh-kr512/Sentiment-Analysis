import string
import os
import pandas as pd

def count_syllables(word):
    vowels = 'aeiouy'
    count = 0
    previous_was_vowel = False
    for char in word:
        if char.lower() in vowels:
            if not previous_was_vowel:
                count += 1
            previous_was_vowel = True
        else:
            previous_was_vowel = False
    if word.endswith(('es', 'ed')):
        count -= 1
    if count == 0:
        count = 1 
    return count

def calculate_avg_syllables_per_word(text):
    words = text.split()
    total_syllables = sum(count_syllables(word) for word in words)
    total_words = len(words)
    avg_syllables_per_word = total_syllables / total_words if total_words != 0 else 0
    return avg_syllables_per_word

def calculate_avg_sentence_length(s_text):
    total_chars = sum(len(sentence.strip()) for sentence in s_text) 
    total_sentences = len(s_text) 
    avg_sentence_length = total_chars / total_sentences
    return avg_sentence_length


def find_complex_words(f_text, threshold):
    complex_words = [word for word in f_text if count_syllables(word) >= threshold]
    return complex_words

# Function to count personal pronouns
def count_personal_pronouns(text):
    pronouns = ['I', 'we', 'my', 'ours', 'us']
    counts = {pronoun: text.lower().count(pronoun) for pronoun in pronouns}
    if 'us' in counts:
        words = text.split()
        if 'us' in words:
            index_us = words.index('us')
            if index_us == len(words) - 1 or not words[index_us + 1].isalpha():
                counts.pop('us')
    return counts

# Read stop words from the folder
def read_stopwords_from_folder(folder_path, stopwords):
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        # Read stop words from the file
        with open(file_path, 'r') as file:
            stopwords.update(line.strip() for line in file)

# Function to analyze text files
def analyze_text_files(folder_path):
    stopwords = set()
    read_stopwords_from_folder('StopWords', stopwords)
    positive_word = open('positive-words.txt').read()
    negative_word = open('negative-words.txt').read()

    results = []

    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        with open(file_path, 'r', encoding='utf-8') as file:
            text = file.read()
            l_text = text.lower()
            c_text = l_text.translate(str.maketrans('', '', string.punctuation))
            t_text = c_text.split()
            f_text = [word for word in t_text if word not in stopwords]

            p_word = sum(1 for word in f_text if word in positive_word)
            n_word = sum(1 for word in f_text if word in negative_word)

            len_f_text = len(f_text)
            p_word=0
            n_word=0
            for word in f_text:
                if word in positive_word:
                    p_word+=1

            for word in f_text:
                if word in negative_word:
                    n_word+=1
            n_word=(-1)*n_word
            polarity_Score = (p_word - n_word) / ((p_word + n_word) + 0.000001)
            subjectivity_Score = (p_word + n_word) / ((len_f_text) + 0.000001)

            word_lenght = [len(word) for word in f_text]
            average_length = (sum(word_lenght) / len(word_lenght))

            threshold = 2
            complex_words = find_complex_words(f_text, threshold)
            p_complex_word = len(complex_words) / len(f_text)
            fog_index = 0.4 * (average_length + p_complex_word)

            avg_syllable = calculate_avg_syllables_per_word(text)

            # Count personal pronouns
            personal_pronouns = count_personal_pronouns(text)

            # Calculate average word length
            total_word_chars = sum(len(word) for word in f_text)
            avg_word_length = total_word_chars / len(f_text)

            # Append results to the list
            results.append({
                'File': filename,
                'Positive Score': p_word,
                'Negative Word': n_word,
                'Polarity Score': polarity_Score,
                'Subjectivity Score': subjectivity_Score,
                'Avg. Sentence Length': average_length,
                'Percentage of Complex words': p_complex_word,
                'Fog Index': fog_index,
                'Complex Word Count': len(complex_words),
                'Word count': len(f_text),
                'Syllable per Word': avg_syllable,
                'Personal Pronouns': personal_pronouns,
                'Average Word Length': avg_word_length
            })

    # Convert the list of dictionaries to a DataFrame
    df = pd.DataFrame(results)

    # Save the DataFrame to an Excel file
    df.to_excel('output_results.xlsx', index=False)

# Call the function to analyze text files in a given folder
analyze_text_files('text_file')  