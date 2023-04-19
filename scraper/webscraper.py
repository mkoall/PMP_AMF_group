#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Created on Thu Apr 13 18:49:10 2023

Webscrape articles on currencies used in carry trades in 
order to get an overview of the news situation.

@author: Martin 'Dashing' Koall
@co-author: Amando 'Handsome' Zanderigo
"""

# IMPORTS
import pprint
import requests
import matplotlib.pyplot as plt
import pandas as pd
# import os
import re
import datetime
import nltk
import smtplib  # To send the emails.
# import json
# import csv

# FROMS
from nltk.corpus import stopwords
from wordcloud import WordCloud
from collections import Counter
from datetime import time
from nltk.tokenize import word_tokenize
from email.message import EmailMessage  # Email modules needed.

nltk.download('stopwords')


class ScrapeAndScore:
    date = ''
    queries = []
    transcript = {}
    endpoint = 'https://newsapi.org/v2/everything?'
    sentiment_words = {}
    scores = {}

    def __init__(self, queries, date):
        """
        ScrapeAndScore class initialiser.

        INPUT:
        - queries: a list of queries to search for.
        - date: a date to run the queries over. 
        """
        self.queries = queries
        self.date = date

    def scrape(self, secret) -> dict:
        """ 
        Method to scrape the actual machine readable news. 
        """
        url = self.endpoint

        result = {}

        for q in self.queries:
            # Specify the query and number of returns
            parameters = {
                'q': q,             # query phrase
                'pageSize': 100,    # maximum is 100
                'apiKey': secret,   # your own API key
                'from': self.date
            }

            # Make the request
            response = requests.get(url, params=parameters)

            # Convert the response to
            # JSON format and pretty print it
            response_json = response.json()
            pprint.pprint(response_json)

            text_combined = ''

            for i in response_json['articles']:
                if i['description'] != None:
                    text_combined += i['description'] + ' '

            result[q] = {'currency': q, 'raw_text': text_combined}

        self.transcript = result
        # return result

    def clean_text(self):
        """
        Clean the transcript text of unnecessary things.
        """
        for title, content in self.transcript.items():
            print('Cleaning \t {}'.format(title))
            # Access raw text
            text = content['raw_text']
            # Convert to lowercase
            text = text.lower()

            # remove everything between long lines
            # #-------
            # #text
            # #-------
            pattern = r"^-{2,}$\n^.*$\n^-{2,}$"
            text = re.sub(pattern, '', text, flags=re.MULTILINE)

            # Remove unicode
            text = re.sub(
                r"(@\[A-Za-z0-9]+)|([^0-9A-Za-z \t])|(\w+:\/\/\S+)|^rt|http.+?", "", text)

            # remove stop words
            stop_words = stopwords.words('english')
            stopwords_dict = Counter(stop_words)
            text = ' '.join([word for word in text.split()
                            if word not in stopwords_dict])

            self.transcript[title]['text'] = text

        # return transcript

    def import_sentiment_words(self, filepath):
        """
        Imports a sentiment dictionary.
        e.g. Loughran and Macdonald financial dictionary. 

        INPUT:
            - filepath: string indicating the sentiment dictionary file location.
        """
        df = pd.read_csv(filepath, sep=',')

        tokeep = ['Word', 'Positive']
        positive = set([x['Word'].lower() for idx, x in df[tokeep].iterrows()
                        if x['Positive'] > 0])

        tokeep = ['Word', 'Negative']
        negative = set([x['Word'].lower() for idx, x in df[tokeep].iterrows()
                        if x['Negative'] > 0])

        tokeep = ['Word', 'Uncertainty']
        uncertainty = set([x['Word'].lower() for idx, x in df[tokeep].iterrows()
                           if x['Uncertainty'] > 0])

        # return {'positive': positive, 'negative': negative, 'uncertainty': uncertainty}
        self.sentiment_words = {'positive': positive,
                                'negative': negative, 'uncertainty': uncertainty}

    def get_scores(self):
        """
        Method to get the scores.
        """
        for title, content in self.transcript.items():

            print('Working on: \t {}'.format(title))
            self.scores[title] = {}

            # get the text for each transcript and split into a list of substrings
            words = content['text']
            words = words.split()

            # Total number of words (to normalize scores)
            totalwords = len(words)

            # score
            sentpos = len(
                [word for word in words if word in self.sentiment_words['positive']])
            sentneg = len(
                [word for word in words if word in self.sentiment_words['negative']])
            sentuncer = len(
                [word for word in words if word in self.sentiment_words['uncertainty']])

            # Collect and prepare for conditional scores
            self.scores[title] = {
                'Positive Sentiment': sentpos,
                'Negative Sentiment': sentneg,
                'Uncertain Sentiment': sentuncer,
                'Sentiment': sentpos - sentneg,
                'Total words': totalwords,
                'Currency': self.transcript[title]['currency'],
            }
        # return scores

    def write_excel(filename, sheetname, dataframe):
        """
        Method to write the output to an Excel sheet. 
        """
        with pd.ExcelWriter(filename, engine='openpyxl', mode='a') as writer:
            workBook = writer.book
            try:
                workBook.remove(workBook[sheetname])
            except:
                print("Worksheet does not exist")
            finally:
                dataframe.to_excel(writer, sheet_name=sheetname, index=False)
                writer.save()

    def collect_in_df(self):
        """
        Method to collect the results in Pandas dataframe.
        """
        scores_df = pd.DataFrame().from_dict(self.scores, orient='index')
        scores_df.index.name = 'Event Name'
        lst = scores_df.Sentiment.to_list()
        lst.insert(0, datetime.date.today())
        df = pd.DataFrame(list(zip(lst, list(initial_data.columns)))).T
        df = pd.DataFrame(df.iloc[0, :]).T
        df.columns = initial_data.columns
        test_data = pd.concat([initial_data, df], ignore_index=True)

        return test_data

    def write_to_txt(self):
        """
        Method to write results to .txt such that they can be emailed/used elsewhere.
        """
        df = self.collect_in_df()  # pd.DataFrame(self.scores)
        df.to_csv("results.txt")

    def email_results(self):
        """
        Method to send the results in an email to Handsome and Dashing. 
        """
        file = "results.txt"
        # Open the plain text file whose name is in textfile for reading.
        with open(file) as fp:
            # Create a text/plain message
            msg = EmailMessage()
            msg.set_content(fp.read())

        # me == the sender's email address
        # you == the recipient's email address
        msg['Subject'] = f'The contents of {file}'
        msg['From'] = "wucloud@localhost"
        msg['To'] = ["amando.zanderigo@wu.ac.at"]

        print(msg)

        # Send the message via our own SMTP server.
        # smtplib.SMTP(host='smtp.gmail.com', port=587)
        s = smtplib.SMTP("localhost")
        s.send_message(msg)
        s.quit()


if __name__ == "__main__":
    initial_data = pd.read_excel('currencies_score.xlsx', 'scores')
    queries = list(initial_data.columns)[1:]
    date = '2023-04-01'

    ScrapeAndScore = ScrapeAndScore(queries, date)

    ScrapeAndScore.import_sentiment_words(
        "./Loughran-McDonald_MasterDictionary_2021.csv")
    ScrapeAndScore.scrape(secret="bf270a2eb98e49a08c967e9fe95eb7c0")
    ScrapeAndScore.clean_text()
    ScrapeAndScore.get_scores()
    ScrapeAndScore.write_to_txt()
    ScrapeAndScore.email_results()

    # write_excel('currencies_score.xlsx', 'scores',
    #             ScrapeAndScore.collect_in_df())
