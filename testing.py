import pandas as pd
from sklearn.metrics import accuracy_score, classification_report
from textblob import TextBlob  # For sentiment analysis
import matplotlib.pyplot as plt


data = pd.read_csv("C:/Users/admin/OneDrive/Documents/VIT/Capstone/testing.csv")

speech_accuracy = accuracy_score(data['recognized_speech'], data['ground_truth'])
intent_accuracy = accuracy_score(data['recognized_intent'], data['expected_intent'])
intent_accuracy_expected = accuracy_score(data['recognized_intent'], data['intent_based_on_recognized_speech'])
intent_report = classification_report(data['expected_intent'], data['recognized_intent'])
common_speech_errors = data[data['recognized_speech'] != data['ground_truth']]['recognized_speech'].value_counts()


print("Speech Recognition Accuracy:", speech_accuracy)
print("Intent Recognition Accuracy:\nOn the basis of expected and recognized intent :", intent_accuracy)
print("Intent Recognition Accuracy:\nOn the basis of recognized speech intent and recognized intent :", intent_accuracy_expected)
print("Intent Recognition Report:\n", intent_report)
print("Common Speech Recognition Errors:\n", common_speech_errors)

import json
with open('C:/Users/admin/OneDrive/Documents/VIT/Capstone/new_intents.json') as file:
    data = json.load(file)
num_intents = len(data['intents'])
num_patterns = sum(len(intent['patterns']) for intent in data['intents'])
print(num_intents)
print(num_patterns)