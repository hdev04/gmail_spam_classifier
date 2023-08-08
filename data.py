import pickle
import streamlit as st
import numpy as np
from win32com.client import Dispatch

def speak(text):
    speak = Dispatch("SAPI.spVoice")
    speak.Speak(text)

classifier = pickle.load(open("spam.pkl", "rb"))

def main():
    st.title("Gmail Spam Classifier")

    msg = st.text_input("Enter a text : ")
    if st.button ("Predict"):
        word_dict_pickle = open("word_pickle.pkl", "rb")

        word_dict_pickle.close()
        sample = []

        for i in word_dict:
            sample.append(msg.split(" ").count(i[0]))

        sample = np.array(sample)
        pred = classifier.predict(sample.reshape(1, 3000))
        result = pred[0]

        if result == 1:
            st.error("This is Spam")
            speak("This is Spam")
        else:
            st.success("This is Ham")
            speak("This is Ham")
main()