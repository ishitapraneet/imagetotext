# -*- coding: utf-8 -*-
"""
Created on Sun Oct  3 04:38:42 2021

@author: Administrator
"""

#Importing necessary libraries
import tkinter as tk
from tkinter import *
from tkinter import filedialog
from tkinter.filedialog import askopenfilename
from tkinter import ttk
from tkinter import messagebox
import numpy as np
from PIL import Image
import pytesseract
import pandas as pd
import cv2
import  mysql.connector

global result


#Creating new root window and defining its size and other properties
root = tk.Tk()
root.config(bg="#1ABC9C")
Title = root.title("OCR GUI")
root.geometry('661x300')
root.resizable(0, 0)


#If user chooses 1st Option then this function will execute
def readfromimage1():
    path = PathTextBox.get('1.0', 'end-1c')
    pytesseract.pytesseract.tesseract_cmd = r'C:/Program Files (x86)/Tesseract-OCR/tesseract.exe'
    # reading image from given path
    im = cv2.imread(path)
    #Image Preprocessing
    # Resize the Image to it's double size using Inter_CUBIC interpolation(a bicubic interpolation over 4×4 pixel neighborhood).
    im = cv2.resize(im, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
    # Converting Input Image to Grayscale.(Pytesseract works in RGB and OpenCV's default is BGR)
    im = cv2.cvtColor(im, cv2.COLOR_BGR2GRAY)
    # Implementing Adaptive thresholding on grayscale Image
    # If pixel value is greater than a threshold value, it is assigned one value (may be white),
    # else it is assigned another value (may be black)
    im = cv2.adaptiveThreshold(im, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 199, 5)
    # A 3*3 kernal passed in 2D convolution to reduce noise and make edges sharp.
    # A LPF(Low Pass Filters) is used to blur or remove noise and HPF(High Pass Filters) is used to make edges sharp.
    kernel = np.array([[-1, -1, -1],
                       [-1, 9, -1],
                       [-1, -1, -1]])
    im= cv2.filter2D(im, -1, kernel)
    # Detect the words from the Image.
    text = pytesseract.image_to_string(im, lang='eng')
    # Reading the knowledge Base through pandas dataframe.
    excel_file = r'C:\Users\Administrator\Desktop\projects\imagetotext\knowledgebase2.xlsx'
    df = pd.read_excel(excel_file)
    # Open text file in read mode and store the text in that text file.
    with open(r'C:\Users\Administrator\Desktop\projects\imagetotext\output.txt', mode='w') as f:
        f.write(text.upper())
    frames = []
    # Go through each word in text file and match that word from knowledge base.
    with open(r"C:\Users\Administrator\Desktop\projects\imagetotext\output.txt", 'r') as file:
        for line in file:
            for word in line.split():
                # If match found , then add it to dataframe df1.
                df1 = df[df['WORDS'] == word]
                # Append dataframe each time to frames list.
                frames.append(df1)
    if frames:
        result = pd.concat(frames)
    else:
        print("No results")

    # On successfull completion A window opens to ask user where
    # to save output excel file
    export_file_path = filedialog.asksaveasfilename(defaultextension='.xlsx')
    result.to_excel(export_file_path, index=False, header=True)


#If user chooses second option this function will execute
def readfromimage2():
    path = PathTextBox.get('1.0', 'end-1c')
    pytesseract.pytesseract.tesseract_cmd = r'C:/Program Files (x86)/Tesseract-OCR/tesseract.exe'
    im = cv2.imread(path)
    im = cv2.resize(im, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
    im = cv2.cvtColor(im, cv2.COLOR_BGR2GRAY)  # Converting Image to Grayscale
    im = cv2.adaptiveThreshold(im, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 199,
                               5)  # Implementing Adaptive thresholding on grayscale Image
    kernel = np.array([[-1, -1, -1],
                       [-1, 9, -1],
                       [-1, -1, -1]])
    im = cv2.filter2D(im, -1, kernel)
    text = pytesseract.image_to_string(im, lang='eng')
    excel_file = r'C:\Users\Administrator\Desktop\projects\imagetotext\knowledgebase1.xlsx'
    df = pd.read_excel(excel_file)
    with open(r'C:\Users\Administrator\Desktop\projects\imagetotext\output2.txt', mode='w') as f:
        f.write(text.upper())
    frames = []
    with open(r'C:\Users\Administrator\Desktop\projects\imagetotext\output2.txt', 'r') as file:
        for line in file:
            for word in line.split():
                df1 = df[df['WORDS'] == word]
                frames.append(df1)
    if frames:
        result = pd.concat(frames)
    else:
        print("No matching Values")

    export_file_path = filedialog.asksaveasfilename(defaultextension='.xlsx')
    result.to_excel(export_file_path, index=False, header=True)


#Dropdown menu function to decide which function will execute when
def checkcmbo():

    if Operation.get() == "Get Synonyms and Antonyms":
        readfromimage1()

    elif Operation.get() == "Get Meaning in different Languages":
        readfromimage2()


#Fuction to browse image from user's System
def openfile():
    name = askopenfilename(initialdir="/",
                           filetypes=(("PNG File", "*.png"),("PDF File", "*.pdf") ,  ("BMP File", "*.bmp"), ("JPEG File", "*.jpeg"), ("JPG File", "*.jpg")),
                           title="Choose a file."
                           )
    PathTextBox.delete("1.0", END)
    PathTextBox.insert(END, name)

#Confirm to exit Function
def confirmation():
    MsgBox = tk.messagebox.askquestion('Exit Application', 'Are you sure you want to exit the application', icon='warning')
    if MsgBox == 'yes':
        root.destroy()
    else:
        tk.messagebox.showinfo('Return', 'You will now return to the application screen')



label1 = Label(root, text = " Text Extraction From Image Using Tesseract OCR ",  background = 'white', foreground ="#1ABC9C", font = ("Times New Roman", 15),anchor=N)
label1.place(x=330, y=20, anchor=CENTER)

label2 = Label(root, text = "Select Operation :", font = ("Times New Roman", 10))
label2.grid(column = 0, row = 2, padx = 10, pady = 60, sticky=W)

n = tk.StringVar()
Operation = ttk.Combobox(root, width = 30, textvariable = n)
Operation['values'] =('Get Synonyms and Antonyms',
                      'Get Meaning in different Languages')
Operation.place(x=135, y=59)

PathLabel = Label(root, text=" Browse Your File (Image): ",font = ("Times New Roman", 10))
PathLabel.place(x=10,y=116)

BrowseButton = Button(root, text=" Browse ", command=openfile, bg='white', fg='#1ABC9C', font=("Times New Roman", 14), borderwidth=0, relief="sunken")
BrowseButton.place(x=573,y=110)

PathTextBox = Text(root, height=2, borderwidth=0, relief="sunken")
PathTextBox.grid(row=12, column=0, padx=10, pady=10)

ReadButton = Button(root, text=" Submit ", command=checkcmbo, bg='white', fg='#1ABC9C', font=("Times New Roman", 14), borderwidth=0, relief="sunken")
ReadButton.place(x=286, y=200)

exitbutton = Button(root, text=" Exit ", command=confirmation, bg="white", fg='#1ABC9C', font=("Times New Roman", 14), borderwidth=0, relief="sunken")
exitbutton.place(x=298, y=250)

Operation.current()
root.mainloop()
