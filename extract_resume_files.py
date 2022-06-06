# -*- coding: utf-8 -*-
"""Extract_Resume_Files.ipynb

Automatically generated by Colaboratory.

Original file is located at
    https://colab.research.google.com/drive/1ijPk2UOxGGBxif8HQpx5j9puRVaYRMiw

**Business objective -** The document classification solution should significantly reduce the manual human effort in the HRM and financial department. It should achieve a higher level of accuracy and automation with minimal human intervention.

**Sample Data Set Details** Resumes and financial documents.


### **Extract the Resume files**
"""

# Using Google Drive Path.
from google.colab import drive
drive.mount('/content/drive')

# Install the Package for Extract the files.
!pip install textract
!pip install python-docx
!python3 -m pip install docx2txt
!sudo apt-get install antiword

# Import Libaries for Extract the files.
import io
import os
import pandas as pd
import textract, docx2txt
from docx import Document

import warnings
warnings.filterwarnings('ignore')

# Original resumes directory.
source_directory = '/content/drive/MyDrive/Colab Notebooks/Projects/Resumes'

file_docx = []
file_path = []

def folder_path(source_directory):
  for file_folder in os.listdir(source_directory):
    if file_folder == 'Developer resumes':
      final_path = os.path.join(source_directory, file_folder)
      for file_name in os.listdir(final_path):
        if file_name.endswith('.docx') or file_name.endswith('.doc'):
          final_folder_path = os.path.join(final_path,file_name)
          file_docx.append(textract.process(final_folder_path))
          file_path.append(file_folder)

    elif file_folder == 'Internship resumes':
      final_path = os.path.join(source_directory, file_folder)
      for file_name in os.listdir(final_path):
        if file_name.endswith('.docx') or file_name.endswith('.doc'):
          final_folder_path = os.path.join(final_path, file_name)
          file_docx.append(textract.process(final_folder_path))
          file_path.append(file_folder)
          
    elif file_folder == 'JS Developer resumes':
      final_path = os.path.join(source_directory,file_folder)
      for file_name in os.listdir(final_path):
        if file_name.endswith('.docx') or file_name.endswith('.doc') or file_name.endswith('.pdf'):
          final_folder_path = os.path.join(final_path, file_name)
          file_docx.append(textract.process(final_folder_path))
          file_path.append(file_folder)
          
    elif file_folder == 'Peoplesoft resumes':
      final_path = os.path.join(source_directory, file_folder)
      for file_name in os.listdir(final_path):
        if file_name.endswith('.docx') or file_name.endswith('.doc'):
          final_folder_path = os.path.join(final_path, file_name)          
          file_docx.append(textract.process(final_folder_path))
          file_path.append(file_folder)
                
    elif file_folder == 'SQL Developer Lightning insight':
      final_path = os.path.join(source_directory, file_folder)
      for file_name in os.listdir(final_path):
        if file_name.endswith('.docx') or file_name.endswith('.doc'):
          final_folder_path = os.path.join(final_path, file_name)          
          file_docx.append(textract.process(final_folder_path))
          file_path.append(file_folder)         

    elif file_folder == 'workday resumes':
      final_path = os.path.join(source_directory, file_folder)
      for file_name in os.listdir(final_path):
        if file_name.endswith('.docx') or file_name.endswith('.doc'):
          final_folder_path = os.path.join(final_path, file_name)         
          file_docx.append(textract.process(final_folder_path))
          file_path.append(file_folder)

# Read the files from source_directory.
folder_path(source_directory)

# length of the files.
len(file_docx)

# Extracted files in docx, doc and pdf.
file_docx

a = '/content/drive/MyDrive/Colab Notebooks/Projects/Resumes/Developer resumes'
b = '/content/drive/MyDrive/Colab Notebooks/Projects/Resumes/workday resumes'
c = '/content/drive/MyDrive/Colab Notebooks/Projects/Resumes/Internship resumes'
d = '/content/drive/MyDrive/Colab Notebooks/Projects/Resumes/SQL Developer Lightning insight'
e = '/content/drive/MyDrive/Colab Notebooks/Projects/Resumes/JS Developer resumes'
f = '/content/drive/MyDrive/Colab Notebooks/Projects/Resumes/Peoplesoft resumes'

name_1 = []
name_2 = []
name_3 = []
name_4 = []
name_5 = []
name_6 = []

for name_file_1 in os.listdir(a):
  if not name_file_1.startswith('.'):
    file, extension = os.path.splitext(name_file_1)
    name_1.append(file)

for name_file_2 in os.listdir(b):
  if not name_file_2.startswith('.'):
    file, extension = os.path.splitext(name_file_2)
    name_2.append(file)

for name_file_3 in os.listdir(c):
  if not name_file_3.startswith('.'):
    file, extension = os.path.splitext(name_file_3)
    name_3.append(file)

for name_file_4 in os.listdir(d):
  if not name_file_4.startswith('.'):
    file, extension = os.path.splitext(name_file_4)
    name_4.append(file)

for name_file_5 in os.listdir(e):
  if not name_file_5.startswith('.'):
    file, extension = os.path.splitext(name_file_5)
    name_5.append(file)

for name_file_6 in os.listdir(f):
  if not name_file_6.startswith('.'):
    file, extension = os.path.splitext(name_file_6)
    name_6.append(file)

df_name_1 = pd.DataFrame(name_1)
df_name_2 = pd.DataFrame(name_2)
df_name_3 = pd.DataFrame(name_3)
df_name_4 = pd.DataFrame(name_4)
df_name_5 = pd.DataFrame(name_5)
df_name_6 = pd.DataFrame(name_6)

df_names = pd.concat([df_name_1,df_name_2,df_name_3,df_name_4,df_name_5,df_name_6], axis=0)

# Create the dataframe and Name the Columns.
df = pd.DataFrame()

df['Names'] = df_names
df['Profiles'] = file_path
df['Summary'] = file_docx
df.head(79)

# Shape of 'df'.
df.shape

# Info of 'df'.
df.info()

# Using lambda function and decode 'df'.
df['Summary'] = df['Summary'].apply(lambda x : x.decode('utf-8'))
df.head()

# Exporting the data into csv files.
df.to_csv('Text.csv', encoding='utf-8', index=True)
df =pd.read_csv("Text.csv", encoding='utf-8')

"""### THE END"""