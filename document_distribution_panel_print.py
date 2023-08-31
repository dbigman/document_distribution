import os
import streamlit as st
import shutil
import pythoncom

pythoncom.CoInitialize()

# specify the directory you want to list
directory = 'test_docx'

# list all files in the directory
files = [f for f in os.listdir(directory) if f.endswith('.docx')]

st.title("Docx Printer")

st.write("Select a file to print:")

# create a temporary directory and delete any existing files
temp_dir = "temp"
if not os.path.exists(temp_dir):
    os.makedirs(temp_dir)
else:
    for file in os.listdir(temp_dir):
        file_path = os.path.join(temp_dir, file)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
        except Exception as e:
            st.write(e)

# create a button for each file
for docx_file in files:
    if st.button(docx_file):
        # move the file to the temporary directory
        temp_file = os.path.join(temp_dir, docx_file)
        shutil.copy(os.path.join(directory, docx_file), temp_file)

        # provide a link to open the file
        st.write('Click the link below to open the file and print it manually:')
        st.markdown(f'<a href="{temp_file}" target="_blank">{docx_file}</a>', unsafe_allow_html=True)
