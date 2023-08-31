import os
import streamlit as st
from docx2pdf import convert
import shutil
import pythoncom

pythoncom.CoInitialize()

# specify the directory you want to list
directory = 'test_docx'

# list all files in the directory
files = [f for f in os.listdir(directory) if f.endswith('.docx')]

st.title("Document Distribution")

st.write("Select a file:")

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
        # convert the selected file to PDF
        input_file = os.path.join(directory, docx_file)
        output_file = os.path.join(directory, docx_file.replace(".docx", ".pdf"))
        convert(input_file, output_file)

        # move the file to the temporary directory
        temp_file = os.path.join(temp_dir, docx_file.replace(".docx", ".pdf"))
        if os.path.exists(temp_file):
            os.remove(temp_file)
        shutil.move(output_file, temp_dir)

        # provide a link to download the file
        st.write('File exported successfully!')
        with open(temp_file, "rb") as file:
            st.download_button("Download PDF", file.read(), file_name=docx_file.replace(".docx", ".pdf"))
