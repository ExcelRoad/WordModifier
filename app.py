import streamlit as st
import zipfile
import os
import shutil

# Function to replace author names in the text content of XML files
def replace_author_in_file(file_path, old_names, new_name):
    with open(file_path, 'r', encoding='utf-8') as file:
        content = file.read()
    
    # Replace occurrences of each old name with the new name
    for old_name in old_names:
        content = content.replace(old_name.strip(), new_name)
    
    with open(file_path, 'w', encoding='utf-8') as file:
        file.write(content)

# Function to process the DOCX file and apply text replacements
def process_docx(docx_file, old_names, new_name):
    # Create a temporary folder to extract files
    temp_dir = "extracted_files"
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    os.makedirs(temp_dir)
    
    # Convert DOCX to ZIP (DOCX is essentially a ZIP file)
    zip_file = docx_file.replace(".docx", ".zip")
    shutil.copy(docx_file, zip_file)
    
    # Extract the ZIP file
    with zipfile.ZipFile(zip_file, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)
    
    # Iterate over the relevant files (document.xml, comments.xml, etc.)
    files_to_process = [
        os.path.join(temp_dir, "word", "document.xml"),
        os.path.join(temp_dir, "word", "comments.xml")
    ]
    
    # Replace old_name with new_name in all files
    for file_path in files_to_process:
        if os.path.exists(file_path):
            replace_author_in_file(file_path, old_names, new_name)

    # Create the modified DOCX file by zipping the folder back
    modified_docx = docx_file.replace(".docx", "_modified.docx")
    with zipfile.ZipFile(modified_docx, 'w', zipfile.ZIP_DEFLATED) as new_zip:
        for root_dir, dirs, files in os.walk(temp_dir):
            for file in files:
                file_path = os.path.join(root_dir, file)
                new_zip.write(file_path, os.path.relpath(file_path, temp_dir))
    
    # Clean up the temporary extracted folder and ZIP file
    shutil.rmtree(temp_dir)
    os.remove(zip_file)
    
    return modified_docx

# Adding custom CSS to handle RTL (Right-to-Left) support
st.markdown("""
<style>
    body {
        direction: rtl;
        text-align: right;
        font-family: 'Arial', sans-serif;
    }
    .streamlit-expanderHeader {
        text-align: right;
    }
    .css-1fpmrmx {
        text-align: right;
    }
</style>
""", unsafe_allow_html=True)

# Streamlit UI setup
st.title("שינוי שמות משתמשים בקובץ WORD")

# File uploader for the DOCX file
uploaded_file = st.file_uploader("אנא בחרי קובץ WORD", type=["docx"])

if uploaded_file is not None:
    # Save the uploaded file to the current directory
    docx_filename = uploaded_file.name
    with open(docx_filename, "wb") as f:
        f.write(uploaded_file.getbuffer())
    
    st.success(f"הקובץ '{docx_filename}' הועלה בהצלחה!")
    
    # Input for new author name
    st.subheader("משתמש חדש")
    new_name = st.text_input("שם משתמש חדש")
    
    if new_name:
        # Input for old author names to replace (comma-separated)
        old_names_input = st.text_area("שמות משתמשים לשינוי (הפרדה בפסיק)")
        
        # Process the file when the button is clicked
        if st.button("תיקון קובץ"):
            if old_names_input:
                old_names = old_names_input.split(",")
                try:
                    modified_docx = process_docx(docx_filename, old_names, new_name)
                    st.success(f"הקובץ שונה בהצלחה. אנא הורידי את הקובץ החדש:")
                    
                    # Provide the download link for the modified file
                    with open(modified_docx, "rb") as f:
                        st.download_button(
                            label="הורדת קובץ",
                            data=f,
                            file_name=modified_docx,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                except Exception as e:
                    st.error(f"An error occurred: {e}")
            else:
                st.warning("אנא ספקי שמות משתמשים לשינוי.")
