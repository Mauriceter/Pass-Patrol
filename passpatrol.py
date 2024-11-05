import os
import json
import shutil
import pdfplumber
from docx import Document
import openpyxl
from pptx import Presentation
from odf.opendocument import load as load_odt
from odf.text import P
import xlrd  # For handling .xls files
import argparse
import threading
from queue import Queue
import datetime
import re
import warnings

# ANSI escape code 
RED_TEXT = '\033[91m'
BOLD_TEXT = "\033[1m"
RESET_TEXT = '\033[0m'

# Define output folder
timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
download_folder = f"downloaded_files_{timestamp}"
os.makedirs(download_folder, exist_ok=True)

# Global variable for debug mode
debug = False
out = False
# Path to the script's directory
script_dir = os.path.dirname(os.path.abspath(__file__))

# Default paths for keywords and extensions
default_keywords_path = os.path.join(script_dir, "keywords.txt")
default_extensions_path = os.path.join(script_dir, "blacklist.txt")

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


# Function to load keywords from a specified file
def load_keywords(keywords_file):
    """Load keywords from the specified file."""
    try:
        with open(keywords_file, 'r') as kf:
            return [line.strip().lower() for line in kf.readlines()]
    except Exception as e:
        print(f"Error loading keywords: {e}")
        return []

# Function to load blacklist or whitelist of file extensions from a specified file
def load_extensions(extensions_file):
    """Load extensions from the specified file."""
    try:
        with open(extensions_file, 'r') as ef:
            return tuple(line.strip().lower() for line in ef.readlines())
    except Exception as e:
        print(f"Error loading extensions: {e}")
        return ()

def search_in_text(text, keywords):
    """Search for keywords in the text, return the keyword if found."""
    for keyword in keywords:
        # Locate the keyword in the text
        keyword_pos = text.lower().find(keyword.lower())
        if keyword_pos != -1:
            return (keyword,keyword_pos)
    return None

def create_snippet(text, keyword_tuple):
    """Generate a snippet around a keyword with up to 50 characters and at most 1 line before and 1 line after,
       highlighting the keyword in red."""
    
    keyword = keyword_tuple[0]
    keyword_pos = keyword_tuple[1]

    # Define the character limits around the keyword
    start_char = max(keyword_pos - 50, 0)
    end_char = min(keyword_pos + len(keyword) + 50, len(text))
    context_text = text[start_char:end_char]
    
    # Limit context_text to 1 line before and 1 line after the keyword line
    lines = context_text.splitlines()
    keyword_line_index = next((i for i, line in enumerate(lines) if keyword.lower() in line.lower()), None)
    if keyword_line_index is None:
        return None

    # Extract up to 1 line before and 1 line after the keyword line
    start_line = max(keyword_line_index - 1, 0)
    end_line = min(keyword_line_index + 2, len(lines))  # +2 to include the keyword line and 1 line after
    snippet = '\n'.join(lines[start_line:end_line])

    # Highlight the keyword in red
    highlighted_snippet = re.sub(re.escape(keyword), f"{RED_TEXT}\\g<0>{RESET_TEXT}", snippet, flags=re.IGNORECASE)
    
    return highlighted_snippet


# Function to handle text files
def handle_text_file(file_path, keywords):
    """Read and search in text files."""
    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
            file_data = file.read()
            keyword =  search_in_text(file_data, keywords)
            if keyword:
                if out:
                    snippet = create_snippet(file_data, keyword)
                    print(f"{BOLD_TEXT}{file_path}{RESET_TEXT}:\n{snippet}\n")
                return file_path
    except Exception as e:
        debug_print(f"Skipping text file {file_path} due to error: {e}")
    return None

# Function to handle PDF files
def handle_pdf_file(file_path, keywords):
    """Read and search in PDF files."""
    try:
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                keyword = search_in_text(text, keywords)
                if keyword:
                    if out:
                        snippet = create_snippet(text, keyword)
                        print(f"{BOLD_TEXT}{file_path}{RESET_TEXT}:\n{snippet}\n")
                    return file_path
    except Exception as e:
        debug_print(f"Skipping PDF file {file_path} due to error: {e}")
    return None

# Function to handle .docx files
def handle_docx_file(file_path, keywords):
    """Read and search in .docx files."""
    try:
        doc = Document(file_path)
        text = ""
        for para in doc.paragraphs:
            text += para.text
        keyword = search_in_text(text, keywords)
        if keyword:
            if out:
                snippet = create_snippet(text, keyword)
                print(f"{BOLD_TEXT}{file_path}{RESET_TEXT}:\n{snippet}\n")
            return file_path
    except Exception as e:
        debug_print(f"Skipping .docx file {file_path} due to error: {e}")
    return None

# Function to handle .xlsx files
def handle_xlsx_file(file_path, keywords):
    """Read and search in .xlsx files."""
    try:
        wb = openpyxl.load_workbook(file_path)
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            for row in ws.iter_rows(values_only=True):
                for cell in row:
                    keyword = search_in_text(str(cell), keywords)
                    if keyword:
                        if out:
                            snippet = create_snippet(str(cell), keyword)
                            print(f"{BOLD_TEXT}{file_path}{RESET_TEXT}:\n{snippet}\n")
                        return file_path
    except Exception as e:
        debug_print(f"Skipping .xlsx file {file_path} due to error: {e}")
    return None

# Function to handle .pptx files
def handle_pptx_file(file_path, keywords):
    """Read and search in .pptx files."""
    try:
        prs = Presentation(file_path)
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    keyword = search_in_text(shape.text, keywords)
                    if keyword:
                        if out:
                            snippet = create_snippet(shape.text, keyword)
                            print(f"{BOLD_TEXT}{file_path}{RESET_TEXT}:\n{snippet}\n")
                        return file_path
    except Exception as e:
        debug_print(f"Skipping .pptx file {file_path} due to error: {e}")
    return None


# Handle .odt files
def handle_odt_file(file_path, keywords):
    """Read and search in .odt files."""
    try:
        doc = load_odt(file_path)
        paragraphs = [element.textContent for element in doc.getElementsByType(P)]
        text = " ".join(paragraphs)
        keyword = search_in_text(text, keywords)
        if keyword:
            if out:
                snippet = create_snippet(text, keyword)
                print(f"{BOLD_TEXT}{file_path}{RESET_TEXT}:\n{snippet}\n")
            return file_path
    except Exception as e:
        debug_print(f"Skipping .odt file {file_path} due to error: {e}")
    return None

# Handle .ppt files
def handle_ppt_file(file_path, keywords):
    """Read and search in .ppt files using python-pptx."""
    try:
        prs = Presentation(file_path)
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    keyword = search_in_text(shape.text, keywords)
                    if keyword:
                        if out:
                            snippet = create_snippet(shape.text, keyword)
                            print(f"{BOLD_TEXT}{file_path}{RESET_TEXT}:\n{snippet}\n")
                        return file_path
    except Exception as e:
        debug_print(f"Skipping .ppt file {file_path} due to error: {e}")
    return None

# Handle .doc files
def handle_doc_file(file_path, keywords):
    """Read and search in .doc files using python-docx."""
    try:
        doc = Document(file_path)
        text = ""
        for para in doc.paragraphs:
            text += para.text
        keyword = search_in_text(text, keywords)
        if keyword:
            if out:
                snippet = create_snippet(text, keyword)
                print(f"{BOLD_TEXT}{file_path}{RESET_TEXT}:\n{snippet}\n")
            return file_path
    except Exception as e:
        debug_print(f"Skipping .doc file {file_path} due to error: {e}")
    return None

# Handle .xls files using xlrd
def handle_xls_file(file_path, keywords):
    """Read and search in .xls files."""
    try:
        workbook = xlrd.open_workbook(file_path)
        for sheet in workbook.sheets():
            for row_idx in range(sheet.nrows):
                row = sheet.row(row_idx)
                for cell in row:
                    keyword = search_in_text(str(row), keywords)
                    if keyword:
                        if out:
                            snippet = create_snippet(str(row), keyword)
                            print(f"{BOLD_TEXT}{file_path}{RESET_TEXT}:\n{snippet}\n")
                        return file_path
    except Exception as e:
        debug_print(f"Skipping .xls file {file_path} due to error: {e}")
    return None

# Function to check if the file should be processed based on its extension
def is_supported_file(file_path, extensions, use_whitelist):
    """Check if the file should be processed based on its extension."""
    if use_whitelist:
        return file_path.lower().endswith(extensions)
    else:
        return not file_path.lower().endswith(extensions)

# Function to handle files based on their type
def handle_file(file_path, keywords, extensions, use_whitelist):
    """Determine file type and search."""
    if is_supported_file(file_path, extensions, use_whitelist):
        if file_path.lower().endswith('.pdf'):
            return handle_pdf_file(file_path, keywords)
        elif file_path.lower().endswith('.docx'):
            return handle_docx_file(file_path, keywords)
        elif file_path.lower().endswith('.xlsx'):
            return handle_xlsx_file(file_path, keywords)
        elif file_path.lower().endswith('.pptx'):
            return handle_pptx_file(file_path, keywords)
        elif file_path.lower().endswith('.odt'):
            return handle_odt_file(file_path, keywords)
        elif file_path.lower().endswith('.ppt'):
            return handle_ppt_file(file_path, keywords)
        elif file_path.lower().endswith('.xls'):
            return handle_xls_file(file_path, keywords)
        elif file_path.lower().endswith('.doc'):
            return handle_doc_file(file_path, keywords)
        else:
            return handle_text_file(file_path, keywords)
    else:
        debug_print(f"Ignoring unsupported file type: {file_path}")
    return None

# Create a metadata dictionary to keep track of original paths
metadata = {}

# Function to ensure the filename is unique in the download folder
def ensure_unique_filename(original_name):
    """Ensure the filename is unique by appending a number if necessary."""
    base, ext = os.path.splitext(original_name)
    count = 1
    new_name = original_name
    while os.path.exists(os.path.join(download_folder, new_name)):
        new_name = f"{base}_{count}{ext}"  # Append a counter to the filename
        count += 1
    return new_name

# Function to copy the found file to the download folder while logging its original path
def copy_file(file_path):
    """Copy the found file to the download folder while logging its original path."""
    try:
        # Get the original file name
        original_name = os.path.basename(file_path)  # Get only the filename without path
        
        # Ensure a unique filename
        safe_filename = ensure_unique_filename(original_name)  
        
        # Set the destination path
        destination_path = os.path.join(download_folder, safe_filename)

        # Copy the file to the destination path
        shutil.copy(file_path, destination_path)

        # Log the original path in the metadata JSON file (Assuming log_metadata is implemented)
        log_metadata(safe_filename, file_path)

        debug_print(f"Copied {file_path} to {destination_path}")
    except Exception as e:
        debug_print(f"Error copying file {file_path}: {e}")


# Function to log metadata to JSON file
def log_metadata(filename, original_path):
    """Log the original path of the copied file in the metadata JSON."""
    metadata_entry = {filename: original_path}

    # Open the metadata file in append mode
    metadata_file_path = os.path.join(script_dir, f'metadata_{timestamp}.json')

    # Check if the metadata file already exists
    if not os.path.exists(metadata_file_path):
        # If not, create the file and write an opening bracket for a JSON array
        with open(metadata_file_path, 'w') as f:
            f.write('[\n')

    # Append the new metadata entry to the JSON file
    with open(metadata_file_path, 'a') as f:
        json.dump(metadata_entry, f, indent=4)
        f.write(',\n')  # Add a comma after each entry

# Function to finalize the metadata JSON file
def finalize_metadata():
    """Finalize the metadata JSON file by replacing the last comma with a closing bracket."""
    metadata_file_path = os.path.join(script_dir, f'metadata_{timestamp}.json')
    # Read the current content of the file
    with open(metadata_file_path, 'r+') as f:
        lines = f.readlines()
        if len(lines) > 1:
            # Remove the last comma and add a closing bracket
            lines[-2] = lines[-2].rstrip(',\n') + '\n'  # Remove the last comma
            lines.append(']\n')  # Add the closing bracket
            f.seek(0)  # Move to the beginning of the file
            f.writelines(lines)  # Write the modified lines back to the file
            f.truncate()  # Truncate the file to the new size
            
# Function to process files in a directory using os.scandir()
def file_generator(mount_directory, queue):
    """Yield file paths in the directory and put them into the queue."""
    try:
        for entry in os.scandir(mount_directory):
            if entry.is_file():
                queue.put(entry.path)
            elif entry.is_dir():  # Recursively traverse subdirectories
                file_generator(entry.path, queue)
    except Exception as e:
        debug_print(f"Error accessing {mount_directory}: {e}")

# Function to process files from the queue
def process_files(queue, keywords, extensions, use_whitelist):
    """Worker function to process files from the queue."""
    while True:
        file_path = queue.get()
        if file_path is None:  # Stop signal
            break
        result = handle_file(file_path, keywords, extensions, use_whitelist)
        if result:
            copy_file(result)
        queue.task_done()

# Function to search and copy files using multithreading
def search_and_copy_files(mount_directory, keywords, extensions, use_whitelist):
    """Search and copy files using multithreading."""
    # Create a queue to hold file paths
    queue = Queue()

    # Create threads for processing files
    threads = []
    for _ in range(os.cpu_count()):  # Use the number of CPU cores
        thread = threading.Thread(target=process_files, args=(queue, keywords, extensions, use_whitelist))
        thread.start()
        threads.append(thread)

    # Start listing files and putting them into the queue
    file_generator(mount_directory, queue)

    # Wait for the queue to finish processing
    queue.join()

    # Stop workers
    for _ in threads:
        queue.put(None)  # Sending stop signal to workers
    for thread in threads:
        thread.join()
    
    # Finalize the metadata JSON file
    finalize_metadata()

# Function for debug printing
def debug_print(message):
    """Print debug messages if debugging is enabled."""
    if debug:
        print(message)

# Main function for argument parsing and execution
def main():
    # Set up argument parser
    parser = argparse.ArgumentParser(description='Search files for specific keywords and copy them.')
    parser.add_argument('--shares', required=True, type=str, help='Path to the directory containing mounted shares.')
    parser.add_argument('--keywords', type=str, help='Path to the file containing keywords.', default=default_keywords_path)
    parser.add_argument('--extensions', type=str, help='Path to the file containing extensions (blacklist or whitelist).', default=default_extensions_path)
    parser.add_argument('--whitelist', action='store_true', help='Use the provided extensions as a whitelist instead of a blacklist.')
    parser.add_argument('--debug', action='store_true', help='Enable debug output.')
    parser.add_argument('--out', action='store_true', help='Print keyword identified.')

    # Parse arguments
    args = parser.parse_args()

    global debug
    debug = args.debug  # Set global debug variable
    global out
    out = args.out

    # Load keywords and extensions
    keywords = load_keywords(args.keywords)
    extensions = load_extensions(args.extensions)

    # Run the search and copy
    search_and_copy_files(args.shares, keywords, extensions, args.whitelist)
    print(f"Search completed. Files containing keywords are saved in '{download_folder}'.")

# Run the script
if __name__ == '__main__':
    main()
