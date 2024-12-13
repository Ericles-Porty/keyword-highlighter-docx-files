# Project: Highlight Keywords in .docx Files

This project processes Word documents (.docx) in an input folder, searches for keywords defined in a text file, and highlights those keywords in the documents found. The modified files are saved in an output folder.

---

## Features

 * Keyword Search: Identifies occurrences of user-defined keywords in .docx files.

 * Word Highlighting: Highlights only the matching words within the paragraphs found.

 * Output Folder: Automatically saves modified files in a specified folder.
  
---

## Requirements
    
 * Python 3.6 or higher
    
 * Libraries
   * python-docx (installable via pip).

To install the required library:
```bash
pip install python-docx
```

---

## File Structure

 * keywords.txt: File containing one keyword per line.
    * Example:
        ```
        example
        test
        keyword
        ```

    * input_docs/: Folder containing the .docx files to be processed.
    * output_docs/: Folder where the modified files will be saved.

---

## How to Use

 1. Setup
      * Create a keywords.txt file containing the keywords to search for, one per line.
      * Place the .docx files you want to process in the input_docs folder.
 2. Run the Script:
      * Copy the code into a Python file (e.g., process_docx.py).
      * Update the paths for keywords_file, input_folder, and output_folder if needed.
      * Run the script:
        ```bash
        python process_docx.py
        ```
 3. Results
   * The .docx files containing keywords will be saved in the output_docs folder with the words highlighted.

---

## Main Code

### Path Configuration
```python
# Configuration
keywords_file = 'keywords.txt'       # Path to the keywords file
input_folder = 'input_docs'          # Folder containing .docx files
output_folder = 'output_docs'        # Folder to save modified files
```

### Script Structure
 1. Load Keywords: Reads keywords from a text file.
 2. Process Documents: Iterates through .docx files in the input folder and searches for keywords.
 3. Highlight Words: Highlights the matching words in the document text.
 4. Save Results: Only documents with found keywords are saved to the output folder.

---

## Contribution
Feel free to suggest improvements or report issues in the code. Pull requests are welcome!

---

## License

This project is open-source and licensed under the MIT License.