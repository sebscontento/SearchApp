#author: Sebastian Solórzano Holbøll


#version 8.0.
#relased version 




import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QFileDialog, QTableWidget, QTableWidgetItem, QHBoxLayout, QHeaderView
from PyQt5.QtCore import QSettings

import os
import re
import glob
import fitz  # PyMuPDF
from docx import Document

class PDFSearchApp(QWidget):
    def __init__(self):
        super().__init__()

        # Load recent directory from settings
        self.settings = QSettings("MyCompany", "PDFSearchApp")
        self.recent_directory = self.settings.value("recent_directory", "")

        self.initUI()

    def initUI(self):
        self.setWindowTitle('SearchApp')

        # Widgets
        self.directory_label = QLabel('Directory:')
        self.directory_entry = QLineEdit()
        self.browse_button = QPushButton('Browse')
        self.search_word_label = QLabel('Search Word:')
        self.search_word_entry = QLineEdit()
        self.search_button = QPushButton('Search')
        self.table = QTableWidget()
        self.download_button = QPushButton('Download Output')

        # Search result widgets
        self.search_result_label = QLabel()
        self.search_result_entry = QLineEdit()
        self.search_result_entry.setPlaceholderText('Search in results')

        # Layout
        layout = QVBoxLayout()
        layout.addWidget(self.directory_label)
        layout.addWidget(self.directory_entry)
        layout.addWidget(self.browse_button)
        layout.addWidget(self.search_word_label)
        layout.addWidget(self.search_word_entry)
        layout.addWidget(self.search_button)

        # Table layout
        table_layout = QHBoxLayout()
        table_layout.addWidget(self.table)

        layout.addLayout(table_layout)

        layout.addWidget(self.download_button)

        # Search result layout
        search_result_layout = QVBoxLayout()
        search_result_layout.addWidget(self.search_result_label)
        search_result_layout.addWidget(self.search_result_entry)

        layout.addLayout(search_result_layout)

        self.setLayout(layout)

        # Set up the table
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(['File', 'Result'])
        self.table.horizontalHeader().setStretchLastSection(True)

        # Set recent directory if available
        if self.recent_directory:
            self.directory_entry.setText(self.recent_directory)

        # Signals and Slots
        self.browse_button.clicked.connect(self.browse_directory)
        self.search_button.clicked.connect(self.execute_search)
        self.download_button.clicked.connect(self.download_output)

        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)


        self.search_result_entry.textChanged.connect(self.filter_results)

    def browse_directory(self):
        directory = QFileDialog.getExistingDirectory(self, 'Select Directory', self.recent_directory)
        if directory:
            self.directory_entry.setText(directory)

    def execute_search(self):
        directory = self.directory_entry.text()
        search_word = self.search_word_entry.text()
        self.table.setRowCount(0)  # Clear previous results

        if directory and search_word:
            pdf_result, docx_result = self.run_document_search(directory, search_word)

          

            pdf_sentences = self.extract_sentences(pdf_result, search_word)
            docx_sentences = self.extract_sentences(docx_result, search_word)

    

       # Populate the table with results
        for idx, (filename, content) in enumerate(pdf_sentences + docx_sentences):
            filename = os.path.basename(filename)  # Extract only the filename from the full path
            content = self.limit_words(content, 30)  # Limit to 30 words
            self.populate_table(idx, filename, content)

        # Resize columns to fit content
        self.table.resizeColumnsToContents()
        # You can adjust the width and height values as needed
        self.resize(1000, 1400)

        # Save recent directory to settings
        self.settings.setValue("recent_directory", directory)

        # Clear the search result entry
        self.search_result_entry.clear()
    def limit_words(self, text, word_limit):
        # Limit the text to a specified number of words
        words = text.split()
        limited_words = ' '.join(words[:word_limit])
        return limited_words

    def populate_table(self, idx, filename, content):
        self.table.insertRow(idx)
        self.table.setItem(idx, 0, QTableWidgetItem(filename))
        self.table.setItem(idx, 1, QTableWidgetItem(content))

    def run_document_search(self, directory, search_word):
        pdf_result = []
        docx_result = []

        # Search for PDF files
        pdf_files = glob.glob(os.path.join(directory, '*.pdf'))
        for pdf_file in pdf_files:
            with fitz.open(pdf_file) as pdf_document:
                for page_number in range(pdf_document.page_count):
                    page = pdf_document[page_number]
                    text = page.get_text()
                    if re.search(fr'\b{re.escape(search_word)}\b', text, flags=re.IGNORECASE):
                        pdf_result.append((pdf_file, text))

        # Search for DOCX files
        docx_files = glob.glob(os.path.join(directory, '*.docx'))
        for docx_file in docx_files:
            doc = Document(docx_file)
            for paragraph in doc.paragraphs:
                if re.search(fr'\b{re.escape(search_word)}\b', paragraph.text, flags=re.IGNORECASE):
                    docx_result.append((docx_file, paragraph.text))

        return pdf_result, docx_result

    def extract_sentences(self, results, search_word):
        sentences = []

        for filename, content in results:
            matches = re.finditer(r'\b{}\b'.format(re.escape(search_word)), content, flags=re.IGNORECASE)

            for match in matches:
                # Extract the sentence containing the search word
                start, end = match.span()
                sentence = content[max(0, start - 50):min(end + 50, len(content))]  # Extract a context of 50 characters around the match
                sentences.append((filename, sentence))

        # Remove empty strings from the content
        sentences = [(filename, content) for filename, content in sentences if content]

        return sentences

    def filter_results(self, search_text):
        # Iterate through all rows in the table and hide those that don't match the search text
        for row in range(self.table.rowCount()):
            filename = self.table.item(row, 0).text()
            content = self.table.item(row, 1).text()
            match_found = search_text.lower() in filename.lower() or search_text.lower() in content.lower()
            self.table.setRowHidden(row, not match_found)

    def download_output(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_dialog, _ = QFileDialog.getSaveFileName(self, "Save Output", "", "Text Files (*.txt);;All Files (*)", options=options)

        if file_dialog:
            with open(file_dialog, 'w') as file:
                # Iterate over rows and columns in the table to write the content to the file
                for row in range(self.table.rowCount()):
                    file.write(self.table.item(row, 0).text() + ": " + self.table.item(row, 1).text() + "\n")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    pdf_search_app = PDFSearchApp()
    pdf_search_app.show()
    sys.exit(app.exec_())
