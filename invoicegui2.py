import os
import pdfplumber
import re
import csv
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QPushButton,
    QLabel, QFileDialog, QTableWidget, QTableWidgetItem, QWidget,
    QMessageBox, QComboBox, QProgressBar, QDialog, QHBoxLayout, QRadioButton,
    QButtonGroup, QLineEdit, QScrollArea
)
from PyQt5.QtCore import Qt


class ReviewDialog(QDialog):
    def __init__(self, file_path, extracted_text, potential_totals, best_guess, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Review Unmatched File")
        self.setGeometry(300, 200, 800, 600)
        
        self.selected_total = None
        
        self.layout = QVBoxLayout()
        
        self.label = QLabel(f"Review totals for file: {os.path.basename(file_path)}")
        self.layout.addWidget(self.label)
        
        # Display extracted text in a scrollable area
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.text_widget = QLabel(extracted_text)
        self.text_widget.setWordWrap(True)
        self.scroll_area.setWidget(self.text_widget)
        self.layout.addWidget(self.scroll_area)
        
        self.radio_label = QLabel("Select the total from the options below:")
        self.layout.addWidget(self.radio_label)
        
        # Radio buttons for potential totals
        self.radio_group = QButtonGroup(self)
        self.radio_layout = QVBoxLayout()
        
        for total in potential_totals:
            radio_button = QRadioButton(total)
            self.radio_group.addButton(radio_button)
            self.radio_layout.addWidget(radio_button)
            if total == best_guess:  # Automatically select the best guess
                radio_button.setChecked(True)
        
        radio_widget = QWidget()
        radio_widget.setLayout(self.radio_layout)
        self.layout.addWidget(radio_widget)
        
        # Editable field for manual input
        self.manual_label = QLabel("Or enter a custom total:")
        self.layout.addWidget(self.manual_label)
        
        self.manual_input = QLineEdit()
        self.layout.addWidget(self.manual_input)
        
        # Buttons for Accept and Reject
        self.buttons_layout = QHBoxLayout()
        
        self.accept_button = QPushButton("Accept")
        self.accept_button.clicked.connect(self.accept)
        self.buttons_layout.addWidget(self.accept_button)
        
        self.reject_button = QPushButton("Reject")
        self.reject_button.clicked.connect(self.reject)
        self.buttons_layout.addWidget(self.reject_button)
        
        self.layout.addLayout(self.buttons_layout)
        self.setLayout(self.layout)
    
    def get_selected_total(self):
        # Check if a radio button is selected
        selected_radio = self.radio_group.checkedButton()
        if selected_radio:
            return selected_radio.text()
        # If no radio button is selected, return the manual input
        return self.manual_input.text()


class InvoiceExtractorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Invoice Total Extractor")
        self.setGeometry(300, 200, 600, 400)
        
        # Layout and Widgets
        self.layout = QVBoxLayout()
        
        self.label = QLabel("Select a folder containing invoice PDFs:")
        self.layout.addWidget(self.label)
        
        self.select_folder_button = QPushButton("Select Folder")
        self.select_folder_button.clicked.connect(self.select_folder)
        self.layout.addWidget(self.select_folder_button)
        
        # Editable Combo Box for Currency Selection
        self.currency_label = QLabel("Select or enter the currency to search for as it appears in the invoice (eg. £ or US$):")
        self.layout.addWidget(self.currency_label)
        
        self.currency_selector = QComboBox()
        self.currency_selector.addItems(["£", "$", "€", "₹", "¥", "₽"])  # Predefined currencies
        self.currency_selector.setEditable(True)  # Make the combo box editable
        self.layout.addWidget(self.currency_selector)
        
        self.extract_button = QPushButton("Extract Totals")
        self.extract_button.setEnabled(False)
        self.extract_button.clicked.connect(self.extract_totals)
        self.layout.addWidget(self.extract_button)
        
        # Progress Bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(self.progress_bar)
        
        self.results_table = QTableWidget()
        self.layout.addWidget(self.results_table)
        
        self.save_button = QPushButton("Save Results to CSV")
        self.save_button.setEnabled(False)
        self.save_button.clicked.connect(self.save_to_csv)
        self.layout.addWidget(self.save_button)
        
        # Central Widget
        central_widget = QWidget()
        central_widget.setLayout(self.layout)
        self.setCentralWidget(central_widget)
        
        self.folder_path = ""
        self.totals = {}
        self.totalsum = 0
        self.unmatched_files = []

    def select_folder(self):
        self.folder_path = QFileDialog.getExistingDirectory(self, "Select Folder")
        if self.folder_path:
            self.label.setText(f"Selected Folder: {self.folder_path}")
            self.extract_button.setEnabled(True)

    def extract_totals(self):
        self.totalsum = 0
        if not os.path.isdir(self.folder_path):
            QMessageBox.critical(self, "Error", "Invalid folder path!")
            return
        
        selected_currency = self.currency_selector.currentText()
        self.totals, self.unmatched_files = self._extract_invoice_totals(self.folder_path, selected_currency)
        
        if self.unmatched_files:
            self._review_unmatched_files()
        
        if self.totals:
            self._populate_table()
            QMessageBox.information(self, "Success", f"Totals extracted successfully! Grand total: {selected_currency}{self.totalsum}")
            self.save_button.setEnabled(True)
        else:
            QMessageBox.warning(self, "No Totals Found", "No totals were found in the PDFs.")

    def _extract_invoice_totals(self, folder_path, currency_symbol):
        totals = {}
        unmatched_files = []
        files_to_process = []

        # Collect all PDF files
        for root, _, files in os.walk(folder_path):
            for filename in files:
                if filename.endswith(".pdf"):
                    files_to_process.append(os.path.join(root, filename))

        total_files = len(files_to_process)
        self.progress_bar.setMaximum(total_files)
        self.progress_bar.setValue(0)

        for index, file_path in enumerate(files_to_process, start=1):
            try:
                with pdfplumber.open(file_path) as pdf:
                    text = "\n".join([page.extract_text() or "" for page in pdf.pages])
                    # Default regex for matching totals
                    pattern = rf'Total\s*[:\-]?\s*{re.escape(currency_symbol)}([\d,]+\.\d{{2}})'
                    match = re.search(pattern, text)
                    if match:
                        total = match.group(1).replace(",", "")
                        totals[file_path] = f"{float(total)}"
                        self.totalsum += float(total)
                    else:
                        unmatched_files.append((file_path, text))
            except Exception as e:
                print(f"Error processing {file_path}: {e}")
            
            # Update progress bar
            self.progress_bar.setValue(index)

        return totals, unmatched_files

    def _populate_table(self):
        self.results_table.setRowCount(len(self.totals))
        self.results_table.setColumnCount(2)
        self.results_table.setHorizontalHeaderLabels(["Filename", "Total"])

        for row, (file_path, total) in enumerate(self.totals.items()):
            # Show only the filename (no root path) in the GUI
            filename = os.path.basename(file_path)
            self.results_table.setItem(row, 0, QTableWidgetItem(filename))
            self.results_table.setItem(row, 1, QTableWidgetItem(total))
        
        # Automatically resize the columns and rows to fit the content
        self.results_table.setColumnWidth(0, int(self.results_table.width() * 0.75))
        self.results_table.setColumnWidth(1, int(self.results_table.width() * 0.18))
        self.results_table.resizeRowsToContents()

    def save_to_csv(self):
        if not self.totals:
            QMessageBox.warning(self, "No Data", "No data to save!")
            return
        
        save_path, _ = QFileDialog.getSaveFileName(self, "Save CSV", "", "CSV Files (*.csv)")
        if save_path:
            with open(save_path, mode="w", newline="") as csvfile:
                writer = csv.writer(csvfile)
                # Write the header
                writer.writerow(["Filename (Full Path)", "Total"])
                for file_path, total in self.totals.items():
                    # Write full path to the CSV for accuracy
                    writer.writerow([file_path, total])
                writer.writerow(["Grand Total", self.totalsum])
            QMessageBox.information(self, "Saved", f"Results saved to {save_path}")


    def _review_unmatched_files(self):
        for file_path, text in self.unmatched_files:
            # Find all potential totals
            potential_totals = re.findall(r'([\d,]+\.\d{2})', text)
            
            # Convert to floats, filter numbers > 10.0, and remove duplicates while preserving order
            unique_totals = []
            seen = set()
            for total in (float(x.replace(",", "")) for x in potential_totals if float(x.replace(",", "")) > 10.0):
                if total not in seen:
                    unique_totals.append(total)
                    seen.add(total)
            
            # Sort totals in descending order and take the top 4–5
            top_totals = [f"{x:,.2f}" for x in sorted(unique_totals, reverse=True)[:5]]
            
            # Best guess: The largest number
            best_guess = top_totals[0] if top_totals else None
            
            dialog = ReviewDialog(file_path, text, top_totals, best_guess, self)
            if dialog.exec_() == QDialog.Accepted:
                total = dialog.get_selected_total()
                if total:
                    self.totals[file_path] = total
                    self.totalsum += float(total)



def main():
    app = QApplication([])
    window = InvoiceExtractorApp()
    window.show()
    app.exec_()

if __name__ == "__main__":
    main()
