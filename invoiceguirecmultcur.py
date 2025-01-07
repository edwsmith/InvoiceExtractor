import os
import pdfplumber
import re
import csv
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QPushButton,
    QLabel, QFileDialog, QTableWidget, QTableWidgetItem, QWidget,
    QMessageBox, QComboBox
)

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

    def select_folder(self):
        self.folder_path = QFileDialog.getExistingDirectory(self, "Select Folder")
        if self.folder_path:
            self.label.setText(f"Selected Folder: {self.folder_path}")
            self.extract_button.setEnabled(True)

    def extract_totals(self):
        if not os.path.isdir(self.folder_path):
            QMessageBox.critical(self, "Error", "Invalid folder path!")
            return
        
        # Get the selected or entered currency symbol
        selected_currency = self.currency_selector.currentText()
        self.totals = self._extract_invoice_totals(self.folder_path, selected_currency)
        
        if self.totals:
            self._populate_table()
            QMessageBox.information(self, "Success", "Totals extracted successfully!")
            self.save_button.setEnabled(True)
        else:
            QMessageBox.warning(self, "No Totals Found", "No totals were found in the PDFs.")

    def save_to_csv(self):
        if not self.totals:
            QMessageBox.warning(self, "No Data", "No data to save!")
            return
        
        save_path, _ = QFileDialog.getSaveFileName(self, "Save CSV", "", "CSV Files (*.csv)")
        if save_path:
            with open(save_path, mode="w", newline="") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(["Filename", "Total"])
                for file, total in self.totals.items():
                    writer.writerow([file, total])
            QMessageBox.information(self, "Saved", f"Results saved to {save_path}")

    def _populate_table(self):
        self.results_table.setRowCount(len(self.totals))
        self.results_table.setColumnCount(2)
        self.results_table.setHorizontalHeaderLabels(["Filename", "Total"])
        
        for row, (filename, total) in enumerate(self.totals.items()):
            self.results_table.setItem(row, 0, QTableWidgetItem(filename))
            self.results_table.setItem(row, 1, QTableWidgetItem(total))
    
    def _extract_invoice_totals(self, folder_path, currency_symbol):
        totals = {}
        
        for root, dirs, files in os.walk(folder_path):
            for filename in files:
                if filename.endswith(".pdf"):
                    file_path = os.path.join(root, filename)
                    try:
                        with pdfplumber.open(file_path) as pdf:
                            for page in pdf.pages:
                                text = page.extract_text()
                                if text:
                                    # Regex for matching totals with the selected currency
                                    pattern = rf'Total\s*[:\-]?\s*{re.escape(currency_symbol)}([\d,]+\.\d{{2}})'
                                    match = re.search(pattern, text)
                                    if match:
                                        total = match.group(1).replace(",", "")
                                        totals[file_path] = f"{currency_symbol}{total}"  # Format total with currency
                                        break  # Stop after finding the total on one page
                    except Exception as e:
                        print(f"Error processing {file_path}: {e}")
        print(f"Extracted totals: {totals}")  # Debug print
        return totals

def main():
    app = QApplication([])
    window = InvoiceExtractorApp()
    window.show()
    app.exec_()

if __name__ == "__main__":
    main()
