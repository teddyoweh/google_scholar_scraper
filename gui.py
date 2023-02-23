#python3 -m pip install google-search-results
#python3 -m pip install openpyxl

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QStandardItemModel, QStandardItem
from PyQt6.QtWidgets import QApplication,        QMessageBox,QMainWindow, QVBoxLayout, QHBoxLayout, QWidget, QLabel, QLineEdit, QPushButton, QFileDialog, QTableView
import sys
import pandas as pd
from serpapi import GoogleSearch

class SearchApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Search App')
        self.setFixedSize(800, 600)

        # Create input widgets
        text_label = QLabel('Search Text:')
        self.text_edit = QLineEdit()
        start_label = QLabel('Start Year:')
        self.start_edit = QLineEdit()
        end_label = QLabel('End Year:')
        self.end_edit = QLineEdit()
        self.search_button = QPushButton('Search')
        self.search_button.clicked.connect(self.search)
        hbox1 = QHBoxLayout()
        self.database_button = QPushButton("Save to JSON")
        self.excel_button = QPushButton("Save to Excel")
        self.html_button = QPushButton("Save to HTML")
        self.csv_button = QPushButton("Save to CSV")
 
        self.database_button.clicked.connect(self.save_to_json)
        self.csv_button.clicked.connect(self.save_to_csv)
        self.excel_button.clicked.connect(self.save_to_excel)
        self.html_button.clicked.connect(self.save_to_html)
        hbox1.addWidget(self.database_button)
        hbox1.addWidget(self.excel_button)
        hbox1.addWidget(self.html_button)
        hbox1.addWidget(self.csv_button)

        # Create table view
        self.table_view = QTableView()
        self.setCentralWidget(self.table_view)

        # Create layout
        input_layout = QHBoxLayout()
        input_layout.addWidget(text_label)
        input_layout.addWidget(self.text_edit)
        input_layout.addWidget(start_label)
        input_layout.addWidget(self.start_edit)
        input_layout.addWidget(end_label)
        input_layout.addWidget(self.end_edit)
        input_layout.addWidget(self.search_button)

        main_layout = QVBoxLayout()
        main_layout.addLayout(input_layout)
        main_layout.addLayout(hbox1)
        main_layout.addWidget(self.table_view)

        # Set main layout
        widget = QWidget()
        widget.setLayout(main_layout)
        self.setCentralWidget(widget)

    def search(self):
        # Get search parameters
        text = self.text_edit.text()
        startdate = self.start_edit.text()
        enddate = self.end_edit.text()
        params = {
          "engine": "google_scholar",
          "q": text,
          "hl": "en",
          "num": "20",
          "as_ylo": enddate,
          "as_yhi": startdate,
          "api_key": "4dcd70032b92445d3f391aa4117d53f2a7453492f957f8f8b23d9c5bc9bd57a3"
        }

        # Perform search
        search = GoogleSearch(params)
        results = search.get_dict()
        df = self.dict_to_df(results['organic_results'])
        self.df =df

        # Display results in table view
        model = QStandardItemModel(df.shape[0], df.shape[1])
        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                model.setItem(i, j, QStandardItem(str(df.iat[i, j])))
        model.setHorizontalHeaderLabels(df.columns)
        self.table_view.setModel(model)

    def dict_to_df(self, arr):
        df = pd.DataFrame(arr)
        for col in df.columns:
            col_data = df[col]
            if isinstance(col_data.iloc[0], dict):
                normalized = pd.json_normalize(col_data)
                for sub_col in normalized.columns:
                    df[f"{col}_{sub_col}"] = normalized[sub_col]
                df.drop(col, axis=1, inplace=True)
        return df

    def save_to_database(self):
        pass

    def save_to_excel(self):
        file_path, _ = QFileDialog.getSaveFileName(self, 'Save to Excel', '', 'Excel files (*.xlsx)')
        if file_path:
            
            self.df.to_excel(file_path, index=False)
            QMessageBox.information(self,'Excel file saved to {}'.format(file_path))

    def save_to_csv(self):
        file_path, _ = QFileDialog.getSaveFileName(self, 'Save to CSV', '', 'CSV files (*.csv)')
        if file_path:
             
            self.df.to_csv(file_path, index=False)
            QMessageBox.information(self,'CSV file saved to {}'.format(file_path))

    def save_to_json(self):
        file_path, _ = QFileDialog.getSaveFileName(self, 'Save to JSON', '', 'JSON files (*.json)')
        if file_path:
            self.df.to_json() 
            QMessageBox.information(self,'JSON file saved to {}'.format(file_path))
 


    def save_to_html(self):
        # Save to HTML file
        file, _ = QFileDialog.getSaveFileName(self, "Save to HTML", "", "HTML Files (*.html)")
        if file:
           
            self.df.to_html(file, index=False)
            QMessageBox.information(self, "Save Successful", f"Results saved to {file}.")

   
    def dict_to_df(self, arr):
        df = pd.DataFrame(arr)
        for col in df.columns:
            col_data = df[col]
            if isinstance(col_data.iloc[0], dict):
                normalized = pd.json_normalize(col_data)
                for sub_col in normalized.columns:
                    df[f"{col}_{sub_col}"] = normalized[sub_col]
                df.drop(col, axis=1, inplace=True)
        return df

if __name__ == '__main__':
    app = QApplication(sys.argv)
    search_app = SearchApp()
    search_app.show()
    sys.exit(app.exec())
