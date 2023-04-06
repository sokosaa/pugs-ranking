import sys
import pandas as pd
import pytesseract
from PyQt5.QtWidgets import QApplication, QLabel, QVBoxLayout, QWidget, QPushButton, QSizePolicy, QRadioButton, QHBoxLayout, QGroupBox, QTextEdit, QTableWidget, QTableWidgetItem, QHeaderView
from PyQt5.QtGui import QPixmap, QImage, QPalette, QColor, QKeySequence, QIcon
from PIL import Image
from PIL.ImageQt import ImageQt
import re

from PyQt5.QtGui import QKeySequence
from PyQt5.QtWidgets import QApplication, QLabel, QVBoxLayout, QWidget, QPushButton, QSizePolicy, QRadioButton, QHBoxLayout, QGroupBox, QTextEdit
from PyQt5.QtCore import Qt, QEvent

from PyQt5.QtWidgets import QTableWidget, QTableWidgetItem, QHeaderView

import pandas as pd
import math
import matplotlib.pyplot as plt
import numpy as np
import os
import subprocess
import platform

from PyQt5.QtWidgets import QMessageBox

from PyQt5.QtWidgets import QPushButton

import elo8

from PyQt5.QtCore import QTimer

from PyQt5.QtWidgets import QGridLayout

def show_popup(message):
    popup = QMessageBox()
    popup.setWindowTitle("Message")
    popup.setText(message)
    popup.exec_()


def close_excel_file(file_name):
    if platform.system() == 'Windows':
        os.system(f'taskkill /F /IM excel.exe /FI "WindowTitle eq {file_name} - Excel"')
    else:
        print("Closing Excel not supported on this platform.")

def open_excel_file(file_name):
    if platform.system() == 'Windows':
        os.startfile(file_name)
    else:
        print("Opening Excel not supported on this platform.")


class ExcelTableWidget(QWidget):
    def __init__(self, file_name):
        super().__init__()

        self.file_name = file_name
        self.table = QTableWidget(self)
        self.layout = QVBoxLayout(self)
        self.layout.addWidget(self.table)

        # Add a button to insert a new row
        self.add_row_button = QPushButton("Add Row")
        self.add_row_button.clicked.connect(self.add_row)
        self.layout.addWidget(self.add_row_button)

        # Add a button to refresh and save the table
        # self.refresh_button = QPushButton("Refresh")
        # self.refresh_button.clicked.connect(self.refresh_and_save)
        # self.layout.addWidget(self.refresh_button)

        self.load_data()

        # Connect the cellChanged signal to handle_table_change method
        self.table.cellChanged.connect(self.handle_table_change)

    def load_data(self):
        df = pd.read_excel(self.file_name)
        df = df[['Player', 'Tank', 'Damage', 'Support', 'High']]  # Select the desired columns
        df = df.dropna()  # Remove rows with missing data

        self.table.setRowCount(df.shape[0])  # Remove - 1
        self.table.setColumnCount(df.shape[1])

        # Set column names
        self.table.setHorizontalHeaderLabels(['Player', 'Tank', 'Damage', 'Support', 'High'])

        for row in range(df.shape[0]):  # Start from row 0, remove - 1
            for col in range(df.shape[1]):
                value = df.iloc[row, col]
                if isinstance(value, (int, float)):
                    value = round(value)
                item = QTableWidgetItem(str(value))
                self.table.setItem(row, col, item)  # Remove - 1

        # Resize columns to fit content
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)




    def handle_table_change(self, row, col):
        # Get the updated cell value
        item = self.table.item(row, col)
        updated_value = item.text() if item else ""

        # Read the data from the file
        df = pd.read_excel(self.file_name)

        # Update the value in the DataFrame
        column_name = df.columns[col + 11]  # Change this line, add 11 to col index

        # Check if the row exists in the DataFrame, if not, add a new row
        if row >= df.shape[0] - 1:
            new_row = pd.DataFrame([[""] * df.shape[1]], columns=df.columns)
            df = pd.concat([df, new_row], ignore_index=True)

        df.iloc[row + 1, df.columns.get_loc(column_name)] = updated_value

        # Save the changes back to the file
        df.to_excel(self.file_name, index=False)





    def add_row(self):
        # Add a new row in the table
        row_count = self.table.rowCount()
        self.table.setRowCount(row_count + 1)

        # Read the data from the file
        df = pd.read_excel(self.file_name)

        # Add a new row in the DataFrame and set the values to empty strings
        new_row = pd.DataFrame([[""] * (df.shape[1])], columns=df.columns)
        df = pd.concat([df, new_row], ignore_index=True)

        # Save the changes back to the file
        df.to_excel(self.file_name, index=False)




    def refresh_and_save(self):
        df = pd.read_excel(self.file_name)
        self.update_table_data(df)
        self.fill_missing_peaks(df)
        self.sort_table_by_peak(df)
        df.to_excel(self.file_name, index=False)
        self.load_data()

    def update_table_data(self, df):
        for row in range(self.table.rowCount()):
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                value = item.text() if item else ""

                column_name = df.columns[col + 11]
                df.iloc[row, df.columns.get_loc(column_name)] = value

    def fill_missing_peaks(self, df):
        for index, row in df.iterrows():
            if pd.isna(row['High']):
                highest_value = max(row['Tank'], row['Damage'], row['Support'])
                df.at[index, 'High'] = highest_value

    def sort_table_by_peak(self, df):
        # Convert the 'High' column to integers
        df['High'] = pd.to_numeric(df['High'], errors='coerce', downcast='integer')
        df.sort_values(by=['High'], ascending=False, inplace=True, ignore_index=True)








class ImagePasteWidget(QWidget):
    def __init__(self):
        super().__init__()

        self.image_label = QLabel(self)
        self.text_box = QTextEdit(self)
        self.text_box.setPlaceholderText("Paste screenshot here (Ctrl+V or Windows+V)")

        layout = QVBoxLayout(self)
        layout.addWidget(self.image_label)
        layout.addWidget(self.text_box)

        self.text_box.installEventFilter(self)  # Listen for Ctrl+V events

    def eventFilter(self, obj, event):
        # Handle Ctrl+V events on the textbox
        if obj is self.text_box and event.type() == QEvent.KeyPress:
            if event.matches(QKeySequence.Paste):
                clipboard = QApplication.clipboard()
                mime_data = clipboard.mimeData()

                if mime_data.hasImage():
                    qimage = QImage(clipboard.image())
                    pixmap = QPixmap.fromImage(qimage)
                    image = Image.fromqpixmap(pixmap)  # Convert QPixmap to PIL.Image
                    # self.image_label.setPixmap(pixmap)
                    # self.image_label.setScaledContents(True)
                    # self.image_label.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Ignored)
                    # self.image_label.setFixedSize(pixmap.size())
                    # self.parent().adjustSize()

                    self.process_image(image)  # Pass the PIL.Image object
                    return True  # Consume the event to prevent the default paste action

        return super().eventFilter(obj, event)

    def process_image(self, image):
        # Perform OCR on the image
        text = pytesseract.image_to_string(image)
        names = [name.strip() for name in text.split('\n') if name.strip()]

        # Function to remove special characters and keep the longest word
        def clean_and_find_longest_word(text):
            words = re.sub('[^a-zA-Z0-9\s]', '', text).split()
            return max(words, key=len) if words else ""

        # Clean the names
        cleaned_names = [clean_and_find_longest_word(name) for name in names[:10]]

        # Write the cleaned names to the Excel file
        file_name = "rat.xlsx"
        df = pd.read_excel(file_name)
        for idx, clean_name in enumerate(cleaned_names):
            df.iloc[0, idx] = clean_name  # Use 0 instead of 1 for row 1
        df.to_excel(file_name, index=False)

        # Set the new text and resize the window
        self.text_box.clear()
        self.text_box.insertPlainText('\n'.join(cleaned_names))
        # self.parent().adjustSize()


class OutcomeSelectionWidget(QWidget):
    def __init__(self, image_paste_widget):
        super().__init__()

        self.image_paste_widget = image_paste_widget


        self.team1_button = QRadioButton("Team 1 (Blue)")
        self.team2_button = QRadioButton("Team 2 (Red)")
        self.draw_button = QRadioButton("Draw")

        self.layout = QVBoxLayout()

        self.group_box = QGroupBox("Select the team that won:")
        self.group_box_layout = QHBoxLayout()
        self.group_box.setLayout(self.group_box_layout)

        self.group_box_layout.addWidget(self.team1_button)
        self.group_box_layout.addWidget(self.team2_button)
        self.group_box_layout.addWidget(self.draw_button)

        self.submit_button = QPushButton("Submit")
        self.submit_button.clicked.connect(self.submit_outcome)
        self.submit_button.clicked.connect(lambda: elo8.process_data_file(file_name)) # connect the process_data_file method
        self.submit_button.clicked.connect(on_submit) # connect the on_submit method

        self.layout.addWidget(self.group_box)
        self.layout.addWidget(self.submit_button)

        self.setLayout(self.layout)
        self.setFixedHeight(100)

    def submit_outcome(self):
        outcome = 0
        if self.team1_button.isChecked():
            outcome = 1
        elif self.team2_button.isChecked():
            outcome = 2

        # Read the text from the text_box
        text = self.image_paste_widget.text_box.toPlainText()
        names = [name.strip() for name in text.split('\n') if name.strip()]

        # Write the names and outcome to the Excel file
        file_name = "rat.xlsx"
        df = pd.read_excel(file_name)

        # Check if the DataFrame has enough columns, and add new columns if needed
        num_columns_needed = 10
        while df.shape[1] < num_columns_needed:
            df[df.shape[1]] = ""

        for idx, name in enumerate(names[:10]):
            df.iloc[0, idx] = name

        df.iloc[0, 10] = outcome
        df.to_excel(file_name, index=False)

        # Check for missing players
        full_player_list = df['Player'].dropna().unique()
        missing_players = [name for name in names if name not in full_player_list]

        if missing_players:
            return False, missing_players
        else:
            return True, []



def on_submit():
    file_name = "rat.xlsx"
    success, missing_players = outcome_selection_widget.submit_outcome()

    if success:
        show_popup("Ranks updated successfully!")
        excel_table_widget.load_data()
    else:
        missing_players_str = ', '.join(missing_players)
        show_popup(f"Rank update failed! Missing players: {missing_players_str}")


def set_dark_theme(app):
    dark_palette = QPalette()

    # Base colors
    dark_palette.setColor(QPalette.Window, QColor(53, 53, 53))
    dark_palette.setColor(QPalette.WindowText, QColor(255, 255, 255))
    dark_palette.setColor(QPalette.Base, QColor(25, 25, 25))
    dark_palette.setColor(QPalette.AlternateBase, QColor(53, 53, 53))
    dark_palette.setColor(QPalette.ToolTipBase, QColor(255, 255, 255))
    dark_palette.setColor(QPalette.ToolTipText, QColor(255, 255, 255))

    # Text colors
    dark_palette.setColor(QPalette.Text, QColor(255, 255, 255))
    dark_palette.setColor(QPalette.Disabled, QPalette.Text, QColor(127, 127, 127))
    dark_palette.setColor(QPalette.ButtonText, QColor(255, 255, 255))

    # Accent colors (orange)
    dark_palette.setColor(QPalette.Highlight, QColor(255, 165, 0))
    dark_palette.setColor(QPalette.HighlightedText, QColor(0, 0, 0))

    # Button colors
    dark_palette.setColor(QPalette.Button, QColor(53, 53, 53))
    dark_palette.setColor(QPalette.Disabled, QPalette.ButtonText, QColor(127, 127, 127))

    # Window text (title) color
    dark_palette.setColor(QPalette.WindowText, QColor(255, 255, 255))

    app.setPalette(dark_palette)
    app.setStyle("Fusion")




if __name__ == "__main__":
    
    file_name = "rat.xlsx"
    close_excel_file(file_name)

    app = QApplication(sys.argv)
    set_dark_theme(app)  # Add this line to apply the dark theme

    main_layout = QGridLayout()

    image_paste_widget = ImagePasteWidget()
    outcome_selection_widget = OutcomeSelectionWidget(image_paste_widget)
    excel_table_widget = ExcelTableWidget(file_name)

    main_layout.addWidget(image_paste_widget, 0, 0)
    main_layout.addWidget(outcome_selection_widget, 1, 0)
    main_layout.addWidget(excel_table_widget, 0, 1, 2, 1)  # Span 2 rows

    main_widget = QWidget()
    main_widget.setLayout(main_layout)
    main_widget.setWindowTitle("Drexel OW PUGs Ranking Experiment")
    main_widget.setWindowIcon(QIcon('icon.png'))  # Add this line to set the application icon
    main_widget.show()

    sys.exit(app.exec_())
