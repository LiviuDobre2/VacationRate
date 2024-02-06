import sys
import os
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QPushButton, QLabel, QFrame,
                             QSizePolicy, QDateEdit, QRadioButton, QButtonGroup,
                             QFormLayout, QDialog, QListWidget, QListWidgetItem, 
                             QAbstractItemView, QLineEdit)
from PyQt5.QtCore import Qt, QDate, pyqtSignal
from PyQt5.QtGui import QFont
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.dates as mdates  
import numpy as np
import pandas as pd
import os



script_directory = os.path.dirname(os.path.abspath(__file__))
excel_file_name = '02. VacationRateApp_Template_Export.xlsx'
excel_file_path = os.path.join(script_directory, excel_file_name)

df = pd.read_excel(excel_file_path)
unique_departments = df["Departament"].unique().tolist()
unique_project = df["Project Name"].unique().tolist()
unique_employee = df["Employee Name"].unique().tolist()
unique_leave = df["Absence Type"].unique().tolist()

print(df)

# Stylesheet for modern look
stylesheet = """
    QMainWindow {
        background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 rgba(173, 216, 230, 255), stop:1 rgba(255, 255, 255, 255));
    }
    QPushButton {
        background-color: rgb(52, 154, 255);  /* Changed button color here */
        border-radius: 5px;
        color: white;
        padding: 6px;
        margin: 6px;
        font-size: 16px;
    }
    QPushButton:hover {
        background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 rgba(0, 0, 180, 255), stop:1 rgba(0, 0, 230, 255));
    }
    QPushButton:pressed {
        background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 rgba(0, 0, 100, 255), stop:1 rgba(0, 0, 150, 255));
    }
    QLabel, QRadioButton {
        color: #003366;
    }
"""

# Custom Dialog for Period Selection
class PeriodDialog(QDialog):
    
    customPeriodSelected = pyqtSignal(str, str)  # Signal for custom period (start date, end date)
    predefinedPeriodSelected = pyqtSignal(str)  # Signal for predefined period

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Select Period")
        self.setupUI()

    def setupUI(self):
        layout = QVBoxLayout(self)

        # Date pickers for custom period
        formLayout = QFormLayout()
        self.startDateEdit = QDateEdit(self)
        self.startDateEdit.setDate(QDate.currentDate())
        self.endDateEdit = QDateEdit(self)
        self.endDateEdit.setDate(QDate.currentDate())
        formLayout.addRow('Start Date:', self.startDateEdit)
        formLayout.addRow('End Date:', self.endDateEdit)

        # Radio buttons for predefined periods
        self.radioButtons = []
        self.radioGroup = QButtonGroup(self)
        self.radioCustom = QRadioButton("Custom")
        self.radioCustom.setChecked(True)
        self.radioGroup.addButton(self.radioCustom)
        formLayout.addRow(self.radioCustom)

        # Add radio buttons for each month
        for i in range(1, 13):
            month_radio = QRadioButton(QDate.longMonthName(i))
            self.radioGroup.addButton(month_radio)
            self.radioButtons.append(month_radio)
            formLayout.addRow(month_radio)

        # Connect radio button group
        self.radioGroup.buttonClicked.connect(self.onRadioButtonClicked)

        # Submit and Cancel buttons
        buttonsLayout = QHBoxLayout()
        self.submitButton = QPushButton('Submit', self)
        self.cancelButton = QPushButton('Cancel', self)
        buttonsLayout.addWidget(self.submitButton)
        buttonsLayout.addWidget(self.cancelButton)

        # Connect buttons
        self.submitButton.clicked.connect(self.onSubmit)
        self.cancelButton.clicked.connect(self.close)

        # Set the dialog layout
        layout.addLayout(formLayout)
        layout.addLayout(buttonsLayout)
        self.setLayout(layout)  # Set the layout manager

        # Set cursor for buttons
        self.setCursorForButtons()

    def setCursorForButtons(self):
        for button in self.findChildren(QPushButton):
            button.setCursor(Qt.PointingHandCursor)

    def onRadioButtonClicked(self, button):
        if button == self.radioCustom:
            self.startDateEdit.setEnabled(True)
            self.endDateEdit.setEnabled(True)
        else:
            self.startDateEdit.setEnabled(False)
            self.endDateEdit.setEnabled(False)

    def onSubmit(self):
        checkedButton = self.radioGroup.checkedButton()
        if checkedButton:
            if checkedButton == self.radioCustom:
                startDate = self.startDateEdit.date().toString(Qt.ISODate)
                endDate = self.endDateEdit.date().toString(Qt.ISODate)
                print(f"Custom period: {startDate} to {endDate}")
                self.customPeriodSelected.emit(startDate, endDate)  # Emit the custom period signal
            else:
                month_name = checkedButton.text()
                print(f"Predefined month selected: {month_name}")
                self.predefinedPeriodSelected.emit(month_name)
        else:
            print("No period selection made.")
        self.close()

# Custom Dialog for Selection
class SelectionDialog(QDialog):
    selectionMade = pyqtSignal(list, str)
    def __init__(self, options, title, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.options = options
        self.setupUI()

    def setupUI(self):
        layout = QVBoxLayout(self)

        # Search input field
        self.searchInput = QLineEdit(self)
        self.searchInput.setPlaceholderText("Search ")
        self.searchInput.textChanged.connect(self.filterOptions)

        layout.addWidget(self.searchInput)
        self.setLayout(layout)

        # List widget for options
        self.listWidget = QListWidget(self)
        for option in self.options:
            item = QListWidgetItem(option)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Unchecked)
            self.listWidget.addItem(item)
        self.listWidget.setSelectionMode(QAbstractItemView.MultiSelection)
        layout.addWidget(self.listWidget)

        # Select All, Submit, and Cancel buttons
        buttonsLayout = QHBoxLayout()
        self.selectAllButton = QPushButton('Select All', self)
        self.submitButton = QPushButton('Submit', self)
        self.cancelButton = QPushButton('Cancel', self)
        buttonsLayout.addWidget(self.selectAllButton)
        buttonsLayout.addWidget(self.submitButton)
        buttonsLayout.addWidget(self.cancelButton)

        # Connect buttons
        self.selectAllButton.clicked.connect(self.selectAll)
        self.submitButton.clicked.connect(self.onSubmit)
        self.cancelButton.clicked.connect(self.close)

        # Set the dialog layout
        layout.addLayout(buttonsLayout)

        # Set cursor for buttons
        self.setCursorForButtons()

    def setCursorForButtons(self):
        for button in self.findChildren(QPushButton):
            button.setCursor(Qt.PointingHandCursor)

    def filterOptions(self):
        search_text = self.searchInput.text().strip()
        if not search_text:
            # If the search input is empty, show all options
            for index in range(self.listWidget.count()):
                item = self.listWidget.item(index)
                item.setHidden(False)
        else:
            # Filter options based on search text
            for index in range(self.listWidget.count()):
                item = self.listWidget.item(index)
                item_text = item.text()
                item.setHidden(search_text.lower() not in item_text.lower())

    def selectAll(self):
        for index in range(self.listWidget.count()):
            item = self.listWidget.item(index)
            item.setCheckState(Qt.Checked)

    def onSubmit(self):
        selectedOptions = [self.listWidget.item(i).text() for i in range(self.listWidget.count()) 
                           if self.listWidget.item(i).checkState() == Qt.Checked]
        print(f"Selected options: {selectedOptions}")
        self.selectionMade.emit(selectedOptions, self.windowTitle())  # Emit the signal with selected options and the dialog title as category
        self.close()

class ApplicationWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.selections = {
            'department': None,
            'project': None,
            'employee': None,
            'leave': None,
            'period': None
        }
        self.title = 'Vacation Rate App'
        self.currentDialog = None  # Add this line
        self.initUI()


    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(100, 100, 1200, 760)  # Adjust the size as needed
        
        # Create Menu Bar
        menubar = self.menuBar()
        viewMenu = menubar.addMenu('View')
        downloadMenu = menubar.addMenu('Download')
        helpMenu = menubar.addMenu('Help')

        # Main layout
        mainLayout = QHBoxLayout()
        
        # Left panel for filter buttons
        leftPanel = QVBoxLayout()

        # Create buttons
        self.periodButton = QPushButton('Period')
        self.periodButton.clicked.connect(self.showPeriodDialog)
        self.modifyButtonAppearance(self.periodButton)  # Call a function to modify button appearance

        self.departmentButton = QPushButton('Department')
        self.departmentButton.clicked.connect(lambda: self.showSelectionDialog(unique_departments, 'Select Department'))
        self.modifyButtonAppearance(self.departmentButton)  # Call a function to modify button appearance

        self.projectButton = QPushButton('Project')
        self.projectButton.clicked.connect(lambda: self.showSelectionDialog(unique_project, 'Select Project'))
        self.modifyButtonAppearance(self.projectButton)  # Call a function to modify button appearance

        self.employeeButton = QPushButton('Employee')
        self.employeeButton.clicked.connect(lambda: self.showSelectionDialog(unique_employee, 'Select Employee'))
        self.modifyButtonAppearance(self.employeeButton)  # Call a function to modify button appearance

        self.typeOfLeaveButton = QPushButton('Type of Leave')
        self.typeOfLeaveButton.clicked.connect(lambda: self.showSelectionDialog(unique_leave, 'Select Type of Leave'))
        self.modifyButtonAppearance(self.typeOfLeaveButton)  # Call a function to modify button appearance

        # List of buttons
        buttons = [
            self.departmentButton,
            self.projectButton,
            self.periodButton,
            self.employeeButton,
            self.typeOfLeaveButton
        ]

        for button in buttons:
            button.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
            leftPanel.addWidget(button)

        # Right panel for graph and metrics overlay
        rightPanel = QVBoxLayout()
        self.figure = Figure()
        self.canvas = FigureCanvas(self.figure)
        self.canvas.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        rightPanel.addWidget(self.canvas)
        
        # Create and display the histogram
        self.createHistogram()

        # Frame for the metrics overlay
        self.metricsFrame = QFrame(self.canvas)
        metricsLayout = QVBoxLayout(self.metricsFrame)
        self.metricsFrame.setLayout(metricsLayout)
        self.metricsFrame.setFrameStyle(QFrame.StyledPanel)
        self.metricsFrame.setStyleSheet("background-color: rgba(255, 255, 255, 128);")  # Semi-transparent background
        self.metricsFrame.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Maximum)

        # Create a QLabel for each metric and add it to the metrics layout
        metrics = ['Metric 1: XXX', 'Metric 2: XXX', 'Metric 3: XXX']  # Example metrics
        for metric in metrics:
            label = QLabel(metric)
            label.setAlignment(Qt.AlignTop | Qt.AlignRight)
            metricsLayout.addWidget(label)

        # Combine layouts with a stretch factor for the right panel to take up more space
        mainLayout.addLayout(leftPanel, 1)
        mainLayout.addLayout(rightPanel, 4)
        
        # Set the central widget and show the main window
        centralWidget = QWidget()
        centralWidget.setLayout(mainLayout)
        self.setCentralWidget(centralWidget)
        self.show()

        # Position the metrics frame after the UI is shown
        self.repositionMetricsFrame()
    
    def modifyButtonAppearance(self, button):
        button.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
        button.setFont(QFont("Arial", 14))  # Adjust the font and size as needed

        # Apply the filter_buttons_stylesheet to the button
        button.setStyleSheet(stylesheet)

        button.setCursor(Qt.PointingHandCursor)  # Change cursor to a hand when hovering


    def onDialogClosed(self):
     self.currentDialog = None

    
    def showPeriodDialog(self):
        if self.currentDialog is not None:
            self.currentDialog.close()
        self.currentDialog = PeriodDialog(self)
        self.currentDialog.customPeriodSelected.connect(self.handleCustomPeriod)  # Connect to a slot for custom period
        self.currentDialog.predefinedPeriodSelected.connect(self.handlePredefinedPeriod)  # Connect to a slot for predefined period
        self.currentDialog.show()
        self.currentDialog.finished.connect(self.onDialogClosed)

    def handleCustomPeriod(self, startDate, endDate):
        # Convert startDate and endDate from strings to datetime objects
        start_date = pd.to_datetime(startDate)
        end_date = pd.to_datetime(endDate)

        # Store the custom period as a tuple of (start_date, end_date)
        self.selections['period'] = (start_date, end_date)
        print(f"Custom period received in main window: {start_date} to {end_date}")
        print("Current selections:", self.selections)
        self.createHistogram()  


    def handlePredefinedPeriod(self, period):
        current_date = pd.to_datetime("today")

        # Check if the period is a month name
        if period in [QDate.longMonthName(i) for i in range(1, 13)]:
            # Find the numeric month for the selected period
            month_num = [QDate.longMonthName(i) for i in range(1, 13)].index(period) + 1
            
            # Calculate the start and end dates for the selected month
            year = current_date.year  # You can adjust this if you want a different year
            start_date = pd.Timestamp(year, month_num, 1)
            end_date = start_date + pd.offsets.MonthEnd()
        else:
            # Handle custom period or add more conditions if needed
            start_date = None
            end_date = None

        # Store the period as a tuple of (start_date, end_date)
        self.selections['period'] = (start_date, end_date)
        print(f"Predefined period received in main window: {period}")
        print(f"Start date: {start_date}, End date: {end_date}")
        print("Current selections:", self.selections)
        self.createHistogram()  # Update the histogram after applying the new period filter

    
    def handleSelection(self, selections, category):
        # category is one of 'Department', 'Project', 'Employee', or 'Type of Leave'
        category_key_map = {
            'Select Department': 'department',
            'Select Project': 'project',
            'Select Employee': 'employee',
            'Select Type of Leave': 'leave'
        }
        category_key = category_key_map.get(category)
        if category_key:
            self.selections[category_key] = selections
            print(f"{category_key} selected: {selections}")
        print("Current selections:", self.selections)
        self.createHistogram()


    def showSelectionDialog(self, options, title):
        if self.currentDialog is not None:
            self.currentDialog.close()
        self.currentDialog = SelectionDialog(options, title, self)
        self.currentDialog.selectionMade.connect(self.handleSelection)  # Connect the signal to handleSelection
        self.currentDialog.show()
        self.currentDialog.finished.connect(self.onDialogClosed)

    def resizeEvent(self, event):
        QMainWindow.resizeEvent(self, event)
        # Reposition the metrics frame when the main window is resized
        self.repositionMetricsFrame()

    def repositionMetricsFrame(self):
        # Calculate the right position with some padding from the right edge
        right_padding = 10  # Adjust this value to increase or decrease the right padding
        top_padding = 10  # Adjust this value to increase or decrease the top padding
        new_right_position = self.canvas.width() - self.metricsFrame.width() - right_padding
        self.metricsFrame.move(new_right_position, top_padding)

    def createHistogram(self):
        filtered_df = self.filterData()

        # Check if filtered_df is empty or if 'From'/'To' columns have only NaN values
        if filtered_df.empty or filtered_df['From'].isna().all() or filtered_df['To'].isna().all():
            print("No data to display for the selected filters.")
            self.displayMessage("No data to display for the selected filters")
            return  # Exit the method

        # Convert 'From' and 'To' columns to datetime
        filtered_df['From'] = pd.to_datetime(filtered_df['From'])
        filtered_df['To'] = pd.to_datetime(filtered_df['To'])

        # Initialize a DataFrame to hold the counts for each month
        # Initialize a DataFrame to hold the counts for each month
        if not pd.isnull(filtered_df['From'].min()) and not pd.isnull(filtered_df['To'].max()):
            start_date = filtered_df['From'].min().replace(day=1)
            end_date = filtered_df['To'].max().replace(day=1) + pd.offsets.MonthEnd(1)
            date_range = pd.date_range(start=start_date, end=end_date, freq='MS')
            absence_counts = pd.DataFrame(index=date_range, columns=['AbsenceDays'])
            absence_counts['AbsenceDays'] = 0

            # Populate absence_counts for each absence record
            for _, row in filtered_df.iterrows():
                start, end = row['From'], row['To']
                while start <= end:
                    month_start = start.replace(day=1)
                    if month_start in absence_counts.index:
                        next_month = month_start + pd.offsets.MonthBegin(1)
                        days_in_month = (min(end, next_month - pd.Timedelta(days=1)) - start).days + 1
                        absence_counts.loc[month_start, 'AbsenceDays'] += days_in_month
                    start = next_month

            # Debug print to understand the state of absence_counts
            print("Absence Counts:\n", absence_counts)

        # Plot the histogram if there are valid absence days
        if not absence_counts['AbsenceDays'].isna().all():
            self.plotHistogram(absence_counts)
        else:
            print("No absence data for the selected filters.")
            self.displayMessage("No absence data for the selected filters")

    def displayMessage(self, message):
        # Clear the previous figure and display a message
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        ax.text(0.5, 0.5, message, horizontalalignment='center', verticalalignment='center', transform=ax.transAxes)
        self.canvas.draw()

    def plotHistogram(self, absence_counts):
        # Clear the previous figure
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        ax.bar(absence_counts.index, absence_counts['AbsenceDays'], width=20, color='blue', alpha=0.7)

        # Formatting the date on X-axis to make it more readable
        ax.xaxis_date()
        ax.xaxis.set_major_locator(mdates.MonthLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
        ax.figure.autofmt_xdate()

        # Set titles and labels
        ax.set_title('Monthly Absence Counts')
        ax.set_xlabel('Month')
        ax.set_ylabel('Total Absence Days')

        # Draw the canvas with the histogram
        self.canvas.draw()
    
    def filterData(self):
        # Start with the unfiltered DataFrame
        filtered_df = df.copy()
        print("Initial DataFrame:", filtered_df)  # Debug print

        # Apply period filter
        if self.selections['period']:
            start_date, end_date = self.selections['period']
            if start_date and end_date:
                # Filter the data to include only records within the selected period
                filtered_df = filtered_df[(filtered_df['From'] >= start_date) & (filtered_df['To'] <= end_date)]

        # Filter the data based on selections
        for category, selection in self.selections.items():
            if selection:  # If there are selections for this category
                print(f"Applying filter for {category}: {selection}")  # Debug print
                if category == 'department':
                    filtered_df = filtered_df[filtered_df['Departament'].isin(selection)]
                elif category == 'project':
                    filtered_df = filtered_df[filtered_df['Project Name'].isin(selection)]
                elif category == 'employee':
                    filtered_df = filtered_df[filtered_df['Employee Name'].isin(selection)]
                elif category == 'leave':
                    filtered_df = filtered_df[filtered_df['Absence Type'].isin(selection)]
                print(f"DataFrame after {category} filter:", filtered_df)  # Debug print

        # Return the filtered DataFrame
        return filtered_df



# Run the application
app = QApplication(sys.argv)


app.setStyleSheet(stylesheet)
ex = ApplicationWindow()
sys.exit(app.exec_())
