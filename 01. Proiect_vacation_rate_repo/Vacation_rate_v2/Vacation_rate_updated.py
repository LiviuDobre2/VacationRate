import sys
import os
from matplotlib import pyplot as plt
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QPushButton, QLabel,
                             QSizePolicy, QDateEdit, QRadioButton, QButtonGroup,
                             QFormLayout, QDialog, QListWidget, QListWidgetItem, 
                             QAbstractItemView, QLineEdit,QTabWidget, QTableWidget, QTableWidgetItem, QVBoxLayout)
from PyQt5.QtCore import Qt, QDate, pyqtSignal
from PyQt5.QtGui import QFont
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.dates as mdates  
import pandas as pd
import calendar
import datetime
from datetime import date
import holidays
from PyQt5.QtGui import QColor

#Ensure that your script's directory path handling is robust for different environments
script_directory = os.path.dirname(os.path.abspath(__file__))
excel_file_name = 'VacationRate.xlsx'
excel_file_path = os.path.join(script_directory, excel_file_name)
# Load the dataset and extract unique values for filtering options

excel_file_path_new=os.path.join(script_directory, excel_file_name)
# Load sheets into DataFrames
df_absences = pd.read_excel(excel_file_path_new, sheet_name='Absences')
df_projects = pd.read_excel(excel_file_path_new, sheet_name='Projects')
merged_df = pd.merge(df_absences, df_projects.drop(columns=['Engineer Name']), on='Employee ID', how='inner')


# Filter merged DataFrame based on condition
filtered_df = merged_df[(merged_df['From'] >= merged_df['Mission start date']) & (merged_df['From'] <= merged_df['Mission end date'])]

# Add 'Project Name' column to the filtered DataFrame
filtered_df['Project Name'] = filtered_df['Project Name']
# Find employees present in Absences but not in Projects
employees_absences_only = df_absences[~df_absences['Employee ID'].isin(df_projects['Employee ID'])]

# Create a DataFrame for these employees with 'No contract' as End Customer and Project Name
employees_absences_only['End Customer'] = 'No contract'
employees_absences_only['Project Name'] = 'No contract'
# Save modified DataFrame back to the 'Absences' sheet
# Find employees who took days off without a project assigned

employees_wrong_dates = pd.merge(df_absences, df_projects, on='Employee ID', how='inner')
employees_wrong_dates = employees_wrong_dates[~((employees_wrong_dates['From'] >= employees_wrong_dates['Mission start date']) & (employees_wrong_dates['From'] <= employees_wrong_dates['Mission end date']))]


# Add 'No project assigned' as Project Name for these employees
employees_wrong_dates['End Customer'] = 'Intercontract'
employees_wrong_dates['Project Name'] = 'Intercontract'

# Concatenate filtered DataFrame, DataFrame for employees with 'No contract', and DataFrame for employees with 'No project assigned'
final_df = pd.concat([filtered_df, employees_absences_only, employees_wrong_dates], ignore_index=True)
final_df.drop(columns=['Engineer Name'], inplace=True)
final_df.to_excel('vacationRate_modified.xlsx', index=False, sheet_name='Absences')

excel_final_name ='vacationRate_modified.xlsx'
excel_final_path = os.path.join(script_directory, excel_final_name)
df = pd.read_excel(excel_final_path)
unique_departments = df["Departament"].unique().tolist()
unique_project = df["Project Name"].unique().tolist()
unique_employee = df["Employee Name"].unique().tolist()
unique_leave = df["Absence Type"].unique().tolist()
# Define a modern-looking stylesheet for the application
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
# This class allows users to select a reporting period for data visualization


class PeriodDialog(QDialog):
    
    customPeriodSelected = pyqtSignal(str, str)  # Signal for custom period (start date, end date)
    predefinedPeriodSelected = pyqtSignal(str)  # Signal for predefined period

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Select Period")
        self.setupUI()

    def setupUI(self):
        # Layout setup, including date pickers and period selection radio buttons
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

        # Add radio button for "All Available Periods"
        self.radioAllPeriods = QRadioButton("All Available Periods")
        self.radioGroup.addButton(self.radioAllPeriods)
        formLayout.addRow(self.radioAllPeriods)


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
         # Handles submission, differentiating between custom and predefined periods
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

class MonthlyTableWindow(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Monthly Data")
        self.current_month = date.today().month
        self.current_year = 2024  # Assuming the initial year
        self.tableWidget = self.createMonthlyTable()
        layout = QVBoxLayout(self)
        self.resize(1600, 400)  # Adjust the width and height as needed
        self.setLayout(layout)

        previous_month = QDate(self.current_year, self.current_month, 1).addMonths(-1).toString("MMMM")
        next_month = QDate(self.current_year, self.current_month, 1).addMonths(1).toString("MMMM")

        
        # Set the text of the buttons
        self.currentMonthLabel = QLabel()
        current_month = QDate(self.current_year, self.current_month, 1).toString("MMMM")
        self.currentMonthLabel.setText(current_month)
        self.prevButton = QPushButton(previous_month)
        self.nextButton = QPushButton(next_month)
        buttonLayout = QHBoxLayout()
        font = QFont()
        font.setFamily("Arial")
        font.setPointSize(20)  # Adjust the size as needed
        font.setBold(True)     # Make the font bold
        # Apply font to the label
        self.currentMonthLabel.setFont(font)
        self.currentMonthLabel.setAlignment(Qt.AlignCenter)
        buttonLayout.addWidget(self.prevButton)
        buttonLayout.addWidget(self.currentMonthLabel)
        buttonLayout.addWidget(self.nextButton)

        layout.addLayout(buttonLayout)
        layout.addWidget(self.tableWidget)
        # Connect button clicks to slots
        self.prevButton.clicked.connect(self.showPreviousMonth)
        self.nextButton.clicked.connect(self.showNextMonth)
    def getSelection(self):
        return ex.selections
    def get_monthly_data(self, year, month):
        filtered_data= self.getSelection()
        # Filter the DataFrame for the specified year and month
        monthly_data = df[(df['From'].dt.year == year) & ((df['From'].dt.month == month) | (df['To'].dt.month==month))]
        if filtered_data['employee'] is not None:
            monthly_data=monthly_data[monthly_data['Employee Name'].isin(filtered_data['employee'])]
        else: 
            if filtered_data['project'] is not None:
                monthly_data=monthly_data[monthly_data['Project Name'].isin(filtered_data['project'])]
            else:
                if filtered_data['department'] is not None:
                    monthly_data = monthly_data[monthly_data['Departament'].isin(filtered_data['department'])]
        if not monthly_data.empty: 
            month_data = {}
            absence_list = []
            absence_type_list = []  # List to store absence types
            # Iterate over each row in the monthly data
            for index, row in monthly_data.iterrows():
                # Extract relevant information from the row
                
                employee_name = row['Employee Name']
                from_date = row['From'].day  # Extract day from 'From' column
                if row["From"].month<self.current_month:
                    row['From']=pd.Timestamp(year=self.current_year,month=self.current_month,day=1)
                    from_date=row['From'].day
                absence_days=row['To'].day-row['From'].day+1

                if row["To"].month>self.current_month:
                    absence_days=calendar.monthrange(self.current_year,self.current_month)[1]-from_date+1
                absence_type = row['Absence Type']  # Extract absence type from specified collumn
                absence_list.append(absence_days)
                absence_type_list.append(absence_type)
                # If the day is not already in the month_data dictionary, initialize it
                if from_date not in month_data:
                    month_data[from_date] = {}

                # Update the dictionary with employee absence days and absence type for the corresponding day
                month_data[from_date][employee_name] = (absence_days, absence_type)

            return month_data, absence_list, absence_type_list
        else:
            # If no data is found for the specified month and year, return empty dictionaries
            return {}, [], []
    def updateTable(self):
        # Clear the table
        self.tableWidget.clearContents()
        num_days = calendar.monthrange(self.current_year, self.current_month)[1]
        self.tableWidget.setColumnCount(num_days + 1)
        light_blue = QColor(173, 216, 230)

        # Color the first two rows into light blue

        # Populate the table with data for the current month and year
        month_data, absence_list, absence_type_list = self.get_monthly_data(self.current_year, self.current_month)
        month_name = QDate.longMonthName(self.current_month)
        self.tableWidget.setHorizontalHeaderItem(0, QTableWidgetItem("Employee"))
        self.tableWidget.setVerticalHeaderItem(0, QTableWidgetItem(month_name))
        for i in range(1,11):
         self.tableWidget.setVerticalHeaderItem(i, QTableWidgetItem(str(i)))   
        
        for i in range(1, num_days + 1):
            date = datetime.date(self.current_year, self.current_month, i)
            week_number = date.isocalendar()[1]  # Get the ISO week number
            week_str = f"CW{week_number:02d}"  # Format the week number
            self.tableWidget.setHorizontalHeaderItem(i, QTableWidgetItem(week_str))
        for i in range(1, num_days + 1):
            date = datetime.date(self.current_year, self.current_month, i)
            day_number = str(i)
            day_name = date.strftime("%A")  # Get the full name of the day
            header_text = f"{day_number}\n{day_name}"
            header_item = QTableWidgetItem(header_text)
            self.tableWidget.setItem(0, i, header_item)
            self.tableWidget.resizeRowsToContents()
        seen_keys = set()
        ordered_names = []

        for inner_dict in month_data.values():
            for key in inner_dict.keys():
                if key not in seen_keys:
                    ordered_names.append(key)
                    seen_keys.add(key)

        self.tableWidget.setHorizontalHeaderItem(0, QTableWidgetItem("Employee"))
        for row_index, employee_name in enumerate(ordered_names, start=1):
            self.tableWidget.setItem(row_index, 0, QTableWidgetItem(employee_name.strip()))
        # Update the table with the retrieved data
        for day, inner_dict in month_data.items():
            for employee_name, (absence_days, absence_type) in inner_dict.items():
                employee_name = employee_name.strip()  # Trim whitespace from employee name
                # Find the row index corresponding to the employee name
                items = self.tableWidget.findItems(employee_name, Qt.MatchExactly)
                if items:
                    row_index = items[0].row()
                    # Calculate the column index based on the day
                    column_index = day
                    # Set the absence days in the table cell
                    for i in range(int(absence_days)):
                        if column_index + i <= num_days:  # Ensure it doesn't exceed the maximum day
                            color, _ = self.get_absence_type_color(absence_type)
                            cell_item = QTableWidgetItem(self.get_absence_letter(absence_type))
                            cell_item.setBackground(color)
                            self.tableWidget.setItem(row_index, column_index+i, cell_item)
        ro_holidays = self.get_national_holidays(self.current_year,self.current_month)

        light_grey = QColor(211,211,211)  # Adjust the RGB values for the desired shade of gray

        for row_index, employee_name in enumerate(ordered_names, start=1):
            for i in range(1, num_days + 1):
                if QDate(self.current_year, self.current_month, i) in ro_holidays:
                    cell_item = QTableWidgetItem('B')
                    cell_item.setBackground(light_grey)
                    self.tableWidget.setItem(row_index, i, cell_item)

        for col in range(self.tableWidget.columnCount()):
            item = self.tableWidget.item(0, col)
            if item is None:
                item = QTableWidgetItem()
                self.tableWidget.setItem(0, col, item)
            item.setBackground(light_blue)
        self.tableWidget.resizeColumnsToContents()
        self.tableWidget.repaint()

    def get_absence_type_color(self, absence_type):
        # Define colors for each absence type and their lighter versions
        color_map = {
            'Sick leave': (QColor(255, 192, 203), QColor(255, 224, 230)),
            'Maternity leave': (QColor(255, 165, 0), QColor(255, 204, 102)),
            'Annual leave': (QColor(144,238,144), QColor(255, 255, 153)),
            'Wedding leave': (QColor(0, 191, 255), QColor(153, 204, 255)),
            'Unpaid leave': (QColor(220, 20, 60), QColor(255, 179, 179)),
            'Floating day': (QColor(65, 105, 225), QColor(173, 216, 230))
        }
        # Return the corresponding color, or default to white if absence type is not found
        return color_map.get(absence_type, (Qt.white, Qt.white))
    def get_absence_letter(self,absence_type):
        if absence_type == "Sick leave":
            return "S"
        elif absence_type == "Maternity leave":
            return "M"
        elif absence_type == "Annual leave":
            return "H"
        elif absence_type == "Wedding leave":
            return "W"
        elif absence_type == "Unpaid leave":
            return "U"
        elif absence_type == "Floating day":
            return "F"
        else:
            return "X"  # Default to 'X' for unknown absence types
    def get_national_holidays(self, year, month):
        # Create a Holidays object for Romania
        ro_holidays = holidays.RO(years=year)

        holidays_in_month = [date for date in ro_holidays.keys() if date.year == year and date.month == month]
        return holidays_in_month
    def showPreviousMonth(self):

        self.current_month -= 1
        if self.current_month < 1:
            self.current_month = 12
            self.current_year -= 1
        # Modify the previous and next buttons with correct months
        current_month = QDate(self.current_year, self.current_month, 1).toString("MMMM")            
        self.currentMonthLabel.setText(current_month)
        previous_month = QDate(self.current_year, self.current_month, 1).addMonths(-1).toString("MMMM")
        self.prevButton.setText(previous_month)
        next_month = QDate(self.current_year, self.current_month, 1).addMonths(+1).toString("MMMM")
        self.nextButton.setText(next_month)
        # Update the table with data for the next month
        self.updateTable()
        
    def showNextMonth(self):
        self.current_month += 1
        if self.current_month > 12:
            self.current_month = 1
            self.current_year += 1
        # Modify the previous and next buttons with correct months
        previous_month = QDate(self.current_year, self.current_month, 1).addMonths(-1).toString("MMMM")
        self.prevButton.setText(previous_month)
        current_month = QDate(self.current_year, self.current_month, 1).toString("MMMM")            
        self.currentMonthLabel.setText(current_month)
        next_month = QDate(self.current_year, self.current_month, 1).addMonths(1).toString("MMMM")
        self.nextButton.setText(next_month)
        # Update the table with data for the next month
        self.updateTable()

    def createMonthlyTable(self):
        tableWidget = QTableWidget()
        num_days = calendar.monthrange(self.current_year, self.current_month)[1]
        tableWidget.setColumnCount(num_days + 1)
        tableWidget.setRowCount(11)
        headers = [str(day) for day in range(1, 31)]  # Assuming maximum 31 days in a month

        tableWidget.setHorizontalHeaderLabels(headers)
        light_blue = QColor(173, 216, 230)

        # Color the first two rows into light blue

        # Populate the table with data for the current month and year
        month_data, absence_list, absence_type_list = self.get_monthly_data(self.current_year, self.current_month)
        month_name = QDate.longMonthName(self.current_month)
        tableWidget.setHorizontalHeaderItem(0, QTableWidgetItem("Employee"))
        tableWidget.setVerticalHeaderItem(0, QTableWidgetItem(month_name))
        for i in range(1,11):
            tableWidget.setVerticalHeaderItem(i, QTableWidgetItem(str(i)))   
        
        for i in range(1, num_days + 1):
            date = datetime.date(self.current_year, self.current_month, i)
            week_number = date.isocalendar()[1]  # Get the ISO week number
            week_str = f"CW{week_number:02d}"  # Format the week number
            tableWidget.setHorizontalHeaderItem(i, QTableWidgetItem(week_str))
        for i in range(1, num_days + 1):
            date = datetime.date(self.current_year, self.current_month, i)
            day_number = str(i)
            day_name = date.strftime("%A")  # Get the full name of the day
            header_text = f"{day_number}\n{day_name}"
            header_item = QTableWidgetItem(header_text)
            tableWidget.setItem(0, i, header_item)
            tableWidget.resizeRowsToContents()
        seen_keys = set()
        ordered_names = []

        for inner_dict in month_data.values():
            for key in inner_dict.keys():
                if key not in seen_keys:
                    ordered_names.append(key)
                    seen_keys.add(key)
        for row_index, employee_name in enumerate(ordered_names, start=1):
            tableWidget.setItem(row_index,0,QTableWidgetItem(employee_name.strip()))
        # Update the table with the retrieved data
        for day, inner_dict in month_data.items():
            for employee_name, (absence_days, absence_type) in inner_dict.items():
                employee_name = employee_name.strip()  # Trim whitespace from employee name
                # Find the row index corresponding to the employee name
                items = tableWidget.findItems(employee_name, Qt.MatchExactly)
                if items:
                    row_index = items[0].row()
                    # Calculate the column index based on the day
                    column_index = day
                    # Set the absence days in the table cell
                    for i in range(int(absence_days)):
                        if column_index + i <= num_days:  # Ensure it doesn't exceed the maximum day
                            color, _ = self.get_absence_type_color(absence_type)
                            cell_item = QTableWidgetItem(self.get_absence_letter(absence_type))
                            cell_item.setBackground(color)
                            tableWidget.setItem(row_index, column_index+i, cell_item)
        ro_holidays = self.get_national_holidays(self.current_year,self.current_month)

        light_grey = QColor(211,211,211)  # Adjust the RGB values for the desired shade of gray

        for row_index, employee_name in enumerate(ordered_names, start=1):
            for i in range(1, num_days + 1):
                if QDate(self.current_year, self.current_month, i) in ro_holidays:
                    cell_item = QTableWidgetItem('B')
                    cell_item.setBackground(light_grey)
                    tableWidget.setItem(row_index, i, cell_item)

        for col in range(tableWidget.columnCount()):
            item = tableWidget.item(0, col)
            if item is None:
                item = QTableWidgetItem()
                tableWidget.setItem(0, col, item)
            item.setBackground(light_blue)
        tableWidget.resizeColumnsToContents()
        tableWidget.repaint()
        return tableWidget
    
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
        self.currentDialog = None  
        self.initUI()


    def openMonthlyTable(self):
        monthly_table_window = MonthlyTableWindow(self)
        monthly_table_window.exec_()


    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(100, 100, 1200, 760)  # Window size

        # Create Menu Bar
        menubar = self.menuBar()
        viewMenu = menubar.addMenu('View')
        downloadMenu = menubar.addMenu('Download')
        helpMenu = menubar.addMenu('Help')

        # Main layout
        mainLayout = QHBoxLayout()

        # Left panel for filter buttons
        leftPanel = QVBoxLayout()

        # Setup filter buttons
        self.openMonthlyTableButton = QPushButton('Monthly Data')
        self.openMonthlyTableButton.clicked.connect(self.openMonthlyTable)
        self.modifyButtonAppearance(self.openMonthlyTableButton)

        self.periodButton = QPushButton('Period')
        self.periodButton.clicked.connect(self.showPeriodDialog)
        self.departmentButton = QPushButton('Department')
        self.departmentButton.clicked.connect(lambda: self.showSelectionDialog(unique_departments, 'Select Department'))
        self.projectButton = QPushButton('Project')
        self.projectButton.clicked.connect(lambda: self.showSelectionDialog(unique_project, 'Select Project'))
        self.employeeButton = QPushButton('Employee')
        self.employeeButton.clicked.connect(lambda: self.showSelectionDialog(unique_employee, 'Select Employee'))
        self.typeOfLeaveButton = QPushButton('Type of Leave')
        self.typeOfLeaveButton.clicked.connect(lambda: self.showSelectionDialog(unique_leave, 'Select Type of Leave'))

        # Add buttons to the layout
        buttons = [
            self.openMonthlyTableButton, self.periodButton, self.departmentButton,
            self.projectButton, self.employeeButton, self.typeOfLeaveButton
        ]
        for button in buttons:
            self.modifyButtonAppearance(button)
            leftPanel.addWidget(button)

        # Right panel for the graph and metrics, now including a tabbed interface
        rightPanel = QVBoxLayout()
        self.tabWidget = QTabWidget()
        self.histogramTab = QWidget()
        self.doughnutChartTab = QWidget()
        self.tabWidget.addTab(self.histogramTab, "Histogram")
        self.tabWidget.addTab(self.doughnutChartTab, "Doughnut Chart")

        # Setup layouts for the histogram and the doughnut chart tabs
        self.histogramLayout = QVBoxLayout(self.histogramTab)
        self.doughnutChartLayout = QVBoxLayout(self.doughnutChartTab)

        # Setup the histogram tab with a matplotlib canvas
        self.figure = Figure()
        self.canvas = FigureCanvas(self.figure)
        self.histogramLayout.addWidget(self.canvas)

        # No need to setup the doughnut chart here; it will be initialized in createDoughnutChart()

        rightPanel.addWidget(self.tabWidget)

        # Combine left and right panels into the main layout
        mainLayout.addLayout(leftPanel, 1)
        mainLayout.addLayout(rightPanel, 4)

        # Set the central widget and show the main window
        centralWidget = QWidget()
        centralWidget.setLayout(mainLayout)
        self.setCentralWidget(centralWidget)
        self.createHistogram()
        self.createDoughnutChart()
        self.show()



    
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
        self.createDoughnutChart() 


    def handlePredefinedPeriod(self, period):
        global df  # Make sure to use the global dataframe if needed

        if period == "All Available Periods":
            # If "All Available Periods" is selected, remove any period filters
            self.selections['period'] = None
            print("All available data will be displayed.")
        else:
            # Convert 'From' and 'To' columns to datetime if not already done
            df['From'] = pd.to_datetime(df['From'])
            df['To'] = pd.to_datetime(df['To'])

            # Determine the month number for the selected period
            month_num = [QDate.longMonthName(i) for i in range(1, 13)].index(period) + 1
            
            # Filter data to find the years available for the selected month
            years_with_data = df[df['From'].dt.month == month_num]['From'].dt.year.unique()
            if years_with_data.size > 0:
                # Prefer the most recent year with data for the selected month
                selected_year = max(years_with_data)
                start_date = pd.Timestamp(year=selected_year, month=month_num, day=1)
                end_date = start_date + pd.offsets.MonthEnd()
                self.selections['period'] = (start_date, end_date)
                print(f"Selected period: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}")
            else:
                print("No data available for the selected period.")
                self.selections['period'] = None

        # After setting the period, update the histogram
        self.createHistogram()
        self.createDoughnutChart()



    
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
        self.createDoughnutChart()



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


    def determine_bin_size(self):
        """Determine the bin size ('day', 'week', 'month') based on the selected period."""
        if self.selections['period']:
            start_date, end_date = self.selections['period']
            delta = end_date - start_date
            # Use days for comparison, avoiding ambiguous 'M' or 'Y' units
            if delta.days <= 30:  # Approximating 1 month as 30 days
                return 'day'
            elif 30 < delta.days <= 90:  # Approximating 3 months as 90 days
                return 'week'
            else:
                return 'month'
        return 'month'  # Default bin size

    

    def aggregate_data(self, bin_size):
        # Use filterData with only_annual_leave=False for general absence aggregation
        filtered_df = self.filterData(only_annual_leave=False)

        if filtered_df.empty:
            return pd.DataFrame()

        # Convert 'From' and 'To' to datetime format for general aggregation
        filtered_df['From'] = pd.to_datetime(filtered_df['From'])
        filtered_df['To'] = pd.to_datetime(filtered_df['To'])

        #Generate a sequence of dates for each row regardless of absence type, marking each as an absence day
        date_sequences = [pd.date_range(row['From'], row['To']).tolist() for index, row in filtered_df.iterrows()]
        all_dates = [date for sublist in date_sequences for date in sublist]
        all_absences_df = pd.DataFrame(all_dates, columns=['Date'])
        all_absences_df['AbsenceDays'] = 1

        # Aggregate all absence days based on the bin_size
        aggregated_all_df = self.aggregate_absences(all_absences_df, bin_size)

        # Use filterData with only_annual_leave=True for "Annual leave" specific aggregation
        filtered_annual_leave_df = self.filterData(only_annual_leave=True)
        annual_leave_sequences = [pd.date_range(row['From'], row['To']).tolist() for index, row in filtered_annual_leave_df.iterrows()]
        annual_leave_dates = [date for sublist in annual_leave_sequences for date in sublist]
        annual_leave_absences_df = pd.DataFrame(annual_leave_dates, columns=['Date'])
        annual_leave_absences_df['AbsenceDays'] = 1

        # Aggregate "Annual Leave" days based on the bin_size for cumulative calculation
        aggregated_annual_leave_df = self.aggregate_absences(annual_leave_absences_df, bin_size)
        aggregated_annual_leave_df['CumulativeAbsenceDays'] = aggregated_annual_leave_df['AbsenceDays'].cumsum()

        # Total entitlement and used days for 'Annual Leave'
        total_entitlement = filtered_df['Sum of Entitlement'].sum()

        # Calculate cumulative percentage based on "Annual Leave" days taken
        aggregated_annual_leave_df['CumulativePercentage'] = (aggregated_annual_leave_df['CumulativeAbsenceDays'] / total_entitlement) * 100

        # Combine aggregated data for all absences with cumulative data for "Annual Leave"
        aggregated_df = aggregated_all_df.merge(aggregated_annual_leave_df[['Date', 'CumulativePercentage']], on='Date', how='left').fillna(method='ffill')

        return aggregated_df

    def aggregate_absences(self, absences_df, bin_size):
        """Aggregate absence days based on the bin_size."""
        absences_df['Date'] = pd.to_datetime(absences_df['Date'])
        if bin_size == 'day':
            aggregated_df = absences_df.groupby(absences_df['Date'].dt.date)['AbsenceDays'].sum().reset_index(name='AbsenceDays')
        elif bin_size == 'week':
            aggregated_df = absences_df.groupby(absences_df['Date'].dt.to_period('W'))['AbsenceDays'].sum().reset_index(name='AbsenceDays')
            aggregated_df['Date'] = aggregated_df['Date'].apply(lambda x: x.start_time.date())
        elif bin_size == 'month':
            aggregated_df = absences_df.groupby(absences_df['Date'].dt.to_period('M'))['AbsenceDays'].sum().reset_index(name='AbsenceDays')
            aggregated_df['Date'] = aggregated_df['Date'].apply(lambda x: x.start_time.date())
        return aggregated_df



    def createDoughnutChart(self):
        filtered_df = self.filterData()

        # Aggregate the leave types
        leave_counts = filtered_df.groupby('Absence Type')['Att./abs. days'].sum()

        # Labels for the pie chart
        labels = leave_counts.index.tolist()

        # Values for each segment
        x = leave_counts.values.tolist()

        # Create a new Figure for the pie chart
        pie_figure = Figure(figsize=(6, 6))
        pie_canvas = FigureCanvas(pie_figure)
        ax = pie_figure.add_subplot(111)

        # Create the pie chart with specified design
        patches, texts, pcts = ax.pie(
            x, labels=labels, autopct='%.1f%%',
            wedgeprops={'linewidth': 3.0, 'edgecolor': 'white'},
            textprops={'size': 'x-large', 'weight': 'bold', 'color': 'black'},
            startangle=90)

        # Set the title of the pie chart
        ax.set_title('Leave Type Distribution', fontsize=18, color='black')

        # For each wedge, set the corresponding text label color to black (or you can match it to the wedge's face color)
        for i, patch in enumerate(patches):
            texts[i].set_color('black')

        # Set percentage text color to white and texts to bold
        plt.setp(pcts, color='white', weight='bold')
        plt.setp(texts, fontweight=600)

        # Ensure the layout is tight so everything fits without overlap
        pie_figure.tight_layout()

        # Clear the existing layout in the doughnut chart tab and add the new canvas
        for i in reversed(range(self.doughnutChartLayout.count())):
            widget_to_remove = self.doughnutChartLayout.itemAt(i).widget()
            self.doughnutChartLayout.removeWidget(widget_to_remove)
            widget_to_remove.setParent(None)

        self.doughnutChartLayout.addWidget(pie_canvas)


    def createHistogram(self):
        bin_size = self.determine_bin_size()
        aggregated_data = self.aggregate_data(bin_size)

        # Clear the previous figure
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        ax2 = ax.twinx()  # Create a secondary y-axis for cumulative percentages

        if not aggregated_data.empty:
            # Dynamically adjust the bar width based on bin size
            bar_width = {'day': 0.7, 'week': 5, 'month': 20}.get(bin_size, 0.7)
            
            # Plot the histogram with the specified color and add a label for the blue bars
            bars = ax.bar(aggregated_data['Date'], aggregated_data['AbsenceDays'], width=bar_width, color=(173/255, 216/255, 230/255), alpha=0.7, label='Absence Days Taken')

            # Plotting the cumulative percentage and adding a label for the red line
            ax2.plot(aggregated_data['Date'], aggregated_data['CumulativePercentage'], color='red', marker='o', linestyle='-', label='Cumulative Days Taken (%)')
            ax2.set_ylabel('Cumulative Days Taken (%)', color='red')
            ax2.tick_params(axis='y', colors='red')
            
            # Adjust x-axis formatting based on bin size
            if bin_size == 'day':
                ax.xaxis.set_major_locator(mdates.DayLocator())
                ax.xaxis.set_major_formatter(mdates.DateFormatter('%d-%m-%Y'))
            elif bin_size == 'week':
                ax.xaxis.set_major_locator(mdates.WeekdayLocator())
                ax.xaxis.set_major_formatter(mdates.DateFormatter('%U - %Y'))
            elif bin_size == 'month':
                ax.xaxis.set_major_locator(mdates.MonthLocator())
                ax.xaxis.set_major_formatter(mdates.DateFormatter('%B %Y'))
            ax.figure.autofmt_xdate()  # Auto-format date labels

            # Annotate each bin with its value
            for bar in bars:
                height = bar.get_height()
                ax.annotate(f'{int(height)}',
                            xy=(bar.get_x() + bar.get_width() / 2, height),
                            xytext=(0, 3),  # 3 points vertical offset
                            textcoords="offset points",
                            ha='center', va='bottom')

            # Add a grid for better readability
            ax.grid(True, which='both', linestyle='--', linewidth=0.5, color='grey', alpha=0.5)

            # Set titles and labels
            ax.set_title('Absence Counts')
            ax.set_xlabel('Period')
            ax.set_ylabel('Total Absence Days')

            # Combining legends from both axes
            handles, labels = ax.get_legend_handles_labels()
            handles2, labels2 = ax2.get_legend_handles_labels()
            ax2.legend(handles + handles2, labels + labels2, loc='upper left')

        else:
            # Display message if no data
            ax.text(0.5, 0.5, 'No data to display for the selected filters', horizontalalignment='center', verticalalignment='center', transform=ax.transAxes)

        self.figure.subplots_adjust(left=0.07, right=0.95, top=0.95, bottom=0.15)
        self.canvas.draw()



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
    
    def filterData(self, only_annual_leave=False):
        global df
        filtered_df = df.copy()

        # Apply period filter
        if self.selections['period']:
            start_date, end_date = self.selections['period']
            filtered_df = filtered_df[
                ((filtered_df['From'] >= start_date) & (filtered_df['To'] <= end_date)) |
                ((filtered_df['From'] <= end_date) & (filtered_df['To'] >= start_date))
            ]

        if only_annual_leave:
            # Filter for "Annual leave" and apply other selections except for 'leave'
            filtered_df = filtered_df[filtered_df['Absence Type'] == 'Annual leave']
            for category, selection in self.selections.items():
                if selection and category not in ['leave']:  # Skip 'leave' type filtering
                    column_map = {
                        'department': 'Departament',
                        'project': 'Project Name',
                        'employee': 'Employee Name',
                    }
                    if category in column_map:
                        filtered_column = column_map[category]
                        filtered_df = filtered_df[filtered_df[filtered_column].isin(selection)]
        else:
            # Apply filters for all categories
            for category, selection in self.selections.items():
                if selection:
                    column_map = {
                        'department': 'Departament',
                        'project': 'Project Name',
                        'employee': 'Employee Name',
                        'leave': 'Absence Type',
                    }
                    if category in column_map:
                        filtered_column = column_map[category]
                        filtered_df = filtered_df[filtered_df[filtered_column].isin(selection)]

        return filtered_df




# Run the application
app = QApplication(sys.argv)


app.setStyleSheet(stylesheet)
ex = ApplicationWindow()
sys.exit(app.exec_())
