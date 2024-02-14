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

    

    def aggregate_data(self, filtered_df, bin_size):
        """Aggregate data by expanding each entry to cover all days between 'From' and 'To', then summing up absence days according to bin size."""
        
        # First, ensure 'From' and 'To' are in datetime format
        filtered_df['From'] = pd.to_datetime(filtered_df['From'])
        filtered_df['To'] = pd.to_datetime(filtered_df['To'])

        # Expand each absence record to include each day within its range
        all_dates = []
        for _, row in filtered_df.iterrows():
            num_days = (row['To'] - row['From']).days + 1
            daily_records = pd.date_range(start=row['From'], periods=num_days, freq='D')
            for day in daily_records:
                all_dates.append({'Date': day, 'AbsenceDays': 1})  # Assuming 1 absence day per day in range
        
        # Convert the list of dictionaries into a DataFrame
        expanded_df = pd.DataFrame(all_dates)

        if 'Date' not in expanded_df.columns:
            print("Error: 'Date' column not found in expanded_df. Available columns:", expanded_df.columns)
            return pd.DataFrame()  # Return an empty DataFrame or handle the error as appropriate

        # Now, aggregate this expanded data frame by the selected bin size
        # Assuming 'AbsenceDays' is the column you want to sum
        if bin_size == 'day':
            aggregated_df = expanded_df.groupby(expanded_df['Date'].dt.date)['AbsenceDays'].sum().reset_index()
        elif bin_size == 'week':
            # Group by week, summing only the 'AbsenceDays' column
            aggregated_df = expanded_df.groupby(expanded_df['Date'].dt.to_period('W'))['AbsenceDays'].sum().reset_index()
            aggregated_df['Date'] = aggregated_df['Date'].apply(lambda x: x.start_time)
        elif bin_size == 'month':
            # Group by month, summing only the 'AbsenceDays' column
            aggregated_df = expanded_df.groupby(expanded_df['Date'].dt.to_period('M'))['AbsenceDays'].sum().reset_index()
            aggregated_df['Date'] = aggregated_df['Date'].apply(lambda x: x.start_time)

        return aggregated_df


    def createHistogram(self):
        filtered_df = self.filterData()
        bin_size = self.determine_bin_size()
        aggregated_data = self.aggregate_data(filtered_df, bin_size)

        # Clear the previous figure
        self.figure.clear()
        ax = self.figure.add_subplot(111)

        if not aggregated_data.empty:
            # Dynamically adjust the bar width based on bin size
            bar_width = {'day': 0.7, 'week': 5, 'month': 20}.get(bin_size, 0.7)
            
            # Plot the histogram
            ax.bar(aggregated_data['Date'], aggregated_data['AbsenceDays'], width=bar_width, color='blue', alpha=0.7)

            # Adjust x-axis formatting based on bin size
            if bin_size == 'day':
                ax.xaxis.set_major_locator(mdates.DayLocator())
                ax.xaxis.set_major_formatter(mdates.DateFormatter('%d-%m-%Y'))
            elif bin_size == 'week':
                # Adjust for a more descriptive week format, e.g., start of week
                ax.xaxis.set_major_locator(mdates.WeekdayLocator())
                ax.xaxis.set_major_formatter(mdates.DateFormatter('%U - %Y'))
            elif bin_size == 'month':
                ax.xaxis.set_major_locator(mdates.MonthLocator())
                ax.xaxis.set_major_formatter(mdates.DateFormatter('%B %Y'))

            ax.figure.autofmt_xdate()  # Auto-format date labels

            # Set titles and labels
            ax.set_title('Absence Counts')
            ax.set_xlabel('Period')
            ax.set_ylabel('Total Absence Days')
        else:
            # Display message if no data
            ax.text(0.5, 0.5, 'No data to display for the selected filters', horizontalalignment='center', verticalalignment='center', transform=ax.transAxes)

        # Draw the canvas
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
    
    def filterData(self):
        # Start with the original DataFrame
        global df  # Ensure you're using the global dataframe or replace with your dataframe variable
        filtered_df = df.copy()

        # Check if a period has been selected
        if self.selections['period']:
            start_date, end_date = self.selections['period']
            # Apply period filter only if it's not None
            filtered_df = filtered_df[
                (filtered_df['From'] >= start_date) & (filtered_df['To'] <= end_date) |
                (filtered_df['From'] <= end_date) & (filtered_df['To'] >= start_date)
            ]

        # Apply filters for other categories (department, project, employee, leave)
        for category, selection in self.selections.items():
            if selection and category in ['department', 'project', 'employee', 'leave']:
                column_map = {
                    'department': 'Departament',
                    'project': 'Project Name',
                    'employee': 'Employee Name',
                    'leave': 'Absence Type'
                }
                filtered_column = column_map[category]
                filtered_df = filtered_df[filtered_df[filtered_column].isin(selection)]

        return filtered_df



# Run the application
app = QApplication(sys.argv)


app.setStyleSheet(stylesheet)
ex = ApplicationWindow()
sys.exit(app.exec_())
