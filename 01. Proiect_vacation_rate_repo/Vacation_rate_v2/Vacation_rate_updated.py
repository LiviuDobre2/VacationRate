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
        self.setGeometry(100, 100, 300, 200)
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
        self.radioGroup = QButtonGroup(self)
        self.radioLastDay = QRadioButton("Last Day")
        self.radioLastWeek = QRadioButton("Last Week")
        self.radioLastMonth = QRadioButton("Last Month")
        self.radioLastQuarter = QRadioButton("Last Quarter")
        self.radioAllPeriod = QRadioButton("All Available Period")
        self.radioGroup.addButton(self.radioLastDay)
        self.radioGroup.addButton(self.radioLastWeek)
        self.radioGroup.addButton(self.radioLastMonth)
        self.radioGroup.addButton(self.radioLastQuarter)
        self.radioGroup.addButton(self.radioAllPeriod)
        formLayout.addRow(self.radioLastDay)
        formLayout.addRow(self.radioLastWeek)
        formLayout.addRow(self.radioLastMonth)
        formLayout.addRow(self.radioLastQuarter)
        formLayout.addRow(self.radioAllPeriod)

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

        # Set cursor for buttons
        self.setCursorForButtons()

    def setCursorForButtons(self):
        for button in self.findChildren(QPushButton):
            button.setCursor(Qt.PointingHandCursor)
            
    def onSubmit(self):
        if not self.radioGroup.checkedButton():
            startDate = self.startDateEdit.date().toString(Qt.ISODate)
            endDate = self.endDateEdit.date().toString(Qt.ISODate)
            print(f"Custom period: {startDate} to {endDate}")
            self.customPeriodSelected.emit(startDate, endDate)  # Emit the custom period signal
        else:
            predefinedPeriod = self.radioGroup.checkedButton().text()
            print(f"Predefined period: {predefinedPeriod}")
            self.predefinedPeriodSelected.emit(predefinedPeriod)  # Emit the predefined period signal
        self.close()


# Custom Dialog for Selection
class SelectionDialog(QDialog):
    selectionMade = pyqtSignal(list, str)
    def __init__(self, options, title, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setGeometry(100, 100, 300, 200)
        self.options = options
        self.setupUI()

    def setupUI(self):
        layout = QVBoxLayout(self)

        # Search input field
        self.searchInput = QLineEdit(self)
        self.searchInput.setPlaceholderText("Search ")
        self.searchInput.textChanged.connect(self.filterOptions)
        layout.addWidget(self.searchInput)

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
        metrics =self.selections  # Example metrics
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
        # Store the custom period as a tuple of (startDate, endDate)
        self.selections['period'] = (startDate, endDate)
        print(f"Custom period received in main window: {startDate} to {endDate}")
        print("Current selections:", self.selections)
        self.updateCustomMetricsOverlay(startDate,endDate)


    def handlePredefinedPeriod(self, period):
        # Handle the predefined period here
        self.selections['period'] = period
        print(f"Predefined period received in main window: {period}")
        print("Current selections:", self.selections)
        self.updateMetricsOverlay()
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
 
        
    def updateMetricsOverlay(self):
        # Clear the previous content of the metrics overlay

        for i in reversed(range(self.metricsFrame.layout().count())):
            item = self.metricsFrame.layout().itemAt(i)
            widget = item.widget()
            if widget:
                widget.setParent(None)

        # Display the selected period in the metrics overlay
        selected_period2 = self.getFilteredDate()
        mask=self.selections
        del mask['period']
        for filter in mask.values():
            if (filter != None):
                filtered_row=selected_period2[selected_period2["Project Name"]==filter[0]]
                print("total")
                print(filtered_row["Sum of Entitlement for 2023"].sum())
                print("luate")
                total=filtered_row["Sum of Entitlement for 2023"].sum()
                luate=filtered_row["Att./abs. days"].sum()
                print(filtered_row["Att./abs. days"].sum())
                metric1=luate*100/total
                period_text = "Days Taken %d%% %s" %(metric1,"total")
                print(metric1)

        
                period_label = QLabel(period_text)
                period_label.setAlignment(Qt.AlignTop | Qt.AlignRight)
                period_label.setWordWrap(True)  # Enable word wrap

                # Set minimum and maximum sizes for the label
                period_label.setMinimumSize(0, period_label.sizeHint().height())
                period_label.setMaximumSize(period_label.sizeHint().width(), period_label.sizeHint().height())

                self.metricsFrame.layout().addWidget(period_label)

        # Reposition the metrics frame
        self.repositionMetricsFrame()
        
    def updateCustomMetricsOverlay(self,start_date,end_date):
       # Clear the previous content of the metrics overlay

        for i in reversed(range(self.metricsFrame.layout().count())):
            item = self.metricsFrame.layout().itemAt(i)
            widget = item.widget()
            if widget:
                widget.setParent(None)

       # Display the selected period in the metrics overlay
        selected_period2 = self.getFilteredDate()
        print(selected_period2)
        period_text = f"Selected Period: {start_date} to {end_date}"
        period_label = QLabel(period_text)
        period_label.setAlignment(Qt.AlignTop | Qt.AlignRight)
        period_label.setWordWrap(True)  # Enable word wrap
        # Set minimum and maximum sizes for the label
        period_label.setMinimumSize(0, period_label.sizeHint().height())
        period_label.setMaximumSize(period_label.sizeHint().width(), period_label.sizeHint().height())

        self.metricsFrame.layout().addWidget(period_label)
        # Reposition the metrics frame
        self.repositionMetricsFrame()

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
        # Clear the previous figure
        self.figure.clear()
        
        # Create an example histogram
        ax = self.figure.add_subplot(111)  # Adding a subplot
        data = [1, 2, 2, 3, 4, 5, 5, 6, 7, 8, 8, 9]  # Example data
        ax.hist(data, bins=8, color='blue', alpha=0.7)
        ax.set_title('Example Histogram')
        ax.set_xlabel('X-axis Label')
        ax.set_ylabel('Y-axis Label')
        
        # Draw the canvas with the histogram
        self.canvas.draw()

    def getFilteredDate(self):
        # Start with the full dataset
        filtered_data = df.copy()

        # Apply other selections as filters
        # Now apply the period filter
        period_selection = self.selections['period']
        if period_selection:
            if isinstance(period_selection, tuple):
                # It's a custom period (startDate, endDate)
                start_date, end_date = pd.to_datetime(period_selection[0]), pd.to_datetime(period_selection[1])
                filtered_data = filtered_data[(filtered_data['From'] >= start_date) & (filtered_data['To'] <= end_date)]
            elif period_selection == "All Available Period": 
                # No date filtering needed, all data is already included
                pass
            else:
                # It's a predefined period like "Last Month"
                start_date, end_date = self.get_predefined_period_dates(period_selection)
                filtered_data = filtered_data[(filtered_data['From'] >= start_date) & (filtered_data['To'] <= end_date)]

        return filtered_data
    
    def get_predefined_period_dates(self, period):
        today = pd.to_datetime('today').normalize()  # Normalize to remove time
        if period == "Last Day":
            end_date = today - pd.Timedelta(days=1)
        if period == "Last Week":
            end_date = today - pd.Timedelta(days=1)  # Exclude today
            start_date = end_date - pd.Timedelta(days=6)  # 7 days including the end date
        if period == "Last Month":
            first_day_this_month = today.replace(day=1)  # First day of the current month
            last_day_last_month = first_day_this_month - pd.Timedelta(days=1)  # Last day of the previous month
            start_date = last_day_last_month.replace(day=1)  # First day of the previous month
            end_date = last_day_last_month
        if period == "Last Quarter":
            current_quarter = ((today.month - 1) // 3) + 1
            first_month_last_quarter = (current_quarter - 2) * 3 + 1
            year_adjustment = today.year if first_month_last_quarter > 0 else today.year - 1
            month_adjustment = first_month_last_quarter if first_month_last_quarter > 0 else first_month_last_quarter + 12
            start_date = pd.Timestamp(year=year_adjustment, month=month_adjustment, day=1)
            end_date = (start_date + pd.DateOffset(months=3)) - pd.Timedelta(days=1)
        return start_date, end_date


# Run the application
app = QApplication(sys.argv)


app.setStyleSheet(stylesheet)
ex = ApplicationWindow()
sys.exit(app.exec_())
