import sys
import os
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QPushButton, QLabel, QFrame,
                             QSizePolicy, QDateEdit, QRadioButton, QButtonGroup,
                             QFormLayout, QDialog, QListWidget, QListWidgetItem, 
                             QAbstractItemView)
from PyQt5.QtCore import Qt, QDate
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import pandas as pd
import os


script_directory = os.path.dirname(os.path.abspath(__file__))
excel_file_name = '02. VacationRateApp_Template_Export.xlsx'
excel_file_path = os.path.join(script_directory, excel_file_name)

excel_row = pd.read_excel(excel_file_path)

# Stylesheet for modern look
stylesheet = """
    QMainWindow {
        background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 rgba(173, 216, 230, 255), stop:1 rgba(255, 255, 255, 255));
    }
    QPushButton {
        background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 rgba(0, 0, 128, 255), stop:1 rgba(0, 0, 255, 255));
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

    def onSubmit(self):
        # Handle period selection logic here
        # Get dates if custom period is selected
        if not self.radioGroup.checkedButton():
            startDate = self.startDateEdit.date()
            endDate = self.endDateEdit.date()
            print(f"Custom period: {startDate.toString(Qt.ISODate)} to {endDate.toString(Qt.ISODate)}")
        else:
            # Or get predefined period
            print(f"Predefined period: {self.radioGroup.checkedButton().text()}")
        self.close()

# Custom Dialog for Selection
class SelectionDialog(QDialog):
    def __init__(self, options, title, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setGeometry(100, 100, 300, 200)
        self.options = options
        self.setupUI()

    def setupUI(self):
        layout = QVBoxLayout(self)

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

    def selectAll(self):
        for index in range(self.listWidget.count()):
            item = self.listWidget.item(index)
            item.setCheckState(Qt.Checked)

    def onSubmit(self):
        # Handle option selection logic here
        selectedOptions = [self.listWidget.item(i).text() for i in range(self.listWidget.count()) 
                           if self.listWidget.item(i).checkState() == Qt.Checked]
        print(f"Selected options: {selectedOptions}")
        self.close()

class ApplicationWindow(QMainWindow):
    def __init__(self):
        super().__init__()
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

        self.departmentButton = QPushButton("Department")
        self.departmentButton.clicked.connect(lambda: self.showSelectionDialog(map(str,excel_row["Departament"].unique()), 'Select Department'))

        self.projectButton = QPushButton('Project')
        self.projectButton.clicked.connect(lambda: self.showSelectionDialog(map(str,excel_row["Project Name"].unique()), 'Select Project'))

        self.employeeButton = QPushButton('Employee')
        self.employeeButton.clicked.connect(lambda: self.showSelectionDialog(map(str,excel_row["Employee Name"].unique()), 'Select Employee'))

        self.typeOfLeaveButton = QPushButton('Absence Type')
        self.typeOfLeaveButton.clicked.connect(lambda: self.showSelectionDialog(map(str,excel_row["Absence Type"].unique()), 'Select Type of Leave'))

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
        
        # Plotting a placeholder graph
        ax = self.figure.add_subplot(111)  # Adding a subplot
        ax.hist([1, 2, 2, 3, 4, 5, 5, 6, 7, 8, 8, 9], bins=8, color='blue', alpha=0.7)  # Example histogram data
        ax.set_title('Histogram Placeholder')
        ax.set_xlabel('X-axis Label')
        ax.set_ylabel('Y-axis Label')
        self.canvas.draw()  # Draw the canvas with the histogram

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
    
    def onDialogClosed(self):
     self.currentDialog = None

    
    def showPeriodDialog(self):
        if self.currentDialog is not None:
            self.currentDialog.close()
        self.currentDialog = PeriodDialog(self)
        self.currentDialog.show()
        self.currentDialog.finished.connect(self.onDialogClosed)

    def showSelectionDialog(self, options, title):
        if self.currentDialog is not None:
            self.currentDialog.close()
        self.currentDialog = SelectionDialog(options, title, self)
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

# Run the application
app = QApplication(sys.argv)


app.setStyleSheet(stylesheet)
ex = ApplicationWindow()
sys.exit(app.exec_())
