import sys
import pyodbc
import os
from datetime import datetime, timedelta
from PySide2.QtWidgets import QApplication, QVBoxLayout, QPushButton, QWidget, QLabel, QTableWidget, \
    QAbstractItemView
from PySide2.QtWidgets import QMessageBox, QTableWidgetItem, QGridLayout, QLineEdit, QComboBox, QDateEdit
from PySide2.QtCore import Qt, QDateTime, QTime, QDate
from PySide2.QtGui import QColor, QFont

'''
def get_user_name():
    user_table = {
        't545751': 'Jan Fasola',
        't721851': 'Jakub Wlóka',
        't594826': 'Tsering Dukatsang',
        't659628': 'Gianmarco Patazzi',
        't127675': 'Priska Dörfinger',
        't114659': 'Angelo Peréz',
        't609321': 'Verona Nehertepe',
        't595527': 'Liri Osman',
        't657331': 'Salvi De Biassi'
    }


def new_id(self):
    t_number = os.getlogin()
    return user_table.get(t_number, "New user")
'''

class MyWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.currentrow = None
        self.hold_start = None
        self.hold_start_time = 0
        self.total_elapsed_time = 0
        self.initUI()
        self.user_name = 'Jakub Włóka'
        self.initDB()
        self.setWindowTitle("Expense Management")
        self.setGeometry(100, 100, 900, 900)

    # Set UI Layout
    def initUI(self):
        mainLayout = QVBoxLayout()

        # ID Display
        self.id_label = QLabel("Selected ID: None")

        # Add new case button
        gridLayout = QGridLayout()
        # Add new case button
        self.newCaseButton = QPushButton("Create new Case")
        gridLayout.addWidget(self.newCaseButton, 0, 1)
        self.newCaseButton.clicked.connect(self.createNewCase)

        # Add SAVE button
        self.saveButton = QPushButton('Save Case')
        gridLayout.addWidget(self.saveButton, 0, 2)
        self.saveButton.clicked.connect(self.save)

        # Add Delete button
        self.deleteButton = QPushButton('Delete Case')
        gridLayout.addWidget(self.deleteButton, 0, 3)
        self.deleteButton.clicked.connect(self.delete)

        mainLayout.addLayout(gridLayout)

        # Receive Date Btn and Lbl

        gridLayout = QGridLayout()

        # self.dateButton = QPushButton("Receive Date")
        # self.dateLabel = QLineEdit()
        # self.dateButton.clicked.connect(self.showDate)


        self.receive_label = QLabel('Receive Date')
        self.receive_date = QDateEdit(self)
        self.receive_date.setCalendarPopup(True)
        self.receive_date.setDate(QDate.currentDate())
        gridLayout.addWidget(self.receive_label, 1, 0)
        gridLayout.addWidget(self.receive_date, 1, 1)
        # self.receive_date.setDisplayFormat('yyyy-MM-dd')

        # Start Date Btn and Lbl
        self.startButton = QPushButton('Start Time')
        self.startLabel = QLineEdit()
        self.startButton.clicked.connect(self.StartDate)
        gridLayout.addWidget(self.startButton, 2, 0)
        gridLayout.addWidget(self.startLabel, 2, 1)

        # Hold Time
        self.holdButton = QPushButton('Hold Time')
        self.holdLabel = QLineEdit()
        self.holdButton.clicked.connect(self.holdTime)
        gridLayout.addWidget(self.holdButton, 3, 0)
        gridLayout.addWidget(self.holdLabel, 3, 1)

        # Finish Time Btn and Lbl
        self.finishButton = QPushButton('Finish Time')
        self.finishLabel = QLineEdit()
        # self.finishButton.clicked.connect(self.finishTime)
        # gridLayout.addWidget(self.finishButton, 4, 0)
        # gridLayout.addWidget(self.finishLabel, 4, 1)
        # Finish Time Btn and Lbl
        self.finishButton = QPushButton('Finish Time')
        self.finishLabel = QLineEdit()
        self.finishButton.clicked.connect(self.FinishDate)
        gridLayout.addWidget(self.finishButton, 4, 0)
        gridLayout.addWidget(self.finishLabel, 4, 1)

        # Total Time
        self.total_time_label = QLabel('Total Time:')
        self.total_time_txt = QLineEdit()
        gridLayout.addWidget(self.total_time_label, 5, 0)
        gridLayout.addWidget(self.total_time_txt, 5, 1)

        # Entered By
        self.entered_by_label = QPushButton('Case entered by:')
        self.entered_by_txt = QComboBox()
        self.entered_by_txt.addItems(['None', 'Nyima Dohnchochog', 'Patricia Bakonyi', 'Sofija Kislitscka',
                                      'Ailisija Kislitscka', 'Marianne Frei', 'Kristine Irieu', 'Sonam Ny',
                                      'Gianmarco Palazzi', 'Jakub Wróka'])
        gridLayout.addWidget(self.entered_by_label, 6, 0)
        gridLayout.addWidget(self.entered_by_txt, 6, 1)
        self.entered_by_label.clicked.connect(self.your_data)

        # Status
        self.status_label = QLabel('Status:')
        # self.status_txt = QLineEdit()
        self.status_txt = QComboBox()
        self.status_txt.addItems(['Open', 'Closed', 'Advised', 'None'])
        gridLayout.addWidget(self.status_label, 7, 0)
        gridLayout.addWidget(self.status_txt, 7, 1)
        # SECOND COLUMN

        # Requestor
        self.requestor_label = QLabel('Requestor CA/QA:')
        self.requestor_txt = QLineEdit()
        gridLayout.addWidget(self.requestor_label,1, 2)
        gridLayout.addWidget(self.requestor_txt, 1, 3)

        # Requestor GPN
        self.requestorG_label = QLabel('Requestor GPN:')
        self.requestorG_txt = QLineEdit()
        gridLayout.addWidget(self.requestorG_label,2, 2)
        gridLayout.addWidget(self.requestorG_txt, 2, 3)

        # No. of Invoices:
        self.invoice_label = QLabel('No of Invoices:')
        self.invoice_txt = QLineEdit()
        gridLayout.addWidget(self.invoice_label, 3, 2)
        gridLayout.addWidget(self.invoice_txt, 3, 3)

        # Handover
        self.handover_label = QLabel('Handover Type:')
        self.handover_txt = QComboBox()
        self.handover_txt.addItems(['None', 'Genesys', 'Other channel'])
        gridLayout.addWidget(self.handover_label, 4, 2)
        gridLayout.addWidget(self.handover_txt, 4, 3)
        # Trip ID
        self.trip_label = QLabel('Case ID:')
        self.trip_txt = QLineEdit()
        gridLayout.addWidget(self.trip_label, 5, 2)
        gridLayout.addWidget(self.trip_txt, 5, 3)
        mainLayout.addLayout(gridLayout)

        # Report Name
        self.report_label = QLabel('Report Name:')
        self.report_txt = QLineEdit()
        gridLayout.addWidget(self.report_label, 6, 2)
        gridLayout.addWidget(self.report_txt, 6, 3)

        # Nr. of iteration
        self.line_label = QLabel('Total entered line items:')
        self.line_txt = QLineEdit()
        gridLayout.addWidget(self.line_label, 7, 2)
        gridLayout.addWidget(self.line_txt, 7, 3)

        # Third COLUMN

        # Nr. of iter
        self.nr_of_iter_label = QLabel('Number of Iterations:')
        self.nr_of_iter_txt = QComboBox()
        # Nr. of iter
        self.nr_of_iter_label = QLabel('Number of Iterations:')
        self.nr_of_iter_txt = QComboBox()
        self.nr_of_iter_txt.addItems([
            'No iteration',
            'Aged Transactions',
            'Iteration - Receipts not handed in',
            'Iteration - Messy submission',
            'Iteration - Missing info on receipts',
            'Iteration - Not due to front',
            'None'
        ])
        gridLayout.addWidget(self.nr_of_iter_label, 1, 4)
        gridLayout.addWidget(self.nr_of_iter_txt, 1, 5)

        # Reason of pending
        self.reason_label = QLabel('Reason of pending:')
        self.reason_txt = QLineEdit()
        gridLayout.addWidget(self.reason_label, 2, 4)
        gridLayout.addWidget(self.reason_txt, 2, 5)

        # Action
        self.action_label = QLabel('Action:')
        self.action_txt = QLineEdit()
        gridLayout.addWidget(self.action_label, 3, 4)
        gridLayout.addWidget(self.action_txt, 3, 5)
        # Filter cases by open
        self.filteroff_label = QLabel('Filter cases by open:')
        self.filteroff_button = QPushButton('Filter Off')
        gridLayout.addWidget(self.filteroff_label, 4, 4)
        gridLayout.addWidget(self.filteroff_button, 4, 5)
        self.filteroff_button.clicked.connect(self.filter_off)

        # Filter cases off
        self.filter_label = QLabel('Filter cases off:')
        self.filter_button = QPushButton('Filter Cases')
        gridLayout.addWidget(self.filter_label, 5, 4)
        gridLayout.addWidget(self.filter_button, 5, 5)
        self.filter_button.clicked.connect(self.filter)

        mainLayout.addLayout(gridLayout)
        mainLayout.addWidget(self.id_label)

        # Table for DB
        self.tableWidget = QTableWidget()
        self.tableWidget.setColumnCount(9)
        col_names = ['ID', 'Start Date', 'Finish Date', 'Event Date', 'Handover Type', 'Requestor CA/QA',
                     'Report name', 'Case Entered', 'Requestor GPN', 'Rank', 'No of Invoices', 'Number of credit card line items', 'No of cash expenses',
                      'Number of Iterations', 'Total entered line items', 'Case Closed', 'Case Pending',
                      'Time spent total in min', 'Receipt email', 'Confirmation email', 'Reason if pending',
                      'Date of Action', 'Total Hold Time', 'Hold Continue', 'Hold Time', 'Action', 'Current Status',
                      'Received Date', 'Total Hold Time', 'Hold Continue']
        self.tableWidget.setHorizontalHeaderLabels(col_names)

        # Set Bold FONT
        bold_font = QFont()
        bold_font.setBold(True)
        self.tableWidget.horizontalHeader().setFont(bold_font)

        # turn off edit mode for table
        self.tableWidget.setEditTriggers(QTableWidget.NoEditTriggers)

        # Double_click functionality
        self.tableWidget.itemDoubleClicked.connect(self.doubleClick)

        # Sorting
        self.tableWidget.setSortingEnabled(True)

        # adding table
        mainLayout.addWidget(self.tableWidget)

        # adding to main layout
        self.setLayout(mainLayout)
        # adding to main layout
        self.setLayout(mainLayout)

        self.tableWidget.setAlternatingRowColors(True)
        self.tableWidget.setStyleSheet("alternate-background-color: white; background-color: lightgray;")

        # Calculating total time
        self.startLabel.textChanged.connect(self.calc_Total_time)
        self.holdLabel.textChanged.connect(self.calc_Total_time)
        self.finishLabel.textChanged.connect(self.calc_Total_time)


    # establish connection
    def initDB(self):
        self.conn = pyodbc.connect(
            r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\Jakub\Desktop\Projekt WSB\TestDB.accdb"
        )
        self.cursor = self.conn.cursor()
        self.loadtableData()


    def loadtableData(self):
        try:
            sqlQuery = "SELECT * FROM Expense_Management ORDER BY ID DESC"
            self.cursor.execute(sqlQuery)
            rows = self.cursor.fetchall()


            self.tableWidget.clearContents()
            self.tableWidget.setRowCount(len(rows))

            for row_index, row_data in enumerate(rows):
                for column_index, col_data in enumerate(row_data):
                    cell = QTableWidgetItem(str(col_data))
                    cell.setFlags(cell.flags() ^ Qt.ItemIsEditable)
                    self.tableWidget.setItem(row_index, column_index, cell)

        except Exception as e:
            print(f"Error: {e}")


    def your_data(self):
        try:
            case = self.entered_by_txt.currentText()
            print(case)

            sqlQuery = "SELECT * FROM Expense_Management WHERE [Case entered by]=? ORDER BY ID DESC"
            self.cursor.execute(sqlQuery, (case,))
            rows = self.cursor.fetchall()

            self.tableWidget.clearContents()
            self.tableWidget.setRowCount(len(rows))

            for row_index, row_data in enumerate(rows):
                for column_index, col_data in enumerate(row_data):
                    if col_data is None or col_data == "" or col_data == " " :
                        cell=QTableWidgetItem(" ")
                    else:
                        cell=QTableWidgetItem(str(col_data))

                    cell.setFlags(cell.flags() ^ Qt.ItemIsEditable)
                    self.tableWidget.setItem(row_index, column_index, cell)

        except Exception as e:
            print(f"Error: {e}")


    def StartDate(self):
        startdate = QDateTime.currentDateTime().toString('dd/MM/yyyy HH:mm:ss')
        self.startLabel.setText(startdate)
        self.calc_Total_time()


    def holdTime(self):
        if self.hold_start is None:
            self.hold_start = datetime.now()
            self.holdLabel.setText("Timing...")
        else:
            hold_end = datetime.now()
            elapsedtime = hold_end - self.hold_start
            elapsed_min = elapsedtime.total_seconds() / 60
            self.holdLabel.setText(str(elapsed_min))

        self.holdLabel.setText(f'{self.total_elapsed_time:.2f} min')
        self.hold_start = None
        self.calc_Total_time()


    def FinishDate(self):
        finishdate = QDateTime.currentDateTime().toString('dd/MM/yyyy HH:mm:ss')
        self.finishLabel.setText(finishdate)
        self.calc_Total_time()


    def createNewCase(self):
        # Fix bug of adding new case and it is blank and not editable
        name="Jakub Wloka"

        '''Start Date', 'Finish Date', 'Event Date', 'Handover Type', 'Requestor CA/QA',
                     'Report name', 'Case Entered', 'Requestor GPN', 'Rank', 'No of Invoices', 'Number of credit card line items', 'No of cash expenses',
                      'Number of Iterations', 'Total entered line items', 'Case Closed', 'Case Pending',
                      'Time spent total in min', 'Receipt email', 'Confirmation email', 'Reason if pending',
                      'Date of Action', 'Total Hold Time', 'Hold Continue', 'Hold Time', 'Action', 'Current Status',
                      'Received Date', 'Total Hold Time', 'Hold Continue'''

        sql = ('INSERT INTO [Expense_Management] ([Start Date], [Finish Date], [Event Date], [Handover Type], '
               '[Requestor CA/ DH], [Report name], [Case entered by], [Requester GPN], '
               '[Number of credit card line items], [No of cash expenses], [Number of Iterations], '
               '[Case Closed], [Case Pending], [Time spent total in min], [Receipt email], [Confirmation email], '
               '[Date of Action], [Action], [Current Status], [Received Date], [Total Hold Time], [Hold Continue], '
               'VALUES (NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '
               'NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL)')
        try:

            self.cursor.execute(sql)
            self.conn.commit()

            # Obtain new ID
            self.cursor.execute('SELECT @@IDENTITY')
            new_id = self.cursor.fetchone()[0]

        except Exception as e:
            print(f"Error entering New Case: {e}")
            QMessageBox.critical(self, 'Error', 'Error entering New Case')

            # manual fill the TableWidget
            row_count = self.tableWidget.rowCount()
            self.tableWidget.insertRow(row_count)
            #self.tableWidget.setItem(row_count, 0, QTableWidgetItem(user_name))
            self.tableWidget.setItem(row_count, 1, QTableWidgetItem(str(new_id)))
            self.loadtableData()
            self.SelectNewCase()
            print("Record added successfully")
            QMessageBox.information(self, 'New Case', 'New Case added')


    def save(self):
        if self.currentrow is not None:
            datetext = self.receive_date.text()
            starttext = self.startLabel.text()
            finishtext = self.finishLabel.text()
            holdtext = self.holdLabel.text()
            case = self.entered_by_txt.currentText()
            invoices = self.invoice_txt.text()
            requestor = self.requestor_txt.text()
            requestorG = self.requestorG_txt.text()
            reason = self.reason_txt.text()
            status = self.status_txt.currentText()
            handover = self.handover_txt.currentText()
            report = self.report_txt.text()
            trip = self.trip_txt.text()
            line = self.line_txt.text()
            action = self.action_txt.text()
            nr_iter = self.nr_of_iter_txt.currentText()
            time_spend = self.total_time_txt.text()

            item = self.tableWidget.item(self.currentrow, 0)
            value = item.text()

            self.tableWidget.setItem(self.currentrow, 1, QTableWidgetItem(starttext))
            self.tableWidget.setItem(self.currentrow, 2, QTableWidgetItem(finishtext))
            self.tableWidget.setItem(self.currentrow, 4, QTableWidgetItem(handover))
            self.tableWidget.setItem(self.currentrow, 5, QTableWidgetItem(requestor))


            self.tableWidget.setItem(self.currentrow, 6, QTableWidgetItem(trip))
            self.tableWidget.setItem(self.currentrow, 7, QTableWidgetItem(report))
            self.tableWidget.setItem(self.currentrow, 8, QTableWidgetItem(case))
            self.tableWidget.setItem(self.currentrow, 9, QTableWidgetItem(requestorG))
            self.tableWidget.setItem(self.currentrow, 10, QTableWidgetItem(invoices))
            self.tableWidget.setItem(self.currentrow, 11, QTableWidgetItem(nr_iter))
            self.tableWidget.setItem(self.currentrow, 12, QTableWidgetItem(line))
            self.tableWidget.setItem(self.currentrow, 13, QTableWidgetItem(time_spend))
            self.tableWidget.setItem(self.currentrow, 14, QTableWidgetItem(reason))
            self.tableWidget.setItem(self.currentrow, 15, QTableWidgetItem(action))
            self.tableWidget.setItem(self.currentrow, 16, QTableWidgetItem(status))
            self.tableWidget.setItem(self.currentrow, 17, QTableWidgetItem(datetext))
            self.tableWidget.setItem(self.currentrow, 18, QTableWidgetItem(holdtext))

            sql = "UPDATE [Expense_Management] SET [Case entered by]=?, [Start Date]=?, [Finish Date]=?, [Hold]=?, [Date]=?, [Handover]=?, [Requestor]=?, [Iterations]=?, [Invoices]=?, [Line Items]=?, [Time Spent]=?, [Reason]=?, [Action]=?, [Status]=?, [Date]=?, [Hold Time]=? WHERE [ID]=?"
            self.cursor.execute(sql, (
            case, starttext, finishtext, holdtext, datetext, handover, requestor, nr_iter, invoices, line, time_spend, reason,
            action, status, datetext, holdtext, value))
            self.conn.commit()
            print("Data saved successfully")
            QMessageBox.information(self, 'Info', 'Data saved successfully')

        else:
            print("Record empty")
            QMessageBox.critical(self, 'Error', 'Record empty')


    def delete(self):
        reply = QMessageBox.question(self, "Delete Confirmation", "Do you really want to delete selected record ?",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            if self.currentrow is not None:
                item = self.tableWidget.item(self.currentrow, 0)
                value = item.text()
                try:
                    sql = "DELETE FROM [Expense_Management] WHERE ID = ?"
                    self.cursor.execute(sql, (value,))
                    self.conn.commit()
                    self.tableWidget.removeRow(self.currentrow)
                    print("Record removed successfully ")
                except Exception as e:
                    print(f"Error: {e}")
        else:
            print("Deletion cancelled")


    def doubleClick(self, _item):
        self.currentrow = None
        self.currentrow = self.tableWidget.currentRow()


        clickedID = self.tableWidget.item(self.currentrow, 0).text()  # ID
        self.id_label.setText(f"Selected ID: {clickedID}")

        startdate = self.tableWidget.item(self.currentrow, 1).text()  # Start Date
        finishdate = self.tableWidget.item(self.currentrow, 2).text()  # Finish Date
        holdtime = self.tableWidget.item(self.currentrow, 28).text()  # HoldTime
        receivedate = self.tableWidget.item(self.currentrow, 25).text()  # Receive Date
        case = self.tableWidget.item(self.currentrow, 8).text()  # Case entered by
        invoices = self.tableWidget.item(self.currentrow, 10).text()  # no of inv
        requestor = self.tableWidget.item(self.currentrow, 5).text()  # requestor
        status = self.tableWidget.item(self.currentrow, 24).text()  # status
        trip = self.tableWidget.item(self.currentrow, 6).text()  # trip ID
        report = self.tableWidget.item(self.currentrow, 7).text()  # report name
        iter = self.tableWidget.item(self.currentrow, 14).text()  # nr of iterations
        requestorG = self.tableWidget.item(self.currentrow, 9).text()  # requestor GPN
        reason = self.tableWidget.item(self.currentrow, 21).text()  # reason of pending
        action = self.tableWidget.item(self.currentrow, 23).text()  # action
        line = self.tableWidget.item(self.currentrow, 15).text()  # Nr of line items
        time_spend = self.tableWidget.item(self.currentrow, 18).text()  # Total time spend
        handover = self.tableWidget.item(self.currentrow, 4).text()  # Handover type

        date = QDate.fromString(receivedate, 'dd/MM/yyyy')  # Start Date
        self.startLabel.setText(startdate)
        self.finishLabel.setText(finishdate)
        self.holdLabel.setText(holdtime)
        self.entered_by_txt.setCurrentText(case)
        self.entered_by_txt.setCurrentText(case)
        self.invoice_txt.setText(invoices)
        self.receive_date.setDate(date)
        self.requestor_txt.setText(requestor)
        self.status_txt.setCurrentText(status)
        self.report_txt.setText(report)
        self.trip_txt.setText(trip)
        self.nr_of_iter_txt.setCurrentText(iter)
        self.requestorG_txt.setText(requestorG)
        self.reason_txt.setText(reason)
        self.action_txt.setText(action)
        self.line_txt.setText(line)
        self.total_time_txt.setText(time_spend)
        self.handover_txt.setCurrentText(handover)


    def SelectNewCase(self):
        if self.tableWidget.rowCount() > 0:
            self.tableWidget.selectRow(0)
            self.tableWidget.scrollToItem(self.tableWidget.item(0, 0), QAbstractItemView.PositionAtTop)

            clickedID = self.tableWidget.item(0, 0).text()  # ID
            self.idLabel.setText(f"Selected ID: {clickedID}")


            startdate = self.tableWidget.item(0, 1).text()  # Start Date
            finishdate = self.tableWidget.item(0, 2).text()  # Finish Date
            holdtime = self.tableWidget.item(0, 28).text()  # HoldTime
            receivedate = self.tableWidget.item(0, 25).text()  # Receive Date
            case = self.tableWidget.item(0, 8).text()  # Case entered by
            invoices = self.tableWidget.item(0, 10).text()  # no of inv
            requestor = self.tableWidget.item(0, 5).text()  # requestor
            status = self.tableWidget.item(0, 24).text()  # status
            trip = self.tableWidget.item(0, 6).text()  # trip ID
            report = self.tableWidget.item(0, 7).text()  # report name
            iter = self.tableWidget.item(0, 14).text()  # nr of iterations
            requestorG = self.tableWidget.item(0, 9).text()  # requestor GPN
            reason = self.tableWidget.item(0, 21).text()  # reason of pending
            action = self.tableWidget.item(0, 23).text()  # action
            line = self.tableWidget.item(0, 15).text()  # Nr of line items
            time_spend = self.tableWidget.item(0, 18).text()  # Total time spend
            handover = self.tableWidget.item(0, 4).text()  # Handover type

            self.startLabel.setText(startdate)  # Start Date
            self.finishLabel.setText(finishdate)  # Finish Date
            self.holdLabel.setText(holdtime)  # HoldTime
            self.entered_by_txt.setCurrentText(case)  # Case entered by
            self.invoice_txt.setText(invoices)  # no of inv
            self.receive_date.setDate(QDate.currentDate())  # Receive Date
            self.requestor_txt.setText(requestor)  # requestor
            self.status_txt.setCurrentText(status)  # status
            self.report_txt.setText(report)  # report name
            self.trip_txt.setText(trip)  # trip ID
            self.nr_of_iter_txt.setCurrentText(iter)  # nr of iterations
            self.requestorG_txt.setText(requestorG)  # requestor GPN
            self.reason_txt.setText(reason)  # reason of pending
            self.action_txt.setText(action)  # action
            self.line_txt.setText(line)  # Nr of line items
            self.total_time_txt.setText(time_spend)  # Total time spend
            self.handover_txt.setCurrentText(handover)  # Handover type


    def filter(self):
        try:
            sqlQuery = "SELECT * FROM Expense_Management WHERE [Current Status]=? AND [Case entered by]=? ORDER BY ID DESC"
            self.cursor.execute(sqlQuery, ('Open', self.user_name,))
            rows = self.cursor.fetchall()

            self.tableWidget.clearContents()
            self.tableWidget.setRowCount(len(rows))

            for row_index, row_data in enumerate(rows):
                for column_index, col_data in enumerate(row_data):
                    cell = QTableWidgetItem(str(col_data))
                    cell.setFlags(cell.flags() ^ Qt.ItemIsEditable)
                    self.tableWidget.setItem(row_index, column_index, cell)

        except Exception as e:
            print(f"Error: {e}")


    def filter_off(self):
        try:
            sqlQuery = "SELECT * FROM Expense_Management ORDER BY ID DESC"
            self.cursor.execute(sqlQuery)
            rows = self.cursor.fetchall()

            self.tableWidget.clearContents()
            self.tableWidget.setRowCount(len(rows))

        except Exception as e:
            print(f"Error: {e}")


    def calc_Total_time(self):
        finish_date_str = self.finishLabel.text()
        start_date_str = self.startLabel.text()
        hold_time_str = self.holdLabel.text()

        # Uncomment the following line if you want to return from the function when start or finish date is not set
        # if not self.finishLabel.text() or not self.startLabel.text():
        #     return

        try:
            hold_time_min = float(hold_time_str.split()[0])
        except:
            hold_time_min = 0


        try:
            hold_time_min = float(hold_time_str.split()[0])
        except:
            hold_time_min = 0

        try:
            finish_date = datetime.strptime(finish_date_str, "%d/%m/%Y %H:%M:%S")
        except:
            finish_date = 0

        if start_date_str == 'None':
            return
        else:
            try:
                start_date = datetime.strptime(start_date_str, "%d/%m/%Y %H:%M:%S")
            except:
                start_date = datetime.strptime(start_date_str, "%Y-%m-%d %H:%M:%S")

        try:
            total_time = (finish_date - start_date).total_seconds() / 60 - hold_time_min
        except:
            return

        minutes = int(total_time)
        seconds = int((total_time - minutes) * 60)

        total_time_str = f"{minutes:02d}:{seconds:02d} min"
        try:
            total_time = (finish_date - start_date).total_seconds() / 60 - hold_time_min
        except:
            return

        minutes = int(total_time)
        seconds = int((total_time - minutes) * 60)

        total_time_str = f"{minutes:02d}:{seconds:02d} min"
        self.total_time_txt.setText(total_time_str)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    sys.exit(app.exec_())
