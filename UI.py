import sys
from pathlib import Path
import pyodbc
import pandas as pd
from datetime import datetime
from PySide2.QtWidgets import QApplication, QVBoxLayout, QPushButton, QWidget, QLabel, QTableWidget, \
    QAbstractItemView
from PySide2.QtWidgets import QMessageBox, QTableWidgetItem, QGridLayout, QLineEdit, QComboBox, QDateEdit
from PySide2.QtCore import Qt, QDateTime, QDate
from PySide2.QtGui import QFont



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

    def get_user_name(self):

        excel_path = Path(__file__).parent / "Book.xlsx"
        df = pd.read_excel(excel_path)
        names = df['Name'].dropna().unique().tolist()
        self.entered_by_txt.addItems(names)

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

        receive_label = QLabel('Receive Date')
        self.receive_date = QDateEdit(self)
        self.receive_date.setCalendarPopup(True)
        self.receive_date.setDate(QDate.currentDate())
        gridLayout.addWidget(receive_label, 1, 0)
        gridLayout.addWidget(self.receive_date, 1, 1)

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
        self.finishButton.clicked.connect(self.FinishDate)
        gridLayout.addWidget(self.finishButton, 4, 0)
        gridLayout.addWidget(self.finishLabel, 4, 1)

        # Total Time
        total_time_label = QLabel('Total Time:')
        self.total_time_txt = QLineEdit()
        gridLayout.addWidget(total_time_label, 5, 0)
        gridLayout.addWidget(self.total_time_txt, 5, 1)

        # Entered By
        self.entered_by_label = QPushButton('Case entered by:')
        self.entered_by_txt = QComboBox()
        self.get_user_name()
        gridLayout.addWidget(self.entered_by_label, 6, 0)
        gridLayout.addWidget(self.entered_by_txt, 6, 1)
        self.entered_by_label.clicked.connect(self.your_data)

        # Status
        status_label = QLabel('Status:')
        self.status_txt = QComboBox()
        self.status_txt.addItems(['Open', 'Closed', 'Advised', 'None'])
        gridLayout.addWidget(status_label, 7, 0)
        gridLayout.addWidget(self.status_txt, 7, 1)


        # SECOND COLUMN

        # Requestor
        requestor_label = QLabel('Requestor CA/QA:')
        self.requestor_txt = QLineEdit()
        gridLayout.addWidget(requestor_label,1, 2)
        gridLayout.addWidget(self.requestor_txt, 1, 3)

        # Requestor GPN
        requestorG_label = QLabel('Requestor GPN:')
        self.requestorG_txt = QLineEdit()
        gridLayout.addWidget(requestorG_label,2, 2)
        gridLayout.addWidget(self.requestorG_txt, 2, 3)

        # No. of Invoices:
        invoice_label = QLabel('No of Invoices:')
        self.invoice_txt = QLineEdit()
        gridLayout.addWidget(invoice_label, 3, 2)
        gridLayout.addWidget(self.invoice_txt, 3, 3)

        # Handover
        self.handover_label = QLabel('Handover Type:')
        self.handover_txt = QComboBox()
        self.handover_txt.addItems(['None', 'Genesys', 'Other channel'])
        gridLayout.addWidget(self.handover_label, 4, 2)
        gridLayout.addWidget(self.handover_txt, 4, 3)

        # Trip ID
        trip_label = QLabel('Case ID:')
        self.trip_txt = QLineEdit()
        gridLayout.addWidget(trip_label, 5, 2)
        gridLayout.addWidget(self.trip_txt, 5, 3)
        mainLayout.addLayout(gridLayout)

        # Report Name
        report_label = QLabel('Report Name:')
        self.report_txt = QLineEdit()
        gridLayout.addWidget(report_label, 6, 2)
        gridLayout.addWidget(self.report_txt, 6, 3)

        # Nr. of iteration
        line_label = QLabel('Total entered line items:')
        self.line_txt = QLineEdit()
        gridLayout.addWidget(line_label, 7, 2)
        gridLayout.addWidget(self.line_txt, 7, 3)

        # Third COLUMN

        # Nr. of iter
        nr_of_iter_label = QLabel('Number of Iterations:')
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
        gridLayout.addWidget(nr_of_iter_label, 1, 4)
        gridLayout.addWidget(self.nr_of_iter_txt, 1, 5)

        # Reason of pending
        reason_label = QLabel('Reason of pending:')
        self.reason_txt = QLineEdit()
        gridLayout.addWidget(reason_label, 2, 4)
        gridLayout.addWidget(self.reason_txt, 2, 5)

        # Action
        action_label = QLabel('Action:')
        self.action_txt = QLineEdit()
        gridLayout.addWidget(action_label, 3, 4)
        gridLayout.addWidget(self.action_txt, 3, 5)

        # Filter cases by open
        filteroff_label = QLabel('Filter cases by open:')
        self.filteroff_button = QPushButton('Filter Open')
        gridLayout.addWidget(filteroff_label, 4, 4)
        gridLayout.addWidget(self.filteroff_button, 4, 5)
        self.filteroff_button.clicked.connect(self.filter_off)

        # Filter cases off
        filter_label = QLabel('Filter cases by close:')
        self.filter_button = QPushButton('Filter Close')
        gridLayout.addWidget(filter_label, 5, 4)
        gridLayout.addWidget(self.filter_button, 5, 5)
        self.filter_button.clicked.connect(self.filter)

        mainLayout.addLayout(gridLayout)
        mainLayout.addWidget(self.id_label)

        # Table for DB
        self.tableWidget = QTableWidget()
        self.tableWidget.setColumnCount(24)
        col_names = ['ID', 'Start Date', 'Finish Date', 'Event Date', 'Handover Type', 'Requestor CA/ DH', 'Trip ID' ,
                     'Report name', 'Case Entered', 'Requestor GPN', 'Number of credit card line items', 'No of cash expenses',
                      'Number of Iterations', 'Case Closed', 'Case Pending', 'Time spent total in min',
                      'Receipt email', 'Confirmation email', 'Date of Action', 'Action', 'Current Status',
                     'Receive Date' ,'Total Hold Time', 'Hold Continue' ]
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

        accdb_path = Path(__file__).parent / "KubaDatabase3.accdb"
        self.conn = pyodbc.connect(
            r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};" + f"DBQ={accdb_path.absolute()}"
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
                    if col_data is None or col_data == "" or col_data == " ":
                        cell = QTableWidgetItem(" ")
                    else:
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
            print('Case 1')
            self.hold_start = datetime.now()
            self.holdLabel.setText("Timing...")
        else:
            hold_end = datetime.now()
            print('Case 2')
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

        name="Jakub Wloka"

        '''sql = ('INSERT INTO [Expense_Management] ([Start Date], [Finish Date], [Event Date], [Handover Type], '
               '[Requestor CA/ DH], [TripID], [Report name], [Case entered by], [Requester GPN], '
               '[Number of credit card line items], [No of cash expenses], [Number of Iterations], '
               '[Case Closed], [Case Pending], [Time spent total in min], [Receipt email], [Confirmation email], '
               '[Date of Action], [Action], [Current Status], [Received Date], [Total Hold Time], [Hold Continue], '
               'VALUES (NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '
               'NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL)')'''

        sql = ('INSERT INTO [Expense_Management] ([Start Date])'
               'VALUES (NULL)')
        try:

            self.cursor.execute(sql)
            self.conn.commit()
            print('records added')

            # Obtain new ID
            self.cursor.execute('SELECT @@IDENTITY')
            new_id = self.cursor.fetchone()[0]

        except Exception as e:
            print(f"Error entering New Case: {e}")
            QMessageBox.critical(self, 'Error', 'Error entering New Case')

        # manual fill the TableWidget
        row_count = self.tableWidget.rowCount()
        self.tableWidget.insertRow(row_count)
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
            self.tableWidget.setItem(self.currentrow, 12, QTableWidgetItem(nr_iter))
            self.tableWidget.setItem(self.currentrow, 14, QTableWidgetItem(time_spend))
            self.tableWidget.setItem(self.currentrow, 14, QTableWidgetItem(reason))
            self.tableWidget.setItem(self.currentrow, 15, QTableWidgetItem(action))
            self.tableWidget.setItem(self.currentrow, 16, QTableWidgetItem(status))
            self.tableWidget.setItem(self.currentrow, 17, QTableWidgetItem(datetext))
            self.tableWidget.setItem(self.currentrow, 18, QTableWidgetItem(holdtext))


            print("Parameters:", case,",", starttext,",", finishtext,",", datetext,",", handover,",", nr_iter,",", invoices,",", line,",", value)


            #sql = "UPDATE [Expense_Management] SET [Case entered by]=?, [Start Date]=?, [Finish Date]=?, [Received Date]=?, [Handover Type]=?, [Number of Iterations]=?, [No of Invoices]=?, [Total entered line items]=? WHERE [ID]=?"

            sql = """
                            UPDATE [Expense_Management] 
                            SET 
                                [Case entered by]=?, 
                                [Start Date]=?, 
                                [Finish Date]=?
                            WHERE [ID]=?
                        """

            self.cursor.execute(sql, (
                case,
                starttext,
                finishtext,
                value  # Assuming 'value' is the ID for the WHERE clause
            ))
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
        self.currentrow = self.tableWidget.currentRow()

        clickedID = self.tableWidget.item(self.currentrow, 0).text() if self.tableWidget.item(self.currentrow,
                                                                                              0) else "No ID"
        self.id_label.setText(f"Selected ID: {clickedID}")

        def get_text(row, col):
            item = self.tableWidget.item(row, col)
            return item.text() if item else ""

        # Use the helper function to get the text for each field
        startdate = get_text(self.currentrow, 1)  # Start Date
        finishdate = get_text(self.currentrow, 2)  # Finish Date
        holdtime = get_text(self.currentrow, 28)  # Hold Time
        receivedate = get_text(self.currentrow, 25)  # Receive Date
        case = get_text(self.currentrow, 8)  # Case entered by
        invoices = get_text(self.currentrow, 10)  # no of inv
        requestor = get_text(self.currentrow, 5)  # requestor
        status = get_text(self.currentrow, 24)  # status
        trip = get_text(self.currentrow, 6)  # trip ID
        report = get_text(self.currentrow, 7)  # report name
        iter = get_text(self.currentrow, 14)  # nr of iterations
        requestorG = get_text(self.currentrow, 9)  # requestor GPN
        reason = get_text(self.currentrow, 21)  # reason of pending
        action = get_text(self.currentrow, 23)  # action
        line = get_text(self.currentrow, 15)  # Nr of line items
        time_spend = get_text(self.currentrow, 18)  # Total time spend
        handover = get_text(self.currentrow, 4)  # Handover type

        date = QDate.fromString(receivedate, 'dd/MM/yyyy') if receivedate else QDate.currentDate()

        # Set the values in the corresponding fields
        self.startLabel.setText(startdate)
        self.finishLabel.setText(finishdate)
        self.holdLabel.setText(holdtime)
        self.receive_date.setDate(date)
        self.entered_by_txt.setCurrentText(case)
        self.invoice_txt.setText(invoices)
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


            def get_text(row, col):
                item = self.tableWidget.item(row, col)
                return item.text() if item else ""


            clickedID = get_text(0, 0)  # ID
            startdate = get_text(0, 1)  # Start Date
            finishdate = get_text(0, 2)  # Finish Date
            holdtime = get_text(0, 28)  # HoldTime
            receivedate = get_text(0, 25)  # Receive Date
            case = get_text(0, 8)  # Case entered by
            invoices = get_text(0, 10)  # no of inv
            requestor = get_text(0, 5)  # requestor
            status = get_text(0, 24)  # status
            trip = get_text(0, 6)  # trip ID
            report = get_text(0, 7)  # report name
            iter = get_text(0, 14)  # nr of iterations
            requestorG = get_text(0, 9)  # requestor GPN
            reason = get_text(0, 21)  # reason of pending
            action = get_text(0, 23)  # action
            line = get_text(0, 15)  # Nr of line items
            time_spend = get_text(0, 18)  # Total time spend
            handover = get_text(0, 4)  # Handover type

            # Set the values in the corresponding fields
            self.startLabel.setText(startdate)
            self.finishLabel.setText(finishdate)
            self.holdLabel.setText(holdtime)
            self.entered_by_txt.setCurrentText(case)
            self.invoice_txt.setText(invoices)
            self.receive_date.setDate(
                QDate.fromString(receivedate, 'dd/MM/yyyy') if receivedate else QDate.currentDate())
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
            self.id_label.setText(f"Selected ID: {clickedID}")


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

            self.loadtableData()

        except Exception as e:
            print(f"Error: {e}")


    def calc_Total_time(self):
        finish_date_str = self.finishLabel.text()
        start_date_str = self.startLabel.text()
        hold_time_str = self.holdLabel.text()

        try:
            hold_time_min = float(hold_time_str.split()[0])
            #print('case1')
        except:
            hold_time_min = 0

        try:
            finish_date = datetime.strptime(finish_date_str, "%d/%m/%Y %H:%M:%S")
        except:
            finish_date = 0

        if start_date_str is None or start_date_str.strip() == '':
            return
        else:
            try:
                start_date = datetime.strptime(start_date_str, "%d/%m/%Y %H:%M:%S")
            except:
                start_date = datetime.strptime(start_date_str, "%Y-%m-%d %H:%M:%S")

        try:
            total_time = (finish_date - start_date).total_seconds() / 60 - hold_time_min
            #print('case2')
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
