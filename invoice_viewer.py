import sys
from PySide2.QtCore import Slot
from PySide2.QtWidgets import (QAction, QApplication, QHeaderView, QHBoxLayout,
                               QMainWindow, QTableWidget, QTableWidgetItem, QWidget)
import Database_management as dm


class Widget(QWidget):
    def __init__(self):
        QWidget.__init__(self)
        self.items = 0

        # Example Data
        self._data = dm.read_record()

        # Left
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["Invoice Number", "Customer Name", "Amount", "Date"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        # Right
        """self.search = QLineEdit()
        self.add = QPushButton("Add")
        self.clear = QPushButton("Clear")
        self.quit = QPushButton("Quit")
        self.plot = QPushButton("Plot")

        self.right = QVBoxLayout()
        self.right.setMargin(10)
        self.right.addWidget(QLabel("Search"))
        self.right.addWidget(self.search)
        self.right.addWidget(QLabel("Description"))
        self.right.addWidget(self.add)
        self.right.addWidget(self.plot)
        self.right.addWidget(self.clear)
        self.right.addWidget(self.quit)"""

        # QWidget Layout
        self.layout = QHBoxLayout()

        # self.table_view.setSizePolicy(size)
        self.layout.addWidget(self.table)

        # set the layout to QWidget
        self.setLayout(self.layout)

        #Fill example data
        self.fill_table()

    def fill_table(self, data=None):
        data = self._data if not data else data
        for invoiceDetail in data:
            invoice_number = invoiceDetail["Invoice_number"]
            customer_name = invoiceDetail["Customer_details"]["Name"]
            amount = invoiceDetail["Amount"]
            date = invoiceDetail["Date"]
            self.table.insertRow(self.items)
            self.table.setItem(self.items, 0, QTableWidgetItem(str(invoice_number)))
            self.table.setItem(self.items, 1, QTableWidgetItem(customer_name))
            self.table.setItem(self.items, 2, QTableWidgetItem(str(amount)))
            self.table.setItem(self.items, 3, QTableWidgetItem(date))
            self.items += 1


class MainWindow(QMainWindow):
    def __init__(self, widget):
        QMainWindow.__init__(self)
        self.setWindowTitle("Invoice Manager")

        # Menu
        self.menu = self.menuBar()
        self.file_menu = self.menu.addMenu("File")

        # Exit Action
        exit_action = QAction("Exit", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.triggered.connect(self.exit_app)

        self.file_menu.addAction(exit_action)
        self.setCentralWidget(widget)

    @Slot()
    def exit_app(self, checked):
        QApplication.quit()


if __name__ == "__main__":
    # Qt Application
    app = QApplication(sys.argv)

    # QWidget
    widget = Widget()

    # QMainWindow using QWidget as central widget
    window = MainWindow(widget)
    window.resize(800, 600)
    window.show()

    # Execute Application
    sys.exit(app.exec_())
