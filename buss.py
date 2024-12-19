from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QTableWidget, QTableWidgetItem,
    QLineEdit, QLabel, QPushButton, QComboBox, QMessageBox, QHBoxLayout
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
import pandas as pd
from datetime import datetime

# File path to store bus data
FILE_PATH = "buses.xlsx"


def load_data():
    try:
        return pd.read_excel(FILE_PATH)
    except FileNotFoundError:
        df = pd.DataFrame(columns=["Bus ID", "Route", "Driver", "Insurance Expiry", "Maintenance Due", "Status"])
        df.to_excel(FILE_PATH, index=False)
        return df


def save_data(df):
    try:
        df.to_excel(FILE_PATH, index=False)
    except Exception as e:
        QMessageBox.critical(None, "Error", f"An error occurred while saving data: {e}")


class BusManagementSystem(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Bus Management System")
        self.setGeometry(200, 200, 1000, 600)

        # Initialize DataFrame
        self.df = load_data()

        # Main widget and layout
        self.main_widget = QWidget()
        self.main_layout = QVBoxLayout(self.main_widget)

        # Title
        title = QLabel("Bus Management System")
        title.setFont(QFont("Arial", 24))
        title.setAlignment(Qt.AlignCenter)
        self.main_layout.addWidget(title)

        # Input Fields
        self.create_input_fields()

        # Table View
        self.create_table_view()

        # Action Buttons
        self.create_action_buttons()

        self.setCentralWidget(self.main_widget)

    def create_input_fields(self):
        input_layout = QVBoxLayout()

        # Bus ID
        self.bus_id_input = self.create_input_row("Bus ID", input_layout)
        # Route
        self.route_input = self.create_input_row("Route", input_layout)
        # Driver
        self.driver_input = self.create_input_row("Driver", input_layout)
        # Insurance Expiry
        self.insurance_input = self.create_input_row("Insurance Expiry (YYYY-MM-DD)", input_layout)
        # Maintenance Due
        self.maintenance_input = self.create_input_row("Maintenance Due (YYYY-MM-DD)", input_layout)
        # Status
        status_label = QLabel("Status")
        self.status_combo = QComboBox()
        self.status_combo.addItems(["Active", "Under Maintenance"])
        input_layout.addWidget(status_label)
        input_layout.addWidget(self.status_combo)

        self.main_layout.addLayout(input_layout)

    def create_input_row(self, label_text, layout):
        label = QLabel(label_text)
        input_field = QLineEdit()
        layout.addWidget(label)
        layout.addWidget(input_field)
        return input_field

    def create_table_view(self):
        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(["Bus ID", "Route", "Driver", "Insurance Expiry", "Maintenance Due", "Status"])
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.main_layout.addWidget(self.table)
        self.load_table_data()

    def create_action_buttons(self):
        button_layout = QHBoxLayout()

        # Buttons
        add_update_button = QPushButton("Add/Update Bus")
        add_update_button.clicked.connect(self.add_or_update_bus)
        button_layout.addWidget(add_update_button)

        view_all_button = QPushButton("View All Buses")
        view_all_button.clicked.connect(self.load_table_data)
        button_layout.addWidget(view_all_button)

        check_insurance_button = QPushButton("Check Expired Insurance")
        check_insurance_button.clicked.connect(self.check_expired_insurance)
        button_layout.addWidget(check_insurance_button)

        check_maintenance_button = QPushButton("Check Maintenance Due")
        check_maintenance_button.clicked.connect(self.check_maintenance_due)
        button_layout.addWidget(check_maintenance_button)

        exit_button = QPushButton("Exit")
        exit_button.clicked.connect(self.close)
        button_layout.addWidget(exit_button)

        self.main_layout.addLayout(button_layout)

    def load_table_data(self):
        self.table.setRowCount(0)
        for _, row in self.df.iterrows():
            row_position = self.table.rowCount()
            self.table.insertRow(row_position)
            for col_index, value in enumerate(row):
                self.table.setItem(row_position, col_index, QTableWidgetItem(str(value)))

    def add_or_update_bus(self):
        bus_id = self.bus_id_input.text()
        route = self.route_input.text()
        driver = self.driver_input.text()
        insurance_expiry = self.insurance_input.text()
        maintenance_due = self.maintenance_input.text()
        status = self.status_combo.currentText()

        if not all([bus_id, route, driver, insurance_expiry, maintenance_due, status]):
            QMessageBox.warning(self, "Input Error", "All fields must be filled!")
            return

        if bus_id in self.df["Bus ID"].values:
            self.df.loc[self.df["Bus ID"] == bus_id, ["Route", "Driver", "Insurance Expiry", "Maintenance Due", "Status"]] = \
                [route, driver, insurance_expiry, maintenance_due, status]
        else:
            new_row = {"Bus ID": bus_id, "Route": route, "Driver": driver,
                       "Insurance Expiry": insurance_expiry, "Maintenance Due": maintenance_due, "Status": status}
            self.df = pd.concat([self.df, pd.DataFrame([new_row])], ignore_index=True)

        save_data(self.df)
        self.load_table_data()
        QMessageBox.information(self, "Success", "Bus details have been updated!")

    def check_expired_insurance(self):
        expired = self.df[pd.to_datetime(self.df["Insurance Expiry"], errors="coerce") < datetime.now()]
        if expired.empty:
            QMessageBox.information(self, "Insurance Check", "No buses with expired insurance.")
        else:
            self.show_popup("Buses with Expired Insurance", expired)

    def check_maintenance_due(self):
        due = self.df[pd.to_datetime(self.df["Maintenance Due"], errors="coerce") <= datetime.now()]
        if due.empty:
            QMessageBox.information(self, "Maintenance Check", "No buses needing maintenance.")
        else:
            self.show_popup("Buses Needing Maintenance", due)

    def show_popup(self, title, data):
        popup = QMainWindow(self)
        popup.setWindowTitle(title)
        popup.setGeometry(300, 300, 600, 400)

        table = QTableWidget()
        table.setColumnCount(len(data.columns))
        table.setHorizontalHeaderLabels(data.columns)
        table.setRowCount(0)

        for _, row in data.iterrows():
            row_position = table.rowCount()
            table.insertRow(row_position)
            for col_index, value in enumerate(row):
                table.setItem(row_position, col_index, QTableWidgetItem(str(value)))

        popup.setCentralWidget(table)
        popup.show()


if __name__ == "__main__":
    app = QApplication([])
    window = BusManagementSystem()
    window.show()
    app.exec()
