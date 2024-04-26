import sys
import sqlite3
from PyQt5.QtWidgets import QInputDialog, QListWidgetItem, QFileDialog, QFormLayout, QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QDialog, QLineEdit, QLabel, QHBoxLayout, QComboBox, QMessageBox, QListWidget
from PyQt5.QtCore import Qt, pyqtSignal
from docx import Document
import pandas as pd
from PyQt5 import QtCore

def init_db():
    connection = sqlite3.connect('AbelyashevTest.db')
    cursor = connection.cursor()

    cursor.execute('''
    CREATE TABLE IF NOT EXISTS Branches(
        BranchID INTEGER PRIMARY KEY AUTOINCREMENT,
        BranchesName TEXT NOT NULL,
        CEO_ID INTEGER,
        FOREIGN KEY (CEO_ID) REFERENCES Employees(EmployeeID)
    )
    ''')

    cursor.execute('''
    CREATE TABLE IF NOT EXISTS Employees (
        EmployeeID INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        surname TEXT NOT NULL,
        branch_id INTEGER NOT NULL,
        FOREIGN KEY (branch_id) REFERENCES Branches(BranchID)
    )
    ''')

    connection.commit()
    connection.close()



class EmployeeForm(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Сотрудник')

        self.name_edit = QLineEdit(self)
        self.surname_edit = QLineEdit(self)
        self.branch_combo = QComboBox(self)
        self.save_button = QPushButton('Сохранить', self)
        self.cancel_button = QPushButton('Отмена', self)

        form_layout = QVBoxLayout()
        form_layout.addWidget(QLabel('Имя:'))
        form_layout.addWidget(self.name_edit)
        form_layout.addWidget(QLabel('Фамилия:'))
        form_layout.addWidget(self.surname_edit)
        form_layout.addWidget(QLabel('Филиал:'))
        form_layout.addWidget(self.branch_combo)

        button_layout = QHBoxLayout()
        button_layout.addWidget(self.save_button)
        button_layout.addWidget(self.cancel_button)
        form_layout.addLayout(button_layout)

        self.setLayout(form_layout)

        self.save_button.clicked.connect(self.save_employee)
        self.cancel_button.clicked.connect(self.reject)

    def save_employee(self):
        name = self.name_edit.text()
        surname = self.surname_edit.text()
        branch_id = self.branch_combo.currentData()
        if name and surname and branch_id:
            connection = sqlite3.connect('AbelyashevTest.db')
            cursor = connection.cursor()
            cursor.execute('''INSERT INTO Employees (name, surname, branch_id) VALUES (?, ?, ?)''', (name, surname, branch_id))
            connection.commit()
            connection.close()
            QMessageBox.information(self, 'Успех', 'Сотрудник сохранён!')
            self.accept()
            self.parent().load_employees()
        else:
            QMessageBox.warning(self, 'Ошибка', 'Все поля должны быть заполнены!')

    def load_branches(self):
        self.branch_combo.clear()
        connection = sqlite3.connect('AbelyashevTest.db')
        cursor = connection.cursor()
        cursor.execute('''SELECT BranchID, BranchesName FROM Branches''')
        branches = cursor.fetchall()
        for branch_id, branch_name in branches:
            self.branch_combo.addItem(branch_name, branch_id)
        connection.close()






class BranchListDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Список филиалов')

        self.branch_list_widget = QListWidget()
        self.branch_list_widget.itemSelectionChanged.connect(self.enable_edit_button)

        self.add_branch_button = QPushButton('Добавить филиал')
        self.add_branch_button.clicked.connect(self.add_branch)

        self.edit_branch_button = QPushButton('Редактировать филиал')
        self.edit_branch_button.setEnabled(False)
        self.edit_branch_button.clicked.connect(self.edit_branch)

        self.delete_branch_button = QPushButton('Удалить филиал')
        self.delete_branch_button.setEnabled(False)
        self.delete_branch_button.clicked.connect(self.delete_branch)

        layout = QVBoxLayout()
        layout.addWidget(self.branch_list_widget)
        layout.addWidget(self.add_branch_button)
        layout.addWidget(self.edit_branch_button)
        layout.addWidget(self.delete_branch_button) 
        self.setLayout(layout)
        finished = QtCore.pyqtSignal()
        self.load_branches()

    def load_branches(self):
        connection = sqlite3.connect('AbelyashevTest.db')
        cursor = connection.cursor()
        cursor.execute('''
            SELECT b.BranchID, b.BranchesName, e.name, e.surname
            FROM Branches AS b
            LEFT JOIN Employees AS e ON b.CEO_ID = e.EmployeeID
        ''')
        branches = cursor.fetchall()
        connection.close()

        self.branch_list_widget.clear()
        if branches:
            for branch_id, branch_name, director_name, director_surname in branches:
                director = f"{director_name} {director_surname}" if director_name and director_surname else "Нет директора"
                self.branch_list_widget.addItem(f"{branch_name} (Директор: {director}) (ID: {branch_id})")
        else:
            QMessageBox.warning(self, 'Внимание', 'Нет доступных филиалов')


    def add_branch(self):
        branch_add_dialog = BranchAddDialog()
        if branch_add_dialog.exec_() == QDialog.Accepted:
            self.load_branches()

    def edit_branch(self):
        selected_item = self.branch_list_widget.currentItem()
        if selected_item:
            branch_id_str = selected_item.text().split('(ID:')[1].strip()[:-1]
            branch_id = int(branch_id_str)
            branch_edit_dialog = BranchEditDialog(branch_id)
            branch_edit_dialog.load_branch(branch_id)  
            if branch_edit_dialog.exec_() == QDialog.Accepted:
                self.load_branches()  



    def delete_branch(self):
        selected_item = self.branch_list_widget.currentItem()
        if selected_item:
            branch_id_str = selected_item.text().split('(ID:')[1].strip()[:-1]
            branch_id = int(branch_id_str)


            connection = sqlite3.connect('AbelyashevTest.db')
            cursor = connection.cursor()
            cursor.execute('SELECT COUNT(*) FROM Employees WHERE branch_id=?', (branch_id,))
            employee_count = cursor.fetchone()[0]

            if employee_count > 0:
                reply = QMessageBox.question(self, 'Удаление филиала',
                                            f'Вы уверены, что хотите удалить филиал? Вместе с ним удалятся {employee_count} сотрудник(ов).',
                                            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                if reply == QMessageBox.Yes:

                    cursor.execute('DELETE FROM Employees WHERE branch_id=?', (branch_id,))

                    cursor.execute('DELETE FROM Branches WHERE BranchID=?', (branch_id,))
                    connection.commit()

                    self.branch_list_widget.takeItem(self.branch_list_widget.row(selected_item))
            else:

                cursor.execute('DELETE FROM Branches WHERE BranchID=?', (branch_id,))
                connection.commit()
                self.branch_list_widget.takeItem(self.branch_list_widget.row(selected_item))

            connection.close()


    def enable_edit_button(self):
        if self.branch_list_widget.selectedItems():
            self.edit_branch_button.setEnabled(True)
            self.delete_branch_button.setEnabled(True)
        else:
            self.edit_branch_button.setEnabled(False)
            self.delete_branch_button.setEnabled(False)

class BranchEditDialog(QDialog):
    def __init__(self, branch_id=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Добавление/Редактирование филиала')

        self.mode = 'Добавление'
        self.branch_id = None
        self.director_id = None
        self.branch_name_edit = QLineEdit()
        self.director_combo = QComboBox()
        self.director_combo.currentIndexChanged.connect(self.update_director)
        self.director_combo.currentIndexChanged.connect(self.update_director)
        self.save_button = QPushButton('Сохранить')
        self.save_button.clicked.connect(self.save_branch)

        layout = QFormLayout()
        layout.addRow('Название филиала:', self.branch_name_edit)
        layout.addRow('Директор филиала:', self.director_combo)
        layout.addWidget(self.save_button)
        self.setLayout(layout)

        self.load_employees(branch_id)

    def load_employees(self, branch_id):
        connection = sqlite3.connect('AbelyashevTest.db')
        cursor = connection.cursor()
        cursor.execute('SELECT EmployeeID, name, surname FROM Employees WHERE branch_id=?', (branch_id,))
        employees = cursor.fetchall()
        connection.close()

        self.director_combo.clear()
        if employees:
            for employee_id, name, surname in employees:
                self.director_combo.addItem(f"{name} {surname}", userData=employee_id)
        else:
            self.director_combo.addItem('Нет сотрудников')
    def update_director(self, index): 
        if index >= 0:
            director_id = self.director_combo.currentData()
            self.director_id = director_id

    def load_branch(self, branch_id):
        self.branch_id = branch_id
        self.mode = 'Редактирование'
        connection = sqlite3.connect('AbelyashevTest.db')
        cursor = connection.cursor()
        cursor.execute('SELECT BranchesName, CEO_ID FROM Branches WHERE BranchID=?', (branch_id,))
        branch_data = cursor.fetchone()
        connection.close()

        if branch_data:
            self.branch_name_edit.setText(branch_data[0])
            director_index = self.director_combo.findData(branch_data[1])
            if director_index != -1:
                self.director_combo.setCurrentIndex(director_index)


    def save_branch(self):
        branch_name = self.branch_name_edit.text()
        director_id = self.director_id  
        
        if not branch_name:
            QMessageBox.warning(self, 'Предупреждение', 'Введите название филиала')
            return

        connection = sqlite3.connect('AbelyashevTest.db')
        cursor = connection.cursor()

        if self.mode == 'Добавление':
            cursor.execute('INSERT INTO Branches (BranchesName, CEO_ID) VALUES (?, ?)', (branch_name, director_id))
        elif self.mode == 'Редактирование':
            cursor.execute('UPDATE Branches SET BranchesName=?, CEO_ID=? WHERE BranchID=?', (branch_name, director_id, self.branch_id))

        connection.commit()
        connection.close()

        self.accept()



class BranchAddDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Добавление филиала')

        self.branch_name_edit = QLineEdit()

        self.save_button = QPushButton('Сохранить')
        self.save_button.clicked.connect(self.save_branch)

        layout = QFormLayout()
        layout.addRow('Название филиала:', self.branch_name_edit)
        layout.addWidget(self.save_button)
        self.setLayout(layout)

    def save_branch(self):
        branch_name = self.branch_name_edit.text()

        if not branch_name:
            QMessageBox.warning(self, 'Предупреждение', 'Введите название филиала')
            return

        connection = sqlite3.connect('AbelyashevTest.db')
        cursor = connection.cursor()

        cursor.execute('INSERT INTO Branches (BranchesName) VALUES (?)', (branch_name,))

        connection.commit()
        connection.close()

        self.accept()

class EmployeeEditDialog(QDialog):
    def __init__(self, employee_id, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Редактирование сотрудника')

        self.employee_id = employee_id

        self.name_label = QLabel('Имя:')
        self.name_edit = QLineEdit()

        self.surname_label = QLabel('Фамилия:')
        self.surname_edit = QLineEdit()

        self.branch_label = QLabel('Филиал:')
        self.branch_combo = QComboBox()

        self.save_button = QPushButton('Сохранить')
        self.save_button.clicked.connect(self.save_employee)

        layout = QVBoxLayout()
        layout.addWidget(self.name_label)
        layout.addWidget(self.name_edit)
        layout.addWidget(self.surname_label)
        layout.addWidget(self.surname_edit)
        layout.addWidget(self.branch_label)
        layout.addWidget(self.branch_combo)
        layout.addWidget(self.save_button)
        self.setLayout(layout)

        self.load_employee_data()
        self.load_branches()

    def load_employee_data(self):
        connection = sqlite3.connect('AbelyashevTest.db')
        cursor = connection.cursor()
        cursor.execute('SELECT name, surname, branch_id FROM Employees WHERE EmployeeID = ?', (self.employee_id,))
        employee_data = cursor.fetchone()
        connection.close()

        if employee_data:
            self.name_edit.setText(employee_data[0])
            self.surname_edit.setText(employee_data[1])
            branch_id = employee_data[2]
            index = self.branch_combo.findData(branch_id)
            if index != -1:
                self.branch_combo.setCurrentIndex(index)

    def load_branches(self):
        connection = sqlite3.connect('AbelyashevTest.db')
        cursor = connection.cursor()
        cursor.execute('SELECT BranchID, BranchesName FROM Branches')
        branches = cursor.fetchall()
        connection.close()

        for branch_id, branch_name in branches:
            self.branch_combo.addItem(branch_name, branch_id)

    def save_employee(self):
        name = self.name_edit.text()
        surname = self.surname_edit.text()
        branch_id = self.branch_combo.currentData()

        connection = sqlite3.connect('AbelyashevTest.db')
        try:
            cursor = connection.cursor()
            cursor.execute('UPDATE Employees SET name = ?, surname = ?, branch_id = ? WHERE EmployeeID = ?', (name, surname, branch_id, self.employee_id))
            connection.commit()
        except sqlite3.Error as e:
            print("Ошибка сохранения сотрудника:", e)
        finally:
            connection.close()

        self.accept()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Управление сотрудниками и филиалами')
        self.employee_list_widget = QListWidget()
        self.employee_list_widget.itemSelectionChanged.connect(self.enable_edit_button)

        self.add_employee_button = QPushButton('Добавить сотрудника')
        self.add_employee_button.clicked.connect(self.add_employee)

        self.edit_employee_button = QPushButton('Редактировать сотрудника')
        self.edit_employee_button.setEnabled(False)
        self.edit_employee_button.clicked.connect(self.edit_employee)

        self.branch_button = QPushButton('Список филиалов')
        self.branch_button.clicked.connect(self.open_branch_list_dialog)

        self.delete_employee_button = QPushButton('Удалить сотрудника')
        self.delete_employee_button.setEnabled(False)
        self.delete_employee_button.clicked.connect(self.delete_employee)

        self.save_data_button = QPushButton('Сохранить данные')
        self.save_data_button.clicked.connect(self.save_data_dialog)

        layout = QVBoxLayout()
        layout.addWidget(self.employee_list_widget)
        layout.addWidget(self.add_employee_button)
        layout.addWidget(self.edit_employee_button)
        layout.addWidget(self.delete_employee_button) 
        layout.addWidget(self.branch_button)
        layout.addWidget(self.save_data_button)

        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        self.employee_form = EmployeeForm(self)

        self.load_employees()

    def update_employee_list(self):
        self.load_employees()

    def open_branch_list_dialog(self):
        dialog = BranchListDialog(self)
        dialog.finished.connect(self.update_employee_list)
        dialog.exec_()

    def open_branch_list_dialog(self):
        dialog = BranchListDialog(self)
        dialog.finished.connect(self.update_employee_list)
        dialog.exec_()

    def save_data_dialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_name, filter_used = QFileDialog.getSaveFileName(self, "Сохранить данные", "", "Word Document (*.docx);;Excel File (*.xlsx)", options=options)
        if file_name:
            print(f"Выбранный путь для сохранения: {file_name}, используемый фильтр: {filter_used}")
            if "Word Document" in filter_used:
                print("Сохранение в формате DOCX")
                self.save_to_docx(file_name)
            elif "Excel File" in filter_used:
                print("Сохранение в формате XLSX")
                self.save_to_xlsx(file_name)
            else:
                print("Формат файла не распознан.")
        else:
            print("Сохранение отменено или путь не выбран.")


    def save_to_docx(self, file_path):
        print(f"Начало процесса сохранения DOCX в {file_path}")
        document = Document()
        employees_data = self.get_employees_data()
        for employee in employees_data:
            name, surname, branch, director = employee
            paragraph = document.add_paragraph(f"Имя: {name}, Фамилия: {surname}, Филиал: {branch}, Директор: {director}")
        document.save(file_path)
        print(f"Документ сохранён: {file_path}")

    def save_to_xlsx(self, file_path):
        if not file_path.endswith('.xlsx'):
            file_path += '.xlsx'

        employees_data = self.get_employees_data()

        df = pd.DataFrame(employees_data, columns=['Имя', 'Фамилия', 'Филиал', 'Директор'])

        df.to_excel(file_path, index=False)



    def get_employees_data(self):
        employees_data = []
        connection = sqlite3.connect('AbelyashevTest.db')
        cursor = connection.cursor()
        cursor.execute('''
            SELECT Employees.name, Employees.surname, Branches.BranchesName, Branches.CEO_ID
            FROM Employees 
            LEFT JOIN Branches ON Employees.branch_id = Branches.BranchID
        ''')
        employees = cursor.fetchall()
        connection.close()
        for employee in employees:
            name, surname, branch, ceo_id = employee
            ceo_name = self.get_employee_name_by_id(ceo_id)
            if ceo_id:
                director = ceo_name
            else:
                director = "Нет директора"
            employees_data.append((name, surname, branch, director))
        return employees_data

    def get_employee_name_by_id(self, employee_id):
        if employee_id is None:
            return "Нет директора"
        connection = sqlite3.connect('AbelyashevTest.db')
        cursor = connection.cursor()
        cursor.execute('SELECT name, surname FROM Employees WHERE EmployeeID = ?', (employee_id,))
        employee = cursor.fetchone()
        connection.close()
        if employee:
            return f"{employee[0]} {employee[1]}"
        else:
            return "Директор не найден"

    def delete_employee(self):
        selected_item = self.employee_list_widget.currentItem()
        if selected_item:
            employee_id = selected_item.data(Qt.UserRole)
            connection = sqlite3.connect('AbelyashevTest.db')
            cursor = connection.cursor()
            cursor.execute('SELECT * FROM Employees WHERE EmployeeID=?', (employee_id,))
            employee_data = cursor.fetchone()
            print("Данные сотрудника:", employee_data)
            if employee_data:
                employee_name = f"{employee_data[1]} {employee_data[2]}"
                branch_id = employee_data[3]
                cursor.execute('SELECT CEO_ID FROM Branches WHERE BranchID=?', (branch_id,))
                director_id = cursor.fetchone()[0]
                if employee_id == director_id:
                    QMessageBox.warning(self, 'Предупреждение', 'Филиал остался без директора')
            cursor.execute('DELETE FROM Employees WHERE EmployeeID=?', (employee_id,))
            connection.commit()
            connection.close()
            self.load_employees()

    def load_employees(self):
        connection = sqlite3.connect('AbelyashevTest.db')
        cursor = connection.cursor()
        cursor.execute('''
            SELECT Employees.EmployeeID, Employees.name, Employees.surname, Branches.BranchesName, Branches.CEO_ID
            FROM Employees 
            LEFT JOIN Branches ON Employees.branch_id = Branches.BranchID
        ''')
        employees = cursor.fetchall()
        connection.close()
        self.employee_list_widget.clear()
        if employees:
            for employee_id, name, surname, branch_name, ceo_id in employees:
                ceo_name = self.get_employee_name_by_id(ceo_id)
                if ceo_id:
                    employee_info = f"{name} {surname} (Филиал: {branch_name}, Директор: {ceo_name})"
                else:
                    employee_info = f"{name} {surname} (Филиал: {branch_name})"
                item = QListWidgetItem(employee_info)
                item.setData(Qt.UserRole, employee_id)
                self.employee_list_widget.addItem(item)
        else:
            QMessageBox.warning(self, 'Внимание', 'Нет доступных сотрудников')




    def edit_employee(self):
        selected_item = self.employee_list_widget.currentItem()
        if selected_item:
            employee_name = selected_item.text().split('(Филиал:')[0].strip()
            connection = sqlite3.connect('AbelyashevTest.db')
            cursor = connection.cursor()
            cursor.execute('SELECT EmployeeID FROM Employees WHERE name || " " || surname = ?', (employee_name,))
            employee_id = cursor.fetchone()[0]
            connection.close()

            edit_dialog = EmployeeEditDialog(employee_id)
            if edit_dialog.exec_() == QDialog.Accepted:
                self.load_employees()


    def add_employee(self):
        self.employee_form.load_branches()
        self.employee_form.show()

    def enable_edit_button(self):
        if self.employee_list_widget.selectedItems():
            self.edit_employee_button.setEnabled(True)
            self.delete_employee_button.setEnabled(True)
        else:
            self.edit_employee_button.setEnabled(False)
            self.delete_employee_button.setEnabled(False)
        self.load_branches()



if __name__ == '__main__':
    app = QApplication(sys.argv)
    init_db()
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
