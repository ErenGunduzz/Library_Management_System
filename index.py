from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
import sys
from PyQt5.uic import loadUiType
import mysql.connector
import datetime
from xlrd import *
from xlsxwriter import *

ui, _ = loadUiType('library.ui')

class MainApp(QMainWindow, ui):
    def __init__(self):
        QMainWindow.__init__(self)
        self.setupUi(self)
        self.handle_buttons()


    def handle_buttons(self):
        self.pushButton.clicked.connect(self.open_day_to_day_tab)
        self.pushButton_2.clicked.connect(self.open_books_tab)
        self.pushButton_4.clicked.connect(self.open_settings_tab)

        # book operations
        self.pushButton_7.clicked.connect(self.add_new_book)
        self.pushButton_12.clicked.connect(self.search_books_by_id)
        self.pushButton_8.clicked.connect(self.edit_books)
        self.pushButton_10.clicked.connect(self.delete_books)

        #category operations
        self.pushButton_16.clicked.connect(self.add_category)

        self.pushButton_6.clicked.connect(self.handle_day_operations)
        self.pushButton_5.clicked.connect(self.show_all_books)
        self.pushButton_3.clicked.connect(self.show_category)
        self.pushButton_9.clicked.connect(self.delete_category)
        self.pushButton_11.clicked.connect(self.search_books_by_name)
        self.pushButton_13.clicked.connect(self.search_books_by_author)
        self.pushButton_14.clicked.connect(self.search_books_by_category)

        #### exporting to excel ####
        self.pushButton_15.clicked.connect(self.export_day_operations)
        self.pushButton_17.clicked.connect(self.export_books)


    ##opening tabs##
    def open_day_to_day_tab(self):
        self.tabWidget.setCurrentIndex(0)


    def open_books_tab(self):
        self.tabWidget.setCurrentIndex(1)

    def open_settings_tab(self):
        self.tabWidget.setCurrentIndex(2)

    ##day operations##
    def handle_day_operations(self):
        book_id = self.lineEdit_13.text()
        book_title = self.lineEdit_14.text()
        client_name = self.lineEdit.text()
        type = self.comboBox.currentText()
        days_number = self.comboBox_2.currentIndex() + 1
        today_date = datetime.date.today()
        to_date = today_date + datetime.timedelta(days=days_number)

        print(today_date)
        print(to_date)

        self.db = mysql.connector.connect(host='localhost', username='root', passwd='root11235813*', db='library')
        self.cur = self.db.cursor()
        self.cur.execute('''
        INSERT INTO dayoperations(id, book_name, client_name, type, days, date, to_date)
        VALUES (%s, %s , %s , %s, %s , %s , %s)
        ''' , (book_id, book_title, client_name, type, days_number, today_date, to_date))

        print("deneme1")
        self.db.commit()
        self.statusBar().showMessage('New Operation Added')


        self.show_all_operations()


    def show_all_operations(self):
        self.db = mysql.connector.connect(host='localhost', username='root', passwd='root11235813*', db='library')
        self.cur = self.db.cursor()

        self.cur.execute('''
        SELECT id, book_name, client_name, type, days, date, to_date FROM dayoperations
        ''')

        data = self.cur.fetchall()

        print(data)

        self.tableWidget.setRowCount(0)
        self.tableWidget.insertRow(0)
        for row, form in enumerate(data):
            for column, item in enumerate(form):
                self.tableWidget.setItem(row, column, QTableWidgetItem(str(item)))
                column += 1

            row_position = self.tableWidget.rowCount()
            self.tableWidget.insertRow(row_position)

    #####books#####
    def add_new_book(self):
        self.db = mysql.connector.connect(host='localhost', username='root', passwd='root11235813*', db='library')
        self.cur = self.db.cursor()

        id = self.lineEdit_2.text()
        book_name = self.lineEdit_3.text()
        book_desc = self.plainTextEdit.toPlainText()
        book_category = self.comboBox_3.currentText()
        book_author = self.lineEdit_8.text()
        book_publisher = self.lineEdit_16.text()
        book_price = self.lineEdit_4.text()

        self.cur.execute('''
            INSERT INTO book(id, book_name, book_desc, book_category,
             book_author, book_publisher, book_price)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
            ''' ,(id, book_name, book_desc, book_category, book_author, book_publisher, book_price))

        print("deneme2")
        self.db.commit()
        self.statusBar().showMessage('New book added')

        self.lineEdit_2.setText('')
        self.lineEdit_3.setText('')
        self.plainTextEdit.setPlainText('')
        self.comboBox_3.setCurrentIndex(-1)
        self.lineEdit_8.setText('')
        self.lineEdit_16.setText('')
        self.lineEdit_4.setText('')


    def show_all_books(self):
        self.db = mysql.connector.connect(host='localhost', username='root',
                                          passwd='root11235813*', db='library')
        self.cur = self.db.cursor()

        self.cur.execute('''
        SELECT * FROM book
         ''')
        data = self.cur.fetchall()

        self.tableWidget_4.setRowCount(0)
        self.tableWidget_4.insertRow(0)

        for row, form in enumerate(data):
            for column, item in enumerate(form):
                self.tableWidget_4.setItem(row, column, QTableWidgetItem(str(item)))
                column += 1
            row_position = self.tableWidget_4.rowCount()
            self.tableWidget_4.insertRow(row_position)

        self.db.close()


    def search_books_by_id(self):
        self.db = mysql.connector.connect(host='localhost', username='root', passwd='root11235813*', db='library')
        self.cur = self.db.cursor()

        book_id = self.lineEdit_15.text()

        self.cur.execute('''SELECT * FROM book WHERE id = %s''', [book_id])

        data = self.cur.fetchall()
        print(data)

        self.tableWidget_4.setRowCount(0)
        self.tableWidget_4.insertRow(0)

        for row, form in enumerate(data):
            for column, item in enumerate(form):
                self.tableWidget_4.setItem(row, column, QTableWidgetItem(str(item)))
                column += 1
            row_position = self.tableWidget_4.rowCount()
            self.tableWidget_4.insertRow(row_position)
        self.db.close()

    def search_books_by_name(self):
        self.db = mysql.connector.connect(host='localhost', username='root', passwd='root11235813*', db='library')
        self.cur = self.db.cursor()

        book_name = self.lineEdit_9.text()

        self.cur.execute('''SELECT * FROM book WHERE book_name = %s''', [book_name])

        data = self.cur.fetchall()
        print(data)

        self.tableWidget_4.setRowCount(0)
        self.tableWidget_4.insertRow(0)

        for row, form in enumerate(data):
            for column, item in enumerate(form):
                self.tableWidget_4.setItem(row, column, QTableWidgetItem(str(item)))
                column += 1
            row_position = self.tableWidget_4.rowCount()
            self.tableWidget_4.insertRow(row_position)
        self.db.close()


    def search_books_by_author(self):
        self.db = mysql.connector.connect(host='localhost', username='root', passwd='root11235813*', db='library')
        self.cur = self.db.cursor()

        author = self.lineEdit_10.text()

        self.cur.execute('''SELECT * FROM book WHERE book_author = %s''', [author])

        data = self.cur.fetchall()
        print(data)

        self.tableWidget_4.setRowCount(0)
        self.tableWidget_4.insertRow(0)

        for row, form in enumerate(data):
            for column, item in enumerate(form):
                self.tableWidget_4.setItem(row, column, QTableWidgetItem(str(item)))
                column += 1
            row_position = self.tableWidget_4.rowCount()
            self.tableWidget_4.insertRow(row_position)
        self.db.close()

    def search_books_by_category(self):
        self.db = mysql.connector.connect(host='localhost', username='root', passwd='root11235813*', db='library')
        self.cur = self.db.cursor()

        category = self.lineEdit_11.text()

        self.cur.execute('''SELECT * FROM book WHERE book_category = %s''', [category])

        data = self.cur.fetchall()
        print(data)

        self.tableWidget_4.setRowCount(0)
        self.tableWidget_4.insertRow(0)

        for row, form in enumerate(data):
            for column, item in enumerate(form):
                self.tableWidget_4.setItem(row, column, QTableWidgetItem(str(item)))
                column += 1
            row_position = self.tableWidget_4.rowCount()
            self.tableWidget_4.insertRow(row_position)
        self.db.close()

    def edit_books(self):
        self.db = mysql.connector.connect(host='localhost', username='root', passwd='root11235813*', db='library')
        self.cur = self.db.cursor()

        id = self.lineEdit_6.text()
        book_name = self.lineEdit_5.text()
        book_desc = self.plainTextEdit_2.toPlainText()
        book_category = self.comboBox_6.currentText()
        book_author = self.lineEdit_17.text()
        book_publisher = self.lineEdit_18.text()
        book_price = self.lineEdit_7.text()

        self.cur.execute('''
            UPDATE book SET book_name=%s, 
            book_desc=%s, book_category=%s, book_author=%s,
            book_publisher=%s, book_price=%s WHERE id=%s
            ''', (book_name, book_desc, book_category,
                  book_author, book_publisher, book_price, id))

        self.db.commit()
        self.statusBar().showMessage('Book updated')

        self.lineEdit_6.setText('')
        self.lineEdit_5.setText('')
        self.plainTextEdit_2.setPlainText('')
        self.comboBox_6.setCurrentIndex(-1)
        self.lineEdit_17.setText('')
        self.lineEdit_18.setText('')
        self.lineEdit_7.setText('')


    def delete_books(self):
        self.db = mysql.connector.connect(host='localhost', username='root', passwd='root11235813*', db='library')
        self.cur = self.db.cursor()

        book_id = self.lineEdit_6.text()

        warning = QMessageBox.warning(self, 'Delete Book', "Are you sure you want to delete this book?",
                                      QMessageBox.Yes | QMessageBox.No)
        if warning == QMessageBox.Yes:
            sql = ''' DELETE FROM book WHERE id  = %s '''
            self.cur.execute(sql, [book_id])
            self.db.commit()
            self.statusBar().showMessage('Book deleted')

            self.show_all_books()

    ########## settings #########
    def add_category(self):

        self.db = mysql.connector.connect(host='localhost', username='root', passwd='root11235813*', db='library')
        self.cur = self.db.cursor()

        category_name = self.lineEdit_29.text()

        self.cur.execute('''
            INSERT INTO category(category_name) VALUES (%s)
            ''', [category_name])
        print("deneme3")

        self.db.commit()
        self.statusBar().showMessage('New category added')
        self.lineEdit_29.setText('')

    def show_category(self):
        self.db = mysql.connector.connect(host='localhost', username='root', passwd='root11235813*', db='library')
        self.cur = self.db.cursor()

        self.cur.execute('''SELECT * FROM category''')
        data = self.cur.fetchall()
        print(data)

        self.tableWidget_2.setRowCount(0)
        self.tableWidget_2.insertRow(0)

        for row, form in enumerate(data):
            for column, item in enumerate(form):
                self.tableWidget_2.setItem(row, column, QTableWidgetItem(str(item)))
                column += 1
            row_position = self.tableWidget_2.rowCount()
            self.tableWidget_2.insertRow(row_position)

        self.db.commit()
        self.db.close()

    def delete_category(self):
        self.db = mysql.connector.connect(host='localhost', username='root', passwd='root11235813*', db='library')
        self.cur = self.db.cursor()

        category_name = self.lineEdit_29.text()

        warning = QMessageBox.warning(self, 'Delete Category', "Are you sure you want to delete this category?",
                                      QMessageBox.Yes | QMessageBox.No)

        if warning == QMessageBox.Yes:
            sql = ''' DELETE FROM category WHERE category_name = %s'''
            self.cur.execute(sql, [category_name])
            self.db.commit()
            self.statusBar().showMessage('Category deleted')
            self.show_category()


    def show_author(self):
        self.db = mysql.connector.connect(host='localhost', username='root', passwd='root11235813*', db='library')
        self.cur = self.db.cursor()

        self.cur.execute(''' SELECT author_name FROM authors''')
        data = self.cur.fetchall()

        if data:
            self.tableWidget_3.setRowCount(0)
            self.tableWidget_3.insertRow(0)
            for row, form in enumerate(data):
                for column, item in enumerate(form):
                    self.tableWidget_3.setItem(row, column, QTableWidgetItem(str(item)))
                    column += 1

                row_position = self.tableWidget_3.rowCount()
                self.tableWidget_3.insertRow(row_position)


#### export data ####

    def export_day_operations(self):
        self.db = mysql.connector.connect(host='localhost', username='root', passwd='root11235813*', db='library')
        self.cur = self.db.cursor()

        self.cur.execute(''' 
                    SELECT id, book_name, client_name, type, days, date, to_date FROM dayoperations
                ''')

        data = self.cur.fetchall()
        wb = Workbook('day_operations.xlsx')
        sheet1 = wb.add_worksheet()

        sheet1.write(0, 0, 'book ISBN')
        sheet1.write(0, 1, 'book_name')
        sheet1.write(0, 2, 'client_name')
        sheet1.write(0, 3, 'type')
        sheet1.write(0, 4, 'days')
        sheet1.write(0, 5, 'date')
        sheet1.write(0, 6, 'to_date')

        row_number = 1
        for row in data:
            column_number = 0
            for item in row:
                sheet1.write(row_number, column_number, str(item))
                column_number += 1
            row_number += 1

        wb.close()
        self.statusBar().showMessage('Report created succesfully!')

    def export_books(self):
        self.db = mysql.connector.connect(host='localhost', username='root', passwd='root11235813*', db='library')
        self.cur = self.db.cursor()

        self.cur.execute(
            ''' SELECT id, book_name, book_author, book_category,
             book_price, book_publisher, book_desc FROM book''')
        data = self.cur.fetchall()

        wb = Workbook('all_books.xlsx')
        sheet1 = wb.add_worksheet()

        sheet1.write(0, 0, 'Book ISBN')
        sheet1.write(0, 1, 'Book Name')
        sheet1.write(0, 2, 'Book Author')
        sheet1.write(0, 3, 'Book Category')
        sheet1.write(0, 4, 'Book Price')
        sheet1.write(0, 5, 'Book Publisher')
        sheet1.write(0, 6, 'Book Description')

        row_number = 1
        for row in data:
            column_number = 0
            for item in row:
                sheet1.write(row_number, column_number, str(item))
                column_number += 1
            row_number += 1

        wb.close()
        self.statusBar().showMessage('Book Report Created Successfully')


def main():
    app = QApplication(sys.argv)
    window = MainApp()
    window.show()
    app.exec_()

if __name__ == '__main__':
    main()



