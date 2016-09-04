#-*- coding: utf-8 -*-
from PyQt5 import QtGui
from PyQt5 import QtWidgets# Import the PyQt4 module we'll need
import sys  # We need sys so that we can pass argv to QApplication
import design  # This file holds our MainWindow and all design related things
import csv, sqlite3
from openpyxl import load_workbook



class ExampleApp(QtWidgets.QMainWindow, design.Ui_MainWindow):


    def __init__(self):

        super(self.__class__, self).__init__()
        self.setupUi(self)
        #list of comboboBoxes in layout
        #need to be exaclty the same as in GUI
        #need to be exactly the same as layout of tables in data_sheet.csv
        comboboxes_list = [self.comboBox, self.comboBox_2, self.comboBox_3, self.comboBox_4, self.comboBox_5,
                           self.comboBox_6, self.comboBox_7, self.comboBox_8, self.comboBox_9]
        labels_list = [self.model, self.uchwyty, self.zasilanie, self.kontrola,
                       self.ochrona, self.ilosc_1, self.typ_1, self.ilosc_2,
                       self.typ_2, self.index_short, self.index_long]
        dict_list = self.populate_data_in_dicts('data_sheet.csv')
        self.fill_comboBoxes(dict_list, comboboxes_list)
        self.make_comboBoxes_signals(dict_list, comboboxes_list)
        self.pushButton.clicked.connect(lambda: self.update_xlsx('choosen_data.xlsx',

                                                                 comboboxes_list, labels_list))
        # self.show_hide()
    # def test_run(self):
    #     voltage_desc_list = ['None', '230 Volts List', '120 Volts List']
    #     voltage_code_list = ['', '230', '120']
    #     plugs_desc_list = ['one plug', 'two plugs', 'three plugs']
    #     plugs_code_list = ['1', '2', '3']
    #     self.comboBox.addItems(voltage_desc_list)
    #     self.comboBox_2.addItems(plugs_desc_list)
    #     self.comboBox.currentIndexChanged.connect(self.lineEditTextSetter)
    #     self.comboBox_2.currentIndexChanged.connect(self.lineEditTextSetter)
    #     self.pushButton.clicked.connect(lambda: self.update_xlsx('spreadsheet_to_insert.xlsx', self.lineEdit.text()))

    # def lineEditTextSetter(self):
    #     voltage_desc_list = ['None', '230 Volts List', '120 Volts List']
    #     voltage_code_list = ['', '230', '120']
    #     voltage_desc = self.comboBox.currentText()
    #     voltage_code = voltage_code_list[voltage_desc_list.index(str(voltage_desc))]
    #
    #     plugs_desc_list = ['one plug', 'two plugs', 'three plugs']
    #     plugs_code_list = ['1', '2', '3']
    #     plugs_desc = self.comboBox_2.currentText()
    #     plugs_code = plugs_code_list[plugs_desc_list.index(str(plugs_desc))]
    #
    #     self.lineEdit.setText('{0}-{1}'.format(voltage_code, plugs_code))
    #     data = self.lineEdit.text()
    #     # return data

    def show_hide(self):
        print(self.comboBox.currentText())
        if self.comboBox.currentText() == 'NPM-V':
            self.comboBox_2.show()
        else:
            self.comboBox_2.hide()

    def update_xlsx(self, dest, comboboxes_list, labels_list):
        # Open an xlsx for reading
        wb = load_workbook(filename=dest)
        # Get the current Active Sheet
        ws = wb.get_active_sheet()
        count_a = 0

        #get labels
        for label in labels_list:
            count_a += 1
            ws.cell('A{0}'.format(count_a)).value = label.text()

        #get comboboxesvalues
        count_b = 0
        for combo in comboboxes_list:
            count_b += 1
            ws.cell('B{0}'.format(count_b)).value = combo.currentText()
        #get linesvalues
        count_b += 1
        ws.cell('B{0}'.format(count_b)).value = self.lineEdit.text()
        count_b += 1
        ws.cell('B{0}'.format(count_b)).value = self.lineEdit_2.text()

        wb.save('choosen_data.xlsx')

    def populate_data_in_dicts(self, csv_file):
        #DICTIONARIES
        model_listwy = {}
        uchwyty = {}
        zasilanie_na_wejsciu = {}
        kontrola_pdu = {}
        ochrona_pdu = {}
        ilosc_gniazd_i_typu = {}
        i_typ_gniazd = {}
        ilosc_gniazd_ii_typu = {}
        ii_typ_gniazd = {}

        dicts_list = [model_listwy, uchwyty, zasilanie_na_wejsciu, kontrola_pdu,
                      ochrona_pdu, ilosc_gniazd_i_typu, i_typ_gniazd, ilosc_gniazd_ii_typu, ii_typ_gniazd]
        csv_lines = []
        with open(csv_file, 'r') as f:
            reader = csv.reader(f)
            for row in reader:
                csv_lines.append(row)
            # print(csv_lines)

            for dictionary in dicts_list:
                count = 0
                for item in csv_lines:
                    count += 1
                    if count > 2:
                        dictionary[item[0]] = item[1]
                        del item[0:3]
                    else:
                        pass
        return dicts_list

    def fill_comboBoxes(self, dicts_list, comboboxes_list):
        i = 0
        for combo in comboboxes_list:
            combo.addItems(sorted(dicts_list[i].keys()))
            i += 1

    def make_comboBoxes_signals(self,dicts_list, comboboxes_list):

        for combo in comboboxes_list:
            combo.currentIndexChanged.connect(lambda: self.make_indexes(dicts_list, comboboxes_list))

    def make_indexes(self, dicts_list, comboboxes_list):

        long_index = ''
        short_index = ''
        count = 0
        for combo, dict in zip(comboboxes_list, dicts_list):
            count += 1
            short_index += dict[combo.currentText()]
            if count == 5:
                short_index += '.'
            elif count == 6:
                short_index += '-'
            elif count == 7:
                short_index += ','
            elif count == 8:
                short_index += '-'
        self.lineEdit.setText(short_index)

        if self.comboBox.currentText() == 'PDU - Basic':
            zero = 'zasilająca'
        else:
            zero = 'zarządzalna'
        # print(self.comboBox_5.currentText())
        long_index += 'Listwa {0} pionowa'.format(zero)
        if '400' in self.comboBox_3.currentText():
            long_index += ' {0}'.format('trójfazowa')
        else:
            pass
        long_index += ' BKT {0} typ {1} x {2}'.format(self.comboBox.currentText(),
                                                     self.comboBox_6.currentText(),
                                                     self.comboBox_7.currentText())
        # print(type(self.comboBox_8.currentText()))
        if self.comboBox_8.currentText() != '00':
            long_index += ' + {0} x {1}'.format(self.comboBox_8.currentText(), self.comboBox_9.currentText())
        else:
            pass
        long_index += ' , {0}'.format(self.comboBox_3.currentText())

        self.lineEdit_2.setText(long_index)


        if self.comboBox.currentText() == 'NPM-V':
            self.comboBox_2.setEnabled(True)
        else:
            self.comboBox_2.setEnabled(False)



def main():
    app = QtWidgets.QApplication(sys.argv)
    form = ExampleApp()
    form.show()
    app.exec_()  # and execute the app


if __name__ == '__main__':
    main()