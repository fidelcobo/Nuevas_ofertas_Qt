
from PyQt5 import QtWidgets
from main_screen import Ui_Oferta_Manual
from maint_screen import Ui_maintenance_Dialog
from procedures import convert_sla, diff_days, fill_offer
from listas import list_uptime_descr, list_manufacturer, list_tech, list_uptime_codes
import locale

locale.setlocale(locale.LC_ALL, 'FR')


class Articulo:

    def __init__(self, unit= 0, tech='', manufacturer='', code='', descripcion_prod= '', list_price = 0.0,
                 coste_prod=0.0, margin_prod=0.0, maintenance= False, init_date='', end_date='', durac= 0,
                 sku_uptime='', descr_uptime='', backout_name='', descr_backout = '', list_price_back=0.0,
                 coste_unit_back=0.0, uplift=0.0, cost_unit_manten=0.0, margen_mant=0.0, venta_mant=0.0):

        self.unit = int(unit)
        self.tech = tech
        self.manufacturer = manufacturer
        self.code = code
        self.descripcion_prod = descripcion_prod
        self.coste_prod = float(coste_prod)
        self.list_price = list_price
        self.margin_prod = float(margin_prod)
        self.venta_prod = round(float(coste_prod/(1-(margin_prod/100))), 2)
        self.maintenance = maintenance
        self.init_date = init_date
        self.end_date = end_date
        self.sku_uptime = sku_uptime
        self.durac = int(durac)
        self.descr_uptime = descr_uptime
        self.backout_name = backout_name
        self.descr_backout = descr_backout
        self.list_price_back = float(list_price_back)
        self.coste_unit_back = float(coste_unit_back)
        self.uplift = float(uplift)
        self.cost_unit_manten = float(cost_unit_manten)
        self.margen_mant = float(margen_mant)
        self.venta_mant = float(venta_mant)

    @staticmethod
    def fill(*args):
        unit = args[0]
        tech = args[1]
        manufacturer = args[2]
        code = args[3]
        coste_prod = args[4]
        margen_prod = args[5]
        venta_prod = args[6]
        manten = args[7]
        init_date = args[8]
        end_date = args[9]
        durac = args[10]
        sku_uptime = args[11]
        descr_uptime = args[12]
        backout_name = args[13]
        list_price_back = args[14]
        coste_unit_back = args[15]
        uplift = args[16]
        cost_unit_manten = args[17]
        margen_mant = args[18]
        venta_mant = args[19]
        return Articulo(unit, tech, manufacturer, code, coste_prod, margen_prod, venta_prod, manten, init_date,
                        end_date, durac, sku_uptime, descr_uptime, backout_name, list_price_back, coste_unit_back,
                        uplift, cost_unit_manten, margen_mant, venta_mant)


class Datos_Mantenimiento(QtWidgets.QDialog, Ui_maintenance_Dialog):

    def __init__(self):
        super(self.__class__, self).__init__()
        self.setupUi(self)
        #
        # self.back_code = ''
        # self.margin = ''
        # self.uplift = ''
        # self.list_back = ''
        # self.cost_back = ''
        # self.back_descr = ''
        # self.pvp = ''
        self.lista_precios = []

        self.sla_uptime.addItems(list_uptime_descr)  # Rellenamos lista desplegable de SLAs

        self.ok_button.clicked.connect(self.process_data)
        self.clear_button.clicked.connect(self.exit_no_maintenance)

    def exit_no_maintenance(self):
        self.lista_precios.append([False, 0, 0, 0, 0])  # Se dice que no hay datos de mantenimiento
        self.close()

    def process_data(self):
        margin = self.margin.text()
        uplift = self.uplift.text()
        list_back = self.list_back.text()
        cost_back = self.cost_back.text()

        prices = self.verify_numeric_values(margin, uplift, list_back, cost_back)
        if prices[0]: # Los valores son coherentes
            self.lista_precios.append(prices)
            print(prices)
            self.close()

    def verify_numeric_values (self, margin, uplift, list_back, cost_back):

        # Verificamos la validez del valor de valor de lista de backout
        if list_back:
            list_back = list_back.replace(',', '.')
            try:
                list_back = float(list_back)
            except ValueError:
                QtWidgets.QMessageBox.critical(self, 'Error', 'Valor de precio de lista incorrecto')
                return False, list_back, cost_back, margin, uplift
        else:  # Precio de lista vacío. Se deja así
            list_back = 0

        # Verificamos la validez del valor de valor de coste de backout
        cost_back = cost_back.replace(',', '.')
        try:
            cost_back = float(cost_back)
        except ValueError:
            QtWidgets.QMessageBox.critical(self, 'Error', 'Valor de coste incorrecto')
            return False, list_back, cost_back, margin, uplift

        # Verificamos la validez del valor de margen
        margin = margin.replace(',', '.')
        try:
            margin = float(margin)
            if margin >= 100:
                QtWidgets.QMessageBox.critical(self, 'Error', 'Valor de margen demasiado alto')
                return (False, list_back, cost_back, margin, uplift)
        except ValueError:
            QtWidgets.QMessageBox.critical(self, 'Error', 'Valor de margen incorrecto')
            return False, list_back, cost_back, margin, uplift

        # Verificamos la validez del valor de uplift
        uplift = uplift.replace(',', '.')
        try:
            uplift = float(uplift)
        except ValueError:
            QtWidgets.QMessageBox.critical(self, 'Error', 'Valor de uplift incorrecto')
            return False, list_back, cost_back, margin, uplift

        return True, list_back, cost_back, margin, uplift


class Oferta(QtWidgets.QDialog, Ui_Oferta_Manual):

    def __init__(self):
        super(self.__class__, self).__init__()
        self.setupUi(self)

        # self.lista_checkpoint = []
        # self.lista_fortinet = []
        # self.lista_hp = []
        # self.lista_riverbed = []
        # self.lista_f5 = []
        # self.lista_cisco = []
        # self.lista_alcatel = []
        # self.lista_paloalto = []
        # self.lista_juniper = []
        # self.lista_bluecoat = []
        # self.lista_brocade = []

        self.lista_articulos = []

        self.manufacturer.addItems(list_manufacturer)
        self.tech.addItems(list_tech)
        self.dict_sla = dict((x,y) for (x, y) in zip(list_uptime_descr, list_uptime_codes))

        self.add_button.clicked.connect(self.add_item)
        self.add_finish_button.clicked.connect(self.complete_offer)

    def add_item(self):
        manuf = self.manufacturer.currentText()
        qty = self.qty.text()
        code = self.code.text()
        code = code.replace('\n', '')  # Si viene de un Excel puede traer un CR
        tech = self.tech.currentText()
        description = self.description.text()
        description = description.replace('\n', '')  # Si viene de un Excel puede traer un CR
        list_price = self.list_price.text()
        cost = self.cost.text()
        margin = self.margin.text()

        # De momento dejamos en blanco los datos de mantenimiento
        sku_uptime, descr_uptime, backout_name, descr_backout = '', '', '', ''
        init_date, end_date = '', ''
        list_price_back, coste_unit_back, uplift, cost_unit_manten, margen_maint, \
        venta_mant, durac_months = 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0

        if not code:  # El campo de código es obligatorio
            QtWidgets.QMessageBox.critical(self, 'Error', 'El campo de código\n no puede estar vacío')
            return

        if self.maintenance.isChecked():
            maintenance = True
        else:
            maintenance = False

        ok, list_price, cost, margin, qty = self.verify_numeric_values(list_price, cost, margin, qty)

        if ok:
            print('Tutti contenti')
            print(manuf, code, qty, description, list_price, cost, margin, maintenance)

            if maintenance:  # Hay que recoger los datos de mantenimiento
                prod = Datos_Mantenimiento()  # Abrimos una nueva pantalla para recoger estos datos
                prod.exec_()
                prod.show()
                try:
                    ok = prod.lista_precios[0][0]  # Es una lista que consta de un sólo elemento tipo lista
                    if ok:  # Los datos de mantenimiento son coherentes
                        list_price_back = prod.lista_precios[0][1]
                        if not list_price_back:
                            list_price_back = 0
                        coste_unit_back = prod.lista_precios[0][2]
                        margen_maint = prod.lista_precios[0][3]
                        uplift = prod.lista_precios[0][4]
                        cost_unit_manten = round(coste_unit_back * (1 + uplift/100), 2)
                        venta_mant = round(cost_unit_manten/(1 - margen_maint/100), 2)
                        backout_name = prod.back_code.text()
                        descr_backout = prod.back_descr.text()
                        descr_uptime = prod.sla_uptime.currentText()
                        sku_uptime = convert_sla(prod.sla_uptime.currentText(), self)
                        init_date = prod.start_date.text()
                        end_date = prod.end_date.text()
                        durac_months = int(round(12 * diff_days(init_date, end_date) / 365, 0))
                        print(type(durac_months), durac_months)
                    else:
                        maintenance = False  # No hay datos de mantenimiento buenos
                except IndexError:
                    QtWidgets.QMessageBox.warning(self, 'Aviso', 'Los datos de mantenimiento del último producto\n no se incluirán en la oferta')

            self.lista_articulos.append(Articulo(qty, tech, manuf, code, description, list_price, cost, margin, maintenance, init_date, end_date, durac_months,
                                                 sku_uptime, descr_uptime, backout_name, descr_backout, list_price_back, coste_unit_back, uplift,
                                                 cost_unit_manten, margen_maint, venta_mant))
            print(len(self.lista_articulos))
            self.clear_values()

    def complete_offer(self):

        # clasificar_articulos(self.lista_articulos, self)
        destination_folder = QtWidgets.QFileDialog.getExistingDirectory(self, "Elegir carpeta de destino")
        fill_offer(self.lista_articulos, destination_folder, self)  # Clasifica por fabricantes y hace el csv
        # hacer_oferta_ms(self.lista_articulos, destination_folder, self)
        QtWidgets.QMessageBox.information(self, 'OK',
                                          'Todo parece haber ido bien')

        self.lista_articulos = []   # Limpiamos la lista de artículos ya pasados a oferta

    def clear_values(self):

        self.code.setText('')
        self.maintenance.setCheckable(True)
        self.margin.setText('')
        self.list_price.setText('')
        self.cost.setText('')
        self.description.setText('')
        self.qty.setText('')
        self.maintenance.setChecked(False)

    def verify_numeric_values(self, list_price, cost, margin, qty):

        try:
            qty = int(qty)
        except ValueError:
            QtWidgets.QMessageBox.critical(self, 'Error', 'Valor de cantidad incorrecto')
            return(False, list_price, cost, margin, qty)

        if list_price:
            list_price = list_price.replace(',', '.')
            try:
                list_price = float(list_price)
            except ValueError:
                QtWidgets.QMessageBox.critical(self, 'Error', 'Valor de precio de lista incorrecto')
                return(False, list_price, cost, margin, qty)
        else:  # Precio de lista vacío. Se deja así
            list_price = 0

        cost = cost.replace(',', '.')
        try:
            cost = float(cost)
        except ValueError:
            QtWidgets.QMessageBox.critical(self, 'Error', 'Valor de coste incorrecto')
            return(False, list_price, cost, margin, qty)

        margin = margin.replace(',', '.')
        try:
            margin = float(margin)
            if margin >= 100:
                QtWidgets.QMessageBox.critical(self, 'Error', 'Valor de margen demasiado alto')
                return (False, list_price, cost, margin, qty)

        except ValueError:
            QtWidgets.QMessageBox.critical(self, 'Error', 'Valor de margen incorrecto')
            return(False, list_price, cost, margin, qty)

        return (True, list_price, cost, margin, qty)   # Los valores son correctos
