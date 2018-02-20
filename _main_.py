from PyQt5 import QtWidgets
import sys
from classes import Oferta


def main():
    app = QtWidgets.QApplication(sys.argv)  # A new instance of QApplication
    progr = Oferta()  # We set the form to be our ExampleApp (design)
    progr.show()  # Show the f
    app.exec_()  # and execute the app
    pass


if __name__ == '__main__':

    main()
