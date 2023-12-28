import sys
import wmi
from PyQt5.QtWidgets import *

OS = processors  = Ram = motherboard = {}
OS_properties=processor_properties=Ram_properties=motherboard_properties=[]

#Создание класса для оконного пользовательского интерфейса
class MyWindow(QMainWindow):

    def __init__(self, OS, OS_properties, processors, processor_properties, Ram, Ram_properties, motherboard, motherboard_properties):
        # def __init__(self):
        super().__init__()

        self.setWindowTitle("Информация об аппаратной части ПК")

        # window size
        # Размер окна (размер окна можно регулировать, доступен полноэкранный режим. Это всего лишь размер окна при запуске программы)
        self.setGeometry(0, 0, 510, 810)

        tab_widget = QTabWidget()

        #Создание талицы для информации о операционной системе
        tab_OS = QWidget()
        layout_OS = QVBoxLayout()
        table_OS = QTableWidget()
        table_OS.setRowCount(len(OS_properties))
        table_OS.setColumnCount(2)

        #Создание талицы для информации о процессорах
        tab_Proccessors = QWidget()
        layout_Processors = QVBoxLayout()
        table_Processors = QTableWidget()
        table_Processors.setRowCount(len(processor_properties))
        table_Processors.setColumnCount(len(processors) + 1)

        # Создание талицы для информации о оперативной памяти
        tab_Ram = QWidget()
        layout_Ram = QVBoxLayout()
        table_Ram = QTableWidget()
        table_Ram.setRowCount(len(Ram_properties))
        table_Ram.setColumnCount(len(Ram) + 1)

        # Создание талицы для информации о материнской плате
        tab_Board = QWidget()
        layout_Board = QVBoxLayout()
        table_Board = QTableWidget()
        table_Board.setRowCount(len(motherboard_properties))
        table_Board.setColumnCount(2)

        # OS
        # Заполнение талицы для информации о операционной системе информацией, полученной при помощи WMI
        for i in range(len(OS_properties)):
            # for j in range(2):
            table_OS.setItem(i, 0, QTableWidgetItem(str(OS_properties[i])))
            if str(OS[OS_properties[i]]) == '':
                table_OS.setItem(i, 1, QTableWidgetItem('Информация отсутствует'))
            else:
                table_OS.setItem(i, 1, QTableWidgetItem(str(OS[OS_properties[i]])))

        layout_OS.addWidget(table_OS)
        tab_OS.setLayout(layout_OS)

        #motherboard
        # Заполнение талицы для информации об операционной системе информацией, полученной при помощи WMI
        for property in range(len(motherboard_properties)):
            table_Board.setItem(property, 0, QTableWidgetItem(str(motherboard_properties[property])))
            if str(OS[OS_properties[property]]) == '':
                table_Board.setItem(property, 1, QTableWidgetItem('Информация отсутствует'))
            else:
                table_Board.setItem(property, 1, QTableWidgetItem(str(motherboard[motherboard_properties[property]])))

        layout_Board.addWidget(table_Board)
        tab_Board.setLayout(layout_Board)


        # Processors
        # Заполнение талицы для информации о процессорах информацией, полученной при помощи WMI
        for i in range(len(processor_properties)):
            table_Processors.setItem(i, 0, QTableWidgetItem(str(processor_properties[i])))
            for j in range(len(processors)):
                if processor_properties[i] not in processors[j]:
                    table_Processors.setItem(i, j + 1, QTableWidgetItem('Информация отсутствует'))
                elif str(processors[j][processor_properties[i]]) == '':
                    table_Processors.setItem(i, j + 1, QTableWidgetItem('Информация отсутствует'))
                else:
                    table_Processors.setItem(i, j + 1, QTableWidgetItem(str(processors[j][processor_properties[i]])))
        layout_Processors.addWidget(table_Processors)
        tab_Proccessors.setLayout(layout_Processors)

        # Ram
        # Заполнение талицы для информации об оперативной памяти информацией, полученной при помощи WMI
        for i in range(len(Ram_properties)):
            table_Ram.setItem(i, 0, QTableWidgetItem(str(Ram_properties[i])))
            for j in range(len(Ram)):
                if Ram_properties[i] not in Ram[j]:
                    table_Ram.setItem(i, j + 1, QTableWidgetItem('Информация отсутствует'))
                elif str(Ram[j][Ram_properties[i]]) == '':
                    table_Ram.setItem(i, j + 1, QTableWidgetItem('Информация отсутствует'))
                else:
                    table_Ram.setItem(i, j + 1, QTableWidgetItem(str(Ram[j][Ram_properties[i]])))
        layout_Ram.addWidget(table_Ram)
        tab_Ram.setLayout(layout_Ram)

        #добавление вкладки "Операционная система" с информацией об операционной системе
        tab_widget.addTab(tab_OS, "Операционная система")
        # добавление вкладки "Процессоры" с информацией о процессорах
        tab_widget.addTab(tab_Proccessors, "Процессоры")
        # добавление вкладки "Оперативная память" с информацией об оперативной памяти
        tab_widget.addTab(tab_Ram, "Оперативная память")
        # добавление вкладки "Материнская плата" с информацией о материнской плате
        tab_widget.addTab(tab_Board, "Материнская плата")

        self.setCentralWidget(tab_widget)


def information():
    global OS, OS_properties, processors, processor_properties, Ram_properties, Ram, motherboard_properties, motherboard
    c = wmi.WMI()
    processors = []

    # OS
    # Получение информации об операционной системе
    for os in c.Win32_OperatingSystem():
        OS = {prop: getattr(os, prop) for prop in os.properties}
        OS_properties = list(os.properties)

    # Processors
    # Получение информации о процессорах
    counter = False
    for processor in c.Win32_Processor():
        processors.append({prop: getattr(processor, prop) for prop in processor.properties})
        if counter is False:
            processor_properties = set(processor.properties)
            counter = True
        else:
            processor_properties = processor_properties | set(processor.properties)
    processor_properties = list(processor_properties)

    # memory
    #Получение информации об оперативной памяти
    Ram = []
    counter = False
    for memory in c.Win32_PhysicalMemory():
        Ram.append({prop: getattr(memory, prop) for prop in memory.properties})
        if counter is False:
            Ram_properties = set(memory.properties)
            counter = True
        else:
            Ram_properties = Ram_properties | set(memory.properties)
    Ram_properties = list(Ram_properties)

    #motherboard
    #Получение информации о материнской плате
    for motherboard_ in c.Win32_BaseBoard():
        motherboard={prop: getattr(motherboard_, prop) for prop in motherboard_.properties}
        motherboard_properties=list(motherboard_.properties)

    return None


app = QApplication(sys.argv)
# Получение информации при помощи WMI
information()
# Создание атрибута класса окна
window = MyWindow(OS, OS_properties, processors, processor_properties, Ram, Ram_properties, motherboard, motherboard_properties)
# Создание окна
window.show()
sys.exit(app.exec_())

