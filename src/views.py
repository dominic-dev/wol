import os
import pyforms

import harvestlist
import hlpickle

from   pyforms          import BaseWidget
from   pyforms.Controls import ControlText
from   pyforms.Controls import ControlButton
from   pyforms.Controls import ControlLabel
from   pyforms.Controls import ControlList
from   pyforms.Controls import ControlEmptyWidget
from   pyforms.Controls import ControlDir
from   pyforms.Controls import ControlFile
from   pyforms.Controls import ControlSaveFile
from PyQt5 import QtCore

class Main:
    # Main window
    root = None
    # Stores the data of the harvest list
    harvest_list = None

    # The HarvestList class
    hl = harvestlist.HarvestList()

    # Stores the object to be pickled
    hl_pickle = hlpickle.HLPickle()

    output_dir = None
    output_path = None

class mainWindow(BaseWidget):
    def __init__(self):

        super().__init__('Oogstlijst Manager')
        self._panel = ControlEmptyWidget()
        self._panel.value = Home()

        Main.root = self

class Home(BaseWidget):
    def __init__(self):
        super().__init__('Home')

        self.button_new = ControlButton('Nieuwe oogstlijst')
        self.button_open = ControlButton('Open bestaande oogstlijst')
        self.formset = ['button_new', 'button_open']

        self.button_new.value = self._new
        self.button_open.value = self._open

    def _new(self):
        window = NewHarvest()
        Main.root._panel.value = window

    def _open(self):
        window = OpenHarvest()
        Main.root._panel.value = window


class OpenHarvest(BaseWidget):
    def __init__(self):
        super().__init__('Open oogstlijst')
        self._file = ControlFile('Project')
        self._file.filter = "OLM files (*.olm)"
        self._button_continue = ControlButton('Kies')

        self._button_continue.value = self.open_file

    def open_file(self):
        Main.hl_pickle = Main.hl.load_pickle(self._file.value)
        Main.hl.harvest_list = Main.hl_pickle.harvest_list
        Main.hl.name = Main.hl_pickle.name

        window = EditHarvest()
        Main.root._panel.value = window

class NewHarvest(BaseWidget):
    def __init__(self):
        super().__init__('Nieuwe oogstlijst')

        self._nameField = ControlText('Naam')
        self._button_continue= ControlButton('Begin')
        self._button_continue.value = self._continue

        self.formset = ['_nameField', '_button_continue']

    def _continue(self):
        name = self._nameField.value
        if not name:
            return

        Main.hl_pickle = hlpickle.HLPickle(name)
        Main.hl.name = name

        window = EditHarvest()
        Main.root._panel.value = window

class EditHarvest(BaseWidget):
    def __init__(self):
        super().__init__('Oogstlijst Manager')
        reference_list = Main.hl.reference_data

        self._list_left = ControlList('Stamlijst')
        self._list_left.readonly = True

        self._list_right = ControlList('Oogstlijst')
        self._list_right.select_entire_row =True
        self._list_right.horizontal_headers = [
            'Oogst',
            'Hoeveelheid (kg)',
        ]

        self._save_button = ControlButton('Verder')
        self._save_button.value = self._continue

        self.formset = ['_list_left', '_list_right', '_save_button']


        def onclick_left(row, column):
            r = self._list_left.get_currentrow_value()
            self._list_right.__add__(r + [0])

            cell = self._list_right.get_cell(0, self._list_right.rows_count - 1)
            cell.setFlags(QtCore.Qt.ItemIsEditable)
            #cell = self._list_right.get_cell(1, self._list_right.rows_count - 1)
            #cell.setTextAlignment(QtCore.Qt.AlignCenter)

        self._list_left.cell_double_clicked_event = onclick_left


        for r in reference_list:
            self._list_left.__add__([' '.join(r[1:3])])

        if Main.hl.harvest_list:
            for r in Main.hl.harvest_list:
                self._list_right.__add__([r[0], r[1]])

    def _continue(self):
        user_input = self._list_right.value
        if not user_input:
            return

        harvest_list = []

        # Combine user input with reference data
        for row in user_input:
            reference = [ ref for ref in Main.hl.reference_data if\
                 ref[1] and ref[1] in row[0] and ref[2] and ref[2] in row[0] ][0]
            if not reference:
                raise Exception('Reference data not found')

            row += [reference[0]] # add prod nr
            row += [reference[2]] # add part
            row += [reference[3]] # add date
            harvest_list.append(row)



        Main.hl.harvest_list = Main.hl_pickle.harvest_list = harvest_list

        window = Rapport()
        Main.root._panel.value = window


class Rapport(BaseWidget):
    def __init__(self):
        super().__init__('Oogstlijst Manager')
        self._plan_button = ControlButton('Genereer planning')
        self._harvest_list_button = ControlButton('Genereer oogstlijst')
        self._save = ControlButton('Opslaan')

        self._plan_button.value = self.generate_plan
        self._harvest_list_button.value = self.generate_list
        self._save.value = self.save_pickle

        self.formset = ['_plan_button', '_harvest_list_button', '_save']

    def generate_plan(self):
        window = SelectFile()
        window.parent = self

        def callback():
            path = os.path.join(Main.output_dir ,Main.hl_pickle.name +\
                              '_planning.xlsx')
            Main.hl.save_plan(path)
            success = Message('Planning gemaakt in {}'.format(path))
            success.parent = self
            success.show()

        window.callback = callback
        window.show()

    def generate_list(self):
        window = SelectFile()
        window.parent = self

        def callback():
            path = os.path.join(Main.output_dir ,Main.hl_pickle.name +\
                              '_oogstlijst.xlsx')
            Main.hl.save_list(path)
            success = Message('Oogstlijst gemaakt in {}'.format(path))
            success.parent = self
            success.show()

        window.callback = callback
        window.show()

    def save_pickle(self):
        #window = SelectFile()
        window = SaveFile()
        window.parent = self

        def callback():
            path = os.path.abspath(Main.output_path)
            #path = os.path.join(Main.output_dir ,Main.hl_pickle.name +\
            #                  '.olm')
            Main.hl.save_pickle(Main.hl_pickle, path)
            success = Message('Project opgeslagen in {}'.format(path))
            success.parent = self
            success.show()

        window.callback = callback
        window.show()


class SelectFile(BaseWidget):
    def __init__(self):
        super().__init__('Oogstlijst Manager')
        self._label = ControlLabel('Kies de map waarin je het bestand wil opslaan')
        self._dir = ControlDir('Map')
        self._select_button = ControlButton('Kies')
        self._select_button.value = self.select

    def select(self):
        directory = self._dir.value
        if not directory:
            return
        Main.output_dir = directory
        self.close()

        if self.callback:
            self.callback()

class SaveFile(BaseWidget):
    def __init__(self):
        super().__init__('Oogstlijst Manager')
        self._file = ControlSaveFile('Opslaan')
        self._file.filter = 'Oogstlijst Manager bestanden (*.olm)'
        self._select_button = ControlButton('Verder')
        self._select_button.value = self.select

        self.formset = ['_file', '_select_button']

    def select(self):
        path = self._file.value
        if not path:
            return

        # extension
        filename, extension = os.path.splitext(path)
        if extension.lower() != '.olm':
            path += '.olm'

        Main.output_path = path
        self.close()

        if self.callback:
            self.callback()

class Message(BaseWidget):
    def __init__(self, message):
        super().__init__('Oogstlijst Manager')
        self._label = ControlLabel(message)


def main():
    pyforms.start_app(mainWindow)
    #pyforms.start_app(EditHarvest)


if __name__ == "__main__":
    main()
