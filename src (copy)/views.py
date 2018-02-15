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

    def _new(self):
        window = NewHarvest()
        Main.root._panel.value = window

    def _open(self):
        pass



class NewHarvest(BaseWidget):
    def __init__(self):
        super().__init__('Nieuwe oogstlijst')

        self._nameField = ControlText('Naam')
        self._button_continue= ControlButton('Begin')
        self._button_continue.value = self._continue

        self.formset = ['_nameField', '_button_continue']

    def _continue(self):
        Main.hl_pickle = hlpickle.HLPickle(self._nameField.value)

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

    def _continue(self):
        user_input = self._list_right.value

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

    def generate_plan(self):
        window = SelectFile()
        window.parent = self

        def callback():
            Main.hl.save_plan(Main.output_dir ,Main.hl_pickle.name + '_planning.xlsx')
            success = Message('Planning gemaakt in {}'.format(Main.output_dir))
            success.parent = self
            success.show()

        window.callback = callback
        window.show()

    def generate_list(self):
        window = SelectFile()
        window.parent = self

        def callback():
            Main.hl.save_list(Main.output_dir ,Main.hl_pickle.name +\
                              '_oogstlijst.xlsx')
            success = Message('Planning gemaakt in {}'.format(Main.output_dir))
            success.parent = self
            success.show()

        window.callback = callback
        window.show()




class SelectFile(BaseWidget):
    def __init__(self):
        super().__init__('Oogstlijst Manager')
        self._label = ControlLabel('Kies de map waarin je het bestand wil opslaan')
        self._dir = ControlDir('Bestand')
        self._select_button = ControlButton('Kies')
        self._select_button.value = self.select

    def select(self):
        Main.output_dir = self._dir.value
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
