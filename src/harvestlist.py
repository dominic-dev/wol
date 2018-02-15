import csv
import os
import pickle
import xlrd
import xlsxwriter

BASE_PATH = os.path.join('..', 'data', 'stamlijst.xlsx')
OUTPUT_DIR = 'output'
OUTPUT_FILE = 'oogstlijst.xlsx'
OUTPUT_PATH = os.path.join('..', OUTPUT_DIR, OUTPUT_FILE)


class HarvestList:
    def __init__(self, harvest_list=None, output_file=OUTPUT_PATH):
        self.harvest_list = harvest_list
        self.output_file = output_file
        self.name = None

        # collect reference data
        workbook = xlrd.open_workbook(os.path.normpath(BASE_PATH))
        sheet = workbook.sheets()[0]
        self.reference_data = [sheet.row_values(i) for i in
                               range(sheet.nrows)][1:]

        # Prepare output directory
        output_dir_path = os.path.join('..', OUTPUT_DIR)
        if not os.path.exists(output_dir_path):
            os.makedirs(output_dir_path)



        # bar rows, keep track of which rows have bars for months
        self.bar_rows = []
        for i in range(16):
            self.bar_rows.append([])

    def save_plan(self, path):
        # output file
        self.workbook = xlsxwriter.Workbook(path)
        self._set_formats()
        self.worksheet = self.workbook.add_worksheet()
        self.worksheet.set_landscape()
        self.worksheet.set_column(0, 11, 7.2)

        self._create_plan_template()
        self._add_harvests()
        self.workbook.close()

    def save_list(self, path):
        self.workbook = xlsxwriter.Workbook(path)
        self._set_formats()
        self.worksheet = self.workbook.add_worksheet()
        self.worksheet.write(0, 0, self.name)
        self.worksheet.set_column('A:A', 20)
        self.worksheet.set_column('B:B', 30)
        self.worksheet.write(1, 0, 'Prod. nr.')
        self.worksheet.write(1, 1, 'Oogst')
        self.worksheet.write(1, 2, 'Gewicht')
        self.worksheet.write(1, 3, 'Datum')

        # Add mean date row to harvest_list
        harvest_list = [r + [Months.mean_as_int(r[-1])] for r in self.harvest_list]
        harvest_list = sorted(harvest_list, key=lambda r: r[-1])

        for i, row in enumerate(harvest_list):
            harvest, weight, prod_nr, part, date_range, date_index = row

            self.worksheet.write(i+2, 0, prod_nr)
            self.worksheet.write(i+2, 1, harvest, self.formats[part])
            self.worksheet.write(i+2, 2, weight)
            self.worksheet.write(i+2, 3, date_range)
        self.workbook.close()

    def save_pickle(self, object_, path):
        pickle.dump(object_, open(path, 'wb'))
        return True

    def load_pickle(self, path):
        data = pickle.load(open(path, 'rb'))
        return data

    def _create_plan_template(self):
        #worksheet.set_column()
        self.worksheet.write(3, 6, self.name)

        for  i, m in enumerate(Months.short_months):
            self.worksheet.write(32, i, m, self.formats['center'])

        # border
        border_format = self.workbook.add_format()
        border_format.set_right(1)
        border_format.set_left(1)
        row = 7
        column = 0
        while row < 32:
            column = 0
            while column < 12:
                self.worksheet.write(row, column, ' ', self.formats['center'])

                column += 1
            row += 1

    def _set_formats(self):
        # formats
        self.formats = {
            'bulbus' : self.workbook.add_format(
                {
                    'bg_color' : 'brown',
                    'font_size' : '10',
                    'font_color' : 'white',
                }
            ),
            'center' : self.workbook.add_format(
                {
                    'align' : 'center',
                    'left' : 1,
                    'right' : 1,
                }
            ),
            'fructuarium' : self.workbook.add_format(
                {
                    'bg_color' : 'red',
                    'font_size' : '10',
                }
            ),
            'fructus' : self.workbook.add_format(
                {
                    'bg_color' : 'red',
                    'font_size' : '10',
                }
            ),
            'folium' : self.workbook.add_format(
                {
                    'bg_color' : 'green',
                    'font_size' : '10',
                    'font_color' : 'white',
                }
            ),
            'folium recens' : self.workbook.add_format(
                {
                    'bg_color' : 'green',
                    'font_size' : '10',
                    'font_color' : 'white',
                }
            ),
            'flos' : self.workbook.add_format(
                {
                    'bg_color' : 'yellow',
                    'font_size' : '10',
                }
            ),
            'herba' : self.workbook.add_format(
                {
                    'bg_color' : '#808000',
                    'font_size' : '10',
                    'font_color' : 'white',
                }
            ),
            'planta tota' : self.workbook.add_format(
                {
                    'bg_color' : '#008000',
                    'font_size' : '10',
                    'font_color' : 'white',
                }
            ),
            'rhizoma' : self.workbook.add_format(
                {
                    'bg_color' : '#993300',
                    'font_size' : '10',
                    'font_color' : 'white',
                }
            ),
            'radix' : self.workbook.add_format(
                {
                    'bg_color' : '#993300',
                    'font_size' : '10',
                    'font_color' : 'white',
                }
            ),
            'summitates' : self.workbook.add_format(
                {
                    'bg_color' : '#ccffcc',
                    'font_size' : '10',
                }
            ),
            'summitates et folium' : self.workbook.add_format(
                {
                    'bg_color' : '#ccffcc',
                    'font_size' : '10',
                }
            ),
        }

    def _get_column_for_month(self, month):
        """Take a month (str) return a column (int)"""
        return Months.full_months.index(month)

    def _add_harvests(self):
        for row in self.harvest_list:
            name, weight, prod_nr, part, date = row
            # Parse date
            begin, end = Months.parse(date)

            insert_row = self._assign_to_row(begin, end)

            for column in range(begin, end+1):
                try:
                    self.worksheet.write(insert_row, column, '',\
                                         self.formats[part])
                except:
                    self.worksheet.write(insert_row, column, '')
            middle = (begin + end) // 2
            try:
                self.worksheet.write(insert_row, middle, name,\
                                    self.formats[part])
            except:
                self.worksheet.write(insert_row, middle, name)

    def _assign_to_row(self, begin, end):
        if not end:
            end = begin

        date_range = range(begin, end+1)
        for i in range(len(self.bar_rows)):
            # check if the range is free
            if any(x in self.bar_rows[i] for x in date_range):
                continue
            self.bar_rows[i] += range(begin, end+1)
            return 31-i

class Months:
    full_months = ['januari', 'februari', 'maart', 'april', 'mei', 'juni',
              'juli', 'augustus',
              'september', 'oktober', 'november', 'december']
    short_months = ['jan', 'feb', 'mrt', 'apr', 'mei', 'jun', 'jul', 'aug',
                   'sept', 'okt', 'nov', 'dec']

    @staticmethod
    def parse(date_range):
        """
        Take date_range(str) and return a tuple containing 
        the beginning and end of the range as integers
        (beign, end)
        """
        # remove words 'begin' and 'eind' from date
        # and split date by -
        date = [d.strip() for d in
                date_range.lower()\
                          .replace('begin', '')\
                          .replace('eind', '')\
                          .replace('–', '-')\
                          .split('-')]

        # Dates may be a range of month - month, or a single month
        begin = Months._month_to_number(date[0])
        if len(date) > 1:
            end = Months._month_to_number(date[-1])
        else:
            end = begin
        return (begin, end)

    @staticmethod
    def mean_as_int(date_range):
        """
        Take a date range (str) and return the index of the mean month
        """
        begin, end = Months.parse(date_range)
        return (begin + end) // 2

    @staticmethod
    def _month_to_number(month):
        """Take a month (str) return an int"""
        return Months.full_months.index(month)

if __name__ == '__main__':

    with open(os.path.join('..', 'oogstlijst 2017.csv')) as f:
        reader = csv.reader(f)
        harvest_list = [r for r in reader]
    h = HarvestList(harvest_list)
    h.save_list()
    #h._create_template()
    #h._add_harvests()
    #print(h.bar_rows)
    #h.workbook.close()

