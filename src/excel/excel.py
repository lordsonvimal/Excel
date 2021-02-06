import pandas

class Column:
    def __init__(self, col_data, col_name, col_index):
        self.data = col_data
        self.index = col_index
        self.name = col_name

    def set_width(self, worksheet, col_width=-1, offset=0):
        width = col_width if col_width > -1 else max(self.data.astype(str).map(len).max(), len(str(self.name))) + offset
        worksheet.set_column(self.index, self.index, width)

class ExcelSheet:
    def __init__(self, data_frame, sheet_name, writer):
        self.data_frame = data_frame
        self.name = sheet_name
        self.writer = writer

    def iter_cols(self, col_name=""):
        for index, col in enumerate(self.data_frame):
            if (col_name == "") or (col_name == col):
                column = Column(self.data_frame[col], col, index)
                yield column

    def to_excel(self):
        self.data_frame.to_excel(self.writer, index=False, sheet_name=self.name)
        worksheet = self.writer.sheets[self.name]
        # We can add any custom formatting here
        # 1. Write formatting rules
        # 2. Connect formatting functions here
        # 3. Move these formatting to separate functions
        for column in self.iter_cols():
            column.set_width(worksheet)

class Excel:
    def __init__(self, input_path, output_path):
        self.data = pandas.read_excel(input_path, sheet_name=None)
        self.writer = pandas.ExcelWriter(output_path, engine="xlsxwriter")
        self.sheets = []

    def update_sheets(self):
        for key, value in self.data.items():
            self.sheets.append(ExcelSheet(pandas.DataFrame(value), key, self.writer))

    def save(self):
        if len(self.sheets) > 0:
            for sheet in self.sheets:
                sheet.to_excel()
            self.writer.save()

# Sample to test the above fns
# This should go to a separate file
# Adding it here for now
# excel = Excel("test.xlsx", "new.xlsx")
# excel.update_sheets()
# excel.save()
