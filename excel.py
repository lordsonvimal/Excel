import pandas

# pip install openpyxl
def read(path):
    data = pandas.read_excel(path, sheet_name=None)
    print(data)
    return data

# pip install xlsxwriter
def write(path, data_frame):
    writer = pandas.ExcelWriter(path, engine="xlsxwriter")
    data_frame.to_excel(writer, index=False)
    # We can add any custom formatting here
    # 1. Write formatting rules
    # 2. Connect formatting functions here
    writer.save()

def dict_to_df(dict):
    for key, value in dict.items():
        yield pandas.DataFrame(value).T

# Sample to test the above fns
data = read("test.xlsx")
for item in dict_to_df(data):
    write("new.xlsx", item)
