import openpyxl

# Buat nyimpen data dari file masukan.xls
data_masukan = {}

# Buat nyimpen data yang nanti disimpan di luaran.xls
data_luaran = {}


# TODO Baca data dari file masukan.xls
path = "masukan.xlsx"
 
# workbook object is created
wb_obj = openpyxl.load_workbook(path)
 
sheet_obj = wb_obj.active
 
max_col = sheet_obj.max_column
 
# Will print a particular row value
for i in range(1, max_col + 1):
    cell_obj = sheet_obj.cell(row = 2, column = i)
    print(cell_obj.value, end = " ")

def read_masukan(self):
    return


# TODO Tulis data hasil olahan observasi ke file luaran.xls
def write_luaran():
    luaran_file = openpyxl.Workbook()
    luaran_sheet = luaran_file.active

    for i in range(1, len(data_luaran) + 1):
        resto_obj = data_luaran[i]
        row = (1 + i)
        luaran_sheet[f'A{row}'] = i
        luaran_sheet[f'B{row}'] = resto_obj['nama_toko']
        luaran_sheet[f'C{row}'] = resto_obj['TOR']

    luaran_file.close()
    luaran_file.save('luaran.xlsx')
    return


# Attribute Grouping:
# Rating:
#   - Enak              : > 3.7
#   - Biasa             : 2 - 4
#   - Kurang Enak       : < 2.3
#
# Rata-rata Harga:
#   - Mahal             : > 50 ribu
#   - Dapat Diterima    : 15 - 55 ribu
#   - Kurang Enak       : < 20 ribu
#
# Jumlah Menu:
#   - Sangat Variatif   : > 17 menu
#   - Variatif          : > 5 - 20 menu
#   - Kurang Variatif   : < 7

# TODO fungsi fuzzifikasi disini
# contoh di slide halaman 54
def fuzzification(self):
    return

# TODO fungsi inference disini
# contoh di slide halaman 55-57
def inference(self):
    return

# TODO fungsi defuzzifikasi disini
# contoh di slide halaman 58
def defuzzification(self):
    return


if __name__ == '__main__':
    write_luaran()
