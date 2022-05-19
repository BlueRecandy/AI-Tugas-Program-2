import openpyxl

# Buat nyimpen data dari file masukan.xls
data_masukan = {}

# Buat nyimpen data yang nanti disimpan di luaran.xls
data_luaran = {}

# TODO Baca data dari file masukan.xls
def read_masukan():
    wb = openpyxl.load_workbook("masukan.xlsx")
    sheet = wb['Sheet1']

    for row in sheet.iter_rows(min_row=2, max_row=101, min_col=1, values_only=True):
        id = row[0]
        masukan = {
            "Nama tempat makan" : row[2],
            "Rating" : row[4],
            "Jumlah menu" : row[3],
            "Harga rata-rata" : row[6],
        }
        data_masukan[id] = masukan
    wb.close()


# TODO Tulis data hasil olahan observasi ke file luaran.xls
def write_luaran():
    luaran_file = openpyxl.Workbook()
    luaran_sheet = luaran_file.active

    luaran_sheet['A1'] = 'ID'
    luaran_sheet['B1'] = 'Nama Tempat Makan'
    luaran_sheet['C1'] = 'TOR'

    for index in range(1, len(data_luaran) + 1):
        resto_obj = data_luaran[index]
        row = (1 + index)
        luaran_sheet[f'A{row}'] = index
        luaran_sheet[f'B{row}'] = resto_obj['nama_toko']
        luaran_sheet[f'C{row}'] = resto_obj['TOR']

    luaran_file.save('luaran.xlsx')
    luaran_file.close()


# Attribute Grouping:
# Rating:
#   - Enak              : > 3.7
#   - Biasa             : 2 - 4
#   - Kurang Enak       : < 2.3
#
# Rata-rata Harga:
#   - Mahal             : > 50 ribu
#   - Dapat Diterima    : 15 - 55 ribu
#   - Murah       : < 20 ribu
#
# Jumlah Menu:
#   - Sangat Variatif   : > 33 menu
#   - Variatif          : > 5 - 35 menu
#   - Kurang Variatif   : < 7

# TODO fungsi fuzzifikasi disini
# contoh di slide halaman 54
def fuzzification():
    for resto_id in data_masukan.keys():
        resto_obj = data_masukan[resto_id]

        rating = resto_obj['rating']
        jumlah_menu = resto_obj['jumlah_menu']
        harga_rata_rata = resto_obj['harga_rata_rata']

        print('-'*10)
        # TODO Calculate attribute
        print('-' * 10)
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
    read_masukan()
    write_luaran()
