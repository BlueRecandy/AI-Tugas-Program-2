import openpyxl

# Buat nyimpen data dari file masukan.xls
data_masukan = {}

# Buat nyimpen data yang nanti disimpan di luaran.xls
data_luaran = {}


# TODO Baca data dari file masukan.xls
def read_masukan(self):
    return


# TODO Tulis data hasil olahan observasi ke file luaran.xls
def write_luaran(self):
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
    print('hello')
