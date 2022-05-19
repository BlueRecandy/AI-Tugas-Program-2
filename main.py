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
            "nama_tempat_makan": row[2],
            "rating": row[4],
            "jumlah_menu": row[3],
            "harga_rata_rata": row[6],
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
        luaran_sheet[f'B{row}'] = resto_obj['nama_tempat_makan']
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

rating_groups = [5, 4, 3.7, 2.3, 2, 0]
menu_groups = [150, 35, 33, 7, 5, 1]
harga_groups = [100000, 55000, 50000, 20000, 15000, 0]


def get_grouping(value: float, groups: list[int]):
    for i in range(0, len(groups)):
        a = groups[i]
        b = groups[i + 1]
        lower = min(a, b)
        upper = max(a, b)

        if lower < value <= upper:
            return lower, upper


def calc_group(value: float, lower: float, upper: float):
    low = -(value - upper) / (upper - lower)
    high = (value - lower) / (upper - lower)
    return low, high

# TODO fungsi fuzzifikasi disini
# contoh di slide halaman 54
def fuzzification():
    fuzzy_result = {}

    for resto_id in data_masukan.keys():
        resto_obj = data_masukan[resto_id]

        rating = resto_obj['rating']
        harga_rata_rata = resto_obj['harga_rata_rata']

        rating_lower, rating_upper = get_grouping(rating, rating_groups)
        harga_lower, harga_upper = get_grouping(harga_rata_rata, harga_groups)

        rating_fuzzy = calc_group(rating, rating_lower, rating_upper)
        harga_fuzzy = calc_group(harga_rata_rata, harga_lower, harga_upper)

        rating_result = {'lower': rating_fuzzy[0], 'upper': rating_fuzzy[1]}
        harga_result = {'lower': harga_fuzzy[0], 'upper': harga_fuzzy[1]}

        result = {
            'rating': rating_result,
            'harga': harga_result
        }

        fuzzy_result[resto_id] = result

    return fuzzy_result


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
    print(fuzzification())
