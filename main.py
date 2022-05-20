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

rating_sets = {
    'Enak': {
        'upper': 5,
        'lower': 3.7
    },
    'Biasa': {
        'upper': 4,
        'lower': 2
    },
    'Kurang': {
        'upper': 2.3,
        'lower': 0
    }
}

harga_sets = {
    'Mahal': {
        'upper': 100000,
        'lower': 50000
    },
    'Diterima': {
        'upper': 55000,
        'lower': 17000
    },
    'Murah': {
        'upper': 20000,
        'lower': 0
    },

}


def get_category_sets(value: float, sets: dict):
    sets_area = []
    for set_type in sets.keys():
        type_obj = sets[set_type]
        upper = type_obj['upper']
        lower = type_obj['lower']

        if lower <= value <= upper:
            sets_area.append(set_type)
    return sets_area


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


def calc_group_v2(value: float, categories: list, lower: float, upper: float):
    if len(categories) == 1:
        return {categories[0]: 1}
    else:
        low = -(value - upper) / (upper - lower)
        high = (value - lower) / (upper - lower)
        return {categories[0]: high, categories[1]: low}


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


def attribute_range_filter(value: float, categories: list[str], category_sets: dict):
    filter = {}

    if len(categories) == 1:
        category = categories[0]
        filter['lower'] = category_sets[category]['lower']
        filter['upper'] = category_sets[category]['upper']
    elif len(categories) == 2:
        category_1 = categories[0]
        category_2 = categories[1]
        filter['lower'] = category_sets[category_1]['lower']
        filter['upper'] = category_sets[category_2]['upper']

    return filter


def fuzzification_v2():
    fuzzy_result = {}

    for resto_id in data_masukan.keys():
        resto_obj = data_masukan[resto_id]
        rating = resto_obj['rating']
        harga = resto_obj['harga_rata_rata']

        rating_categories = get_category_sets(rating, rating_sets)
        rating_filter = attribute_range_filter(rating, rating_categories, rating_sets)
        rating_result_values = calc_group_v2(rating, rating_categories, rating_filter['lower'], rating_filter['upper'])
        rating_result = {
            'value': rating,
            'range': {
                'category': rating_categories,
                'range': {
                    'lower': rating_filter['lower'],
                    'upper': rating_filter['upper']
                }
            },
            'result': rating_result_values
        }

        harga_categories = get_category_sets(harga, harga_sets)
        harga_filter = attribute_range_filter(harga, harga_categories, harga_sets)
        harga_result_values = calc_group_v2(harga, harga_categories, harga_filter['lower'], harga_filter['upper'])
        harga_result = {
            'value': harga,
            'range': {
                'category': harga_categories,
                'range': {
                    'lower': harga_filter['lower'],
                    'upper': harga_filter['upper']
                }
            },
            'result': harga_result_values
        }

        resto_fuzzy = {
            'rating': rating_result,
            'harga': harga_result
        }


        fuzzy_result[resto_id] = resto_fuzzy
        return fuzzy_result

    return fuzzy_result


# TODO fungsi inference disini
# contoh di slide halaman 55-57
def inference(fuzzy_result):

    # Loop tiap resto
    # Ambil attribut rating dan harga dari tiap resto
    # Ambil semua kategori yang ada dari rating dan harga
    # Pasangkan kategori rating dan harga
    # Lakukan rule dari NK

    return


# TODO fungsi defuzzifikasi disini
# contoh di slide halaman 58
def defuzzification(self):
    return


if __name__ == '__main__':
    read_masukan()
    print(fuzzification_v2())
