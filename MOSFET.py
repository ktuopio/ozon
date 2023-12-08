import pandas as pd

# Загрузка данных из исходных файлов
mosfet_df = pd.read_excel(r'C:\Users\Betekhtina\Downloads\mosfet_new_.xlsx', sheet_name='mosfet')
wayon_df = pd.read_excel(r'C:\Users\Betekhtina\Downloads\mosfet_new_.xlsx', sheet_name='wayon')

# Создание нового DataFrame для результата
result_df = pd.DataFrame(columns=[
    "Код номенклатуры", "Наименование", "Производитель", "Package Type", "VDS", "ID 25", "ID 100", "RDSON, 4.5V typ",
    "RDSON, 4.5V max", "ID1", "RDSON, -4.5V typ", "RDSON, -4.5V max", "ID2",
    "RDSON, 10V typ", "RDSON, 10V max", "ID3", "RDSON, -10V typ", "RDSON, -10V max", "ID4",
    "RDSON, 2.5V typ", "RDSON, 2.5V max", "ID5", "RDSON, -2.5V typ", "RDSON, -2.5V max", "ID6",
    "RDSON, 1.8V typ", "RDSON, 1.8V max", "ID7", "RDSON, -1.8V typ", "RDSON, -1.8V max", "ID8",
    "RDSON, 6V typ", "RDSON, 6V max", "ID9", "RDSON, -6V typ", "RDSON, -6V max", "ID10", "QG typ", "CISS", "PD",
    "Код номенклатуры_замены", "Наименование_замены", "Производитель_замены", "Package Type_замены", "VDS_замены",
    "ID 25_замены", "ID 100_замены", "RDSON, 4.5V typ_замены", "RDSON, 4.5V max_замены", "ID1_замены",
    "RDSON, -4.5V typ_замены", "RDSON, -4.5V max_замены", "ID2_замены",
    "RDSON, 10V typ_замены", "RDSON, 10V max_замены", "ID3_замены", "RDSON, -10V typ_замены",
    "RDSON, -10V max_замены", "ID4_замены",
    "RDSON, 2.5V typ_замены", "RDSON, 2.5V max_замены", "ID5_замены", "RDSON, -2.5V typ_замены",
    "RDSON, -2.5V max_замены", "ID6_замены",
    "RDSON, 1.8V typ_замены", "RDSON, 1.8V max_замены", "ID7_замены", "RDSON, -1.8V typ_замены",
    "RDSON, -1.8V max_замены", "ID8_замены",
    "RDSON, 6V typ_замены", "RDSON, 6V max_замены", "ID9_замены", "RDSON, -6V typ_замены", "RDSON, -6V max_замены",
    "ID10_замены", "QG typ_замены", "CISS_замены", "PD_замены",
])

# Столбцы RDSON, которые вы хотите учитывать
rdson_columns = [
    "RDSON, 4.5V typ", "RDSON, 4.5V max", "RDSON, -4.5V typ", "RDSON, -4.5V max",
    "RDSON, 10V typ", "RDSON, 10V max", "RDSON, -10V typ", "RDSON, -10V max",
    "RDSON, 2.5V typ", "RDSON, 2.5V max", "RDSON, -2.5V typ", "RDSON, -2.5V max",
    "RDSON, 1.8V typ", "RDSON, 1.8V max", "RDSON, -1.8V typ", "RDSON, -1.8V max",
    "RDSON, 6V typ", "RDSON, 6V max", "RDSON, -6V typ", "RDSON, -6V max"
]
columns_to_fillna = [
    "Package Type", "VDS", "ID 25", "ID 100", "QG typ", "CISS", "PD"
]
mosfet_df[rdson_columns] = mosfet_df[rdson_columns].fillna(0)
wayon_df[rdson_columns] = wayon_df[rdson_columns].fillna(0)

# Проходим по каждой строке в таблице mosfet
for index, row in wayon_df.iterrows():
    package_type = row["Package Type"]
    vds = row["VDS"]
    id_25 = row["ID 25"]
    id_100 = row["ID 100"]
    qg_typ = row["QG typ"]
    ciss = row["CISS"]
    rdsontyp45 = row["RDSON, 4.5V typ"]
    rdsonmax45 = row["RDSON, 4.5V max"]
    rdsontyp045 = row["RDSON, -4.5V typ"]
    rdsonmax045 = row["RDSON, -4.5V max"]
    rdsontyp10 = row["RDSON, 10V typ"]
    rdsonmax10 = row["RDSON, 10V max"]
    rdsontyp010 = row["RDSON, -10V typ"]
    rdsonmax010 = row["RDSON, -10V max"]
    rdsontyp25 = row["RDSON, 2.5V typ"]
    rdsonmax25 = row["RDSON, 2.5V max"]
    rdsontyp025 = row["RDSON, -2.5V typ"]
    rdsonmax025 = row["RDSON, -2.5V max"]
    rdsontyp18 = row["RDSON, 1.8V typ"]
    rdsonmax18 = row["RDSON, 1.8V max"]
    rdsontyp018 = row["RDSON, -1.8V typ"]
    rdsonmax018 = row["RDSON, -1.8V max"]
    rdsontyp6 = row["RDSON, 6V typ"]
    rdsonmax6 = row["RDSON, 6V max"]
    rdsontyp06 = row["RDSON, -6V typ"]
    rdsonmax06 = row["RDSON, -6V max"]
    powerd = row["PD"]


# Значения RDSON из таблицы mosfet
    rdson_values = [row[rdson_col] for rdson_col in rdson_columns]

# Флаг для отслеживания выполнения всех условий
    conditions_met = True

# Фильтруем таблицу wayon по условиям
    filtered_mosfet = mosfet_df[
        (mosfet_df["Package Type"] == package_type) &
        (mosfet_df["VDS"] * 0.85 <= vds) & (vds <= mosfet_df["VDS"] * 1.5) &
        (mosfet_df["ID 25"] * 0.85 < id_25) & (id_25 <= mosfet_df["ID 25"] * 1.5) &
        (mosfet_df["RDSON, 4.5V typ"] * 0.5 <= rdsontyp45) & (rdsontyp45 <= mosfet_df["RDSON, 4.5V typ"] * 1.15) &
        (mosfet_df["RDSON, 4.5V max"] * 0.5 <= rdsonmax45) & (rdsonmax45 <= mosfet_df["RDSON, 4.5V max"] * 1.15) &
        (mosfet_df["RDSON, -4.5V typ"] * 0.5 <= rdsontyp045) & (rdsontyp045 <= mosfet_df["RDSON, -4.5V typ"] * 1.15) &
        (mosfet_df["RDSON, -4.5V max"] * 0.5 <= rdsonmax045) & (rdsonmax045 <= mosfet_df["RDSON, -4.5V max"] * 1.15) &
        (mosfet_df["RDSON, 10V typ"] * 0.5 <= rdsontyp10) & (rdsontyp10 <= mosfet_df["RDSON, 10V typ"] * 1.15) &
        (mosfet_df["RDSON, 10V max"] * 0.5 <= rdsonmax10) & (rdsonmax10 <= mosfet_df["RDSON, 10V max"] * 1.15) &
        (mosfet_df["RDSON, -10V typ"] * 0.5 <= rdsontyp010) & (rdsontyp010 <= mosfet_df["RDSON, -10V typ"] * 1.15) &
        (mosfet_df["RDSON, -10V max"] * 0.5 <= rdsonmax010) & (rdsonmax010 <= mosfet_df["RDSON, -10V max"] * 1.15) &
        (mosfet_df["RDSON, 2.5V typ"] * 0.5 <= rdsontyp25) & (rdsontyp25 <= mosfet_df["RDSON, 2.5V typ"] * 1.15) &
        (mosfet_df["RDSON, 2.5V max"] * 0.5 <= rdsonmax25) & (rdsonmax25 <= mosfet_df["RDSON, 2.5V max"] * 1.15) &
        (mosfet_df["RDSON, -2.5V typ"] * 0.5 <= rdsontyp025) & (rdsontyp025 <= mosfet_df["RDSON, -2.5V typ"] * 1.15) &
        (mosfet_df["RDSON, -2.5V max"] * 0.5 <= rdsonmax025) & (rdsonmax025 <= mosfet_df["RDSON, -2.5V max"] * 1.15) &
        (mosfet_df["RDSON, 1.8V typ"] * 0.5 <= rdsontyp18) & (rdsontyp18 <= mosfet_df["RDSON, 1.8V typ"] * 1.15) &
        (mosfet_df["RDSON, 1.8V max"] * 0.5 <= rdsonmax18) & (rdsonmax18 <= mosfet_df["RDSON, 1.8V max"] * 1.15) &
        (mosfet_df["RDSON, -1.8V typ"] * 0.5 <= rdsontyp018) & (rdsontyp018 <= mosfet_df["RDSON, -1.8V typ"] * 1.15) &
        (mosfet_df["RDSON, -1.8V max"] * 0.5 <= rdsonmax018) & (rdsonmax018 <= mosfet_df["RDSON, -1.8V max"] * 1.15) &
        (mosfet_df["RDSON, 6V typ"] * 0.5 <= rdsontyp6) & (rdsontyp6 <= mosfet_df["RDSON, 6V typ"] * 1.15) &
        (mosfet_df["RDSON, 6V max"] * 0.5 <= rdsonmax6) & (rdsonmax6 <= mosfet_df["RDSON, 6V max"] * 1.15) &
        (mosfet_df["RDSON, -6V typ"] * 0.5 <= rdsontyp06) & (rdsontyp06 <= mosfet_df["RDSON, -6V typ"] * 1.15) &
        (mosfet_df["RDSON, -6V max"] * 0.5 <= rdsonmax06) & (rdsonmax06 <= mosfet_df["RDSON, -6V max"] * 1.15) &
        (mosfet_df["QG typ"] * 0.7 <= qg_typ) & (qg_typ <= mosfet_df["QG typ"] * 1.3) &
        (mosfet_df["CISS"] * 0.7 <= ciss) & (ciss <= mosfet_df["CISS"] * 1.3) &
        (mosfet_df["PD"] * 0.7 <= powerd) & (powerd <= mosfet_df["PD"] * 1.3)
        ]

    # Если есть соответствующие записи в таблице wayon, добавляем их в результат
    if conditions_met and not filtered_mosfet.empty:
        for _, mosfet_row in filtered_mosfet.iterrows():
            new_row = {
                "Код номенклатуры": row["Код номенклатуры"],
                "Наименование": row["Наименование"],
                "Производитель": row["Производитель"],
                "Package Type": row["Package Type"],
                "VDS": row["VDS"],
                "ID 25": row["ID 25"],
                "ID 100": row["ID 100"],
                "RDSON, 4.5V typ": row["RDSON, 4.5V typ"],
                "RDSON, 4.5V max": row["RDSON, 4.5V max"],
                "ID1": row["ID1"],
                "RDSON, -4.5V typ": row["RDSON, -4.5V typ"],
                "RDSON, -4.5V max": row["RDSON, -4.5V max"],
                "ID2": row["ID2"],
                "RDSON, 10V typ": row["RDSON, 10V typ"],
                "RDSON, 10V max": row["RDSON, 10V max"],
                "ID3": row["ID3"],
                "RDSON, -10V typ": row["RDSON, -10V typ"],
                "RDSON, -10V max": row["RDSON, -10V max"],
                "ID4": row["ID4"],
                "RDSON, 2.5V typ": row["RDSON, 2.5V typ"],
                "RDSON, 2.5V max": row["RDSON, 2.5V max"],
                "ID5": row["ID5"],
                "RDSON, -2.5V typ": row["RDSON, -2.5V typ"],
                "RDSON, -2.5V max": row["RDSON, -2.5V max"],
                "ID6": row["ID6"],
                "RDSON, 1.8V typ": row["RDSON, 1.8V typ"],
                "RDSON, 1.8V max": row["RDSON, 1.8V max"],
                "ID7": row["ID7"],
                "RDSON, -1.8V typ": row["RDSON, -1.8V typ"],
                "RDSON, -1.8V max": row["RDSON, -1.8V max"],
                "ID8": row["ID8"],
                "RDSON, 6V typ": row["RDSON, 6V typ"],
                "RDSON, 6V max": row["RDSON, 6V max"],
                "ID9": row["ID9"],
                "RDSON, -6V typ": row["RDSON, -6V typ"],
                "RDSON, -6V max": row["RDSON, -6V max"],
                "ID10": row["ID10"],
                "QG typ": row["QG typ"],
                "CISS": row["CISS"],
                "PD": row["PD"],
                "Код номенклатуры_замены": mosfet_row["Код номенклатуры"],
                "Наименование_замены": mosfet_row["Наименование"],
                "Производитель_замены": mosfet_row["Производитель"],
                "Package Type_замены": mosfet_row["Package Type"],
                "VDS_замены": mosfet_row["VDS"],
                "ID 25_замены": mosfet_row["ID 25"],
                "ID 100_замены": mosfet_row["ID 100"],
                "RDSON, 4.5V typ_замены": mosfet_row["RDSON, 4.5V typ"],
                "RDSON, 4.5V max_замены": mosfet_row["RDSON, 4.5V max"],
                "ID1_замены": mosfet_row["ID1"],
                "RDSON, -4.5V typ_замены": mosfet_row["RDSON, -4.5V typ"],
                "RDSON, -4.5V max_замены": mosfet_row["RDSON, -4.5V max"],
                "ID2_замены": mosfet_row["ID2"],
                "RDSON, 10V typ_замены": mosfet_row["RDSON, 10V typ"],
                "RDSON, 10V max_замены": mosfet_row["RDSON, 10V max"],
                "ID3_замены": mosfet_row["ID3"],
                "RDSON, -10V typ_замены": mosfet_row["RDSON, -10V typ"],
                "RDSON, -10V max_замены": mosfet_row["RDSON, -10V max"],
                "ID4_замены": mosfet_row["ID4"],
                "RDSON, 2.5V typ_замены": mosfet_row["RDSON, 2.5V typ"],
                "RDSON, 2.5V max_замены": mosfet_row["RDSON, 2.5V max"],
                "ID5_замены": mosfet_row["ID5"],
                "RDSON, -2.5V typ_замены": mosfet_row["RDSON, -2.5V typ"],
                "RDSON, -2.5V max_замены": mosfet_row["RDSON, -2.5V max"],
                "ID6_замены": mosfet_row["ID6"],
                "RDSON, 1.8V typ_замены": mosfet_row["RDSON, 1.8V typ"],
                "RDSON, 1.8V max_замены": mosfet_row["RDSON, 1.8V max"],
                "ID7_замены": mosfet_row["ID7"],
                "RDSON, -1.8V typ_замены": mosfet_row["RDSON, -1.8V typ"],
                "RDSON, -1.8V max_замены": mosfet_row["RDSON, -1.8V max"],
                "ID8_замены": mosfet_row["ID8"],
                "RDSON, 6V typ_замены": mosfet_row["RDSON, 6V typ"],
                "RDSON, 6V max_замены": mosfet_row["RDSON, 6V max"],
                "ID9_замены": mosfet_row["ID9"],
                "RDSON, -6V typ_замены": mosfet_row["RDSON, -6V typ"],
                "RDSON, -6V max_замены": mosfet_row["RDSON, -6V max"],
                "ID10_замены": mosfet_row["ID10"],
                "QG typ_замены": mosfet_row["QG typ"],
                "CISS_замены": mosfet_row["CISS"],
                "PD_замены": mosfet_row["PD"]
            }

            result_df = pd.concat([result_df, pd.DataFrame([new_row])], ignore_index=True)

# Сохраняем результат в новый файл crossref.xlsx
result_df.to_excel(r"C:\Users\Betekhtina\Downloads\mosfet_crossref_PP.xlsx", index=False)

# Выводим сообщение об успешном выполнении
print("Скрипт успешно завершен. Результат сохранен в файл crossref.xlsx.")
