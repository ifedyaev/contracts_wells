import re

# colors Excel
YELLOW = "00FFD500"
GREEN = "0000FF4D"
#
gSHEETS_REGEX = re.compile(r"\d{1,2}.\d{4}")

# key переменные
g_key_number = "№ п.п."
g_key_contract = "Договор №"
g_key_NGDU = "НГДУ"
g_key_field = "Месторождение"
g_key_well_number = "Скв. №"
g_key_bush_number = "Куст №"
g_key_liquid = "Жидкость"
g_key_count_day = "Количество суток"
g_key_data_install = "Дата монтажа"
g_key_data_uninstall = "Дата демонтажа"
g_key_reason_stop = "Причина сотановки"
g_key_type_YECN = "Тип УЭЦН"
g_key_owner = "Собственник"
g_key_sum_tex_close = "Сумма техзакрытия"
g_key_TK = "ТК"
g_key_EE = "ЭЭ"
g_key_fund = "Фонд"
g_key_move_fund = "Движение Фонда"
g_key_refusal_CNO = "Отказы СНО"
g_key_refusal_MRP = "Отказы МРП"
g_key_sum_day_rent = "Стоимость суток проката"
g_key_type_stop = "Признак остановки"
# border data
g_arr_border = [
    (15.0, 200.0),
    (250.0, 600.0),
    (700.0, 700.0),
    (800.0, 800.0),
    (1000.0, 1000.0),
    (1250.0, 1250.0),
    (1500.0, 1500.0)
]
