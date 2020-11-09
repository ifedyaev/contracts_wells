# -*- coding utf-8 -*-
import os
import sys
import time

# data
from alexey_fedyaev.data.container import BorderStatistWells
from alexey_fedyaev.const.const import *
from alexey_fedyaev.const.error import *
from alexey_fedyaev.logger.logger import err_log
from alexey_fedyaev.read_xls.real_sheet import extract_xls, read_list_sheets
from alexey_fedyaev.write_xls.workbook_write import WorkBookContractWells

# set to environment path PYTHONPATH
sys.path.append(os.path.dirname(os.path.abspath(__file__)))


def main():
    path = r"./test/input_data.xlsm"
    # check path
    # проверка на существование пути и что по пути лежит файл
    if not os.path.exists(path) or not os.path.isfile(path):
        err_log(ERR_PATH_FILE.format(path))
    # get sheet
    arr_all_sheets = read_list_sheets(path_xls=path)
    # filtering sheets
    arr_sheets = [s for s in arr_all_sheets if gSHEETS_REGEX.match(s)]

    # check filtering sheets
    if len(arr_sheets) == 0:
        err_log(mess=ERR_SHEET)
    # check
    data_sheet_by_month = {}
    for name_sheet in arr_sheets:
        month, year = name_sheet.split(".")
        key_month = int(month)
        # check month
        if not (1 <= key_month <= 12):
            err_log(mess=ERR_SHEET_MONTH.format(key_month))
        #
        data_sheet_by_month[key_month] = extract_xls(path_xls=path,
                                                     name_sheet=name_sheet)

    arr_month_keys = list(data_sheet_by_month.keys())
    arr_month_keys.sort()
    #
    df = data_sheet_by_month[arr_month_keys[0]]
    #
    arr_contracts = df[g_key_contract].drop_duplicates().values
    work_book = WorkBookContractWells()
    # loop contracts
    for i, contract in enumerate(arr_contracts):
        print("*" * 80)
        print("contracts: {0}".format(contract))  # format - подставляет значения
        # create sheet
        sheet = work_book.create_sheet(title=contract, index=i + 1)
        # map by contracts
        data_contract_by_border = {}
        # loop month
        for key_month in arr_month_keys:
            df = data_sheet_by_month[key_month]
            # loop border
            for lb, ub in g_arr_border:
                df_contract = df[(df[g_key_contract] == contract) \
                                 & (df[g_key_type_YECN] >= lb) & (df[g_key_type_YECN] <= ub) \
                                 & (df[g_key_NGDU] == 1)]
                border_statistic_wells = BorderStatistWells()
                # wells
                df_life_well = df_contract[df_contract[g_key_number].notna()]
                border_statistic_wells.count_wells, _ = df_life_well.shape
                # count days
                border_statistic_wells.count_days = df_contract[g_key_count_day].astype(int).sum()
                # const costs
                border_statistic_wells.sum_costs = df_contract[g_key_sum_tex_close].astype(float).sum()
                # costs day
                border_statistic_wells.compute_costs_day()
                key_lub = (lb, ub)
                if not key_lub in data_contract_by_border:
                    data_contract_by_border[key_lub] = {}
                data_contract_by_border[key_lub][key_month] = border_statistic_wells
            # end loop border
        # end loop month
        beg_row = 2
        beg_col = 0
        # write to
        # loop border
        for lb, ub in g_arr_border:
            # get statistic
            all_border_statistic_wells = data_contract_by_border[(lb, ub)]
            # write to sheet
            beg_row = work_book.write_to_sheet(beg_row=beg_row, beg_col=beg_col,
                                               sheet=sheet,
                                               lb=lb, ub=ub,
                                               all_border_statistic_wells=all_border_statistic_wells)
        # end loop border
        work_book.auto_format_cell_width(sheet)
        print("*" * 80)
    # end loop contracts
    work_book.save(path=os.path.join(os.path.dirname(path), "test.xlsm"))
    return


if __name__ == '__main__':
    t0 = time.time()
    main()
    print("job run time: {0:.2f} sec.".format(time.time() - t0))
