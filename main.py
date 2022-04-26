import pandas as pd
import os
import shutil
from openpyxl import Workbook
from openpyxl import load_workbook
import numpy as np


def find_father(levels_list: list, current_index: int) -> int:
    current_level = levels_list[current_index]
    iterated_levels = levels_list[0:current_index]
    iterated_levels = iterated_levels[::-1]
    for i, level in enumerate(iterated_levels):
        if level < current_level:
            return current_index - i-1


def find_direct_children(levels_list: list, current_index: int) -> list:
    children = []
    current_level = levels_list[current_index]
    for i, level in enumerate(levels_list[current_index+1:]):
        if level is current_level:
            break
        elif level is (current_level+1):
            children.append(current_index+i+1)
    return children


def find_all_children(levels_list: list, current_index: int) -> list:
    children = []
    current_level = levels_list[current_index]
    for i, level in enumerate(levels_list[current_index + 1:]):
        if level <= current_level:
            break
        else:
            children.append(current_index + i + 1)
    return children


def find_pn_indices(pns_list: list, required_pn: str) -> list:
    indices = []
    for i, pn in enumerate(pns_list):
        if pn == required_pn:
            indices.append(i)

    return indices


def all_sons_matched(match_indices: np.array, levels_list: list):
    for i, level in enumerate(levels_list):
        children_indices = find_all_children(levels_list, len(levels_list)-i-1)
        if len(children_indices) != 0:
            all_children_matched = True
            for child in children_indices:
                if child not in match_indices:
                    all_children_matched = False
                    break
            if all_children_matched is True:
                match_indices = np.append(match_indices, len(levels_list)-i-1)
    return match_indices


def find_match_indices(bom_list: list, compare_list: list, levels_list: list):
    match_indices = np.array([], dtype="int16")
    for compare_pn in compare_list:
        pn_indices = find_pn_indices(bom_list, compare_pn)
        match_indices = np.append(match_indices, pn_indices) if pn_indices is not None else []

        for pn_index in pn_indices:
            match_indices = np.append(match_indices, find_all_children(levels_list, pn_index)) if \
                find_all_children(levels_list, pn_index) is not None else []
    match_indices = all_sons_matched(match_indices, levels_list)
    match_indices = np.unique(match_indices)

    return match_indices


def create_bom_comparison(bom_path: str, match_indices: np.array):
    # if os.path.exists(os.path.dirname(bom_path)+"\\Compare_results.xlsx"):
    #     os.remove(os.path.dirname(bom_path)+"\\Compare_results.xlsx")
    # shutil.copy(bom_path, os.path.dirname(bom_path)+"\\Compare_results.xlsx")
    match_indices = match_indices.astype("int")
    wb = load_workbook(os.path.dirname(bom_path)+"\\BOM.xlsx")
    ws = wb.worksheets[0]
    ws.insert_cols(1)
    ws["A1"] = "Is match?"
    for index in match_indices:
        ws[f"A{index + 2}"] = "Match!"
    pass
    wb.save(os.path.dirname(bom_path)+'\\Compare_results.xlsx')
    wb.close()


def main():
    bom_path = "D:\\BOM\\BOM.xlsx"
    compare_path = "D:\\BOM\\Compare.xlsx"
    bom = pd.read_excel(bom_path, sheet_name=0)
    compare = pd.read_excel(compare_path, sheet_name=0)
    level_list = bom['רמת מוצר'].tolist()
    level_list = [int(level.replace(".", "")) for level in level_list]
    compare_pns = compare["מק'ט"].tolist()
    pns_list = bom["מק'ט"].tolist()
    # bom = []
    # pd.ExcelWriter()
    # father_index = 0
    # for i, pn in enumerate(pns_list):
    #     bom.append(PartNumber(pn))
    #     if level_list[father_index] < level_list[i]:
    #         bom[i].add_father(pns_list[father_index])
    trial_index = 50
    #
    # print(find_father(level_list, trial_index))
    # print(find_direct_children(level_list, trial_index))
    # print(find_all_children(level_list, trial_index))
    #
    # print(find_pn_indices(pns_list, compare_pns[1]))
    match_indices = find_match_indices(pns_list, compare_pns, level_list)
    create_bom_comparison(bom_path, match_indices)
    pass


if __name__ == '__main__':
    main()
