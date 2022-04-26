import pandas as pd
import os
import shutil
import xlsxwriter


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
        if level is current_level:
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


def compare_bom(bom_list: list, compare_list: list, levels_list: list, bom_path: str):
    shutil.copy(bom_path, os.path.dirname(bom_path)+"\\Compare_results.xlsx")


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

    print(find_father(level_list, trial_index))
    print(find_direct_children(level_list, trial_index))
    print(find_all_children(level_list, trial_index))

    print(find_pn_indices(pns_list, compare_pns[1]))

    pass


if __name__ == '__main__':
    main()
