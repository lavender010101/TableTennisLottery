import utils
import random

file_name = ''
headers = []


def shuffle(sheet, groups):
    seed, unseed = divide_seeds(sheet, headers)
    random.shuffle(unseed) # 非种子选手随机打乱
    items = seed + unseed

    rows = len(items)
    group_maps = {}
    for i in range(rows):
        if i % groups in group_maps.keys():
            group_maps[i % groups].append(items[i] + [i % groups + 1])
        else:
            group_maps[i % groups] = [items[i] + [i % groups + 1]]

    data_headers = ['学号', '姓名', '小组']
    data = [data_headers]
    for group in group_maps.values():
        for item in group:
            data.append(item)
        data.append([])
        data.append([])

    return data


def divide_seeds(sheet, headers):
    seed = []
    unseed = []
    colums = utils.get_columns(sheet, headers)

    for row in sheet.get_rows():
        if row[colums['签到']].value.upper() not in ['Y', 'YES']:
            continue
        if row[colums['种子']].value.upper() in ['Y', 'YES']:
            seed.append([row[colums['学号']].value, row[colums['名字']].value])
        elif row[colums['种子']].value.upper() in ['N', 'NO']:
            unseed.append([row[colums['学号']].value, row[colums['名字']].value])

    return seed, unseed


if __name__ == "__main__":
    male_groups = input("请输入男子分组数：")
    female_groups = input("请输入女子分组数：")
    male, female = utils.get_sheets(file_name)

    male_data = shuffle(male, male_groups)
    female_data = shuffle(female, female_groups)

    utils.save_data('../res/分组名单.xls', ['男子', '女子'], [male_data, female_data])
