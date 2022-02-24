import utils

# 在引号内填入源文件的路径，包含后缀名
input_file = '' # 需要进行提取的文件
output_file = '' # 提取后输出文件的文件名
headers = ['学号', '姓名', '学院', '性别']


# 简单地通过学号是否为12位判断学号是否有效
def isValid(student_no):
    if len(student_no) != 12:
        return False

    return True


def extract_data(sheet, headers):
    columns = utils.get_columns(sheet, headers)

    informations = {}

    for i in range(1, sheet.nrows):
        student_no = sheet.cell(i, columns['学号']).value
        if isValid(student_no) == False:
            continue

        items = []
        for j in columns.values():
            items.append(sheet.cell(i, j).value)

        informations[student_no] = items

    ret = [headers]
    for info in informations.values():
        ret.append(info)

    return ret


def divide_by_gender(data):
    gender_col = data[0].index('性别')

    male = [data[0]]
    female = [data[0]]
    for item in data[1:]:
        if '男' in item[gender_col]:
            # print(item[gender_col])
            male.append(item)
        else:
            female.append(item)

    return male, female


if __name__ == "__main__":

    sheets = utils.get_sheets(input_file)
    data = extract_data(sheets[0], headers)

    male, female = divide_by_gender(data)

    utils.save_data(output_file, ['男子', '女子'], [male, female])
