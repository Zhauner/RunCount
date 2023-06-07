import pandas
import os

datas = []
text_datas = []
copy_text_datas = []
copy_datas = []
all_runners_nickname = []

for files in os.listdir("Run tables"):
    if files.split('.')[-1] == "xlsx":

        name = files.split('.')[0][-1]

        sheet_names = pandas.ExcelFile(f"Run tables\\{files}").sheet_names
        sheets = pandas.read_excel(f"Run tables\\{files}", sheet_name=sheet_names[0])

        all_data = []

        for x in sheets.columns.ravel():
            all_data.append(sheets[x].to_list())

        columns_count = len(sheets.columns.ravel())
        last_name_count = len(all_data[1])

        for data in range(last_name_count):
            for column in range(columns_count):

                with open(f"Run tables\\{name}.txt", "a", encoding="utf-8") as parse_data:
                    parse_data.write(str(all_data[column][data]).strip(" ") + ' ')
                    if column == 6:
                        parse_data.write("\n")
                    parse_data.close()

for files_txt in os.listdir("Run tables"):

    data = []

    if files_txt.split('.')[-1] == "txt":
        read_data = open(f"Run tables\\{files_txt}", "r", encoding="utf-8").readlines()
        read_data = [x.strip('\n').strip(' ') for x in read_data]
        for x in read_data:
            with open(f"Run tables\\run_{files_txt}", "a", encoding="utf-8") as file:
                file.write(
                    x.split(' ')[1] + ' ' + x.split(' ')[2] \
                    + ' ' + x.split(' ')[3] + ' ' + x.split(' ')[-2] + ' ' + x.split(' ')[-1] + "\n"
                )
        os.remove(f"Run tables\\{files_txt}")
count_file_run = 1
for run_files_txt in os.listdir("Run tables"):

    if run_files_txt.split('.')[-1] == "txt":
        read_data = open(f"Run tables\\{run_files_txt}", "r", encoding="utf-8").readlines()
        read_data_copy = open(f"Run tables\\{run_files_txt}", "r", encoding="utf-8").readlines()

        text_run = read_data
        text_run_copy = read_data_copy
        text_run = [x.strip('\n').strip(' ') for x in text_run]
        text_run_copy = [' '.join(x.strip('\n').strip(' ').split(' ')[:-1]) for x in text_run_copy]

        text_datas.append(text_run)
        copy_text_datas.append(text_run_copy)

        read_data = [x.strip('\n').strip(' ').split(' ') for x in read_data]
        read_data_copy = [y.strip('\n').strip(' ').split(' ')[:-1] for y in read_data_copy]

        datas.append(read_data)
        copy_datas.append(read_data_copy)

        os.remove(f"Run tables\\run_{count_file_run}.txt")
        count_file_run += 1


for list_count_for_merge in range(len(datas)):
    for nickname in datas[list_count_for_merge]:
        all_runners_nickname.append(' '.join(nickname[:-1]))
all_runners_nickname_set = list(set(all_runners_nickname))

for check in all_runners_nickname_set:
    for check_nick in range(len(copy_datas)):
        if check.split(' ') in copy_datas[check_nick]:
            with open(f"Run tables\\result_datas.txt", "a", encoding="utf-8") as all_result_with_name:

                all_result_with_name.write(
                    datas[check_nick][copy_datas[check_nick].index(check.split(' '))][0] + " " + \
                    datas[check_nick][copy_datas[check_nick].index(check.split(' '))][1] + " " + \
                    datas[check_nick][copy_datas[check_nick].index(check.split(' '))][2] + " " + \
                    datas[check_nick][copy_datas[check_nick].index(check.split(' '))][3] + " " + \
                    datas[check_nick][copy_datas[check_nick].index(check.split(' '))][4] + "\n"
                )
            all_result_with_name.close()

        else:
            with open(f"Run tables\\result_datas.txt", "a", encoding="utf-8") as all_result_with_name_t:

                all_result_with_name_t.write(
                    check.split(' ')[0] + " " + \
                    check.split(' ')[1] + " " + \
                    check.split(' ')[2] + " " + \
                    check.split(' ')[3] + " " + \
                    "-" + "\n"
                )
            all_result_with_name_t.close()

result_open = open("Run tables\\result_datas.txt", "r", encoding="utf-8").readlines()

result_open = [x.strip('\n').split(' ') for x in result_open]
counter_users = 1


pre_write_in_exel = []

for write_result_in_list in range(len(result_open)):

    if counter_users > len(datas):
        counter_users = 1

    if counter_users == 1:

        pre_write_in_exel.extend([result_open[write_result_in_list]])
        counter_users += 1

    elif counter_users > 1:

        pre_write_in_exel[-1].extend([result_open[write_result_in_list][-1]])
        counter_users += 1

for pre in range(len(pre_write_in_exel)):
    pre_write_in_exel[pre].append(str(len(datas) - pre_write_in_exel[pre].count("-")))


for count_of_plus in pre_write_in_exel:
    zach = 0
    all_zach = 0
    for run_len in range(len(datas)):
        if count_of_plus[4:-1][run_len] in ['660', '665', '670', '675', '680', '690', '700']:
            zach += 1
        if count_of_plus[4:-1][run_len] != "0" and count_of_plus[4:-1][run_len] != "-":
            all_zach += 1
    count_of_plus.append(str(zach))
    count_of_plus.append(str(all_zach))

for count_of_sum_mark in pre_write_in_exel:
    sum_mark = 0
    for sum_num in range(len(datas)):
        if count_of_sum_mark[4:-2][sum_num] != '-':
            sum_mark += int(count_of_sum_mark[4:-2][sum_num])
    count_of_sum_mark.append(str(sum_mark))

pre_write_in_exel_sort = sorted(pre_write_in_exel, key=lambda win: int(win[-1]), reverse=True)

for win_place, value in enumerate(pre_write_in_exel_sort, start=1):
    value.extend([str(win_place)])

sort_members_by_run_count = reversed(sorted(pre_write_in_exel_sort, key=lambda run: int(run[-4])))
os.remove("Run tables\\result_datas.txt")
# Write data in Exel .xlsx
for datas_for_xlsx in sort_members_by_run_count:
    with open("Run tables\\Merged tables.txt", "a", encoding="utf-8") as result:
        result.write(' '.join(datas_for_xlsx) + "\n")
    result.close()
