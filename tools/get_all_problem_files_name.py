import os


def deal_input_number(user_input):
    return_list = []
    user_input_list = user_input.strip().split(' ')
    for item in user_input_list:
        if '-' in item:
            first_num = int(float(item.split('-')[0]))
            second_num = int(float(item.split('-')[1]))
            for i in range(first_num, second_num + 1):
                return_list.append("{:03d}".format(i))
        else:
            return_list.append("{:03d}".format(int(float(item))))
    return return_list


def get_all_problem_files_name(company_name, user_input_num):
    problem_nums_list = deal_input_number(user_input_num)
    all_files = os.listdir(f"{company_name}")
    all_problem_files_name = []
    for problem_num in problem_nums_list:
        for file in all_files:
            if problem_num == file.split('-')[1][-3:]:
                all_problem_files_name.append(f"{company_name}/{file}")
    return all_problem_files_name
