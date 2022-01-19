import glob
import os

import PyPDF4


# def get_file_list(path="output", file_type=".pdf") -> list:
#     """
#     Create list of files in directory
#     :param path: like '.'
#     :param file_type: like '.pdf' or '*.pdf'
#     :return: list of files in directory
#     """
#     # files_ = [f for f in os.listdir(path) if os.path.isfile(f) and f.endswith(file_type)]
#     # files = []
#     # for f in os.listdir(''):
#     #     if os.path.isfile(f) and f.endswith('.pdf'):
#     #         files.append(f)
#     files_ = []
#     # os.chdir(path)
#     # for file_ in glob.glob(file_type):
#     for fname in os.listdir(path):
#         if fname.endswith(file_type):
#             files_.append(fname)
#     return files_


# def get_data_from_files(files: list, num_page: int) -> list:
#     """
#     Get some data from files
#     :param num_page: number of page (start from 0)
#     :param files: list of files
#     :return: list of dictionary with some data from files
#     """
#     text_list = []
#     for file in files:
#         file_obj = open('output/' + file, 'rb')
#         file_reader = PyPDF4.PdfFileReader(file_obj)
#         page = file_reader.getPage(num_page)
#         pages_text_list = page.extractText().replace(' \n', '').split("\n")
#         file_dict = {'file': file.title()}
#         for i in range(len(pages_text_list)):
#             # if 'Name of this Investment:' in pages_text_list[i] or 'Unique Investment Identifier (UII):' in \
#             #         pages_text_list[
#             #             i]:
#             #     pages_text_list[i] = pages_text_list[i].replace('1. Name of this Investment:',
#             #                                                     'Name of this Investment').replace(
#             #         '2. Unique Investment Identifier (UII):', 'UII')
#             #
#             #     file_dict[pages_text_list[i]] = pages_text_list[i + 1]
#             if 'Name of this Investment:' in pages_text_list[i]:
#                 # file_dict[pages_text_list[i + 3]] = pages_text_list[i + 1]
#                 file_dict['investment'] = pages_text_list[i + 1]
#                 file_dict['uii'] = pages_text_list[i + 3]
#         text_list.append(file_dict)
#     return text_list


# def compare_data(data_1: list, data_2: list, control_list: list) -> bool:
#     """
#
#     :param data_1:
#     :param data_2:
#     :param control_list:
#     :return:
#     """
#     flag = False
#     # control_list = ['001', '002', '003']
#     for i in control_list:
#         dict_1 = next(item for item in data_1 if item["uii"] == i)
#         dict_2 = next(item for item in data_2 if item["uii"] == i)
#         for j in dict_1:
#             if j not in dict_2:
#                 flag = False
#             else:
#                 flag = True
#     if flag:
#         print('Links and downloads have been matched.')
#     else:
#         print('Links and downloads not matched.')
#     return flag

# if __name__ == '__main__':
#     list_for_check = [{'Data Management and Delivery': '422-000000004'}, {'iTRAK': '422-000001327'}, {'Mission Support Systems': '422-000001328'}]
#     list_data = get_data_from_files(get_file_list('.', '.pdf'), 0)
#     print(compare_data(list_for_check, list_data))
