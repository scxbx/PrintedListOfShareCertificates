import os
import time

import openpyxl


class VillageCommittee:

    def __init__(self, path):
        self.__path = path
        self.__committee_list = []
        self.__fetch_info()

    @property
    def committee_list(self):
        return self.__committee_list

    def __fetch_info(self):
        try:
            group_paths = os.listdir(self.__path)
            for one_path in group_paths:
                full_path = os.path.join(self.__path, one_path)
                village_group = VillageGroup(full_path)
                # print('village_group', village_group)
                if village_group.success:
                    self.__committee_list.extend(village_group.group_list)
        except NotADirectoryError:
            print('{} 不是文件夹，请重试。'.format(self.__path))
        except FileNotFoundError as e:
            print(e.strerror, e.filename)

    def sort_certificate(self):
        def take_certificate(family):
            return family.certificate

        self.__committee_list.sort(key=take_certificate)


class VillageGroup:
    def __init__(self, path):
        self.__path = path
        self.__group_list = []
        self.__success = False
        self.__fetch_info()

    @property
    def group_list(self):
        return self.__group_list

    @property
    def success(self):
        return self.__success

    def __fetch_info(self):
        try:
            family_paths = os.listdir(self.__path)
            for one_path in family_paths:
                full_path = os.path.join(self.__path, one_path)
                # print('full_path', full_path)
                family = Family(full_path)
                if family.success:
                    self.__group_list.append(family)
            self.__success = True
        except NotADirectoryError:
            print('{} 不是文件夹，跳过。'.format(self.__path))
            self.__success = False


class Family:
    @property
    def org_name(self):
        return self.__org_name

    @property
    def credit_code(self):
        return self.__credit_code

    @property
    def certificate(self):
        return self.__certificate

    @property
    def master_name(self):
        return self.__master_name

    @property
    def member_num(self):
        return self.__member_num

    @property
    def path(self):
        return self.__path

    @property
    def success(self):
        return self.__success

    def __init__(self, path):
        self.__success = False
        self.__path = path
        self.__org_name = self.__credit_code = self.__certificate = self.__master_name = self.__member_num = None
        self.__fetch_info()

    def __fetch_info(self):
        if os.path.isdir(self.__path) or '~$' in self.__path or not self.__path.endswith('.xlsx'):
            print('股权证路径：{} 错误!'.format(self.__path))
            self.__success = False
            return

        wb = openpyxl.load_workbook(self.__path, data_only=True)
        ws1 = wb['1']
        ws3 = wb['3']

        self.__org_name = ws1['V31'].value
        self.__credit_code = ws1['X22'].value
        self.__certificate = ws1['W81'].value
        self.__master_name = ws3['S24'].value
        # self.__member_num = ws3['T39']
        try:
            mem_num = int(int(ws3['T39'].value) / 10)
        except ValueError:
            mem_num = ''
            print('股权数合计无法转为整型！请手动修改。')
        self.__member_num = mem_num
        self.__success = True
        wb.close()

    def print_info(self):
        print(self.__org_name, self.__credit_code, self.__certificate, self.__master_name, self.__member_num)

    def print_info2(self):
        print(self.__org_name, self.__credit_code, self.__certificate, self.__master_name, self.__member_num,
              self.__path)


if __name__ == '__main__':
    print('Hello Tuesday!')

    start_time = time.time()
    committee_path = r'股权证\祖坡七个村小组'
    committee = VillageCommittee(committee_path)
    end_time = time.time()
    print('Elapsed time:', end_time - start_time)

    # print('old')
    # for item in committee.committee_list:
    #     item.print_info()
    # print('new')
    committee.sort_certificate()
    # for item in committee.committee_list:
    #     item.print_info()
