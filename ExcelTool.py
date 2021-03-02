import os

import xlwt
import xlrd

import platform
import time

'''
excel 工具之一
需求是，需要把数据汇总，某些项不能重复，如人名

难点：打开一张表后，如果发现里面的数据，在前面的表是不存在的，比如新入职员工，为避免汇总后出现重复或者缺人，
就需要人名同步给前面的表，但是此人的相关数据均为0,；另外如果有离职员工同理，需要再后面的表上依然添加该员工，
只是此人的相关数据均为0

作用是，把 N 份格式一致的表的指定数据提取，然后按照一定的规则，比如a表有b表没有的数据 或者a表没有b表有的数据，
进行处理，汇总到一张表上
'''


class ExcelTool:

    def __init__(self, key_word = ['费用明细表', '社保减免部分', '劳务费', '合计']):
        self.key_word_list = key_word

    def set_key_word(self,key_word_list):
        for key in key_word_list:
            self.key_word_list.append(key)

    def get_key_word_list(self):
        return self.key_word_list

    def get_excel_data(self, path):
        fpath = path
        fname = path.split(r'/')[-1]
        print('fname: ', fname)
        # 初始化
        date = 197001

        key_word = self.get_key_word_list()

        try:
            # # 文件标识，作为字典的key用
            # windows 格式，使用windows是使用
            if platform.system().lower() == 'windows':
                fname = path.split('\\')[-1]
                ftag = fname.split(key_word[2])[0]
                date = ftag
            else:
                # MAC的格式，用苹果电脑时使用
                fname = path.split(r'/')[-1]
                ftag = fname.split(key_word[2])[0]
                date = ftag

        except Exception as e:
            print("{0} 名字格式好像不对:{1}".format(fname, e))

        name_list = []
        price_list = []

        bk = xlrd.open_workbook(fpath)
        sheet_name_list = bk.sheet_names()
        for sheet_name in sheet_name_list:
            # 判断是否有关键字页签，基本上每个excel都有2个
            if key_word[0] in sheet_name:
                sh = bk.sheet_by_name(sheet_name)
                nrows = sh.nrows
                row_item = 0
                row_item_end = 0
                print('{0} 找到页签 {1}'.format(fname, sheet_name))

                # 获取关键字行数
                for i in range(nrows):
                    row_values = sh.row_values(i)
                    # print(row_values)

                    # 找到关键字后退出
                    for item in row_values:
                        # 找开始行
                        if key_word[1] in str(item):
                            row_item = i

                        # 找结束行
                        if key_word[3] == str(item) and i > 5:
                            row_item_end = i

                    # 找到了开始和结束点直接退出循环
                    if row_item != 0 and row_item_end != 0:
                        print('{0} {1} 找到 {2} 栏目'.format(fname, sheet_name, key_word[1]))
                        break

                # 判断是否有关键字行
                if row_item != 0:
                    row_item_values = sh.row_values(row_item)
                    # 需要初始化
                    col = 0
                    for item in row_item_values:
                        # 判断关键字在第几列
                        if key_word[1] in str(item):
                            col = row_item_values.index(item)
                            break

                    if col != 0:
                        # 开始收集数据，从第row_item+1行开始
                        for j in range(row_item + 1, nrows):
                            if j < row_item_end:
                                name = sh.cell(j, 1).value
                                name_list.append(name)
                                price = sh.cell(j, col).value
                                price_list.append(price)
                            else:
                                break
        print('date1:', date)
        print('name_list1: ', name_list)
        print('price_list:', price_list)
        if (len(name_list) == len(price_list)) and (name_list != []):
            return date, name_list, price_list
        else:
            print('%s 没有对应数据' % fname)
            return None

    def file_name(self, file_dir):
        '''
         获取excel名
        :param file_dir:
        :return:
        '''
        files_list = []
        for root, dirs, files in os.walk(file_dir):
            # # 当前目录路径
            # print(root)
            # # 当前路径下所有子目录
            # print(dirs)
            # # 当前路径下所有非目录子文件
            # print(files)
            for file in files:
                if '.xls' in file:
                    files_list.append(file)
            return files_list

    def data_compare(self, root_path, files_list):
        '''
        数据处理，需要处理，新入职员工，把入职前的表补齐这个员工信息，数据均为0
        还需要处理离职员工，把后面离职后的表补齐，数据为0
        :param root_path:
        :param files_list:
        :return:
        '''
        # {date:{name:price,name:price}}
        data_dict = {}
        # 用来记录日期
        date_list = []
        count = 0
        for file in files_list:
            print('data_dict: ', data_dict)
            print('date_list: ', date_list)
            print('count: ', count)
            if '劳务费' in file:
                path = os.path.join(root_path, file)
                data = self.get_excel_data(path)
                if data != None:
                    date = data[0]
                    date_list.append(date)
                    name_list = data[1]
                    price_list = data[2]

                    # 首次加入数据
                    if data_dict == {}:
                        data_dict[date] = dict(zip(name_list, price_list))

                    else:
                        last_date = date_list[count - 1]
                        last_data = data_dict[last_date]
                        last_name_list = last_data.keys()
                        # 新入职
                        for name in name_list:
                            # 发现有新入职的人
                            if name not in last_name_list:
                                # 以前的list要加上新入职的，但是入职之前的费用为0
                                for old_date in date_list[:-1]:
                                    print('old_date: ', old_date)
                                    data_dict[old_date][name] = 0

                        # 已离职
                        for name in last_name_list:
                            # 发现有人离职，需要把离职的也添加上，但是费用为0
                            if name not in name_list:
                                name_list.append(name)
                                price_list.append(0)

                        # 处理完写入数据
                        data_dict[date] = dict(zip(name_list, price_list))

                    count += 1
        return data_dict

    def write_to_excel(self, data, path=None):
        '''
        创建新表，并把数据写入新表
        因为对日期先后有要求，所以需要排序一下
        :param data:
        :param path:
        :return:
        '''

        workbook = xlwt.Workbook(encoding='utf-8')
        # 创建一个worksheet
        worksheet = workbook.add_sheet('data')

        # 写入excel
        date_list = list(data.keys())
        # 排序
        date_list.sort()
        count = 0
        for date in date_list:
            name_list = list(data[date].keys())
            name_list.sort()
            # 日期
            worksheet.write(0, count + 1, label=date)
            for i in range(1, len(name_list)+1):
                # 参数对应 行, 列, 值
                if count == 0:
                    # 人名
                    worksheet.write(i, 0, label=name_list[i-1])
                # 费用
                worksheet.write(i, count+1, label=data[date][name_list[i-1]])

            count += 1

        if path is None:
            now = time.strftime("%Y-%m-%d-%H_%M_%S", time.localtime(time.time()))
            path = now + '_result' + '.xls'

        # 保存
        workbook.save(path)


if __name__ == '__main__':
    root_path = r'D:\PycharmProjects\tool\lilitool\excel_tool\数据'

    path = os.path.join(root_path, r'202012劳务费结算表（宝安应急管理局).xls')

    key_word = ['费用明细表', '社保减免部分', '劳务费', '合计']

    et = ExcelTool(key_word)

    print(et.get_excel_data(path))

    file_list = et.file_name(root_path)
    print(file_list)
    data = et.data_compare(root_path, file_list)
    print(data)
    # # out_put = r'D:\PycharmProjects\tool\lilitool\excel_tool\output.xls'
    et.write_to_excel(data)

