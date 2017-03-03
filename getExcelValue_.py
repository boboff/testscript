"""
Created on 2016年7月4日

@author: lil03
"""
import datetime
import xlrd
import os
import sys
# 获取当前时间并且格式化
date = datetime.datetime.now()
datevalue = date.strftime('%m%d%H%M')

curpath = os.getcwd()
path_object = curpath.replace('Test_PubScripts', 'Test_Data/object_source')
path_data = curpath.replace('Test_PubScripts', 'Test_Data/data_config')
# 设置变量文件列表。
o_paths = []
d_paths = []
# 遍历起始文件
ofiles = os.listdir(path=path_object)
dfiles = os.listdir(path=path_data)
# 添加对象文件
for of in ofiles:
    path = os.path.join(path_object, of)
    if os.path.isfile(path):
        if '.xlsx' == os.path.splitext(of)[1]:
            o_paths.append(path)
# 添加数据文件
for df in dfiles:
    path = os.path.join(path_data, df)
    if os.path.isfile(path):
        if '.xlsx' == os.path.splitext(df)[1]:
            d_paths.append(path)


class DataClass(object):

    def __init__(self, type = None):
        #App_type
        """
        eg. 'android','ios'
        :param type: 传入的app类型,选择加载iOS对象库,or 安卓对象库
        """
        self.appType = type
        # 对象列表集合
        self.objectPaths = o_paths
        # 数据列表集合
        self.dataPaths = d_paths

    @property
    def object_Dict(self):
        o_Dict = {}
        # 遍历所有对象文件
        for o_path in self.objectPaths:
            FileList = self.get_object_info(o_path)
            o_Dict = dict(o_Dict, **FileList)
        return o_Dict

    @staticmethod
    def get_object_info(f_path):
        """

        :param f_path: 单个文件地址
        :return:
        """
        objectDict = {}
        try:
            Workbook = xlrd.open_workbook(f_path)
        except IOError:
            print('数据初始化异常，请检查Object对应Excel文件')
        # 设置数据重复判断点
        # 遍历sheet页的数据
        for sheet in Workbook.sheet_names():
            curSheet = Workbook.sheet_by_name(sheet)
            # 总行数
            rowCount = curSheet.nrows
            # 第一列为标题，忽略掉。
            warningTag = False
            for row in range(1, rowCount):
                # 1,2,3代表需求字段在Excel的列Index
                # 1：对象名称
                # 2：对象属性名
                # 3：对象属性值
                objProperty = str(curSheet.cell_value(row, 2))
                ObjProValue = str(curSheet.cell_value(row, 3))
                objectName = str(curSheet.cell_value(row, 1))
                # 当前列存在重复则提示

                if objectName != '' and objProperty != '' and ObjProValue != '':
                    # 属性以及属性值组成数组
                    PropertyArr = [objProperty, ObjProValue]
                    if objectName in objectDict.keys():
                        warningTag = True
                        print('warning: 当前对象名称%s存在重复，在<%s>页第%s行！' % (objectName, sheet, (row + 1)))

                    # 属性以及对象名称组成字典
                    objectDict[objectName] = PropertyArr
            if warningTag:
                print('         >>>>请纠正后重试...')
                sys.exit(0)
        return objectDict

if __name__ == '__main__':
    value = DataClass().object_Dict
    print(len(value))
    for k,v in value.items():
        print(k,v)
