# -*- coding: utf-8 -*-

from xmindparser import xmind_to_dict
import xlwt

#用例字段
CASE_TITLE = '用例名称'
CASE_TYPE = '用例类型'
CASE_STATUS = '用例状态'
CASE_LEVEL = '用例等级'
CASE_CREATOR = '创建人'
CASE_STEP = '用例步骤'
CASE_PRE = '前置条件'
CASE_EXPECTED = '预期结果'
CASE_DIR = '用例目录'
REQURIEMENT_ID = '需求ID'

DEFAULT_PRE_DIR = '需求管理-V1.0-'
DEFAULT_LEVEL = '中'

#脑图中的用例字段
STEP_KEYS = ['用例步骤','操作步骤：','操作步骤']
EXPECTED_KEYS = ['预期结果','预期结果：']
LEVEL_KEYS = ['用例等级','优先级：']

class XmindToExcel:

    def __init__(self,path):
        self.path = path
        content = xmind_to_dict(self.path)
        self.datas = content[0]['topic']['topics']
        self.excel_col = [CASE_TITLE,CASE_TYPE,CASE_STATUS,CASE_LEVEL,CASE_CREATOR,CASE_STEP,CASE_EXPECTED,CASE_DIR,REQURIEMENT_ID]
        self.book = xlwt.Workbook()
        self.sheet = self.book.add_sheet('cases')

        self.case_types = []
        self.case_status = []
        self.case_levels = []
        self.case_steps = []
        self.case_expecteds = []
        self.case_titles = []
        self.case_dir = []
        self.requirement_ids = []

    def getColIndexByName(self,col_name):
        return self.excel_col.index(col_name)

    def dataSet(self):
        '''
        处理一些默认数据
        :return:
        '''
        self.setCaseTypes()

    def setCaseTypes(self):
        if len(self.case_types) < self.getCaseNums():
            if len(self.case_types) == 0:
                self.case_types = ['功能测试' for i in range(self.getCaseNums())]
            else:
                tmp = ['功能测试' for i in range(len(self.case_types), self.getCaseNums())]
                for type in tmp:
                    self.case_types.append(type)

    def writeToExecl(self):
        row  = 0
        col = 0
        for col_name in self.excel_col:
            self.sheet.write(row,col,col_name)
            col += 1

        row = 1
        for case_title in self.case_titles:
            self.sheet.write(row,self.getColIndexByName(CASE_TITLE),case_title)
            self.sheet.write(row,self.getColIndexByName(CASE_TYPE),self.case_types[row-1])
            self.sheet.write(row,self.getColIndexByName(CASE_STEP),self.case_steps[row-1])
            self.sheet.write(row, self.getColIndexByName(CASE_EXPECTED), self.case_expecteds[row - 1])
            self.sheet.write(row, self.getColIndexByName(CASE_DIR), self.case_dir[row - 1])
            self.sheet.write(row, self.getColIndexByName(CASE_LEVEL), self.case_levels[row - 1])
            self.sheet.write(row, self.getColIndexByName(REQURIEMENT_ID), self.requirement_ids[row - 1])
            row += 1

        self.book.save('./test.xls')

    def getCaseNums(self):
        return len(self.case_titles)

    # 获取用例等级。eg：用例等级：高 或 优先级：高
    def getCaseLevel(self,case_data):
        info = ''
        for tmp in case_data:
            if CASE_LEVEL in tmp['title'] or '优先级' in tmp['title']:  # 用例等级解析
                info = tmp['title'].split('：')[1]  # 获取用例等级 eg：用例等级：高
                break
            else:
                info = DEFAULT_LEVEL
        self.case_levels.append(info)

    # 获取用例等级。eg：需求ID：xxxx 或 关联需求ID：xxxxx
    def getRequirementID(self,case_data):
        info = ''
        for tmp in case_data:
            if REQURIEMENT_ID in tmp['title'] :  # 用例等级解析
                info = tmp['title'].split('：')[1]  # 获取用例等级 eg：用例等级：高
                break
        self.requirement_ids.append(info)

    # 根据脑图中的key获取用例内容，主要用来获取用例步骤和预期结果
    def getCaseInfoByKey(self,case_data,case_info=[CASE_TITLE]):
        info = ''
        for tmp in case_data:
            if tmp['title'] in case_info :
                tmp_data = tmp['topics']
                if isinstance(tmp_data,list):
                    for tmp_reslt in tmp_data:
                        info += tmp_reslt['title']
                        info += '\n'
                break
        if len(set(EXPECTED_KEYS) & set(case_info)) != 0:
            self.case_expecteds.append(info)
        elif len(set(STEP_KEYS) & set(case_info)) !=0:
            self.case_steps.append(info)
        return str(info)

    def parse(self):
        for mod in self.datas:
            # 模块下确认有子模块以及用例
            if 'topics' in mod:
                sub_mods = mod['topics']
                #print("========================子模块==========================")
                #print("sub_mods",sub_mods)
                #print("========================子模块==========================")
                for sub_mod in sub_mods:
                    #print("sub_mod",sub_mod)
                    # 子模块下有用例
                    if 'topics' in sub_mod:
                        test_cases = sub_mod['topics']
                        #print('test_cases',test_cases)
                        #print(len(test_cases))
                        for test_case in test_cases:
                            case_title = test_case['title']
                            self.case_titles.append(case_title)
                            if 'topics' in test_case:
                                case_data = test_case['topics']
                                self.getCaseLevel(case_data)
                                self.getRequirementID(case_data)
                                self.getCaseInfoByKey(case_data, STEP_KEYS)
                                self.getCaseInfoByKey(case_data,EXPECTED_KEYS)
                            else:
                                self.case_expecteds.append("")
                                self.case_steps.append("")
                                self.case_levels.append(DEFAULT_LEVEL)
                                self.requirement_ids.append("")

                            # 设置用例目录=默认名-模块名-子模块
                            self.case_dir.append(DEFAULT_PRE_DIR+mod['title']+'-'+sub_mod['title'])

    def write_data(self):
        with open('./files/data.txt', 'w+') as f:
            f.write(str(self.datas))


if __name__ == '__main__':
    xmind = XmindToExcel('files/需求管理_1.0.xmind')
    xmind.parse()
    xmind.dataSet()
    xmind.writeToExecl()