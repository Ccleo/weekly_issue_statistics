import time
import xlsxwriter
import os
import sys
from utils.db import DB
from utils.db_search import DbSearch
from utils.jira import JiraSta
from utils.redmine import RedmineSta
from sheet.write_proj_sheet import write_Proj_sheet
from sheet.write_summaryChart_sheet01 import write_summaryChart_sheet01
from sheet.write_summary_sheet02 import write_summary_sheet02
from utils.common import Logger
import config

filename = "log_test" + time.strftime('%Y-%m-%d_%H_%M_%S') + ".txt"
filename_path = os.path.join(config.root_dir, "log", filename)
sys.stdout = Logger(filename_path)
start_time = time.time()
print(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())) + "---loading completed!")


class Statistics(object):
    # options = {"server": "http://192.168.1.212:8088"}
    # auth = ("lig", "lig")  # user_name:test user_passwd:test
    connect = DB().conn()
    """
    type：缺陷管理工具类型
    redmine_project_id：redmine接口识别的项目代号（查看redmine项目配置获取）
    redmine_tracker_id：redmine项目缺陷的跟踪代号（通过打印redmine单个issue的tracker_id获取）
    jira_tag：jira接口识别的项目代号（通过选择JIRA高级查询查看）
    jira_tracker_id：jira查询语句中代表该项目缺陷的代号，一般是“缺陷”/“故障”，新建项目默认为“缺陷”，通过Redmine_find_proj_conf.py脚本获取
    title：项目单独报表的标题
    index：项目在报表中的位置
    sheet_name：项目工作簿名称
    """
    PROJ_DICT = config.PROJ_DICT
    ALL_PROJ_HISTORY_DATA = None
    TOTAL_DATA = DbSearch().search_allDate_history(PROJ_DICT)
    PROJ_BEFORE_BUG = {}  # 新建项目之前未纳入wbs数据库的bug字典
    NEW_BUG = {}  # 所有项目新建bug字典
    LEAVE_BUG = {}  # 所有项目遗留bug字典
    WEEK = DbSearch().week()  # 当前周
    print("当前周:%s" % WEEK)
    NEW_PROJECT = []  # 新增项目列表
    currentWeekSolveRate = {}  # 当前周解决速率
    currentWeekAddRate = {}  # 当前周新增速率
    correSituation = {}  # 当前周对应状况
    filename = '每周项目测试缺陷状况%s.xlsx' % time.strftime('%Y-%m-%d_%H-%M-%S', time.localtime(time.time()))
    reportName = os.path.join(config.root_dir, 'report', filename)

    workbook = xlsxwriter.Workbook(reportName)

    # 类属性更新，一次从数据库获取关键数据
    def data_init(self):
        """
        EBID包含了JIRA和Redmine两部分的数据，从2019年12月份开始，天津专家库和公规院项目开始使用Redmine，所以要计算二者之和
        """
        print("--------------------------data init-----------------------------")
        # dbs = DbSearch()
        # self.TOTAL_DATA = DbSearch().search_allDate_history(self.PROJ_DICT)
        # self.WEEK = DbSearch().week()
        # 更新NEW_PROJECT列表,将新加入的项目加入NEW_PROJECT列表
        self.NEW_PROJECT = DbSearch().search_new_project(self.PROJ_DICT, self.WEEK, self.NEW_PROJECT)
        begintime = time.strftime('%Y-%m-01', time.localtime(time.time()))
        endtime = time.strftime('%Y-%m-%d', time.localtime(time.time()))
        print("统计时间：%s~%s" % (begintime, endtime))
        for proj in self.PROJ_DICT:
            if self.PROJ_DICT[proj]['type'] == 'redmine':
                # redmine_tag = self.PROJ_DICT[proj]['redmine_tag']
                self.NEW_BUG[proj] = RedmineSta().Redmine_build_bug(proj, self.NEW_PROJECT, self.PROJ_DICT)
                self.LEAVE_BUG[proj] = RedmineSta().Redmine_leave_bug(proj, self.PROJ_DICT)
            elif self.PROJ_DICT[proj]['type'] == 'jira':
                # jira_tag = self.PROJ_DICT[proj]['jira_tag']
                self.NEW_BUG[proj] = JiraSta().JIRA_build_bug(proj, self.PROJ_DICT)
                self.LEAVE_BUG[proj] = JiraSta().JIRA_leave_bug(proj, self.PROJ_DICT)
            # todo: 某些项目在jira和redmine都存有数据，将redmine和jira的issue情况存于不同的数据表，方便处理

        self.PROJ_BEFORE_BUG = DbSearch().search_proj_before_bug()
        print('项目之前新建bug', self.PROJ_BEFORE_BUG)

    def main(self):
        self.data_init()
        self.currentWeekSolveRate, self.currentWeekAddRate, self.correSituation = write_summaryChart_sheet01(
            self.workbook,
            self.NEW_BUG,
            self.LEAVE_BUG,
            self.TOTAL_DATA,
            self.PROJ_DICT,
            self.NEW_PROJECT,
            self.PROJ_BEFORE_BUG,
            self.currentWeekSolveRate,
            self.currentWeekAddRate,
            self.correSituation,
            self.WEEK)
        write_summary_sheet02(self.workbook, self.PROJ_DICT, self.WEEK, self.NEW_PROJECT, self.currentWeekSolveRate,
                              self.currentWeekAddRate, self.correSituation)
        # sta.write_BIM_sheet()
        write_Proj_sheet(self.workbook, self.PROJ_DICT, self.PROJ_BEFORE_BUG, self.LEAVE_BUG)
        self.workbook.close()


if __name__ == "__main__":
    sta = Statistics()
    sta.main()
    print(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())) + "---statistics completed!")
    end_time = time.time()
    statiscs_time = str('%.2f' % (end_time - start_time))
    print('用时：%s秒' % statiscs_time)
