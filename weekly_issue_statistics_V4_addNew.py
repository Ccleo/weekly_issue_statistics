from redminelib import Redmine
import datetime
import time
from jira import JIRA
import pymysql
import matplotlib.pyplot as plt
import matplotlib
import matplotlib.ticker as mtick
import xlsxwriter
import os
import shutil
from pathlib import Path


def conn():
    connect = pymysql.Connect(
        host='192.168.1.245',
        port=3306,
        user='root',
        passwd='123456',
        db='wbs',
        charset='utf8'
    )
    return connect


def reConnect(self):
    try:
        self.ping()
    except:
        self.connection()


class Statistics(object):
    redmine = Redmine('http://192.168.1.212:7777/redmine', username='redmine', password='redmine')
    # options = {"server": "http://192.168.1.212:8088"}
    # auth = ("lig", "lig")  # user_name:test user_passwd:test
    jira = JIRA({"server": "http://192.168.1.212:8088"}, basic_auth=("lig", "lig"))
    connect = conn()
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
    PROJ_DICT = {
        'BIM': {'type': 'redmine', 'redmine_project_id': 'fjbim', 'redmine_tracker_id': '7', 'title': '福建BIM项目测试Bug状况',
                'index': 1, 'sheet_name': '福建BIM项目'},
        'EBID': {'type': 'jira', 'jira_project_id': 'EBID', 'jira_tracker_id': '缺陷', 'title': '电子招投标项目测试Bug状况',
                 'index': 2, 'sheet_name': '电子招投标系统'},
        'OA': {'type': 'jira', 'jira_project_id': 'OA', 'jira_tracker_id': '故障', 'title': '华杰OA系统项目测试Bug状况',
               'index': 3, 'sheet_name': '华杰OA系统'},
        'EBID-CCCC': {'type': 'redmine', 'redmine_project_id': 'ebid-cccc', 'redmine_tracker_id': '7',
                      'title': '公规院电子招标采购系统项目测试Bug状况', 'index': 6, 'sheet_name': '公规院电子招标采购项目'},
        'FAS': {'type': 'redmine', 'redmine_project_id': 'hsdfas', 'redmine_tracker_id': '7',
                'title': '勘察设计外业采集系统项目测试Bug状况',
                'index': 4, 'sheet_name': '勘察设计外业采集系统'},
        'EVAL': {'type': 'redmine', 'redmine_project_id': 'eval', 'redmine_tracker_id': '15', 'title': '考评系统项目测试Bug状况',
                 'index': 5, 'sheet_name': '考评系统'},
        'EXPERT_TJ': {'type': 'redmine', 'redmine_project_id': 'expert_tj', 'redmine_tracker_id': '7',
                      'title': '天津评标专家管理系统项目测试Bug状况', 'index': 7, 'sheet_name': '天津评标专家管理系统项目'},
        'PMS': {'type': 'redmine', 'redmine_project_id': 'pms', 'redmine_tracker_id': '15',
                'title': '国际工程项目信息管理系统项目测试Bug状况', 'index': 8, 'sheet_name': '国际工程项目信息管理系统'}}
    # 'ECAP': {'type': 'redmine', 'redmine_project_id': 'ecap', 'redmine_tracker_id': '7',
    #          'title': '应急资源管理平台项目测试Bug状况', 'index': 9, 'sheet_name': '应急资源管理平台'}}
    ALL_PROJ_HISTORY_DATA = None
    TOTAL_DATA = None
    PROJ_BEFORE_BUG = {}
    NEW_BUG = {}  # 所有项目新建bug字典
    LEAVE_BUG = {}  # 所有项目遗留bug字典
    YEAR = time.strftime('%Y', time.localtime(time.time()))
    MONTH = time.strftime('%m', time.localtime(time.time()))
    WEEK = None  # 当前周
    NEW_PROJECT = []  # 新增项目列表
    currentWeekSolveRate = {}  # 当前周解决速率
    currentWeekAddRate = {}  # 当前周新增速率
    correSituation = {}  # 当前周对应状况
    # workbook = xlwt.Workbook(encoding='utf-8')
    current_time = time.strftime('%Y-%m-%d_%H-%M-%S', time.localtime(time.time()))
    reportName = os.path.join('D:\\Data_AT_Statistics\\每周项目测试缺陷状况\weekly_issue_statistics\\report\\',
                              '--每周项目测试缺陷状况%s.xlsx' % time.strftime('%Y-%m-%d_%H-%M-%S', time.localtime(time.time())))
    workbook = xlsxwriter.Workbook(reportName)

    def reConnect(self):
        try:
            self.connect.ping()
        except:
            self.connect = conn()

    # 类属性更新，一次从数据库获取关键数据
    def data_init(self):
        """
        EBID包含了JIRA和Redmine两部分的数据，从2019年12月份开始，天津专家库和公规院项目开始使用Redmine，所以要计算二者之和
        """
        print("--------------------------data init-----------------------------")
        self.WEEK = self.week()
        self.search_new_project()  # 更新NEW_PROJECT列表,将新加入的项目加入NEW_PROJECT列表
        self.TOTAL_DATA = self.search_allDate_history()
        begintime = time.strftime('%Y-%m-01', time.localtime(time.time()))
        endtime = time.strftime('%Y-%m-%d', time.localtime(time.time()))
        print("统计时间：%s~%s" % (begintime, endtime))
        for proj in self.PROJ_DICT:
            if self.PROJ_DICT[proj]['type'] == 'redmine':
                # redmine_tag = self.PROJ_DICT[proj]['redmine_tag']
                self.NEW_BUG[proj] = self.Redmine_build_bug(proj)
                self.LEAVE_BUG[proj] = self.Redmine_leave_bug(proj)
            elif self.PROJ_DICT[proj]['type'] == 'jira':
                # jira_tag = self.PROJ_DICT[proj]['jira_tag']
                self.NEW_BUG[proj] = self.JIRA_build_bug(proj)
                self.LEAVE_BUG[proj] = self.JIRA_leave_bug(proj)
            # todo: 某些项目在jira和redmine都存有数据，将redmine和jira的issue情况存于不同的数据表，方便处理

        self.search_proj_before_bug()
        print(self.PROJ_BEFORE_BUG)

    # redmine当前月的bug新建情况
    def Redmine_build_bug(self, proj):
        """
        :param tag:
        :param start_time: 07-01
        :param end_time: 07-31
        :return: bug字典,当前时间段bug总数
        """
        # project = self.redmine.project.get(tag)
        # start_time = time.strftime('%m-01', time.localtime(time.time()))
        # end_time = time.strftime('%m-%d', time.localtime(time.time()))

        start_time = time.strftime('%Y-%m-01', time.localtime(time.time())) + " 08:00:0"
        end_time = time.strftime('%Y-%m-%d', time.localtime(time.time())) + " 08:00:0"

        start_time = datetime.datetime.utcfromtimestamp(
            int(time.mktime(time.strptime(start_time, "%Y-%m-%d %H:%M:%S")))).strftime("%Y-%m-%d")
        end_time = datetime.datetime.utcfromtimestamp(
            int(time.mktime(time.strptime(end_time, "%Y-%m-%d %H:%M:%S")))).strftime("%Y-%m-%d")

        if proj in self.NEW_PROJECT:  # 新建项目，统计所有已创建bug
            issues_list = self.redmine.issue.filter(project_id=self.PROJ_DICT[proj]['redmine_project_id'],
                                                    status_id="*",
                                                    tracker_id=self.PROJ_DICT[proj]['redmine_tracker_id'],
                                                    )
            # print(len(issues_list))
            issues_list_before = self.redmine.issue.filter(project_id=self.PROJ_DICT[proj]['redmine_project_id'],
                                                           status_id="*",
                                                           tracker_id=self.PROJ_DICT[proj]['redmine_tracker_id'],
                                                           created_on='<=' + str(start_time))

            created_bug_num = len(issues_list)  # 非新建项目，只统计当前月新建bug数量
            print("【Redmine】%s 新建bug数量为:%d " % (proj, created_bug_num))
            proj_before_bug_num = int(len(issues_list_before))
            print("【Redmine】%s %s月份之前新建bug数量为:%d " % (proj, self.MONTH, created_bug_num))
            if proj_before_bug_num > 0:
                connect = conn()
                cur = connect.cursor()
                sql = 'insert into wbg_proj_before_bug (proj,beforeNum) values (%(proj)s,%(beforeNum)s)'
                value = {'proj': proj, 'before': proj_before_bug_num}
                # print('------------------------')
                # print(value)
                cur.execute(sql, value)
                connect.commit()
            return created_bug_num

        else:
            issues_list = self.redmine.issue.filter(project_id=self.PROJ_DICT[proj]['redmine_project_id'],
                                                    status_id="*",
                                                    tracker_id=self.PROJ_DICT[proj]['redmine_tracker_id'],
                                                    created_on='><' + str(start_time) + '|' + str(end_time))
            # print(len(issues_list))
            created_bug_num = len(issues_list)  # 当前新建bug数量
            print("【Redmine】%s 新建bug数量为:%d " % (proj, created_bug_num))
            return created_bug_num

    # JIRA当前月的bug新建情况
    def JIRA_build_bug(self, proj):
        begintime = time.strftime('%Y-%m-01', time.localtime(time.time()))
        # print(begintime)
        # t = time.strftime('%d', time.localtime(time.time()))
        endtime = (datetime.datetime.now() + datetime.timedelta(days=1)).strftime("%Y-%m-%d")
        # print(endtime)
        JQL = "project in (%s) AND issuetype = %s AND created >= %s AND created <= %s" % (
            self.PROJ_DICT[proj]['jira_project_id'], self.PROJ_DICT[proj]['jira_tracker_id'], begintime, endtime)
        # print(JQL)
        issue_list = self.jira.search_issues(JQL, maxResults=100000)
        # project_name = {"EBID": "EBID", "OA": "OA系统"}
        print("【JIRA】%s 新建bug数量: %d" % (proj, len(issue_list)))
        bug_build_num = len(issue_list)
        return bug_build_num

    # rdemine每次统计时所有bug遗留情况
    def Redmine_leave_bug(self, proj):
        """统计状态为打开的bug以及开发遗留bug情况(某项目）"""
        # 统计开发遗留bug数量和遗留时长
        issues_list = self.redmine.issue.filter(project_id=self.PROJ_DICT[proj]['redmine_project_id'], status="打开",
                                                tracker_id=self.PROJ_DICT[proj]['redmine_tracker_id'])
        leave_bug_num = len(issues_list)  # 当前新建bug数量
        print("【Redmine】%s 遗留bug数量为:%d " % (proj, leave_bug_num))
        bug_owner_list = []  #
        bug_owner_1week = {}
        bug_owner_1_2week = {}
        bug_owner_2week_1month = {}
        bug_owner_1month = {}
        bugStatusYear = {self.YEAR: {'1': 0, '2': 0, '3': 0, '4': 0, '5': 0, '6': 0, '7': 0, '8': 0,
                                     '9': 0, '10': 0, '11': 0, '12': 0}}
        for issue in issues_list:
            leave_days = (datetime.datetime.now() - issue.created_on).days
            issueCreateMonth = str(issue.created_on.month)
            issueCreateYear = str(issue.created_on.year)
            if issueCreateYear == self.YEAR:
                bugStatusYear[self.YEAR][issueCreateMonth] += 1
            else:
                if issueCreateYear not in bugStatusYear:
                    bugStatusYear[issueCreateYear] = 1
                else:
                    bugStatusYear[issueCreateYear] += 1

            try:
                bug_owner = issue.assigned_to.name
            except:
                print("【没有指派】———%s" % issue.id)
            else:
                if bug_owner not in bug_owner_list:
                    bug_owner_list.append(bug_owner)
                if leave_days >= 0 and leave_days <= 7:
                    if bug_owner not in bug_owner_1week:
                        bug_owner_1week[issue.assigned_to.name] = 1
                    else:
                        bug_owner_1week[issue.assigned_to.name] += 1
                elif leave_days > 7 and leave_days <= 14:
                    if bug_owner not in bug_owner_1_2week:
                        bug_owner_1_2week[issue.assigned_to.name] = 1
                    else:
                        bug_owner_1_2week[issue.assigned_to.name] += 1
                elif leave_days > 14 and leave_days <= 30:
                    if bug_owner not in bug_owner_2week_1month:
                        bug_owner_2week_1month[issue.assigned_to.name] = 1
                    else:
                        bug_owner_2week_1month[issue.assigned_to.name] += 1
                elif leave_days > 30:
                    if bug_owner not in bug_owner_1month:
                        bug_owner_1month[issue.assigned_to.name] = 1
                    else:
                        bug_owner_1month[issue.assigned_to.name] += 1
        # print(bug_owner_1week)
        # print(bug_owner_1_2week)
        # print(bug_owner_2week_1month)
        # print(bug_owner_1month)
        owner_sta_dict = {}
        for i in bug_owner_list:
            owner_sta_dict[i] = {'一周': 0, "一周~二周": 0, "二周~一月": 0, "一月以上": 0}
            if i in bug_owner_1week:
                owner_sta_dict[i]['一周'] = bug_owner_1week[i]
            if i in bug_owner_1_2week:
                owner_sta_dict[i]['一周~二周'] = bug_owner_1_2week[i]
            if i in bug_owner_2week_1month:
                owner_sta_dict[i]['二周~一月'] = bug_owner_2week_1month[i]
            if i in bug_owner_1month:
                owner_sta_dict[i]['一月以上'] = bug_owner_1month[i]
        print("proj:", proj)
        owner_sta_dict['proj'] = proj
        owner_sta_dict['lbn'] = leave_bug_num
        owner_sta_dict['bugStatusYear'] = bugStatusYear
        print(owner_sta_dict)
        return owner_sta_dict

    # JIRA每次统计时所有bug遗留情况
    def JIRA_leave_bug(self, proj):

        JQL = 'project = %s AND issuetype = %s AND resolution = Unresolved ORDER BY priority DESC, updated DESC' % (
            self.PROJ_DICT[proj]['jira_project_id'], self.PROJ_DICT[proj]['jira_tracker_id'])
        issues_list = self.jira.search_issues(JQL, maxResults=10000)
        leave_bug = len(issues_list)

        bug_owner_list = []
        bug_owner_1week = {}
        bug_owner_1_2week = {}
        bug_owner_2week_1month = {}
        bug_owner_1month = {}
        bugStatusYear = {self.YEAR: {'1': 0, '2': 0, '3': 0, '4': 0, '5': 0, '6': 0, '7': 0, '8': 0,
                                     '9': 0, '10': 0, '11': 0, '12': 0}}
        for issue in issues_list:
            leave_days = (datetime.datetime.now() - datetime.datetime.strptime(
                issue.raw['fields']['created'].split('T')[0], '%Y-%m-%d')).days
            issueCreateMonth = str(
                datetime.datetime.strptime(issue.raw['fields']['created'].split('T')[0], '%Y-%m-%d').month)
            issueCreateYear = str(
                datetime.datetime.strptime(issue.raw['fields']['created'].split('T')[0], '%Y-%m-%d').year)

            if issueCreateYear == self.YEAR:
                bugStatusYear[self.YEAR][issueCreateMonth] += 1
            else:
                if issueCreateYear not in bugStatusYear:
                    bugStatusYear[issueCreateYear] = 1
                else:
                    bugStatusYear[issueCreateYear] += 1
            try:
                bug_owner = issue.raw['fields']['assignee']['displayName']
            except:
                print("【没有指派】———%s" % issue)
            else:
                if bug_owner not in bug_owner_list:
                    bug_owner_list.append(bug_owner)
                if leave_days >= 0 and leave_days <= 7:
                    if bug_owner not in bug_owner_1week:
                        bug_owner_1week[bug_owner] = 1
                    else:
                        bug_owner_1week[bug_owner] += 1
                elif leave_days > 7 and leave_days <= 14:
                    if bug_owner not in bug_owner_1_2week:
                        bug_owner_1_2week[bug_owner] = 1
                    else:
                        bug_owner_1_2week[bug_owner] += 1
                elif leave_days > 14 and leave_days <= 30:
                    if bug_owner not in bug_owner_2week_1month:
                        bug_owner_2week_1month[bug_owner] = 1
                    else:
                        bug_owner_2week_1month[bug_owner] += 1
                elif leave_days > 30:
                    if bug_owner not in bug_owner_1month:
                        bug_owner_1month[bug_owner] = 1
                    else:
                        bug_owner_1month[bug_owner] += 1
        # print(bug_owner_1week)
        # print(bug_owner_1_2week)
        # print(bug_owner_2week_1month)
        # print(bug_owner_1month)
        owner_sta_dict = {}
        for i in bug_owner_list:
            owner_sta_dict[i] = {'一周': 0, "一周~二周": 0, "二周~一月": 0, "一月以上": 0}
            if i in bug_owner_1week:
                owner_sta_dict[i]['一周'] = bug_owner_1week[i]
            if i in bug_owner_1_2week:
                owner_sta_dict[i]['一周~二周'] = bug_owner_1_2week[i]
            if i in bug_owner_2week_1month:
                owner_sta_dict[i]['二周~一月'] = bug_owner_2week_1month[i]
            if i in bug_owner_1month:
                owner_sta_dict[i]['一月以上'] = bug_owner_1month[i]
        owner_sta_dict['proj'] = proj
        owner_sta_dict['lbn'] = leave_bug
        owner_sta_dict['bugStatusYear'] = bugStatusYear
        print("【%s】JIRA遗留bug" % proj, owner_sta_dict)
        return owner_sta_dict

    def search_proj_before_bug(self):
        title = '查询项目纳入统计之前的新建bug数'
        print(title.center(40, '-'))
        connect = self.connect
        self.reConnect()
        cur = connect.cursor()
        sql = 'select * from wbg_proj_before_bug'
        cur.execute(sql)
        result = cur.fetchall()
        print(result)
        for i in result:
            self.PROJ_BEFORE_BUG[i[0]] = i[1]

    def update_wbg_year_sta(self):
        print("--------------------------update wbg_year_sta-----------------------------")
        # newBug = {'BIM': self.BIM['newBug'], 'EBID': self.EBID['newBug'], 'OA': self.OA['newBug'],
        #           'EDU': self.EDU['newBug'], 'FAS': self.FAS['newBug'], 'EVAL': self.Redmine_build_bug("eval")}
        #
        # # 更新数据库wbg_year_sta
        # leaveBug = {'BIM': self.BIM['leaveBug'], 'EBID': self.EBID['leaveBug'],
        #             'OA': self.OA['leaveBug'], 'EDU': self.EDU['leaveBug'],
        #             'FAS': self.FAS['leaveBug'], 'EVAL': self.EVAL['leaveBug']}

        # 根据当前年月去数据库 wbg_year_sta表查询数据，更新每月遗留bug数，新建bug数
        currentYear = self.YEAR
        print('year:', currentYear)
        currentMonth = int(self.MONTH)
        print('month:', currentMonth)
        connect = self.connect
        self.reConnect()
        cur = connect.cursor()
        updateTime = datetime.datetime.now()
        # updateTime = '2019-10-08 20:28:26'
        for proj in self.NEW_BUG:
            # 更新当前年份的之前月份（不用更新build_bug_num），直接更新
            # if self.YEAR in self.NEW_BUG[proj]['bugStatusYear']:
            print(proj)
            # print(self.NEW_BUG)
            for i in range(1, currentMonth):
                sql = 'update wbg_year_sta set leave_bug_num=%(lbn)s,update_time=%(ut)s where year=%(currentYear)s and month=%(currentMonth)s and project=%(proj)s'
                value = {'lbn': self.LEAVE_BUG[proj]['bugStatusYear'][self.YEAR][str(i)],
                         'currentYear': currentYear,
                         'currentMonth': str(i), 'proj': proj, 'ut': updateTime}
                cur.execute(sql, value)
                connect.commit()

            # 更新当前月
            sql = 'select * from wbg_year_sta where year=%(currentYear)s and month=%(currentMonth)s and project=%(proj)s'
            value = {'currentYear': currentYear, 'currentMonth': str(currentMonth), 'proj': proj}

            try:
                cur.execute(sql, value)
                result = cur.fetchone()
            except:
                print('查询wbg_year_sta表出错')
            else:
                if result:  # 如果查到数据，就更新数据表
                    sql = 'update wbg_year_sta set build_bug_num = %(bbn)s,leave_bug_num=%(lbn)s,update_time=%(ut)s where year=%(currentYear)s and month=%(currentMonth)s and project=%(proj)s'
                    value = {'bbn': self.NEW_BUG[proj],
                             'lbn': self.LEAVE_BUG[proj]['bugStatusYear'][self.YEAR][str(currentMonth)],
                             'currentYear': currentYear,
                             'currentMonth': currentMonth, 'proj': proj, 'ut': updateTime}
                    cur.execute(sql, value)
                    connect.commit()
                else:
                    sql = 'insert into wbg_year_sta (project,year,month,leave_bug_num,build_bug_num,update_time) values (%(proj)s,%(year)s,%(month)s,%(lbn)s,%(bbn)s,%(ut)s)'
                    value = {'proj': proj, 'year': currentYear, 'month': currentMonth,
                             'lbn': self.LEAVE_BUG[proj]['bugStatusYear'][self.YEAR][str(currentMonth)],
                             'bbn': self.NEW_BUG[proj], 'ut': updateTime}
                    # print('------------------------')
                    # print(value)
                    cur.execute(sql, value)

            # 更新之前年份
            before_year_dict = self.LEAVE_BUG[proj]['bugStatusYear']
            before_year_dict.pop(self.YEAR)  # 删除当前年份
            for year in before_year_dict:
                sql = 'select * from wbg_year_sta where year=%(currentYear)s and project=%(proj)s'
                value = {'currentYear': year, 'proj': proj}
                try:
                    cur.execute(sql, value)
                    result = cur.fetchone()
                except:
                    print('查询wbg_year_sta表出错')
                else:
                    if result:  # 如果查到数据，就更新数据表
                        sql = 'update wbg_year_sta set leave_bug_num=%(lbn)s,update_time=%(ut)s where year=%(currentYear)s and  project=%(proj)s'
                        value = {'lbn': before_year_dict[year],
                                 'currentYear': year, 'proj': proj, 'ut': updateTime}
                        cur.execute(sql, value)

        connect.commit()

    # 更新wbg_history_data表

    def update_wbg_history_data(self, proj, wn, st, bln, btn, blr, ban, bsn):
        connect = self.connect
        self.reConnect()
        cur = connect.cursor()
        sql = 'select * from wbg_history_data where week_num=%(wn)s and project = %(proj)s'
        value = {'wn': self.WEEK, 'proj': proj}
        cur.execute(sql, value)
        result = cur.fetchall()
        # print(result)
        # print(type(blr))
        # print(blr)
        if result:
            print("%s周数据已存入数据表,更新当前周【%s】项目数据" % (self.WEEK, proj))
            sql = 'update wbg_history_data set project=%(proj)s ,week_num=%(wn)s,statis_time=%(st)s,bug_leave_num=%(bln)s,bug_total_num=%(btn)s,bug_leave_rate=%(blr)s,bug_add_num=%(ban)s,bug_solve_num=%(bsn)s where project=%(proj)s and week_num=%(wn)s'
            value = {'proj': proj, 'wn': wn, 'st': st, 'bln': bln, 'btn': btn, 'blr': blr, 'ban': ban, 'bsn': bsn}
            cur.execute(sql, value)
            connect.commit()

        else:
            print("更新wbg_histroy_data数据表，插入%s周【%s】项目数据" % (self.WEEK, proj))
            sql = 'insert into wbg_history_data (project,week_num,statis_time,bug_leave_num,bug_total_num,bug_leave_rate,bug_add_num,bug_solve_num) values (%(proj)s,%(wn)s,%(st)s,%(bln)s,%(btn)s,%(blr)s,%(ban)s,%(bsn)s)'
            value = {'proj': proj, 'wn': wn, 'st': st, 'bln': bln, 'btn': btn, 'blr': blr, 'ban': ban, 'bsn': bsn}
            cur.execute(sql, value)
            connect.commit()

    # 统计当前周数
    def week(self):
        connect = self.connect
        self.reConnect()
        cur = connect.cursor()
        sql = 'select last_week from wbg_week where id = 0'
        cur.execute(sql)
        week = cur.fetchall()[0][0]
        print("当前周数：", week)
        # print(type(week))
        connect.close()
        return week

    # 查询是否为新增项目(查询wbg_history_data表中的上周数据是否有该proj的值)
    def search_new_project(self):
        title = "查询新增项目"
        print(title.center(40, '-'))
        connect = self.connect
        self.reConnect()
        cur = connect.cursor()
        for proj in self.PROJ_DICT:
            sql = 'select project from wbg_history_data where week_num = %(week_num)s and project = %(project)s'
            # week_num = str(self.WEEK-1)+'周'
            value = {'week_num': str(self.WEEK - 1) + '周', 'project': proj}
            # print(self.WEEK)
            # print(type(self.WEEK))
            cur.execute(sql, value)
            projDate = cur.fetchone()
            if not projDate:
                self.NEW_PROJECT.append(proj)

    # 查询并计算当前周bug总数
    def search_current_week_totalBugNum(self, proj):
        title = "查询当前%s项目当前周bug总数" % proj
        print(title.center(40, '-'))
        connect = self.connect
        self.reConnect()
        cur = connect.cursor()
        sql = 'select year,month,build_bug_num from wbg_year_sta where project = %(proj)s'
        value = {'proj': proj}
        cur.execute(sql, value)
        result = cur.fetchall()
        print(result)
        total_num = 0
        for i in result:
            if i[2]:
                total_num += i[2]
        print('-------', total_num)
        return total_num

    # 查询wbg_history_data某项目所有历史数据，返回成列表，用来绘制首页图表
    def search_allDate_history(self):
        # 查询某项目某周数据
        title = "查询当前所有项目所有历史数据"
        print(title.center(40, '-'))
        connect = self.connect
        self.reConnect()
        cur = connect.cursor()
        sql = 'select project,week_num,bug_leave_num,bug_total_num,bug_leave_rate,bug_add_num,bug_solve_num from wbg_history_data where del_flag is NULL '
        # value = {'proj': 'BIM'}
        cur.execute(sql)
        result = cur.fetchall()
        # print(result)
        resultDict = {}
        for proj in self.PROJ_DICT:
            resultDict[proj] = {'wkn': [], 'bln': [], 'btn': [], 'blr': [], 'ban': [], 'bsn': []}
        for date in result:
            if date[0] in resultDict:
                resultDict[date[0]]['wkn'].append(date[1][0:-1])  # 周数列表
                resultDict[date[0]]['bln'].append(int(date[2]))  # bug遗留数列表
                resultDict[date[0]]['btn'].append(int(date[3]))  # bug总数列表
                resultDict[date[0]]['blr'].append(float(date[4][0:5]) * 100)  # bug遗留率
                resultDict[date[0]]['ban'].append(int(date[5]))  # bug新增数
                resultDict[date[0]]['bsn'].append(int(date[6]))  # bug解决数

        # print(resultDict)
        return resultDict

    # 在wbg_history_data中根据周数查询每周历史数据
    def search_weekDate_history(self, weekNum, proj=None, ):
        # 查询某项目某周数据
        title = "查询当前%s项目%s周历史数据" % (proj, weekNum)
        print(title.center(40, '-'))
        connect = self.connect
        self.reConnect()
        cur = connect.cursor()
        sql = 'select project,week_num,statis_time,bug_leave_num,bug_total_num from wbg_history_data where week_num = %(week_num)s'
        if proj != None:
            sql = 'select project,week_num,statis_time,bug_leave_num,bug_total_num from wbg_history_data where (week_num = %(week_num)s and project=%(proj)s)'
        value = {'week_num': weekNum, 'proj': proj}
        # print(sql)
        cur.execute(sql, value)
        weekData = cur.fetchall()
        # print(weekData)
        print(weekData)
        return weekData

    # 单元格样式（传color则为黄色，不传则无背景色，微软雅黑、实线、垂直居中）
    def style_of_cell(self, color=None):
        """垂直居中,微软雅辉，实线,颜色可选"""
        style = self.workbook.add_format({
            'bold': False,  # 字体加粗
            'border': 1,  # 单元格边框宽度
            'align': 'center',  # 水平对齐方式
            'valign': 'vcenter',  # 垂直对齐方式
            # 'fg_color': color,  # 单元格背景颜色
            'text_wrap': True,  # 是否自动换行
            'font_size': 11,  # 字体
            'font_name': u'微软雅黑'
        })
        if color == 'yellow':
            style = self.workbook.add_format({
                'bold': False,  # 字体加粗
                'border': 1,  # 单元格边框宽度
                'align': 'center',  # 水平对齐方式
                'valign': 'vcenter',  # 垂直对齐方式
                'fg_color': 'yellow',  # 单元格背景颜色
                'text_wrap': True,  # 是否自动换行
                'font_size': 11,  # 字体
                'font_name': u'微软雅黑'
            })
        elif color == 'gray':
            style = self.workbook.add_format({
                'bold': False,  # 字体加粗
                'border': 1,  # 单元格边框宽度
                'align': 'center',  # 水平对齐方式
                'valign': 'vcenter',  # 垂直对齐方式
                'fg_color': '#BEBEBE',  # 单元格背景颜色
                'text_wrap': True,  # 是否自动换行
                'font_size': 11,  # 字体
                'font_name': u'微软雅黑'
            })
        elif color == '14':
            style = self.workbook.add_format({
                'bold': True,  # 字体加粗
                'align': 'left',  # 水平对齐方式
                'valign': 'vcenter',  # 垂直对齐方式
                # 'fg_color': '#BEBEBE',  # 单元格背景颜色
                'text_wrap': False,  # 是否自动换行
                'font_size': 14,  # 字体
                'font_name': u'微软雅黑'
            })
        elif color == 'noBold':
            style = self.workbook.add_format({
                # 'bold': True,  # 字体加粗
                'align': 'left',  # 水平对齐方式
                'valign': 'vcenter',  # 垂直对齐方式
                # 'fg_color': '#BEBEBE',  # 单元格背景颜色
                'text_wrap': True,  # 是否自动换行
                'font_size': 11,  # 字体
                'font_name': u'微软雅黑'
            })
        elif color == 'bold':
            style = self.workbook.add_format({
                'border': 1,  # 单元格边框宽度
                'bold': True,  # 字体加粗
                'align': 'left',  # 水平对齐方式
                'valign': 'vcenter',  # 垂直对齐方式
                # 'fg_color': '#BEBEBE',  # 单元格背景颜色
                'text_wrap': True,  # 是否自动换行
                'font_size': 11,  # 字体
                'font_name': u'微软雅黑',
            })

        return style

    # 绘制遗留数&遗留率统计图表
    def write_index_chart(self):
        # title = "开始绘制《遗留数&遗留率》图表"
        totalDate = self.TOTAL_DATA
        print(totalDate)
        for proj in totalDate:
            title = "开始绘制%s《遗留数&遗留率》图表" % proj
            print(title.center(40, '='))
            matplotlib.rcParams['font.sans-serif'] = ['SimHei']
            matplotlib.rcParams['font.family'] = 'sans-serif'
            matplotlib.rcParams['axes.unicode_minus'] = False
            # 取做图数据
            lx = totalDate[proj]['wkn']  # 横坐标（周数）
            x = range(len(lx))
            y1 = totalDate[proj]['bln']  # 柱状图数据(遗留数）
            y2 = totalDate[proj]['blr']  # 折线图数据（遗留率）
            if len(lx) > 40:  # 优化x坐标
                for i in range(len(lx)):
                    if i % 2 != 0:
                        lx[i] = None
            elif len(lx) == 1:
                lx_t = [None]
                y_t = [0]
                # if proj in self.NEW_PROJECT:
                lx = lx + lx_t
                y1 = y1 + y_t
                y2 = y2 + y_t
                x = range(len(lx))
            # 设置图形大小
            plt.rcParams['figure.figsize'] = (7.0, 2.0)
            fig = plt.figure()
            # 画柱子
            ax1 = fig.add_subplot(111)
            width = 0.01 * len(lx) + 0.05
            print("chart宽度：", width)
            chart1 = ax1.bar(x, y1, alpha=.7, color='#8B0000', width=width, label=u'遗留数')
            # ax1.set_ylabel('遗留数', fontsize='15')
            ax1.set_title(self.PROJ_DICT[proj]['title'], fontsize='10')
            plt.yticks(fontsize=9)
            plt.xticks(x, lx)
            plt.xticks(fontsize=9)
            # ax = plt.gca()
            # 折线图
            ax2 = ax1.twinx()  # 这个很重要噢
            fmt = '%.1f%%'
            yticks = mtick.FormatStrFormatter(fmt)
            ax2.yaxis.set_major_formatter(yticks)
            if proj not in self.NEW_PROJECT:
                chart2 = ax2.plot(x, y2, 'r', color='cornflowerblue', lw='2', mec='r', mfc='w', label=u'遗留率')

            plt.yticks(fontsize=9)
            plt.xticks(x, lx)
            plt.xticks(fontsize=9)
            plt.grid(True)
            # plt.legend()
            ax1.legend(loc=2, fontsize=9)
            if proj not in self.NEW_PROJECT:
                ax2.legend(loc=1, fontsize=9)
            # plt.show()
            # plt.show()
            # 判断是否存在图片存放文件夹
            my_path = os.getcwd()
            my_dir = Path(my_path + '/chart')
            # print(my_dir)

            plt.savefig('%s/%s01' % (str(my_dir), proj), bbox_inches='tight')
            plt.close()
        print("图表1绘制完毕")

    # 绘制新增数&解决数统计图表
    def write_index_chart2(self):
        # title = "开始绘制《新增数&解决数》图表"
        # print(title.center(40, '='))
        totalDate = self.TOTAL_DATA
        for proj in totalDate:
            title = "开始绘制%s《新增数&解决数》图表" % proj
            print(title.center(40, '='))
            matplotlib.rcParams['font.sans-serif'] = ['SimHei']
            matplotlib.rcParams['font.family'] = 'sans-serif'
            matplotlib.rcParams['axes.unicode_minus'] = False
            # 取做图数据
            lx = totalDate[proj]['wkn']  # 横坐标（周数）
            x = range(len(lx))
            y1 = totalDate[proj]['ban']  # 折线图数据(新增数）
            y2 = totalDate[proj]['bsn']  # 折线图数据（解决数）
            if len(lx) > 40:
                for i in range(len(lx)):
                    if i % 2 != 0:
                        lx[i] = None
            # elif len(lx) == 1:
            #     lx_t = [None]
            #     y_t = [0]
            #     # if proj in self.NEW_PROJECT:
            #     lx = lx + lx_t
            #     y1 = y1 + y_t
            #     y2 = y2 + y_t
            #     x = range(len(lx))

            # 设置图形大小
            plt.rcParams['figure.figsize'] = (6.5, 2.0)

            plt.plot(x, y1, 'r', color='red', lw='2', mec='r', mfc='w', label=u'新增数')
            plt.plot(x, y2, 'r', color='green', lw='2', mec='r', mfc='w', label=u'解决数')

            plt.yticks(fontsize=9)
            plt.xticks(x, lx)
            plt.xticks(fontsize=9)
            plt.grid(True)
            #
            plt.legend(loc=2, fontsize=9)
            plt.title(self.PROJ_DICT[proj]['title'], fontsize=10)

            # plt.show()
            # plt.show()
            my_path = os.getcwd()
            my_dir = Path(my_path + '/chart')
            plt.savefig('%s/%s02' % (str(my_dir), proj), bbox_inches='tight')
            plt.close()
        print("图表2绘制完毕")

    # 编写汇总（图表）sheet
    def write_summaryChart_sheet01(self):
        title = "开始编写【汇总（图表）】工作表"
        print(title.center(40, '='))
        self.update_wbg_year_sta()  # 对数据初始化，会更新数据表中的数据
        workbook = self.workbook
        # for sheetName in ['汇总(图表)','汇总','BIM','EBID','OA系统','交通云教育','外业采集系统','考评系统']:
        # 编写汇总（图表）===================================表头1/2行==================================================

        worksheet = workbook.add_worksheet('汇总(图表)')

        worksheet.hide_gridlines(option=2)  # 隐藏网格线
        # 单元格格式为垂直居中，无背景色，字体11
        style = self.style_of_cell()
        style_yellow = self.style_of_cell('yellow')
        style_nobold = self.style_of_cell('nobold')
        # style_small_width = self.style_of_cell()
        worksheet.merge_range(0, 1, 1, 1, '周别', style)
        worksheet.merge_range(0, 2, 1, 2, '统计时间', style)
        for proj in self.PROJ_DICT:
            index = self.PROJ_DICT[proj]['index']
            sheet_name = self.PROJ_DICT[proj]['sheet_name']
            worksheet.merge_range(0, 5 * index - 2, 0, 5 * index + 2, sheet_name, style)
        for i in range(len(self.PROJ_DICT)):
            worksheet.write(1, 5 * i + 3, '遗留数', style)
            worksheet.write(1, 5 * i + 4, '总数', style)
            worksheet.write(1, 5 * i + 5, '遗留率', style_yellow)
            worksheet.write(1, 5 * i + 6, '新增数', style_yellow)
            worksheet.write(1, 5 * i + 7, '解决数', style_yellow)
        for row in range(1, 9):
            worksheet.set_row(row, 20)  # 设置行高
        worksheet.set_column(0, 0, 1)  # 设置列宽
        worksheet.set_column(1, 1, 9)  # 设置列宽
        worksheet.set_column(3, 3 + 5 * len(self.PROJ_DICT), 6)  # 设置列宽
        rowIndex = len(self.PROJ_DICT)
        for row in range(rowIndex):
            worksheet.set_column((row + 1) * 5, (row + 1) * 5, 7)

        week = self.WEEK
        # 最近5周数据展示 ==========================2/3/4/5/6行=========================================================
        # ('BIM', '34周', datetime.datetime(2019, 8, 24, 0, 0), '119', '877')
        # 前4周
        last4weekDate = [self.search_weekDate_history(week - 4), self.search_weekDate_history(week - 3),
                         self.search_weekDate_history(week - 2), self.search_weekDate_history(week - 1)]  # 前四周所有数据
        dataFirstIndex = {}
        for proj in self.PROJ_DICT:
            index = self.PROJ_DICT[proj]['index']
            dataFirstIndex[proj] = 5 * index - 2
        # dataFirstIndex = {'BIM': 3, 'EBID': 8, 'OA': 13, 'EDU': 18, 'FAS': 23, 'EVAL': 28}
        for j in range(0, 4):
            temp_dict = {}
            noData_project_list = []
            k = last4weekDate[j]  # 依次获取前4周某周数据
            for p in k:
                if p[0] in dataFirstIndex:
                    temp_dict[p[0]] = p
            print('1', temp_dict)
            for proj in self.PROJ_DICT:
                if proj not in temp_dict:
                    noData_project_list.append(proj)
            # print(k)
            # worksheet.write(j + 2, 1, week+j-4, style)  # 周别
            # worksheet.write(j + 2, 2, i[2].strftime('%m') + '月' + i[2].strftime('%d') + '日', style)  # 统计时间

            for i in temp_dict:
                i = temp_dict[i]
                print('2', i)
                worksheet.write(j + 2, 1, i[1], style)  # 周别
                worksheet.write(j + 2, 2, i[2].strftime('%m') + '月' + i[2].strftime('%d') + '日', style)  # 统计时间
                # i:('BIM', '48周', datetime.datetime(2019, 11, 30, 14, 19, 38), '130', '1019')
                firstCol = dataFirstIndex[i[0]]
                worksheet.write(j + 2, firstCol, i[3], style)  # 遗留数
                worksheet.write(j + 2, firstCol + 1, i[4], style)  # 总数
                worksheet.write(j + 2, firstCol + 2, format(int(i[3]) / int(i[4]), '.1%'),
                                style_yellow)  # 遗留率（遗留数/总数）
                swh_resilt = self.search_weekDate_history(week - (5 - j), i[0])
                if swh_resilt:
                    worksheet.write(j + 2, firstCol + 3, int(i[4]) - int(swh_resilt[0][4]), style_yellow)  # 新增数(总数-上总数）
                    worksheet.write(j + 2, firstCol + 4,
                                    int(swh_resilt[0][3]) + int(i[4]) - int(i[3]) - int(swh_resilt[0][4]),
                                    style_yellow)  # 解决数（上遗留数+总数-遗留数-上总数）
                else:
                    worksheet.write(j + 2, firstCol + 3, 0, style_yellow)  # 新增数(总数-上总数）
                    worksheet.write(j + 2, firstCol + 4, 0, style_yellow)  # 解决数（上遗留数+总数-遗留数-上总数）

            for i in noData_project_list:
                firstCol = dataFirstIndex[i]
                worksheet.write(j + 2, firstCol, '—', style)  # 遗留数
                worksheet.write(j + 2, firstCol + 1, '—', style)  # 总数
                worksheet.write(j + 2, firstCol + 2, '—', style_yellow)  # 遗留率
                worksheet.write(j + 2, firstCol + 3, '—', style_yellow)  # 新增数
                worksheet.write(j + 2, firstCol + 4, '—', style_yellow)  # 解决数

            for i in self.NEW_PROJECT:
                # new_proj_list.append(i[0])
                firstCol = dataFirstIndex[i]
                worksheet.write(j + 2, firstCol, '—', style)  # 遗留数
                worksheet.write(j + 2, firstCol + 1, '—', style)  # 总数
                worksheet.write(j + 2, firstCol + 2, '—', style_yellow)  # 遗留率（遗留数/总数）
                worksheet.write(j + 2, firstCol + 3, '—', style_yellow)  # 新增数(总数-上总数）
                worksheet.write(j + 2, firstCol + 4, '—', style_yellow)  # 解决数（上遗留数+总数-遗留数-上总数）

        # 本周 =======================================第6行=============================================================
        weekNum = str(week) + '周'
        worksheet.write(6, 1, weekNum, style)  # 周别
        month = datetime.datetime.now().strftime('%m')
        day = datetime.datetime.now().strftime('%d')
        worksheet.write(6, 2, month + '月' + day + '日', style)  # 统计时间
        worksheet.merge_range(7, 1, 9, 1, '最近一周\nbug趋势', style_yellow)
        worksheet.write(7, 2, '解决速率', style_yellow)
        worksheet.write(8, 2, '新增速率', style_yellow)
        worksheet.write(9, 2, '对应状况', style_yellow)
        # currentWeekDict = {'BIM': self.BIM, 'EBID': self.EBID, 'OA': self.OA, 'EDU': self.EDU, 'FAS': self.FAS,
        #                    'EVAL': self.EVAL}

        for proj in self.LEAVE_BUG:
            # print(proj)
            weekBugLeaveNum = int(self.LEAVE_BUG[proj]['lbn'])
            row = dataFirstIndex[proj]
            # print(row)
            worksheet.write(6, row, weekBugLeaveNum, style)  # 遗留数
            weekBugTotalNum = int(self.search_current_week_totalBugNum(proj))
            print('xxxx', self.PROJ_BEFORE_BUG)
            if proj in self.PROJ_BEFORE_BUG:
                weekBugTotalNum = int(self.PROJ_BEFORE_BUG[proj]) + weekBugTotalNum
                worksheet.write(6, row + 1, weekBugTotalNum, style)  # 总数
            else:
                worksheet.write(6, row + 1, weekBugTotalNum, style)  # 总数
            weekBugLeaveRate = str(weekBugLeaveNum / weekBugTotalNum)[0:5]
            worksheet.write(6, row + 2, format(weekBugLeaveNum / weekBugTotalNum, '.1%'), style_yellow)  # 遗留率（遗留数/总数）
            if proj in self.NEW_PROJECT:
                worksheet.write(6, row + 3, '0', style_yellow)  # 新加入的项目bug新增数为空
                worksheet.write(6, row + 4, '0', style_yellow)  # 新增项目的解决数为空
                nowTime = datetime.datetime.now()
                self.update_wbg_history_data(proj, weekNum, nowTime, weekBugLeaveNum, weekBugTotalNum, weekBugLeaveRate,
                                             str('0'), str('0'))
                worksheet.merge_range(7, row, 7, row + 4, '0%', style_yellow)  # 解决速率
                worksheet.merge_range(8, row, 8, row + 4, '0%', style_yellow)  # 新增速率
                worksheet.merge_range(9, row, 9, row + 4, '无应对', style_yellow)  # 对应状况

            else:
                beforeBugTotalNum = int(self.search_weekDate_history(week - 1, proj)[0][4])
                bugAddNum = weekBugTotalNum - beforeBugTotalNum
                worksheet.write(6, row + 3, bugAddNum, style_yellow)  # 新增数(总数-上总数）
                beforeBugLeaveNum = int(
                    self.search_weekDate_history(week - 1, proj)[0][
                        3])  # project,week_num,statis_time,bug_leave_num,bug_total_num
                bugSolveNum = beforeBugLeaveNum + weekBugTotalNum - weekBugLeaveNum - beforeBugTotalNum
                worksheet.write(6, row + 4, bugSolveNum, style_yellow)  # 解决数（上遗留数+总数-遗留数-上总数）

                nowTime = datetime.datetime.now()
                self.update_wbg_history_data(proj, weekNum, nowTime, weekBugLeaveNum, weekBugTotalNum, weekBugLeaveRate,
                                             str(bugAddNum), str(bugSolveNum))

                # 7/8/9行数据写入 ==============================================================================================
                solveRate = bugSolveNum / weekBugTotalNum
                self.currentWeekSolveRate[proj] = format(solveRate, '.1%')
                worksheet.merge_range(7, row, 7, row + 4, format(solveRate, '.1%'), style_yellow)  # 解决速率 row =
                addRate = (weekBugTotalNum - beforeBugTotalNum) / weekBugTotalNum
                self.currentWeekAddRate[proj] = format(addRate, '.1%')
                worksheet.merge_range(8, row, 8, row + 4, format(addRate, '.1%'), style_yellow)  # 新增速率
                if solveRate == 0:
                    worksheet.merge_range(9, row, 9, row + 4, '无应对', style_yellow)  # 对应状况
                    self.correSituation[proj] = '无应对'
                else:
                    if addRate > solveRate:
                        worksheet.merge_range(9, row, 9, row + 4, '对应缓慢', style_yellow)  # 对应状况
                        self.correSituation[proj] = '对应缓慢'
                    else:
                        worksheet.merge_range(9, row, 9, row + 4, '积极应对', style_yellow)  # 对应状况
                        self.correSituation[proj] = '积极应对'

        # 插入图片
        my_path = os.getcwd()
        my_dir = Path(my_path + '/chart')
        if my_dir.is_dir():
            print("存在chart文件夹，删除该文件夹")
            shutil.rmtree(str(my_dir))
            os.mkdir(str(my_dir))
            print("重新创建chart文件夹")
        else:
            print("不存在dir")
            os.mkdir(str(my_dir))
        worksheet.set_row(10, 2)
        self.write_index_chart()  # 绘制图表1 （proj+01）
        self.write_index_chart2()  # 绘制图表2 （proj+02)
        chartIndex = {}
        for proj in self.PROJ_DICT:
            index = self.PROJ_DICT[proj]['index']
            chartIndex[proj + '01'] = 'B' + str(11 * index + 1)
            chartIndex[proj + '02'] = 'N' + str(11 * index + 1)

        # chartIndex = {'BIM01': 'B12', 'EBID01': 'B23', 'OA01': 'B34', 'EDU01': 'B45', 'FAS01': 'B56', 'EVAL01': 'B67',
        #               'BIM02': 'N12', 'EBID02': 'N23', 'OA02': 'N34', 'EDU02': 'N45', 'FAS02': 'N56', 'EVAL02': 'N67'}
        for image in chartIndex:
            worksheet.insert_image(chartIndex[image], '%s\%s.png' % (str(my_dir), image))

        style_3 = self.workbook.add_format({
            # 'bold': True,  # 字体加粗
            'align': 'left',  # 水平对齐方式
            'valign': 'vcenter',  # 垂直对齐方式
            # 'fg_color': '#BEBEBE',  # 单元格背景颜色
            'text_wrap': True,  # 是否自动换行
            'font_size': 11,  # 字体
            'font_name': u'微软雅黑'
        })
        if 'BIM' in self.PROJ_DICT:
            index = self.PROJ_DICT['BIM']['index']  # 一般是1
            worksheet.merge_range(11 * index + 2, 25, 11 * index + 8, 32,
                                  '* 因BIM项目的缺陷管理已由JIRA转移至redmine，故从39周起不再统计JIRA上遗留bug数量', style_3)
        # workbook.close()
        if 'EBID-CCCC' in self.PROJ_DICT:
            index = self.PROJ_DICT['EBID-CCCC']['index']
            worksheet.merge_range(11 * index + 2, 25, 11 * index + 8, 32,
                                  '* 公规院电子招标采购项目2019年12月12号之前bug存于JIRA的EBID项目中，在JIRA中没有单独建立项目，故未统计其在JIRA中的bug，只统计其在Redmine中的bug',
                                  style_3)

        if 'EXPERT_TJ' in self.PROJ_DICT:
            index = self.PROJ_DICT['EXPERT_TJ']['index']
            worksheet.merge_range(11 * index + 2, 25, 11 * index + 8, 32,
                                  '* 天津评标专家管理系统项目2019年12月4号之前bug存于JIRA的EBID项目中，在JIRA中没有单独建立项目，故未统计其在JIRA中的bug，只统计其在Redmine中的bug',
                                  style_3)

        print("-" * 20 + "【汇总（图表）】工作表编写完毕" + "-" * 20)

    # 编写汇总sheet
    def write_summary_sheet02(self):
        title = "开始编写【汇总】工作表"
        print(title.center(40, '='))
        workbook = self.workbook
        worksheet = workbook.add_worksheet('汇总')
        worksheet.hide_gridlines(option=2)  # 隐藏网格线
        style = self.style_of_cell()
        style_lightGray = self.style_of_cell('gray')
        # style_small_width = self.style_of_cell()
        worksheet.set_row(0, 20)  # 设置行高
        worksheet.set_column('A:A', 2)  # 设置列宽
        worksheet.set_column('B:B', 24)  # 设置列宽
        worksheet.merge_range(1, 1, 2, 1, '项目', style)
        projDict = {}
        for proj in self.PROJ_DICT:
            index = self.PROJ_DICT[proj]['index']
            proj_name = self.PROJ_DICT[proj]['sheet_name']
            projDict[proj] = index + 2
            worksheet.write(index + 2, 1, proj_name, style)

        # 查询上周汇报时间
        connect = self.connect
        self.reConnect()
        cur = connect.cursor()
        sql = 'select project,week_num,statis_time,bug_leave_num,bug_total_num,bug_leave_rate,bug_add_num,bug_solve_num from wbg_history_data where week_num = %(week_num)s'
        # week_num = str(self.WEEK-1)+'周'
        value = {'week_num': str(self.WEEK - 1) + '周'}
        # print(self.WEEK)
        # print(type(self.WEEK))
        cur.execute(sql, value)
        totalDate = cur.fetchall()
        beforeWeek = totalDate[0][2]
        beforeWeek = datetime.datetime.strftime(beforeWeek, '%m-%d').replace('-', '月') + '日'
        worksheet.merge_range(1, 2, 1, 6, beforeWeek, style)
        worksheet.write(2, 2, '遗留数', style)
        worksheet.write(2, 3, '总数', style)
        worksheet.write(2, 4, '遗留率', style)
        worksheet.write(2, 5, '新增数', style)
        worksheet.write(2, 6, '解决数', style)
        # projDict = {'BIM': 3, 'EBID': 4, 'OA': 5, 'EDU': 6, 'FAS': 7, 'EVAL': 8}
        for i in totalDate:
            if i[0] in self.PROJ_DICT:
                worksheet.write(projDict[i[0]], 2, i[3], style)
                worksheet.write(projDict[i[0]], 3, i[4], style)
                worksheet.write(projDict[i[0]], 4, format(float(i[5]), '.1%'), style)
                worksheet.write(projDict[i[0]], 5, i[6], style)
                worksheet.write(projDict[i[0]], 6, i[7], style)
        for i in self.NEW_PROJECT:
            worksheet.write(projDict[i], 2, '—', style)
            worksheet.write(projDict[i], 3, '—', style)
            worksheet.write(projDict[i], 4, '—', style)
            worksheet.write(projDict[i], 5, '—', style)
            worksheet.write(projDict[i], 6, '—', style)

        # 查询本周汇报数据
        sql = 'select project,week_num,statis_time,bug_leave_num,bug_total_num,bug_leave_rate,bug_add_num,bug_solve_num from wbg_history_data where week_num = %(week_num)s'
        value = {'week_num': self.WEEK}
        cur.execute(sql, value)
        totalDate = cur.fetchall()
        beforeWeek = totalDate[0][2]
        beforeWeek = datetime.datetime.strftime(beforeWeek, '%m-%d').replace('-', '月') + '日'
        worksheet.merge_range(1, 7, 1, 11, beforeWeek, style)
        worksheet.write(2, 7, '遗留数', style)
        worksheet.write(2, 8, '总数', style)
        worksheet.write(2, 9, '遗留率', style)
        worksheet.write(2, 10, '新增数', style)
        worksheet.write(2, 11, '解决数', style)
        # projDict = {'BIM': 3, 'EBID': 4, 'OA': 5, 'EDU': 6, 'FAS': 7, 'EVAL': 8}
        for i in totalDate:
            if i[0] in self.PROJ_DICT:
                worksheet.write(projDict[i[0]], 7, i[3], style)
                worksheet.write(projDict[i[0]], 8, i[4], style)
                worksheet.write(projDict[i[0]], 9, format(float(i[5]), '.1%'), style)
                worksheet.write(projDict[i[0]], 10, i[6], style)
                worksheet.write(projDict[i[0]], 11, i[7], style)

        # 当周bug整体情况
        worksheet.merge_range(1, 12, 1, 14, '当周bug整体状况', style_lightGray)
        worksheet.write(2, 12, '解决速率', style_lightGray)
        worksheet.write(2, 13, '新增速率', style_lightGray)
        worksheet.write(2, 14, '对应状况', style_lightGray)
        for proj in projDict:
            if proj in self.currentWeekSolveRate:
                worksheet.write(projDict[proj], 12, self.currentWeekSolveRate[proj], style_lightGray)
            else:
                worksheet.write(projDict[proj], 12, '0%', style_lightGray)
            if proj in self.currentWeekSolveRate:
                worksheet.write(projDict[proj], 13, self.currentWeekAddRate[proj], style_lightGray)
            else:
                worksheet.write(projDict[proj], 13, '0%', style_lightGray)
            if proj in self.correSituation:
                worksheet.write(projDict[proj], 14, self.correSituation[proj], style_lightGray)
            else:
                worksheet.write(projDict[proj], 14, '无应对', style_lightGray)

        for i in range(1, 9):
            worksheet.set_row(i, 22)
        print("-" * 20 + "【汇总】工作表编写完毕" + "-" * 20)
        # workbook.close()

    def write_Proj_sheet(self):

        workbook = self.workbook
        connect = self.connect
        self.reConnect()
        cur = connect.cursor()
        PROJ_DICT = self.PROJ_DICT
        PROJ_DICT = sorted(PROJ_DICT.items(), key=lambda PROJ_DICT: PROJ_DICT[1]['index'])
        for proj in PROJ_DICT:
            proj = proj[0]
            title = "开始编写【%s】工作表" % self.PROJ_DICT[proj]['sheet_name']
            print(title.center(40, '='))
            worksheet = workbook.add_worksheet(self.PROJ_DICT[proj]['sheet_name'])
            worksheet.hide_gridlines(option=2)  # 隐藏网格线
            style = self.style_of_cell()
            style_title = self.style_of_cell('14')
            # style_title2 = self.style_of_cell('noBold')
            # style_bold = self.style_of_cell('bold')
            sql = "select year from wbg_year_sta where project=%(proj)s"
            value = {'proj': proj}
            cur.execute(sql, value)
            result = cur.fetchall()
            # print(result)
            yearList = []  # [2017,2018,2019]
            for i in result:
                if int(i[0]) not in yearList:
                    yearList.append(int(i[0]))
            yearList.sort()
            print(yearList)
            worksheet.write('B2', self.PROJ_DICT[proj]['sheet_name'], style_title)
            worksheet.write(3, 1, '年份', style)
            worksheet.write(3, 2, '遗留BUG数', style)
            worksheet.write(3, 3, '新建BUG总数', style)
            worksheet.write(3, 4, '遗留率', style)
            r = len(yearList)
            for index in range(r):
                worksheet.write(4 + index, 1, yearList[index], style)
                sql = "select year,leave_bug_num,build_bug_num from wbg_year_sta where project=%(proj)s and year=%(year)s"
                value = {'year': yearList[index], 'proj': proj}
                cur.execute(sql, value)
                result = cur.fetchall()
                lbn = 0
                nbn = 0
                for i in result:
                    lbn += i[1]
                    nbn += i[2]
                worksheet.write(4 + index, 2, lbn, style)
                if proj in self.PROJ_BEFORE_BUG:
                    nbn += self.PROJ_BEFORE_BUG[proj]
                worksheet.write(4 + index, 3, nbn, style)
                if nbn == 0:
                    worksheet.write(4 + index, 4, 0, style)
                else:
                    worksheet.write(4 + index, 4, format(int(lbn) / int(nbn), '.1%'), style)

            style_1 = self.workbook.add_format({
                'bold': False,  # 字体加粗
                # 'border': 1,  # 单元格边框宽度
                'align': 'left',  # 水平对齐方式
                'valign': 'vcenter',  # 垂直对齐方式
                # 'fg_color': color,  # 单元格背景颜色
                'text_wrap': True,  # 是否自动换行
                'font_size': 11,  # 字体
                'font_name': u'微软雅黑'
            })
            # 本年bug状况
            worksheet.merge_range(5 + r, 1, 5 + r, 3, '%s年BUG状况' % self.YEAR, style_1)
            worksheet.write(6 + r, 1, '月份', style)
            worksheet.write(6 + r, 2, '遗留BUG数', style)
            worksheet.write(6 + r, 3, '新建BUG总数', style)
            worksheet.write(6 + r, 4, '遗留率', style)
            sql = "select month,leave_bug_num,build_bug_num from wbg_year_sta where project=%(proj)s and year=%(year)s"
            value = {'proj': proj, 'year': self.YEAR}
            cur.execute(sql, value)
            result = cur.fetchall()
            # print(result)
            monthdict = {}
            for i in result:
                monthdict[int(i[0])] = (i[1], i[2])
            print(monthdict)
            row = len(monthdict)  # 9/3
            monthdict = sorted(monthdict.items(), key=lambda x: x[0])
            # currentMonth = datetime.datetime.now().month
            index = 0
            for j in monthdict:
                mlbn = j[1][0]
                mbbn = j[1][1]
                month = j[0]
                index += 1
                # # todo
                # if row == currentMonth:
                #     mlbn = monthdict[str(j + 1)][0]
                #     mbbn = monthdict[str(j + 1)][1]
                #     month = j + 1
                # else:
                #     mlbn = monthdict[str(currentMonth - row + 1 + j)][0]
                #     mbbn = monthdict[str(currentMonth - row + 1 + j)][1]
                #     month = currentMonth - row + 1 + j
                worksheet.write(7 + r + index, 1, month, style)
                worksheet.write(7 + r + index, 2, mlbn, style)
                worksheet.write(7 + r + index, 3, mbbn, style)
                if mbbn == 0:
                    worksheet.write(7 + r + index, 4, '0.0%', style)
                else:
                    worksheet.write(7 + r + index, 4, format(int(mlbn) / int(mbbn), '.1%'), style)  # 遗留率

            # bug遗留时效
            worksheet.merge_range(row + 8 + r, 1, row + 9 + r, 1, '姓名', style)
            worksheet.merge_range(row + 8 + r, 2, row + 9 + r, 2, '遗留bug数 ', style)
            worksheet.merge_range(row + 8 + r, 3, row + 8 + r, 6, 'bug遗留时效 ', style)
            worksheet.write(row + 9 + r, 3, '一周', style)
            worksheet.write(row + 9 + r, 4, '一周~二周', style)
            worksheet.write(row + 9 + r, 5, '二周~一月', style)
            worksheet.write(row + 9 + r, 6, '一月以上', style)
            bugData = self.LEAVE_BUG[proj]
            print(bugData)
            del bugData['proj']
            del bugData['lbn']
            del bugData['bugStatusYear']

            r1 = len(bugData)
            k = 0
            total1 = 0
            total2 = 0
            total3 = 0
            total4 = 0
            total5 = 0
            # for k in range(r):  # 列
            for name in bugData:  # 行
                total = bugData[name]['一周'] + bugData[name]['一周~二周'] + bugData[name]['二周~一月'] + bugData[name]['一月以上']
                worksheet.write(row + 10 + r + k, 1, name, style)  # 姓名
                worksheet.write(row + 10 + r + k, 2, total, style)  # 遗留bug数
                worksheet.write(row + 10 + r + k, 3, bugData[name]['一周'], style)  # 一周
                worksheet.write(row + 10 + r + k, 4, bugData[name]['一周~二周'], style)  # 一周~二周
                worksheet.write(row + 10 + r + k, 5, bugData[name]['二周~一月'], style)  # 二周~一月
                worksheet.write(row + 10 + r + k, 6, bugData[name]['一月以上'], style)  # 一月以上
                k += 1
                total1 += total
                total2 += bugData[name]['一周']
                total3 += bugData[name]['一周~二周']
                total4 += bugData[name]['二周~一月']
                total5 += bugData[name]['一月以上']

            style_2 = self.workbook.add_format({
                'bold': True,  # 字体加粗
                'border': 1,  # 单元格边框宽度
                'align': 'center',  # 水平对齐方式
                'valign': 'vcenter',  # 垂直对齐方式
                # 'fg_color': color,  # 单元格背景颜色
                'text_wrap': True,  # 是否自动换行
                'font_size': 11,  # 字体
                'font_name': u'微软雅黑'
            })
            worksheet.write(row + 10 + r + r1, 1, '汇总', style_2)
            worksheet.write(row + 10 + r + r1, 2, total1, style_2)
            worksheet.write(row + 10 + r + r1, 3, total2, style_2)
            worksheet.write(row + 10 + r + r1, 4, total3, style_2)
            worksheet.write(row + 10 + r + r1, 5, total4, style_2)
            worksheet.write(row + 10 + r + r1, 6, total5, style_2)
            if proj == 'BIM':
                worksheet.merge_range(row + 11 + r + r1, 1, row + 13 + r + r1, 6,
                                      "备注：2018年6月之前采用C++进行编码，此版本已经废弃，但Jira上保留C++编码版本Bug记录，本次统计已排除6月前Bug记录.2018年6月份之前的Bug共计41个",
                                      style_1)
            worksheet.set_column(1, 6, 12)
            worksheet.set_column(0, 0, 4)
            print("新增项目列表：", self.NEW_PROJECT)
            print("-" * 20 + "【%s】工作表编写完毕" % proj + "-" * 20)
        self.connect.close()
        workbook.close()


if __name__ == "__main__":
    sta = Statistics()
    sta.data_init()
    sta.write_summaryChart_sheet01()
    sta.write_summary_sheet02()
    # sta.write_BIM_sheet()
    sta.write_Proj_sheet()
    # -------------
    sta.write_index_chart()
    sta.write_index_chart2()
