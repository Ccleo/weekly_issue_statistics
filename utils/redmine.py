from utils.db import DB
import time, datetime
from redminelib import Redmine
from utils.db_search import DbSearch


class RedmineSta():

    def __init__(self):
        self.YEAR = time.strftime('%Y', time.localtime(time.time()))
        self.MONTH = time.strftime('%m', time.localtime(time.time()))
        self.redmine = Redmine('http://192.168.1.212:7777/redmine', username='redmine', password='redmine')

    # redmine当前月的bug新建情况
    def Redmine_build_bug(self, proj, NEW_PROJECT, PROJ_DICT):
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

        if proj in NEW_PROJECT:  # 新建项目，统计所有已创建bug
            issues_list = self.redmine.issue.filter(project_id=PROJ_DICT[proj]['redmine_project_id'],
                                                    status_id="*",
                                                    tracker_id=PROJ_DICT[proj]['redmine_tracker_id'],
                                                    )
            # print(len(issues_list))
            issues_list_before = self.redmine.issue.filter(project_id=PROJ_DICT[proj]['redmine_project_id'],
                                                           status_id="*",
                                                           tracker_id=PROJ_DICT[proj]['redmine_tracker_id'],
                                                           created_on='<=' + str(start_time))

            created_bug_num = len(issues_list)  # 非新建项目，只统计当前月新建bug数量
            print("【Redmine】%s 新建bug数量为:%d " % (proj, created_bug_num))
            proj_before_bug_num = int(len(issues_list_before))
            print("【Redmine】%s %s月份之前新建bug数量为:%d " % (proj, self.MONTH, created_bug_num))
            if proj_before_bug_num > 0:
                connect = DB().conn()
                cur = connect.cursor()
                updateTime = datetime.datetime.now()
                PROJ_BEFORE_BUG = DbSearch().search_proj_before_bug()
                if proj in PROJ_BEFORE_BUG:
                    sql = 'insert into wbg_proj_before_bug (proj,beforeNum,updateTime) values (%(proj)s,%(beforeNum)s,%(updateTime)s)'
                    value = {'proj': proj, 'beforeNum': proj_before_bug_num, 'updateTime': updateTime}
                    # print('------------------------')
                    # print(value)
                    cur.execute(sql, value)
                    connect.commit()
            return created_bug_num

        else:
            issues_list = self.redmine.issue.filter(project_id=PROJ_DICT[proj]['redmine_project_id'],
                                                    status_id="*",
                                                    tracker_id=PROJ_DICT[proj]['redmine_tracker_id'],
                                                    created_on='><' + str(start_time) + '|' + str(end_time))
            # print(len(issues_list))
            created_bug_num = len(issues_list)  # 当前新建bug数量
            print("【Redmine】%s 新建bug数量为:%d " % (proj, created_bug_num))
            return created_bug_num

    # rdemine每次统计时所有bug遗留情况
    def Redmine_leave_bug(self, proj, PROJ_DICT):
        """统计状态为打开的bug以及开发遗留bug情况(某项目）"""
        # 统计开发遗留bug数量和遗留时长
        issues_list = self.redmine.issue.filter(project_id=PROJ_DICT[proj]['redmine_project_id'], status="打开",
                                                tracker_id=PROJ_DICT[proj]['redmine_tracker_id'])
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
