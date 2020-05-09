import time, datetime
from jira import JIRA


class JiraSta():

    def __init__(self):
        self.jira = JIRA({"server": "http://192.168.1.212:8088"}, basic_auth=("lig", "lig"))
        self.YEAR = time.strftime('%Y', time.localtime(time.time()))

    # JIRA当前月的bug新建情况
    def JIRA_build_bug(self, proj,PROJ_DICT):
        begintime = time.strftime('%Y-%m-01', time.localtime(time.time()))
        # print(begintime)
        # t = time.strftime('%d', time.localtime(time.time()))
        endtime = (datetime.datetime.now() + datetime.timedelta(days=1)).strftime("%Y-%m-%d")
        # print(endtime)
        JQL = "project in (%s) AND issuetype = %s AND created >= %s AND created <= %s" % (
            PROJ_DICT[proj]['jira_project_id'], PROJ_DICT[proj]['jira_tracker_id'], begintime, endtime)
        # print(JQL)
        issue_list = self.jira.search_issues(JQL, maxResults=100000)
        # project_name = {"EBID": "EBID", "OA": "OA系统"}
        print("【JIRA】%s 新建bug数量: %d" % (proj, len(issue_list)))
        bug_build_num = len(issue_list)
        return bug_build_num

    # JIRA每次统计时所有bug遗留情况
    def JIRA_leave_bug(self, proj,PROJ_DICT):

        JQL = 'project = %s AND issuetype = %s AND resolution = Unresolved ORDER BY priority DESC, updated DESC' % (
            PROJ_DICT[proj]['jira_project_id'], PROJ_DICT[proj]['jira_tracker_id'])
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
