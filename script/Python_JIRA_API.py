from redminelib import Redmine
import datetime
import time
from jira import JIRA
import pymysql
import matplotlib.pyplot as plt
import matplotlib
import matplotlib.ticker as mtick
import xlsxwriter


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


class Statistics(object):
    redmine = Redmine('http://192.168.1.212:7777/redmine', username='redmine', password='redmine')
    # options = {"server": "http://192.168.1.212:8088"}
    # auth = ("lig", "lig")  # user_name:test user_passwd:test
    jira = JIRA({"server": "http://192.168.1.212:8088"}, basic_auth=("lig", "lig"))
    PROJ_DICT = {
        'BIM': {'type': 'redmine', 'redmine_project_id': 'fjbim', 'redmine_tracker_id': '7', 'title': '福建BIM项目测试Bug状况',
                'index': 1, 'sheet_name': '福建BIM项目'},
        # 'EBID': {'type': 'jira', 'jira_project_id': 'EBID', 'jira_tracker_id': '缺陷', 'title': '电子招投标项目测试Bug状况',
        #          'index': 2, 'sheet_name': '电子招投标系统'},
        'OA': {'type': 'jira', 'jira_project_id': 'OA', 'jira_tracker_id': '故障', 'title': '华杰OA系统项目测试Bug状况',
               'index': 2, 'sheet_name': '华杰OA系统'},
        'EBID-CCCC': {'type': 'redmine', 'redmine_project_id': 'ebid-cccc', 'redmine_tracker_id': '7',
                      'title': '公规院电子招标采购系统项目测试Bug状况', 'index': 3, 'sheet_name': '公规院电子招标采购项目'},
        'FAS': {'type': 'redmine', 'redmine_project_id': 'hsdfas', 'redmine_tracker_id': '7',
                'title': '勘察设计外业采集系统项目测试Bug状况',
                'index': 4, 'sheet_name': '勘察设计外业采集系统'},
        'EVAL': {'type': 'redmine', 'redmine_project_id': 'eval-defectmanagement', 'redmine_tracker_id': '15',
                 'title': '考评系统项目测试Bug状况',
                 'index': 5, 'sheet_name': '考评系统'},  # 考评系统下面还有子视图
        'EXPERT_TJ': {'type': 'redmine', 'redmine_project_id': 'expert_tj', 'redmine_tracker_id': '7',
                      'title': '天津评标专家管理系统项目测试Bug状况', 'index': 6, 'sheet_name': '天津评标专家管理系统项目'},
        # 'PMS': {'type': 'redmine', 'redmine_project_id': 'pms', 'redmine_tracker_id': '15',
        #         'title': '国际工程项目信息管理系统项目测试Bug状况', 'index': 7, 'sheet_name': '国际工程项目信息管理系统'}}
        'ECAP': {'type': 'redmine', 'redmine_project_id': 'ecap', 'redmine_tracker_id': '7',
                 'title': '应急资源管理平台项目测试Bug状况', 'index': 7, 'sheet_name': '应急资源管理平台'}}

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
        issue = issue_list[1]
        fl = dir(issue)
        for i in fl:
            if hasattr(issue, i):
                print(i, "-->", getattr(issue, i))
        # project_name = {"EBID": "EBID", "OA": "OA系统"}

    # JIRA每次统计时所有bug遗留情况
    def JIRA_leave_bug(self, proj):

        JQL = 'project = %s AND issuetype = %s AND resolution = Unresolved ORDER BY priority DESC, updated DESC' % (
            self.PROJ_DICT[proj]['jira_project_id'], self.PROJ_DICT[proj]['jira_tracker_id'])
        issues_list = self.jira.search_issues(JQL, maxResults=10000)
        leave_bug = len(issues_list)

    # 类属性更新，一次从数据库获取关键数据

    #     pass


if __name__ == "__main__":
    sta = Statistics()
    sta.JIRA_build_bug('OA')

a = {'self': 'http://192.168.1.212:8088/rest/api/2/issue/23140', 'key': 'OA-4479', 'fields': {
    'votes': {'hasVoted': False, 'self': 'http://192.168.1.212:8088/rest/api/2/issue/OA-4479/votes', 'votes': 0},
    'components': [{'self': 'http://192.168.1.212:8088/rest/api/2/component/10200', 'id': '10200', 'name': '财务管理'}],
    'customfield_10005': '0|i01xqn:', 'aggregateprogress': {'total': 0, 'progress': 0}, 'issuelinks': [],
    'aggregatetimeestimate': None, 'customfield_10101': None, 'timeoriginalestimate': None, 'resolution': None,
    'creator': {'displayName': '李炜', 'active': True,
                'avatarUrls': {'24x24': 'http://192.168.1.212:8088/secure/useravatar?size=small&avatarId=10346',
                               '48x48': 'http://192.168.1.212:8088/secure/useravatar?avatarId=10346',
                               '16x16': 'http://192.168.1.212:8088/secure/useravatar?size=xsmall&avatarId=10346',
                               '32x32': 'http://192.168.1.212:8088/secure/useravatar?size=medium&avatarId=10346'},
                'emailAddress': 'liwei@huajie.com.cn',
                'self': 'http://192.168.1.212:8088/rest/api/2/user?username=liwei', 'timeZone': 'Asia/Shanghai',
                'name': 'liwei', 'key': 'liwei'}, 'customfield_10100': None, 'created': '2020-01-09T16:50:33.000+0800',
    'aggregatetimespent': None, 'updated': '2020-01-09T16:50:33.000+0800',
    'priority': {'self': 'http://192.168.1.212:8088/rest/api/2/priority/5', 'id': '5', 'name': 'Lowest',
                 'iconUrl': 'http://192.168.1.212:8088/images/icons/priorities/lowest.svg'}, 'lastViewed': None,
    'watches': {'self': 'http://192.168.1.212:8088/rest/api/2/issue/OA-4479/watchers', 'isWatching': False,
                'watchCount': 1}, 'customfield_10202': None, 'workratio': -1, 'customfield_10200': None,
    'issuetype': {'id': '10100', 'avatarId': 10303,
                  'iconUrl': 'http://192.168.1.212:8088/secure/viewavatar?size=xsmall&avatarId=10303&avatarType=issuetype',
                  'self': 'http://192.168.1.212:8088/rest/api/2/issuetype/10100',
                  'description': '测试过程，维护过程发现影响系统运行的问题。', 'name': '故障', 'subtask': False},
    'progress': {'total': 0, 'progress': 0}, 'labels': [], 'duedate': None, 'versions': [],
    'aggregatetimeoriginalestimate': None, 'customfield_10204': None, 'assignee': {'displayName': '王光宇', 'active': True,
                                                                                   'avatarUrls': {
                                                                                       '24x24': 'http://192.168.1.212:8088/secure/useravatar?size=small&avatarId=10341',
                                                                                       '48x48': 'http://192.168.1.212:8088/secure/useravatar?avatarId=10341',
                                                                                       '16x16': 'http://192.168.1.212:8088/secure/useravatar?size=xsmall&avatarId=10341',
                                                                                       '32x32': 'http://192.168.1.212:8088/secure/useravatar?size=medium&avatarId=10341'},
                                                                                   'emailAddress': 'wangguangyu@huajie.com.cn',
                                                                                   'self': 'http://192.168.1.212:8088/rest/api/2/user?username=wanggy',
                                                                                   'timeZone': 'Asia/Shanghai',
                                                                                   'name': 'wanggy', 'key': 'wanggy'},
    'description': '【复现步骤】\r\n # 费用分摊中有数据，支付内容中删除该数据\r\n\r\n【预期结果】\r\n * 费用分摊中有数据，支付内容中删除该记录需二次确认\r\n\r\n【实际结果】\r\n * 费用分摊中有数据，支付内容中可以直接删除该记录无提示',
    'summary': '【报销管理】【日常费用报销】费用分摊中有数据，支付内容中删除该记录需二次确认', 'timespent': None, 'resolutiondate': None,
    'project': {'self': 'http://192.168.1.212:8088/rest/api/2/project/10300', 'name': '华杰OA综合管理系统', 'id': '10300',
                'avatarUrls': {'24x24': 'http://192.168.1.212:8088/secure/projectavatar?size=small&avatarId=10324',
                               '48x48': 'http://192.168.1.212:8088/secure/projectavatar?avatarId=10324',
                               '16x16': 'http://192.168.1.212:8088/secure/projectavatar?size=xsmall&avatarId=10324',
                               '32x32': 'http://192.168.1.212:8088/secure/projectavatar?size=medium&avatarId=10324'},
                'key': 'OA'}, 'timeestimate': None, 'environment': None, 'subtasks': [],
    'status': {'id': '10000', 'iconUrl': 'http://192.168.1.212:8088/',
               'self': 'http://192.168.1.212:8088/rest/api/2/status/10000', 'description': '', 'name': '待办',
               'statusCategory': {'colorName': 'blue-gray',
                                  'self': 'http://192.168.1.212:8088/rest/api/2/statuscategory/2', 'name': '待办',
                                  'id': 2, 'key': 'new'}}, 'customfield_10201': None, 'customfield_10203': None,
    'customfield_10000': None, 'fixVersions': [], 'reporter': {'displayName': '李炜', 'active': True, 'avatarUrls': {
        '24x24': 'http://192.168.1.212:8088/secure/useravatar?size=small&avatarId=10346',
        '48x48': 'http://192.168.1.212:8088/secure/useravatar?avatarId=10346',
        '16x16': 'http://192.168.1.212:8088/secure/useravatar?size=xsmall&avatarId=10346',
        '32x32': 'http://192.168.1.212:8088/secure/useravatar?size=medium&avatarId=10346'},
                                                               'emailAddress': 'liwei@huajie.com.cn',
                                                               'self': 'http://192.168.1.212:8088/rest/api/2/user?username=liwei',
                                                               'timeZone': 'Asia/Shanghai', 'name': 'liwei',
                                                               'key': 'liwei'}, 'customfield_10004': None},
     'id': '23140', 'expand': 'operations,versionedRepresentations,editmeta,changelog,renderedFields'}
