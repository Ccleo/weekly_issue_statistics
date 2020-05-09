import os
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

root_dir = os.getcwd()
print('root_dir:',os.getcwd())