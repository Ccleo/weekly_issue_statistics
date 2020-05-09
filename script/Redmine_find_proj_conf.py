from redminelib import Redmine
import datetime
import time
from matplotlib import pyplot as plt

# from pylab import *


# class RedmineStatistics(object):
redmine = Redmine('http://192.168.1.212:7777/redmine', username='redmine', password='redmine')

issue = redmine.issue.get(7273)
# 获取issue下面的所有方法
# func_list = dir(issue.project)
# print(dir(issue.project))
# for i in func_list:
#     if hasattr(issue.project, i):
#         print(i, "-->", getattr(issue.project, i))

# 获取redmine中某项目对应的tracker_id的值
# print("=================================tracker信息=============================")
print("project_name:", issue.project)
print("tracker_id:", issue.tracker.id)
print("tracker_name:", issue.tracker.name)
