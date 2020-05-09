from utils.db import DB


class DbSearch():

    def __init__(self):
        print('connect')
        self.connect = DB().conn()

    def search_proj_before_bug(self):
        PROJ_BEFORE_BUG = {}
        title = '查询项目纳入统计之前的新建bug数'
        print(title.center(40, '-'))
        cur = self.connect.cursor()
        sql = 'select * from wbg_proj_before_bug'
        cur.execute(sql)
        result = cur.fetchall()
        print(result)
        for i in result:
            PROJ_BEFORE_BUG[i[0]] = i[1]
        print("项目纳入统计之前的新建bug数:" ,PROJ_BEFORE_BUG)
        return PROJ_BEFORE_BUG

    # 统计当前周数
    def week(self):
        cur = self.connect.cursor()
        sql = 'select last_week from wbg_week where id = 0'
        cur.execute(sql)
        week = cur.fetchall()[0][0]
        print("当前周数：", week)
        # print(type(week))
        self.connect.close()
        return week

    # 查询是否为新增项目(查询wbg_history_data表中的上周数据是否有该proj的值)
    def search_new_project(self,PROJ_DICT,WEEK,NEW_PROJECT):
        title = "查询新增项目"
        print(title.center(40, '-'))
        cur = self.connect.cursor()
        for proj in PROJ_DICT:
            sql = 'select project from wbg_history_data where week_num = %(week_num)s and project = %(project)s'
            # week_num = str(self.WEEK-1)+'周'
            value = {'week_num': str(WEEK - 1) + '周', 'project': proj}
            # print(self.WEEK)
            # print(type(self.WEEK))
            cur.execute(sql, value)
            projDate = cur.fetchone()
            if not projDate:
                NEW_PROJECT.append(proj)
        print('新增项目查询结束'.center(40, '-'))
        print("新增项目列表:", NEW_PROJECT)
        return NEW_PROJECT

    # 查询并计算当前周bug总数
    def search_current_week_totalBugNum(self, proj):
        title = "查询当前%s项目当前周bug总数" % proj
        print(title.center(40, '-'))
        cur = self.connect.cursor()
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
    def search_allDate_history(self,PROJ_DICT):
        # 查询某项目某周数据
        title = "查询当前所有项目所有历史数据"
        print(title.center(40, '-'))
        cur = self.connect.cursor()
        sql = 'select project,week_num,bug_leave_num,bug_total_num,bug_leave_rate,bug_add_num,bug_solve_num from wbg_history_data where del_flag is NULL '
        # value = {'proj': 'BIM'}
        cur.execute(sql)
        result = cur.fetchall()
        # print(result)
        resultDict = {}
        for proj in PROJ_DICT:
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
        cur = self.connect.cursor()
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
