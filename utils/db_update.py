import datetime,time
from utils.db import DB
from utils.db_search import DbSearch

class DbUpdate():

    def __init__(self):
        self.YEAR = time.strftime('%Y', time.localtime(time.time()))
        self.MONTH = time.strftime('%m', time.localtime(time.time()))
        self.connect = DB().conn()
        self.WEEK = DbSearch().week()


    def update_wbg_year_sta(self,NEW_BUG,LEAVE_BUG):
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
        cur = connect.cursor()
        updateTime = datetime.datetime.now()
        # updateTime = '2019-10-08 20:28:26'
        for proj in NEW_BUG:
            # 更新当前年份的之前月份（不用更新build_bug_num），直接更新
            # if self.YEAR in self.NEW_BUG[proj]['bugStatusYear']:
            print(proj)
            # print(self.NEW_BUG)
            for i in range(1, currentMonth):
                sql = 'update wbg_year_sta set leave_bug_num=%(lbn)s,update_time=%(ut)s where year=%(currentYear)s and month=%(currentMonth)s and project=%(proj)s'
                value = {'lbn': LEAVE_BUG[proj]['bugStatusYear'][self.YEAR][str(i)],
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
                    value = {'bbn': NEW_BUG[proj],
                             'lbn': LEAVE_BUG[proj]['bugStatusYear'][self.YEAR][str(currentMonth)],
                             'currentYear': currentYear,
                             'currentMonth': currentMonth, 'proj': proj, 'ut': updateTime}
                    cur.execute(sql, value)
                    connect.commit()
                else:
                    sql = 'insert into wbg_year_sta (project,year,month,leave_bug_num,build_bug_num,update_time) values (%(proj)s,%(year)s,%(month)s,%(lbn)s,%(bbn)s,%(ut)s)'
                    value = {'proj': proj, 'year': currentYear, 'month': currentMonth,
                             'lbn': LEAVE_BUG[proj]['bugStatusYear'][self.YEAR][str(currentMonth)],
                             'bbn': NEW_BUG[proj], 'ut': updateTime}
                    # print('------------------------')
                    # print(value)
                    cur.execute(sql, value)

            # 更新之前年份
            before_year_dict = LEAVE_BUG[proj]['bugStatusYear']
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

    def update_wbg_history_data(self, proj, wn, st, bln, btn, blr, ban, bsn):
        connect = self.connect
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
