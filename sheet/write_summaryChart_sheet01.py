from utils.db_update import DbUpdate
from utils.styles import Styles
from utils.db_search import DbSearch
from chart.write_chart import Chart
import datetime
import os
import shutil
from pathlib import Path


# 编写汇总（图表）sheet
def write_summaryChart_sheet01(workbook, NEW_BUG, LEAVE_BUG, TOTAL_DATA, PROJ_DICT, NEW_PROJECT, PROJ_BEFORE_BUG,
                               currentWeekSolveRate, currentWeekAddRate, correSituation, WEEK):
    title = "开始编写【汇总（图表）】工作表"
    print(title.center(40, '='))
    DbUpdate().update_wbg_year_sta(NEW_BUG, LEAVE_BUG)  # 对数据初始化，会更新数据表中的数据
    # for sheetName in ['汇总(图表)','汇总','BIM','EBID','OA系统','交通云教育','外业采集系统','考评系统']:
    # 编写汇总（图表）===================================表头1/2行==================================================

    worksheet = workbook.add_worksheet('汇总(图表)')

    worksheet.hide_gridlines(option=2)  # 隐藏网格线
    # 单元格格式为垂直居中，无背景色，字体11
    style = Styles(workbook).style_of_cell()
    style_yellow = Styles(workbook).style_of_cell('yellow')
    style_nobold = Styles(workbook).style_of_cell('nobold')
    # style_small_width = Styles(workbook).style_of_cell()
    worksheet.merge_range(0, 1, 1, 1, '周别', style)
    worksheet.merge_range(0, 2, 1, 2, '统计时间', style)
    for proj in PROJ_DICT:
        index = PROJ_DICT[proj]['index']
        sheet_name = PROJ_DICT[proj]['sheet_name']
        worksheet.merge_range(0, 5 * index - 2, 0, 5 * index + 2, sheet_name, style)
    for i in range(len(PROJ_DICT)):
        worksheet.write(1, 5 * i + 3, '遗留数', style)
        worksheet.write(1, 5 * i + 4, '总数', style)
        worksheet.write(1, 5 * i + 5, '遗留率', style_yellow)
        worksheet.write(1, 5 * i + 6, '新增数', style_yellow)
        worksheet.write(1, 5 * i + 7, '解决数', style_yellow)
    for row in range(1, 9):
        worksheet.set_row(row, 20)  # 设置行高
    worksheet.set_column(0, 0, 1)  # 设置列宽
    worksheet.set_column(1, 1, 9)  # 设置列宽
    worksheet.set_column(3, 3 + 5 * len(PROJ_DICT), 6)  # 设置列宽
    rowIndex = len(PROJ_DICT)
    for row in range(rowIndex):
        worksheet.set_column((row + 1) * 5, (row + 1) * 5, 7)

    # 最近5周数据展示 ==========================2/3/4/5/6行=========================================================
    # ('BIM', '34周', datetime.datetime(2019, 8, 24, 0, 0), '119', '877')
    # 前4周
    last4weekDate = [DbSearch().search_weekDate_history(WEEK - 4), DbSearch().search_weekDate_history(WEEK - 3),
                     DbSearch().search_weekDate_history(WEEK - 2),
                     DbSearch().search_weekDate_history(WEEK - 1)]  # 前四周所有数据
    dataFirstIndex = {}
    for proj in PROJ_DICT:
        index = PROJ_DICT[proj]['index']
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
        for proj in PROJ_DICT:
            if proj not in temp_dict:
                noData_project_list.append(proj)
        # print(k)
        # worksheet.write(j + 2, 1, WEEK+j-4, style)  # 周别
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
            swh_resilt = DbSearch().search_weekDate_history(WEEK - (5 - j), i[0])
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

        for i in NEW_PROJECT:
            # new_proj_list.append(i[0])
            firstCol = dataFirstIndex[i]
            worksheet.write(j + 2, firstCol, '—', style)  # 遗留数
            worksheet.write(j + 2, firstCol + 1, '—', style)  # 总数
            worksheet.write(j + 2, firstCol + 2, '—', style_yellow)  # 遗留率（遗留数/总数）
            worksheet.write(j + 2, firstCol + 3, '—', style_yellow)  # 新增数(总数-上总数）
            worksheet.write(j + 2, firstCol + 4, '—', style_yellow)  # 解决数（上遗留数+总数-遗留数-上总数）

    # 本周 =======================================第6行=============================================================
    weekNum = str(WEEK) + '周'
    worksheet.write(6, 1, weekNum, style)  # 周别
    month = datetime.datetime.now().strftime('%m')
    day = datetime.datetime.now().strftime('%d')
    worksheet.write(6, 2, month + '月' + day + '日', style)  # 统计时间
    worksheet.merge_range(7, 1, 9, 1, '最近一周\nbug趋势', style_yellow)
    worksheet.write(7, 2, '解决速率', style_yellow)
    worksheet.write(8, 2, '新增速率', style_yellow)
    worksheet.write(9, 2, '对应状况', style_yellow)
    # currentWeekDict = {'BIM': BIM, 'EBID': EBID, 'OA': OA, 'EDU': EDU, 'FAS': FAS,
    #                    'EVAL': EVAL}

    for proj in LEAVE_BUG:
        # print(proj)
        weekBugLeaveNum = int(LEAVE_BUG[proj]['lbn'])
        row = dataFirstIndex[proj]
        # print(row)
        worksheet.write(6, row, weekBugLeaveNum, style)  # 遗留数
        weekBugTotalNum = int(DbSearch().search_current_week_totalBugNum(proj))
        print('xxxx', PROJ_BEFORE_BUG)
        if proj in PROJ_BEFORE_BUG:
            weekBugTotalNum = int(PROJ_BEFORE_BUG[proj]) + weekBugTotalNum
            worksheet.write(6, row + 1, weekBugTotalNum, style)  # 总数
        else:
            worksheet.write(6, row + 1, weekBugTotalNum, style)  # 总数
        weekBugLeaveRate = str(weekBugLeaveNum / weekBugTotalNum)[0:5]
        worksheet.write(6, row + 2, format(weekBugLeaveNum / weekBugTotalNum, '.1%'), style_yellow)  # 遗留率（遗留数/总数）
        if proj in NEW_PROJECT:
            worksheet.write(6, row + 3, '0', style_yellow)  # 新加入的项目bug新增数为空
            worksheet.write(6, row + 4, '0', style_yellow)  # 新增项目的解决数为空
            nowTime = datetime.datetime.now()
            DbUpdate().update_wbg_history_data(proj, weekNum, nowTime, weekBugLeaveNum, weekBugTotalNum,
                                               weekBugLeaveRate,
                                               str('0'), str('0'))
            worksheet.merge_range(7, row, 7, row + 4, '0%', style_yellow)  # 解决速率
            worksheet.merge_range(8, row, 8, row + 4, '0%', style_yellow)  # 新增速率
            worksheet.merge_range(9, row, 9, row + 4, '无应对', style_yellow)  # 对应状况

        else:
            beforeBugTotalNum = int(DbSearch().search_weekDate_history(WEEK - 1, proj)[0][4])
            bugAddNum = weekBugTotalNum - beforeBugTotalNum
            worksheet.write(6, row + 3, bugAddNum, style_yellow)  # 新增数(总数-上总数）
            beforeBugLeaveNum = int(
                DbSearch().search_weekDate_history(WEEK - 1, proj)[0][
                    3])  # project,week_num,statis_time,bug_leave_num,bug_total_num
            bugSolveNum = beforeBugLeaveNum + weekBugTotalNum - weekBugLeaveNum - beforeBugTotalNum
            worksheet.write(6, row + 4, bugSolveNum, style_yellow)  # 解决数（上遗留数+总数-遗留数-上总数）

            nowTime = datetime.datetime.now()
            DbUpdate().update_wbg_history_data(proj, weekNum, nowTime, weekBugLeaveNum, weekBugTotalNum,
                                               weekBugLeaveRate,
                                               str(bugAddNum), str(bugSolveNum))

            # 7/8/9行数据写入 ==============================================================================================
            solveRate = bugSolveNum / weekBugTotalNum
            currentWeekSolveRate[proj] = format(solveRate, '.1%')
            worksheet.merge_range(7, row, 7, row + 4, format(solveRate, '.1%'), style_yellow)  # 解决速率 row =
            addRate = (weekBugTotalNum - beforeBugTotalNum) / weekBugTotalNum
            currentWeekAddRate[proj] = format(addRate, '.1%')
            worksheet.merge_range(8, row, 8, row + 4, format(addRate, '.1%'), style_yellow)  # 新增速率
            if solveRate == 0:
                worksheet.merge_range(9, row, 9, row + 4, '无应对', style_yellow)  # 对应状况
                correSituation[proj] = '无应对'
            else:
                if addRate > solveRate:
                    worksheet.merge_range(9, row, 9, row + 4, '对应缓慢', style_yellow)  # 对应状况
                    correSituation[proj] = '对应缓慢'
                else:
                    worksheet.merge_range(9, row, 9, row + 4, '积极应对', style_yellow)  # 对应状况
                    correSituation[proj] = '积极应对'

    # 插入图片
    my_path = os.getcwd()
    my_dir = Path(my_path + '/chart/chart')
    if my_dir.is_dir():
        print("存在chart文件夹，删除该文件夹")
        shutil.rmtree(str(my_dir))
        os.mkdir(str(my_dir))
        print("重新创建chart文件夹")
    else:
        print("不存在dir")
        os.mkdir(str(my_dir))
    worksheet.set_row(10, 2)
    Chart(TOTAL_DATA,PROJ_DICT,NEW_PROJECT).write_index_chart()  # 绘制图表1 （proj+01）
    Chart(TOTAL_DATA,PROJ_DICT,NEW_PROJECT).write_index_chart2()  # 绘制图表2 （proj+02)
    chartIndex = {}
    for proj in PROJ_DICT:
        index = PROJ_DICT[proj]['index']
        chartIndex[proj + '01'] = 'B' + str(11 * index + 1)
        chartIndex[proj + '02'] = 'N' + str(11 * index + 1)

    # chartIndex = {'BIM01': 'B12', 'EBID01': 'B23', 'OA01': 'B34', 'EDU01': 'B45', 'FAS01': 'B56', 'EVAL01': 'B67',
    #               'BIM02': 'N12', 'EBID02': 'N23', 'OA02': 'N34', 'EDU02': 'N45', 'FAS02': 'N56', 'EVAL02': 'N67'}
    for image in chartIndex:
        worksheet.insert_image(chartIndex[image], '%s\%s.png' % (str(my_dir), image))

    style_3 = Styles(workbook).style_of_cell(3)
    if 'BIM' in PROJ_DICT:
        index = PROJ_DICT['BIM']['index']  # 一般是1
        worksheet.merge_range(11 * index + 2, 25, 11 * index + 8, 32,
                              '* 因BIM项目的缺陷管理已由JIRA转移至redmine，故从39周起不再统计JIRA上遗留bug数量', style_3)
    # workbook.close()
    if 'EBID-CCCC' in PROJ_DICT:
        index = PROJ_DICT['EBID-CCCC']['index']
        worksheet.merge_range(11 * index + 2, 25, 11 * index + 8, 32,
                              '* 公规院电子招标采购项目2019年12月12号之前bug存于JIRA的EBID项目中，在JIRA中没有单独建立项目，故未统计其在JIRA中的bug，只统计其在Redmine中的bug',
                              style_3)

    if 'EXPERT_TJ' in PROJ_DICT:
        index = PROJ_DICT['EXPERT_TJ']['index']
        worksheet.merge_range(11 * index + 2, 25, 11 * index + 8, 32,
                              '* 天津评标专家管理系统项目2019年12月4号之前bug存于JIRA的EBID项目中，在JIRA中没有单独建立项目，故未统计其在JIRA中的bug，只统计其在Redmine中的bug',
                              style_3)

    print("【汇总（图表）】工作表编写完毕".center(40, '-'))
    return currentWeekSolveRate, currentWeekAddRate, correSituation
