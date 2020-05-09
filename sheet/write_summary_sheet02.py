from utils.styles import Styles
import datetime
from utils.db import DB

    

def write_summary_sheet02(workbook,PROJ_DICT,WEEK,NEW_PROJECT,currentWeekSolveRate,currentWeekAddRate,correSituation):
    title = "开始编写【汇总】工作表"
    print(title.center(40, '='))
    worksheet = workbook.add_worksheet('汇总')
    worksheet.hide_gridlines(option=2)  # 隐藏网格线
    style = Styles(workbook).style_of_cell()
    style_lightGray = Styles(workbook).style_of_cell('gray')
    # style_small_width = Styles(workbook).style_of_cell()
    worksheet.set_row(0, 20)  # 设置行高
    worksheet.set_column('A:A', 2)  # 设置列宽
    worksheet.set_column('B:B', 24)  # 设置列宽
    worksheet.merge_range(1, 1, 2, 1, '项目', style)
    projDict = {}
    for proj in PROJ_DICT:
        index = PROJ_DICT[proj]['index']
        proj_name = PROJ_DICT[proj]['sheet_name']
        projDict[proj] = index + 2
        worksheet.write(index + 2, 1, proj_name, style)

    # 查询上周汇报时间
    connect = DB().conn()
    sql = 'select project,week_num,statis_time,bug_leave_num,bug_total_num,bug_leave_rate,bug_add_num,bug_solve_num from wbg_history_data where week_num = %(week_num)s'
    # week_num = str(WEEK-1)+'周'
    value = {'week_num': str(WEEK - 1) + '周'}
    # print(WEEK)
    # print(type(WEEK))
    totalDate = DB().search(sql,value,connect)
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
        if i[0] in PROJ_DICT:
            worksheet.write(projDict[i[0]], 2, i[3], style)
            worksheet.write(projDict[i[0]], 3, i[4], style)
            worksheet.write(projDict[i[0]], 4, format(float(i[5]), '.1%'), style)
            worksheet.write(projDict[i[0]], 5, i[6], style)
            worksheet.write(projDict[i[0]], 6, i[7], style)
    for i in NEW_PROJECT:
        worksheet.write(projDict[i], 2, '—', style)
        worksheet.write(projDict[i], 3, '—', style)
        worksheet.write(projDict[i], 4, '—', style)
        worksheet.write(projDict[i], 5, '—', style)
        worksheet.write(projDict[i], 6, '—', style)

    # 查询本周汇报数据
    sql = 'select project,week_num,statis_time,bug_leave_num,bug_total_num,bug_leave_rate,bug_add_num,bug_solve_num from wbg_history_data where week_num = %(week_num)s'
    value = {'week_num': WEEK}
    totalDate = DB().search(sql,value,connect)
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
        if i[0] in PROJ_DICT:
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
        if proj in currentWeekSolveRate:
            worksheet.write(projDict[proj], 12, currentWeekSolveRate[proj], style_lightGray)
        else:
            worksheet.write(projDict[proj], 12, '0%', style_lightGray)
        if proj in currentWeekSolveRate:
            worksheet.write(projDict[proj], 13, currentWeekAddRate[proj], style_lightGray)
        else:
            worksheet.write(projDict[proj], 13, '0%', style_lightGray)
        if proj in correSituation:
            worksheet.write(projDict[proj], 14, correSituation[proj], style_lightGray)
        else:
            worksheet.write(projDict[proj], 14, '无应对', style_lightGray)

    for i in range(1, 9):
        worksheet.set_row(i, 22)
    print("【汇总】工作表编写完毕".center(40, '-'))
    # workbook.close()