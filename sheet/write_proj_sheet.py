from utils.styles import Styles
from utils.db import DB
from utils.db_search import DbSearch
import time


def write_Proj_sheet(workbook, PROJ_DICT, PROJ_BEFORE_BUG, LEAVE_BUG):
    YEAR = time.strftime('%Y', time.localtime(time.time()))
    # MONTH = time.strftime('%m', time.localtime(time.time()))
    # WEEK = DbSearch().week()  # 当前周
    connect = DB().conn()
    PROJ_DICT_sorted = sorted(PROJ_DICT.items(), key=lambda PROJ_DICT: PROJ_DICT[1]['index'])
    for proj in PROJ_DICT_sorted:
        proj = proj[0]
        title = "开始编写【%s】工作表" % PROJ_DICT[proj]['sheet_name']
        print(title.center(40, '='))
        worksheet = workbook.add_worksheet(PROJ_DICT[proj]['sheet_name'])
        worksheet.hide_gridlines(option=2)  # 隐藏网格线
        style = Styles(workbook).style_of_cell()
        style_title = Styles(workbook).style_of_cell('14')
        # style_title2 = style_of_cell('noBold')
        # style_bold = style_of_cell('bold')
        sql = "select year from wbg_year_sta where project=%(proj)s"
        value = {'proj': proj}
        result = DB().search(sql, value, connect)
        # print(result)
        yearList = []  # [2017,2018,2019]
        for i in result:
            if int(i[0]) not in yearList:
                yearList.append(int(i[0]))
        yearList.sort()
        print(yearList)
        worksheet.write('B2', PROJ_DICT[proj]['sheet_name'], style_title)
        worksheet.write(3, 1, '年份', style)
        worksheet.write(3, 2, '遗留BUG数', style)
        worksheet.write(3, 3, '新建BUG总数', style)
        worksheet.write(3, 4, '遗留率', style)
        r = len(yearList)
        for index in range(r):
            worksheet.write(4 + index, 1, yearList[index], style)
            sql = "select year,leave_bug_num,build_bug_num from wbg_year_sta where project=%(proj)s and year=%(year)s"
            value = {'year': yearList[index], 'proj': proj}
            result = DB().search(sql, value, connect)
            lbn = 0
            nbn = 0
            for i in result:
                lbn += i[1]
                nbn += i[2]
            worksheet.write(4 + index, 2, lbn, style)
            if proj in PROJ_BEFORE_BUG:
                nbn += PROJ_BEFORE_BUG[proj]
            worksheet.write(4 + index, 3, nbn, style)
            if nbn == 0:
                worksheet.write(4 + index, 4, 0, style)
            else:
                worksheet.write(4 + index, 4, format(int(lbn) / int(nbn), '.1%'), style)

        style_1 = Styles(workbook).style_of_cell(1)
        # 本年bug状况
        worksheet.merge_range(5 + r, 1, 5 + r, 3, '%s年BUG状况' % YEAR, style_1)
        worksheet.write(6 + r, 1, '月份', style)
        worksheet.write(6 + r, 2, '遗留BUG数', style)
        worksheet.write(6 + r, 3, '新建BUG总数', style)
        worksheet.write(6 + r, 4, '遗留率', style)
        sql = "select month,leave_bug_num,build_bug_num from wbg_year_sta where project=%(proj)s and year=%(year)s"
        value = {'proj': proj, 'year': YEAR}
        result = DB().search(sql, value, connect)
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
        bugData = LEAVE_BUG[proj]
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

        style_2 = workbook.add_format({
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
        # print("新增项目列表：", NEW_PROJECT)
        print("【%s】工作表编写完毕".center(40, '-') % proj)
    connect.close()
