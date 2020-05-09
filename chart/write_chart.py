from pathlib import Path
import matplotlib.pyplot as plt
import matplotlib
import matplotlib.ticker as mtick
import os


class Chart():
    # 绘制遗留数&遗留率统计图表
    def __init__(self, TOTAL_DATA, PROJ_DICT, NEW_PROJECT):
        self.TOTAL_DATA = TOTAL_DATA
        self.PROJ_DICT = PROJ_DICT
        self.NEW_PROJECT = NEW_PROJECT
        self.my_dir = Path(os.getcwd() + '/chart/chart')

    def write_index_chart(self):
        # title = "开始绘制《遗留数&遗留率》图表"
        TOTAL_DATA = self.TOTAL_DATA
        PROJ_DICT = self.PROJ_DICT
        NEW_PROJECT = self.NEW_PROJECT
        print(TOTAL_DATA)
        for proj in TOTAL_DATA:
            title = "开始绘制%s《遗留数&遗留率》图表" % proj
            print(title.center(40, '='))
            matplotlib.rcParams['font.sans-serif'] = ['SimHei']
            matplotlib.rcParams['font.family'] = 'sans-serif'
            matplotlib.rcParams['axes.unicode_minus'] = False
            # 取做图数据
            lx = TOTAL_DATA[proj]['wkn']  # 横坐标（周数）
            x = range(len(lx))
            y1 = TOTAL_DATA[proj]['bln']  # 柱状图数据(遗留数）
            y2 = TOTAL_DATA[proj]['blr']  # 折线图数据（遗留率）
            if len(lx) > 40:  # 优化x坐标
                for i in range(len(lx)):
                    if i % 2 != 0:
                        lx[i] = None
            elif len(lx) == 1:
                lx_t = [None]
                y_t = [0]
                # if proj in self.NEW_PROJECT:
                lx = lx + lx_t
                y1 = y1 + y_t
                y2 = y2 + y_t
                x = range(len(lx))
            # 设置图形大小
            plt.rcParams['figure.figsize'] = (7.0, 2.0)
            fig = plt.figure()
            # 画柱子
            ax1 = fig.add_subplot(111)
            width = 0.01 * len(lx) + 0.05
            print("chart宽度：", width)
            chart1 = ax1.bar(x, y1, alpha=.7, color='#8B0000', width=width, label=u'遗留数')
            # ax1.set_ylabel('遗留数', fontsize='15')
            ax1.set_title(PROJ_DICT[proj]['title'], fontsize='10')
            plt.yticks(fontsize=9)
            plt.xticks(x, lx)
            plt.xticks(fontsize=9)
            # ax = plt.gca()
            # 折线图
            ax2 = ax1.twinx()  # 这个很重要噢
            fmt = '%.1f%%'
            yticks = mtick.FormatStrFormatter(fmt)
            ax2.yaxis.set_major_formatter(yticks)
            if proj not in NEW_PROJECT:
                chart2 = ax2.plot(x, y2, 'r', color='cornflowerblue', lw='2', mec='r', mfc='w', label=u'遗留率')

            plt.yticks(fontsize=9)
            plt.xticks(x, lx)
            plt.xticks(fontsize=9)
            plt.grid(True)
            # plt.legend()
            ax1.legend(loc=2, fontsize=9)
            if proj not in NEW_PROJECT:
                ax2.legend(loc=1, fontsize=9)
            # plt.show()
            # plt.show()

            # print(my_dir)

            plt.savefig('%s/%s01' % (str(self.my_dir), proj), bbox_inches='tight')
            plt.close()
        print("图表1绘制完毕")

    # 绘制新增数&解决数统计图表
    def write_index_chart2(self):
        # title = "开始绘制《新增数&解决数》图表"
        # print(title.center(40, '='))
        TOTAL_DATA = self.TOTAL_DATA
        for proj in TOTAL_DATA:
            title = "开始绘制%s《新增数&解决数》图表" % proj
            print(title.center(40, '='))
            matplotlib.rcParams['font.sans-serif'] = ['SimHei']
            matplotlib.rcParams['font.family'] = 'sans-serif'
            matplotlib.rcParams['axes.unicode_minus'] = False
            # 取做图数据
            lx = TOTAL_DATA[proj]['wkn']  # 横坐标（周数）
            x = range(len(lx))
            y1 = TOTAL_DATA[proj]['ban']  # 折线图数据(新增数）
            y2 = TOTAL_DATA[proj]['bsn']  # 折线图数据（解决数）
            if len(lx) > 40:
                for i in range(len(lx)):
                    if i % 2 != 0:
                        lx[i] = None
            # elif len(lx) == 1:
            #     lx_t = [None]
            #     y_t = [0]
            #     # if proj in self.NEW_PROJECT:
            #     lx = lx + lx_t
            #     y1 = y1 + y_t
            #     y2 = y2 + y_t
            #     x = range(len(lx))

            # 设置图形大小
            plt.rcParams['figure.figsize'] = (6.5, 2.0)

            plt.plot(x, y1, 'r', color='red', lw='2', mec='r', mfc='w', label=u'新增数')
            plt.plot(x, y2, 'r', color='green', lw='2', mec='r', mfc='w', label=u'解决数')

            plt.yticks(fontsize=9)
            plt.xticks(x, lx)
            plt.xticks(fontsize=9)
            plt.grid(True)
            #
            plt.legend(loc=2, fontsize=9)
            plt.title(self.PROJ_DICT[proj]['title'], fontsize=10)

            # plt.show()
            # plt.show()
            plt.savefig('%s/%s02' % (str(self.my_dir), proj), bbox_inches='tight')
            plt.close()
        print("图表2绘制完毕")
