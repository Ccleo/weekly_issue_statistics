class Styles():

    def __init__(self, workbook):
        self.workbook = workbook

    # 单元格样式（传color则为黄色，不传则无背景色，微软雅黑、实线、垂直居中）
    def style_of_cell(self, color=None):
        """垂直居中,微软雅辉，实线,颜色可选"""
        style = self.workbook.add_format({
            'bold': False,  # 字体加粗
            'border': 1,  # 单元格边框宽度
            'align': 'center',  # 水平对齐方式
            'valign': 'vcenter',  # 垂直对齐方式
            # 'fg_color': color,  # 单元格背景颜色
            'text_wrap': True,  # 是否自动换行
            'font_size': 11,  # 字体
            'font_name': u'微软雅黑'
        })
        if color == 'yellow':
            style = self.workbook.add_format({
                'bold': False,  # 字体加粗
                'border': 1,  # 单元格边框宽度
                'align': 'center',  # 水平对齐方式
                'valign': 'vcenter',  # 垂直对齐方式
                'fg_color': 'yellow',  # 单元格背景颜色
                'text_wrap': True,  # 是否自动换行
                'font_size': 11,  # 字体
                'font_name': u'微软雅黑'
            })
        elif color == 'gray':
            style = self.workbook.add_format({
                'bold': False,  # 字体加粗
                'border': 1,  # 单元格边框宽度
                'align': 'center',  # 水平对齐方式
                'valign': 'vcenter',  # 垂直对齐方式
                'fg_color': '#BEBEBE',  # 单元格背景颜色
                'text_wrap': True,  # 是否自动换行
                'font_size': 11,  # 字体
                'font_name': u'微软雅黑'
            })
        elif color == '14':
            style = self.workbook.add_format({
                'bold': True,  # 字体加粗
                'align': 'left',  # 水平对齐方式
                'valign': 'vcenter',  # 垂直对齐方式
                # 'fg_color': '#BEBEBE',  # 单元格背景颜色
                'text_wrap': False,  # 是否自动换行
                'font_size': 14,  # 字体
                'font_name': u'微软雅黑'
            })
        elif color == 'noBold':
            style = self.workbook.add_format({
                # 'bold': True,  # 字体加粗
                'align': 'left',  # 水平对齐方式
                'valign': 'vcenter',  # 垂直对齐方式
                # 'fg_color': '#BEBEBE',  # 单元格背景颜色
                'text_wrap': True,  # 是否自动换行
                'font_size': 11,  # 字体
                'font_name': u'微软雅黑'
            })
        elif color == 'bold':
            style = self.workbook.add_format({
                'border': 1,  # 单元格边框宽度
                'bold': True,  # 字体加粗
                'align': 'left',  # 水平对齐方式
                'valign': 'vcenter',  # 垂直对齐方式
                # 'fg_color': '#BEBEBE',  # 单元格背景颜色
                'text_wrap': True,  # 是否自动换行
                'font_size': 11,  # 字体
                'font_name': u'微软雅黑',
            })

        elif color == 1:
            style = self.workbook.add_format({
                'bold': False,  # 字体加粗
                # 'border': 1,  # 单元格边框宽度
                'align': 'left',  # 水平对齐方式
                'valign': 'vcenter',  # 垂直对齐方式
                # 'fg_color': color,  # 单元格背景颜色
                'text_wrap': True,  # 是否自动换行
                'font_size': 11,  # 字体
                'font_name': u'微软雅黑'
            })
        elif color == 3:
            style = self.workbook.add_format({
                # 'bold': True,  # 字体加粗
                'align': 'left',  # 水平对齐方式
                'valign': 'vcenter',  # 垂直对齐方式
                # 'fg_color': '#BEBEBE',  # 单元格背景颜色
                'text_wrap': True,  # 是否自动换行
                'font_size': 11,  # 字体
                'font_name': u'微软雅黑'
            })

        return style

    # def style_1(self):
    #     style = self.workbook.add_format({
    #         'bold': False,  # 字体加粗
    #         # 'border': 1,  # 单元格边框宽度
    #         'align': 'left',  # 水平对齐方式
    #         'valign': 'vcenter',  # 垂直对齐方式
    #         # 'fg_color': color,  # 单元格背景颜色
    #         'text_wrap': True,  # 是否自动换行
    #         'font_size': 11,  # 字体
    #         'font_name': u'微软雅黑'
    #     })
    #     return style_01
    
    # def style_3(self):
    #     style_03 = self.workbook.add_format({
    #         # 'bold': True,  # 字体加粗
    #         'align': 'left',  # 水平对齐方式
    #         'valign': 'vcenter',  # 垂直对齐方式
    #         # 'fg_color': '#BEBEBE',  # 单元格背景颜色
    #         'text_wrap': True,  # 是否自动换行
    #         'font_size': 11,  # 字体
    #         'font_name': u'微软雅黑'
    #     })
    #     return style_03