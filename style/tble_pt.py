import xlwt
from xlwt import *



# Оформление таблиц
style_body = XFStyle()
style_top = XFStyle()
style_bottom = XFStyle()
style_table_integr_top = XFStyle()
style_table_integr_body = XFStyle()

#  tyle = easyxf ('align: horizontal left;')

style_top.font.name = 'Arial Cyr'
style_top.font.bold = 2
style_top.borders.left = 5
style_top.borders.right = 5
style_top.borders.top = 5
style_top.borders.bottom = 5
style_top.alignment.wrap = 1
style_top.alignment.horz = 2
style_top.alignment.vert = 1

style_body.font.name = 'Arial Cyr'
#  style.font
#  style.alignment.HORZ_CENTER = 1
style_body.num_format_str = '#0.000'
style_body.borders.left = 1
style_body.borders.right = 1
style_body.borders.top = 1
style_body.borders.bottom = 1

style_bottom.num_format_str = '#0.000'
#  style_bottom.borders.MEDIUM = 10
style_bottom.borders.left = 2
style_bottom.borders.right = 2
style_bottom.borders.top = 2
style_bottom.borders.bottom = 2

style_table_integr_top.num_format_str = '#0.000'
#  style_bottom.borders.MEDIUM = 10
style_table_integr_top.borders.left = 2
style_table_integr_top.borders.right = 2
style_table_integr_top.borders.top = 2
style_table_integr_top.borders.bottom = 2
style_table_integr_top.alignment.wrap = 1
style_table_integr_top.alignment.horz = 2
style_table_integr_top.alignment.vert = 1
