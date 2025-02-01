import openpyxl as xl
import subprocess
import time

from openpyxl.styles import PatternFill


class PleaseInputNumberOfColorsYouWantToCreate():


    def play(exshell):
        message = """\
🙋　Please input
-----------------
作りたい色の数を 1 以上、常識的な数以下の整数で入力してください。

    Guide
    -----
    *   `3` - ３色
    * `100` - １００色

    Example of input
    ----------------
    7

Input
-----
"""
        line = input(message)
        number_of_color_samples = int(line)
        print() # 空行


#         # ワークブックを新規生成
#         wb = xl.Workbook()

#         # ワークシート
#         ws = wb['Sheet']

#         cell = ws[f'A1']
#         cell.value = "No"

#         cell = ws[f'B1']
#         cell.value = "色"


#         for index, row_th in enumerate(range(2, 2 + number_of_color_samples)):

#             web_safe_color = '#FFFFFF'
#             xl_color = web_safe_color[1:]
#             try:
#                 pattern_fill = PatternFill(
#                         patternType='solid',
#                         fgColor=xl_color)
#             except:
#                 print(f'{xl_color=}')
#                 raise

#             # 連番
#             cell = ws[f'A{row_th}']
#             cell.value = index

#             # 色
#             cell = ws[f'B{row_th}']
#             cell.fill = pattern_fill

#             # コメント
#             cell = ws[f'C{row_th}']
#             cell.value = '未定'


#         # ワークブック保存
#         exshell.save_workbook(wb=wb)

#         # エクセルを開く
#         exshell.open_virtual_display()


#         message = f"""\
# 🙋　Please input
# -----------------
# 開いたワークシートは、サンプルです。
# 次に進むために、こちらに何も文字を入力せずエンターキーを入力してください。

#     Example of input
#     ----------------
    

# Input
# -----
# """
#         line = input(message)
#         print() # 空行


#         # エクセルを閉じる
#         exshell.close_virtual_display()


        return number_of_color_samples
