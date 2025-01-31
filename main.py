import openpyxl as xl
import math
import random
import traceback

from openpyxl.styles import PatternFill
from src.create_color_pallete import Color


def main():
    message = """\
作りたい色の数を入力してください。

Example
-------
3

Input
-----
"""
    line = input(message)
    number_of_color = int(line)


    # ワークブックを新規生成
    wb = xl.Workbook()

    # ワークシート
    ws = wb['Sheet']

    color_obj = create_first_color_obj()
    print(f'{color_obj.to_web_safe_color()=}')

    xl_color = color_obj.to_web_safe_color()[1:]

    pattern_fill = PatternFill(
            patternType='solid',
            fgColor=xl_color)

    for row_th in range(2, 2+number_of_color):
        cell = ws[f'B{row_th}']
        cell.fill = pattern_fill

        cell = ws[f'C{row_th}']
        cell.value = xl_color

    wb.save('./temp/hello.xlsx')


def create_first_color_obj():
    """基準となる最初の１色を決めます。
    """

    # とりあえず、モノクロは避けるとします。

    # RGB値の下限をとりあえず決めます
    low = random.randrange(0, 240)
    # RGB値の上限をとりあえず決めます
    high = random.randrange(low, 255)
    # low と high の中間
    middle = math.floor((low + high) / 2)
    
    return Color(low, high, middle)
    #color = Color(0xFF, 0x66, 0x00)


##########################
# MARK: コマンドから実行時
##########################

if __name__ == '__main__':
    try:
        main()

    except Exception as err:
        print(f"""\
おお、残念！　例外が投げられてしまった！
{type(err)=}  {err=}

以下はスタックトレース表示じゃ。
{traceback.format_exc()}
""")