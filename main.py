import openpyxl as xl
import traceback

from openpyxl.styles import PatternFill


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

    pattern_fill = PatternFill(
            patternType='solid',
            fgColor='FF6600')

    cell = ws[f'B2']
    cell.fill = pattern_fill

    wb.save('./temp/hello.xlsx')

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