import openpyxl as xl
import math
import random
import traceback

from openpyxl.styles import PatternFill
from src.create_color_pallete import Color, ToneSystem


MAX_scalar = 255


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
    number_of_color_samples = int(line)


    # ワークブックを新規生成
    wb = xl.Workbook()

    # ワークシート
    ws = wb['Sheet']

    color_obj = create_first_color_obj(
            number_of_color_samples=number_of_color_samples)

    for row_th in range(2, 2 + number_of_color_samples):

        print(f'{color_obj.to_web_safe_color()=}')
        xl_color = color_obj.to_web_safe_color()[1:]
        pattern_fill = PatternFill(
                patternType='solid',
                fgColor=xl_color)

        cell = ws[f'B{row_th}']
        cell.fill = pattern_fill

        cell = ws[f'C{row_th}']
        cell.value = xl_color

        color_obj = create_next_color_obj(
                previous_color=color_obj,
                number_of_color_samples=number_of_color_samples)


    wb.save('./temp/hello.xlsx')


def create_first_color_obj(number_of_color_samples):
    """基準となる最初の１色を決めます。

    Parameters
    ----------
    number_of_color_samples : int
        色の標本数
    """

    # NOTE ウェブ・セーフ・カラーは、暗い色の幅が多めに取られています。 0～255 のうち、 180 ぐらいまで暗い色です。

    # NOTE 色の標本数が多くなると、 low, high は極端にできません。変化の幅が狭まってしまいます。
    # number_of_color_samples は 1 以上の整数とします。
    if number_of_color_samples == 1:
        # １色しか標本がないのなら、基準色は、色バーの全部が使えます
        freedom_qty = MAX_scalar
    else:
        # ２色しか標本がないのなら、基準色は、控えめに色バーの半分だけ使うことにします
        freedom_qty = MAX_scalar // number_of_color_samples

    half_freedom_qty = freedom_qty // 2

    # 基準彩度の下限
    min_base_scalar = half_freedom_qty
    # 基準彩度の上限
    max_base_scalar = MAX_scalar - half_freedom_qty

    # とりあえず基準彩度の中間点は、幅の間でランダムに決めます
    mid_scalar = random.randrange(min_base_scalar, max_base_scalar)

    # 彩度の下限と上限を決めると、だいたい彩度が決まります。
    # グレーに近づけたければ彩度を小さく、ビビッドに近づけたければ彩度を大きくします。
    # 彩度
    # NOTE モノクロに近づくと、標本数が多くなると、色の違いを出しにくいです。
    #saturation = random.randrange(0, freedom_qty)
    saturation = freedom_qty

    # 彩度の下限
    low_scalar = mid_scalar - saturation
    high_scalar = mid_scalar + saturation

    print(f"""\
{freedom_qty=}
{half_freedom_qty=}
{min_base_scalar=}
{max_base_scalar=}
{mid_scalar=}
{saturation=}
{low_scalar=}
{high_scalar=}""")
    
    return Color(low_scalar, high_scalar, mid_scalar)
    #color = Color(0xFF, 0x66, 0x00)


def create_next_color_obj(previous_color, number_of_color_samples):
    """次の色を算出

    初期状態を、以下の通りとする：
        （１）G値、B値が下限スカラーであり、R値は上限スカラーだ。
            つまり赤

    できること：
        （１）B値が下限彩度なら、R値を上限彩度まで増やす
            赤系
        （２）R値が上限彩度なら、G値を下限彩度まで減らす
        （３）G値が下限彩度なら、B値を上限彩度まで増やす
        （４）B値が上限彩度なら、R値を下限彩度まで減らす
        （５）G値が下限彩度なら、B値を上限彩度まで増やす
    """
    return previous_color


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