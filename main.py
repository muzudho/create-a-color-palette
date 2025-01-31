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


    message = f"""\
彩度を 0 以上 {MAX_scalar} 以下の整数で入力してください。
0 に近いほどグレー、{MAX_scalar} に近いほどビビッドに近づきます。

Example
-------
180

Input
-----
"""
    line = input(message)
    saturation = int(line)


    message = f"""\
明度を {saturation} 以上 {MAX_scalar} 以下の整数で入力してください。
0 に近いほど黒、{MAX_scalar} に近いほど白に近づきます。

Example
-------
{saturation}

Input
-----
"""
    line = input(message)
    brightness = int(line)


    # ワークブックを新規生成
    wb = xl.Workbook()

    # ワークシート
    ws = wb['Sheet']

    low, high = create_tone(
            saturation=saturation,
            brightness=brightness)
    
    # 色相 [0.0, 1.0]
    cur_hue = random.uniform(0, 1)
    step_hue = 1 / number_of_color_samples
#     print(f"""\
# {step_hue=}""")


    cell = ws[f'A1']
    cell.value = "色相"

    cell = ws[f'B1']
    cell.value = "色相種類"

    cell = ws[f'C1']
    cell.value = "色相内段階"

    cell = ws[f'D1']
    cell.value = "色"

    cell = ws[f'E1']
    cell.value = "コード"


    for row_th in range(2, 2 + number_of_color_samples):

        tone_system = ToneSystem(
                low=low,
                high=high,
                hue=cur_hue)
#         print(f"""\
# {low=}
# {high=}
# {cur_hue=}""")

        color_obj = Color(tone_system.get_red(), tone_system.get_green(), tone_system.get_blue())
#         print(f"""\
# {color_obj.to_web_safe_color()=}""")
    
        web_safe_color = color_obj.to_web_safe_color()
        xl_color = web_safe_color[1:]
        try:
            pattern_fill = PatternFill(
                    patternType='solid',
                    fgColor=xl_color)
        except:
            print(f'{xl_color=}')
            raise

        cell = ws[f'A{row_th}']
        cell.value = cur_hue

        cell = ws[f'B{row_th}']
        cell.value = tone_system.get_phase_name()

        cell = ws[f'C{row_th}']
        cell.value = tone_system.get_value_of_hue_in_phase()

        cell = ws[f'D{row_th}']
        cell.fill = pattern_fill

        cell = ws[f'E{row_th}']
        cell.value = web_safe_color

        cur_hue += step_hue
        if 1 < cur_hue:
            cur_hue -= 1


    file_path_to_write = './temp/hello.xlsx'
    wb.save(file_path_to_write)
    print(f"Please look 📄［{file_path_to_write}］ file.")


def create_tone(saturation, brightness):
    """色調を１つに決めます。

    Parameters
    ----------
    saturation : int
        彩度。[0, 255] の整数
        NOTE モノクロに近づくと、標本数が多くなると、色の違いを出しにくいです。
    brightness : int
        明度
    """

    # NOTE ウェブ・セーフ・カラーは、暗い色の幅が多めに取られています。 0～255 のうち、 180 ぐらいまで暗い色です。
    # NOTE 色の標本数が多くなると、 low, high は極端にできません。変化の幅が狭まってしまいます。

    # 上限
    high = brightness
    # 下限
    low = brightness - saturation

    if 255 < high:
        raise ValueError(f'{high=} Others: {brightness=} {saturation=}')

    if low < 0:
        raise ValueError(f'{low=} Others: {brightness=} {saturation=}')


    return low, high


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