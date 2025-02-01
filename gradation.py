import math
import openpyxl as xl
import os
import random
import subprocess
import time
import traceback

from openpyxl.styles import PatternFill
from pathlib import Path

from src.create_color_pallete import Color, ToneSystem
from src.create_color_pallete.wizards import PleaseInputHue, PleaseInputNumberOfColorsYouWantToCreate
from src.create_color_pallete.exshell import Exshell, ExshellBuilder


PATH_TO_EXSHELL_CONFIG = './exshell_config.toml'
PATH_TO_CONTENTS = './temp/gradation.xlsx'
MAX_SCALAR = 255


class Context():
    """現在の作業状態
    """


    def __init__(self):
        self._number_of_color_samples = None


    @property
    def number_of_color_samples(self):
        return self._number_of_color_samples


    @number_of_color_samples.setter
    def number_of_color_samples(self, value):
        self._number_of_color_samples = value


def main():

    exshell_builder = ExshellBuilder(
            abs_path_to_workbook=Path(PATH_TO_CONTENTS).resolve())

    # エクシェル設定ファイル読込
    exshell_builder.load_config(abs_path=Path(PATH_TO_EXSHELL_CONFIG).resolve())
    # エクシェル設定ファイルが不完全ならチュートリアル開始
    if not exshell_builder.config_is_ok():
        exshell_builder.start_tutorial()

    # エクシェルの生成
    exshell = exshell_builder.build()

    # 現在の状態を保持するオブジェクト
    context_rw = Context()

    while True:

        # 基準となる色相
        PleaseInputHue.play(
                exshell=exshell)

        # 色の数
        context_rw.number_of_color_samples = PleaseInputNumberOfColorsYouWantToCreate.play(
                exshell=exshell)

        subroutine(
                context_rw=context_rw,
                exshell=exshell)


def subroutine(context_rw, exshell):

    message = f"""\
Message
-------
彩度を 0 以上 {MAX_SCALAR} 以下の整数で入力してください。

    Guide
    -----
    *   `0` -   0 に近いほどグレー
    * `{MAX_SCALAR:3}` - {MAX_SCALAR:3} に近いほどビビッド

    Example of input
    ----------------
    {MAX_SCALAR*2//3:3}

Input
-----
"""
    line = input(message)
    saturation = int(line)
    print() # 空行


    high_brightness = MAX_SCALAR
    low_brightness = saturation
    mid_brightness = (high_brightness + low_brightness) // 2

    message = f"""\
Message
-------
明度を {low_brightness} 以上 {high_brightness} 以下の整数で入力してください。

    Guide
    -----
    *   `0` - Black
    * `100` - Dark
    * `220` - Bright
    * `{MAX_SCALAR:3}` - White

    Example of input
    ----------------
    {mid_brightness}

Input
-----
"""
    line = input(message)
    brightness = int(line)
    print() # 空行


    # ワークブックを新規生成
    wb = xl.Workbook()

    # ワークシート
    ws = wb['Sheet']

    low, high = create_tone(
            saturation=saturation,
            brightness=brightness)
    
    # 色相 [0.0, 1.0]
    cur_hue = random.uniform(0, 1)
    step_hue = 1 / context_rw.number_of_color_samples
#     print(f"""\
# {step_hue=}""")


    cell = ws[f'A1']
    cell.value = "No"

    cell = ws[f'B1']
    cell.value = "色"

    cell = ws[f'C1']
    cell.value = "ウェブ・セーフ・カラー"

    # デバッグ用情報
    # cell = ws[f'A1']
    # cell.value = "色相"

    # cell = ws[f'B1']
    # cell.value = "色相種類"

    # cell = ws[f'C1']
    # cell.value = "色相内段階"


    for index, row_th in enumerate(range(2, 2 + context_rw.number_of_color_samples)):

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

        # 連番
        cell = ws[f'A{row_th}']
        cell.value = index

        # 色
        cell = ws[f'B{row_th}']
        cell.fill = pattern_fill

        # ウェブ・セーフ・カラー
        cell = ws[f'C{row_th}']
        cell.value = web_safe_color

        # デバッグ情報
        # cell = ws[f'A{row_th}']
        # cell.value = cur_hue

        # cell = ws[f'B{row_th}']
        # cell.value = tone_system.get_phase_name()

        # cell = ws[f'C{row_th}']
        # cell.value = tone_system.get_value_of_hue_in_phase()

        cur_hue += step_hue
        if 1 < cur_hue:
            cur_hue -= 1


    # ワークブック保存
    exshell.save_workbook(wb=wb)


    is_successful = False

    # エクセル開く
    exshell.open_virtual_display()


    message = f"""\
Message
-------
自動的に開いた Excel アプリケーションを閉じたい場合は y を、
そうでない場合は　それ以外を入力してください。

    Example of input
    ----------------
    y

Input
-----
"""
    line = input(message)
    print() # 空行

    if line == 'y':
        # エクセル閉じる
        exshell.close_virtual_display()


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
    low = saturation

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