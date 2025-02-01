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


        return number_of_color_samples
