import traceback

from src.create_color_pallete import Color, ToneSystem


def test():
    low = 0
    high = 256
    print(f"""\
Low High Value Phase   R   G   B
--- ---- ----- ----- --- --- ---""")

    for big_value in range(0,100):
        value = big_value / 100
        tone_system = ToneSystem(
                low=low,
                high=high,
                value=value)

        print(f"""\
{low:3} {high:4} {value:<5} {tone_system.get_phase():5} {tone_system.get_red():3} {tone_system.get_green():3} {tone_system.get_blue():3}""")


##########################
# MARK: コマンドから実行時
##########################

if __name__ == '__main__':
    try:
        test()

    except Exception as err:
        print(f"""\
おお、残念！　例外が投げられてしまった！
{type(err)=}  {err=}

以下はスタックトレース表示じゃ。
{traceback.format_exc()}
""")