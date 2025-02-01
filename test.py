import traceback

from src.create_color_palette import Color, ToneSystem


def test():
    low = 0
    high = 256
    print(f"""\
Low High Hue Phase   R   G   B
--- ---- --- ----- --- --- ---""")

    for big_hue in range(0,100):
        hue = big_hue / 100
        tone_system = ToneSystem(
                low=low,
                high=high,
                hue=hue)

        print(f"""\
{low:3} {high:4} {hue:<3} {tone_system.get_phase():5} {tone_system.get_red():3} {tone_system.get_green():3} {tone_system.get_blue():3}""")


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