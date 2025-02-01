MAX_SCALAR = 255


class PleaseInputBrightness():
    """明度を入力させる

    明度は high として使われる
    """


    def play(saturation):

        # high は上限まで使用可能
        high_brightness = MAX_SCALAR

        # low は彩度以上が必要
        low_brightness = saturation

        mid_brightness = (high_brightness + low_brightness) // 2

        message = f"""\
🙋　Please input
-----------------
明度を {low_brightness} 以上 {high_brightness} 以下の整数で入力してください。

    Guide
    -----
    *   `0` - Black out
    * `100` - Dark
    * `220` - Bright
    * `{MAX_SCALAR:3}` - White out

    Example of input
    ----------------
    {mid_brightness}

Input
-----
"""
        line = input(message)
        brightness = int(line)
        print() # 空行

        return brightness
