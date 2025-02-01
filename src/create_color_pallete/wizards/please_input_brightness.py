MAX_SCALAR = 255


class PleaseInputBrightness():
    """明度を入力させる
    """


    def play(saturation):

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

        return brightness
