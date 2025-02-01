MAX_SCALAR = 255


class PleaseInputSaturation():
    """彩度を入力させる
    """


    def play():
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

        return saturation
