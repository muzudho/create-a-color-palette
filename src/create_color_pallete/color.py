class Color():


    def __init__(self, red, green, blue):
        self._red = red
        self._green = green
        self._blue = blue


    @property
    def red(self):
        return self._red


    @property
    def green(self):
        return self._green


    @property
    def blue(self):
        return self._blue


    def to_web_safe_color(self):
        r = format(self._red, '02x')
        g = format(self._green, '02x')
        b = format(self._blue, '02x')
        return f'#{r}{g}{b}'
