import math


class ToneSystem():
    """色相システム

    状態は大きく分けて、以下の６つある：
        （０）［赤～黄相］R値が最高、B値が最低であり、G値は最低 ～ 最高 - 1 のいずれかだ。
                  Low        High
                R |xxxxxxxxxx|
                G |>>>>>>>>> |
                B |          |

        （１）［黄～緑相］G値が最高、B値が最低であり、R値は最高 ～ 最低 + 1 のいずれかだ。
                R | <<<<<<<<<|
                G |xxxxxxxxxx|
                B |          |

        （２）［緑～シアン相］G値が最高であり、R値は最低であり、B値は最低 ～ 最高 - 1 のいずれかだ。
                R |          |
                G |xxxxxxxxxx|
                B |>>>>>>>>> |

        （３）［シアン～青相］B値が最高であり、R値は最低であり、G値は最高 ～ 最低 + 1 のいずれかだ。
                R |          |
                G | <<<<<<<<<|
                B |xxxxxxxxxx|

        （４）［青～マゼンタ相］B値が最高であり、G値は最低であり、R値は最低 ～ 最高 - 1 のいずれかだ。
                R |>>>>>>>>> |
                G |          |
                B |xxxxxxxxxx|

        （５）［マゼンタ～赤相］R値が最高であり、G値は最低であり、B値は最高 ～ 最低 + 1 のいずれかだ。
                R |xxxxxxxxxx|
                G |          |
                B | <<<<<<<<<|
    
    例えば、 PHASE_NUM = 6 、 Low High 間のサイズを saturation とし、
        （０）［赤～黄相］R値が最高、B値が最低で、G値が最低から最高 - 1に達したとき、 phase = 1 / PHASE_NUM とする
                  Low        High
                R |xxxxxxxxxx|      R= 255
                G |   --->   |      G= phase * PHASE_NUM * saturation
                B |          |      B= 0

        （１）［黄～緑相］G値が最高、B値が最低で、R値が最高から最低 + 1に達したとき、 phase = 2 / PHASE_NUM とする
                  Low        High
                R |xxx<---xxx|      R= saturation - (phase - 1/PHASE_NUM) * PHASE_NUM * saturation
                G |xxxxxxxxxx|      G= 255
                B |          |      B= 0

        （２）［緑～シアン相］G値が最高、R値が最低で、B値が最低から最高 - 1に達したとき、 phase = 3 / PHASE_NUM とする
                  Low        High
                R |          |      B= 0
                G |xxxxxxxxxx|      G= 255
                B |   --->   |      B= (phase - 2/PHASE_NUM) * PHASE_NUM * saturation

        （３）［シアン～青相］B値が最高、R値が最低で、G値が最高から最低 + 1に達したとき、 phase = 4 / PHASE_NUM とする
                  Low        High
                R |          |      R= 0
                G |   <---   |      G= saturation - (phase - 3/PHASE_NUM) * PHASE_NUM * saturation
                B |>>>>>>>>>>|      B= 255

        （４）［青～マゼンタ相］B値が最高、G値が最低で、R値が最低から最高 - 1に達したとき、 phase = 5 / PHASE_NUM とする
                  Low        High
                R |   --->   |      R= (phase - 4/PHASE_NUM) * PHASE_NUM * saturation
                G |          |      G= 0
                B |xxxxxxxxxx|      B= 255

        （５）［マゼンタ～赤相］R値が最高、B値が最低で、B値が最高から最低 + 1に達したとき、 phase = 0 / PHASE_NUM とする
                  Low        High
                R |xxxxxxxxxx|      R= 255
                G |          |      G= 0
                B |xxx<---xxx|      B= saturation - (phase - 5/PHASE_NUM) * PHASE_NUM * saturation
    """


    @classmethod
    @property
    def RED_TO_YELLOW(clazz):
        return 0


    @classmethod
    @property
    def YELLOW_TO_GREEN(clazz):
        return 1


    @classmethod
    @property
    def GREEN_TO_CYAN(clazz):
        return 2


    @classmethod
    @property
    def CYAN_TO_BLUE(clazz):
        return 3


    @classmethod
    @property
    def BLUE_TO_MAGENTA(clazz):
        return 4


    @classmethod
    @property
    def MAGENTA_TO_RED(clazz):
        return 5


    @classmethod
    @property
    def PHASE_NUM(clazz):
        return 6


    def __init__(self, low, high, value):
        """
        Parameters
        high : int
            high はその数を含みません。
        """
        self._low = low
        self._high = high
        self._value = value


    @property
    def low(self):
        return self._low


    @property
    def high(self):
        return self._high


    @property
    def saturation(self):
        return self._high - self._low


    @property
    def value(self):
        if self._value < 0 or 1 <= self._value:
            raise ValueError(f"[0,1) である必要があります。 {self._value=}")

        return self._value


    def get_phase(self):
        """色相の６分類を返す
        """
        if self.value < 1 / ToneSystem.PHASE_NUM:
            return ToneSystem.RED_TO_YELLOW

        if self.value < 2 / ToneSystem.PHASE_NUM:
            return ToneSystem.YELLOW_TO_GREEN

        if self.value < 3 / ToneSystem.PHASE_NUM:
            return ToneSystem.GREEN_TO_CYAN

        if self.value < 4 / ToneSystem.PHASE_NUM:
            return ToneSystem.CYAN_TO_BLUE

        if self.value < 5 / ToneSystem.PHASE_NUM:
            return ToneSystem.BLUE_TO_MAGENTA

        return ToneSystem.MAGENTA_TO_RED


    def get_value_in_phase(self):
        phase = self.get_phase()
        return self.value - phase * (1 / ToneSystem.PHASE_NUM)


    def get_red(self):
        phase = self.get_phase()

        if phase == ToneSystem.RED_TO_YELLOW:
            return 255

        if phase == ToneSystem.YELLOW_TO_GREEN:
            value_in_phase = self.get_value_in_phase()
            return self.saturation - math.ceil(value_in_phase * self.saturation)

        if phase == ToneSystem.GREEN_TO_CYAN:
            return 0

        if phase == ToneSystem.CYAN_TO_BLUE:
            return 0

        if phase == ToneSystem.BLUE_TO_MAGENTA:
            value_in_phase = self.get_value_in_phase()
            return math.ceil(value_in_phase * self.saturation)

        return 255


    def get_green(self):
        phase = self.get_phase()

        if phase == ToneSystem.RED_TO_YELLOW:
            value_in_phase = self.get_value_in_phase()
            return math.ceil(value_in_phase * self.saturation)

        if phase == ToneSystem.YELLOW_TO_GREEN:
            return 255

        if phase == ToneSystem.GREEN_TO_CYAN:
            return 255

        if phase == ToneSystem.CYAN_TO_BLUE:
            value_in_phase = self.get_value_in_phase()
            return self.saturation - math.ceil(value_in_phase * self.saturation)

        if phase == ToneSystem.BLUE_TO_MAGENTA:
            return 0

        return 0


    def get_blue(self):
        phase = self.get_phase()

        if phase == ToneSystem.RED_TO_YELLOW:
            return 0

        if phase == ToneSystem.YELLOW_TO_GREEN:
            return 0

        if phase == ToneSystem.GREEN_TO_CYAN:
            value_in_phase = self.get_value_in_phase()
            return math.ceil(value_in_phase * self.saturation)

        if phase == ToneSystem.CYAN_TO_BLUE:
            return 255

        if phase == ToneSystem.BLUE_TO_MAGENTA:
            return 255

        value_in_phase = self.get_value_in_phase()
        return self.saturation - math.ceil(value_in_phase * self.saturation)
