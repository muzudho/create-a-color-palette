import traceback

from pathlib import Path

from src.create_color_palette.wizards import OutputGradation, PleaseInputBrightness, PleaseInputHue, PleaseInputNumberOfColorsYouWantToCreate, PleaseInputSaturation
from src.exshell import ExshellBuilder


PATH_TO_EXSHELL_CONFIG = './exshell_config.toml'
PATH_TO_CONTENTS = './temp/gradation.xlsx'
MAX_SCALAR = 255


def main():

    exshell_builder = ExshellBuilder(
            abs_path_to_workbook=Path(PATH_TO_CONTENTS).resolve())

    # エクシェル設定ファイル読込
    exshell_builder.load_config(
            abs_path=Path(PATH_TO_EXSHELL_CONFIG).resolve(),
            create_if_not_exists=True)

    # エクシェル設定ファイルが不完全ならチュートリアル開始
    if not exshell_builder.config_is_ok():
        exshell_builder.start_tutorial()

    # エクシェルの生成
    exshell = exshell_builder.build()


    while True:

        # 基準となる色相
        is_error, start_hue = PleaseInputHue.play(
                exshell=exshell)
        
        if is_error:
            continue

        # 色の数
        number_of_color_samples = PleaseInputNumberOfColorsYouWantToCreate.play(
                exshell=exshell)

        # 彩度を入力させる
        saturation = PleaseInputSaturation.play()

        # 明度を入力させる
        brightness = PleaseInputBrightness.play(saturation=saturation)
        
        # グラデーションを出力する
        is_exit = OutputGradation.play(
                number_of_color_samples=number_of_color_samples,
                start_hue=start_hue,
                saturation=saturation,
                brightness=brightness,
                exshell=exshell)

        if is_exit:
            break


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