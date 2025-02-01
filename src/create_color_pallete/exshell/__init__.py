import subprocess
import time


class Exshell():
    """エクシェル
    """


    def __init__(self):
        self._opened_excel_process = None


    @property
    def opened_excel_process(self):
        return self._opened_excel_process


    @opened_excel_process.setter
    def opened_excel_process(self, value):
        self._opened_excel_process = value


    def open_virtual_display(self, excel_application_path, abs_path_to_workbook):
        """仮想ディスプレイを開く
        """
        print(f"""\
🔧　Open virtual display...
""")
        # 外部プロセスを開始する（エクセルを開く）
        self.opened_excel_process = subprocess.Popen([excel_application_path, abs_path_to_workbook])   # Excel が開くことを期待
        time.sleep(1)


    def close_virtual_display(self):
        """仮想ディスプレイを閉じる
        """

        print(f"""\
🔧　Close virtual display...
""")
        # 外部プロセスを終了する（エクセルを閉じる）
        self.opened_excel_process.terminate()
        time.sleep(1)
