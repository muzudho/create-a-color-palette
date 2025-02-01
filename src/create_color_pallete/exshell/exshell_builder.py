from tomlkit import parse as toml_parse


class ExshellBuilder():


    def __init__(self, abs_path_to_workbook):
        self._abs_path_to_workbook = abs_path_to_workbook
        self._abs_path_to_config = None


    @property
    def abs_path_to_workbook(self):
        return self._abs_path_to_workbook

    @property
    def abs_path_to_config(self):
        return self._abs_path_to_config


    def load_config(self, abs_path):
        """設定ファイル読取
        """
        self._abs_path_to_config = abs_path
        print(f'🔧　read 📄［ {self._abs_path_to_config} ］config file...')
        with open(self._abs_path_to_config, mode='r', encoding='utf-8') as f:
            config_text = f.read()

        self.config_doc_rw = toml_parse(config_text)
