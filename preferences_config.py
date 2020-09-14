'''This file contains PreferencesConfig class.
'''


class PreferencesConfig:
    '''PreferencesConfig class contains all configuration used in PreferencesDialog.
    It has methods to read and write config.
    '''

    def __init__(self):
        self.current_bible_format = ''
        self.bible_rootdir = ''
        self.current_bible_version = ''

    def read_config(self, config):
        '''read_config reads all configuration from config class.
        '''
        self.current_bible_format = self._read_one_string(config, 'current_bible_format', '')
        self.bible_rootdir = self._read_one_string(config, 'bible_rootdir', '')
        self.current_bible_version = self._read_one_string(config, 'current_bible_version', '')

    def _read_one_string(self, config, label, default_value):
        value = default_value

        try:
            value = config.Read(label, default_value)
        except ValueError:
            pass

        return value

    def _read_one_integer(self, config, label, default_value):
        value = default_value

        try:
            value = config.ReadInt(label, default_value)
        except ValueError:
            pass

        return value

    def write_config(self, config):
        '''write_config writes all configuration to config class.
        '''
        config.Write('current_bible_format', self.current_bible_format)
        config.Write('bible_rootdir', self.bible_rootdir)
        config.Write('current_bible_version', self.current_bible_version)
