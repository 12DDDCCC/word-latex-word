import importlib.util
import unittest
from pathlib import Path
from unittest.mock import patch


MODULE_PATH = Path(__file__).with_name('tex_to_word.py')
SPEC = importlib.util.spec_from_file_location('tex_to_word_runtime', MODULE_PATH)
MODULE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(MODULE)


class RuntimeResolutionTest(unittest.TestCase):
    def test_find_pandoc_uses_path_first(self):
        with patch.object(MODULE.shutil, 'which', return_value='pandoc-bin'):
            self.assertEqual(MODULE.find_pandoc(), 'pandoc-bin')

    def test_find_pandoc_prefers_visible_user_path_entry(self):
        system_path = 'C:/Program Files/WinGet/Packages/Pandoc/pandoc-3.9.0.2'
        user_path = 'C:/Users/user/AppData/Local/Microsoft/WinGet/Packages/Pandoc/pandoc-3.9.0.2'
        path = MODULE.os.pathsep.join([system_path, user_path])

        def exists(candidate):
            return 'AppData' in str(candidate)

        with patch.object(MODULE.shutil, 'which', return_value=None), \
                patch.dict(MODULE.os.environ, {'PATH': path}), \
                patch.object(Path, 'exists', exists):
            self.assertEqual(MODULE.find_pandoc(), str(Path(user_path) / 'pandoc.exe'))

    def test_find_pandoc_uses_powershell_command_source(self):
        result = MODULE.subprocess.CompletedProcess(
            args=[], returncode=0, stdout='C:/Tools/Pandoc/pandoc.EXE\n')
        with patch.object(MODULE.shutil, 'which', return_value=None), \
                patch.object(MODULE.os, 'name', 'nt'), \
                patch.dict(MODULE.os.environ, {'PATH': 'C:/Windows'}, clear=True), \
                patch.object(MODULE.subprocess, 'run', return_value=result):
            self.assertEqual(MODULE.find_pandoc(), 'C:/Tools/Pandoc/pandoc.EXE')

    def test_find_pandoc_uses_windows_package_discovery(self):
        root = Path('C:/Program Files') / 'WinGet' / 'Packages'
        match = root / 'vendor' / 'pandoc.exe'
        result = MODULE.subprocess.CompletedProcess(args=[], returncode=0, stdout='')
        with patch.object(MODULE.shutil, 'which', return_value=None), \
                patch.object(MODULE.os, 'name', 'nt'), \
                patch.object(MODULE.subprocess, 'run', return_value=result), \
                patch.dict(MODULE.os.environ, {
                    'PATH': 'C:/Windows',
                    'ProgramFiles': 'C:/Program Files',
                }, clear=True), \
                patch.object(Path, 'is_dir', return_value=True), \
                patch.object(Path, 'rglob', return_value=[match]):
            self.assertEqual(MODULE.find_pandoc(), str(match))


if __name__ == '__main__':
    unittest.main()
