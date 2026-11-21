import tempfile
import unittest
from pathlib import Path
from office_tool import choose_app, validate_file, restore_regions, is_setup_dialog, write_update_policy

class SafetyTests(unittest.TestCase):
    def test_update_policy_uses_real_office_16_registry_path(self):
        import office_tool
        self.assertIn('16.0', office_tool.POLICY_KEY.split(chr(92)))
        self.assertFalse(any(ord(character) < 32 for character in office_tool.POLICY_KEY))

    def test_open_with_filename_opens_without_tool_window(self):
        import office_tool
        from unittest.mock import patch
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'جدول من الزر الأيمن.xlsx'
            path.write_bytes(b'fixture')
            result = {'status': 'opened_in_existing_session'}
            with patch.object(office_tool.sys, 'argv', ['OfficeFileTool.exe', str(path)]), patch.object(office_tool, 'gui') as window, patch.object(office_tool, 'start_trial', return_value=result) as opener:
                self.assertEqual(office_tool.main(), result)
                opener.assert_called_once_with('Excel', str(path.resolve()))
                window.assert_not_called()

    def test_launch_without_filename_keeps_tool_window(self):
        import office_tool
        from unittest.mock import patch
        with patch.object(office_tool.sys, 'argv', ['OfficeFileTool.exe']), patch.object(office_tool, 'gui') as window, patch.object(office_tool, 'start_trial') as opener:
            office_tool.main()
            window.assert_called_once_with('Word', '')
            opener.assert_not_called()

    def test_only_office_setup_dialog_matches(self):
        self.assertTrue(is_setup_dialog('NUIDialog', 'تسجيل الدخول لإعداد Office'))
        self.assertTrue(is_setup_dialog('NUIDialog', 'Sign in to set up Office'))
        self.assertFalse(is_setup_dialog('OpusApp', 'تسجيل الدخول لإعداد Office'))
        self.assertFalse(is_setup_dialog('NUIDialog', 'Save changes to Office document?'))
        self.assertFalse(is_setup_dialog('#32770', 'Sign in to set up Office'))

    def test_arbitrary_unicode_path(self):
        with tempfile.TemporaryDirectory() as folder:
            p = Path(folder) / 'اختبار بأي اسم.DOCX'
            p.write_bytes(b'test')
            self.assertEqual(validate_file(str(p), 'Word'), p.resolve())

    def test_update_buttons_only_change_registry(self):
        import winreg
        import uuid
        base = 'Software\\OfficeFileToolTest_' + uuid.uuid4().hex
        try:
            for enabled in [False, True]:
                write_update_policy(enabled, hive=winreg.HKEY_CURRENT_USER, policy_key=base+'\\Policy', config_key=base+'\\Config')
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, base+'\\Policy') as key:
                    self.assertEqual(winreg.QueryValueEx(key, 'enableautomaticupdates')[0], int(enabled))
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, base+'\\Config') as key:
                    self.assertEqual(winreg.QueryValueEx(key, 'UpdatesEnabled')[0], 'True' if enabled else 'False')
        finally:
            for suffix in ['Policy','Config']:
                try:winreg.DeleteKey(winreg.HKEY_CURRENT_USER, base+'\\'+suffix)
                except FileNotFoundError:pass
            try:winreg.DeleteKey(winreg.HKEY_CURRENT_USER, base)
            except FileNotFoundError:pass

    def test_frozen_launcher_does_not_need_source(self):
        import office_tool
        from unittest.mock import patch
        with patch.object(office_tool.sys, 'frozen', True, create=True), patch.object(office_tool.sys, 'executable', r'C:\Portable\OfficeFileTool.exe'):
            self.assertEqual(office_tool.self_command('start', '--app', 'Excel'), [r'C:\Portable\OfficeFileTool.exe', 'start', '--app', 'Excel'])

    def test_installation_configuration_matches_supported_build(self):
        import office_tool
        import xml.etree.ElementTree as ET
        root = ET.fromstring(office_tool.installation_xml())
        self.assertEqual(root.find('Add').get('Version'), '16.0.19628.20132')
        self.assertEqual(root.find('Add').get('OfficeClientEdition'), '32')
        self.assertEqual(root.find('Updates').get('Enabled'), 'FALSE')

    def test_other_office_types(self):
        for file, app in [('جدول.XLSX','Excel'), ('book.xlsm','Excel'), ('data.csv','Excel'), ('db.accdb','Access'), ('old.mdb','Access'), ('design.pub','Publisher'), ('note.one','OneNote'), ('message.msg','Outlook')]:
            self.assertEqual(choose_app(file), app)

    def test_powerpoint_detection(self):
        self.assertEqual(choose_app('slides.PPTX'), 'PowerPoint')
        self.assertEqual(choose_app('old.ppt'), 'PowerPoint')

    def test_mismatch_rejected(self):
        with tempfile.TemporaryDirectory() as folder:
            p = Path(folder) / 'slides.pptx'
            p.write_bytes(b'test')
            with self.assertRaises(ValueError):
                validate_file(str(p), 'Word')

    def test_restore_refuses_foreign_bytes_before_writes(self):
        regions = [{'address': 1, 'original': '55', 'replacement': 'c3'},
                   {'address': 2, 'original': '55', 'replacement': 'c3'}]
        memory = {1: b'\xc3', 2: b'\x90'}
        with self.assertRaises(RuntimeError):
            restore_regions(regions, lambda a,n: memory[a], lambda a,b: memory.update({a:b}))
        self.assertEqual(memory[1], b'\xc3')

    def test_restore_is_repeatable(self):
        regions = [{'address': 1, 'original': '55', 'replacement': 'c3'}]
        memory = {1: b'\xc3'}
        for _ in range(2):
            restore_regions(regions, lambda a,n: memory[a], lambda a,b: memory.update({a:b}))
        self.assertEqual(memory[1], b'\x55')

if __name__ == '__main__':
    unittest.main()
