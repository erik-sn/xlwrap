import os
import unittest

from xlwrap import ExcelManager

DIRECTORY = os.path.dirname(os.path.realpath(__file__)).replace('\\', '/')


class TestExcelManagerUnit(unittest.TestCase):

    def setUp(self):
        self.manager = ExcelManager(os.path.join(DIRECTORY, 'test_excel/read_test.xls'))

    def tearDown(self):
        self.manager.close()

    def test_parse_row_column_from_args(self):
        self.assertEqual(self.manager._ExcelManager__parse_row_column_from_args('A1'), (1, 1))
        self.assertEqual(self.manager._ExcelManager__parse_row_column_from_args('B5'), (5, 2))
        self.assertEqual(self.manager._ExcelManager__parse_row_column_from_args('AA1'), (1, 27))
        self.assertEqual(self.manager._ExcelManager__parse_row_column_from_args('AC55'), (55, 29))

    def test_parse_row_column_from_args_invalid(self):
        with self.assertRaises(ValueError):
            self.manager._ExcelManager__parse_row_column_from_args('50BB')

    def test_parse_row_column_from_args_fails(self):
        with self.assertRaises(ValueError):
            self.assertEqual(self.manager._ExcelManager__parse_row_column_from_args(''), (1, 1))
        with self.assertRaises(ValueError):
            self.assertEqual(self.manager._ExcelManager__parse_row_column_from_args('C-1'), (1, 1))
        with self.assertRaises(ValueError):
            self.assertEqual(self.manager._ExcelManager__parse_row_column_from_args('11'), (1, 1))
        with self.assertRaises(ValueError):
            self.assertEqual(self.manager._ExcelManager__parse_row_column_from_args('CB'), (1, 1))
        with self.assertRaises(ValueError):
            self.assertEqual(self.manager._ExcelManager__parse_row_column_from_args('1C'), (1, 1))

    def test_check_file_extension_fails(self):
        with self.assertRaises(ValueError):
            self.manager._ExcelManager__check_file_extension('test_name')

    def test_check_file_extension_valid(self):
        self.assertIsNone(self.manager._ExcelManager__check_file_extension('test_name.xls'), None)
        self.assertIsNone(self.manager._ExcelManager__check_file_extension('test_name.xlsx'), None)
        self.assertIsNone(self.manager._ExcelManager__check_file_extension('test_name.xlsm'), None)


class TestExcelManagerIntegration(unittest.TestCase):

    sheet = [
        ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['', 'testb2', '', 'testd2', '', '', '', '', '', 'testj2', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'testaa2'],
        ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['', '', 'testc4', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['', 'testb7', 'testc7', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['', 'testb10', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    ]
    col2 = ['', 'testb2', '', '', '', '', 'testb7', '', '', 'testb10']
    test_indexes = [(2, 2), (2, 4), (2, 10), (2, 27), (4, 3), (7, 2), (7, 3), (10, 2)]

    def setUp(self):
        self.xls = ExcelManager(os.path.join(DIRECTORY, 'test_excel', 'read_test.xls'))
        self.xlsx = ExcelManager(os.path.join(DIRECTORY, 'test_excel', 'read_test.xlsx'))
        self.xlsm = ExcelManager(os.path.join(DIRECTORY, 'test_excel', 'read_test.xlsm'))

    def tearDown(self):
        self.xls.close()
        self.xlsx.close()
        self.xlsm.close()
        if os.path.exists(os.path.join(DIRECTORY, 'test_excel', 'write_test.xls')):
            os.remove(os.path.join(DIRECTORY, 'test_excel', 'write_test.xls'))
        if os.path.exists(os.path.join(DIRECTORY, 'test_excel', 'write_test.xlsx')):
            os.remove(os.path.join(DIRECTORY, 'test_excel', 'write_test.xlsx'))
        if os.path.exists(os.path.join(DIRECTORY, 'test_excel', 'write_test.xlsm')):
            os.remove(os.path.join(DIRECTORY, 'test_excel', 'write_test.xlsm'))

    def test_file_not_found(self):
        with self.assertRaises(FileNotFoundError):
            ExcelManager(os.path.join(DIRECTORY, 'test_excel', 'notfound.xls'))

    def test_xls_change_sheet_name(self):
        self.assertEqual(self.xls.sheet.name, 'Sheet1')
        self.xls.change_sheet(name='other_sheet')
        self.assertEqual(self.xls.sheet.name, 'other_sheet')

    def test_xls_change_sheet_index(self):
        self.assertEqual(self.xls.sheet.name, 'Sheet1')
        self.xls.change_sheet(index=2)
        self.assertEqual(self.xls.sheet.name, 'other_sheet')

    def test_xlsx_change_sheet_name(self):
        self.assertEqual(self.xlsx.sheet.title, 'Sheet1')
        self.xlsx.change_sheet(name='other_sheet')
        self.assertEqual(self.xlsx.sheet.title, 'other_sheet')

    def test_xlsx_change_sheet_index(self):
        self.assertEqual(self.xlsx.sheet.title, 'Sheet1')
        self.xlsx.change_sheet(index=2)
        self.assertEqual(self.xlsx.sheet.title, 'other_sheet')

    def test_xlsm_change_sheet_name(self):
        self.assertEqual(self.xlsm.sheet.title, 'Sheet1')
        self.xlsm.change_sheet(name='other_sheet')
        self.assertEqual(self.xlsm.sheet.title, 'other_sheet')

    def test_xlsm_change_sheet_index(self):
        self.assertEqual(self.xlsm.sheet.title, 'Sheet1')
        self.xlsm.change_sheet(index=2)
        self.assertEqual(self.xlsm.sheet.title, 'other_sheet')

    def test_xls_read(self):
        self.assertEqual(self.xls.read(1, 1), '')
        self.assertEqual(self.xls.read(2, 2), 'testb2')
        self.assertEqual(self.xls.read('B2'), 'testb2')
        self.assertEqual(self.xls.read(10, 2), 'testb10')
        self.assertEqual(self.xls.read('B10'), 'testb10')

    def test_xlsx_read(self):
        self.assertEqual(self.xlsx.read(1, 1), '')
        self.assertEqual(self.xlsx.read(2, 2), 'testb2')
        self.assertEqual(self.xlsx.read('B2'), 'testb2')
        self.assertEqual(self.xlsx.read(10, 2), 'testb10')
        self.assertEqual(self.xlsx.read('B10'), 'testb10')

    def test_xlsm_read(self):
        self.assertEqual(self.xlsm.read(1, 1), '')
        self.assertEqual(self.xlsm.read(2, 2), 'testb2')
        self.assertEqual(self.xlsm.read('B2'), 'testb2')
        self.assertEqual(self.xlsm.read(10, 2), 'testb10')
        self.assertEqual(self.xlsm.read('B10'), 'testb10')

    def test_xls_write(self):
        with self.assertRaises(TypeError):
            self.xls.write(2, 2, value='test_write')

    def test_xls_save(self):
        with self.assertRaises(TypeError):
            self.xls.save()

    def test_invalid_cell_target(self):
        with self.assertRaises(ValueError):
            self.xlsx.write(1, 2, 3, value='test')
        with self.assertRaises(ValueError):
            self.xlsx.write('wrong', value='test')
        with self.assertRaises(ValueError):
            self.xlsx.write('a1', value=os.path.join)

    def test_xlsx_write(self):
        self.xlsx.write(2, 2, value='test_write')
        self.xlsx.write('C5', value='test_write')
        self.xlsx.save(os.path.join(DIRECTORY, 'test_excel', 'write_test.xlsx'))

        manager = ExcelManager(os.path.join(DIRECTORY, 'test_excel', 'write_test.xlsx'))
        self.assertEqual(manager.read(2, 2), 'test_write')
        self.assertEqual(manager.read(5, 3), 'test_write')

    def test_xlsm_write(self):
        self.xlsm.write(2, 2, value='test_write')
        self.xlsm.write('C5', value='test_write')
        self.xlsm.save(os.path.join(DIRECTORY, 'test_excel', 'write_test.xlsm'))

        manager = ExcelManager(os.path.join(DIRECTORY, 'test_excel', 'write_test.xlsm'))
        self.assertEqual(manager.read(2, 2), 'test_write')
        self.assertEqual(manager.read(5, 3), 'test_write')

    def test_xlsx_save(self):
        new_file = os.path.join(DIRECTORY, 'test_excel', 'write_test.xlsx')
        self.xlsx.save(new_file)
        self.assertTrue(os.path.exists(new_file))

    def test_xlsm_save(self):
        new_file = os.path.join(DIRECTORY, 'test_excel', 'write_test.xlsm')
        self.xlsm.save(new_file)
        self.assertTrue(os.path.exists(new_file))

    def test_xls_cell(self):
        self.assertEqual(self.xls.cell(2, 2).value, 'testb2')
        self.assertEqual(self.xls.cell(10, 2).value, 'testb10')

    def test_xlsx_cell(self):
        self.assertEqual(self.xlsx.cell(2, 2).value, 'testb2')
        self.assertEqual(self.xlsx.cell(10, 2).value, 'testb10')

    def test_xlsm_cell(self):
        self.assertEqual(self.xlsm.cell(2, 2).value, 'testb2')
        self.assertEqual(self.xlsm.cell(10, 2).value, 'testb10')

    def test_xls_row(self):
        self.assertEqual(self.xls.row(2), self.sheet[1])

    def test_xlsx_row(self):
        self.assertEqual(self.xlsx.row(2), self.sheet[1])

    def test_xlsm_row(self):
        self.assertEqual(self.xlsm.row(2), self.sheet[1])

    def test_xls_col(self):
        self.assertEqual(self.xls.column(2), self.col2)
        self.assertEqual(self.xls.column('B'), self.col2)
        self.assertEqual(self.xls.column('b'), self.col2)

    def test_xlsx_col(self):
        self.assertEqual(self.xlsx.column(2), self.col2)
        self.assertEqual(self.xlsx.column('B'), self.col2)
        self.assertEqual(self.xlsx.column('b'), self.col2)

    def test_xlsm_col(self):
        self.assertEqual(self.xlsm.column(2), self.col2)
        self.assertEqual(self.xlsm.column('B'), self.col2)
        self.assertEqual(self.xlsm.column('b'), self.col2)

    def test_xls_array(self):
        self.assertEqual(self.xls.array(), self.sheet)

    def test_xlsx_array(self):
        self.assertEqual(self.xlsx.array(), self.sheet)

    def test_xlsm_array(self):
        self.assertEqual(self.xlsm.array(), self.sheet)

    def test_xls_search(self):
        self.assertEqual(self.xls.search('testb2'), (2, 2))
        self.assertEqual(self.xls.search('test', contains=True, match=2), (2, 4))
        self.assertEqual(self.xls.search('does_not_exist'), (None, None))

    def test_xlsx_search(self):
        self.assertEqual(self.xlsx.search('testb2'), (2, 2))
        self.assertEqual(self.xlsx.search('test', contains=True, match=2), (2, 4))
        self.assertEqual(self.xlsx.search('does_not_exist'), (None, None))

    def test_xlsm_search(self):
        self.assertEqual(self.xlsm.search('testb2'), (2, 2))
        self.assertEqual(self.xlsm.search('test', contains=True, match=2), (2, 4))
        self.assertEqual(self.xlsm.search('does_not_exist'), (None, None))

    def test_xls_searches(self):
        self.assertEqual(self.xls.search('testb2', many=True), [(2, 2)])
        self.assertEqual(self.xls.search('test', contains=True, many=True), self.test_indexes)
        self.assertEqual(self.xls.search('does_not_exist', many=True), [])

    def test_xlsx_searches(self):
        self.assertEqual(self.xlsx.search('testb2', many=True), [(2, 2)])
        self.assertEqual(self.xlsx.search('test', contains=True, many=True), self.test_indexes)
        self.assertEqual(self.xlsx.search('does_not_exist', many=True), [])

    def test_xlsm_searches(self):
        self.assertEqual(self.xlsm.search('testb2', many=True), [(2, 2)])
        self.assertEqual(self.xlsm.search('test', contains=True, many=True), self.test_indexes)
        self.assertEqual(self.xlsm.search('does_not_exist', many=True), [])

    def test_xls_info(self):
        info = self.xls.info()
        self.assertEqual(info['file'], os.path.join(DIRECTORY, 'test_excel', 'read_test.xls'))
        self.assertEqual(info['sheet'], 'Sheet1')
        self.assertEqual(info['reads'], 0)
        self.assertEqual(info['writes'], 0)

        info_str = 'File: {}\nSheet: {}\nReads: {}\nWrites: {}' \
            .format(info['file'], info['sheet'], info['reads'], info['writes'])
        self.assertEqual(self.xls.info(string=True), info_str)

    def test_xlsx_info(self):
        info = self.xlsx.info()
        self.assertEqual(info['file'], os.path.join(DIRECTORY, 'test_excel', 'read_test.xlsx'))
        self.assertEqual(info['sheet'], 'Sheet1')
        self.assertEqual(info['reads'], 0)
        self.assertEqual(info['writes'], 0)

        info_str = 'File: {}\nSheet: {}\nReads: {}\nWrites: {}' \
            .format(info['file'], info['sheet'], info['reads'], info['writes'])
        self.assertEqual(self.xlsx.info(string=True), info_str)

    def test_xlsm_info(self):
        info = self.xlsm.info()
        self.assertEqual(info['file'], os.path.join(DIRECTORY, 'test_excel', 'read_test.xlsm'))
        self.assertEqual(info['sheet'], 'Sheet1')
        self.assertEqual(info['reads'], 0)
        self.assertEqual(info['writes'], 0)

        info_str = 'File: {}\nSheet: {}\nReads: {}\nWrites: {}' \
            .format(info['file'], info['sheet'], info['reads'], info['writes'])
        self.assertEqual(self.xlsm.info(string=True), info_str)
