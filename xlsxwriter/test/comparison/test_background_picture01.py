# -*- coding: utf-8 -*-
# !/usr/bin/env python
# @Time    : 2021/5/7 5:20 下午
# @Author  : lidong@test.com
# @Site    : 
# @File    : test_background_picture01.py
from ..excel_comparison_test import ExcelComparisonTest
from ...workbook import Workbook


class TestCompareBackgroundPicture(ExcelComparisonTest):
    def setUp(self):
        self.set_filename('watermark.xlsx')

    def test_create_file_and_add_background_picture(self):
        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()
        filename = self.image_dir + "watermark.png"
        # watermark = open(filename, "rb")
        worksheet.write('B1', 0)
        worksheet.write('B2', 0)
        worksheet.write('B3', 0)
        worksheet.write('C1', 0)
        worksheet.write('C2', 0)
        worksheet.write('C3', 0)
        worksheet.insert_image(0, 0, filename, options={"set_background_picture": True})

        # watermark.close()
        workbook.close()

        self.assertExcelEqual()

