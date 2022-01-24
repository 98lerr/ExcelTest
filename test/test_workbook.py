from workbook import Workbook


class TestExcelEvaluator:
    def test_セルの値を読み取る(self):
        worksheet = Workbook.get_workbook(workbook="./test/samplebook.xlsx")

        actual = worksheet.get_cell_value(worksheet_name="Sheet1", cell_address="A1")
        expected = "計算用入力"
        assert actual == expected

    def test_セルの数式から値として読み取る(self):
        worksheet = Workbook.get_workbook(workbook="./test/samplebook.xlsx")

        actual = worksheet.get_cell_value(worksheet_name="Sheet1", cell_address="C1")
        expected = 100
        assert actual == expected

    def test_セルの数式が参照する値を更新することで計算結果が変わる(self):
        worksheet = Workbook.get_workbook(workbook="./test/samplebook.xlsx")
        worksheet.set_cell_value(worksheet_name="Sheet1", cell_address="C1", cell_value="200")

        actual = worksheet.get_cell_value(worksheet_name="Sheet1", cell_address="C5")
        expected = 220.00000000000003
        assert actual == expected

    def test_セルの数式が参照する値を更新することで計算結果が変わる_文字編(self):
        workbook = Workbook.get_workbook(workbook="./test/samplebook.xlsx")
        workbook.set_cell_value(worksheet_name="Sheet1", cell_address="C2", cell_value="新しい本")

        actual = workbook.get_cell_value(worksheet_name="Sheet1", cell_address="C6")
        expected = "書籍 - 新しい本"
        assert actual == expected

    def test_セルの数式が参照する値を更新することで計算結果が変わる_他のシートの文字編(self):
        workbook = Workbook.get_workbook(workbook="./test/samplebook.xlsx")
        workbook.set_cell_value(worksheet_name="Sheet1", cell_address="C2", cell_value="新しい本")

        actual = workbook.get_cell_value(worksheet_name="other_sheet", cell_address="B1")
        expected = "書籍 - 新しい本"
        assert actual == expected

    def test_セルの数式が参照する値を更新することで計算結果が変わる_vlookup編(self, two_hundred):
        workbook = Workbook.get_workbook(workbook="./test/samplebook.xlsx")
        workbook.set_cell_value(worksheet_name="Sheet1", cell_address="C1", cell_value=200)

        actual = workbook.get_cell_value(worksheet_name="Sheet1", cell_address="C7")
        expected = two_hundred
        assert actual == expected
