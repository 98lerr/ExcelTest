from xlcalculator import ModelCompiler
from xlcalculator import Evaluator


class Workbook:
    def __init__(self,
                 evaluator: Evaluator,
                 workbook: str = "./test/samplebook.xlsx"):
        self.evaluator = evaluator
        self.workbook_name = workbook

    @staticmethod
    def get_workbook(workbook: str = "./test/samplebook.xlsx"):
        compiler = ModelCompiler()
        new_model = compiler.read_and_parse_archive(workbook)
        evaluator = Evaluator(new_model)
        return Workbook(evaluator=evaluator, workbook=workbook)

    def get_cell_value(self, worksheet_name: str, cell_address: str):
        return self.evaluator.evaluate("{}!{}".format(worksheet_name, cell_address)).value

    def set_cell_value(self, worksheet_name: str, cell_address: str, cell_value: str):
        self.evaluator.set_cell_value("{}!{}".format(worksheet_name, cell_address), cell_value)


