# Excel Test

## 趣旨

Excel 手順書をやめられない人向けの、Excel手順書テストサンプル。

## 使用方法

- testフォルダに、テスト対象のExcelを配置します。
- test 内の test_excel_evaluator のケースを書き換える、あるいはコピーして "test_"で始まるファイル名でテストファイルを作ります。

## 実行方法

```bash: ExcelTest直下にて
pip install -r requirements.txt
python -m pytest
```

## 実行結果サンプル

以下は、サンプルケースを `pytest -v` で実行(成功込みで表示)とした場合です。

```
$ python -m pytest -v
========================================================================== test session starts ===========================================================================
platform darwin -- Python 3.7.9, pytest-6.2.5, py-1.11.0, pluggy-1.0.0 -- /Users/hoge/ExcelTest/venv/bin/python
cachedir: .pytest_cache
rootdir: /Users/hoge/ExcelTest
collected 6 items                                                                                                                                                        

test/test_workbook.py::TestExcelEvaluator::test_セルの値を読み取る PASSED                                                                                          [ 16%]
test/test_workbook.py::TestExcelEvaluator::test_セルの数式から値として読み取る PASSED                                                                              [ 33%]
test/test_workbook.py::TestExcelEvaluator::test_セルの数式が参照する値を更新することで計算結果が変わる PASSED                                                      [ 50%]
test/test_workbook.py::TestExcelEvaluator::test_セルの数式が参照する値を更新することで計算結果が変わる_文字編 PASSED                                               [ 66%]
test/test_workbook.py::TestExcelEvaluator::test_セルの数式が参照する値を更新することで計算結果が変わる_他のシートの文字編 PASSED                                   [ 83%]
test/test_workbook.py::TestExcelEvaluator::test_セルの数式が参照する値を更新することで計算結果が変わる_vlookup編 PASSED                                            [100%]

=========================================================================== 6 passed in 0.70s ============================================================================
```
