Groovy Excel Test
===

##Data of sample Excel file

|id|name|pid|date|
|---|---|---|---|
|1234567890|	ほげ	|1112233|2014/11/14|
|1234567891| ほげ,ほげ	|1112234|2014/11/14|
|1234567892|“hoge”    |1112235|2014/11/14|

##Log sample

    2015-11-18 16:21:28 INFO  ExcelReader:? - Let's process the file:test.xls
    2015-11-18 16:21:28 INFO  ExcelReader:? - workbook class:org.apache.poi.hssf.usermodel.HSSFWorkbook
    2015-11-18 16:21:28 INFO  ExcelReader:? - number of sheets:3
    2015-11-18 16:21:28 INFO  ExcelReader:? - ---------------------------------
    2015-11-18 16:21:28 INFO  ExcelReader:? - sheet index:0
    2015-11-18 16:21:28 INFO  ExcelReader:? - sheet name:Sheet1
    2015-11-18 16:21:28 INFO  ExcelReader:? - row count:4
    2015-11-18 16:21:28 INFO  ExcelReader:? - col count:4
    2015-11-18 16:21:28 INFO  ExcelReader:? - cell type analysis>===============
    2015-11-18 16:21:28 INFO  ExcelReader:? - cell index:0, type:CELL_TYPE_NUMERIC, data:1.23456789E9as long:1234567890
    2015-11-18 16:21:28 INFO  ExcelReader:? - cell index:1, type:CELL_TYPE_STRING, data:ほげ
    2015-11-18 16:21:28 INFO  ExcelReader:? - cell index:2, type:CELL_TYPE_NUMERIC, data:1112233.0as int:1112233
    2015-11-18 16:21:28 INFO  ExcelReader:? - cell index:3, type:CELL_TYPE_NUMERIC, data:41957.0as date:Fri Nov 14 00:00:00 CET 2014
    2015-11-18 16:21:28 INFO  ExcelReader:? - ================<analysis end!
    2015-11-18 16:21:28 INFO  ExcelReader:? - ---------------------------------
    
##Memo

- 整数も小数も日付も、POI側からは、Cell Type:Numeric となってしまう 
  => DBカラム定義を参照して、日付ならgetDateCellValue()を呼ぶとか、数値ならlong/int/doubleで型変換するとかが必要そう
- もしかするとデータがダブルクォーテーションで囲まれていても抽出すると消えてしまう？（データはOpenOfficeで視認しているのでそのせいかも）
    
    