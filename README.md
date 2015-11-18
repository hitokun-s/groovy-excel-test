Groovy Excel Test
===

##Data of sample Excel file

|id|name|pid|date|
|---|---|---|---|
|1234567890|	�ق�	|1112233|2014/11/14|
|1234567891| �ق�,�ق�	|1112234|2014/11/14|
|1234567892|�ghoge�h    |1112235|2014/11/14|

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
    2015-11-18 16:21:28 INFO  ExcelReader:? - cell index:1, type:CELL_TYPE_STRING, data:�ق�
    2015-11-18 16:21:28 INFO  ExcelReader:? - cell index:2, type:CELL_TYPE_NUMERIC, data:1112233.0as int:1112233
    2015-11-18 16:21:28 INFO  ExcelReader:? - cell index:3, type:CELL_TYPE_NUMERIC, data:41957.0as date:Fri Nov 14 00:00:00 CET 2014
    2015-11-18 16:21:28 INFO  ExcelReader:? - ================<analysis end!
    2015-11-18 16:21:28 INFO  ExcelReader:? - ---------------------------------
    
##Memo

- ���������������t���APOI������́ACell Type:Numeric �ƂȂ��Ă��܂� 
  => DB�J������`���Q�Ƃ��āA���t�Ȃ�getDateCellValue()���ĂԂƂ��A���l�Ȃ�long/int/double�Ō^�ϊ�����Ƃ����K�v����
- ����������ƃf�[�^���_�u���N�H�[�e�[�V�����ň͂܂�Ă��Ă����o����Ə����Ă��܂��H�i�f�[�^��OpenOffice�Ŏ��F���Ă���̂ł��̂��������j
    
    