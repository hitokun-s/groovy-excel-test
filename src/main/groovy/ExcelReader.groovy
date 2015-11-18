import groovy.util.logging.Log4j
import org.apache.poi.hssf.usermodel.*
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFSheet

@Log4j
class ExcelReader {

    File file

    ExcelReader(File file) {
        log.info "Let's process the file:${file.name}"
        this.file = file
    }

    public void exec() {
        new FileInputStream(file).withStream { fis ->
            //ブックオブジェクトを作成
            // Workbookは、HSSFWorkbook, XSSFWorkbookらのインターフェース
            Workbook workbook = WorkbookFactory.create(fis)

            log.info "workbook class:${workbook.class.name}"
            log.info "number of sheets:${workbook.numberOfSheets}"

            // HSSFWorkbookの場合、下記の方法だとエラーになる。
            //  workbook.eachWithIndex { Sheet sheet, int idx ->

            workbook.numberOfSheets.times { int idx ->
                Sheet sheet = workbook.getSheetAt(idx)
                // Sheetは、XSSFSheet、HSSFSheetらのインターフェース
                log.info "---------------------------------"
                log.info "sheet index:$idx"
                log.info "sheet name:${sheet.sheetName}"
                log.info "row count:${sheet.getLastRowNum() + 1}"
                log.info "col count:${sheet.lastRowNum > 0 ? sheet.getRow(0).asList().size() : 'There is no row!'}"
                if (sheet.lastRowNum > 1) {
                    log.info "cell type analysis>==============="
                    sheet.getRow(1).asList().eachWithIndex { Cell cell, int idx2 ->
                        def content =  "cell index:$idx2, type:${getCellTypeName(cell.getCellType())}, data:${getCellData(cell)}"
                        // <= poiバージョンが3.13だとここでClassNotFoundエラーになる
                        // 参考：http://stackoverflow.com/questions/10330593/apache-poi-exception-in-reading-xlsx-files

                        if(idx2 == 0){
                            content += ", as long:${cell.getNumericCellValue() as long}"
                        }
                        if(idx2 == 2){
                            content += ", as int:${cell.getNumericCellValue() as int}"
                        }
                        if(idx2 == 3){
                            content += ", as date:${cell.getDateCellValue()}"
                        }
                        log.info content
                    }
                    log.info "================<analysis end!"
                }
            }
        }
    }

    private Object getCellData(Cell cell){
        // 参考：http://stackoverflow.com/questions/5578535/get-cell-value-from-excel-sheet-with-apache-poi
        switch (cell.cellType) {
            case Cell.CELL_TYPE_BOOLEAN:
                return cell.getBooleanCellValue()
            case Cell.CELL_TYPE_NUMERIC:
                return cell.getNumericCellValue()
            case Cell.CELL_TYPE_STRING:
                return cell.getStringCellValue()
            case Cell.CELL_TYPE_BLANK:
                return null
            case Cell.CELL_TYPE_ERROR:
                return cell.getErrorCellValue()
            case Cell.CELL_TYPE_FORMULA:
                throw new RuntimeException("CELL_TYPE_FORMULA is not permitted!")
        }
    }

    // TODO too ugly. You must die now.
    private String getCellTypeName(int type) {
        switch (type) {
            case Cell.CELL_TYPE_BLANK:
                return "CELL_TYPE_BLANK"
            case Cell.CELL_TYPE_BOOLEAN:
                return "CELL_TYPE_BOOLEAN"
            case Cell.CELL_TYPE_ERROR:
                return "CELL_TYPE_ERROR"
            case Cell.CELL_TYPE_FORMULA:
                return "CELL_TYPE_FORMULA"
            case Cell.CELL_TYPE_NUMERIC:
                return "CELL_TYPE_NUMERIC"
            case Cell.CELL_TYPE_STRING:
                return "CELL_TYPE_STRING"
            default:
                throw new RuntimeException("unknown cell type:$type")
        }
    }
}
