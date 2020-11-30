package test.excel.writing

import org.scalatest.flatspec.AnyFlatSpec


import org.apache.poi.xssf.streaming.SXSSFWorkbook
import org.apache.poi.xssf.streaming.SXSSFSheet
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import java.io.FileOutputStream
import org.apache.poi.ss.util.CellReference

class TestExcelStreaming extends AnyFlatSpec {

  "it " should "create a spreadsheet" in {


    val wb = new SXSSFWorkbook(1000)
    wb.setCompressTempFiles(true) // temp files will be gzipped

    val sh = wb.createSheet
    for (rownum <- 0 until 10000) {
      val row = sh.createRow(rownum)
      for (cellnum <- 0 until 10) {
        val cell = row.createCell(cellnum)
        val address = new CellReference(cell).formatAsString
        cell.setCellValue(address)
      }
      /*
      // manually control how rows are flushed to disk
      if (rownum % 100 == 0) {
        sh.asInstanceOf[SXSSFSheet].flushRows(100) // retain 100 last rows and flush all others

        // ((SXSSFSheet)sh).flushRows() is a shortcut for ((SXSSFSheet)sh).flushRows(0),
        // this method flushes all rows
      }

       */
    }
    val out = new FileOutputStream("/Users/hugo/sxssf.xlsx")
    wb.write(out)
    out.close()
    // dispose of temporary files backing this workbook on disk

  }

}