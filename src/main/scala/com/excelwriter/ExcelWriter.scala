package com.excelwriter

import java.io.FileOutputStream

import org.apache.poi.xssf.streaming.SXSSFWorkbook
import org.apache.poi.ss.util.CellRangeAddress

case class ExcelWriter (numRows: Int = 1000) {


  val wb = new SXSSFWorkbook(numRows)
  wb.setCompressTempFiles(true) // temp files will be gzipped

  //ExcelSheet[TSrc](writer: ExcelWriter, sheetName: String, data: Iterable[TSrc])

  def addSheet[TSrc](sheet: ExcelSheet[TSrc]) = {
    this
  }

  def addSheet[TSrc](sheetName: String,
                  data: Iterable[TSrc]
  ): ExcelWriter = {
    this
  }

  def writeToFile(path: String, fileName: String): Unit = {
    val fullName = s"$path$fileName"
    val out = new FileOutputStream(fullName)
    wb.write(out)
    out.close()
  }

}
