package com.excelwriter

import org.apache.poi.xssf.streaming.{SXSSFCell, SXSSFSheet}

import scala.collection.mutable.ListBuffer
import scala.reflect.runtime.universe._

case class ExcelSheet[TSrc](writer: ExcelWriter,
                            sheetName: String,
                            data: Iterator[TSrc],
                            rowsPerSheet: Int = 1000000,
                            maximumRows: Int = 2000000) {

  val sh = writer.wb.createSheet(sheetName)

  private var _writers = new ListBuffer[(SXSSFCell, TSrc) => Unit]()
  private var _columnNames = new ListBuffer[Option[String]]()

  private def writeValue[T: TypeTag](cell: SXSSFCell, src: TSrc)(writer: TSrc => T) = {

    val value = writer(src)
    typeOf[T] match {
      case t if t =:= typeOf[String] => cell.setCellValue(value.asInstanceOf[String])
      case t if t =:= typeOf[Int] => cell.setCellValue(value.asInstanceOf[Int])
      case t if t =:= typeOf[Double] => cell.setCellValue(value.asInstanceOf[Double])
      case _ => cell.setCellValue(value.toString)
    }
  }

  private def writeHeaderRow(rowCnt: Int) = {
    val colsWithIndex = _columnNames.zipWithIndex
    val sheetRow = sh.createRow(rowCnt)
    for(col <- colsWithIndex) {
      val cell = sheetRow.createCell(col._2)
      cell.setCellValue(col._1.getOrElse(""))
    }
    rowCnt + 1
  }

  def write(totalRowCntAllSheets: Int = 0 ): ExcelSheet[TSrc] = {
    var rowCnt = 0
    var runningRowCount = totalRowCntAllSheets
    var ret = this
    if (_columnNames.filter((x => !x.isEmpty)).length > 0) {
      val newCnt = writeHeaderRow(rowCnt)
      rowCnt = newCnt
      runningRowCount = runningRowCount + newCnt
    }
    val colWriter = _writers.zipWithIndex

    while (data.hasNext && rowCnt <= rowsPerSheet && runningRowCount <= maximumRows  ) {
      val row = data.next()
      val sheetRow = sh.createRow(rowCnt)
      for (writer <- colWriter) {
        val cell = sheetRow.createCell(writer._2)
        writer._1(cell, row)
      }
      rowCnt = rowCnt + 1
      runningRowCount = runningRowCount + 1
      ret = if (rowCnt == rowsPerSheet && runningRowCount <= maximumRows) {
        val newSheetName = s"$sheetName${runningRowCount / rowCnt}"
        val newSheetCopy = this.copy(sheetName = newSheetName)
        newSheetCopy._columnNames = this._columnNames
        newSheetCopy._writers = this._writers
        newSheetCopy.write(runningRowCount)
      } else {
        this
      }
    }
    /*
    for (row <- data) {

    }
    */

    ret
  }

    def addColumn[T: TypeTag](columnName: Option[String],
                              columnAccessor: TSrc => T) = {
      _columnNames += columnName
      val func = writeValue(_: SXSSFCell, _: TSrc)(columnAccessor)
      _writers += func
      this
    }

}



  /*
  def addSheet[TSrc](sheetName: String,
                     data: Iterable[TSrc],
                     columnAccessors: Iterable[Column[TSrc]],
                     freezePanes: Boolean = false,
                     autoFilters: Boolean = false,
                     maxRows: Int=1000000 ,
                     totalRowCnt: Int=0): ExcelWriter = {

    val sh = wb.createSheet(sheetName)
    val cols = columnAccessors.toList
    val colsWithIdx = cols.zipWithIndex
    val hasHeaderRow = cols.filter(x => !x.columnName.isEmpty).length > 0
    var rowCnt = 0;
    var total = totalRowCnt;

    if(hasHeaderRow) {
      val sheetRow = sh.createRow(rowCnt)
      for (col <- colsWithIdx) {
        val cell = sheetRow.createCell(col._2)
        cell.setCellValue(col._1.columnName.getOrElse(" "))
        if(freezePanes) sh.createFreezePane(0,1)
      }
      rowCnt = rowCnt + 1
    }

    data.foreach(row => {
      val sheetRow = sh.createRow(rowCnt)
      colsWithIdx.foreach(col => {
        val cell = sheetRow.createCell(col._2)
        val value = col._1.fieldValueAccessor(row)
        cell.setCellValue(value.toString)
      })
      rowCnt = rowCnt + 1;
      total = total + rowCnt
      if(rowCnt == 1000000) this.addSheet(s"$sheetName", data, cols, freezePanes, autoFilters, maxRows, rowCnt)
    })

    if(autoFilters) sh.setAutoFilter(new CellRangeAddress(0, rowCnt, 0, cols.length))
    this
  }
*/

