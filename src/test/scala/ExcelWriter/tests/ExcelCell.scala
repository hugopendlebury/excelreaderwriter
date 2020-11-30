package ExcelWriter.tests

import ExcelWriter.tests.ExcelImplicits.{ExcelStringConverters}

case class ExcelCell (address: String, value: String, formattedValue: String) {
  val row: Int = address.toExcelRowNumber
  val column: Int = address.toExcelColumnNumber

}
