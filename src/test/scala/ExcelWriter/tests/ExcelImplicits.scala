package ExcelWriter.tests

object ExcelImplicits {

  implicit class ExcelIntConverters(src: Int) {

    def toExcelColumnName: String = {
      var dividend = src
      var colName = ""
      while (dividend > 0) {
        val modulo = (dividend - 1) % 26
        colName = (65 + modulo).toChar + colName
        dividend = (dividend - modulo) / 26
      }
      colName
    }
  }

  implicit class ExcelStringConverters(src: String) {

    def toExcelRowNumber: Int = {
      src.filter(x => x >= '0' && x <= '9').toInt
    }

    def toExcelColumnNumber: Int = {
      src.filter(x => x >= 'A' && x <= 'Z')
        .foldLeft(0)((a,b) => (a * 26) + (b - 'A' + 1))
    }

  }


}
