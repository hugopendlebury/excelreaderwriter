package com.excelwriter

import scala.reflect.runtime.universe._
import org.apache.poi.xssf.usermodel.XSSFCell

case class Column[TSrc, T : TypeTag](columnName: Option[String],
                     fieldValueAccessor: TSrc => T
                    ) {

}

case class ExcelColumn[TSrc]() {


  def MapColumn[T : TypeTag](accessor: TSrc => T, columnName: Option[String] ) = {
    val func = (src: TSrc) => {
      accessor(src)
    }


  }

}
