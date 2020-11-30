package ExcelWriter.tests

import org.scalatest.flatspec.AnyFlatSpec
import ExcelWriter.tests.ExcelImplicits.{ExcelIntConverters, ExcelStringConverters}

class ImplicitTests extends AnyFlatSpec {

  "with an input of 1 we" should "return A" in {
    val columnName = 1.toExcelColumnName
    assert(columnName == "A")
  }

  "with an input of A we" should "return 1" in {
    val columnName = "A".toExcelColumnNumber
    assert(columnName == 1)
  }

  "with an input of 26 we" should "return Z" in {
    val columnName = 26.toExcelColumnName
    assert(columnName == "Z")
  }

  "with an input of Z we" should "return 26" in {
    val columnName = "Z".toExcelColumnNumber
    assert(columnName == 26)
  }

  "with an input of 27 we" should "return AA" in {
    val columnName = 27.toExcelColumnName
    assert(columnName == "AA")
  }

  "with an input of AA we" should "return 27" in {
    val columnName = "AA".toExcelColumnNumber
    assert(columnName == 27)
  }

  "with an input of 702 we" should "return ZZ" in {
    val columnName = 702.toExcelColumnName
    assert(columnName == "ZZ")
  }

  "with an input of ZZ we" should "return 702" in {
    val columnName = "ZZ".toExcelColumnNumber
    assert(columnName == 702)
  }

  "with an input of 703 we" should "return AAA" in {
    val columnName = 703.toExcelColumnName
    assert(columnName == "AAA")
  }

  "with an input of AAA we" should "return 703" in {
    val columnName = "AAA".toExcelColumnNumber
    assert(columnName == 703)
  }

  "with an input of A1 we" should "return 1" in {
    val row = "A1".toExcelRowNumber
    assert(row == 1)
  }

  "with an input of Z123 we" should "return 123" in {
    val row = "A123".toExcelRowNumber
    assert(row == 123)
  }

}
