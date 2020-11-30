package ExcelWriter.tests

import java.util.zip.ZipFile

import org.scalatest.flatspec.AnyFlatSpec

class RawParserTests extends AnyFlatSpec {


  "we" should "be able to read sheet data" in {
    val file = "/Users/hugo/TestMissing.xlsx"
    val zipFile = new ZipFile(file)
    var rawParser = RawExcel()
    val strings = rawParser.getSharedStrings(zipFile)
    val sheets = rawParser.getSheetMeta(zipFile).toList
    val z  = 10
  }

  "we" should "be able to read the main data" in {
    val file = "/Users/hugo/TestMissing.xlsx"
    val zipFile = new ZipFile(file)
    var rawParser = RawExcel()
    val strings = rawParser.getSharedStrings(zipFile)
    val styles = rawParser.getStyles(zipFile)
    var sheet = RawSheetParser(zipFile, "Sheet1", strings, styles)
    val data = sheet.foreach(x => x.foreach(c => println(s"test call cell is ${c}")))
    val z  = 10

  }

  "we" should "be able to read formatted data" in {
    val file = "/Users/hugo/DateAndNumberFormats.xlsx"
    val zipFile = new ZipFile(file)
    val rawParser = RawExcel()
    val strings = rawParser.getSharedStrings(zipFile)
    val styles = rawParser.getStyles(zipFile)
    styles.foreach(x => println(x))
    val sheet = RawSheetParser(zipFile, "Sheet1", strings, styles)
    val data = sheet.toList
    val firstRow = data(0)
    val cells = firstRow.map(x => (x.address, x)).toMap
    assert(cells("A1").value == "43875")
    assert(cells("A1").formattedValue == "14/02/2020")
    assert(cells("B1").formattedValue == "14/01/2020 15:30:00")
    assert(cells("C1").formattedValue == "123.12")
    assert(cells("D1").formattedValue == "123")
    assert(cells("E1").formattedValue == "123.1")

  }

}
