package ExcelWriter.tests

import java.io.InputStream
import java.util
import org.apache.poi.util.XMLHelper
import org.apache.poi.openxml4j.opc.OPCPackage
import org.apache.poi.xssf.eventusermodel.XSSFReader
import org.apache.poi.xssf.model.SharedStringsTable
import org.xml.sax.Attributes
import org.xml.sax.ContentHandler
import org.xml.sax.InputSource
import org.xml.sax.SAXException
import org.xml.sax.XMLReader
import org.xml.sax.helpers.DefaultHandler
import javax.xml.parsers.ParserConfigurationException
import Implicits.XMLStreamReaderStream


import org.scalatest.flatspec.AnyFlatSpec

class ReaderTest extends AnyFlatSpec {

  import org.apache.poi.openxml4j.opc.OPCPackage
  import org.apache.poi.util.XMLHelper
  import org.apache.poi.xssf.eventusermodel.XSSFReader
  import org.apache.poi.xssf.model.SharedStringsTable
  import org.xml.sax.XMLReader
  import javax.xml.parsers.ParserConfigurationException

  @throws[Exception]
  def processOneSheet(filename: String): Unit = {
    val pkg = OPCPackage.open(filename)
    val r = new XSSFReader(pkg)
    val sst = r.getSharedStringsTable
    val parser = fetchSheetParser(sst)
    // To look up the Sheet Name / Sheet Order / rID,
    //  you need to process the core Workbook stream.
    // Normally it's of the form rId# or rSheet#
    val sheet2 = r.getSheet("rId2")
    val sheetSource = new InputSource(sheet2)
    parser.parse(sheetSource)
    sheet2.close
  }


  def processAllSheets(filename: String): Unit = {
    val pkg = OPCPackage.open(filename)
    val r = new XSSFReader(pkg)
    val sst = r.getSharedStringsTable
    val parser = fetchSheetParser(sst)
    val sheets = r.getSheetsData
    while ( {
      sheets.hasNext
    }) {
      System.out.println("Processing new sheet:\n")
      val sheet = sheets.next
      val sheetSource = new InputSource(sheet)
      parser.parse(sheetSource)
      sheet.close
      System.out.println("")
    }
  }


  def fetchSheetParser(sst: SharedStringsTable): XMLReader = {
    val parser = XMLHelper.newXMLReader
    val handler = new SheetHandler(sst)
    parser.setContentHandler(handler)
    parser
  }

  /**
   * See org.xml.sax.helpers.DefaultHandler javadocs
   */
  private class SheetHandler (var sst: SharedStringsTable) extends DefaultHandler {
    private var lastContents: String = null
    private var nextIsString = false

    override def startElement(uri: String, localName: String, name: String, attributes: Attributes): Unit = { // c => cell
      if (name == "c") { // Print the cell reference
        print(attributes.getValue("r") + " - ")
        // Figure out if the value is an index in the SST
        val cellType = attributes.getValue("t")
        if (cellType != null && cellType == "s") nextIsString = true
        else nextIsString = false
      }
      // Clear contents cache
      lastContents = ""
    }

    override def endElement(uri: String, localName: String, name: String): Unit = { // Process the last contents as required.
      // Do now, as characters() may be called more than once
      if (nextIsString) {
        val idx = lastContents.toInt
        lastContents = sst.getItemAt(idx).getString
        nextIsString = false
      }
      // v => contents of a cell
      // Output after we've seen the string contents
      if (name == "v") System.out.println(lastContents)
    }

    override def characters(ch: Array[Char], start: Int, length: Int): Unit = {
      lastContents += new String(ch, start, length)
    }
  }


  import org.apache.poi.xssf.eventusermodel.XSSFReader


  "we" should "be able to read a sheet" in {
    processAllSheets("/Users/hugo/TestMissing.xlsx")
  }

  "we" should "be able to read the shared strings" in {
    val file = "/Users/hugo/TestMissing.xlsx"
    val pkg = OPCPackage.open(file)
    val r = new XSSFReader(pkg)
    SharedStringParser.populateSharedStrings(r)
  }

  "we" should "be able to read formatting" in {
    val file = "/Users/hugo/extracts/formatting/Formatting.xlsx"
    val pkg = OPCPackage.open(file)
    val r = new XSSFReader(pkg)
    StylesParser.populateNumberFormats(r)
  }

  "we" should "be able to read sheet data" in {
    val file = "/Users/hugo/TestMissing.xlsx"
    val pkg = OPCPackage.open(file)
    val r = new XSSFReader(pkg)
    val sharedStrings = SharedStringParser.populateSharedStrings(r)
    val s = SheetParser()
    s.parse(r, sharedStrings)
    val results = s.map(s => s).toList

  }

}
