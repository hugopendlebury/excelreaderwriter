package ExcelWriter.tests

import javafx.css.StyleableIntegerProperty
import org.xml.sax.Attributes
import org.xml.sax.helpers.DefaultHandler

import scala.xml.InputSource
import org.apache.poi.xssf.eventusermodel.XSSFReader
import org.xml.sax.helpers.XMLReaderFactory

object StylesParser {

  def populateNumberFormats(r: XSSFReader): Map[Int, Int] = {
    var cache = Map[Int, Int]()

    try {
      val is = r.getStylesData
      try {

        var styleIndex = 0;
        val parser = XMLReaderFactory.createXMLReader


        val handler: DefaultHandler = new DefaultHandler {

          override def startDocument(): Unit = {
            super.startDocument
          }

          override def startElement(uri: String, localName: String, name: String, attributes: Attributes): Unit = {
            if (name == "xf") {
              cache += styleIndex -> attributes.getValue("numFmtId").toInt
              styleIndex = styleIndex + 1;
            }

          }
      }

      parser.setContentHandler(handler)
      if (is != null) parser.parse(new InputSource(is))

      } finally if (is != null) is.close()
    }
    cache
  }
}
