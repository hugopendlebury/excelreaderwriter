package ExcelWriter.tests

import org.xml.sax.Attributes
import org.xml.sax.helpers.DefaultHandler

import scala.collection.mutable.ListBuffer
import scala.xml.InputSource
import org.apache.poi.xssf.eventusermodel.XSSFReader
import org.xml.sax.helpers.XMLReaderFactory

object SharedStringParser {

  def populateSharedStrings(r: XSSFReader): List[String] = {
    try {
      val is = r.getSharedStringsData
      try {
        var si = false
        var t = false

        val parser = XMLReaderFactory.createXMLReader
        val ss = new ListBuffer[String]();

        val handler: DefaultHandler = new DefaultHandler {

          override def startDocument(): Unit = {
            super.startDocument
          }

          override def startElement(uri: String, localName: String, name: String, attributes: Attributes): Unit = {
            if (name == "si") {
              si = true
            } else if(name == "t") {
              t = true;
            }
          }

          override def endElement(uri: String, localName: String, qName: String): Unit = {
            if (localName == "si") {
              si = false
            } else if(localName == "t") {
              t = true;
            }
          }

          override def characters(ch: Array[Char], start: Int, length: Int): Unit = {
            if(si && t) {
              ss += new String(ch, start, length)
            }
          }

        }

        parser.setContentHandler(handler)
        if(is != null) parser.parse(new InputSource(is))

        ss.toList

      } finally if (is != null) is.close()
    }
  }
}
