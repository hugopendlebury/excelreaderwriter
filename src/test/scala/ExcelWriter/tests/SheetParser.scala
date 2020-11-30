package ExcelWriter.tests

import org.apache.poi.util.XMLHelper
import org.apache.poi.xssf.eventusermodel.XSSFReader
import java.util.{ArrayList, Iterator}

import javax.xml.stream.events.XMLEvent
import javax.xml.stream.{XMLEventReader, XMLStreamConstants, XMLStreamReader}


case class SheetParser () extends Iterable[String] {


  private val rowCache = new ArrayList[String]
  private var rowCacheIterator = rowCache.iterator
  private val rowCacheSize = 10
  private var parser : XMLStreamReader = null

  def parse(r: XSSFReader, sharedStrings: List[String]) {

    //for now take the first sheet
    val sheets = r.getSheetsData
    val hasnext = sheets.hasNext
    val is = sheets.next()
    val factory = XMLHelper.newXMLInputFactory()
    parser = factory.createXMLStreamReader(is)
    val z = 0
  }

  def hasMore() = {
    val s = rowCache.size < rowCacheSize
    val y = parser.hasNext
    s && y
  }

  private def getRow = try {
    //owCache.clear
    while (hasMore()) {
      parser.next()
      val zz = 10
      //handleEvent(parser.nextEvent)
    }
    rowCacheIterator = rowCache.iterator
    rowCacheIterator.hasNext
  } catch {
    case e => throw new Exception("Error reading XML stream", e)
  }

  private def handleEvent(event: XMLEvent): Unit = {
    val z = 10;
  }

  override def iterator: scala.Iterator[String] = {
    //val it = new StreamingRowIterator().asInstanceOf[scala.Iterator[String]]
    val it = new StreamingRowIterator()
    it
  }

  class StreamingRowIterator() extends scala.Iterator[String] {
    if (rowCacheIterator == null) hasNext

    def hasNext: Boolean = (rowCacheIterator != null && rowCacheIterator.hasNext) || getRow

    def next: String = rowCacheIterator.next

  }



}
