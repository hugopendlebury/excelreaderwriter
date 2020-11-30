package ExcelWriter.tests

import java.util.zip.ZipFile
import java.util.{Optional}
import org.apache.poi.ss.usermodel.DataFormatter
import ExcelWriter.tests.ExcelImplicits.{ExcelIntConverters}

import javax.xml.stream.{XMLInputFactory, XMLStreamConstants, XMLStreamReader}

import scala.collection.mutable.ListBuffer

case class RawSheetParser ( zipFile: ZipFile,
                            sheetName: String,
                            sharedString: Optional[List[String]],
                            styles: Map[Int, Option[String]]
                          ) extends Iterable[List[ExcelCell]] {

  private val rowCache = new ListBuffer[List[ExcelCell]]
  private var rowCacheIterator = rowCache.iterator
  private val rowCacheSize = 10
  private var stop = false
  private var startSheetSection = false
  private var fillMissing = false

  private val sheets = RawExcelHelper.getFiles(zipFile)
  private val sheet = sheets.filter(x => x._1.toLowerCase() == s"xl/worksheets/$sheetName.xml".toLowerCase())
                            .findFirst().get()
  val inputFactory = XMLInputFactory.newInstance()
  val reader: XMLStreamReader = inputFactory.createXMLStreamReader(zipFile.getInputStream(sheet._2))
  val stream = Wrapper(reader)
  advanceToSheetData
  val (minCol: Int, maxCol: Int) = bufferRows


  private def advanceToSheetData = {
    stream.advanceTo(() => reader.getEventType == XMLStreamConstants.START_ELEMENT &&
        reader.getLocalName == "sheetData" )
    startSheetSection = true;
  }

  private def bufferRows = {
    val buffer = this.take(10)
    //Get the maximum column count -- values can be sparse in the xml
    val minCol = buffer.foldLeft(1)((a,b) => a.min(b.minBy(x => x.column).column))
    val maxCol = buffer.foldLeft(0)((a,b) => a.max(b.maxBy(x => x.column).column))
    val filledRows = fillMissingRows(buffer, minCol, maxCol)
    rowCacheIterator = filledRows.iterator
    fillMissing = true
    (minCol, maxCol)
  }

  def fillMissingRows(buffer: Iterable[List[ExcelCell]], minCol: Int, maxCol: Int ) = {
    buffer.map(row => fillMissingRow(row, minCol, maxCol))
  }

  def fillMissingRow(row: List[ExcelCell], minCol: Int = this.minCol, maxCol: Int = this.maxCol) = {
      val requiredColumns = (minCol to maxCol).toSet
      val rowId = row.minBy(x => x.row).row
      val columns = row.map(x => x.column)
      val missingCols = requiredColumns diff columns.toSet
      if (missingCols.isEmpty) {
        row
      } else {
        val newCells = missingCols.toList.map(x => ExcelCell(s"${x.toExcelColumnName}$rowId", "", ""))
        (row ++ newCells).sortBy(x => x.column)
      }

  }

  def hasMore() = {
    val s = rowCache.size < rowCacheSize
    val y = reader.hasNext
    s && y
  }

  private def getCell() = {

    val attrCnt = stream.reader.getAttributeCount
    val attributes = (0 until attrCnt).map(x => {
      (reader.getAttributeLocalName(x), reader.getAttributeValue(x))
    }).toMap
    val cellAddress = attributes.get("r").get
    val t = attributes.get("t")
    val isString = !t.isEmpty && t.get == "s"
    var value: String = ""
    val styleIdx = attributes.get("s")
    stream.take(1)
    //println (s"cell info type = ${reader.getEventType} name = ${reader.getLocalName}")
    while (stream.reader.hasNext && !isEndOfCell) {
      if (isValue) {
        val v = stream.reader.getElementText
        value = if (isString) sharedString.get()(v.toInt) else v
      }
      stream.take(1)
    }

    val fmt = if(styleIdx.nonEmpty) {
      val index = styleIdx.getOrElse("0").toInt
      styles.get(index).getOrElse(None)
    } else {
      None
    }

    val formattedValue = if(styleIdx.nonEmpty && fmt.nonEmpty) {
      if(value != "") {
        new DataFormatter().formatRawCellContents(value.toDouble, -1, fmt.get)
      } else {
        value
      }
    }
    else {
      value
    }

    ExcelCell(cellAddress, value, formattedValue)

  }

  private def getRow(): List[ExcelCell] = {
    val rowData = ListBuffer[ExcelCell]()
     do{
       stream.advanceTo(() => isNewCell())
       val cell = getCell()
       rowData += cell
       stream.take(1)
    } while (stream.reader.hasNext && !isEndOfRow())
    stream.take(1) //advance stream
    if(isEndOfData()) {
      stop = true
    }
    println(s"returning $rowData")

    if(fillMissing) {
      fillMissingRow(rowData.toList)
    } else {
      rowData.toList
    }

  }

  private def isNewCell() = {
    reader.getEventType == XMLStreamConstants.START_ELEMENT && reader.getLocalName == "c"
  }

  private def isEndOfCell() = {
    reader.getEventType == XMLStreamConstants.END_ELEMENT && reader.getLocalName == "c"
  }

  private def isEndOfRow() = {
    reader.getEventType == XMLStreamConstants.END_ELEMENT && reader.getLocalName == "row"
  }

  private def hasMoreCells() = {
    reader.getEventType != XMLStreamConstants.END_ELEMENT && reader.getLocalName != "row"
  }

  private def isNewRow() = {
    reader.getEventType == XMLStreamConstants.START_ELEMENT && reader.getLocalName == "row"
  }

  private def isEndOfData() = {
    reader.getEventType == XMLStreamConstants.END_ELEMENT && reader.getLocalName == "sheetData"
  }

  private def isValue = {
    reader.getEventType == XMLStreamConstants.START_ELEMENT && reader.getLocalName == "v"
  }

  private def getData  = try {

      rowCache.clear
      while (hasMore() && !isEndOfRow() && !stop) {
        if(!isNewRow()) stream.take(1)
        if(isNewRow) {
          rowCache += getRow()
        }
      }
      rowCacheIterator = rowCache.iterator
      rowCacheIterator.hasNext || !stop
    } catch {
      case e => throw new Exception("Error reading XML stream", e)
    }


  override def iterator: scala.Iterator[List[ExcelCell]] = {
    val it = new StreamingRowIterator()
    it
  }

  class StreamingRowIterator() extends scala.Iterator[List[ExcelCell]] {

    if (rowCacheIterator == null) hasNext

    def hasNext: Boolean = (rowCacheIterator != null && rowCacheIterator.hasNext) || getData

    def next: List[ExcelCell] = rowCacheIterator.next

  }

}
