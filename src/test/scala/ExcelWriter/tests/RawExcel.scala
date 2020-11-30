package ExcelWriter.tests

import java.util.zip.ZipFile

import scala.collection.mutable.ListBuffer


case class RawExcel() {

  import javax.xml.stream.XMLInputFactory
  import javax.xml.stream.XMLStreamConstants
  import javax.xml.stream.XMLStreamReader
  import Implicits.XMLStreamReaderStream


  def getSheetMeta(zipFile: ZipFile) = {
    //we can have workbook in _rels we want the one which is not in _rels
    val workbook = RawExcelHelper.getFiles(zipFile, "workbook.xml").filter(x => x._1.endsWith("workbook.xml")).findFirst().get()
    val inputFactory = XMLInputFactory.newInstance()
    var reader: XMLStreamReader = null
    try {
      reader = inputFactory.createXMLStreamReader(zipFile.getInputStream(workbook._2))
      reader.toStreamHP.map(event => {
          if (event == XMLStreamConstants.START_ELEMENT) {
            val elementName = reader.getLocalName
            if (elementName == "sheet") {
              val attrCnt = reader.getAttributeCount
              val attributes = (0 until attrCnt).map(x => {
                (reader.getAttributeLocalName(x), reader.getAttributeValue(x))
              }).toMap

              Some(SheetDetails(attributes.getOrElse("name", ""),
                                attributes.getOrElse("sheetId", ""),
                                attributes.getOrElse("id", ""))
              )
            } else {
              None
            }
          } else {
            None
          }
        }
      ).filter(x => !x.isEmpty)
        .map(x => x.get)
    } finally if (reader != null) reader.close()
  }

  def getSharedStrings(zipFile: ZipFile) = {
    val strings = RawExcelHelper.getFiles(zipFile, "sharedStrings")

    strings.map(x => {
      val reader = XMLInputFactory.newInstance().createXMLStreamReader(zipFile.getInputStream(x._2))

      val strings = reader.toStreamHP.withFilter(_ => reader.getEventType == XMLStreamConstants.CHARACTERS &&
                                                    reader.getText.trim.length > 0)
        .map(_ => reader.getText).toList
      strings
    }).findFirst()


  }

  private def getCustomNumFormats(stream: Wrapper) = {
    var customFormats = new scala.collection.mutable.HashMap[Int, String]()
    do {
      if(stream.reader.isStartElement && stream.reader.getLocalName == "numFmt") {
        val attrCnt = stream.reader.getAttributeCount
        val attributes = (0 until attrCnt).map(x => {
          (stream.reader.getAttributeLocalName(x), stream.reader.getAttributeValue(x))
        }).toMap
        customFormats += (attributes("numFmtId").toInt ->  attributes("formatCode"))
      }
      stream.reader.next
    }  while (stream.reader.hasNext &&
      ! (stream.reader.isEndElement && stream.reader.getLocalName == "numFmts"))
    customFormats.foreach(x => println(s"custom format $x"))
    customFormats.toMap
  }

  private def getDefinedFormats(stream: Wrapper) = {
    val definedFormats = ListBuffer[Int]()
    do {
      if(stream.reader.isStartElement && stream.reader.getLocalName == "xf") {
        val attrCnt = stream.reader.getAttributeCount
        val attributes = (0 until attrCnt).map(x => {
          (stream.reader.getAttributeLocalName(x), stream.reader.getAttributeValue(x))
        }).toMap
         definedFormats += attributes("numFmtId").toInt
      }
      stream.reader.next
    } while (stream.reader.hasNext &&
      ! (stream.reader.isEndElement && stream.reader.getLocalName == "cellXfs"))
    definedFormats.toList
  }

  def getStyles(zipFile: ZipFile) = {
    val stylesSheet = RawExcelHelper.getFiles(zipFile, "styles").findFirst().get()
    val reader = XMLInputFactory.newInstance().createXMLStreamReader(zipFile.getInputStream(stylesSheet._2))
    val stream = Wrapper(reader)
    var customFormats = Map[Int, String]()
    var definedFormats = List[Int]()
    stream.foreach(_ => {
      if(stream.reader.isStartElement && stream.reader.getLocalName == "numFmts") {
        customFormats  = getCustomNumFormats(stream)
      }
      if(stream.reader.isStartElement && stream.reader.getLocalName == "cellXfs") {
        definedFormats  = getDefinedFormats(stream)
      }
    })
    //now map the formats
    definedFormats.zipWithIndex.map{ case(numFmtId, idx) => {
      val fmt = Formats.getFormat(numFmtId)
      val customFmt = customFormats.get(numFmtId)
      //Give the custom format preference
      val fmtStr = if(customFmt.nonEmpty) customFmt else fmt
      (idx, fmtStr)
    }}.toMap

  }


  def getSheetsXML(zipFile: ZipFile) = {
    RawExcelHelper.getFiles(zipFile, "sheet")
  }

}
