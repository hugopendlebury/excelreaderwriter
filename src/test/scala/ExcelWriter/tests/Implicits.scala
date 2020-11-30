package ExcelWriter.tests

import javax.xml.stream.XMLStreamReader

object Implicits {

  implicit class XMLStreamReaderStream(xmlreader: XMLStreamReader) {

    def toStreamHP = {
      println("getting stream1")
      new Iterator[Int] {
        def hasNext = xmlreader.hasNext

        def next() = xmlreader.next()
      }.toStream
    }

    def toStream2 = {
        print("getting stream")
        val cache = xmlreader
        Iterator.continually(cache).takeWhile(x => x.hasNext).map(_.next())
    }

  }
}