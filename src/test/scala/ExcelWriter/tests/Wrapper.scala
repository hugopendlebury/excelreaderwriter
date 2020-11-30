package ExcelWriter.tests

import javax.xml.stream.{ XMLStreamReader }

case class Wrapper(reader: XMLStreamReader)  extends Iterable[Int] {

  override def iterator: Iterator[Int] = new Iterator[Int] {
    def hasNext = reader.hasNext
    def next = reader.next()
  }

  def advanceTo(func: () => Boolean) = {
    if(!func()) {
      this.takeWhile(_ => !func())
    }
  }

}
