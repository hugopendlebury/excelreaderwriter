package ExcelWriter.tests

import org.scalatest.flatspec.AnyFlatSpec

import scala.collection.IterableOnce.iterableOnceExtensionMethods

class StreamTest extends AnyFlatSpec  {

  lazy val fibs: Stream[BigInt] = BigInt(0) #:: BigInt(1) #:: fibs.zip(fibs.tail).map { n => n._1 + n._2 }
  /*
  def doIt(size: Int) = {
    val s : Stream[Int]
    s #:: (0 until size).zip(_)

  }
  */

  lazy val rows: Stream[Int] = (0 until 100).map(_ * 2).toStream

  "it " should "read a stream" in {

  }

}
