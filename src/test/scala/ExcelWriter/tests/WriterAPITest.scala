package ExcelWriter.tests

import org.scalatest.flatspec.AnyFlatSpec
import com.excelwriter._

class WriterAPITest extends AnyFlatSpec {

  case class People (Name: String, Occupation: String, Age: Int)

  case class Computers (Name: String, Type: String, YearOfManufacture: Int, Value: Int)

  val family = List(
    People("Hugo","Developer",43),
    People("Karen","Data Analyst",43),
    People("Holly","Student",9),
  )

  val comps = List(
    Computers("Hugo Macbook Pro", "Laptop", 2013, 250),
    Computers("Karen Macbook Pro", "Laptop", 2013, 250),
    Computers("Holly Air", "Laptop", 2015, 350),
    Computers("Mac Mini", "Desktop", 2018, 500),
  )



  "it " should "look elegant" in {
    var writer = ExcelWriter(1000)
      writer.addSheet(
        ExcelSheet(writer, "Family", family.iterator)
        .addColumn(Some("First Name"), x => x.Name)
        .addColumn(Some("Occupation"), x => x.Occupation)
        .addColumn(Some("Age"), x=>x.Age)
        .write())
      writer.addSheet(
        ExcelSheet(writer, "Computers", comps.iterator)
          .addColumn(Some("Name"), x => x.Name)
          .addColumn(Some("Type"), x => x.Type)
          .addColumn(Some("Year of Manufacture"), x=>x.YearOfManufacture)
          .addColumn(Some("Value"), x => x.Value)
          .write())
    .writeToFile("/Users/hugo/", "ScalaExcel.xlsx")

  }

  "it " should "write big sheets" in {
    var writer = ExcelWriter(1000)
    val data = (1 until 1000000).map(x => { family(x % 3) }).iterator
    writer.addSheet(
      ExcelSheet(writer, "Family", data, 100)
        .addColumn(Some("First Name"), x => x.Name)
        .addColumn(Some("Occupation"), x => x.Occupation)
        .addColumn(Some("Age"), x=>x.Age).write()
    )
        //.write)
      .writeToFile("/Users/hugo/", "ScalaExcel.xlsx")

  }

  "it " should "create multiple sheets" in {
    var writer = ExcelWriter(1000)
    val data = (1 to 4).map(x => { family(x % 3) }).iterator
    writer.addSheet(
      ExcelSheet(writer, "Family", data, 3)
        .addColumn(Some("First Name"), x => x.Name)
        .addColumn(Some("Occupation"), x => x.Occupation)
        .addColumn(Some("Age"), x=>x.Age).write()
    )
      //.write)
      .writeToFile("/Users/hugo/", "ScalaExcel.xlsx")

  }

  "it " should "stop when we get to the max rows limit" in {
    var writer = ExcelWriter(1000)
    val data = (1 to 50).map(x => { family(x % 3) }).iterator
    writer.addSheet(
      ExcelSheet(writer, "Family", data, 50, 10)
        .addColumn(Some("First Name"), x => x.Name)
        .addColumn(Some("Occupation"), x => x.Occupation)
        .addColumn(Some("Age"), x=>x.Age).write()
    )
      //.write)
      .writeToFile("/Users/hugo/", "ScalaExcel.xlsx")

  }

  "when multiple sheet and we hit max rows it " should "stop processing" in {
    var writer = ExcelWriter(1000)
    val data = (1 to 10).map(x => { family(x % 3) }).iterator
    writer.addSheet(
      ExcelSheet(writer, "Family", data, 3, 5)
        .addColumn(Some("First Name"), x => x.Name)
        .addColumn(Some("Occupation"), x => x.Occupation)
        .addColumn(Some("Age"), x=>x.Age).write()
    )
      //.write)
      .writeToFile("/Users/hugo/", "ScalaExcel.xlsx")

  }

}
