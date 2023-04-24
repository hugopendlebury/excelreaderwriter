# excelreader

Scala library to read and write excel files. 

Fluent API and Low Memory usage via streaming.

Sample usage below

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


In reality we would probably read the above in from something like a CSV or database. The iterator and lambda expressions
enable us to avoid typos in column names and avoid the need to write any loops. The library will do the write the data for you based on your iterator

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
