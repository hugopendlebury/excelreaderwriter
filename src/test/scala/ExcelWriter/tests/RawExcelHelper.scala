package ExcelWriter.tests

import java.util.stream
import java.util.zip.{ZipEntry, ZipFile}

object RawExcelHelper {

  def getFiles(zipFile: ZipFile): stream.Stream[(String, ZipEntry)] = {
    zipFile.stream().map(x => (x.getName, x))
  }

  def getFiles(zipFile: ZipFile, name: String): stream.Stream[(String, ZipEntry)] = {
    getFiles(zipFile).filter(x => x._1.contains(name))
  }

}
