package com.excelwriter.Traits

trait IAccessColumn[TSrc] {

  val columnName: String

  def columnAccess[T]: TSrc => T


}
