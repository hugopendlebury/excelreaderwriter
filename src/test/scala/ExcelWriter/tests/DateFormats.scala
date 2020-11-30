package ExcelWriter.tests

object Formats {

  private val builtInFormats = Map (
    1 -> "0",
    2 -> "0.00",
    3 -> "#,##0",
    4 -> "#,##0.00",
    5 -> "$#,##0_);($#,##0)",
    6 -> "$#,##0_);[Red]($#,##0)",
    7 -> "$#,##0.00_);($#,##0.00)",
    8 -> "$#,##0.00_);[Red]($#,##0.00)",
    9 -> "0%",
    10 -> "0.00%",
    11 -> "0.00E+00",
    12 -> "# ?/?",
    13 -> "# ??/??",
    //Note 14 is not consistent and seems to vary from the standard which states
    //it is m/d/yyyy - but try it in excel and see that it is dd/mm/yyyy
    14 -> "dd/mm/yyyy",
    15 -> "d-mmm-yy",
    16 -> "d-mmm",
    17 -> "mmm-yy",
    18 -> "h:mm AM/PM",
    19 -> "h:mm:ss AM/PM",
    20 -> "h:mm",
    21 -> "h:mm:ss",
    22 -> "M/d/yyyy h:mm",
    37 -> "#,##0_);(#,##0)",
    38 -> "#,##0_);[Red](#,##0)",
    39 -> "#,##0.00_);(#,##0.00)",
    40 -> "#,##0.00_);[Red](#,##0.00)",
    45 -> "mm:ss",
    46 -> "[h]:mm:ss",
    47 -> "mm:ss.0",
    48 -> "##0.0E+0",
    49 -> "@"
  )

  private val dateFormats = Map (
    14 -> "dd/MM/yyyy",
    15 -> "d-MMM-yy",
    16 -> "d-MMM",
    17 -> "MMM-yy",
    18 -> "h:mm AM/PM",
    19 -> "h:mm:ss AM/PM",
    20 -> "h:mm",
    21 -> "h:mm:ss",
    22 -> "M/d/yy h:mm",
    30 -> "M/d/yy",
    34 -> "yyyy-MM-dd",
    45 -> "mm:ss",
    46 -> "[h]:mm:ss",
    47 -> "mmss.0",
    51 -> "MM-dd",
    52 -> "yyyy-MM-dd",
    53 -> "yyyy-MM-dd",
    55 -> "yyyy-MM-dd",
    56 -> "yyyy-MM-dd",
    165 -> "M/d/yy",
    166 -> "dd MMMM yyyy",
    167 -> "dd/MM/yyyy",
    168 -> "dd/MM/yy",
    169 -> "d.M.yy",
    170 -> "yyyy-MM-dd",
    171 -> "dd MMMM yyyy",
    172 -> "d MMMM yyyy",
    173 -> "M/d",
    174 -> "M/d/yy",
    175 -> "MM/dd/yy",
    176 -> "d-MMM",
    177 -> "d-MMM-yy",
    178 -> "dd-MMM-yy",
    179 -> "MMM-yy",
    180 -> "MMMM-yy"
  )

  def getFormat(styleId: Int) = {
    val builtIn = builtInFormats.get(styleId)
    if(builtIn.nonEmpty) builtIn else dateFormats.get(styleId)

  }

}
