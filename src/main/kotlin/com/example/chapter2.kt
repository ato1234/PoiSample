package com.example

import org.apache.poi.ss.usermodel.WorkbookFactory
import java.nio.file.Paths

fun main(args: Array<String>) {

    val file = Paths.get("PoiSampleWorkbook.xlsx").toFile()
    val book = WorkbookFactory.create(file)
    val sheet = book.getSheet("Sheet1")

    // indexは0オリジン
    val row = sheet.getRow(1)
    if (row != null) {
        val cell = row.getCell(0)
        if (cell != null) {
            println(cell.stringCellValue)
        }
    }
}
