package com.example

import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.nio.file.Paths

fun main(args: Array<String>) {

    val file = Paths.get("PoiSampleWorkbook.xlsx").toFile()
    val book = WorkbookFactory.create(file)
    val sheet = book.getSheet("Sheet1")

    sheet.rowIterator().asSequence().forEach { row ->
        val values = row.cellIterator().asSequence().map { cell ->
            when (cell.cellType) {
                CellType.NUMERIC -> cell.numericCellValue.toString()
                CellType.STRING -> cell.stringCellValue
                else -> throw RuntimeException("CellType=${cell.cellType}]")
            }
        }
        println(values.toList())
    }
}
