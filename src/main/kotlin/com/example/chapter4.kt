package com.example

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.ss.util.CellUtil
import java.nio.file.Paths

fun main(args: Array<String>) {

    val file = Paths.get("PoiSampleWorkbook.xlsx").toFile()
    val book = WorkbookFactory.create(file)
    val sheet = book.getSheet("Sheet1")

    fun Sheet.cells(row: Int = 0, col: Int = 0) = CellUtil.getCell(CellUtil.getRow(row, this), col)
    fun Cell.value() = when (cellType) {
        CellType.NUMERIC -> numericCellValue.toString()
        CellType.STRING -> stringCellValue
        else -> throw RuntimeException("CellType=$cellType]")
    }

    println(sheet.cells(1, 0).value())
    println(sheet.cells(1, 1).value())
}


