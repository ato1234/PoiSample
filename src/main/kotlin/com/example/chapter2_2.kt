package com.example

import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.ss.util.CellUtil
import java.nio.file.Paths

fun main(args: Array<String>) {

    val file = Paths.get("PoiSampleWorkbook.xlsx").toFile()
    val book = WorkbookFactory.create(file)
    val sheet = book.getSheet("Sheet1")

    // indexは0オリジン
    val row = CellUtil.getRow(1, sheet)
    val cell = CellUtil.getCell(row, 0)
    println(cell.stringCellValue)

    // 存在しないセルを指定してもぬるぽにならず、CellType.BLANKのセルが取得できる
    val row2 = CellUtil.getRow(999, sheet)
    val cell2 = CellUtil.getCell(row2, 999)
    println("CellType=${cell2.cellType}, value=${cell2.stringCellValue}")
}
