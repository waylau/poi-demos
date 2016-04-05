package com.waylau.poi.hssf;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Demonstrates how to create borders around cells.
 *
 * @author waylau.com
 * @date  2015-1-27
 */
public class Borders {
    public static void main(String[] args) throws IOException {
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet("new sheet");

        // Create a row and put some cells in it. Rows are 0 based.
        HSSFRow row = sheet.createRow(1);

        // Create a cell and put a value in it.
        HSSFCell cell = row.createCell(1);
        cell.setCellValue(4);

        // Style the cell with borders all around.
        HSSFCellStyle style = wb.createCellStyle();
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBottomBorderColor(HSSFColor.BLACK.index);
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style.setLeftBorderColor(HSSFColor.GREEN.index);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setRightBorderColor(HSSFColor.BLUE.index);
        style.setBorderTop(HSSFCellStyle.BORDER_MEDIUM_DASHED);
        style.setTopBorderColor(HSSFColor.ORANGE.index);
        cell.setCellStyle(style);

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
        wb.close();
    }
}