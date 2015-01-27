package com.waylau.poi.hssf;

import org.apache.poi.hssf.usermodel.*;

import java.io.IOException;
import java.io.FileOutputStream;

/**
 * @author waylau.com
 * @date  2015-1-27
 */
public class RepeatingRowsAndColumns {
    public static void main(String[] args) throws IOException {
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet1 = wb.createSheet("first sheet");
        wb.createSheet("second sheet");
        wb.createSheet("third sheet");

        HSSFFont boldFont = wb.createFont();
        boldFont.setFontHeightInPoints((short)22);
        boldFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);

        HSSFCellStyle boldStyle = wb.createCellStyle();
        boldStyle.setFont(boldFont);

        HSSFRow row = sheet1.createRow(1);
        HSSFCell cell = row.createCell(0);
        cell.setCellValue("This quick brown fox");
        cell.setCellStyle(boldStyle);

        // Set the columns to repeat from column 0 to 2 on the first sheet
        wb.setRepeatingRowsAndColumns(0,0,2,-1,-1);
        // Set the rows to repeat from row 0 to 2 on the second sheet.
        wb.setRepeatingRowsAndColumns(1,-1,-1,0,2);
        // Set the the repeating rows and columns on the third sheet.
        wb.setRepeatingRowsAndColumns(2,4,5,1,2);

        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }
}