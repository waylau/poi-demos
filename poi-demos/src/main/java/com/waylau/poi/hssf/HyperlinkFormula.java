package com.waylau.poi.hssf;

import org.apache.poi.hssf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Test if hyperlink formula, with url that got more than 127 characters, works
 *
 * @author waylau.com
 * @date  2015-1-27
 */
public class HyperlinkFormula {
    public static void main(String[] args) throws IOException {
    	HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet("new sheet");
        HSSFRow row = sheet.createRow(0);

        HSSFCell cell = row.createCell(0);
        cell.setCellType(HSSFCell.CELL_TYPE_FORMULA);
        cell.setCellFormula("HYPERLINK(\"http://127.0.0.1:8080/toto/truc/index.html?test=aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa\", \"test\")");

        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
        wb.close();
    }
}