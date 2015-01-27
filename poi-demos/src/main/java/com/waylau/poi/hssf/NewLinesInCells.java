package com.waylau.poi.hssf;

import org.apache.poi.hssf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Demonstrates how to use newlines in cells.
 *
 * @author waylau.com
 * @date  2015-1-27
 */
public class NewLinesInCells {
    public static void main( String[] args ) throws IOException {

        HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet s = wb.createSheet();
		HSSFRow r = null;
		HSSFCell c = null;
		HSSFCellStyle cs = wb.createCellStyle();
		HSSFFont f2 = wb.createFont();

		cs = wb.createCellStyle();

		cs.setFont(f2);
		// Word Wrap MUST be turned on
		cs.setWrapText(true);

		r = s.createRow(2);
		r.setHeight((short) 0x349);
		c = r.createCell(2);
		c.setCellType(HSSFCell.CELL_TYPE_STRING);
		c.setCellValue("Use \n with word wrap on to create a new line");
		c.setCellStyle(cs);
		s.setColumnWidth(2, (int) ((50 * 8) / ((double) 1 / 20)));

		FileOutputStream fileOut = new FileOutputStream("workbook.xls");
		wb.write(fileOut);
		fileOut.close();
    }
}