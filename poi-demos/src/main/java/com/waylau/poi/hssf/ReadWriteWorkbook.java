package com.waylau.poi.hssf;

import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFCell;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
/**
 * This example demonstrates opening a workbook, modifying it and writing
 * the results back out.
 *
 * @author waylau.com
 * @date  2015-1-27
 */
public class ReadWriteWorkbook {
    public static void main(String[] args) throws IOException {
        FileInputStream fileIn = null;
        FileOutputStream fileOut = null;
        HSSFWorkbook wb = null;
        try
        {
            fileIn = new FileInputStream("workbook.xls");
            POIFSFileSystem fs = new POIFSFileSystem(fileIn);
            wb = new HSSFWorkbook(fs);
            HSSFSheet sheet = wb.getSheetAt(0);
            HSSFRow row = sheet.getRow(2);
            if (row == null)
                row = sheet.createRow(2);
            HSSFCell cell = row.getCell(3);
            if (cell == null)
                cell = row.createCell(3);
            cell.setCellType(HSSFCell.CELL_TYPE_STRING);
            cell.setCellValue("a test");

            // Write the output to a file
            fileOut = new FileOutputStream("workbookout.xls");
            wb.write(fileOut);
        } finally {
            if (fileOut != null)
                fileOut.close();
            if (fileIn != null)
                fileIn.close();
            if (wb != null)
                wb.close();
 
        }
    }
}