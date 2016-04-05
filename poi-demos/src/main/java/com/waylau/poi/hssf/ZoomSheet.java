package com.waylau.poi.hssf;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;

import java.io.IOException;
import java.io.FileOutputStream;

/**
 * Sets the zoom magnication for a sheet.
 *
 * @author waylau.com
 * @date  2015-1-27
 */
public class ZoomSheet
{
    public static void main(String[] args)
        throws IOException
    {
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet1 = wb.createSheet("new sheet");
        sheet1.setZoom(3,4);   // 75 percent magnification
        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
        wb.close();
    }
}