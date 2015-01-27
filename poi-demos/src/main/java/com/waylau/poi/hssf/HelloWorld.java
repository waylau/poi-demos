package com.waylau.poi.hssf;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;

/**
 * Hello world!
 * 
 * @author waylau.com
 * @date  2015-1-27
 */
public class HelloWorld 
{
    public static void main( String[] args ) throws IOException
    {
        //创建工作簿
        Workbook wb = new HSSFWorkbook();  // 或 new XSSFWorkbook();
        
        //创建 sheet
        Sheet sheet1 = wb.createSheet("第1个 sheet");
        Sheet sheet2 = wb.createSheet("第2个 sheet");

        // 注意,Excel中sheet的名字是不得超过31个字符
        // 并且，必须不能包含下面字符:
        // 0x0000
        // 0x0003
        // colon (:)
        // backslash (\)
        // asterisk (*)
        // question mark (?)
        // forward slash (/)
        // opening square bracket ([)
        // closing square bracket (])

        // 使用org.apache.poi.ss.util.WorkbookUtil#createSafeSheetName(String nameProposal)}
        // 用安全的方法来创建有效的名称,这个工具用 (' ') 替换无效字符
        String safeName = WorkbookUtil.createSafeSheetName("[Waylau's Blog*?]"); // 返回 " Waylau's Blog   "

        Sheet sheet3 = wb.createSheet(safeName);
        
        //创建数据 行。Row 是从  0 开始的
        Row row = sheet1.createRow((short)0);
        
        //创建 cell 放值进去
        Cell cell0 = row.createCell(0);
        cell0.setCellValue(1);
        
        //或者在一行代码实现
        row.createCell(1).setCellValue(1.2);
        row.createCell(2).setCellValue(true);
        
        //创建 CreationHelper 处理实例化具体类的不同实例
        CreationHelper createHelper = wb.getCreationHelper();
        row.createCell(3).setCellValue(createHelper.createRichTextString("This is a string"));
        
        //创建 CellStyle,用来设置   Cell 的样式
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setDataFormat(
            createHelper.createDataFormat().getFormat("m/d/yy h:mm"));
        
        Cell cell4 = row.createCell(4);
        cell4.setCellValue(new Date());
        cell4.setCellStyle(cellStyle);
        
        row.createCell(5).setCellValue(Calendar.getInstance());
        row.createCell(6).setCellType(Cell.CELL_TYPE_ERROR);
        
        //生成文件
        FileOutputStream fileOut = new FileOutputStream("helloword.xls");
        wb.write(fileOut);
        fileOut.close();
        
        System.out.println( "已经生成文件");
    }
}
