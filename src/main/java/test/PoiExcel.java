package test;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class PoiExcel {
    public static void main(String[] args) {
        int cellNo = 1;// 单元格下标
        Cell cell = null;// 单元格对象
        int rowNo = 0;// 行下标
        Row row = null;// 行对象

        // 创建工作簿
        Workbook workbook = new HSSFWorkbook();
        // 创建工作表
        Sheet sheet = workbook.createSheet();
        // 设置单元格宽
        sheet.setColumnWidth(cellNo++, 30*256);
        sheet.setColumnWidth(cellNo++, 30*256);
        sheet.setColumnWidth(cellNo++, 30*256);

        //--------------设置大标题--------------
        row = sheet.createRow(rowNo);// 创建大标题的行对象
        row.setHeightInPoints(36);// 设置行高
        cellNo = 1;// 重置单元格下标为1
        cell = row.createCell(cellNo);// 在当前行上创建一个单元格对象
        sheet.addMergedRegion(new CellRangeAddress(rowNo, rowNo, cellNo, cellNo+2));// 合并单元格
        cell.setCellValue("学生成绩表");// 设置单元格内容
//        cell.setCellStyle(this.bigTitle(workbook));// 设置单元格样式

        //--------------设置小标题--------------
        row = sheet.createRow(++rowNo);
        row.setHeightInPoints(26.25f);// 设置行高
        String titles[] = {"学号","姓名","成绩（单位：分）"};
        // 创建单元格对象，设置内容与样式
        for (String title : titles) {
            cell = row.createCell(cellNo++);
            cell.setCellValue(title);
//            cell.setCellStyle(this.title(workbook));
        }

        //--------------模拟数据输出--------------
        row = sheet.createRow(++rowNo);
        row.setHeightInPoints(24);// 设置行高

        cellNo = 1;// 重置单元格下标为1
        cell = row.createCell(cellNo++);
        cell.setCellValue("200000");// 设置单元格内容，学号200000
//        cell.setCellStyle(this.text(workbook));

        cell = row.createCell(cellNo++);
        cell.setCellValue("老王");// 设置单元格内容，老王
//        cell.setCellStyle(this.text(workbook));

        cell = row.createCell(cellNo++);
        cell.setCellValue("59.9");// 设置单元格内容，59.9分
//        cell.setCellStyle(this.text(workbook));

        // 保存，关闭流对象，在C盘生成excel测试.xls文件
        OutputStream os = null;
        try {
            os = new FileOutputStream("D:\\excel测试.xls");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        try {
            workbook.write(os);
        } catch (IOException e) {
            e.printStackTrace();
        }
        try {
            os.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
