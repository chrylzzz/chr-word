package com.chryl.excelToWordTable;

/**
 * excel 数据转为 word表格
 * Created by Chr.yl on 2024/3/9.
 *
 * @author Chr.yl
 */

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public class ChrWordTable2 {
    public static void main(String[] args) throws Exception {
        List<LinkedHashMap<String, String>> linkedHashMapList = new ArrayList<>();
        Map<String, String> LinkedHashMap = new LinkedHashMap<>();

        /**
         * read excel
         */
        try (FileInputStream fis = new FileInputStream("/Users/chryl/Downloads/temp3.xlsx");
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {
            // 获取第一个工作表
            XSSFSheet sheet = workbook.getSheetAt(0);

            /**
             * read excel
             */


            /**
             * write word
             */
            // 创建一个新的Word文档
            XWPFDocument document = new XWPFDocument();

            // 创建一个表格
            //excel表格，方法rowNum从0开始，所以行数+1
//            XWPFTable table = document.createTable(sheet.getLastRowNum() + 1, 2);
            XWPFTable table = document.createTable(sheet.getLastRowNum() + 1, 3);

            //--------------------------------------------------------
            // 遍历每一行
            for (Row row : sheet) {
                // 读取每一列的数据
                String firFunction = row.getCell(2).getStringCellValue();
                String function = row.getCell(3).getStringCellValue();
//                String functionDesc = row.getCell(4).getStringCellValue();
                int rowNum = row.getRowNum();
                System.out.println("rowNum: " + rowNum);

//                System.out.println(function);
//                System.out.println(functionDesc);
               //写入word中
                table.getRow(rowNum).getCell(0).setText(firFunction);
                table.getRow(rowNum).getCell(1).setText(function);
//                table.getRow(rowNum).getCell(2).setText(functionDesc);

                // 处理数据...
            }

            //--------------------------------------------------------

            // 将文档保存到文件系统
//        FileOutputStream out = new FileOutputStream("/Users/chryl/Downloads/table_example.docx");
            FileOutputStream out = new FileOutputStream("/Users/chryl/Downloads/table_example.doc");
            document.write(out);
            out.close();

            System.out.println("Word文档已创建，包含表格。");

            /**
             * write word
             */

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}