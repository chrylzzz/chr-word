package com.chryl.demo;

/**
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

public class ChrWordTableDemo {

    public static void main(String[] args) throws Exception {
        List<LinkedHashMap<String, String>> linkedHashMapList = new ArrayList<>();
        Map<String, String> LinkedHashMap = new LinkedHashMap<>();

        /**
         * read excel
         */
        try (FileInputStream fis = new FileInputStream("/Users/chryl/Downloads/temp.xlsx");
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {
            // 获取第一个工作表
            XSSFSheet sheet = workbook.getSheetAt(0);


            // 遍历每一行
            for (Row row : sheet) {
                // 读取每一列的数据
//                String name = row.getCell(0).getStringCellValue();
//                String firFunction = row.getCell(2).getStringCellValue();
//                if(StringUtils.isBlank(){
//
//                }

                String function = row.getCell(3).getStringCellValue();
                String functionDesc = row.getCell(4).getStringCellValue();
                System.out.println(row.getRowNum());
                System.out.println(function);
                System.out.println(functionDesc);

                LinkedHashMap.put(function, functionDesc);

                // 处理数据...
            }
            /**
             * read excel
             */


            /**
             * write word
             */
            // 创建一个新的Word文档
            XWPFDocument document = new XWPFDocument();

            // 创建一个表格
            XWPFTable table = document.createTable(2, 2); // 2行2列的表格

            // 填充表格数据
            table.getRow(0).getCell(0).setText("单元格1");
            table.getRow(0).getCell(1).setText("单元格2");
            table.getRow(1).getCell(0).setText("单元格3");
            table.getRow(1).getCell(1).setText("单元格4");

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