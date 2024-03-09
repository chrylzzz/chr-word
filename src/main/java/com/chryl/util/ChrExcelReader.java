package com.chryl.util;

/**
 * Created by Chr.yl on 2024/3/9.
 *
 * @author Chr.yl
 */

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class ChrExcelReader {
    public static void main(String[] args) {
        try (FileInputStream fis = new FileInputStream("/Users/chryl/Downloads/temp.xlsx");
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {
            // 获取第一个工作表
            XSSFSheet sheet = workbook.getSheetAt(0);

            // 遍历每一行
            for (Row row : sheet) {
                // 读取每一列的数据
//                String name = row.getCell(0).getStringCellValue();
                String function = row.getCell(3).getStringCellValue();
                String functionDesc = row.getCell(4).getStringCellValue();
                System.out.println(function);
                System.out.println(functionDesc);

                // 处理数据...
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}