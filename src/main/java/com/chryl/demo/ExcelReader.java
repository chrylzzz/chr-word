package com.chryl.demo;

/**
 * Created by Chr.yl on 2024/3/9.
 *
 * @author Chr.yl
 */
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {
    public static void main(String[] args) {
        try (FileInputStream fis = new FileInputStream("students.xlsx");
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {
            // 获取第一个工作表
            XSSFSheet sheet = workbook.getSheetAt(0);

            // 遍历每一行
            for (Row row : sheet) {
                // 读取每一列的数据
                String name = row.getCell(0).getStringCellValue();
                int age = (int) row.getCell(1).getNumericCellValue();
                double score = row.getCell(2).getNumericCellValue();

                // 处理数据...
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}