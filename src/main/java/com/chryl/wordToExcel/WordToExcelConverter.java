package com.chryl.wordToExcel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.FileInputStream;
import java.io.FileOutputStream;
/**
 * 参考模板
 * Created by Chr.yl on 2024/3/28.
 *
 * @author Chr.yl
 */
public class WordToExcelConverter {
    public static void main(String[] args) throws Exception {
        // 加载Word文档
        FileInputStream wordInputStream = new FileInputStream("/Users/chryl/Downloads/document.docx");
        XWPFDocument wordDocument = new XWPFDocument(wordInputStream);

        // 创建Excel工作簿
        Workbook excelWorkbook = new XSSFWorkbook();
        Sheet excelSheet = excelWorkbook.createSheet("标题列表");

        // 读取标题并写入Excel
        int rowNum = 0;
        XWPFParagraph paragraph;
        for (IBodyElement element : wordDocument.getBodyElements()) {
            if (element instanceof XWPFParagraph) {
                paragraph = (XWPFParagraph) element;
                String text = paragraph.getText();
                if (text != null && !text.trim().isEmpty()) {
                    Row row = excelSheet.createRow(rowNum++);
                    Cell cell = row.createCell(0);
                    cell.setCellValue(text);
                }
            }
        }

        // 将Excel写入文件
        FileOutputStream excelOutputStream = new FileOutputStream("/Users/chryl/Downloads/titles.xlsx");
        excelWorkbook.write(excelOutputStream);
        excelOutputStream.close();
        excelWorkbook.close();
        wordInputStream.close();

        System.out.println("Excel文档已创建，包含表格。");
    }
}

