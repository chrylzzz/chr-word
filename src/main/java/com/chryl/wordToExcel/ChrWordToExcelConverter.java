package com.chryl.wordToExcel;

/**
 * Created by Chr.yl on 2024/3/28.
 *
 * @author Chr.yl
 */

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

public class ChrWordToExcelConverter {
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
                //-------

                // 检查粗体格
                String styleID = paragraph.getStyleID();
                if (paragraph.getCTP().getPPr().getPStyle() != null) {//标题

                    Integer styleIntId = Integer.valueOf(styleID);
                    if (styleIntId == 5) {
                        if (text != null && !text.trim().isEmpty()) {
                            Row row = excelSheet.createRow(rowNum++);
                            Cell cell = row.createCell(0);
                            cell.setCellValue(text);
                        }
                    } else if (styleIntId == 6) {
                        if (text != null && !text.trim().isEmpty()) {
                            Row row = excelSheet.createRow(rowNum++);
                            Cell cell = row.createCell(1);
                            cell.setCellValue(text);
                        }
                    } else if (styleIntId == 7) {
                        if (text != null && !text.trim().isEmpty()) {
                            Row row = excelSheet.createRow(rowNum++);
                            Cell cell = row.createCell(2);
                            cell.setCellValue(text);
                        }
                    }
                } else {//正文
                    if (text != null && !text.trim().isEmpty()) {
                        Row row = excelSheet.createRow(rowNum++);
                        Cell cell = row.createCell(3);
                        cell.setCellValue(text);
                    }
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

