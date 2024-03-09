package com.chryl.demo;

/**
 * Created by Chr.yl on 2024/3/9.
 *
 * @author Chr.yl
 */

import org.apache.poi.xwpf.usermodel.*;

import java.io.FileOutputStream;

public class WordTableExample {
    public static void main(String[] args) throws Exception {
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
    }
}
