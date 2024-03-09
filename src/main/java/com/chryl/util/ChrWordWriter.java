package com.chryl.util;

/**
 * Created by Chr.yl on 2024/3/9.
 *
 * @author Chr.yl
 */

import com.chryl.po.Student;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ChrWordWriter {
    public static void main(String[] args) {
        try (FileOutputStream fos = new FileOutputStream("/Users/chryl/Downloads/chrtotemp.doc");
             XWPFDocument document = new XWPFDocument()) {
            // 创建一个段落
            XWPFParagraph paragraph = document.createParagraph();

            // 创建一个运行
            XWPFRun run = paragraph.createRun();

            // 设置文本内容
            run.setText("学生信息报告");

            // 添加一个表格
            XWPFTable table = document.createTable();

            // 在表格中添加表头
            XWPFTableRow headerRow = table.getRow(0);
            headerRow.getCell(0).setText("姓名");
//            headerRow.getCell(1).setText("年龄");


            List<Student> students = new ArrayList<>();
            students.add(new Student("chiyulin","17"));
            students.add(new Student("maytun","62"));
            // 在表格中添加数据行
            for (Student student : students) {
                XWPFTableRow dataRow = table.createRow();
                dataRow.getCell(0).setText(student.getName());
//                dataRow.getCell(1).setText(student.getAge());
            }
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
