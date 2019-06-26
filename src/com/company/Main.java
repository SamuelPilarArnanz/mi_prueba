package com.company;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import java.io.File;
import java.io.FileOutputStream;

//jdbc
public class Main {

    public static void main(String[] args) {

        try {
            XWPFDocument document = new XWPFDocument();
            FileOutputStream out = new FileOutputStream(new File("c::/poidemo.docx"));

            XWPFParagraph paragraph =  document.createParagraph();
            XWPFRun run = paragraph.createRun();
            run.setText("This is my prove");
            document.write(out);
            out.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
