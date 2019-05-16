package fun.prodev.learn.apache.poi.word;

import fun.prodev.learn.apache.poi.Constants;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @author prodev
 * @date 2019/5/13 10:14
 * @description POI Word
 */
@RestController
@RequestMapping("word")
public class Word {

    public static void createWord() throws IOException {

        // 创建
        XWPFDocument xwpfDocument = new XWPFDocument();
        // 段落
        XWPFParagraph paragraph = xwpfDocument.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText("At tutorialspoint.com, we strive hard to provide quality tutorials for self-learning purpose in the domains of Academics, Information Technology, Management and Computer Programming Languages.");

        // 表格
        XWPFTable table = xwpfDocument.createTable();
        //create first row
        XWPFTableRow tableRowOne = table.getRow(0);
        tableRowOne.getCell(0).setText("col one, row one");
        tableRowOne.addNewTableCell().setText("col two, row one");
        tableRowOne.addNewTableCell().setText("col three, row one");
        //create second row
        XWPFTableRow tableRowTwo = table.createRow();
        tableRowTwo.getCell(0).setText("col one, row two");
        tableRowTwo.getCell(1).setText("col two, row two");
        tableRowTwo.getCell(2).setText("col three, row two");
        //create third row
        XWPFTableRow tableRowThree = table.createRow();
        tableRowThree.getCell(0).setText("col one, row three");
        tableRowThree.getCell(1).setText("col two, row three");
        tableRowThree.getCell(2).setText("col three, row three");

        // 写出
        File newWord = new File(Constants.TEMP_PATH + "new Word.docx");
        FileOutputStream fileOutputStream = new FileOutputStream(newWord);
        xwpfDocument.write(fileOutputStream);
        fileOutputStream.close();

        XWPFWordExtractor we = new XWPFWordExtractor(xwpfDocument);
        System.out.println(we.getText());
    }

    public static void main(String[] args) throws IOException {
        createWord();
    }
}
