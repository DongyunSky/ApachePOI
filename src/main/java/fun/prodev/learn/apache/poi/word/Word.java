package fun.prodev.learn.apache.poi.word;

import fun.prodev.learn.apache.poi.Constants;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;

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

        // 创建标题
        XWPFParagraph titleParagraph = xwpfDocument.createParagraph();
        titleParagraph.setAlignment(ParagraphAlignment.CENTER); // 设置段落居中
        XWPFRun titleParagraphRun = titleParagraph.createRun();

        titleParagraphRun.setText("Java PoI");
        titleParagraphRun.setColor("000000");
        titleParagraphRun.setFontSize(20);

        // 段落
        XWPFParagraph paragraph = xwpfDocument.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText("At tutorialspoint.com, we strive hard to provide quality tutorials for self-learning purpose in the domains of Academics, Information Technology, Management and Computer Programming Languages.");

        // 换行
        XWPFParagraph paragraph1 = xwpfDocument.createParagraph();
        XWPFRun paragraphRun1 = paragraph1.createRun();
        paragraphRun1.setText("\r");

        // 表格
        XWPFTable table = xwpfDocument.createTable();
        // 填充数据 表格生成后自带一行一个单元格
        XWPFTableRow row = table.getRow(0);
        row.getCell(0).setText("标题列1");
        row.addNewTableCell().setText("第2个标题列");
        row.addNewTableCell().setText("第3列");
        row.addNewTableCell().setText("第4列");
        for (int i = 0; i < 10; i++) {
            row = table.createRow();
            for (int j = 0; j < 4; j++) {
                row.getCell(j).setText((i+1) + "行" + j + "列");
            }
        }
        table.getCTTbl().getTblPr().unsetTblBorders(); // 去表格边框
        // 列宽自动分割
        CTTblWidth infoTableWidth = table.getCTTbl().addNewTblPr().addNewTblW();
        infoTableWidth.setType(STTblWidth.DXA);
        infoTableWidth.setW(BigInteger.valueOf(9072));

        CTSectPr sectPr = xwpfDocument.getDocument().getBody().addNewSectPr();
        XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(xwpfDocument, sectPr);
        // 添加页眉
        CTP ctpHeader = CTP.Factory.newInstance();
        CTR ctrHeader = ctpHeader.addNewR();
        CTText ctHeader = ctrHeader.addNewT();
        String headerText = "ctpHeader";
        ctHeader.setStringValue(headerText);
        XWPFParagraph headerParagraph = new XWPFParagraph(ctpHeader, xwpfDocument);
        // 设置为右对齐
        headerParagraph.setAlignment(ParagraphAlignment.RIGHT);
        XWPFParagraph[] parsHeader = new XWPFParagraph[1];
        parsHeader[0] = headerParagraph;
        policy.createHeader(XWPFHeaderFooterPolicy.DEFAULT, parsHeader);

        // 添加页脚
        CTP ctpFooter = CTP.Factory.newInstance();
        CTR ctrFooter = ctpFooter.addNewR();
        CTText ctFooter = ctrFooter.addNewT();
        String footerText = "ctpFooter";
        ctFooter.setStringValue(footerText);
        XWPFParagraph footerParagraph = new XWPFParagraph(ctpFooter, xwpfDocument);
        headerParagraph.setAlignment(ParagraphAlignment.CENTER);
        XWPFParagraph[] parsFooter = new XWPFParagraph[1];
        parsFooter[0] = footerParagraph;
        policy.createFooter(XWPFHeaderFooterPolicy.DEFAULT, parsFooter);

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
