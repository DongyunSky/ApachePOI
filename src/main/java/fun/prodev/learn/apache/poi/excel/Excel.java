package fun.prodev.learn.apache.poi.excel;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellType;

import java.io.*;
import java.util.Date;

/**
 * @author prodev
 * @date 2019/5/12 20:55
 * @description POI 操作 Excel 文件
 */
public class Excel {

    private static final String filePath = "D:\\Download\\Temp\\template.xls";

    public static void createExcel() throws IOException {
        // 获取生成文件的路径
        // FileSystemView fsv = FileSystemView.getFileSystemView();
        // String desktop = fsv.getHomeDirectory().getPath();
        // String filePath = desktop + "/template.xls";
        System.out.println(filePath);
        File excel = new File(filePath);
        FileOutputStream fileOutputStream = new FileOutputStream(excel);

        HSSFWorkbook hssfWorkbook = new HSSFWorkbook(); // Excel 文件
        HSSFSheet sheet1 = hssfWorkbook.createSheet("Sheet1");
        HSSFRow row = sheet1.createRow(0);
        row.createCell(0).setCellValue("id");
        row.createCell(1).setCellValue("订单号");
        row.createCell(2).setCellValue("下单时间");
        row.createCell(3).setCellValue("个数");
        row.createCell(4).setCellValue("单价");
        row.createCell(5).setCellValue("订单金额");
        row.setHeightInPoints(30); // 设置行的高度

        HSSFRow row1 = sheet1.createRow(1);
        row1.createCell(0).setCellValue("1");
        row1.createCell(1).setCellValue("NO00001");

        // 日期格式化
        HSSFCellStyle cellStyle2 = hssfWorkbook.createCellStyle();
        HSSFCreationHelper creationHelper = hssfWorkbook.getCreationHelper();
        cellStyle2.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));
        sheet1.setColumnWidth(2, 20 * 256); // 设置列的宽度

        HSSFCell cell2 = row1.createCell(2);
        cell2.setCellStyle(cellStyle2);
        cell2.setCellValue(new Date());
        row1.createCell(3).setCellValue(2);

        // 保留两位小数
        HSSFCellStyle cellStyle3 = hssfWorkbook.createCellStyle();
        cellStyle3.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));
        HSSFCell cell4 = row1.createCell(4);
        cell4.setCellStyle(cellStyle3);
        cell4.setCellValue(29.5);

        // 货币格式化
        HSSFCellStyle cellStyle4 = hssfWorkbook.createCellStyle();
        HSSFFont font = hssfWorkbook.createFont();
        font.setFontName("华文行楷");
        font.setFontHeightInPoints((short) 15);
        font.setColor(new HSSFColor().getIndex());
        cellStyle4.setFont(font);

        HSSFCell cell5 = row1.createCell(5);
        cell5.setCellFormula("D2*E2");  // 设置计算公式

        // 获取计算公式的值
        HSSFFormulaEvaluator e = new HSSFFormulaEvaluator(hssfWorkbook);
        cell5 = e.evaluateInCell(cell5);
        System.out.println(cell5.getNumericCellValue());


        hssfWorkbook.setActiveSheet(0);
        hssfWorkbook.write(fileOutputStream);
        fileOutputStream.close();

    }

    public static void readExcel() throws IOException {
        FileInputStream fileInputStream = new FileInputStream(filePath);
        BufferedInputStream bufferedInputStream = new BufferedInputStream(fileInputStream);
        POIFSFileSystem fileSystem = new POIFSFileSystem(bufferedInputStream);
        HSSFWorkbook workbook = new HSSFWorkbook(fileSystem);
        HSSFSheet sheet = workbook.getSheet("Sheet1");
        int lastRowIndex = sheet.getLastRowNum();
        System.out.println(lastRowIndex);
        for (int i = 0; i <= lastRowIndex; i++) {
            HSSFRow row = sheet.getRow(i);
            if (row == null) {
                break;
            }
            short lastCellNum = row.getLastCellNum();
            for (int j = 0; j < lastCellNum; j++) {
                HSSFCell cell = row.getCell(j);
                cell.setCellType(CellType.STRING);
                String cellValue = cell.getStringCellValue();
                System.out.println(cellValue);
            }
        }
        bufferedInputStream.close();
    }

}
