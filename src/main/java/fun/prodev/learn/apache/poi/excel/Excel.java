package fun.prodev.learn.apache.poi.excel;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;
import java.io.*;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @author prodev
 * @date 2019/5/12 20:55
 * @description POI 操作 Excel 文件
 */
@RestController
@RequestMapping("excel")
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

    @RequestMapping("export")
    public void exportExcel(HttpServletResponse response, HttpSession session, String name) throws IOException {
        String[] tableHeaders = {"id", "姓名", "年龄"};
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("Sheet1");
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        Font font = workbook.createFont();
        font.setBold(true);
        cellStyle.setFont(font);

        // 将第一行的三个单元格给合并
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 2));
        HSSFRow row = sheet.createRow(0);
        HSSFCell beginCell = row.createCell(0);
        beginCell.setCellValue("通讯录");
        beginCell.setCellStyle(cellStyle);

        row = sheet.createRow(1);
        // 创建表头
        for (int i = 0; i < tableHeaders.length; i++) {
            HSSFCell cell = row.createCell(i);
            cell.setCellValue(tableHeaders[i]);
            cell.setCellStyle(cellStyle);
        }
        List<User> users = new ArrayList<>();
        users.add(new User("1", "20", "张三"));
        users.add(new User("2", "22", "李四"));
        users.add(new User("3", "21", "王五"));
        for (int i = 0; i < users.size(); i++) {
            row = sheet.createRow(i + 2);
            User user = users.get(i);
            row.createCell(0).setCellValue(user.getId());
            row.createCell(1).setCellValue(user.getName());
            row.createCell(2).setCellValue(user.getAge());
        }
        OutputStream outputStream = response.getOutputStream();
        response.reset();
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-disposition", "attachment;filename=template.xls");
        workbook.write(outputStream);
        outputStream.flush();
        outputStream.close();
    }

    @RequestMapping("import")
    public void importExcel(@RequestParam("file") MultipartFile file) throws IOException {
        InputStream inputStream = file.getInputStream();
        BufferedInputStream bufferedInputStream = new BufferedInputStream(inputStream);
        POIFSFileSystem fileSystem = new POIFSFileSystem(bufferedInputStream);
        HSSFWorkbook workbook = new HSSFWorkbook(fileSystem);

        HSSFSheet sheet = workbook.getSheetAt(0);
        int lastRowNum = sheet.getLastRowNum();
        for (int i = 2; i <= lastRowNum; i++) {
            HSSFRow row = sheet.getRow(i);
            String id = row.getCell(0).getStringCellValue();
            String name = row.getCell(1).getStringCellValue();
            String age = row.getCell(2).getStringCellValue();
            System.out.println(id + "-" + name + "-" + age);
        }
    }

    class User {
        String id;
        String age;
        String name;

        public User() {
        }

        public User(String id, String age, String name) {
            this.id = id;
            this.age = age;
            this.name = name;
        }

        public String getId() {
            return id;
        }

        public void setId(String id) {
            this.id = id;
        }

        public String getAge() {
            return age;
        }

        public void setAge(String age) {
            this.age = age;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        @Override
        public String toString() {
            return "User{" +
                    "id='" + id + '\'' +
                    ", age='" + age + '\'' +
                    ", name='" + name + '\'' +
                    '}';
        }
    }

}
