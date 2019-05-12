package fun.prodev.learn.apache.poi;

import fun.prodev.learn.apache.poi.excel.Excel;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.IOException;

/**
 * @author prodev
 * @date 2019/5/12 22:13
 * @description
 */
@RunWith(SpringRunner.class)
@SpringBootTest
public class TestPOI {

    @Test
    public void testCreateExcel() throws IOException {
        Excel.createExcel();
        Excel.readExcel();
    }
}
