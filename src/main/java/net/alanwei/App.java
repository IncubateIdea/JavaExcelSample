package net.alanwei;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 */
public class App {
    public static void main(String[] args) throws Exception {
        FileInputStream file = new FileInputStream(new File(System.getProperty("user.dir") + "\\src\\main\\resources\\sample.xlsx"));
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);

        //读取每一行的数据, 存到集合里
        List<ExcelData> datas = new ArrayList<ExcelData>();
        for (Row row : sheet) {
            //遍历行
            ExcelData data = new ExcelData();
            data.setColumnA(row.getCell(0).getStringCellValue()); //第一个单元格的值
            data.setColumnB(row.getCell(1).getStringCellValue()); //第二个单元格的值
            datas.add(data);
        }

        System.out.println("总共读取了" + datas.size() + "行数据");
        //比较数据是否一致
        int index = 1;
        for (ExcelData data : datas) {
            if (data.getColumnA().equals(data.getColumnB())) {
                System.out.println("第" + index + "行的值相等");
            } else {
                System.out.println("第" + index + "行的值不相等: A: " + data.getColumnA() + ", B: " + data.getColumnB());
            }
            ++index;
        }
    }
}
