package net.alanwei;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;

/**
 * Hello world!
 */
public class App {
    public static void main(String[] args) throws Exception {
        FileInputStream file = new FileInputStream(new File("E:\\扬子石化总部端2018年盘查项目 - Copy.xls"));
        Workbook workbook = new HSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);

        //读取每一行的数据, 存到集合里
        List<ExcelData> datas = new ArrayList<ExcelData>();
        int index = 0;
        for (Row row : sheet) {
            if (index++ == 0) {
                //第一行是列头, 不读取
                continue;
            }
            //遍历行
            ExcelData data = new ExcelData();


            data.setColumnF(getCellValue(row.getCell(5))); //NODE_ID
            data.setColumnG(getCellValue(row.getCell(6))); //NODE_NAME
            data.setColumnH(getCellValue(row.getCell(7))); //NODE_ID
            data.setColumnM(getCellValue(row.getCell(12))); //NODE_NAME

            datas.add(data);

        }


        System.out.println("总共读取了" + datas.size() + "行数据");
        //比较数据是否一致
        index = 2;
        for (ExcelData data : datas) {
            if(data.getColumnG().equals(data.getColumnM())){
                System.out.println("第" + index + "行相同, ID分别为: " + data.getColumnF() + ", " + data.getColumnM());
            }
        }
    }

    static String getCellValue(Cell cell) {
        if(cell == null){
            return "";
        }
        Object cellValue ;
        switch (cell.getCellType()) {
            case NUMERIC:
                cellValue = cell.getNumericCellValue();
                break;
            case BOOLEAN:
                cellValue = cell.getBooleanCellValue();
                break;
            case STRING:
                cellValue = cell.getStringCellValue();
                break;
            default:
                cellValue = "";
        }
        return String.valueOf(cellValue);
    }
}
