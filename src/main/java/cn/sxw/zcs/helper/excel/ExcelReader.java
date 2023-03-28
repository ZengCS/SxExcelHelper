package cn.sxw.zcs.helper.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;

/**
 * @Author ZengCS
 * @Date 2023/3/28 15:57
 * @Copyright 四川生学教育科技有限公司
 * @Address 成都市天府软件园E3-3F
 */
public class ExcelReader {
    private static ArrayList<StudentImportBean> stuList = new ArrayList<>();

    public static void main(String[] args) {
        try {
            new ExcelReader(new File("/Users/zengcs/Documents/华西心理促进/宝石小学-信息采集模板（4-6学生）.xlsx"));

            for (StudentImportBean stu : stuList) {
                System.out.println(stu.toString());
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    public ExcelReader(File file) throws Exception {
        String fileName = file.getName();
        if (!fileName.endsWith(".xlsx") && !fileName.endsWith(".xls")) {
            throw new IllegalArgumentException("请勿上传非Excel文件~");
        }

        FileInputStream fis = new FileInputStream(file.getAbsolutePath());
        // 获取工作簿
        Workbook workbook = fileName.endsWith(".xlsx") ? new XSSFWorkbook(fis) : new HSSFWorkbook(fis);
        // 获取第2张表
        Sheet sheetAt = workbook.getSheetAt(1);
        Iterator<Row> rowIterator = sheetAt.rowIterator();
        rowIterator.next();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            StudentImportBean student = new StudentImportBean();
            stuList.add(student);

            student.setGradeName(row.getCell(0).getStringCellValue());
            student.setClassNum((int) row.getCell(1).getNumericCellValue());
            student.setName(row.getCell(2).getStringCellValue());
            student.setGender(row.getCell(3).getStringCellValue());
            student.setCode(row.getCell(4).getStringCellValue());
            student.setNation(row.getCell(5).getStringCellValue());
        }

        fis.close();
    }
}
