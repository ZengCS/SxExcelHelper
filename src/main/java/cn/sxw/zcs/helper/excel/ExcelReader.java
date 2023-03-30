package cn.sxw.zcs.helper.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @Author ZengCS
 * @Date 2023/3/28 15:57
 * @Copyright 四川生学教育科技有限公司
 * @Address 成都市天府软件园E3-3F
 */
public class ExcelReader {
    private static final String[] IMPORT_TITLES = new String[]{"学生姓名", "学籍号", "年级/班级", "性别", "民族", "完成时间"};
    private static final String[] EXPORT_TITLES = new String[]{"学生姓名", "学籍号", "学校", "年级/班级", "用户Id", "性别", "家长联系电话", "最后作答时间", "用时（分钟）"};
    private static final String[] EXPORT_SHEETS = new String[]{"数据采集名单", "已完成名单", "未完成名单", "异常已完成", "异常未完成"};

    public List<StudentImportBean> stuList = new ArrayList<>();
    public List<StudentExportBean> completeList = new ArrayList<>();
    public List<StudentExportBean> unCompleteList = new ArrayList<>();
    public List<StudentExportBean> errorCompleteList = new ArrayList<>();
    public List<StudentExportBean> errorUnCompleteList = new ArrayList<>();


    private File file1;
    private File file2;

    public ExcelReader(File file1, File file2) {
        this.file1 = file1;
        this.file2 = file2;
    }

    public File analysis() throws Exception {
        this.readBaseExcel();
        this.readOutExcel();
        return this.mergeExcels();
    }

//    public static void main(String[] args) {
//        try {
//            File f1 = new File("/Users/zengcs/Documents/华西心理促进/宝石小学-信息采集模板（4-6学生）.xlsx");
//            File f2 = new File("/Users/zengcs/Documents/华西心理促进/宝石小学已完成.xls");
//            ExcelReader excelReader = new ExcelReader(f1, f2);
//            excelReader.readBaseExcel();
//            excelReader.readOutExcel();
//            File file = excelReader.mergeExcels();
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
//    }

    private XSSFCellStyle makeStyle(Workbook workbook, boolean isTitle, boolean isState) {
        XSSFCellStyle cellStyle = (XSSFCellStyle) workbook.createCellStyle();
        if (isTitle) {
            // 设置单元格颜色
            cellStyle.setFillForegroundColor(HSSFColor.GREY_40_PERCENT.index);
            // 设定填充单色
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        // 边框
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);

        // 设置水平居中
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        // 设置垂直居中
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        // 字体
        XSSFFont font = (XSSFFont) workbook.createFont();
        // 2代表红色字体
        if (isState) {
            font.setColor(HSSFColor.GREEN.index);
        }

        // 字体加粗
        font.setBold(isTitle || isState);
        font.setFontHeightInPoints((short) (isTitle ? 14 : 12));
        cellStyle.setFont(font);
        return cellStyle;
    }

    private File mergeExcels() throws Exception {
        String path = file1.getParentFile().getAbsolutePath();
        // 创建一个工作簿07
        Workbook workbook = new XSSFWorkbook();
        XSSFCellStyle titleStyle = makeStyle(workbook, true, false);
        XSSFCellStyle cellStyle = makeStyle(workbook, false, false);
        XSSFCellStyle stateStyle = makeStyle(workbook, false, true);

        int[] counts = new int[]{stuList.size(), completeList.size(), unCompleteList.size(), errorCompleteList.size(), errorUnCompleteList.size()};

        int index = 0;
        for (String name : EXPORT_SHEETS) {
            // 创建一个工作表
            Sheet sheet = workbook.createSheet(name + "(" + counts[index] + "人)");
            // 只冻结第一行
            sheet.createFreezePane(0, 1, 0, 1);
            sheet.setColumnWidth(0, 256 * 20);
            sheet.setColumnWidth(1, 256 * 30);
            if (index == 0) {
                sheet.setColumnWidth(2, 256 * 20);
                sheet.setColumnWidth(3, 256 * 12);
                sheet.setColumnWidth(4, 256 * 12);
                sheet.setColumnWidth(5, 256 * 30);
            } else {
                sheet.setColumnWidth(2, 256 * 30);
                sheet.setColumnWidth(3, 256 * 20);
                sheet.setColumnWidth(4, 256 * 12);
                sheet.setColumnWidth(5, 256 * 12);
                sheet.setColumnWidth(6, 256 * 20);
                sheet.setColumnWidth(7, 256 * 30);
                sheet.setColumnWidth(8, 256 * 20);
            }


            int c = 0;
            Row titleRow = sheet.createRow(0);

            if (index == 0) {
                for (String title : IMPORT_TITLES) {
                    Cell cell = titleRow.createCell(c++);
                    cell.setCellValue(title);
                    cell.setCellStyle(titleStyle);
                }
                int rowIndex = 1;
                for (StudentImportBean stu : stuList) {
                    // 判断他是否已完成
                    boolean complete = false;
                    String completeTime = "";
                    for (StudentExportBean cstu : completeList) {
                        if (cstu.getCode().equals(stu.getCode())) {
                            complete = true;
                            completeTime = cstu.getLastAnswerTime();
                            break;
                        }
                    }

                    Row row = sheet.createRow(rowIndex++);
                    int j = 0;
                    for (String attr : stu.getAttributes()) {
                        Cell cell = row.createCell(j++);
                        cell.setCellValue(attr);
                        cell.setCellStyle(cellStyle);
                    }
                    Cell cell = row.createCell(j);
                    cell.setCellValue(completeTime);
                    if (complete) {
                        cell.setCellStyle(stateStyle);
                    } else {
                        cell.setCellStyle(cellStyle);
                    }
                }
            } else {

                for (String title : EXPORT_TITLES) {
                    Cell cell = titleRow.createCell(c++);
                    cell.setCellValue(title);
                    cell.setCellStyle(titleStyle);
                }
                int rowIndex = 1;
                List<StudentExportBean> targetList;
                if (index == 1) {
                    targetList = completeList;
                } else if (index == 2) {
                    targetList = unCompleteList;
                } else if (index == 3) {
                    targetList = errorCompleteList;
                } else {
                    targetList = errorUnCompleteList;
                }
                for (StudentExportBean stu : targetList) {
                    Row row = sheet.createRow(rowIndex++);
                    int j = 0;
                    for (String attr : stu.getAttributes()) {
                        Cell cell = row.createCell(j++);
                        cell.setCellValue(attr);
                        cell.setCellStyle(cellStyle);
                    }
                }
            }
            index++;
        }

        //创建一个工作表 07版本使用xlsx结尾！
        SimpleDateFormat sdf = new SimpleDateFormat("MMdd_HHmmss", Locale.getDefault());
        String currDateTime = sdf.format(new Date());
        String fileName = path + "/" + file2.getName().split("\\.")[0] + "_数据分析_" + currDateTime + ".xlsx";
        FileOutputStream fileOutputStream = new FileOutputStream(fileName);
        workbook.write(fileOutputStream);
        //关闭流
        fileOutputStream.close();
        System.out.println("excel表生成完毕 --> " + fileName);
        return new File(fileName);
    }

    public void readBaseExcel() throws Exception {
        String fileName = file1.getName();
        if (!fileName.endsWith(".xlsx") && !fileName.endsWith(".xls")) {
            throw new IllegalArgumentException("请勿上传非Excel文件~");
        }

        FileInputStream fis = new FileInputStream(file1.getAbsolutePath());
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
            Cell cell1 = row.getCell(1);
            if (cell1 != null) {
                try {
                    student.setClassNum((int) cell1.getNumericCellValue());
                } catch (Exception exception) {
                    student.setClassNum(Integer.parseInt(cell1.getStringCellValue().trim()));
                }
            } else {
                throw new IllegalArgumentException("文件【" + file1.getName() + "】\n中存在班级为空的错误数据~");
            }
            if (student.getClassNum() <= 0) {
                throw new IllegalArgumentException("文件【" + file1.getName() + "】\n中存在班级为空的错误数据~");
            }

            student.setName(row.getCell(2).getStringCellValue());
            student.setGender(row.getCell(3).getStringCellValue());
            student.setCode(row.getCell(4).getStringCellValue());
            student.setNation(row.getCell(5).getStringCellValue());
        }

        fis.close();
    }

    public void readOutExcel() throws Exception {
        String fileName = file2.getName();
        if (!fileName.endsWith(".xlsx") && !fileName.endsWith(".xls")) {
            throw new IllegalArgumentException("请勿上传非Excel文件~");
        }

        FileInputStream fis = new FileInputStream(file2.getAbsolutePath());
        // 获取工作簿
        Workbook workbook = fileName.endsWith(".xlsx") ? new XSSFWorkbook(fis) : new HSSFWorkbook(fis);
        // 获取第1张表
        Sheet sheet1 = workbook.getSheetAt(0);
        Iterator<Row> rowIterator1 = sheet1.rowIterator();
        rowIterator1.next();
        while (rowIterator1.hasNext()) {
            StudentExportBean student = analysisRow(rowIterator1.next());
            // 判断用户是否在导入表中
            boolean isNormal = false;
            for (StudentImportBean stu : stuList) {
                if (stu.getCode().equals(student.getCode())) {
                    isNormal = true;
                    break;
                }
            }
            if (isNormal) {
                completeList.add(student);
            } else {
                errorCompleteList.add(student);
            }
        }
        // 获取第2张表
        Sheet sheet2 = workbook.getSheetAt(1);
        Iterator<Row> rowIterator2 = sheet2.rowIterator();
        rowIterator2.next();
        while (rowIterator2.hasNext()) {
            StudentExportBean student = analysisRow(rowIterator2.next());
            // 判断用户是否在导入表中
            boolean isNormal = false;
            for (StudentImportBean stu : stuList) {
                if (stu.getCode().equals(student.getCode())) {
                    isNormal = true;
                    break;
                }
            }
            if (isNormal) {
                unCompleteList.add(student);
            } else {
                errorUnCompleteList.add(student);
            }
        }

        fis.close();
    }

    private StudentExportBean analysisRow(Row row) {
        StudentExportBean student = new StudentExportBean();
        Cell c0 = row.getCell(0);
        student.setCode(c0 != null ? c0.getStringCellValue() : "");

        Cell c1 = row.getCell(1);
        student.setSchool(c1 != null ? c1.getStringCellValue() : "");

        Cell c2 = row.getCell(2);
        student.setGradeClass(c2 != null ? c2.getStringCellValue() : "");

        Cell c3 = row.getCell(3);
        student.setUserId(c3 != null ? c3.getStringCellValue() : "");

        Cell c4 = row.getCell(4);
        student.setName(c4 != null ? c4.getStringCellValue() : "");

        Cell c5 = row.getCell(5);
        student.setGender(c5 != null ? c5.getStringCellValue() : "");

        Cell c6 = row.getCell(6);
        student.setParentPhone(c6 != null ? c6.getStringCellValue() : "");

        Cell c7 = row.getCell(7);
        student.setLastAnswerTime(c7 != null ? c7.getStringCellValue() : "");

        Cell c8 = row.getCell(8);
        student.setUseTime(c8 != null ? c8.getStringCellValue() : "");
        return student;
    }
}
