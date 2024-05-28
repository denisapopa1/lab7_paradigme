package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class Main {
    public static void main(String[] args) {
        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet sheet = workbook.createSheet("Employee Data");

        Map<String, Object[]> data = new TreeMap<String, Object[]>();
        data.put("1", new Object[]{"Name", "Surname", "Grade 1", "Grade 2", "Grade 3", "Grade 4","Max","Average"});
        data.put("2", new Object[]{"Amit", "Shukla", 9, 8, 7, 5});
        data.put("3", new Object[]{"Lokesh", "Gupta", 8, 9, 6, 7});
        data.put("4", new Object[]{"John", "Adwards", 8, 8, 7, 6});
        data.put("5", new Object[]{"Brian", "Schultz", 7, 6, 8, 9});



        Set<String> keyset = data.keySet();
        int rownum = 0;
        for (String key : keyset) {
            Row row = sheet.createRow(rownum++);
            Object[] objArr = data.get(key);
            int cellnum = 0;
            int max = 0;
            int sum = 0;
            int numGrades = 0;
            for (Object obj : objArr) {
                Cell cell = row.createCell(cellnum++);
                if (obj instanceof String || obj instanceof Integer) {
                    cell.setCellValue(obj.toString());
                    if(rownum==1){
                        CellStyle style= workbook.createCellStyle();
                        Font font=workbook.createFont();
                        font.setBold(true);
                        style.setFont(font);
                        cell.setCellStyle(style);
                    }
                    if (obj instanceof Integer) {
                        int grade = (Integer) obj;
                        max = Math.max(max, grade);
                        sum += grade;
                        numGrades++;
                    }
                }
            }
            Cell maxCell = row.createCell(cellnum++);
            maxCell.setCellValue(max);

            Cell avgCell = row.createCell(cellnum++);
            avgCell.setCellValue(numGrades > 0 ? (double) sum / numGrades : 0);

            if(rownum>1){
                maxCell.setCellStyle(getBackgroundStyle(workbook));
                avgCell.setCellStyle(getBackgroundStyle(workbook));
            }
        }


        try {
            FileOutputStream out = new FileOutputStream(new File("output.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("output.xlsx written successfully on disk.");
        } catch (Exception e) {
            e.printStackTrace();
        }


    }

    private static CellStyle getBackgroundStyle(XSSFWorkbook workbook) {
        CellStyle style=workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.PINK.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }
}