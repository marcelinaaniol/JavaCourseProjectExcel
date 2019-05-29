package Excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class App {

    private static String[] columns = {"TIME", "SUBJECT DESCRIPTION", "DAY OF WEEK", "NUMBER OF HOURS", "DATE"};
    private static List<CourseDetails> courseDetails = new ArrayList<>();


    static {
        courseDetails.add(new CourseDetails("18:30 - 19:00", "Introduction to Java", "Tuesday", 3, "02/04/2019"));
        courseDetails.add(new CourseDetails("18:30 - 19:00", "Classes and Objects", "Thursday", 3, "04/04/2019"));
        courseDetails.add(new CourseDetails("18:30 - 19:00", "Inheritance", "Tuesday", 3, "09/04/2019"));
        courseDetails.add(new CourseDetails("18:30 - 19:00", "Abstract Class and Polymorphism", "Thursday", 3, "11/04/2019"));
        courseDetails.add(new CourseDetails("18:30 - 19:00", "Interface", "Tuesday", 3, "16/04/2019"));
        courseDetails.add(new CourseDetails("18:30 - 19:00", "Generics", "Thursday", 3, "18/04/2019"));
        courseDetails.add(new CourseDetails("18:30 - 19:00", "Solid Principles", "Tuesday", 3, "23/04/2019"));
        courseDetails.add(new CourseDetails("18:30 - 19:00", "Maven", "Thursday", 3, "25/04/2019"));
        courseDetails.add(new CourseDetails("18:30 - 19:00", "Summary", "Tuesday", 3, "30/04/2019"));
    }


    public static void main(String[] args) throws IOException {

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Java Course Schedule");
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 16);
        headerFont.setColor(IndexedColors.BLACK.getIndex());

        {
            Row row = sheet.createRow(0);
            Cell cell = row.createCell(0);
            cell.setCellValue("JAVA COURSE SCHEDULE");
            Font cellFont = workbook.createFont();
            cellFont.setBold(true);
            cellFont.setFontHeightInPoints((short) 16);
            cellFont.setColor(IndexedColors.GREEN.getIndex());
            CellStyle headerCellStyle = workbook.createCellStyle();
            headerCellStyle.setFont(headerFont);
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setFont(cellFont);
            cell.setCellStyle(cellStyle);
        }

        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        Row headerRow = sheet.createRow(2); //wiersz naglowka

        for (int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellStyle(headerCellStyle);
            cell.setCellValue(columns[i]);
        }

        int rowNumber = 3;

        for (CourseDetails courseDetails : courseDetails) {
            Row row = sheet.createRow(rowNumber++);
            row.createCell(0).setCellValue(courseDetails.TIME);
            row.createCell(1).setCellValue(courseDetails.SUBJECTDESCRIPTION);
            row.createCell(2).setCellValue(courseDetails.DAYOFWEEK);
            row.createCell(3).setCellValue(courseDetails.NUMBEROFHOURS);
            row.createCell(4).setCellValue(courseDetails.DATE);
        }

        for (int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }

        int rowCount = 6;
        Row rowTotal = sheet.createRow(rowCount + 7);
        Cell cellTotalText = rowTotal.createCell(2);
        cellTotalText.setCellValue("Total:");
        Cell cellTotal = rowTotal.createCell(3);
        cellTotal.setCellFormula("SUM(D4:D12)");


        FileOutputStream fileOut = new FileOutputStream("Java Course.xlsx");
        workbook.write(fileOut);
        fileOut.close();
        workbook.close();

    }


}