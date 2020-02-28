package javaapplication1;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class JavaApplication1 {
    final static String EXCEL_FILE_NAME = "test.xlsx";

    // Запись данных в XLSX-файл
    public static void writeXLSXFile() throws IOException {
        XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(EXCEL_FILE_NAME));
        XSSFSheet sheet = wb.getSheetAt(0);
        for (int r = 0; r < 3; r++) {
            XSSFRow row = sheet.createRow(r+9);
            for (int c = 0; c < 3; c++) {
                row.createCell(c + 8).setCellValue(String.valueOf(2*(r+c)));
            }
        }
        FileOutputStream fos = new FileOutputStream(EXCEL_FILE_NAME);
        wb.write(fos); fos.flush(); fos.close();
    }

     // Чтение данных из XLSX-файла
   public static void readXLSXFile() throws IOException {
        InputStream ExcelFileToRead = new FileInputStream(EXCEL_FILE_NAME);
        XSSFSheet sheet = new XSSFWorkbook(ExcelFileToRead).getSheetAt(0);
        XSSFRow row; XSSFCell cell;
        // Считывание текстовых и цифровых данных из файла
        Iterator rows = sheet.rowIterator();
        while (rows.hasNext()) {
            row = (XSSFRow) rows.next();
            Iterator cells = row.cellIterator();
            while (cells.hasNext()) {
                cell = (XSSFCell) cells.next();
                if (cell.getCellType() == CellType.STRING) {
                    System.out.print(cell.getStringCellValue() + " ");
                } else if (cell.getCellType() == CellType.NUMERIC) {
                    System.out.print(cell.getNumericCellValue() + " ");
                } else {
                }
            }
            System.out.println();
        }
    }

    public static void main(String[] args) {
        try {
            writeXLSXFile();
            readXLSXFile();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
