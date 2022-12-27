package ApachePOI;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;

public class _10_Soru_CarpimTablosu1 {
    public static void main(String[] args) throws IOException {

/** Soru 1:
 *  yeni excel
 *  Çarpım tablosunu excele yazdırınız.
 *  1 x 1 = 1 şeklinde işaretleri de yazdırınız.
 *  sıfırdan excel oluşturarak.
 *  her bir onluktan sonra 1 satır boşluk bırakarak alt alta
 */
        // MULTIPLICATION TABLE - VERSION - I

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Sheet1");

        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("Table X");

        int count = 1;
        for (int i = 1; i <= 10; i++) {
            for (int j = 1; j <= 10; j++) {
                row = sheet.createRow(count);
                cell = row.createCell(0);
                cell.setCellValue(i + "x" + j + " =" + (i * j));
                count++;
            }
            sheet.createRow(count);
            row.createCell(0);
            cell.setCellValue("     ");

        }
        String newPath = "src/test/java/ApachePOI/resource/ExcelTableX.xlsx";

        try {
            FileOutputStream outputStream = new FileOutputStream(newPath);
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();
        } catch (IOException ex) {
            throw new RuntimeException(ex);
        }

        System.out.println("Version I is ready!");
    }


    /*
    public static void changeCellBackgroundColorWithPattern(Cell cell) {
        CellStyle cellStyle = cell.getCellStyle();
        if(cellStyle == null) {
            cellStyle = cell.getSheet().getWorkbook().createCellStyle();
        }
        cellStyle.setFillBackgroundColor(IndexedColors.BLUE.getIndex());
        cellStyle.setFillPattern(FillPatternType.BIG_SPOTS);
        cellStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        cell.setCellStyle(cellStyle);
    }   */



        }










