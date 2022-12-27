package ApachePOI;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class _11_Soru_CarpimTablosu2 {
    public static void main(String[] args) throws IOException {

/** Soru 1:
 *  yeni excel
 *  Çarpım tablosunu excele yazdırınız.
 *  1 x 1 = 1 şeklinde işaretleri de yazdırınız.
 *  sıfırdan excel oluşturarak.
 *  her bir onluktan sonra 1 satır boşluk bırakarak yanyana
 */
        // MULTIPLICATION TABLE - VERSION - 2

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Sheet1");

        for (int i = 1; i <= 10; i++) {

            Row row = sheet.createRow(i - 1);
            int cellCount=row.getPhysicalNumberOfCells();

            for (int j = 1; j <= 10; j++) {

                Cell cell = row.createCell(cellCount);
                cellCount+=2;
                cell.setCellValue(i + "x" + j + " =" + (i * j) );
            }
        }
        FileOutputStream outputStream = new FileOutputStream("src/test/java/ApachePOI/resource/ExcelTable2.xlsx");
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();    }
}












