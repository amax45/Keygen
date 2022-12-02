import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;

public class main {
    public static void main(String[] args) throws IOException {
        ArrayList<String> arr = new ArrayList<>();
        String path = "dictionary.xlsx";
        FileInputStream fileInputStream = new FileInputStream(path);
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
        int rows = sheet.getLastRowNum();
        int cells = sheet.getRow(1).getLastCellNum();

        for (int r = 2; r <= rows; r++) {
            XSSFRow row = sheet.getRow(r);
            for (int c = 3; c <= cells - 1; c++) {
                XSSFCell cell = row.getCell(c);
                String value = cell.getStringCellValue();
                if (!value.equals("")) {
                    arr.add(value);
                }
            }
        }
        for (int i = 0; i < arr.size(); i++) {
            System.out.println(".addKeyword(\""+arr.get(i)+"\")");
        }
    }
}
