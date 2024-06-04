package test;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class ManipulateExcelFile
{
    public static void main(String[] args)
    {
        try
        {
            FileInputStream fileInputStream =
                    new FileInputStream("C:\\Users\\Acer\\IdeaProjects\\selenium\\src\\test\\resources\\Task 13.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            Sheet sheet = workbook.getSheet("Sheet1");
            Map<String, Map<String,String>> testdata = new HashMap<>();
            Row headerRow = sheet.getRow(0);
            for (int i = 1; i < sheet.getLastRowNum(); i++)
            {
                Row row = sheet.getRow(i);
                Map<String,String> colValues = new HashMap<>();
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    colValues.put(headerRow.getCell(j).getStringCellValue(), row.getCell(j).getStringCellValue());
                }
                    System.out.println(colValues);
            }
        workbook.close();
        }catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}

