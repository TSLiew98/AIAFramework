import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;


//Testing out writing into excel file using Apache POI (add new data row)
public class WriteToExcel2 {
    public static  void main(String args[]) throws IOException {
        
    	//Create an object of File class to open xlsx file
        File file =    new File("E:\\AIA documents\\AIA testdata excelsheet.xlsx");
        
        //Create an object of FileInputStream class to read excel file
        FileInputStream inputStream = new FileInputStream(file);
        
        //creating workbook instance that refers to .xls file
        XSSFWorkbook wb=new XSSFWorkbook(inputStream);
        
        //creating a Sheet object using the sheet Name
        XSSFSheet sheet=wb.getSheet("TESTCASE_DATA");
        
        //Create a row object to retrieve row at index 5
        XSSFRow row6=sheet.createRow(5);
        
        //create a cell object to enter value in it using cell Index
        row6.createCell(0).setCellValue("AIA_TC_1");
        row6.createCell(1).setCellValue("Federick Wong");
        row6.createCell(2).setCellValue("Test@123");
        
        //write the data in excel using output stream
        FileOutputStream outputStream = new FileOutputStream("E:\\AIA documents\\AIA testdata excelsheet.xlsx");
        wb.write(outputStream);
        wb.close();

    }
}