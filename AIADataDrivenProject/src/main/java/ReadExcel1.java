import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

//Testing out reading excel file using Apache POI
public class ReadExcel1 {
public static  void main(String args[]) throws IOException {
        
        //Create an object of File class to open xlsx file
        File file =    new File("E:\\AIA documents\\AIA testdata excelsheet.xlsx");
        
        //Create an object of FileInputStream class to read excel file
        FileInputStream inputStream = new FileInputStream(file);
        
        //Creating workbook instance that refers to .xls file
        XSSFWorkbook wb=new XSSFWorkbook(inputStream);
        
        //Creating a Sheet object using the sheet Name
        XSSFSheet sheet=wb.getSheet("TESTCASE_DATA");
        
        //Create a row object to retrieve row at index 2
        XSSFRow row3=sheet.getRow(2);
        
        //Create a cell object to retreive cell at index 1
        XSSFCell cell=row3.getCell(1);
        
        //Get the address in a variable
        String username= cell.getStringCellValue();
        
        //Printing the username
        System.out.println("The username is : "+ username);
        
        ///////////////////////////////////////////////////////////
        System.out.println("");
        System.out.println("/////////////NEW SECTION/////////");
        System.out.println("");
        
        //get all rows in the sheet
        int rowCount=sheet.getLastRowNum()-sheet.getFirstRowNum();
        
        //iterate over all the row to print the data present in each cell.
        for(int i=0;i<=rowCount;i++){
            
            //get cell count in a row
            int cellcount=sheet.getRow(i).getLastCellNum();
            
            //iterate over each cell to print its value
            System.out.println("Row"+ i+" data is :");
            
            for(int j=0;j<cellcount;j++){
                System.out.print(sheet.getRow(i).getCell(j).getStringCellValue() +",");
            }
            System.out.println();
        }
	}
}
