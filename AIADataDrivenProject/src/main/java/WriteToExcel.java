import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

//Testing out writing into excel file using Apache POI
public class WriteToExcel {

	public static  void main(String args[]) throws IOException {
		
		File src= new File("E:\\AIA documents\\AIA testdata excelsheet.xlsx");
		
		FileInputStream fis=new FileInputStream(src);
		
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		
		XSSFSheet sheet1=wb.getSheetAt(0);
		
		sheet1.getRow(1).createCell(3).setCellValue("Pass");
		
		sheet1.getRow(2).createCell(3).setCellValue("Fail");
		
		FileOutputStream fout=new FileOutputStream(src);
		
		wb.write(fout);
		
		fis.close();
		wb.close();
		
	}

}
