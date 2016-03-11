package TempFramework;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.Test;

public class DataDrivenUsingExcel {
    WebDriver driver =new FirefoxDriver();

	//public static void main(String[] args) throws Exception {
	@Test
	public void ExcelTestng(){ 
		
    driver.get("https://www.google.com/?gws_rd=ssl");
    driver.manage().window().maximize();
    WebElement searchbox = driver.findElement(By.name("q"));
    try{ 
    
    FileInputStream f = new FileInputStream(new File("C:\\SEProjects\\TestNGExamples\\src\\TestNGSamples\\Test1.xlsx"));
    XSSFWorkbook wb= new XSSFWorkbook(f);
    XSSFSheet sh1 = wb.getSheetAt(0);
    
    for(int i=1; i<=sh1.getLastRowNum(); i++){
    	String keyword = sh1.getRow(i).getCell(0).getStringCellValue();
    	searchbox.sendKeys(keyword);
    	searchbox.submit();
    	driver.manage().timeouts().implicitlyWait(100000,TimeUnit.MILLISECONDS);
    }	
wb.close();
f.close();
	}
	catch (FileNotFoundException fnfe){
		
		fnfe.printStackTrace();
		
	}
	catch(IOException ioe){
		
		ioe.printStackTrace();
	}
}}

	/*catch (FileNotFoundException fnfe) {
    	fnfe.printStackTrace();
    	
    }
	 catch (IOException ioe) {
    	
    	ioe.printStackTrace();
    }
    

	}

}*/
