package utils;
import java.io.FileInputStream;
	import java.io.FileOutputStream;
	import java.util.Properties;

	import org.apache.poi.ss.usermodel.ExcelStyleDateFormatter;
	import org.apache.poi.xssf.usermodel.XSSFCell;
	import org.apache.poi.xssf.usermodel.XSSFRow;
	import org.apache.poi.xssf.usermodel.XSSFSheet;
	import org.apache.poi.xssf.usermodel.XSSFWorkbook;

	public class ExcelUtils {

		public static XSSFSheet ExcelWSheet;

		public static XSSFWorkbook ExcelWBook;

		public static XSSFCell Cell;

		public static XSSFRow Row;

	//This method is to set the File path and to open the Excel file, Pass Excel Path and Sheetname as Arguments to this method

	public static void setExcelFile(String Path,String SheetName) throws Exception {

	try {

			// Open the Excel 		
			FileInputStream ExcelFile = new FileInputStream(Path);
			//Access the required sheet
			ExcelWBook = new XSSFWorkbook(ExcelFile);
			ExcelWSheet = ExcelWBook.getSheet(SheetName);

	} catch (Exception e){
			throw (e);
	}

	}

	public static int getRowCount(String sheetName){
		int index = ExcelWBook.getSheetIndex(sheetName);
		ExcelWSheet = ExcelWBook.getSheetAt(index);
		int number=ExcelWSheet.getLastRowNum()+1;
		
		return number;
	}}

