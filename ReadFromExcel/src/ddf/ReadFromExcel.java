package ddf;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadFromExcel {

	public static void main(String[] args) throws IOException {
		
		//Locate the file
		//FileInputStream file = new FileInputStream("C:\\Users\\Tenant\\eclipse-workspace\\ReadFromExcel\\input.xlsx");
		
		FileInputStream file = new FileInputStream(new File(System.getProperty("user.dir")+"\\input.xlsx"));

		
		//To initialize workbook   Opening file
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		
		//To initialize Sheet 1
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		//To read data string
		
		
		  String fname = sheet.getRow(0).getCell(0).getStringCellValue();
		  
		  System.out.println(fname); System.out.println("-----------------------");
		 
		
		// to read data numeric
		
		   Number num= sheet.getRow(1).getCell(0).getNumericCellValue();
		   
		   System.out.println(num);
		
		file.close();
		
		
		
		//class obj = new class(parameters);

		
	}

}
