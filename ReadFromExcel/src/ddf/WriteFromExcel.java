package ddf;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteFromExcel {

	public static void main(String[] args) throws IOException {
		FileInputStream file = new FileInputStream(new File(System.getProperty("user.dir")+"\\input.xlsx"));
		
		//To initialize workbook   Opening file
				XSSFWorkbook workbook = new XSSFWorkbook(file);
				
				//To initialize Sheet 1
				XSSFSheet sheet = workbook.getSheetAt(0);
				
				// cell allocation
				Cell write = sheet.getRow(0).getCell(0);// cell is allocated to update data
				
				// value set to particar Cell
				
				write.setCellValue("Techbodhi");
				
				//Read data
				String fname = sheet.getRow(0).getCell(0).getStringCellValue();// to read data from particalr cell
				
				System.out.println(fname);
				
				// give the path of Output file 
				
				FileOutputStream filewrite = new FileOutputStream(new File(System.getProperty("user.dir")+"\\output.xlsx"));
				
				workbook.write(filewrite);
		
				filewrite.close();
				
				
				
	}

}
