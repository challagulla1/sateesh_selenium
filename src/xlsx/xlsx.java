package xlsx;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class xlsx {
 public static void main(String[] args) 
 {
	 XSSFWorkbook workbook = new XSSFWorkbook();
	try {
		workbook = new XSSFWorkbook(new FileInputStream(new File("C:\\Users\\KEETECH\\Downloads\\Keetech Employee Tracker123 sateesh1123 (1).xlsx")));
	  } catch (FileNotFoundException e) {
		
		e.printStackTrace();
	} catch (IOException e) {
		
		e.printStackTrace();
	}
	 XSSFSheet sheet = workbook.getSheetAt(0);
	 
			 int rows = sheet.getLastRowNum();
	 for(int i=0;i<rows;i++)
	 {
		 XSSFRow row = sheet.getRow(i);
		 int cellscount = row.getLastCellNum();
		 for(int j=0;j<cellscount;j++)
		 {
			 XSSFCell cell = row.getCell(j);
			 System.out.println("data in row , cell:"+i+","+j+"is : "+cell.getStringCellValue());
		 }
		 
	 }
		 
		 try {
			workbook.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		 
		 
		 
 }
}

