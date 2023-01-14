package DDT;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingExcel2 {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		FileInputStream file = new FileInputStream(System.getProperty("user.dir")+"\\testdata\\Book1.xlsx");
		
		XSSFWorkbook Workbook = new XSSFWorkbook(file);
		
		XSSFSheet sheet = Workbook.getSheet("Sheet1");
		
		int totalrows    = sheet.getLastRowNum();
		int totalcolumns = sheet.getRow(1).getLastCellNum();
		
		
		System.out.println( totalrows);
		System.out.println(totalcolumns );
		
		
		 	}

}
