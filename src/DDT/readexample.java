package DDT;


import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class readexample {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		 FileInputStream file = new FileInputStream(System.getProperty("user.dir")+"\\testdata\\caldata.xlsx");
		    //FileInputStream file=new FileInputStream(System.getProperty("user.dir")+"\\testdata\\data.xlsx");
		 XSSFWorkbook  w = new XSSFWorkbook(file) ;
		 XSSFSheet sheet = w.getSheet("Sheet1");
		 
		 int totalrows = sheet.getLastRowNum();
		 int totalcell = sheet.getRow(1).getLastCellNum();
		 
		 System.out.println("Total Rows   :"+ totalrows);
		 System.out.println("Total cells  :"+  totalcell);
		 
		 for(int r=0 ; r <= totalrows ; r++) {
			 
			 XSSFRow CurrentRow = sheet.getRow(r);
			 
			 for(int c=0 ; c <totalcell ; c++)
			 
			 {
				 String value = CurrentRow.getCell(c).toString();
				 System.out.print(value+"      ");
			 }
		 }
		 
		
		

	}

}
