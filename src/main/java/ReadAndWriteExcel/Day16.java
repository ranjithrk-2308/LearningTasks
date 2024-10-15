package ReadAndWriteExcel;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.*;

public class Day16 {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		XSSFWorkbook book = new XSSFWorkbook();

		//Create Sheet
		XSSFSheet sheet = book.createSheet("Details");
		
		// Store the details -> Name (String) Age(Int) City (String)
		Object[][] data = {
				{"Name", "Age", "City"},
				{"Ranjith",	"25"	,"Sivakasi"},
				{"Vicky", 	"25",	"RS Mangalam"},
				{"Kamal",	"22"	,"Ambai"}
		};
		// Put the data into the sheet. And COunt the data as well.
		int rowCOunt = 0;
		
		// Get into each rows using for each iteration.
		for (Object[] row1: data ) {
		
			XSSFRow row  = sheet.createRow(rowCOunt++);
			int columncount = 0;
			
			//Each data of the row for For each to get columns:
			
			for(Object col : row1) {
			XSSFCell column = 	row.createCell(columncount++);
				// Checking the type of data and making the entry : 
			// Keyword instance of 
			    if (col instanceof String) {
			    
			    	column.setCellValue((String)col);
			    }else if (col instanceof Integer) {
			    	column.setCellValue((Integer)col);
			    }
			}	
	}
			try {
			FileOutputStream output = new FileOutputStream ("C:\\Users\\ralaguchamy\\OneDrive - IGT PLC\\Desktop\\Learning\\Eclipse 9  9 2024\\ExcelFileOperation\\src\\main\\java\\ReadAndWriteExcel\\Details");
			book.write(output);
			
			}catch(Exception e){
		   e.printStackTrace();
	   		}	
		book.close();
	}
}
