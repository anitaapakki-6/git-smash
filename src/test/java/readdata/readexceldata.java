package readdata;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;



public class readexceldata {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		

		//1. Provide file location
		    File file=new File("C:\\Users\\lumbu\\Desktop\\Book11.xlsx");
		    	//clear
		    	
		//2.   create con and provide location (file) 	
		    	
		    	FileInputStream fis = new FileInputStream(file);
				
		    
		//3. Workbook File
		    Workbook wb =WorkbookFactory.create(fis) ;
			
		    
		              //get the sheet(1)
		    Sheet sheet= wb.getSheetAt(0);
		    
		    
		    //sheet --> row
		    
		      Row row= sheet.getRow(1);
		     
		     //row ---> cell
		      
		      Cell cell= row.getCell(1);
		      
		      System.out.println(cell);
	}





	
	}


