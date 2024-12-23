package readdata;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;



public class readdataexcel {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		

		//1. Provide file location
		    File file=new File("./Book11.xlsx");
		    	//clear
		    	
		//2.   create con and provide location (file) 	
		    	
		    	FileInputStream fis = null;
				try {
					fis = new FileInputStream(file);
				} catch (FileNotFoundException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
		    
		//3. Workbook File
		    Workbook wb = null;
			try {
				wb = WorkbookFactory.create(fis);
			} catch (EncryptedDocumentException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		    
		              //get the sheet(1)
		    Sheet sheet= wb.getSheetAt(0);
		    
		    
		    //sheet --> row
		    
		      Row row= sheet.getRow(0);
		     
		     //row ---> cell
		      
		      Cell cell= row.getCell(0);
		      
		      System.out.println(cell);
	}

}
