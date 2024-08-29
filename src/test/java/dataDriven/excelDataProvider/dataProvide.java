package dataDriven.excelDataProvider;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class dataProvide 
{
	// Here we will send five sets of Data in form of Arrays through dataprovider
   // Then the test will run 5 times with 5 seprate sets of Data 
	DataFormatter formatter = new DataFormatter();
	@Test(dataProvider = "driveTest")
	public void testCaseData(String greeting , String communication ,String id)
	{
		System.out.println(greeting+communication+id);
		
	}
	
	
	
	@DataProvider(name ="driveTest")
	public Object[][] getData() throws IOException 
	{
	//Every row of excel should be stored in one Array
		FileInputStream fis = new FileInputStream("C:\\Users\\Gaurav Singh\\Documents\\excelDriven.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet =wb.getSheetAt(0);
		int rowCount=sheet.getPhysicalNumberOfRows();
		XSSFRow row =sheet.getRow(0);
		int colcount =row.getLastCellNum();
		Object data[][] = new Object[rowCount-1][colcount];
		
		for(int i=0; i<(rowCount-1) ;i++)
		{
			
			row = sheet.getRow(i+1);
			
			for(int j=0; j<colcount; j++)	
			{
			 XSSFCell cell = row.getCell(j)	;
			
				
		     data [i][j]= formatter.formatCellValue(cell);
			}
			
		}	
		
		return data ;
		
		// Object[][] data = {{"hello" ,"text","1"},{"bye" ,"message","143"},{"solo" ,"call","453"}}	;
		//return data ;
	}
	
	
	
	

}
