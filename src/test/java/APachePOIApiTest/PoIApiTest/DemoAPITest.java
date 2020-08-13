package APachePOIApiTest.PoIApiTest;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DemoAPITest {
	
	public static void main(String[] args) throws IOException {
		
		 ArrayList<String> alist = getDataFromExcelFile("Register");
		 
		 for(String a : alist) {
			 
			 System.out.println(a);
		 }
		
	}

	public static ArrayList<String> getDataFromExcelFile(String TestName) throws IOException {
		
		ArrayList<String> alist = new ArrayList<String>();
		
		//FileInputStream fis = new FileInputStream("C:\\Users\\smanchegowda\\Desktop\\ExcelTestData.xlsx");
		
		FileInputStream fis = new FileInputStream("ExcelTestData.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		int sheetcount = workbook.getNumberOfSheets();
		
		//System.out.println("Number of sheets in excel file is :" +sheetcount);
		
		for(int i=0; i<sheetcount; i++) {
			
			if(workbook.getSheetName(i).equalsIgnoreCase("SheetA")) {
				
				XSSFSheet sheet = workbook.getSheetAt(0);
				
				 Iterator<Row> rows = sheet.iterator();
				 
				Row firstrow = rows.next();
				
				Iterator<Cell> firstrowCells = firstrow.iterator();
				
				int c=0;
				
				int TestColumnPosition =0;
				
				while (firstrowCells.hasNext()) {
					//System.out.println(firstrowCells.next().getStringCellValue());
					
					Cell firstrowCell = firstrowCells.next();
					
					if (firstrowCell.getRichStringCellValue().equals("Tests")) {
						
						TestColumnPosition = c;
						
					}
					c++;
				}
				 
				while(rows.hasNext()) {
					
					 Row row = rows.next();
					 
					 Cell cell = row.getCell(TestColumnPosition);
					 
					 if(cell.getStringCellValue().equalsIgnoreCase(TestName)){
						 
						Iterator<Cell> cells = row.iterator();
						
						cells.next();
						
						while(cells.hasNext()) {
							
						 Cell currentcell = cells.next();
						 
						 if(currentcell.getCellType()==CellType.STRING) {
							 
							 alist.add(currentcell.getStringCellValue());
							 
							 //System.out.println(currentcell.getStringCellValue());
							 
						 } else if (currentcell.getCellType()==CellType.NUMERIC){
							 
							 alist.add(NumberToTextConverter.toText(currentcell.getNumericCellValue()));
							 
							// System.out.println(currentcell.getNumericCellValue());
							 
						 }
							 
							
						}
						
						 
						 
					 }
				}
				
				
				
				
			}
		}
		
		return alist;
		

	}

}
