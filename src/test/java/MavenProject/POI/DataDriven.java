package MavenProject.POI;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDriven {

	public static void main(String[] args) throws IOException {
		
		FileInputStream fis = new FileInputStream("C://Users/Julie//workspace//POI//TestCases.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		
		int sheets = workbook.getNumberOfSheets();
		for(int i=0;i<sheets;i++){
			if(workbook.getSheetName(i).equalsIgnoreCase("TestData")){
			XSSFSheet sheet = workbook.getSheetAt(i);
			Iterator<Row> rows = sheet.iterator();
			Row firstrow = rows.next();
			//collection of cells
			Iterator<Cell> ce = firstrow.cellIterator();
			int k=0;
			int column;
			while(ce.hasNext())
			{
				Cell value=ce.next();
				value.getStringCellValue().equalsIgnoreCase("TestCases")
				{
						//desired column
					column=k;
				}
			k++;
			}
			
			}
		}
		
	}

}
