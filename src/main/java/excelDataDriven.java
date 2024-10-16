import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excelDataDriven {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		FileInputStream fis=new FileInputStream("C://Selenium_2024//excelDriven//TestData.xlsx");
		XSSFWorkbook workBook=new XSSFWorkbook(fis);
		
		int sheets=workBook.getNumberOfSheets();
		for(int i=0;i<sheets;i++) {
			
			if(workBook.getSheetName(i).equalsIgnoreCase("TestData")) {
				XSSFSheet sheet=workBook.getSheetAt(i);
				
//				Indentify TestCase column by scanning entire first row
				Iterator<Row> rows= sheet.iterator(); //sheet is collection of rows
				Row firstRow=rows.next();
				Iterator<Cell> ce= firstRow.cellIterator(); //row is collection of cells
				int k=0;
				int column=0;
				while (ce.hasNext()) {
					Cell value=ce.next();
					if(value.getStringCellValue().equalsIgnoreCase("TestCase")) {
						column=k;
					}
					k++;
				}
				System.out.println(column);
			}
			
		}
		

	}

}
