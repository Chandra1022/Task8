package javapackage;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.*;

public class ReadingExcel {

	public static void main(String[] args) throws IOException {
		
		String excelFilePath=".\\Resource\\JAVA.xlsx";
		FileInputStream inputstream=new FileInputStream(excelFilePath);
		XSSFWorkbook workbook=new XSSFWorkbook(inputstream);
		XSSFSheet sheet=workbook.getSheetAt(0);
		
		//Using for loop
		
		int rows=sheet.getLastRowNum();
		int col=sheet.getRow(1).getLastCellNum();
		
		//System.out.println(rows);
		//System.out.println(col);
		
		for(int r=0;r<=rows;r++) {
			XSSFRow row=sheet.getRow(r);
			for(int c=0;c<col;c++) {
				XSSFCell cell=row.getCell(c);
				
				switch(cell.getCellType()) {
				case STRING : System.out.print(cell.getStringCellValue());break;
				case NUMERIC: System.out.print(cell.getNumericCellValue());break;
				case BOOLEAN: System.out.print(cell.getBooleanCellValue());break;
				}
				System.out.print(" | ");
			}
			System.out.println();
		}
		
	}

}
