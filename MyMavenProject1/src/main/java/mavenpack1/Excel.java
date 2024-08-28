package mavenpack1;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {
	XSSFSheet sh;
	
	public Excel() throws IOException {
		FileInputStream f = new FileInputStream("D:\\Test Excel\\Test.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(f);
		sh=wb.getSheet("Sheet1");	
	}

	public String readData(int i, int j) {
		Row r=sh.getRow(i);
		Cell c=r.getCell(j);
		int Celltype=c.getCellType();
		switch(Celltype) {
		case Cell.CELL_TYPE_NUMERIC:
		{
			double d=c.getNumericCellValue();
			return String.valueOf(d);
			
		}
		
		case Cell.CELL_TYPE_STRING:
		{
			String s=c.getStringCellValue();
			return String.valueOf(s);
		}
		}
		
		return null;
	}

}
