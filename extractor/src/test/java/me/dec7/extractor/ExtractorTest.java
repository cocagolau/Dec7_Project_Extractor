package me.dec7.extractor;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class ExtractorTest {
	
	
	private static final Logger logger = LoggerFactory.getLogger(ExtractorTest.class);
	
	
	@Test
	public void test() throws IOException {
		FileInputStream is = new FileInputStream("/Users/Dec7/Desktop/excel-test.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(is);
		
		int rowIdx = 0;
		int colIdx = 0;
		
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		int rows = sheet.getPhysicalNumberOfRows();
		for (rowIdx=1; rowIdx<rows; rowIdx++) {
			XSSFRow row = sheet.getRow(rowIdx);
			if (row != null) {
				
				int cells = row.getPhysicalNumberOfCells();
				for (colIdx=1; colIdx<cells; colIdx++) {
					XSSFCell cell = row.getCell(colIdx);
					String value = "";
					
					if (cell == null) {
						continue;
					} else {
						switch (cell.getCellType()) {
						case XSSFCell.CELL_TYPE_FORMULA :
							value = cell.getCellFormula();
							break;
						case XSSFCell.CELL_TYPE_NUMERIC :
							value = cell.getNumericCellValue() + "";
							break;
						case XSSFCell.CELL_TYPE_STRING :
							value = cell.getStringCellValue();
							break;
						case XSSFCell.CELL_TYPE_BOOLEAN :
							value = cell.getBooleanCellValue() + "";
							break;
						case XSSFCell.CELL_TYPE_ERROR :
							value = cell.getErrorCellString();
							break;
						}
					}
					
					System.out.println(value);
				}
				
				
			}
		}
	}

}
