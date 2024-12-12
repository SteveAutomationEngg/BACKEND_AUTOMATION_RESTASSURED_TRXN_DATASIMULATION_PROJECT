package Sim;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelUtility {
	
	String testDataPath = Driver.path;  //as Path is static in Driver Class    
	
	public String getDataFromExcel(String sheetName, int rowNum , int celNum) throws Throwable {
		FileInputStream fis = new FileInputStream(testDataPath);
		Workbook wb = WorkbookFactory.create(fis);
		String data = wb.getSheet(sheetName).getRow(rowNum).getCell(celNum).getStringCellValue();
		wb.close();
		
		return data;	
	}
	
	
	public int getRowcount(String sheetName) throws Throwable {
		FileInputStream fis = new FileInputStream(testDataPath);
		Workbook wb = WorkbookFactory.create(fis);
		int rowCount = wb.getSheet(sheetName).getLastRowNum();
		wb.close();
		
		return rowCount;
	}
	
	
	public void setDataIntoExcel(String sheetName, int RowNum, int CelNum,  String data) throws Throwable {
		
		FileInputStream fis = new FileInputStream(testDataPath);
		Workbook wb = WorkbookFactory.create(fis);
		Cell cel= wb.getSheet(sheetName).getRow(RowNum).createCell(CelNum);
		cel.setCellType(CellType.STRING);
		cel.setCellValue(data);
		
		FileOutputStream fos = new FileOutputStream(testDataPath);
		wb.write(fos);
		wb.close();
	}
	

	
}
