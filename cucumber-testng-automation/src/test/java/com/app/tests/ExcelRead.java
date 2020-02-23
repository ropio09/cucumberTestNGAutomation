package com.app.tests;

import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
/**
 * Creating excell sheet and using it with apache 
 * @author arod
 *
 */
public class ExcelRead {
public static void main(String[] args) throws Exception {
	

	
	String filePath="/Users/arod/Desktop/Employees.xlsx";

	FileInputStream inStream = new FileInputStream(filePath);
	
	Workbook workbook = WorkbookFactory.create(inStream);
	Sheet worksheet= workbook.getSheetAt(0); 
	Row row = worksheet.getRow(0);
	// row.getcell will get first line 
	Cell cell= row.getCell(0);
	System.out.println(cell.toString());	
	
	//Find out how many rows in Excel Sheet
	int rowsCount= worksheet.getPhysicalNumberOfRows();//getLastRowNum(); // will print empty rows 
	System.out.println("Number of Rows: "+rowsCount);
	for (int rowNum = 1; rowNum < rowsCount; rowNum++) {
		row = worksheet.getRow(rowNum);
		System.out.println(rowNum+" - "+cell.toString());
	//	System.out.println(rowNum+" - "+worksheet.getRow(rowNum).getCell(0));
	}
	
	//print the job id of nancy
	System.out.println(worksheet.getRow(5).getCell(2).toString());
	Cell NancyJob = worksheet.getRow(5).getCell(2);
	
	for (int i = 1; i < worksheet.getPhysicalNumberOfRows(); i++) {
		Row myrow= worksheet.getRow(i);
		if(myrow.getCell(0).toString().equals("Nancy")) {
			// print job ID from same row
			System.out.println("Nancy works as : "+myrow.getCell(2).toString());
			break;
		}
	}
	
	workbook.close();
	inStream.close();
}
}
