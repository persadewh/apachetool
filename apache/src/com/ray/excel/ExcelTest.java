package com.ray.excel;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;


public class ExcelTest {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		String excelFilePath = "/home/rayweihao/temp/1.xlsx";
		String excelFilePath1 = "/home/rayweihao/temp/2.xls";
		
		try {
			//Excel.createExcelOnly(excelFilePath);
			//Excel.createExcelOnly(excelFilePath1);
			
//			List<String> sheetNameList = new ArrayList<String>();
//			sheetNameList.add("1");
//			sheetNameList.add("2");
//			sheetNameList.add("3");
//			sheetNameList.add("4");
//			sheetNameList.add("5");
//			sheetNameList.add("6");
//			
//			Excel.createExcelOnlyWithSheetNames(excelFilePath,sheetNameList);
			
			Workbook wb = Excel.createWorkbook(excelFilePath);
			Excel.createNewSheetWithSheetName(wb, "Y");
			
			Sheet sheet = Excel.getSheetByName(wb, "Y");
			
			Row row = Excel.createRow(sheet, 0);
			Cell cell = Excel.createCell(row, 0);
			cell.setCellValue("This is my test cell");
			
			CellStyle cellStyle = Excel.createCellStyle(wb);
			
			Excel.setAlignment(cellStyle, HorizontalAlignment.CENTER, VerticalAlignment.CENTER);
			Excel.setBorder(cellStyle, BorderStyle.DOUBLE, 
					BorderStyle.DOUBLE, 
					BorderStyle.DOUBLE, 
					BorderStyle.DOUBLE, 
					IndexedColors.RED.getIndex(),
					IndexedColors.RED.getIndex(), 
					IndexedColors.RED.getIndex(), 
					IndexedColors.RED.getIndex());
			
			cell.setCellStyle(cellStyle);
			
			Cell dateCell = Excel.createDateCell(wb, row, 1, "yyyy/mm/dd hh:MM:ss");
			dateCell.setCellValue(new Date());
			
			Excel.saveWorkbook(excelFilePath, wb);
			
			
//			Workbook wb1 = Excel.openWorkbookWithFactory(excelFilePath1);
//			Sheet sheet1 = Excel.getSheetByName(wb1, "Y");
//			Row row1 = sheet1.getRow(0);
//			System.out.println(row1.getCell(0).getStringCellValue());
//			System.out.println(row1.getCell(1).getDateCellValue());
//			
//			Row row2 = Excel.createRow(sheet1, 1);
//			Cell cell1 = Excel.createCell(row2, 0);
//			cell1.setCellValue("This is my test cell");
//			
//			Cell dateCell1 = Excel.createDateCell(wb1, row2, 1, "yyyy/mm/dd hh:MM:ss");
//			dateCell1.setCellValue(new Date());
//			
//			Excel.saveWorkbook("/home/rayweihao/temp/3.xls", wb1);
			
			
			Workbook wb1 = Excel.openWorkbookWithFactoryByFile(excelFilePath1);
			Sheet sheet1 = Excel.getSheetByName(wb1, "Y");
			Row row1 = sheet1.getRow(0);
			System.out.println(Excel.getCellValue(row1.getCell(1)));
			System.out.println(Excel.getCellValue(row1.getCell(0)));
			
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (EncryptedDocumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}

}
