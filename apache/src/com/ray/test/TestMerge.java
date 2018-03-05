package com.ray.test;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.ray.excel.Excel;

public class TestMerge {

	public static final String excelFilePath1 = "/home/rayweihao/temp/1.xlsx";
	
	public static final String excelFilePath2 = "/home/rayweihao/temp/4.xlsx";
	
	public static void main(String[] args) {
		// TODO Auto-generated method stub

		try {
			Workbook wb1 = Excel.openWorkbookWithFactoryByFile(excelFilePath1);
			
			Sheet sheet = Excel.getSheetAt(wb1, 0);
			Excel.mergeCells(sheet, 0, 10, 0, 10);
			
			Excel.saveWorkbook(excelFilePath2, wb1);
			
		} catch (EncryptedDocumentException | InvalidFormatException | IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}

}
