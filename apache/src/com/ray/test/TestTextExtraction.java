package com.ray.test;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;

import com.ray.excel.Excel;

public class TestTextExtraction {
	
	public static final String excelFilePath1 = "/home/rayweihao/temp/1.xls";
	
	public static void main(String[] args) {
		
		try {
			Workbook wb1 = Excel.openWorkbookWithFactoryByFile(excelFilePath1);
			
			System.out.println(Excel.textExtraction((HSSFWorkbook) wb1));
			
		} catch (EncryptedDocumentException | InvalidFormatException | IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}

}
