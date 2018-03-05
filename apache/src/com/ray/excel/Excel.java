package com.ray.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {

	public static void createExcelOnly(String excelFilePath) throws IOException {
		createExcelOnlyWithSheetNames(excelFilePath, null);
	}
	
	public static void createExcelOnlyWithSheetName(String excelFilePath,String sheetName) throws IOException {
		List<String> sheetNameList = new ArrayList<String>();
		sheetNameList.add(sheetName);
		createExcelOnlyWithSheetNames(excelFilePath, sheetNameList);
	}
	
	public static void createExcelOnlyWithSheetNames(String excelFilePath, List<String> sheetNameList) throws IOException {
		Workbook wb = null;
		if(null != excelFilePath) {
			if(excelFilePath.toLowerCase().endsWith(".xls")) {
				wb = new HSSFWorkbook();
			}else if(excelFilePath.toLowerCase().endsWith(".xlsx")) {
				wb = new XSSFWorkbook();
			}
			
			if(null != sheetNameList && sheetNameList.size() > 0) {
				for(String sheetName : sheetNameList){
					if(null != sheetName && !sheetName.equals("")) {
						wb.createSheet(sheetName);
					}
				}
			}
			else {
				wb.createSheet(Constants.DEFAULTSHEETNAME);
			}
			
			FileOutputStream fileOut = new FileOutputStream(excelFilePath);
		    wb.write(fileOut);
		    fileOut.close();
		}
	}
	
	
	/**
	 * You must invoke saveWorkbook function to save your excel file.
	 * @param excelFilePath
	 * @return Workbook
	 */
	public static Workbook createWorkbook(String excelFilePath) {
		Workbook wb = null;
		if(null != excelFilePath) {
			if(excelFilePath.toLowerCase().endsWith(".xls")) {
				wb = new HSSFWorkbook();
			}else if(excelFilePath.toLowerCase().endsWith(".xlsx")) {
				wb = new XSSFWorkbook();
			}
		}
		return wb;
	}
	
	/**
	 * 
	 * @param excelFilePath
	 * @param wb
	 * @throws IOException
	 */
	public static void saveWorkbook(String excelFilePath, Workbook wb) throws IOException {
		FileOutputStream fileOut = new FileOutputStream(excelFilePath);
	    wb.write(fileOut);
	    fileOut.close();
	}
	
	public static void createNewSheet(Workbook wb) {
		createNewSheets(wb, null);
	}
	
	public static void createNewSheetWithSheetName(Workbook wb, String sheetName) {
		List<String> sheetNameList = new ArrayList<String>();
		sheetNameList.add(sheetName);
		createNewSheets(wb, sheetNameList);
	}
	
	public static void createNewSheets(Workbook wb, List<String> sheetNameList) {
		if(null != sheetNameList && sheetNameList.size() > 0) {
			for(String sheetName : sheetNameList){
				if(null != sheetName && !sheetName.equals("")) {
					wb.createSheet(sheetName);
				}
			}
		}
		else {
			wb.createSheet(Constants.DEFAULTSHEETNAME);
		}
	}
	
	public static String getValidationSheetName(String sheetName) {
		return WorkbookUtil.createSafeSheetName(sheetName);
	}
	
	public static Sheet getSheetByName(Workbook wb, String sheetName) {
		return wb.getSheet(sheetName);
	}
	
	public static Sheet getSheetAt(Workbook wb, int num) {
		return wb.getSheetAt(num);
	}
	
	/**
	 * Create a row. Rows are 0 based
	 * @param sheet
	 * @param rowNum
	 * @return
	 */
	public static Row createRow(Sheet sheet, int rowNum) {
		return sheet.createRow(rowNum);
	}
	
	/**
	 * Create a cell. Cells are 0 based
	 * @param row
	 * @param columnNum
	 * @return
	 */
	public static Cell createCell(Row row, int columnNum) {
		return row.createCell(columnNum);
	}
	
	/**
	 * The formatting style follow the Java formatting for Date
	 * @param wb
	 * @param row
	 * @param columnNum
	 * @param dataFormat
	 * @return Date Cell
	 */
	public static Cell createDateCell(Workbook wb, Row row, int columnNum, String dataFormat) {
		Cell cell = row.createCell(columnNum);
		CellStyle cellStyle = wb.createCellStyle();
		CreationHelper createHelper = wb.getCreationHelper();
		cellStyle.setDataFormat(createHelper.createDataFormat().getFormat(dataFormat));
		cell.setCellStyle(cellStyle);
		return cell;
	}
	
	
	
	public static RichTextString createRichTextString(Workbook wb, String str) {
		if(null != wb) {
			CreationHelper createHelper = wb.getCreationHelper();
			return createHelper.createRichTextString(str);
		}
		return null;
	}
	
	/**
	 * Use a file
	 * @param excelFilePath
	 * @return
	 * @throws EncryptedDocumentException
	 * @throws InvalidFormatException
	 * @throws IOException
	 */
	public static Workbook openWorkbookWithFactoryByFile(String excelFilePath) throws EncryptedDocumentException, InvalidFormatException, IOException {
		return WorkbookFactory.create(new File(excelFilePath));
	}
	
	/**
	 * Use an InputStream, needs more memory
	 * @param excelFilePath
	 * @return
	 * @throws EncryptedDocumentException
	 * @throws InvalidFormatException
	 * @throws IOException
	 */
	public static Workbook openWorkbookWithFactoryByStream(String excelFilePath) throws EncryptedDocumentException, InvalidFormatException, IOException {
		return WorkbookFactory.create(new FileInputStream(excelFilePath));
	}
	
}
