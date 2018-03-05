package com.ray.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.examples.CellTypes;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.microsoft.schemas.office.visio.x2012.main.CellType;

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
	
	public static void autoSizeColumn(Sheet sheet) {
		sheet.autoSizeColumn((short)2);
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
	
	public static void mergeCells(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
		sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
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
	
	/**
	 * HorizontalAlignment
	 * VerticalAlignment
	 * @param cellStyle
	 * @param halign
	 * @param valign
	 */
	public static void setAlignment(CellStyle cellStyle, HorizontalAlignment halign, VerticalAlignment valign) {
		cellStyle.setAlignment(halign);
		cellStyle.setVerticalAlignment(valign);
	}
	
	/**
	 * BorderStyle
	 * IndexedColors
	 * @param cellStyle
	 * @param topBorderStyle
	 * @param bottomBorderStyle
	 * @param leftBorderStyle
	 * @param rightBorderStyle
	 * @param topBorderColor
	 * @param bottomBorderColor
	 * @param leftBorderColor
	 * @param rightBorderColor
	 */
	public static void setBorder(CellStyle cellStyle, 
			BorderStyle topBorderStyle, 
			BorderStyle bottomBorderStyle, 
			BorderStyle leftBorderStyle, 
			BorderStyle rightBorderStyle, 
			short topBorderColor, 
			short bottomBorderColor, 
			short leftBorderColor, 
			short rightBorderColor) {
		
		cellStyle.setBorderBottom(bottomBorderStyle);
		cellStyle.setBottomBorderColor(bottomBorderColor);
	    cellStyle.setBorderLeft(leftBorderStyle);
	    cellStyle.setLeftBorderColor(leftBorderColor);
	    cellStyle.setBorderRight(rightBorderStyle);
	    cellStyle.setRightBorderColor(rightBorderColor);
	    cellStyle.setBorderTop(topBorderStyle);
	    cellStyle.setTopBorderColor(topBorderColor);
		
	}
	
	/**
	 * style.setFillPattern(CellStyle.BIG_SPOTS);
	 * @param cellStyle
	 * @param backGroundColor
	 * @param foreGroundColor
	 * @param fillPatternType
	 */
	public static void setFill(CellStyle cellStyle, 
			short backGroundColor, 
			short foreGroundColor, 
			FillPatternType fillPatternType ) {
		cellStyle.setFillBackgroundColor(backGroundColor);
		cellStyle.setFillForegroundColor(foreGroundColor);
		cellStyle.setFillPattern(fillPatternType);
	}
	
	/**
	 * You can use \n with word wrap on to create a new line in cell
	 * But you must invoke setWrapText function
	 * @param cellStyle
	 */
	public static void setWrapText(CellStyle cellStyle) {
		cellStyle.setWrapText(true);
	}
	
	public static CellStyle createCellStyle(Workbook wb) {
		return wb.createCellStyle();
	}
	
	/**
	 * You should re-use fonts in your applications instead of creating a font for each cell.
	 * 
	 * @param cellStyle
	 * @param font
	 */
	public static void setFont(CellStyle cellStyle, Font font) {
		cellStyle.setFont(font);
	}
	
	public static Font createFont(Workbook wb) {
		return wb.createFont();
	}
	
	public static String getCellValue(Cell cell) {
		String value = "";
		if(null != cell) {
			switch(cell.getCellType()) {
			case Cell.CELL_TYPE_STRING:
				value = cell.getRichStringCellValue().getString();
				break;
			case Cell.CELL_TYPE_NUMERIC:
				if(DateUtil.isCellDateFormatted(cell)) {
					Date date = cell.getDateCellValue();
					value = date.toLocaleString();
				}else {
					value = String.valueOf(cell.getNumericCellValue());
				}
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				value = String.valueOf(cell.getBooleanCellValue());
				break;
			case Cell.CELL_TYPE_FORMULA:
				value = cell.getCellFormula();
				break;
			case Cell.CELL_TYPE_BLANK:
				value = "";
				break;
			default:
				value = "";
			}
		}
		return value;
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
	 * @return Workbook
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
	 * @return Workbook
	 * @throws EncryptedDocumentException
	 * @throws InvalidFormatException
	 * @throws IOException
	 */
	public static Workbook openWorkbookWithFactoryByStream(String excelFilePath) throws EncryptedDocumentException, InvalidFormatException, IOException {
		return WorkbookFactory.create(new FileInputStream(excelFilePath));
	}
	
	/**
	 * get all text in excel
	 * just for HSSF(xls)
	 * @param wb
	 * @return text
	 */
	public static String textExtraction(HSSFWorkbook wb) {
		ExcelExtractor extractor = new ExcelExtractor(wb);
		extractor.setFormulasNotResults(true);
		extractor.setIncludeSheetNames(false);
		return extractor.getText();
	}
	
}
