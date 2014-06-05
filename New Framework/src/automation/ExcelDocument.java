package automation;


import java.io.File;  
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.Properties;
import java.util.Set;
import java.util.Vector;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.log4j.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dom4j.Document;
import org.w3c.dom.Element;


/**
 * ExcelDocument class is generated to perfrom excel operation based on apache
 * POI Library
 * 
 * @author ckatta@cisco.com
 * 
 */
public class ExcelDocument {

	Logger consoleLogger = Logger.getRootLogger();

	HashMap<String, Integer> sequentialPatternCounts = new HashMap<String, Integer>();
	int sequentialPatternCount = 0;
	HashMap<String, Integer> sequentialPatternCounts1 = new HashMap<String, Integer>();
	int sequentialPatternCount1 = 0;

	/**
	 * Default constructor
	 * 
	 * <pre>
	 * {@code
	 * ExcelDocument excelDocument = new ExcelDocumnt(String fileName);
	 * excelDocument.selectSheet(String sheetName);
	 * }
	 * </pre>
	 */
	public ExcelDocument() {
		workbook = null;
		workingSheet = null;
	}

	/**
	 * Constructor, File path as single input argument
	 * 
	 * @param fileURL
	 *            <pre>
	 * {@code
	 * ExcelDocument excelDocument = new ExcelDocumnt(String fileName);
	 * excelDocument.selectSheet(String sheetName);
	 * }
	 * </pre>
	 */
	public ExcelDocument(String fileURL) {
		initiateFileInstance(fileURL);
	}

	/**
	 * Constructor, Directory & File paths are input arguments
	 * 
	 * @param directoryPath
	 * @param filePath
	 *            <pre>
	 * {@code
	 * ExcelDocument excelDocument = new ExcelDocumnt(String fileName);
	 * excelDocument.selectSheet(String sheetName);
	 * }
	 * </pre>
	 */
	public ExcelDocument(String directoryPath, String filePath) {
	}

	/**
	 * Helping variables, logging.
	 */
	static Logger logger = Logger.getLogger(ExcelDocument.class);

	/**
	 * Excel related variables declarations
	 */
	Workbook workbook;
	Sheet workingSheet;

	/**
	 * 
	 */
	String direcotyPath;
	String workingFileName;

	/**
	 * @return the direcotyPath
	 */
	public String getDirecotyPath() {
		return direcotyPath;
	}

	/**
	 * @param directoryPath
	 *            the direcotyPath to set
	 */
	public void setDirecotyPath(String direcotyPath) {
		logger.info("'" + direcotyPath + "' is current working directory.");
		this.direcotyPath = direcotyPath;
	}

	/**
	 * @return the workingFileName
	 */
	public String getWorkingFileName() {
		return workingFileName;
	}

	/**
	 * @param workingFileName
	 *            the workingFileName to set
	 */
	public void setWorkingFileName(String workingFileName) {
		logger.info("'" + workingFileName + "' is current working file");
		this.workingFileName = workingFileName;
	}

	/**
	 * Saving Excel file after modifications made in excel file.
	 */
	public void saveExcel() {
		try {
			FileOutputStream fos = new FileOutputStream(direcotyPath
					+ workingFileName);
			workbook.write(fos);
			fos.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	/**
	 * Saving Excel file after modifications made in excel file.
	 */
	public void saveExcel(String filePath) {
		try {
			FileOutputStream fos = new FileOutputStream(filePath);
			workbook.write(fos);
			fos.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	/**
	 * Creating new excel file for given file name. Creates, if file not exists
	 * in given location
	 * 
	 * @return
	 */
	public boolean createFile() {
		File file = new File(direcotyPath + workingFileName);
		if (!file.exists()) {
			workbook = new XSSFWorkbook();
			workbook.createSheet("variables");
			saveExcel();
		}
		initiateFileInstance();
		return true;
	}
	
	public Sheet getSheet(){
		return workingSheet;
	}

	public boolean createFile(String filePath) {
		File file = new File(filePath);
		consoleLogger.debug(filePath);
		if (!file.exists()) {
			workbook = new XSSFWorkbook();
			workbook.createSheet("variables");
			saveExcel(filePath);
		}
		return true;
	}

	/**
	 * Creating new Sheet with name
	 * 
	 * @param sheetName
	 *            , new sheet name
	 */
	public void createSheet(String sheetName) {
		Sheet sheet = workbook.getSheet(sheetName);
		if (sheet == null) {
			sheet = workbook.createSheet(sheetName);
		}
	}

	/**
	 * Select the sheet with name
	 * 
	 * @param sheetName
	 *            , new sheet name
	 */
	public void selectSheet(String sheetName) {
		workingSheet = workbook.getSheet(sheetName);
		if (workingSheet == null) {
			workingSheet = workbook.createSheet(sheetName);
		}
	}

	/**
	 * Select the sheet with number
	 * 
	 * @param sheetName
	 *            , new sheet name
	 */
	public void selectSheet(int sheetNo) {
		workingSheet = workbook.getSheetAt(sheetNo);
		if (workingSheet == null) {
			workingSheet = workbook.createSheet("temp");
		}
	}

	public void initiateFileInstance(String filePath) {
		createFile(filePath);
		try {
			FileInputStream fis = new FileInputStream(filePath);
			workbook = WorkbookFactory.create(fis);
			workingSheet = workbook.getSheetAt(0);
			// consoleLogger.debug("IN EXCEL");
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	public void initiateFileInstance() {
		try {
			FileInputStream fis = new FileInputStream(direcotyPath
					+ workingFileName);
			workbook = WorkbookFactory.create(fis);
			workingSheet = workbook.getSheetAt(0);
			// consoleLogger.debug("IN EXCEL");
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	/**
	 * 
	 * @param startingRow
	 * @param keyColumnNum
	 * @param valueColumnNum
	 * @param inputData
	 */
	public void addDictionaryDataTOExcelDocument(int startingRow,
			int keyColumnNum, int valueColumnNum,
			HashMap<String, String> inputData) {
		logger.info("DATA IN");
		int rowCount = startingRow;
		Set<String> keys = inputData.keySet();
		Iterator<String> keyIterator = keys.iterator();
		while (keyIterator.hasNext()) {
			String identifierString = keyIterator.next();
			String valueString = inputData.get(identifierString);
			Row startRow = workingSheet.getRow(rowCount);
			if (startRow == null) {
				startRow = workingSheet.createRow(rowCount);

			}
			updateCellValue(rowCount, keyColumnNum, identifierString);
			updateCellValue(rowCount, valueColumnNum, valueString);
			rowCount++;
		}
	}

	public void addDictionaryDataTOExcelDocumentBySortingKeys(int startingRow,
			int keyColumnNum, int valueColumnNum,
			HashMap<String, LinkedList<String>> inputData) {
		logger.info("DATA IN");
		int rowCount = startingRow;
		Set<String> keys = inputData.keySet();
		int keysLength = keys.size();
		int iterator = 1;
		while (iterator < keysLength) {
			String identifierString = String.valueOf(iterator);
			LinkedList<String> valueStrings = inputData.get(identifierString);
			Row startRow = workingSheet.getRow(rowCount);
			if (startRow == null) {
				startRow = workingSheet.createRow(rowCount);

			}
			// updateCellValue(rowCount, keyColumnNum, identifierString);
			int vcn = valueColumnNum;
			boolean mainHeader = false;
			boolean subTotalHeader = false;
			for (String valueString : valueStrings) {
				XSSFCellStyle style = (XSSFCellStyle) workbook
						.createCellStyle();
				if (valueString.startsWith("Main")) {
					mainHeader = true;
				} else if (valueString.startsWith("Subtotal")) {
					subTotalHeader = true;
				}
				if (mainHeader) {
					// style.setFillBackgroundColor(IndexedColors.GREY_80_PERCENT.getIndex());
					style.setFillForegroundColor(new XSSFColor(
							new java.awt.Color(238, 238, 238)));
					style.setFillPattern(CellStyle.SOLID_FOREGROUND);
				} else if (subTotalHeader) {
					// style.setFillBackgroundColor(IndexedColors.AQUA.getIndex());
					style.setFillForegroundColor(new XSSFColor(
							new java.awt.Color(233, 253, 234)));
					style.setFillPattern(CellStyle.SOLID_FOREGROUND);
				} else {
					// style.setFillBackgroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
					style.setFillForegroundColor(new XSSFColor(
							new java.awt.Color(255, 255, 255)));
					style.setFillPattern(CellStyle.SOLID_FOREGROUND);
				}
				updateCellValue(style, rowCount, vcn - 1, valueString);
				vcn++;
			}
			rowCount++;
			iterator++;
		}
	}

	/**
	 * 
	 * @param startingRow
	 * @param columnNum
	 * @param allVariables
	 */
	@SuppressWarnings("rawtypes")
	public void addListDataVerticallyToExcelColumn(int startingRow,
			int columnNum, Vector<String> allVariables) {
		logger.info("Excel Data Storage Properties are: "
				+ "Data starts from Row " + startingRow + " & Columun "
				+ columnNum + ".");
		if (allVariables != null) {
			Iterator iterator = allVariables.iterator();
			int count = startingRow;
			while (iterator.hasNext()) {
				String string = (String) iterator.next();
				if (!isValueExitsInGivenCell(count, columnNum, string)
						&& getVlaueFromSameRow(columnNum, columnNum, string)
								.length() == 0) {
					logger.info("new identifier " + string
							+ " is added to excel");
					addNewVlaueToLastRow(startingRow, columnNum, string);
				}
				count++;
			}
		}
	}

	/**
	 * 
	 * @param startingRow
	 * @param keyColumNum
	 * @param columnNums
	 * @return
	 */
	public HashMap<String, LinkedList<String>> getSelectedExcelDataIntoMap(
			int startingRow, int keyColumNum, LinkedList<Integer> columnNums) {
		HashMap<String, LinkedList<String>> keyValueMap = new HashMap<String, LinkedList<String>>();
		int rowCount = startingRow;
		for (Row row : workingSheet) {
			int rowNumber = row.getRowNum();
			if (rowNumber == rowCount) {
				String keyString = getValueFromExcel(rowCount, keyColumNum);
				if (keyString != null) {
					LinkedList<String> dependentStrings = new LinkedList<String>();
					for (Integer columnNum : columnNums) {
						String dependentString = getValueFromExcel(rowCount,
								columnNum);
						if (dependentString != null) {
							dependentStrings.add(dependentString);
						} else {
							dependentStrings.add("");
						}
					}
					keyValueMap.put(keyString, dependentStrings);
				}
				rowCount++;
			}
		}
		return keyValueMap;
	}

	/**
	 * 
	 * @param identifierString
	 * @param identifierStringcolumnNum
	 * @param IdentifierStringValueColumnNum
	 * @return
	 */
	public HashMap<String, HashMap<String, Object>> createDataStructure(
			String identifierString, int identifierStringcolumnNum,
			int IdentifierStringValueColumnNum) {
		try {
			consoleLogger.debug("Tab Identifeir String: " + identifierString);
			consoleLogger.debug("Tab Identifier Column: "
					+ identifierStringcolumnNum);
			consoleLogger.debug("Tab Name Column: "
					+ IdentifierStringValueColumnNum);
			HashMap<String, HashMap<String, Object>> hashMap = new HashMap<String, HashMap<String, Object>>();
			int identifierCount = 0;
			int rowCount = 0;
			String previousKeyWord = new String();
			for (Row selectedRow : workingSheet) {
				Cell selectedCell = selectedRow
						.getCell(identifierStringcolumnNum);
				String cellStringValue = getValueFromCell(selectedCell);
				if (identifierString.equalsIgnoreCase(cellStringValue)) {
					if (identifierCount > 0) {
						HashMap<String, Object> innerhashMap = hashMap
								.get(previousKeyWord);
						innerhashMap.put("endingrow", rowCount);
					}
					HashMap<String, Object> innerHashMap = new HashMap<String, Object>();
					innerHashMap.put("startingrow", rowCount);
					String cellStringValue2 = getValueFromCell(selectedRow
							.getCell(IdentifierStringValueColumnNum));
					previousKeyWord = cellStringValue2;
					hashMap.put(cellStringValue2, innerHashMap);
					identifierCount++;
				}
				rowCount++;
			}
			HashMap<String, Object> innerHashMap = hashMap.get(previousKeyWord);
			innerHashMap.put("endingrow", rowCount + 1);
			return hashMap;
		} catch (Exception e) {
			consoleLogger.error("Error: \n" + e.getStackTrace());
		}
		return null;
	}

	/**
	 * 
	 * @param startingRow
	 * @param endingRow
	 * @param keyColumNum
	 * @param columnNums
	 * @return
	 */
	public HashMap<String, LinkedList<String>> getSelectedExcelDataIntoMap(
			int startingRow, int endingRow, int keyColumNum,
			LinkedList<Integer> columnNums) {
		HashMap<String, LinkedList<String>> keyValueMap = new HashMap<String, LinkedList<String>>();
		int rowCount = startingRow;
		for (Row row : workingSheet) {
			int rowNumber = row.getRowNum();
			if (rowCount == endingRow) {
				break;
			}
			if (rowNumber == rowCount) {
				String keyString = getValueFromExcel(rowCount, keyColumNum);
				if (keyString != null) {
					LinkedList<String> dependentStrings = new LinkedList<String>();
					for (Integer columnNum : columnNums) {
						String dependentString = getValueFromExcel(rowCount,
								columnNum);
						if (dependentString != null) {
							dependentStrings.add(dependentString);
						} else {
							dependentStrings.add("");
						}
					}
					keyValueMap.put(keyString, dependentStrings);
				}
				rowCount++;
			}
		}
		return keyValueMap;
	}

	/**
	 * 
	 * @param startingRow
	 * @param endingRow
	 * @param keyColumNum
	 * @param columnNums
	 * @return
	 */
	public HashMap<String, LinkedList<String>> getSelectedExcelDataIntoMap2(
			int startingRow, int endingRow, int keyColumNum,
			LinkedList<Integer> columnNums) {
		HashMap<String, LinkedList<String>> keyValueMap = new HashMap<String, LinkedList<String>>();
		int rowCount = startingRow;
		int kCount = 0;
		for (Row row : workingSheet) {
			int rowNumber = row.getRowNum();
			if (rowCount == endingRow) {
				break;
			}
			if (rowNumber == rowCount) {
				String keyString = getValueFromExcel(rowCount, keyColumNum);
				if (keyString != null) {
					LinkedList<String> dependentStrings = new LinkedList<String>();
					for (Integer columnNum : columnNums) {
						String dependentString = getValueFromExcel(rowCount,
								columnNum);
						if (dependentString != null) {
							dependentStrings.add(dependentString);
						} else {
							dependentStrings.add("");
						}
					}
					keyValueMap.put(String.valueOf(kCount), dependentStrings);
				}
				rowCount++;
				kCount++;
			}
		}
		return keyValueMap;
	}

	/**
	 * 
	 * @param rowIndex
	 * @param columnIndex
	 * @return
	 */
	public String getValueFromExcel(int rowIndex, int columnIndex) {
		String cellString = "";
		Row selectedRow = workingSheet.getRow(rowIndex);
		if (selectedRow != null) {
			Cell selectedCell = selectedRow.getCell(columnIndex);
			if (selectedCell != null) {
				cellString = getValueFromCell(selectedCell);
			}
		}
		if (cellString.length() == 0) {
			return "";
		} else {
			return cellString;
		}
	}

	public HashMap<Integer, String> getValuesFromSelectedColumn(int ColumnNo) {
		HashMap<Integer, String> columnNosAndColumnValues = new HashMap<Integer, String>();
		for (Row selectedRow : workingSheet) {
			int rowNum = selectedRow.getRowNum();
			Cell selectedCell = selectedRow.getCell(ColumnNo);
			if (selectedCell != null) {
				String cellValue = getValueFromCell(selectedCell);
				columnNosAndColumnValues.put(rowNum, cellValue);
			}
		}
		return columnNosAndColumnValues;
	}

	/**
	 * This method is defined to perform the operations of key, value pairs. Key
	 * stores in one column, value stores in another column
	 * 
	 * @param inputColumnIndex
	 *            - Input column number
	 * @param outputColumnIndex
	 *            - Output column number
	 * @param inputValue
	 *            - Simply it is key value to get the value information
	 * @return
	 */
	public String getVlaueFromSameRow(int inputColumnIndex,
			int outputColumnIndex, String inputValue) {
		for (Row row : workingSheet) {
			Cell cellInput = row.getCell(inputColumnIndex);
			String stringFromExcel = getValueFromCell(cellInput);
			if (inputValue.equalsIgnoreCase(stringFromExcel)) {
				Cell cellOutput = row.getCell(outputColumnIndex);
				stringFromExcel = getValueFromCell(cellOutput);
				return stringFromExcel;
			}
		}
		return "";
	}

	/**
	 * This method is defined to perform the operations of key, value pairs. Key
	 * stores in one column, value stores in another column
	 * 
	 * @param inputColumnIndex
	 *            - Input column number
	 * @param outputColumnIndex
	 *            - Output column number
	 * @param inputValue
	 *            - Simply it is key value to get the value information
	 * @return
	 */
	public String getVlaueFromSameRowBasedOnSequentialPattern(String inputValue) {
		int presentMatchedStringCount = 0;
		for (Row row : workingSheet) {
			Cell cellInput = row.getCell(2);
			String stringFromExcel = getValueFromCell(cellInput);
			if (inputValue.equalsIgnoreCase(stringFromExcel)) {
				presentMatchedStringCount++;
				Cell cellOutputInput = row.getCell(4);
				if (sequentialPatternCounts.containsKey(inputValue)) {
					sequentialPatternCount = sequentialPatternCounts
							.get(inputValue);
					sequentialPatternCounts.remove(inputValue);
				} else {
					sequentialPatternCount = 0;
				}
				stringFromExcel = getValueFromCell(cellOutputInput);
				if (sequentialPatternCount < presentMatchedStringCount) {
					sequentialPatternCount++;
					sequentialPatternCounts.put(inputValue,
							sequentialPatternCount);
					return stringFromExcel;
				}
			}
		}
		return "";
	}

	/**
	 * This method is defined to perform the operations of key, value pairs. Key
	 * stores in one column, value stores in another column
	 * 
	 * @param inputColumnIndex
	 *            - Input column number
	 * @param outputColumnIndex
	 *            - Output column number
	 * @param inputValue
	 *            - Simply it is key value to get the value information
	 * @return
	 */
	public String getVlaueFromSameRowBasedOnSequentialPattern1(String inputValue) {
		int presentMatchedStringCount = 0;
		for (Row row : workingSheet) {
			Cell cellInput = row.getCell(2);
			String stringFromExcel = getValueFromCell(cellInput);
			if (inputValue.equalsIgnoreCase(stringFromExcel)) {
				presentMatchedStringCount++;
				Cell cellOutputInput = row.getCell(5);
				if (sequentialPatternCounts1.containsKey(inputValue)) {
					sequentialPatternCount1 = sequentialPatternCounts1
							.get(inputValue);
				} else {
					sequentialPatternCount1 = 0;
					sequentialPatternCounts1.put(inputValue, 0);
				}
				stringFromExcel = getValueFromCell(cellOutputInput);
				if (sequentialPatternCount1 < presentMatchedStringCount) {
					sequentialPatternCount1++;
					return stringFromExcel;
				}
			}
		}
		return "";
	}

	/**
	 * Method to get value from cell depends on type
	 * 
	 * @param cell
	 *            - Input cell to read value
	 * @return returns cell string value, won't care about cell type
	 * @see
	 */
	public String getValueFromCell(Cell cell) {
		String cellStringValue = "";
		if (cell != null) {
			int cellType = cell.getCellType();
			switch (cellType) {
			case 0:
				cellStringValue = String.valueOf(new Double(cell
						.getNumericCellValue()).intValue());
				break;
			case 1:
				cellStringValue = cell.getStringCellValue();
				break;
			case 2:
				cellStringValue = cell.getCellFormula();
				break;
			case 3:
				cellStringValue = "";
				break;
			case 4:
				cellStringValue = String.valueOf(cell.getBooleanCellValue());
				break;
			case 5:
				cellStringValue = String.valueOf(cell.getErrorCellValue());
				break;
			default:
				cellStringValue = "";
			}
		}
		return cellStringValue;
	}

	/**
	 * Adding new value to new row in given column number
	 * 
	 * @param columnNo
	 *            where to add new value in last row
	 * @param valueString
	 *            input string value
	 */
	public void addNewVlaueToLastRow(int startingRow, int columnNo,
			String valueString) {
		int lastRowNum = workingSheet.getLastRowNum();
		if (lastRowNum == 0) {
			lastRowNum = startingRow - 1;
		}
		Row row = workingSheet.createRow(lastRowNum + 1);
		Cell varCell = row.createCell(columnNo);
		varCell.setCellValue(valueString);
	}

	/**
	 * checking that is given value is matched string or not
	 * 
	 * @param columnNo
	 * @param valueString
	 * @return
	 */
	public boolean isValueExitsInGivenCell(int rowNum, int columnNum,
			String valueString) {
		Row selectedRow = workingSheet.getRow(rowNum);
		if (selectedRow != null) {
			Cell cell = selectedRow.getCell(columnNum);
			if (cell != null) {
				String actualValue = cell.getStringCellValue();
				if (valueString.equalsIgnoreCase(actualValue))
					return true;
			}
		}
		return false;
	}

	/**
	 * finding all strings, where as they contain given pattern
	 * 
	 * @param pattern
	 * @return
	 */
	public Vector<String> searchAllKeywordsWithGivenPattern(String pattern) {
		Vector<String> matchedStrings = new Vector<String>();
		for (Row row : workingSheet) {
			for (Cell cell : row) {
				String stringValueInCell = getValueFromCell(cell);
				if (stringValueInCell.contains(pattern)) {
					matchedStrings.add(stringValueInCell);
				}
			}
		}
		return matchedStrings;
	}

	/**
	 * 
	 * @param rowNum
	 * @return
	 */
	public LinkedList<Integer> getColumnInfoFromExcel(int rowNum) {
		LinkedList<Integer> columnsInfo = new LinkedList<Integer>();
		Row selectedRow = workingSheet.getRow(rowNum);
		for (Cell selectedCell : selectedRow) {
			columnsInfo.add(selectedCell.getColumnIndex());
		}
		return columnsInfo;
	}
	
	/**
	 * 
	 * @param columnNum
	 * @param startingRowNum
	 * @return
	 */
	public LinkedList<String> getColumnInfoFromExcel(int columnNum, int startingRowNum) {
		LinkedList<String> columnsInfo = new LinkedList<String>();
		for(int i=startingRowNum; i<=workingSheet.getLastRowNum(); i++) {
			Row row = workingSheet.getRow(i);
			Cell cell = row.getCell(columnNum);
			String columnInfo = getValueFromCell(cell);
			if(columnInfo.length() != 0) {
				columnsInfo.add(columnInfo);
			}
		}
		return columnsInfo;
	}

	/**
	 * Update given cell (Row X Column) with input value
	 * 
	 * @param rowNum
	 *            - Row Position
	 * @param columnNum
	 *            - Column Position
	 * @param value
	 */
	public void updateCellValue(int rowNum, int columnNum, String value) {
		Row selectedRow = workingSheet.getRow(rowNum);
		if (selectedRow == null) {
			selectedRow = workingSheet.createRow(rowNum);
		}
		Cell selectedCell = selectedRow.getCell(columnNum);
		if (selectedCell == null) {
			selectedCell = selectedRow.createCell(columnNum);
		}
		selectedCell.setCellType(Cell.CELL_TYPE_STRING);
		selectedCell.setCellValue(value);
	}

	/**
	 * 
	 * @param cellStyle
	 * @param rowNum
	 * @param columnNum
	 * @param value
	 */
	public void updateCellValue(CellStyle cellStyle, int rowNum, int columnNum,
			String value) {
		Row selectedRow = workingSheet.getRow(rowNum);
		if (selectedRow == null) {
			selectedRow = workingSheet.createRow(rowNum);
		}
		Cell selectedCell = selectedRow.getCell(columnNum);
		if (selectedCell == null) {
			selectedCell = selectedRow.createCell(columnNum);
		}
		selectedCell.setCellType(Cell.CELL_TYPE_STRING);
		selectedCell.setCellValue(value);
		workingSheet.autoSizeColumn(columnNum);
		selectedCell.setCellStyle(cellStyle);
	}

	/**
	 * Used to get the row details
	 * 
	 * @param rowNum
	 *            , Row number integer
	 * @return row details in HashMap format
	 */
	public HashMap<String, String> getRowDetails(int rowNum) {

		HashMap<String, String> rowMap = new HashMap<String, String>();
		Row selectedRow = workingSheet.getRow(rowNum);
		for (Cell cell : selectedRow) {
			rowMap.put(String.valueOf(cell.getColumnIndex()),
					getValueFromCell(cell));
		}
		return rowMap;

	}

	public HashMap<String, String> getRowDetailsInReverseOrder(int rowNum) {

		HashMap<String, String> rowMap = new HashMap<String, String>();
		Row selectedRow = workingSheet.getRow(rowNum);
		for (Cell cell : selectedRow) {
			rowMap.put(getValueFromCell(cell),
					String.valueOf(cell.getColumnIndex()));
		}
		return rowMap;

	}

//	/**
//	 * Getting data from excel based header stings
//	 * @param headerRowInfo
//	 * @param startingRow
//	 * @return
//	 */
//	public LinkedList<GraphObject> getExcelDataIntoGraphObjects(
//			HashMap<String, String> headerRowInfo, int startingRow) {
//		LinkedList<GraphObject> graphObjects = new LinkedList<GraphObject>();
//		int count = 1;
//		for (Row selectedRow : workingSheet) {
//			if (count >= startingRow) {
//				GraphObject graphObject = new GraphObject();
//				Properties properties = graphObject.getObjectProperties();
//				Set<String> headerStringsSet = headerRowInfo.keySet();
//				Iterator<String> headerStringIterator = headerStringsSet.iterator();
//				while (headerStringIterator.hasNext()) {
//					String headerString = headerStringIterator.next();
//					String valueString = headerRowInfo.get(headerString);
//					Cell selectedCell = selectedRow.getCell(Integer.parseInt(valueString));
//					String cellValue = getValueFromCell(selectedCell);
//					properties.put(headerString, cellValue);
//				}
//				graphObjects.add(graphObject);
//			}
//			count++;
//		}
//		return graphObjects;
//	}
	
	public LinkedList<String> getExcelData(
			HashMap<String, String> headerRowInfo, int startingRow) {
		LinkedList<String> excelData = new LinkedList<String>();
		int count = 1;
		LinkedList<String> headerStrings = new LinkedList<String>();
		for (Row selectedRow : workingSheet) {
			if (count >= startingRow) {
				String rowString = "";
				for (Cell selectedcell : selectedRow) {
					String columnValue = String.valueOf(selectedcell
							.getColumnIndex());
					String modifiedColumnValue = "<"
							+ headerRowInfo.get(columnValue) + ">";
					if (isValuePresent(headerStrings, modifiedColumnValue)) {
						headerStrings.remove(modifiedColumnValue);
					}
					rowString += modifiedColumnValue
							+ getValueFromCell(selectedcell) + ": +";
					headerStrings.add(rowString);
					// System.out.println(rowString);
					break;
				}
				String finalString = "";
				for (String string : headerStrings) {
					finalString += string;
				}
				excelData.add(finalString + rowString);
			}
			count++;
		}
		return excelData;
	}

	public org.w3c.dom.Document getExcelDatatoXML(
			HashMap<String, String> headerRowInfo, int startingRow) {
		try {
			DocumentBuilderFactory documentBuilderFactory = DocumentBuilderFactory
					.newInstance();
			DocumentBuilder documentBuilder = documentBuilderFactory
					.newDocumentBuilder();
			org.w3c.dom.Document document = documentBuilder.newDocument();
			Element excelElement = document.createElement("N-1:ExcelDocument");
			document.appendChild(excelElement);
			int previousRowColumnNo = -1;
			int count = 1;
			for (Row selectedRow : workingSheet) {
				if (count >= startingRow) {
					String rowString = "";
					for (Cell selectedcell : selectedRow) {
						int columnNo = selectedcell.getColumnIndex();
						String columnValue = String.valueOf(columnNo);
						String elementName = headerRowInfo.get(columnValue);
						System.out.println(columnNo + "--" + elementName);
						elementName = elementName.replace(" ", "_");
						elementName = elementName.replace("/", "_");
						// System.out.println("hi "+elementName);
						elementName = elementName.trim();
						while (true) {
							if (columnNo > previousRowColumnNo) {
								String cellValue = getValueFromCell(selectedcell);
								if (cellValue.length() > 0) {
									Element newElement = document
											.createElement("N" + columnNo + ":"
													+ elementName);
									newElement.setTextContent(cellValue);
									excelElement.appendChild(newElement);
									excelElement = newElement;
									previousRowColumnNo = columnNo;
								}
								break;
							} else if (columnNo == previousRowColumnNo) {

								String cellValue = getValueFromCell(selectedcell);
								if (cellValue.length() > 0) {
									excelElement = (Element) excelElement
											.getParentNode();
									Element newElement = document
											.createElement("N" + columnNo + ":"
													+ elementName);
									newElement.setTextContent(cellValue);
									excelElement.appendChild(newElement);
									excelElement = newElement;
									previousRowColumnNo = columnNo;
								}
								break;
							} else {
								excelElement = (Element) excelElement
										.getParentNode();
								previousRowColumnNo = Integer
										.parseInt(excelElement.getTagName()
												.split(":")[0].replace("N", ""));
							}
						}
					}
				}
				count++;
			}
			return document;
		} catch (ParserConfigurationException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (NullPointerException e) {
			e.printStackTrace();
		}
		return null;
	}

	/**
	 * 
	 * @param startingRow
	 * @param endingRow
	 * @param identifierString
	 * @return
	 */
	public String getValueFromExcelBasedOnMatchedString(int startingRow,
			int endingRow, String identifierString) {
		String returnString = "";
		for (int count = startingRow; count < endingRow; count++) {
			Row selectedRow = workingSheet.getRow(count);
			for (Cell cell : selectedRow) {
				String cellValue = getValueFromCell(cell);
				if (cellValue.contains(identifierString)) {
					returnString = cellValue.split(identifierString)[1];
					// returnString.replace(":", "");
				}
			}
		}
		return returnString;
	}

	/**
	 * 
	 * @param keywordRowValue
	 * @param keywordColumnValue
	 * @param valueRowValue
	 * @param valueClumnValue
	 * @param identifierString
	 * @return
	 */
	public String getValueFromExcelBasedOnKeyword(int keywordRowValue,
			int keywordColumnValue, int valueRowValue, int valueClumnValue,
			String identifierString) {

		return "";
	}

	/**
	 * 
	 * @param allStrings
	 * @param findValue
	 * @return
	 */
	public boolean isValuePresent(LinkedList<String> allStrings,
			String findValue) {
		for (String string : allStrings) {
			if (string.contains(findValue)) {
				return true;
			}
		}
		return false;
	}

	public void printExcelData() {
		for (Row selectedRow : workingSheet) {
			for (Cell cell : selectedRow) {
				System.out.print(getValueFromCell(cell).trim() + " ");
			}
			System.out.print("\n--- link ---\n");
		}
	}

	public HashMap<String, String> getMergedCellInformation(int columnNo) {
		HashMap<String, String> regionMap = new HashMap<String, String>();
		int totalMergerRegionCount = workingSheet.getNumMergedRegions();
		// System.out.println("No of merged regions: "+workingSheet.getNumMergedRegions());
		for (int i = 0; i < totalMergerRegionCount; i++) {
			int mergedFirstColumnNumber = workingSheet.getMergedRegion(i)
					.getFirstColumn();
			if (mergedFirstColumnNumber == columnNo) {
				regionMap.put(String.valueOf(workingSheet.getMergedRegion(i)
						.getFirstRow()), String.valueOf(workingSheet
						.getMergedRegion(i).getLastRow()));
				// System.out.println(workingSheet.getMergedRegion(i).getFirstRow()+"---"+workingSheet.getMergedRegion(i).getLastRow());
				// System.out.println(getValueFromExcel(workingSheet.getMergedRegion(i).getFirstRow(),
				// columnNo));
			}
		}
		return regionMap;
	}

}
