package com.basereality.ExcelConverter;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;

import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Locale;
import java.util.StringTokenizer;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;


import java.util.Vector;

import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.DataFormatter;


public class ExcelConverter {

	/**
	 * Get a commons logger, it will use log4j if possible
	 */
	protected Logger logger = LoggerFactory.getLogger(getClass());

	private Vector<Vector<String>> sheetContents = new Vector<Vector<String>>();

	private static final int maxLines = 5000;

	private static final int maxData = maxLines * 200;

	private static FormulaEvaluator evaluator = null;

	private boolean removeEmptyColumns = false;

	private boolean skipEmptyRows = true;

	private static final String decimalNumberFormat = "######0.0000";

	private DecimalFormat decimalFormat = null;

	// Create a formatter, do this once
	private DataFormatter formatter = new DataFormatter(Locale.UK);


	public ExcelConverter() {
		setFormulaNumberMode("decimal");
	}

	/**
	 * Set how formulas that resolve to Excel 'numbers' are evaluated into Java type numbers.
	 *
	 * @param numberMode - how to interpret numbers from formulas
	 *	float
	 *  double
	 *  integer - all decimal places are thrown away
	 *  decimal - numbers are rounded to 4 decimal places
	 */
	public void setFormulaNumberMode(String numberMode) {

		if ("float".compareToIgnoreCase(numberMode) == 0) {
			throw new UnsupportedOperationException("Formatting formula to float not implemented yet.");
		} else if ("double".compareToIgnoreCase(numberMode) == 0) {
			throw new UnsupportedOperationException("Formatting formula to float not implemented yet.");
		} else if ("integer".compareToIgnoreCase(numberMode) == 0) {
			decimalFormat = new DecimalFormat("######0");
		} else if ("decimal".compareToIgnoreCase(numberMode) == 0) {
			decimalFormat = new DecimalFormat(decimalNumberFormat);
		} else {
			throw new UnsupportedOperationException("Unrecognized number mode [" + numberMode + "] for parsing.");
		}
	}

	/**
	 * Read an OLE input stream, parse it through the appropriate Apache POI reader and return it as a 2d array of strings that represent the contents.
	 *
	 * @param inputStream
	 * @param sheetNumber
	 * @return
	 * @throws InvalidDataException
	 * @throws IOException
	 * @throws InvalidFormatException
	 */
	public Vector<Vector<String>> readInputStreamAsOLE(InputStream inputStream, int sheetNumber) throws IOException, InvalidFormatException, InvalidDataException {

		Workbook workbook = org.apache.poi.ss.usermodel.WorkbookFactory.create(inputStream);

		workbook.setMissingCellPolicy(Row.CREATE_NULL_AS_BLANK);

		int totalData = 0;

		Sheet sheet = workbook.getSheetAt(0);

		evaluator = workbook.getCreationHelper().createFormulaEvaluator();

		int rowCount = 0;

		for (Row row : sheet) {

			Vector<String> rowContents = new Vector<String>();

			int cellCount = 0;

			if (rowCount > maxLines) {
				throw new InvalidDataException("Data file has too many rows, maxLines set at " + ExcelConverter.maxLines + " and the file has " + rowCount + " rows.");
			}

			for (Cell cell : row) {
				String[] cellContentsArray = getCellContents(cell, cellCount, rowCount);

				if (cellContentsArray != null) {

					for (String cellContents : cellContentsArray) {
						rowContents.add(cellContents);
						cellCount++;

						totalData += cellContents.length();

						if (totalData > ExcelConverter.maxData) {
							throw new InvalidDataException("Data file is too large, maxSize set at " + ExcelConverter.maxData);
						}
					}
				} else {
					rowContents.add("");
					cellCount++;
				}
			}

			if(skipEmptyRows == true && rowContents.size() == 0){
				//Skip this line as it's empty
			}
			else{
				sheetContents.add(rowContents);
			}

			rowCount++;
		}

		if (removeEmptyColumns == true) {
			return purgeEmptyColumns(sheetContents);
		}

		return sheetContents;
	}

	/**
	 * Remove all empty columns from the imported data. This is to make the data import look
	 * nicer and be easier to understand when there are multiple empty columns, which is confusing.
	 *
	 * @param sheetContents
	 * @return
	 */
	public Vector<Vector<String>> purgeEmptyColumns(Vector<Vector<String>> sheetContents) {

		Vector<Vector<String>> newSheetContents = new Vector();

		HashMap<Integer, Boolean> validColumns = new HashMap();

		//Find all the column which are empty in all rows,
		//by looping over all cells, and marking the valid columns
		for (Vector<String> rowContents : sheetContents) {
			int columnCounter = 0;

			for (String cell : rowContents) {//Loop over all cells
				if (cell != null && cell.trim().length() != 0) {
					//It's a valid column
					validColumns.put(columnCounter, Boolean.TRUE);
				}

				columnCounter++;
			}
		}

		//Extract all the valid contents by copying them to a new structure
		for (Vector<String> rowContents : sheetContents) {
			Vector<String> purgedRowContents = new Vector();

			int columnCounter = 0;

			for (String cell : rowContents) {//Loop over all cells

				Boolean isValidColumn = validColumns.get(columnCounter);
				if (isValidColumn == null || isValidColumn.booleanValue() == false) {
					//Skip this column
				} else {
					purgedRowContents.add(cell);
				}

				columnCounter++;
			}

			newSheetContents.add(purgedRowContents);
		}

		return newSheetContents;
	}


	/**
	 * Read an inputstream, parse it as CSV data and return a 2d vector of the strings.
	 *
	 * @param inp
	 * @param sheetNumber
	 * @return
	 * @throws IOException
	 */
	public Vector<Vector<String>> readInputStreamAsCSV(InputStream inp, int sheetNumber) throws IOException {

		BufferedReader bufferedReader = new BufferedReader(new InputStreamReader(inp));

		String strLine = null;
		StringTokenizer st = null;

		int rowCount = 0;

		char[] allowedSeparators = {
				';',
				'\t',
				',',
		};

		char bestSeparator = '.';
		int bestSeparatorScore = 0;

		while ((strLine = bufferedReader.readLine()) != null) {

			Vector<String> rowContents = new Vector<String>();

			for (char separator : allowedSeparators) {
				int separatorScore = ExcelConverterUtil.getNumberTimesCharacterOccursInString(separator, strLine);

				if (separatorScore > bestSeparatorScore) {
					bestSeparator = separator;
					bestSeparatorScore = separatorScore;
				}
			}

			if (bestSeparatorScore == 0) {
				bestSeparator = ',';    //Blank lines should not explode the parser.
			}

			ArrayList<String> rowCells = ExcelConverterUtil.parseCSV(strLine, "" + bestSeparator);

			for (String string : rowCells) {
				rowContents.add(string);
			}

			if(skipEmptyRows == true && rowContents.size() == 0){
				//Skip this line as it's empty
			}
			else{
				sheetContents.add(rowContents);
			}

			rowCount++;
		}

		bufferedReader.close();

		if (removeEmptyColumns == true) {
			return purgeEmptyColumns(sheetContents);
		}
		return sheetContents;
	}


	/**
	 * Get the contents of an OLE type cell as well as all the required empty cells before this cell that
	 * need to be generated.
	 * This is copied from the internet and is assumed to be correct.
	 *
	 * @param cell
	 * @param columnCount
	 * @return An array of strings
	 */
	public String[] getCellContents(Cell cell, int columnCount, int rowCount) {

		//Excel skips over the first columns if they are blank.
		//Need to generate some empty columns
		int columnsToGenerate = cell.getColumnIndex() - columnCount + 1;

		String[] retStrings = new String[columnsToGenerate];

		if (columnsToGenerate <= 0) {
			logger.warn("Hmm, columnsToGenerate is " + columnsToGenerate + ", so not doing that and just returning a single blank column.");
			String[] emptyArray = {""};
			return emptyArray;
		}

		for (int i = 0; i < columnsToGenerate - 1; i++) {
			retStrings[i] = "";
		}

		switch (cell.getCellType()) {
			case Cell.CELL_TYPE_STRING: {
				retStrings[columnsToGenerate - 1] = cell.getRichStringCellValue().getString();
				break;
			}
			case Cell.CELL_TYPE_NUMERIC: {
				if (DateUtil.isCellDateFormatted(cell)) {
					retStrings[columnsToGenerate - 1] = cell.getDateCellValue().toString();
					break;
				} else {
					//Format the cell as excel would format it.
					retStrings[columnsToGenerate - 1] = formatter.formatCellValue(cell);
					break;
				}
			}
			case Cell.CELL_TYPE_BOOLEAN: {
				retStrings[columnsToGenerate - 1] = String.valueOf(cell.getBooleanCellValue());
				break;
			}
			case Cell.CELL_TYPE_FORMULA: {

				retStrings[columnsToGenerate - 1] = cell.getCellFormula(); //Get the original formula first, so if we totally
				//fail to parse the formula, at least there's something there.
				try {
					//Please note - some formulas are not supported in Apache POI
					//http://www.chipkillmar.net/2008/04/28/evaluating-excel-formulas-with-apache-poi/
					CellValue cellValue = evaluator.evaluate(cell);
					retStrings[columnsToGenerate - 1] = parseCellValue(cellValue);
				} catch (Exception e) {
					logger.debug("Exception of class " + e.getClass().toString() + " thrown when evaluating formula. Exception is: " + e.toString());
				}

				break;
			}

			case Cell.CELL_TYPE_BLANK: {
				retStrings[columnsToGenerate - 1] = "";
				break;
			}

			default: {
				logger.info("skipping cell type " + cell.getCellType());
				//Could throw an exception here as we're skipping data.
				return null;
			}
		}

		return retStrings;
	}

	/**
	 * Formulas get converted to CellValues, and then we convert the CellValue to a string.
	 *
	 * @param cellValue The cellValue generated by the formulaEvaluator
	 * @return
	 */
	public String parseCellValue(CellValue cellValue) {

		switch (cellValue.getCellType()) {

			case (Cell.CELL_TYPE_NUMERIC): {
				double cellValueAsDouble = cellValue.getNumberValue();
				String cellString = "" + cellValueAsDouble;

				try {
					cellString = decimalFormat.format(cellValueAsDouble);
				} catch (Exception e) {
					//I guess it wasn't a valid number.
				}

				return cellString;
			}

			case (Cell.CELL_TYPE_STRING): {
				return cellValue.getStringValue();
			}

			case (Cell.CELL_TYPE_FORMULA): {
				//Recursion detected - formulas should not evaluate to a formula
				return "";
			}

			case (Cell.CELL_TYPE_BLANK): {
				return "";
			}

			case (Cell.CELL_TYPE_BOOLEAN): {
				if (cellValue.getBooleanValue() == true) {
					return "true";
				}
				return "false";
			}

			case (Cell.CELL_TYPE_ERROR): {
				//Probably an error in the formula - what to do?
				//We aren't an editor do nothing sensible to display, lets just return a blank string.
				return "";
			}

			default: {
				//Unknown type - could throw error here.
				return "";
			}
		}
	}
}
