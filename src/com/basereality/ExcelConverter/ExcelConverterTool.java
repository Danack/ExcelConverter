package com.basereality.ExcelConverter;


import java.io.*;
import java.util.Vector;

import net.sf.json.JSONArray;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;


public class ExcelConverterTool {


	private static String filename = null;

	private static String numberMode = null;

	/**
	 * List of all the items that can be passed in to the program, java can't switch on strings properly.
	 */
	public enum ArgumentType {
		NONE,
		FILENAME,
		MAXLINES,
		MAXDATA,
		NUMBERFORMAT,
	}

	public static final String ARGUMENT_FILE = "-file";
	public static final String ARGUMENT_FILE_BRIEF = "-f";

	public static final String ARGUMENT_MAX_LINES = "-maxLines";
	public static final String ARGUMENT_MAX_DATA = "-maxData";
	public static final String ARGUMENT_NUMBER_FORMAT = "-number";


	/**
	 * Run the excel reader.
	 *
	 * @param args The args to parse
	 */
	public static void main(String[] args) {

		try {
			ExcelConverterTool.readArguments(args);

			if (filename == null) {
				System.out.println("Failed to read the filename from the arguments.");
				System.exit(-2);
			}

			Vector<Vector<String>> stringsFromSpreadSheet = ExcelConverterTool.processFile();
			ExcelConverterTool.exportData(stringsFromSpreadSheet);
		} catch (InvalidDataException ide) {
			System.out.println("InvalidDataException processing file [" + filename + "]" + ide.getMessage());
			System.exit(-1);
		} catch (Exception e) {
			System.out.println("Unexpected exception of class [" + e.getClass() + "] processing file [" + filename + "]");
			System.out.println("Exception is: " + e.getMessage());

			e.printStackTrace(System.out);
			System.exit(-2);
		}
	}

	/**
	 * Process the file.
	 *
	 * @return
	 * @throws FileNotFoundException
	 * @throws InvalidDataException
	 * @throws IOException
	 * @throws InvalidFormatException
	 */
	public static Vector<Vector<String>> processFile() throws FileNotFoundException, IOException, InvalidFormatException, InvalidDataException {

		Vector<Vector<String>> stringsFromSpreadSheet = null;
		boolean attemptToReadFileAsCSV = false;

		System.err.println("Opening file [" + filename + "]");

		InputStream inputStream = new FileInputStream(filename);

		ExcelConverter excelConverter = new ExcelConverter();

		if (numberMode != null) {
			excelConverter.setFormulaNumberMode(numberMode);
		}

		try {//We always try to read the file as an good OLE stream as some suppliers have the wrong
			//file extension e.g. the file has an extension xls, but is actually CSV...
			stringsFromSpreadSheet = excelConverter.readInputStreamAsOLE(inputStream, 0);
		} catch (IllegalArgumentException iae) {
			System.err.println("Failed to parse the the stream as an OLE, so going to try as csv." + iae.getMessage());
			attemptToReadFileAsCSV = true;
		} finally {
			inputStream.close();
		}

		if (attemptToReadFileAsCSV == true) {

			inputStream = new FileInputStream(filename);

			try {
				stringsFromSpreadSheet = excelConverter.readInputStreamAsCSV(inputStream, 0);
			} catch (IllegalArgumentException iae) {
				System.err.println("Failed to parse the the stream as CSV, so going to throw an exception.");
				throw iae;
			} finally {
				inputStream.close();
			}
		}

		return stringsFromSpreadSheet;
	}

	/**
	 * Read any and all arguments from the command line.
	 *
	 * @param args
	 */
	public static void readArguments(String[] args) {

		ArgumentType nextArgumentType = ArgumentType.NONE;

		for (String argValue : args) {

			boolean argumentMustBeFlag = false;

			switch (nextArgumentType) {

				case FILENAME: {
					filename = argValue;
					nextArgumentType = ArgumentType.NONE;
					break;
				}

				case MAXLINES: {
					int maxLines = Integer.parseInt(argValue);
					if (maxLines <= 0) {
						throw new UnsupportedOperationException("Invalid max lines value [" + argValue + "]");
					}
					ExcelConverter.setMaxLines(maxLines);
					nextArgumentType = ArgumentType.NONE;
					break;
				}

				case MAXDATA: {
					int maxData = Integer.parseInt(argValue);
					if (maxData <= 0) {
						throw new UnsupportedOperationException("Invalid max data value [" + argValue + "]");
					}
					ExcelConverter.setMaxData(maxData);
					nextArgumentType = ArgumentType.NONE;
					break;
				}

				case NUMBERFORMAT: {

					if ("i".compareToIgnoreCase(argValue) == 0) {
						numberMode = "integer";
					} else if ("f".compareToIgnoreCase(argValue) == 0) {
						numberMode = "float";
					} else if ("d".compareToIgnoreCase(argValue) == 0) {
						numberMode = "double";
					}
					nextArgumentType = ArgumentType.NONE;
					break;
				}

				default: {
					argumentMustBeFlag = true;
					break;
				}
			}

			if (argumentMustBeFlag == true) {
				if (argValue.compareTo(ARGUMENT_FILE) == 0 ||
						argValue.compareTo(ARGUMENT_FILE_BRIEF) == 0) {
					nextArgumentType = ArgumentType.FILENAME;
				} else if (argValue.compareTo(ARGUMENT_MAX_LINES) == 0) {
					nextArgumentType = ArgumentType.MAXLINES;
				} else if (argValue.compareTo(ARGUMENT_MAX_DATA) == 0) {
					nextArgumentType = ArgumentType.MAXDATA;
				} else if (argValue.compareTo(ARGUMENT_NUMBER_FORMAT) == 0) {
					nextArgumentType = ArgumentType.NUMBERFORMAT;
				} else {
					System.out.println("Unknown argument " + argValue);
					System.out.println("Try something more like:");
					System.out.println("java -jar ExcelConverter.jar -file filename [-maxLines maxLinesInFile][-maxData maxDataSize][-number integer|bstool|float]");
					System.exit(-1);
				}
			}
		}

		if (nextArgumentType != ArgumentType.NONE) {
			System.out.println("Missing last argument " + nextArgumentType.toString());
			System.exit(-1);
		}

		if (filename == null) {
			System.out.println("filename is not set - please set via command -file");
			System.exit(-1);
		}
	}

	/**
	 * Export the data as a big old json array.
	 *
	 * @param sheetContents
	 */
	public static void exportData(Vector<Vector<String>> sheetContents) {
		JSONArray json = new JSONArray();
		json.addAll(sheetContents);
		//json.setExpandElements(true);
		System.out.println(json.toString());
		System.exit(1);
}
}
