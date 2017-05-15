package com.gitdemo.operations;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataMoveOperation {

	public static void readXLSXFile(String inputFileName, String outputFileName) throws IOException {

		int csvFlag = 0;
		int rowCount = 0;
		int batchVal = 0;
		StringBuffer data = new StringBuffer();
		InputStream ExcelFileToRead = new FileInputStream(inputFileName);
		XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);
		XSSFSheet sheet = wb.getSheetAt(0);
		XSSFRow row;
		XSSFCell cell;

		Iterator<Row> rowIterator = sheet.rowIterator();

		while (rowIterator.hasNext()) {
			rowCount = rowCount + 1;
			batchVal = batchVal + 1;

			row = (XSSFRow) rowIterator.next();

			Iterator<Cell> cellIterator = row.cellIterator();
			while (cellIterator.hasNext()) {
				cell = (XSSFCell) cellIterator.next();

				switch (cell.getCellType()) {
				case Cell.CELL_TYPE_BOOLEAN:
					data.append(cell.getBooleanCellValue() + ",");
					break;

				case Cell.CELL_TYPE_NUMERIC:
					data.append(cell.getNumericCellValue() + ",");
					break;

				case Cell.CELL_TYPE_STRING:
					data.append(cell.getStringCellValue().replace(",", " ") + ",");
					break;

				case Cell.CELL_TYPE_BLANK:
					data.append("" + ",");
					break;

				default:
					data.append(cell + ",");
				}

			}
			data.append("=GOOGLETRANSLATE(B" + rowCount + "|\"auto\"|\"en\")");
			data.append("\n");

			if (csvFlag == 0 && batchVal == 5000) {
				createCSVFile(outputFileName);
				writeCSVFile(outputFileName,data.toString());
				data = new StringBuffer();
				batchVal = 0;
				csvFlag = 1;
			} else if (csvFlag != 0 && batchVal == 5000) {
				writeCSVFile(outputFileName,data.toString());
				data = new StringBuffer();
				batchVal = 0;
			}

		}
		writeCSVFile(outputFileName,data.toString());

	}

	public static void writeCSVFile(String csvFilename,String content) {
		BufferedWriter bw = null;
		FileWriter fw = null;

		try {

			File fileDir = new File(csvFilename);

			FileOutputStream out = new FileOutputStream(fileDir,true);
			out.write(content.getBytes("UTF-8"));

			out.flush();
			out.close();

		} catch (IOException e) {

			e.printStackTrace();

		} finally {

			try {

				if (bw != null)
					bw.close();

				if (fw != null)
					fw.close();

			} catch (IOException ex) {

				ex.printStackTrace();

			}

		}
	}

	public static void createCSVFile(String csvFileName) {

		try {

			File file = new File(csvFileName);

			file.createNewFile()

		} catch (IOException e) {
			e.printStackTrace();

		}
	}



}