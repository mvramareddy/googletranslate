package com.gitdemo.trans;

import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Pattern;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.AppendCellsRequest;
import com.google.api.services.sheets.v4.model.BatchUpdateSpreadsheetRequest;
import com.google.api.services.sheets.v4.model.CellData;
import com.google.api.services.sheets.v4.model.ExtendedValue;
import com.google.api.services.sheets.v4.model.GridProperties;
import com.google.api.services.sheets.v4.model.Request;
import com.google.api.services.sheets.v4.model.RowData;
import com.google.api.services.sheets.v4.model.Sheet;
import com.google.api.services.sheets.v4.model.SheetProperties;
import com.google.api.services.sheets.v4.model.Spreadsheet;
import com.google.api.services.sheets.v4.model.SpreadsheetProperties;

public class DemoTranslate {

	public static void sheetOperations(Sheets sheetsService, String inputFileName, int batchSize) throws Exception {
		process(sheetsService, inputFileName, batchSize);

	}

	public static void process(Sheets sheetsService, String inputFileName, int batchSize) throws IOException {

		List<Sheet> s = new ArrayList<Sheet>();
		s.add(new Sheet().setProperties(new SheetProperties().setTitle("DemoTrans-Sheet1").setSheetId(0)
				.setGridProperties(new GridProperties().setColumnCount(4))));

		Timestamp ts = new Timestamp(System.currentTimeMillis());
		SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");
		String timestamp = sdf.format(ts);
		Spreadsheet requestBody = new Spreadsheet();

		Sheets.Spreadsheets.Create request = sheetsService.spreadsheets().create(requestBody
				.setProperties(new SpreadsheetProperties().setTitle("Demo_Translate_" + timestamp)).setSheets(s));

		Spreadsheet response = request.execute();
		String id = response.getSpreadsheetId();

		List<RowData> rows = new ArrayList<RowData>();
		List<Request> requests = null;

		String csvFile = inputFileName;
		String line = "";
		int endLine = 0;

		int totalCount = 0;
		AppendCellsRequest appendCellReq = new AppendCellsRequest();

		try (BufferedReader br = new BufferedReader(new FileReader(csvFile))) {
			requests = new ArrayList<>();
			String pattern = ".*_x.*_";
			while ((line = br.readLine()) != null) {
				List<CellData> values = new ArrayList<>();
				RowData row = new RowData();
				totalCount++;
				String[] data = line.split(",");
				String splitArray[] = null;
				StringBuilder sb = new StringBuilder();
				if (Pattern.matches(pattern, data[1])) {
					splitArray = data[1].split("_x");

					for (int i = 0; i < splitArray.length; i++) {

						if (splitArray[i].contains("_")) {
							char c = (char) Integer.parseInt(splitArray[i].substring(0, splitArray[i].indexOf("_")),
									16);
							sb.append(c);
							System.out.println("Special Chracter......................" + c);
						} else {
							sb.append(splitArray[i]);
							System.out.println("Normal Value is..............." + splitArray[i]);
						}

					}

				} else {
					sb.append(data[1]);
				}

				values.add(new CellData().setUserEnteredValue(new ExtendedValue().setStringValue(data[0])));
				values.add(new CellData().setUserEnteredValue(new ExtendedValue().setStringValue(sb.toString())));
				values.add(new CellData().setUserEnteredValue(new ExtendedValue().setStringValue(data[2])));
				values.add(new CellData()
						.setUserEnteredValue(new ExtendedValue().setFormulaValue(data[3].replace("|", ","))));

				row.setValues(values);

				endLine = endLine + 1;
				rows.add(row);

				appendCellReq.setRows(rows);
				appendCellReq.setFields("userEnteredValue,userEnteredFormat.numberFormat");

				if (endLine % batchSize == 0) {
					endLine = 0;

					requests.add(new Request().setAppendCells(appendCellReq));

					BatchUpdateSpreadsheetRequest batchUpdateRequest = new BatchUpdateSpreadsheetRequest()
							.setRequests(requests);
					sheetsService.spreadsheets().batchUpdate(id, batchUpdateRequest).execute();

					System.out.println("INFO :: Total Records Processed -->" + totalCount + "\n");
					rows = new ArrayList<RowData>();
					requests = new ArrayList<>();

					System.out
							.println("INFO :: Next Batch Starting At -->" + new Timestamp(System.currentTimeMillis()));
					System.out.println("INFO :: Processing For Cell Range --> " + "A" + totalCount + ":D"
							+ (totalCount + batchSize) + "\n\n");

				}

			}
			requests.add(new Request().setAppendCells(appendCellReq));

			BatchUpdateSpreadsheetRequest batchUpdateRequest = new BatchUpdateSpreadsheetRequest()
					.setRequests(requests);
			sheetsService.spreadsheets().batchUpdate(id, batchUpdateRequest).execute();
			System.out.println("INFO :: Total Records Processed, Final Count-->" + totalCount);
			

		} catch (Exception e) {
			e.printStackTrace();
		}

	}

}
