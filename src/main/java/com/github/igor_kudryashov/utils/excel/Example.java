package com.github.igor_kudryashov.utils.excel;

import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;

public class Example {
	public static void main(String[] args) {
		// create an Excel workbook
		ExcelWriter writer = new ExcelWriter();
		// create an Excel worksheet
		Sheet sheet1 = writer.createSheet("Sheet1");
		// create header of table
		writer.createRow(sheet1, new String[] { "Column 1", "Column 2", "Column 3" }, true);
		for (int x = 0; x < 3; x++) {
			// create simple row
			writer.createRow(sheet1, new Object[] { "Cell 1", "Cell 2", "Cell 3" }, false);
		}
		writer.setAutoSizeColumns(sheet1, true);
		Sheet sheet2 = writer.createSheet("Sheet2");
		// create header of table
		writer.createRow(sheet2, new String[] { "Column 1", "Column 2", "Column 3" }, true);
		for (int x = 0; x < 3; x++) {
			// create simple row
			writer.createRow(sheet2, new Object[] { "Cell 1", "Cell 2", "Cell 3" }, false);
		}
		writer.setAutoSizeColumns(sheet2, true);
		try {
			// save an Excel workbook file
			writer.saveToFile("file.xlsx");
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}
