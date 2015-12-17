/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

package com.github.igor_kudryashov.utils.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

public class ExcelWriter {

	private SXSSFWorkbook workbook;
	private SXSSFSheet sheet;
	private int rowNum = 0;
	private String sheetName;
	private int lastCellNum = 0;
	// min colunm width (pic)
	private int MIN_WIDTH = 3000;
	// max column width (pic)
	private int MAX_WIDTH = 15000;
	// the styles of cells
	Map<String, XSSFCellStyle> styles = new HashMap<String, XSSFCellStyle>();
	// the width of columns
	Map<Integer, Integer> columnWidth = new HashMap<Integer, Integer>();

	
	/**
	 * Class constructor
	 * @param name - name of sheet or <code>null</code> then the sheet will not be created.
	 */
	public ExcelWriter(String name) {

		workbook = new SXSSFWorkbook();
		// temp files will be gzipped
		workbook.setCompressTempFiles(true);

		if (name != null) {
			sheet = createSheet(name);
		}
	}

	/**
	 * Create sheet with given name 
	 * @param name - name of the sheet.
	 * @return created sheet.
	 */
	public SXSSFSheet createSheet(String name) {
		sheetName = name.replaceAll("\\\\", "-").trim();
		sheet = workbook.createSheet(sheetName);
		return sheet;
	}

	/**
	 * Sets autofilter header and autosize of columns's width   
	 * @param withHeader - <code>true</code> for sets autofilter header or <code>false</code> otherwise.
	 * @throws IOException
	 */
	public void setAutoSizeColumns(boolean withHeader) throws IOException {

		// flushRows();

		if (withHeader) {
			org.apache.poi.ss.util.CellRangeAddress range = new org.apache.poi.ss.util.CellRangeAddress(0, 0, 0,
					lastCellNum - 1);
			sheet.setAutoFilter(range);
			sheet.createFreezePane(0, 1);
		}

		for (Map.Entry<Integer, Integer> entry : columnWidth.entrySet()) {
			sheet.setColumnWidth(entry.getKey(), entry.getValue());
		}

	}

	/**
	 * It creates a style for the object of the specified type and saves the style in collection.
	 * @param entry 
	 * @param header - <code>true</code> for the header or <code>false</code> if not.
	 * @return created style.
	 */
	private XSSFCellStyle getCellStyle(Object entry, boolean header) {
		XSSFCellStyle style = null;
		String name = entry.getClass().getName();
		if (header) {
			name += "_header";
		} else if ((rowNum % 2) == 0) {
			name += "_even";
		}
		if (styles.containsKey(name)) {
			style = styles.get(name);
		} else {
			style = (XSSFCellStyle) workbook.createCellStyle();

			style.setVerticalAlignment(VerticalAlignment.TOP);

			// ligth grey color
			XSSFColor borderColor = new XSSFColor(new java.awt.Color(192, 192, 192));
			style.setBorderBottom(CellStyle.BORDER_THIN);
			style.setBorderColor(XSSFCellBorder.BorderSide.BOTTOM, borderColor);
			style.setBorderLeft(CellStyle.BORDER_THIN);
			style.setBorderColor(XSSFCellBorder.BorderSide.LEFT, borderColor);
			style.setBorderRight(CellStyle.BORDER_THIN);
			style.setBorderColor(XSSFCellBorder.BorderSide.RIGHT, borderColor);

			XSSFDataFormat fmt = (XSSFDataFormat) workbook.createDataFormat();
			short format = 0;
			if (name.contains("Date")) {
				format = fmt.getFormat(BuiltinFormats.getBuiltinFormat(0xe));
				style.setAlignment(CellStyle.ALIGN_LEFT);
			} else if (name.contains("Double")) {
				format = fmt.getFormat(BuiltinFormats.getBuiltinFormat(2));
				style.setAlignment(CellStyle.ALIGN_RIGHT);
			} else if (name.contains("Integer")) {
				format = fmt.getFormat(BuiltinFormats.getBuiltinFormat(1));
				style.setAlignment(CellStyle.ALIGN_RIGHT);
			} else {
				style.setAlignment(CellStyle.ALIGN_LEFT);
				if (!header) {
					style.setWrapText(true);
				}
			}
			if (header) {
				// gray color
				XSSFColor color = new XSSFColor(new java.awt.Color(210, 210, 210));
				style.setFillForegroundColor(color);
				style.setFillPattern(CellStyle.SOLID_FOREGROUND);
			} else if (name.contains("_even")) {
				// ligth grey color
				XSSFColor color = new XSSFColor(new java.awt.Color(239, 239, 239));
				style.setFillForegroundColor(color);
				style.setFillPattern(CellStyle.SOLID_FOREGROUND);
			}
			style.setDataFormat(format);
			styles.put(name, style);
		}
		return style;
	}

	/**
	 * Create row of sheet
	 * @param values - the array of values for the row.
	 * @param header <code>true</code> if row is an header or <code>false</code> if not.
	 */
	public void createRow(Object[] values, boolean header) {

		if (lastCellNum < values.length) {
			lastCellNum = values.length;
		}

		SXSSFRow row = sheet.createRow(rowNum++);
		for (int x = 0; x < values.length; x++) {
			Object o = values[x];
			Cell cell = row.createCell(x);
			if (o != null) {
				if (o.getClass().getName().contains("String")) {
					String value = (String) values[x];
					cell.setCellValue(value);
					saveColumnWidth(x, value);
					Integer width = 0;
					if (columnWidth.containsKey(x)) {
						width = columnWidth.get(x);
					}
					if (width < value.length()) {
						width = value.length();
					}

				} else if (o.getClass().getName().contains("Double")) {
					cell.setCellValue((Double) values[x]);
				} else if (o.getClass().getName().contains("Integer")) {
					cell.setCellValue((Integer) values[x]);
				} else if (o.getClass().getName().contains("Date")) {
					cell.setCellValue((Date) values[x]);
				}
				cell.setCellStyle(getCellStyle(values[x], header));
			}

			saveColumnWidth(x, o);

		}
	}

	/**
	 * Calculate width of column as save it.
	 * @param x - the colunm position. 
	 * @param value - the value of cell.
	 */
	private void saveColumnWidth(int x, Object value) {
		int w = 0;
		String className = value.getClass().getName();
		if (className.contains("String")) {
			w = ((String) value).length() * 256;
		}
		if (className.contains("Double") || className.contains("Integer")) {
			w = value.toString().length() * 256;
		}
		if (className.contains("Date")) {
			w = 2560;
		}
		if (w < MIN_WIDTH) {
			w = MIN_WIDTH;
		}
		if (w > MAX_WIDTH) {
			w = MAX_WIDTH;
		}
		if (columnWidth.containsKey(x)) {
			int i = columnWidth.get(x);
			if (i < w) {
				columnWidth.put(x, w);
			}
		} else {
			columnWidth.put(x, w);
		}

	}

	/**
	 * save the excel workbook into file.
	 * @param fileName -  filename for excel workbook.
	 * @return
	 * @throws IOException
	 */
	public boolean saveToFile(String fileName) throws IOException {
		flushRows();
		// Write the output to a file
		FileOutputStream fileOut = new FileOutputStream(fileName);
		workbook.write(fileOut);
		fileOut.close();
		// dispose of temporary files backing this workbook on disk
		workbook.dispose();
		return (true);
	}

	/**
	 * this method flushes all rows
	 */
	private void flushRows() throws IOException {
		if (!sheet.areAllRowsFlushed()) {
			sheet.flushRows();
		}
	}

}
