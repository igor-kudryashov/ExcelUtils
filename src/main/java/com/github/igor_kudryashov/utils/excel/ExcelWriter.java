package com.github.igor_kudryashov.utils.excel;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;

public class ExcelWriter {

	private SXSSFWorkbook workbook;

	private final XSSFColor colorBorder = new XSSFColor(new java.awt.Color(192, 192, 192));
	private final XSSFColor colorHeaderBackground = new XSSFColor(new java.awt.Color(210, 210, 210));
	private final XSSFColor colorEvenCellBackground = new XSSFColor(new java.awt.Color(239, 239, 239));
	// min colunm width (pic)
	private final int MIN_WIDTH = 3000;
	// max column width (pic)
	private final int MAX_WIDTH = 15000;

	// the styles of cells
	Map<String, XSSFCellStyle> styles = new HashMap<String, XSSFCellStyle>();
	Map<String, Map<Integer, Integer>> columnWidth = new HashMap<String, Map<Integer, Integer>>();
	// the rows counter
	Map<String, Integer> rows = new HashMap<String, Integer>();

	public ExcelWriter() {
		workbook = new SXSSFWorkbook();
	}

	public Sheet createSheet(String name) {
		String sheetName = name.replaceAll("\\\\", "-").trim();
		SXSSFSheet sheet = workbook.getSheet(name);
		if (sheet == null) {
			sheet = workbook.createSheet(sheetName);
		}
		return sheet;
	}

	public Workbook getWorkbook() {
		return workbook;
	}

	public Row createRow(Sheet sheet, Object[] values, boolean header) {
		Row row;
		String sheetName = sheet.getSheetName();
		int rownum;
		if (rows.containsKey(sheetName)) {
			rownum = rows.get(sheetName);
		} else {
			rownum = 0;
		}
		row = sheet.createRow(rownum);
		for (int x = 0; x < values.length; x++) {
			Object o = values[x];
			Cell cell = row.createCell(x);
			if (o != null) {
				if (o.getClass().getName().contains("String")) {
					String value = (String) values[x];
					cell.setCellValue(value);
					saveColumnWidth(sheet, x, value);
				} else if (o.getClass().getName().contains("Double")) {
					cell.setCellValue((Double) values[x]);
				} else if (o.getClass().getName().contains("Integer")) {
					cell.setCellValue((Integer) values[x]);
				} else if (o.getClass().getName().contains("Date")) {
					cell.setCellValue((Date) values[x]);
				}
				cell.setCellStyle(getCellStyle(rownum, values[x], header));
			}
			saveColumnWidth(sheet, x, o);
		}
		rows.put(sheetName, ++rownum);
		return row;
	}

	private XSSFCellStyle getCellStyle(int rownum, Object entry, boolean header) {
		XSSFCellStyle style = null;
		String name = entry.getClass().getName();
		if (header) {
			name += "_header";
		} else if ((rownum % 2) == 0) {
			name += "_even";
		}
		if (styles.containsKey(name)) {
			style = styles.get(name);
		} else {
			style = (XSSFCellStyle) workbook.createCellStyle();

			style.setVerticalAlignment(VerticalAlignment.TOP);
			style.setBorderBottom(CellStyle.BORDER_THIN);
			style.setBorderColor(XSSFCellBorder.BorderSide.BOTTOM, colorBorder);
			style.setBorderLeft(CellStyle.BORDER_THIN);
			style.setBorderColor(XSSFCellBorder.BorderSide.LEFT, colorBorder);
			style.setBorderRight(CellStyle.BORDER_THIN);
			style.setBorderColor(XSSFCellBorder.BorderSide.RIGHT, colorBorder);

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
				style.setFillForegroundColor(colorHeaderBackground);
				style.setFillPattern(CellStyle.SOLID_FOREGROUND);
			} else if (name.contains("_even")) {
				style.setFillForegroundColor(colorEvenCellBackground);
				style.setFillPattern(CellStyle.SOLID_FOREGROUND);
			}
			style.setDataFormat(format);
			styles.put(name, style);
		}
		return style;
	}

	private void saveColumnWidth(Sheet sheet, int x, Object value) {
		String sheetName = sheet.getSheetName();
		Map<Integer, Integer> width;

		if (columnWidth.containsKey(sheetName)) {
			width = columnWidth.get(sheetName);
		} else {
			width = new HashMap<Integer, Integer>();
			columnWidth.put(sheetName, width);
		}
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
		if (width.containsKey(x)) {
			int i = width.get(x);
			if (i < w) {
				width.put(x, w);
			}
		} else {
			width.put(x, w);
		}

	}

	public void setAutoSizeColumns(Sheet sheet, boolean withHeader) {

		if (withHeader) {
			int x = sheet.getRow(1).getLastCellNum();
			CellRangeAddress range = new CellRangeAddress(0, 0, 0, x - 1);
			sheet.setAutoFilter(range);
			sheet.createFreezePane(0, 1);
		}
	}

	public boolean saveToFile(String fileName) throws IOException {
		Iterator<Sheet> it = workbook.iterator();
		while (it.hasNext()) {
			SXSSFSheet sheet = (SXSSFSheet) it.next();
			if (!sheet.areAllRowsFlushed()) {
				sheet.flushRows();
			}
		}
		// Write the output to a file
		FileOutputStream fileOut = new FileOutputStream(fileName);
		workbook.write(fileOut);
		fileOut.close();
		// dispose of temporary files backing this workbook on disk
		workbook.dispose();
		return (true);
	}

}
