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

/**
 * The simple helper for create Microsoft Excel workbook via Apache POI
 */
public class ExcelWriter {

    private SXSSFWorkbook workbook;
    private final XSSFColor colorBorder = new XSSFColor(new java.awt.Color(192, 192, 192));
    private final XSSFColor colorHeaderBackground = new XSSFColor(new java.awt.Color(210, 210, 210));
    private final XSSFColor colorEvenCellBackground = new XSSFColor(new java.awt.Color(239, 239, 239));
    private final int MIN_COLUMN_WIDTH = 3000;
    private final int MAX_COLUMN_WIDTH = 15000;

    // the styles of cells
    Map<String, XSSFCellStyle> styles = new HashMap<String, XSSFCellStyle>();
    // the storage for column widths
    Map<String, Map<Integer, Integer>> columnWidth = new HashMap<String, Map<Integer, Integer>>();
    // the rows counter
    Map<String, Integer> rows = new HashMap<String, Integer>();

    /**
     * Default constructor
     */
    public ExcelWriter() {
        //create new workbook
        workbook = new SXSSFWorkbook();
    }

    /**
     * Creates a new sheet in the workbook
     *
     * @param name Name of sheet
     * @return Created sheet
     */
    public Sheet createSheet(String name) {
        // delete back slash in name
        String sheetName = name.replaceAll("\\\\", "-").trim();
        // create new worksheet
        SXSSFSheet sheet = workbook.getSheet(name);
        if (sheet == null) {
            sheet = workbook.createSheet(sheetName);
        }
        return sheet;
    }

    /**
     * Returns active workbook
     *
     * @return active workbook
     */
    public Workbook getWorkbook() {
        return workbook;
    }

    /**
     * Creates new row in the worksheet
     *
     * @param sheet  Name of Sheet
     * @param values the value of the new cell line
     * @param header <code>true</code> if this row is the header, otherwise <code>false</code>
     * @return created row
     */
    public Row createRow(Sheet sheet, Object[] values, boolean header) {
        Row row;
        String sheetName = sheet.getSheetName();
        int rownum = 0;
        if (rows.containsKey(sheetName)) {
            rownum = rows.get(sheetName);
        }
        // create new row
        row = sheet.createRow(rownum);
        // create a cells of row
        for (int x = 0; x < values.length; x++) {
            Object o = values[x];
            Cell cell = row.createCell(x);
            if (o != null) {
                if (o.getClass().getName().contains("String")) {
                    String value = (String) values[x];
                    cell.setCellValue(value);
                } else if (o.getClass().getName().contains("Double")) {
                    cell.setCellValue((Double) values[x]);
                } else if (o.getClass().getName().contains("Integer")) {
                    cell.setCellValue((Integer) values[x]);
                } else if (o.getClass().getName().contains("Date")) {
                    cell.setCellValue((Date) values[x]);
                }
                cell.setCellStyle(getCellStyle(rownum, values[x], header));
            }
            // save max column width
            saveColumnWidth(sheet, x, o);
        }
        // save the last number of row for this worksheet
        rows.put(sheetName, ++rownum);
        return row;
    }

    /**
     * Returns a style of cell
     *
     * @param rownum the number of row for count odd/even rows
     * @param entry  value of cell
     * @param header <code>true</code> if this row is the header, otherwise <code>false</code>
     * @return the cell style
     */
    private XSSFCellStyle getCellStyle(int rownum, Object entry, boolean header) {
        XSSFCellStyle style;
        String name = entry.getClass().getName();
        if (header) {
            name += "_header";
        } else if ((rownum % 2) == 0) {
            name += "_even";
        }
        if (styles.containsKey(name)) {
            // if we already have a style for this class, return it
            style = styles.get(name);
        } else {
            // create new style
            style = (XSSFCellStyle) workbook.createCellStyle();
            style.setVerticalAlignment(VerticalAlignment.TOP);
            style.setBorderBottom(CellStyle.BORDER_THIN);
            style.setBorderColor(XSSFCellBorder.BorderSide.BOTTOM, colorBorder);
            style.setBorderLeft(CellStyle.BORDER_THIN);
            style.setBorderColor(XSSFCellBorder.BorderSide.LEFT, colorBorder);
            style.setBorderRight(CellStyle.BORDER_THIN);
            style.setBorderColor(XSSFCellBorder.BorderSide.RIGHT, colorBorder);
            // format data
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
                // for header
                style.setFillForegroundColor(colorHeaderBackground);
                style.setFillPattern(CellStyle.SOLID_FOREGROUND);
            } else if (name.contains("_even")) {
                // for even rows
                style.setFillForegroundColor(colorEvenCellBackground);
                style.setFillPattern(CellStyle.SOLID_FOREGROUND);
            }
            style.setDataFormat(format);
            // keep the style for reuse
            styles.put(name, style);
        }
        return style;
    }

    /**
     * Stores the maximum width of the column
     *
     * @param sheet Name of worksheet
     * @param x     number of column
     * @param value cell value
     */
    private void saveColumnWidth(Sheet sheet, int x, Object value) {
        String sheetName = sheet.getSheetName();
        Map<Integer, Integer> width;

        if (columnWidth.containsKey(sheetName)) {
            width = columnWidth.get(sheetName);
        } else {
            width = new HashMap<Integer, Integer>();
            columnWidth.put(sheetName, width);
        }
        // calculate width of column by data value
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
        if (w < MIN_COLUMN_WIDTH) {
            w = MIN_COLUMN_WIDTH;
        }
        if (w > MAX_COLUMN_WIDTH) {
            w = MAX_COLUMN_WIDTH;
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

    /**
     * Format a table of worksheet
     *
     * @param sheet      Name of sheet
     * @param withHeader <code>true</code> for create auto filter and freeze pane in first row, otherwise <code>false</code>
     */
    public void setAutoSizeColumns(Sheet sheet, boolean withHeader) {
        if (withHeader) {
            int x = sheet.getRow(1).getLastCellNum();
            CellRangeAddress range = new CellRangeAddress(0, 0, 0, x - 1);
            sheet.setAutoFilter(range);
            sheet.createFreezePane(0, 1);
        }
        // auto-sizing columns
        if (columnWidth.containsKey(sheet.getSheetName())) {
            Map<Integer, Integer> width = columnWidth.get(sheet.getSheetName());
            for (Map.Entry<Integer, Integer> entry : width.entrySet()) {
                sheet.setColumnWidth(entry.getKey(), entry.getValue());
            }
        }
    }

    /**
     * Save a workbook in file
     *
     * @param fileName filename
     * @return <code>true</code> if saved successfully, otherwise <code>false</code>
     * @throws IOException
     */
    public boolean saveToFile(String fileName) throws IOException {
        for (Sheet aWorkbook : workbook) {
            SXSSFSheet sheet = (SXSSFSheet) aWorkbook;
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
