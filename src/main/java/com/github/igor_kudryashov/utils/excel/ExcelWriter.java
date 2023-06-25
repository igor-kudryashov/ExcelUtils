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

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

/**
 * The simple helper for create Microsoft Excel workbook via Apache POI
 *
 * @author igor.kudryashov@gmail.com
 * @version 2023-06-25
 */
public class ExcelWriter {

    private final SXSSFWorkbook workbook;

    // the styles of cells
    private final Map<String, XSSFCellStyle> styles = new HashMap<>();
    private final CreationHelper createHelper;
    // the storage for column widths
    private final Map<String, Map<Integer, Integer>> columnWidth = new HashMap<>();
    // the rows counter
    private final Map<String, Integer> rows = new HashMap<>();

    /**
     * Default constructor
     */
    public ExcelWriter() {
        // create new workbook with 100 unflushed records
        workbook = new SXSSFWorkbook(100);
        // When a new node is created via createRow() and the total number of unflushed
        // records would exceed the specified value, then the row with the lowest index
        // value is flushed and cannot be accessed via getRow() anymore.
        // A value of -1 indicates unlimited access. In this case all records that have
        // not been flushed by a call to flush() are available for random access.
        // A value of 0 is not allowed because it would flush any newly created row
        // without having a chance to specify any cells.

        createHelper = workbook.getCreationHelper();

    }

    /**
     * Creates a new sheet in the workbook
     *
     * @param name Name of sheet
     * @return Created sheet
     */
    public Sheet createSheet(String name) {
        // delete backslash in name
        String sheetName = name.replaceAll("\\\\", "-").trim();
        // create new worksheet
        Sheet sheet = workbook.getSheet(name);
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
     * @param sheet  Sheet
     * @param values the value of the new cell line
     * @param style  the style of cells
     * @return created row
     */

    @SuppressWarnings("unchecked")
    public Row createRow(Sheet sheet, Object[] values, XSSFCellStyle style) {
        Row row;
        String sheetName = sheet.getSheetName();
        int rownum = 0;
        if (rows.containsKey(sheetName)) {
            rownum = rows.get(sheetName);
        }
        // create new row
        row = sheet.createRow(rownum);
        // create a cells of row by array values
        for (int x = 0; x < values.length; x++) {
            Object o = values[x];
            Cell cell = row.createCell(x);
            if (o instanceof String) {
                String value = (String) o;
                cell.setCellValue(value);
                String lower = value.toLowerCase();
                if (lower.startsWith("http://") || lower.startsWith("https://") || lower.startsWith("notes://") || lower.startsWith("www.")) {
                    // hyperlink string
                    createHyperlink(sheet, rownum, x, value, style);
                } else {
                    // simple string
                    cell.setCellStyle(getDefaultCellStyle(o, rownum, style));
                }
            } else if (o instanceof Number) {
                cell.setCellValue((Double) o);
                cell.setCellStyle(getDefaultCellStyle(o, rownum, style));
            } else if (o.getClass().getName().contains("Date")) {
                cell.setCellValue((Date) o);
                cell.setCellStyle(getDefaultCellStyle(o, rownum, style));
            } else if (o instanceof Collection) {
                // join collection into multiline string
                String s;
                int maxLine = 0;
                ArrayList<Object> list = new ArrayList<>((Collection<Object>) o);
                if (list.size() > 1) {
                    StringBuilder sb = new StringBuilder();
                    sb.append(list.get(0).toString());
                    for (int i = 1; i < list.size(); i++) {
                        sb.append("\n");
                        String line = list.get(i).toString();
                        sb.append(line);
                        if (maxLine < line.length()) {
                            maxLine = line.length();
                        }
                    }
                    s = sb.toString();
                } else {
                    s = (String) list.get(0);
                    maxLine = s.length();
                }
                cell.setCellValue(s);
                cell.setCellStyle(getDefaultCellStyle(o, rownum, style));
                // save max column width
                if (rownum == 0) {
                    saveColumnWidth(sheet, x, new String(new char[maxLine]).replace('\0', ' '));
                }
            }
            // save max column width
            if (rownum != 0) {
                saveColumnWidth(sheet, x, o);
            }
        }
        // save the last number of row for this worksheet
        rows.put(sheetName, ++rownum);
        return row;

    }

    /**
     * Adds a hyperlink style to a cell. The content of the cell remains the same.
     * Don't forget to fill in the contents of the cell before adding a hyperlink style.
     * If the string is already cleared, this method does not work!
     *
     * @param sheet  Sheet
     * @param rownum number of row
     * @param column number of column
     * @param style  the style of cell
     * @param url    hyperlink
     */
    public void createHyperlink(Sheet sheet, int rownum, int column, String url, XSSFCellStyle style) {
        // get current Row
        Row row = sheet.getRow(rownum);
        if (url != null && !"".equals(url)) {
            // get Cell by column number
            Cell cell = row.getCell(column);
            // create hyperlink style
            XSSFHyperlink hyperlink = (XSSFHyperlink) createHelper.createHyperlink(HyperlinkType.URL);
            hyperlink.setAddress(url);
            cell.setHyperlink(hyperlink);
            XSSFCellStyle hyperlinkStyle = (XSSFCellStyle) workbook.createCellStyle();
            String name = "hyperlink";
            if (style != null) {
                hyperlinkStyle.cloneStyleFrom(style);
                name = name + "_" + style.getIndex();
            } else {
                hyperlinkStyle = (XSSFCellStyle) workbook.createCellStyle();
                name = name + "_" + hyperlinkStyle.getIndex();
            }
            if (styles.containsKey(name)) {
                hyperlinkStyle = styles.get(name);
            } else {
                // set font style
                XSSFFont font = (XSSFFont) workbook.createFont();
                if (font != null) {
                    font.setUnderline(XSSFFont.U_SINGLE);
                    font.setColor(IndexedColors.BLUE.getIndex());
                    hyperlinkStyle.setFont(font);
                }
                styles.put(name, hyperlinkStyle);
            }
            cell.setCellStyle(hyperlinkStyle);
        }
    }

    /**
     * Returns a style of cell
     *
     * @param value value of cell
     * @return the cell style
     */
    private XSSFCellStyle getDefaultCellStyle(Object value, int rownum, XSSFCellStyle s) {
        String styleName = (value == null) ? "null" : value.getClass().getName();
        if (rownum == 0) {
            styleName = styleName + ".header";
        }
        // create new style
        XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
        if (s != null) {
            style.cloneStyleFrom(s);
            styleName = styleName + "." + s.getIndex();
        }
        if (styles.containsKey(styleName)) {
            // if we already have a style for this class, return it
            style = styles.get(styleName);
        } else {
            // format data
            XSSFDataFormat fmt = (XSSFDataFormat) workbook.createDataFormat();
            short format = 0;
            if (styleName.contains("header")) {
                if (s == null) {
                    XSSFFont font = (XSSFFont) workbook.createFont();
                    font.setBold(true);
                    style.setFont(font);
                }
            } else if (styleName.contains("Date")) {
                format = fmt.getFormat(BuiltinFormats.getBuiltinFormat(0xe));
                style.setAlignment(HorizontalAlignment.LEFT);
            } else if (styleName.contains("Double")) {
                format = fmt.getFormat(BuiltinFormats.getBuiltinFormat(2));
                style.setAlignment(HorizontalAlignment.RIGHT);
            } else if (styleName.contains("Integer")) {
                format = fmt.getFormat(BuiltinFormats.getBuiltinFormat(1));
                style.setAlignment(HorizontalAlignment.RIGHT);
            } else {
                style.setAlignment(HorizontalAlignment.LEFT);
            }
            style.setDataFormat(format);
            // keep the style for reuse
            styles.put(styleName, style);
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
        if (value == null) {
            value = "";
        }
        if (columnWidth.containsKey(sheetName)) {
            width = columnWidth.get(sheetName);
        } else {
            width = new HashMap<>();
            columnWidth.put(sheetName, width);
        }
        // calculate width of column by data value
        int w = 0;
        String className = value.getClass().getName();
        if (className.contains("String")) {
            w = ((String) value).length() * 256;
        }
        if (className.contains("Double") || className.contains("Integer") || className.contains("Long")) {
            w = value.toString().length() * 256;
        }
        if (className.contains("Date")) {
            w = 2560;
        }
        int MIN_COLUMN_WIDTH = 3500;
        if (w < MIN_COLUMN_WIDTH) {
            w = MIN_COLUMN_WIDTH;
        }
        int MAX_COLUMN_WIDTH = 15000;
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
     * @param withHeader <code>true</code> for create auto filter and freeze pane in
     *                   first row, otherwise <code>false</code>
     */
    public void setAutoSizeColumns(Sheet sheet, boolean withHeader) {
        if (sheet.getLastRowNum() > 0) {
            if (withHeader) {
                // columns headers with drop-down button
                int x = sheet.getRow(sheet.getLastRowNum()).getLastCellNum();
                CellRangeAddress range = new CellRangeAddress(0, 0, 0, x - 1);
                sheet.setAutoFilter(range);
                sheet.createFreezePane(0, 1);
            }
            if (columnWidth.containsKey(sheet.getSheetName())) {
                // columns width
                Map<Integer, Integer> width = columnWidth.get(sheet.getSheetName());
                width.forEach(sheet::setColumnWidth);
            }
        }
    }

    /**
     * Save a workbook in file
     *
     * @param fileName filename
     */
    public void saveToFile(String fileName) throws IOException {
        for (int x = 0; x < workbook.getNumberOfSheets(); x++) {
            SXSSFSheet sheet = workbook.getSheetAt(x);
            sheet.flushRows();
        }
        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream(fileName);
        workbook.write(fileOut);
        fileOut.close();
        // dispose of temporary files backing this workbook on disk
        workbook.dispose();
    }

}