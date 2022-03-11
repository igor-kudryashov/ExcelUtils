# ExcelUtils
The ExcelWriter is a simple helper for create Microsoft Excel workbook via Apache POI
#ussage
```java
public class Example {
	public static void main(String[] args) {
		// create an Excel workbook
		ExcelWriter writer = new ExcelWriter();

		// create a first worksheet
		Sheet sheet1 = writer.createSheet("Sheet1");
		// create a second worksheet
		Sheet sheet2 = writer.createSheet("Sheet2");

		// create style
		XSSFCellStyle style1 = (XSSFCellStyle) writer.getWorkbook().createCellStyle();
		style1.setFillForegroundColor(IndexedColors.LIME.getIndex());
		style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		// create header of table for first worksheet with style
		writer.createRow(sheet1, new String[] { "Column 1", "Column 2", "Column 3" }, style1);
		// create header of table for second worksheet without style
		writer.createRow(sheet2, new String[] { "Column 1", "Column 2", "Column 3" }, null);

		// create style
		XSSFCellStyle style2 = (XSSFCellStyle) writer.getWorkbook().createCellStyle();
		style2.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
		style2.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		// work with first worksheet
		for (int x = 0; x < 3; x++) {
			// create simple row with style
			Row row = writer.createRow(sheet1, new Object[] { "Cell 1", "Cell 2", "Cell 3" }, style2);
			// append hyperlink
			writer.createHyperlink(sheet1, row.getRowNum(), 2, "http://www.microsoft.com", style2);
		}

		// work with second worksheet
		for (int x = 0; x < 3; x++) {
			// create simple row without style
			Row row = writer.createRow(sheet2, new Object[] { "Cell 1", "Cell 2", "Cell 3" }, null);
			// append hyperlink
			writer.createHyperlink(sheet2, row.getRowNum(), 2, "http://www.ibm.com", null);
		}

		// format first worksheet
		writer.setAutoSizeColumns(sheet1, true);
		// format second worksheet
		writer.setAutoSizeColumns(sheet2, false);

		// save the workbook to file
		try {
			writer.saveToFile("file.xlsx");
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
```
