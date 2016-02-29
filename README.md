# utils.excel
The ExcelWriter is a simple helper for create Microsoft Excel workbook via Apache POI
#ussage
```java
package com.github.igor_kudryashov.utils.excel;

import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;

public class Example {
	public static void main(String[] args) {
		// create an Excel workbook
		ExcelWriter writer = new ExcelWriter();

		// create a first worksheet
		Sheet sheet1 = writer.createSheet("Sheet1");
		// create a second worksheet
		Sheet sheet2 = writer.createSheet("Sheet2");

		// create header of table for first worksheet
		writer.createRow(sheet1, new String[] { "Column 1", "Column 2", "Column 3" }, true);
		// create header of table for second worksheet
		writer.createRow(sheet2, new String[] { "Column 1", "Column 2", "Column 3" }, true);

		// work with first worksheet
		for (int x = 0; x < 3; x++) {
			// create simple row
			writer.createRow(sheet1, new Object[] { "Cell 1", "Cell 2", "Cell 3" }, false);
		}

		// work with second worksheet
		for (int x = 0; x < 3; x++) {
			// create simple row
			writer.createRow(sheet2, new Object[] { "Cell 1", "Cell 2", "Cell 3" }, false);
		}

		// format first worksheet
		writer.setAutoSizeColumns(sheet1, true);
		// format second worksheet
		writer.setAutoSizeColumns(sheet2, true);

		// save the workbook to file
		try {
			writer.saveToFile("file.xlsx");
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}

```
