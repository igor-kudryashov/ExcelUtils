# ExcelWriter
The ExcelWriter is a simple helper for create Microsoft Excel workbook via Apache POI. 
ExcelWriter using the streaming usermodel API https://poi.apache.org/components/spreadsheet/how-to.html#sxssf and can generate the very large spreadsheets. 
SXSSF achieves its low memory footprint by limiting access to the rows that are within a sliding window. 
ExcelWriter requires Java 1.8 or higher and Apache POI 4.1.1 for create a Microsoft Excel workbook. To create the worksheet content in the Exmple class, I am using the "Java Fake" project from https://github.com/DiUS/java-faker.

## Ussage
```java
import com.github.javafaker.Faker;
import com.github.javafaker.service.FakeValuesService;
import com.github.javafaker.service.RandomService;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import java.io.IOException;
import java.util.Locale;

/**
 * An example of using the ExcelWriter class
 *
 * @author igor.kudryashov@gmail.com
 * @version 2023-06-25
 */
public class Example {
    public static void main(String[] args) {
        // create an Excel workbook
        ExcelWriter writer = new ExcelWriter();

        // create a first worksheet
        Sheet sheet1 = writer.createSheet("Sheet1");
        // create a second worksheet
        Sheet sheet2 = writer.createSheet("Sheet2");

        // create style for header row
        XSSFCellStyle headerStyle = (XSSFCellStyle) writer.getWorkbook().createCellStyle();
        XSSFFont headerFont = (XSSFFont) sheet2.getWorkbook().createFont();
        headerStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setFont(headerFont);

        // custom color example
        // byte[] rgb = {(byte)153, (byte)204, (byte) 255};
        // headerStyle.setFillForegroundColor(new XSSFColor(rgb, new DefaultIndexedColorMap()));
        // headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // create table header with default style
        writer.createRow(sheet1, new String[]{"Name", "Email", "Birthday", "Sum", "Site"}, null);

        // create table header with custom header row style
        writer.createRow(sheet2, new String[]{"Name", "Email", "Birthday", "Sum", "Site"}, headerStyle);

        // create custom style for alternate row
        XSSFCellStyle styleEven = (XSSFCellStyle) writer.getWorkbook().createCellStyle();
        styleEven.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        styleEven.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // to create worksheet content I use the "Java Faker" project https://github.com/DiUS/java-faker
        FakeValuesService fakeValuesService = new FakeValuesService(new Locale("en-GB"), new RandomService());

        // create content for first worksheet with row with default style
        for (int x = 0; x < 100; x++) {
            Faker faker = new Faker();
            // create simple row with style
            writer.createRow(sheet1, new Object[]{faker.name().fullName(), faker.internet().emailAddress(),
                    faker.date().birthday(), Double.valueOf(faker.commerce().price().replace(",", ".")),
                    faker.internet().url()}, null);
        }

        // create content for second worksheet with row with custom style
        for (int x = 0; x < 100; x++) {
            Faker faker = new Faker();
            if (x % 2 == 0) {
                writer.createRow(sheet2, new Object[]{faker.name().fullName(), faker.internet().emailAddress(),
                        faker.date().birthday(), Double.valueOf(faker.commerce().price().replace(",", ".")),
                        faker.internet().url()}, styleEven);
            } else {
                writer.createRow(sheet2, new Object[]{faker.name().fullName(), faker.internet().emailAddress(),
                        faker.date().birthday(), Double.valueOf(faker.commerce().price().replace(",", ".")),
                        faker.internet().url()}, null);
            }
        }

        // format worksheets
        writer.setAutoSizeColumns(sheet1, true);
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
