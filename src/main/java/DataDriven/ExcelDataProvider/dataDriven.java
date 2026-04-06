package DataDriven.ExcelDataProvider;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class dataDriven {

	@Test(dataProvider = "ExcelData")
	public void testCaseData(String TestCases, String Data1, String Data2, String Data3) {

		System.out.println(TestCases + Data1 + Data2 + Data3);

	}

	@DataProvider(name = "ExcelData")
	public Object[][] getData() throws IOException {

		DataFormatter formatter = new DataFormatter();
		String excelFilePath = System.getProperty("user.dir") + "//SeleniumDataDrivenWorkbook.xlsx";
		FileInputStream fis = new FileInputStream(excelFilePath);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheetAt(0);
		int rowCount = sheet.getPhysicalNumberOfRows();
		XSSFRow row = sheet.getRow(0);
		int colCount = row.getLastCellNum();
		Object data[][] = new Object[rowCount - 1][colCount];
		for (int i = 0; i < rowCount - 1; i++) {

			row = sheet.getRow(i + 1);
			for (int j = 0; j < colCount; j++) {

				XSSFCell cell = row.getCell(j);
				data[i][j] = formatter.formatCellValue(cell);

			}

		}
		return data;

	}

}
