import java.util.Hashtable;

import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class TestB {
	@Test(dataProvider = "getData")
	public void TestA(Hashtable<String, String> data) {

		System.out.println(data.get("username"));
		System.out.println(data.get("password"));
		System.out.println(data.get("Col3"));
		System.out.println(data.get("Col4"));

	}

	@DataProvider
	public Object[][] getData() {

		Xls_Reader xls = new Xls_Reader("D:\\Selenium\\Workspace\\ReadingData\\Data.xlsx");
		String sheetName = "Data";

		int testStartRow = 1;
		while (!xls.getCellData(sheetName, 0, testStartRow).equals("TestB")) {
			testStartRow++;
		}

		System.out.println("test row starts: " + testStartRow);

		int colStartRowNum = testStartRow + 1;
		int dataStartRowNum = testStartRow + 2;

		// calculate number colnumber

		int cols = 0;
		while (!xls.getCellData(sheetName, cols, colStartRowNum).equals("")) {
			cols++;
		}
		System.out.println("total number of col: " + cols);

		int rows = 0;
		while (!xls.getCellData(sheetName, 0, dataStartRowNum + rows).equals("")) {
			rows++;
		}

		System.out.println("total number of rows: " + rows);

		Object[][] data = new Object[rows][1];

		int dataRow = 0;
		Hashtable<String, String> table;

		for (int rowNum = dataStartRowNum; rowNum < dataStartRowNum + rows; rowNum++) {

			table = new Hashtable<String, String>();

			for (int colNum = 0; colNum < cols; colNum++) {

				String key = xls.getCellData(sheetName, colNum, colStartRowNum);
				String value = xls.getCellData(sheetName, colNum, rowNum);

				table.put(key, value);

//				data[dataRow][colNum] = xls.getCellData("Data", colNum, rowNum);
				// 0,0 0,1 1,0 ...
			}
			data[dataRow][0] = table;

			dataRow++;
		}

		return data;

	}
}
