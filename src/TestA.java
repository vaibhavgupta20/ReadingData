import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class TestA {

	@Test(dataProvider = "getData")
	public void TestA(String uName, String pWord) {

		
		
		
	}

	@DataProvider
	public Object[][] getData() {

		Xls_Reader xls = new Xls_Reader("D:\\Selenium\\Workspace\\ReadingData\\Data.xlsx");

		
		
		
		
		
		
		int colCount = xls.getColumnCount("Data");
		int rowCount = xls.getRowCount("Data");

		Object[][] data = new Object[3][2];

		int dataRow = 0;
		for (int rowNum = 3; rowNum <= rowCount; rowNum++) {
			for (int colNum = 0; colNum < colCount; colNum++) {

				data[dataRow][colNum] = xls.getCellData("Data", colNum, rowNum);
				// 0,0 0,1 1,0 ...
			}
			dataRow++;
		}
		
		return data;

	}

}
