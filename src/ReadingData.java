import java.util.Hashtable;

public class ReadingData {

	public static void main(String[] args) {
		// TODO Auto-generated method stub

		Xls_Reader xls = new Xls_Reader("D:\\Selenium\\Workspace\\ReadingData\\Data.xlsx");
		
		int colCount = xls.getColumnCount("Data");
		int rowCount = xls.getRowCount("Data");
		
		for (int rowNum = 1; rowNum <=rowCount ; rowNum++) {
			for (int colNum = 0; colNum < colCount; colNum++) {
				
				System.out.println(xls.getCellData("Data",colNum, rowNum));
			}
			
		}

		Hashtable<String, String>table=new Hashtable<>();
		table.put("username", "a1");
		table.put("password", "a2");
		
		
	}

}
