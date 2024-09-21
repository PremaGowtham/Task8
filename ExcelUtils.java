package Task8;

import java.io.IOException;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.
CTSheet;

public class ExcelUtils {
	
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;
	private static String projectPath;

	public static void main(String[] args) {
		getRowCount();
		getCellDataString(0, 0);
		getCellDataNumeric(1,1);
	}
	
	public static void getRowCount(){
		
		try {
			String projectPath=System.getProperty("user.dir");
			workbook = new XSSFWorkbook(projectPath+ "\\excel\\sheet1.xlsx");
			sheet = workbook.getSheet("sheet1");
			int rowCount = sheet.getPhysicalNumberOfRows();
			System.out.println("Numer of Rows" +rowCount);
		}catch(IOException e) {
			e.printStackTrace();
		}
	}
	
	public static void getCellDataString(int rowNum, int ColNum) {
		try {
			String projectPath=System.getProperty("user.dir");
			workbook = new XSSFWorkbook(projectPath+ "\\excel\\data.xlsx");
			sheet = workbook.getSheet("sheet1");
			String cellData = sheet.getRow(rowNum).getCell(ColNum).getStringCellValue();
			System.out.println(cellData);
		}catch(Exception e) {
			e.printStackTrace();
		}
	}
	
	public static void getCellDataNumeric(int rowNum, int ColNum) {
		try {
			
			String projectPathSystem = System.getProperty("user.dir");
			workbook = new XSSFWorkbook(projectPath+ "\\excel\\data.xlsx");
			sheet = workbook.getSheet("sheet1");
			int cellData = (int) sheet.getRow(rowNum).getCell(ColNum).getNumericCellValue();
			System.out.println(cellData);
		}catch(Exception e) {
			e.printStackTrace();
		}
	
			
	}
}
