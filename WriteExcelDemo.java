package Task8;

import java.io.FileOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcelDemo {

	public static void main(String[] args) throws IOException{
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("EmployeeData");
		
		Object empData[][] = {{"EmpID", "Name", "Designation"}, {101, "Krish", "Software Engineer"},{102, "Kavin", "Software Tester"}, {103, "Gowtham", "HR"}
	};
		
		int rows = empData.length;
		int cols = empData[0].length;
		
		System.out.println("Number of Rows" +rows);
		System.out.println("Number of Column" +cols);
		
		for(int r = 0; r < rows; r++) {

        XSSFRow row = sheet.createRow(r);
        
        for(int c = 0; c < cols; c++) {
        	XSSFCell cell = row.createCell(c);
        
        	Object value = empData[r][c];
        	if(value instanceof String)
        		cell.setCellValue((String) value);
        	
        	if(value instanceof Integer)
        		cell.setCellValue((Integer) value);
        	
        	if(value instanceof Boolean)
        		cell.setCellValue((Boolean) value);
        }
        
			}
		String filePath = ".\\excel\\Employees.xlsx";
		FileOutputStream fos = new FileOutputStream(filePath);
		workbook.write(fos);
		fos.close();
		System.out.println("Employee file written successfully");

	}
}
	