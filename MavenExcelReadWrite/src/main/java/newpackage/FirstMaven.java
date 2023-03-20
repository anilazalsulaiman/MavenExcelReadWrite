package newpackage;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FirstMaven {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		File obj=new File("F:\\Anilaz Projects\\Software Testing\\Java\\Programs\\ExcelReadWrite\\mavenExcel.xlsx");
		FileInputStream obj1=new FileInputStream(obj);
		XSSFWorkbook wb=new XSSFWorkbook(obj1);
		XSSFSheet sheet=wb.getSheetAt(0);
		
		for(Row r:sheet) {
			for(Cell c:r) {
				System.out.println(c);
			}
		}
	}

}
