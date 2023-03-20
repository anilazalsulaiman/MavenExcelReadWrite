package newpackage;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteDataClass {
	static String name="Anilaz value changed";
	
public static void main(String[] args) throws IOException {
	XSSFWorkbook wb=new XSSFWorkbook();
	XSSFSheet sheet=wb.createSheet("FirstSheet");
	
	XSSFRow rw=sheet.createRow(0);
	Cell cl=rw.createCell(0);
	cl.setCellValue(name);
	File objfile=new File("F:\\Anilaz Projects\\Software Testing\\Java\\Programs\\ExcelReadWrite\\mavenExcel1.xlsx");
	FileOutputStream fl=new FileOutputStream(objfile);
	wb.write(fl);
	fl.close();
	System.out.println("Succesfully written");
}
}
