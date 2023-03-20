package multipleWriteExcel;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelMultipleWrite {
	
	public static void main(String[] args) throws IOException {
		  //Create blank workbook
	      XSSFWorkbook workbook = new XSSFWorkbook();
	      
	      //Create a blank sheet
	      XSSFSheet spreadsheet = workbook.createSheet( " Employee Info ");

	      //Create row object
	      XSSFRow row;

	      //This data needs to be written (Object[])
	      Map < Integer, Object[] > empinfo = new TreeMap < Integer, Object[] >();
	      empinfo.put( 1, new Object[] {
	         "EMP ID", "EMP NAME", "DESIGNATION" });
	      
	      empinfo.put( 2, new Object[] {
	         "tp01", "Gopal", "Technical Manager" });
	      
	      empinfo.put( 3, new Object[] {
	         "tp02", "Manisha", "Proof Reader" });
	      
	      empinfo.put( 4, new Object[] {
	         "tp03", "Masthan", "Technical Writer" });
	      
	      empinfo.put( 10, new Object[] {
	         "tp04", "Satish", "Technical Writer" });
	      
	      empinfo.put( 6, new Object[] {
	         "tp05", "Krishna", "Technical Writer" });

	      //Iterate over data and write to sheet
	      Set < Integer > keyid = empinfo.keySet();
	      System.out.println(keyid);
	      int rowid = 0;
	      
	      for (Integer key : keyid) {
	         row = spreadsheet.createRow(rowid++);
	         Object [] objectArr = empinfo.get(key);
	         int cellid = 0;
	         
	         for (Object obj : objectArr){
	            Cell cell = row.createCell(cellid++);
	            cell.setCellValue((String)obj);
	         }
	      }
	      //Write the workbook in file system
	      FileOutputStream out = new FileOutputStream(new File("F:\\\\Anilaz Projects\\\\Software Testing\\\\Java\\\\Programs\\\\ExcelReadWrite\\\\WriteMultipleData.xlsx"));
	      
	      workbook.write(out);
	      out.close();
	      System.out.println("multiple datas written successfully");
		
	}
}
