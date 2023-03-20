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

public class NewExelmultipleData {
	public static void main(String[] args) throws IOException {
		//Create blank workbook
	      XSSFWorkbook workbook = new XSSFWorkbook();
	      
	      //Create a blank sheet
	      XSSFSheet spreadsheet = workbook.createSheet( " Employee Info ");

	      //Create row object
	      XSSFRow row;

	      //This data needs to be written (Object[])
	      Map < String, String > empinfo = new TreeMap < String, String >();
	      empinfo.put( "1", "Name1");
	      
	      empinfo.put( "2", "Name2");
	      
	      empinfo.put( "3", "Name3");
	      
	      empinfo.put( "4", "Name4");
	      
	      empinfo.put( "5", "Name5");
	      
	      empinfo.put( "6", "Name6");

	      //Iterate over data and write to sheet
	      Set < String > keyid = empinfo.keySet();
	      int rowid = 0;
	      
	      for (String key : keyid) {
	         row = spreadsheet.createRow(rowid++);
	         int cellid = 0;
	         
	         
	            Cell cell = row.createCell(cellid++);
	            cell.setCellValue((String)key);
	            Cell cell1 = row.createCell(cellid++);
	            cell1.setCellValue((String)empinfo.get(key));
	         
	      }
	      //Write the workbook in file system
	      FileOutputStream out = new FileOutputStream(new File("F:\\\\Anilaz Projects\\\\Software Testing\\\\Java\\\\Programs\\\\ExcelReadWrite\\\\WriteMultipleData.xlsx"));
	      
	      workbook.write(out);
	      out.close();
	      System.out.println("multiple datas written successfully");
	}

}
