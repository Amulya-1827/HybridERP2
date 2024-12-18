package utilities;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.impl.xb.xsdschema.Public;

public class ExcelFileUtil {
XSSFWorkbook wb;
//constructor for reading excel path
public ExcelFileUtil(String Excelpath) throws Throwable
{
	FileInputStream fi = new FileInputStream(Excelpath);
	wb = new XSSFWorkbook(fi);
}
//method for counting no of rows in a sheet
public int rowCount(String sheetname) {
	return wb.getSheet(sheetname).getLastRowNum();
}
//method for reading cell data
public String getcellData(String sheetname,int row,int column)
{
	String data="";
	if(wb.getSheet(sheetname).getRow(row).getCell(column).getCellType()==CellType.NUMERIC)
	{
		int celldata = (int) wb.getSheet(sheetname).getRow(row).getCell(column).getNumericCellValue();
	    data = String.valueOf(celldata);
	}
	else
	{
		data = wb.getSheet(sheetname).getRow(row).getCell(column).getStringCellValue();
	}
	return data;
	}
//method for writing data
public void setcelldata(String sheetname,int row,int column,String status,String WriteExcel) throws Throwable
{
	//get sheet from wb
	XSSFSheet ws= wb.getSheet(sheetname);
	//get row in sheet 
	XSSFRow rownum = ws.getRow(row);
	//create cell
	XSSFCell cell = rownum.createCell(column);
	//write status
	cell.setCellValue(status);
	if(status.equalsIgnoreCase("PASS"))
	{
		XSSFCellStyle style	 = wb.createCellStyle();
		XSSFFont font = wb.createFont();
		font.setColor(IndexedColors.GREEN.getIndex());
		font.setBold(true);
		style.setFont(font);
		rownum.getCell(column).setCellStyle(style);
	}
	else if(status.equalsIgnoreCase("FAIL"))
	{
		XSSFCellStyle style = wb.createCellStyle();
		XSSFFont font = wb.createFont();
		font.setColor(IndexedColors.RED.getIndex());
		font.setBold(true);
		style.setFont(font);
		rownum.getCell(column).setCellStyle(style);
	}
	else if(status.equalsIgnoreCase("BLOCKED"))
	{
		XSSFCellStyle style = wb.createCellStyle();
		XSSFFont font = wb.createFont();
		font.setColor(IndexedColors.BLUE.getIndex());
		font.setBold(true);
		style.setFont(font);
		rownum.getCell(column).setCellStyle(style);
	}
    FileOutputStream fo = new FileOutputStream(WriteExcel);
    wb.write(fo);
}
    public static void main(String[] args) throws Throwable {
    	ExcelFileUtil xl = new ExcelFileUtil("D:\\Sample.xlsx");
    	int rc = xl.rowCount("Emp");
    	System.out.println(rc);
    	for(int i=1;i<=rc;i++)
    	{
    		String fname = xl.getcellData("Emp", i, 0);
    		String mname = xl.getcellData("Emp", i, 1);
    		String lname = xl.getcellData("Emp", i, 2);
    		String eid	= xl.getcellData("Emp", i, 3);
    		System.out.println(fname+ "     "+mname+"     "+lname+"    "+eid);
            //xl.setcelldata("Emp", i, 4, "Pass","d:\\Results.xlsx");
            //xl.setcelldata("Emp", i, 4, "Fail","d:\\Results.xlsx");
            xl.setcelldata("Emp", i, 4, "Blocked","d:\\Results.xlsx");
    	}
    	
    	
} 

}

//completed....executed.... saved the output in Results file in d-drive.
