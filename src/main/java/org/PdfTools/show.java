package org.PdfTools;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class show {
	
	public  static void show(File file) throws BiffException, IOException {
		InputStream inputStream = new FileInputStream(file.getAbsolutePath());
		Workbook workbook = Workbook.getWorkbook(inputStream);
		
		Sheet sheet = workbook.getSheet(0);
		int rows = sheet.getRows();
		
		 for (int i = 1; i < rows; i++) {
			 
			 Cell cell = sheet.getCell(8,i);
			 String s = cell.getContents();
			 System.out.println(s);
		 }
	}
	public static void main(String[] args) throws BiffException, IOException {
		File file = new File("F:\\20180926_F18FTSNWKF1790_I-WAbyS005_Mures_tang\\F18FTSNWKF1790_I-WAbyS005_20180926\\Submit\\BGI_result\\Significant\\1.xls");
		show(file);
	}
}
