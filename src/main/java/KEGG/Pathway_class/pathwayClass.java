package KEGG.Pathway_class;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class pathwayClass{
	public static void getinfo(File file1,File file2) throws Exception{
		
		InputStream inputStream = new FileInputStream(file2.getAbsolutePath());
		Workbook workbook = Workbook.getWorkbook(inputStream);
		
		Sheet sheet = workbook.getSheet(0);
		int rows = sheet.getRows();
		
		InputStream inputStream2 = new FileInputStream(file1.getAbsolutePath());
		Workbook workbook2 = Workbook.getWorkbook(inputStream2);
		
		Sheet sheet2 = workbook2.getSheet(0);
		int rows2 = sheet2.getRows();
		
		List<String>list1 = new ArrayList<String>();
		List<String>list2 = new ArrayList<String>();
		List<String>list3 = new ArrayList<String>();
		
		for (int i = 0; i < rows; i++) {
			//我从这里拿到的s1就是蛋白ID  通过蛋白ID输出一些东西
			Cell cell1 = sheet.getCell(1,i);
			String s1 = cell1.getContents();
			
			for (int j = 0; j < rows2; j++) {
				//在另一个表中的第三列
				Cell cell2 = sheet2.getCell(3,j);
				String s2 = cell2.getContents();
				//Pathway term
				Cell cell3 = sheet2.getCell(1,j);
				String s3 = cell3.getContents();
				//Ontoloty
				Cell cell4 = sheet2.getCell(0,j);
				String s4 = cell4.getContents();
				if (s2.contains(s1)) {
	//				System.out.println(s1+" "+s3+" "+s4);
					list1.add(s1);
					list2.add(s3);
					list3.add(s4);
				}
				
			}
		}
		 WritableWorkbook book = Workbook.createWorkbook(new File("F://pathwayClass.xls"));
		 
		 WritableSheet sheet3 = book.createSheet("test", 0);
		 
		 //插入蛋白行
		 Label label1 = new Label(0,0,"gene_id");
		 //插入综合注释行
		 Label label2 = new Label(1,0,"kegg_level2");
		 //插入go_level1_mix
		 Label label3 = new Label(2,0,"kegg_level1");
		 
		 sheet3.addCell(label1);
		 sheet3.addCell(label2);
		 sheet3.addCell(label3);
		
		
		for (int i = 0; i < list1.size(); i++) {
			String s1 = list1.get(i);
			Label label4 = new Label(0,i+1,s1);
			sheet3.addCell(label4);
			
			String s2 = list2.get(i);
			Label label5 = new Label(1,i+1,s2);
			sheet3.addCell(label5);
			
			String s3 = list3.get(i);
			Label label6 = new Label(2,i+1,s3);
			sheet3.addCell(label6);
		}
		
		book.write();
		book.close();
	}
	
	public static void main(String[] args) {
		
		File file1 = new File("F:\\20180926_F18FTSNWKF1790_I-WAbyS005_Mures_tang\\F18FTSNWKF1790_I-WAbyS005_20180926\\Submit\\BGI_result\\Function_Analyse\\DiffExpressedProtein\\Pathway_Enrichment\\A-VS-B_Pathway_enrichment\\3.xls");
		
		File file2 = new File("F:\\20180926_F18FTSNWKF1790_I-WAbyS005_Mures_tang\\F18FTSNWKF1790_I-WAbyS005_20180926\\Submit\\BGI_result\\Significant\\1.xls");
		
		try {
			getinfo(file1, file2);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
