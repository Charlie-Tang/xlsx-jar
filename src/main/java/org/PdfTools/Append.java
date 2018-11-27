package org.PdfTools;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;

public class Append {
	//我是直接用蛋白ID从我们的文件中查找信息，并添加  还是通过GO注释呢
	
	public static List<String> get(File file1,File file2)throws Exception
	{
		InputStream inputStream = new FileInputStream(file1.getAbsolutePath());
		Workbook workbook = Workbook.getWorkbook(inputStream);
		Sheet sheet = workbook.getSheet(0);
		int rows = sheet.getRows();
		
		InputStream inputStream1 = new FileInputStream(file2.getAbsolutePath());
		Workbook workbook1 = Workbook.getWorkbook(inputStream1);
		Sheet sheet1 = workbook1.getSheet(0);
		int rows1 = sheet1.getRows();
		
		List<String> list = new ArrayList<String>();
		
		StringBuilder sb = new StringBuilder();
		
		 for (int i = 1; i < rows; i++) {
			 Cell cell = sheet.getCell(1,i);
			 //获取我们的蛋白字段名字
			 String id = cell.getContents();
			 //作为我的标识位，将id先加入看看
			 
			 for (int j = 1; j < rows1; j++) {
				 //蛋白的列
				Cell cell2 = sheet1.getCell(3,j);
				Cell cell3 = sheet1.getCell(1,j);
				Cell cell4 = sheet1.getCell(0,j);
				String totalid = cell2.getContents();
				
				if (totalid.contains(id)) {
					
					String biological_process = cell4.getContents();
					String GO_term = cell3.getContents();
					String mix_GO_term = null;
//					System.out.println(id);
					
					if (sb.toString().contains("[P]")&&biological_process.equals("biological_process")) {
						mix_GO_term = GO_term+";";
					}
					else if (sb.toString().contains("[C]")&&biological_process.equals("cellular_component")) {
						mix_GO_term = GO_term +";";
					}
					else if(sb.toString().contains("[F]")&&biological_process.equals("molecular_function")) 
					{
						mix_GO_term = GO_term + ";";
					}
					else {
						if (biological_process.equals("biological_process")) {
							mix_GO_term = "[P]"+GO_term+";";
						}
						else if (biological_process.equals("cellular_component")) {
							mix_GO_term = "[C]"+GO_term+";";
						}
						else if (biological_process.equals("molecular_function")) {
							mix_GO_term = "[F]"+GO_term+";";
						}
					}
					sb.append(mix_GO_term);
				}
			}	
			 
			 
			 	list.add(sb.toString());
			 	sb.setLength(0);
		 }
		 
		 List<String> finalPCF = new ArrayList<String>();
		 
		 for (int i = 0; i < list.size(); i++) {
			 String PCF = list.get(i);
			 int index = PCF.lastIndexOf(";");
			 finalPCF.add(PCF.substring(0, index));
		}
		 
		 return finalPCF;
		 
	}
	
	
	//可以直接把参数暴露出来 好操作
	public static void main(String[] args) {
		File file1 = new File("F:\\20180926_F18FTSNWKF1790_I-WAbyS005_Mures_tang\\F18FTSNWKF1790_I-WAbyS005_20180926\\Submit\\BGI_result\\Significant\\1.xls");
		File file2 = new File("F:\\20180926_F18FTSNWKF1790_I-WAbyS005_Mures_tang\\F18FTSNWKF1790_I-WAbyS005_20180926\\Submit\\BGI_result\\Function_Analyse\\DiffExpressedProtein\\GO_Enrichment\\A-VS-B_GO_enrichment\\1.xls");
		try {
			get(file1,file2);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
