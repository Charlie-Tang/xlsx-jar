package org.PdfTools;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import com.itextpdf.text.xml.simpleparser.NewLineHandler;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class GoRichGene {
	
	public static void readColumn(File file,int index) throws Exception
	{
		InputStream inputStream = new FileInputStream(file.getAbsolutePath());
		Workbook workbook = Workbook.getWorkbook(inputStream);
		
		Sheet sheet = workbook.getSheet(0);
		int rows = sheet.getRows();
//		int column = sheet.getColumns();
		
		List<String> list1 = new ArrayList<String>();
		List<String> list2 = new ArrayList<String>();
		List<String> list3 = new ArrayList<String>();
		List<String> list4 = new ArrayList<String>();
		
//		System.out.println(rows);
//		System.out.println(column);
		
		 WritableWorkbook book = Workbook.createWorkbook(new File("F://GoRichGene.xls"));
		 
		 WritableSheet sheet2 = book.createSheet("test", 0);
		 
		 //插入蛋白行
		 Label label1 = new Label(0,0,"gene_id");
		 //插入综合注释行
		 Label label2 = new Label(1,0,"go_term_mix");
		 //插入go_level1_mix
		 Label label3 = new Label(2,0,"go_level1_mix");
		 
		 Label label4 = new Label(3, 0, "go_level2_mix");
		 
		 Label label5 = new Label(4,0,"diffexp_log2fc");
		 
		 sheet2.addCell(label1);
		 sheet2.addCell(label2);
		 sheet2.addCell(label3);
		 sheet2.addCell(label4);
		 sheet2.addCell(label5);
        
		 //循环读取表
		 for (int i = 1; i < rows; i++) {
			 
			 Cell cell1 = sheet.getCell(index,i);//[P]
			 Cell cell2 = sheet.getCell(index+1,i);//[F]
			 Cell cell3 = sheet.getCell(index+2,i);//[C]
			 Cell cell4 = sheet.getCell(1,i);
			 Cell cell5 = sheet.getCell(8,i);
			 String s1 = "";
			 String s2 = "";
			 String s3 = "";
			 String go_level1_mix = "";
			 
			 if (cell1.getContents().equals("") && cell2.getContents()!=null && cell3.getContents()!=null) {
				
				 s1 = "";
				 s2 = "[F]"+cell2.getContents()+";";
				 s3 = "[C]"+cell3.getContents()+";";
				 go_level1_mix="cellular_component;molecular_function";
			}
			 else if (cell1.getContents().equals("") && cell2.getContents().equals("") && cell3.getContents()!=null) {
				
				 s1 = "";
				 s2 = "";
				 s3 = "[C]"+cell3.getContents()+";";
				 
				 go_level1_mix = "molecular_function";
			}
			 else if (cell1.getContents().equals("") && cell2.getContents()!=null && cell3.getContents().equals("")) {
				
				 s1 = "";
				 s2 = "[F]"+cell2.getContents()+";";
				 s3 = "";
				 
				 go_level1_mix = "molecular_function";
			}
			 else if (cell1.getContents().equals("") && cell2.getContents().equals("") && cell3.getContents().equals("")) {
				
				 s1 = "";
				 s2 = "";
				 s3 = "";
				 go_level1_mix = "";
			}
			 else if (cell1.getContents()!=null && cell2.getContents().equals("") && cell3.getContents()!=null) {
				s1 = "[P]"+cell1.getContents()+";";
				s2 = "";
				s3 = "[C]"+cell3.getContents()+";";
				
				 go_level1_mix = "biological_process;cellular_component";
			}
			 else if(cell1.getContents()!=null && cell2.getContents().equals("") && cell3.getContents().equals(""))
			 {
				 s1 = "[P]"+cell1.getContents()+";";
				 s2 = "";
				 s3 = "";
				 
				 go_level1_mix="biological_process";
			}
			 else if (cell1.getContents()!=null && cell2.getContents()!=null && cell3.getContents().equals("")) {
				s1 = "[P]"+cell1.getContents()+";";
				s2 = "[F]"+cell2.getContents()+";";
				s3 = "";
				
				go_level1_mix="biological_process;molecular_function";
			}
			 else if (cell1.getContents()!=null && cell2.getContents()!=null && cell3.getContents()!=null) {
				 s1 = "[P]"+cell1.getContents()+";";
				 s2 = "[F]"+cell2.getContents()+";";
				 s3 = "[C]"+cell3.getContents()+";";
				 
				 go_level1_mix = "biological_process;cellular_component;molecular_function";
			 }
			 String s = s1+s2+s3;
			 String id = cell4.getContents();
			 String Mean_Ratio = cell5.getContents();
			 
			 list1.add(id);
			 list2.add(s);
			 list3.add(go_level1_mix);
			 list4.add(Mean_Ratio);
		}
		 //第一列
		 for (int i = 0; i < list1.size(); i++) {
			String s = list1.get(i);
			Label label = new Label(0,i+1,s);
			sheet2.addCell(label);
		}
		 //第二列
		 for (int i = 0; i < list2.size(); i++) {
			 String s = list2.get(i);
			 Label label6 = new Label(1,i+1,s);
			 sheet2.addCell(label6);
		}
		 //第三列
		 for (int j = 0; j < list3.size(); j++) {
			String s = list3.get(j);
			Label label7 = new Label(2,j+1,s);
			sheet2.addCell(label7);
		}
		 @SuppressWarnings("static-access")
		List<String> list = new Append().get(new File("F:\\20180926_F18FTSNWKF1790_I-WAbyS005_Mures_tang\\F18FTSNWKF1790_I-WAbyS005_20180926\\Submit\\BGI_result\\Significant\\1.xls"), 
				 new File("F:\\20180926_F18FTSNWKF1790_I-WAbyS005_Mures_tang\\F18FTSNWKF1790_I-WAbyS005_20180926\\Submit\\BGI_result\\Function_Analyse\\DiffExpressedProtein\\GO_Enrichment\\A-VS-B_GO_enrichment\\1.xls"));
		 
		 //第四列 调用数据
		 for (int i = 0; i < list.size(); i++) {
			String s = list.get(i);
			Label label9 = new Label(3,i+1,s);
			sheet2.addCell(label9);
		}
		 //第五列
		 for (int i = 0; i < list4.size(); i++) {
			 String s = list4.get(i);
			 Label label8 = new Label(4,i+1,s);
			 sheet2.addCell(label8);
		}
		 book.write();
		 book.close();
		 
	}
	 public static void main(String[] args) {
		File file = new File("F:\\20180926_F18FTSNWKF1790_I-WAbyS005_Mures_tang\\F18FTSNWKF1790_I-WAbyS005_20180926\\Submit\\BGI_result\\Significant\\1.xls");
		try {
			readColumn(file, 31);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
}
