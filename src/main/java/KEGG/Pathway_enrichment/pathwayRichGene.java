package KEGG.Pathway_enrichment;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class pathwayRichGene {
	public static void main(String[] args) throws RowsExceededException, BiffException, WriteException, IOException {
		File file1 = new File("F:\\20180926_F18FTSNWKF1790_I-WAbyS005_Mures_tang\\F18FTSNWKF1790_I-WAbyS005_20180926\\Submit\\BGI_result\\Significant\\1.xls");
		File file2 = new File("F:\\20180926_F18FTSNWKF1790_I-WAbyS005_Mures_tang\\F18FTSNWKF1790_I-WAbyS005_20180926\\Submit\\BGI_result\\Function_Analyse\\DiffExpressedProtein\\Pathway_Enrichment\\A-VS-B_Pathway_enrichment\\3.xls");
		File file3 = new File("F:\\20180926_F18FTSNWKF1790_I-WAbyS005_Mures_tang\\F18FTSNWKF1790_I-WAbyS005_20180926\\Submit\\BGI_result\\Function_Analyse\\DiffExpressedProtein\\Pathway_Enrichment\\A-VS-B_Pathway_enrichment\\4.xls");
		
		getinfo(file1,file2,file3);
	}
	
	public static void getinfo(File file1,File file2,File file3) throws BiffException, IOException, RowsExceededException, WriteException
	{
		InputStream inputStream = new FileInputStream(file1.getAbsolutePath());
		Workbook workbook = Workbook.getWorkbook(inputStream);
		
		Sheet sheet = workbook.getSheet(0);
		int rows = sheet.getRows();
		
		InputStream inputStream1 = new FileInputStream(file2.getAbsolutePath());
		Workbook workbook1 = Workbook.getWorkbook(inputStream1);
		
		Sheet sheet2 = workbook1.getSheet(0);
		int rows1 = sheet2.getRows();
		
		InputStream inputStream2 = new FileInputStream(file3.getAbsolutePath());
		Workbook workbook2 = Workbook.getWorkbook(inputStream2);
		
		Sheet sheet3 = workbook2.getSheet(0);
		int rows2 = sheet3.getRows();
		
		WritableWorkbook book = Workbook.createWorkbook(new File("F://pathwayRichGene.xls"));
		 //页面名称
		//TODO 这里可以暴露出一个参数
		WritableSheet sheet1 = book.createSheet("pathwayRichGene", 0);
		
		Label label1 = new Label(0,0,"gene_id");
		Label label2 = new Label(1,0,"kegg_term_mix");
		Label label3 = new Label(2,0,"kegg_level1_mix");
		Label label4 = new Label(3,0,"kegg_level2_mix");
		Label label5 = new Label(4,0,"diffexp_log2fc");
		
		sheet1.addCell(label1);
		sheet1.addCell(label2);
		sheet1.addCell(label3);
		sheet1.addCell(label4);
		sheet1.addCell(label5);
		
		//蛋白ID
		List<String> Protein = new ArrayList<String>();
		//我们表的比较值 但是在另一个表中是各样品中的基因表达量比值
		List<String> log2 = new ArrayList<String>();
		
		List<String> Pathway1 = new ArrayList<String>();
		
		List<String> Pathway2 = new ArrayList<String>();
		
		List<String> Ko = new ArrayList<String>();
		
		StringBuffer stringBuilder1 = new StringBuffer();
		StringBuffer stringBuilder2 = new StringBuffer();
		
		//该表是含有表头的  所以从1开始
		for (int i = 1; i < rows; i++) {
			Cell cell1 = sheet.getCell(1,i);
			String Protein_Id = cell1.getContents();
			Protein.add(Protein_Id);
			
			Cell cell2 = sheet.getCell(8,i);
			String log = cell2.getContents();
			log2.add(log);
			
			for (int j = 0; j < rows1; j++) {
				
				Cell cell3 = sheet2.getCell(3,j);
				String mix_Protein_Id = cell3.getContents();
				
				if (mix_Protein_Id.contains(Protein_Id)) {
					//如果包含  则把这些信息打印出来
					Cell cell4 = sheet2.getCell(0,j);
					String Pathway_level1 = cell4.getContents();
					stringBuilder1.append(Pathway_level1+";");
					
					Cell cell5 = sheet2.getCell(1,j);
					String Pathway_level2 = cell5.getContents();
					stringBuilder2.append(Pathway_level2+";");
					
				}
			}
			Pathway1.add(stringBuilder1.toString());
			Pathway2.add(stringBuilder2.toString());
			
			//这两个字符缓存器在用完之后需要清空
			stringBuilder1.delete(0,stringBuilder1.length());
			stringBuilder2.delete(0,stringBuilder2.length());
			
			//读取第三个表获取我们的ko信息
			for (int j = 0; j < rows2; j++) {
				Cell cell6 = sheet3.getCell(0,j);
				String match_Protein = cell6.getContents();
				//如果表中还有才去查找
				if (match_Protein.equals(Protein_Id)) {
					
					Cell cell7 = sheet3.getCell(1,j);
					String ko_information = cell7.getContents();
					//如果信息不为空时才输出
					if (ko_information.length()>10) {
						//获取到ko_id
						String ko_id = ko_information.substring(0, 6);
						
						int index = ko_information.lastIndexOf("|");
						String ko_info = ko_information.substring(index+1,ko_information.length());
						String ko = ko_id+"//"+ko_info;
						Ko.add(ko);
					}
					else
					{
						//匹配不到也插入一个空值
						Ko.add("");
					}
				}
			}
			
		}			
		for (int j = 0; j < Protein.size(); j++) {
			String Protein_Id = Protein.get(j);
			Label label6 = new Label(0,j+1,Protein_Id);
			sheet1.addCell(label6);
			
			String Pathway_level1 = Pathway1.get(j).toString();
			Label label7 = new Label(2,j+1,Pathway_level1);
			sheet1.addCell(label7);
			String Pathway_level2 = Pathway2.get(j).toString();
			Label label8 = new Label(3,j+1,Pathway_level2);
			sheet1.addCell(label8);
			
			String log = log2.get(j);
			Label label9 = new Label(4,j+1,log);
			sheet1.addCell(label9);
			
			String ko = Ko.get(j);
			Label label20 = new Label(1,j+1,ko);
			sheet1.addCell(label20);
			
		}
		
		
		book.write();
		book.close();
	}
}
