package KEGG.Pathway_enrichment;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import jxl.Cell;
import jxl.NumberCell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class pathwayRichTerm {
	
public static void getinfo(File file) throws Exception{
		
		InputStream inputStream = new FileInputStream(file.getAbsolutePath());
		Workbook workbook = Workbook.getWorkbook(inputStream);
		
		Sheet sheet = workbook.getSheet(0);
		int rows = sheet.getRows();
		
		WritableWorkbook book = Workbook.createWorkbook(new File("F://pathwayRichTerm.xls"));
		 //页面名称
		WritableSheet sheet1 = book.createSheet("pathwayRichTerm", 0);
		
		Label label1 = new Label(0,0,"kegg_term_id");
		Label label2 = new Label(1,0,"kegg_term");
		Label label3 = new Label(2,0,"kegg_level1");
		Label label4 = new Label(3,0,"kegg_level2");
		Label label5 = new Label(4,0,"kegg_term_candidate_gene_num");
		Label label6 = new Label(5,0,"kegg_total_candidate_gene_num");
		Label label7 = new Label(6,0,"kegg_term_gene_num");
		Label label8 = new Label(7,0,"kegg_total_gene_num");
		Label label9 = new Label(8,0,"kegg_rich_ratio");
		Label label10 = new Label(9,0,"kegg_pvalue");
		
		sheet1.addCell(label1);
		sheet1.addCell(label2);
		sheet1.addCell(label3);
		sheet1.addCell(label4);
		sheet1.addCell(label5);
		sheet1.addCell(label6);
		sheet1.addCell(label7);
		sheet1.addCell(label8);
		sheet1.addCell(label9);
		sheet1.addCell(label10);
		
		List<String> list1 = new ArrayList<String>();
		List<String> list2 = new ArrayList<String>();
		List<String> list3 = new ArrayList<String>();
		List<String> list4 = new ArrayList<String>();
		List<String> list5 = new ArrayList<String>();
		List<String> list6 = new ArrayList<String>();
		List<Double> list7 = new ArrayList<Double>();
		
		for (int i = 1; i < rows; i++) {
			Cell cell1 = sheet.getCell(4,i);
			String Pathway_ID = cell1.getContents();
			list1.add(Pathway_ID);
			
			Cell cell2 = sheet.getCell(0,i);
			String Pathway = cell2.getContents();
			list2.add(Pathway);
			
			Cell cell3 = sheet.getCell(5,i);
			String Level1 = cell3.getContents();
			list3.add(Level1);
			
			Cell cell4 = sheet.getCell(6,i);
			String Level2 = cell4.getContents();
			list4.add(Level2);
			
			Cell cell5 = sheet.getCell(1,i);
			String Sample1 = cell5.getContents();
			list5.add(Sample1);
			
			Cell cell6 = sheet.getCell(2,i);
			String Sample2 = cell6.getContents();
			list6.add(Sample2);
			
			NumberCell cell7 = (NumberCell) sheet.getCell(3,i);
			Double Pvalue = cell7.getValue();
			list7.add(Pvalue);
		}	
		
		for (int i = 0; i < list1.size(); i++) {
			String Pathway_ID = list1.get(i);
			Label label11 = new Label(0,i+1,Pathway_ID);
			sheet1.addCell(label11);
			
			String Pathway = list2.get(i);
			Label label12 = new Label(1,i+1,Pathway);
			sheet1.addCell(label12);
			
			String Level1 = list3.get(i);
			Label label13 = new Label(2,i+1,Level1);
			sheet1.addCell(label13);
			
			String Level2 = list4.get(i);
			Label label14 = new Label(3,i+1,Level2);
			sheet1.addCell(label14);
			
			String Sample1 = list5.get(i);
			Label label15 = new Label(4,i+1,Sample1);
			sheet1.addCell(label15);
			
			String total_candidate_gene_num = list5.get(0);
			Label label16 = new Label(5,i+1,total_candidate_gene_num);
			sheet1.addCell(label16);
			
			String Sample2 = list6.get(i);
			Label label17 = new Label(6,i+1,Sample2);
			sheet1.addCell(label17);
			
			String kegg_total_gene_num = list6.get(0);
			Label label18 = new Label(7,i+1,kegg_total_gene_num);
			sheet1.addCell(label18);
			
			Double rich_ratio = Double.valueOf(Sample1)/Double.valueOf(Sample2);
			Label label19 = new Label(8,i+1,String.valueOf(rich_ratio));
			sheet1.addCell(label19);
			
			
			Double Pvalue = list7.get(i);
			Label label20 = new Label(9,i+1,String.valueOf(Pvalue));
			sheet1.addCell(label20);
			
			
		}
		
		book.write();
		book.close();
	}
	
public static void main(String[] args) throws Exception {
		
		File file = new File("F:\\20180926_F18FTSNWKF1790_I-WAbyS005_Mures_tang\\F18FTSNWKF1790_I-WAbyS005_20180926\\Submit\\BGI_result\\Function_Analyse\\DiffExpressedProtein\\Pathway_Enrichment\\A-VS-B_Pathway_enrichment\\2.xls");
		getinfo(file);
	
	}
}	
