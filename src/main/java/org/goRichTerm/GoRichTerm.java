package org.goRichTerm;

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

public class GoRichTerm {
	
	public static void getinfo(File file) throws Exception{
		
		InputStream inputStream = new FileInputStream(file.getAbsolutePath());
		Workbook workbook = Workbook.getWorkbook(inputStream);
		
		Sheet sheet = workbook.getSheet(0);
		int rows = sheet.getRows();
		
		WritableWorkbook book = Workbook.createWorkbook(new File("F://GORichTerm.xls"));
		 //页面名称
		WritableSheet sheet1 = book.createSheet("goRichTerm", 0);
		
		Label label1 = new Label(0,0,"go_term_id");
		Label label2 = new Label(1,0,"go_term");
		Label label3 = new Label(2,0,"go_level1");
		Label label4 = new Label(3,0,"go_levele2");
		Label label5 = new Label(4,0,"go_term_candidate_gene_num");
		Label label6 = new Label(5,0,"go_total_candidate_gene_num");
		Label label7 = new Label(6,0,"go_term_gene_num");
		Label label8 = new Label(7,0,"go_total_gene_num");
		Label label9 = new Label(8,0,"go_rich_ratio");
		Label label10 = new Label(9,0,"go_pvalue");
		
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
		
		List<String> go_term_id = new ArrayList<String>();
		List<String> go_term = new ArrayList<String>();
		List<String> go_term_candidate_gene_num = new ArrayList<String>();
		List<String> go_total_candidate_gene_num = new ArrayList<String>();
		List<String> go_term_gene_num = new ArrayList<String>();
		List<String> go_total_gene_num = new ArrayList<String>();
		List<Double> go_rich_ratio = new ArrayList<Double>();
		List<Double> go_pvalue = new ArrayList<Double>();
		
		for (int i = 0; i < rows; i++) {
			Cell cell1 = sheet.getCell(0,i);
			String s1 = cell1.getContents();
			go_term_id.add(s1);
			
			Cell cell2 = sheet.getCell(1,i);
			String s2 = cell2.getContents();
			go_term.add(s2);
			
			Cell cell3 = sheet.getCell(2,i);
			String s3 = cell3.getContents();
			int index1 = s3.indexOf("of");
			//这是注释到此GO Term的蛋白数
			String candidate_gene_num = s3.substring(0,index1-1);
			go_term_candidate_gene_num.add(candidate_gene_num);
			
			//这是总的选定集合蛋白数
			int index2 = s3.indexOf("in");
			String total_candidate_gene_num = s3.substring(index1+3,index2-1);
			go_total_candidate_gene_num.add(total_candidate_gene_num);
			
			Cell cell4 = sheet.getCell(3,i);
			String s4 = cell4.getContents();
			//本物种注释到此GO Term的蛋白数
			int index3 = s4.indexOf("of");
			String term_gene_num = s4.substring(0, index3-1);
			go_term_gene_num.add(term_gene_num);
			
			//本物种总的蛋白数
			int index4 = s4.indexOf("in");
			String total_gene_num = s4.substring(index3+3,index4-1);
			go_total_gene_num.add(total_gene_num);
			
			//P-value 通过numvercell可以将其全部取完
			NumberCell cell5 = (NumberCell) sheet.getCell(4,i);
			double s5 = cell5.getValue();
			go_pvalue.add(s5);
			
			
		}
		//level1
		if (sheet.getName().contains("P")) {
			for (int i = 0; i < go_term_id.size(); i++) {
				String s1 = go_term_id.get(i);
				Label label11 = new Label(0,i+1,s1);
				sheet1.addCell(label11);
				
				Label label13 = new Label(2,i+1,"biological_process");
				sheet1.addCell(label13);
			}
		}
		else if (sheet.getName().contains("C")) {
			
		}
		else if (sheet.getName().contains("F")) {
			
		}
		for (int i = 0; i < go_term_candidate_gene_num.size(); i++) {
			String s1 = go_term_candidate_gene_num.get(i);
			Label label14 = new Label(4,i+1,s1);
			sheet1.addCell(label14);
			
			String s2 = go_total_candidate_gene_num.get(i);
			Label label15 = new Label(5,i+1,s2);
			sheet1.addCell(label15);
			
			String s3 = go_term_gene_num.get(i);
			Label label16 = new Label(6,i+1,s3);
			sheet1.addCell(label16);
			
			String s4 = go_total_gene_num.get(i);
			Label label17 = new Label(7,i+1,s4);
			sheet1.addCell(label17);
			
			double s5 = go_pvalue.get(i);
			Label label18 = new Label(9,i+1,(String.valueOf(s5)));
			sheet1.addCell(label18);
			
			//go_rich_ratio = go_term_candidate_gene_num / go_term_gene_num
			double s6 = (Double.valueOf(s1))/(Double.valueOf(s3));
			String s7 = (String.valueOf(s6));
			Label label19 = new Label(8,i+1,s7);
			sheet1.addCell(label19);
		}
		
		
		
		for (int i = 0; i < go_term.size(); i++) {
			String s2 = go_term.get(i);
			Label label12 = new Label(1,i+1,s2);
			sheet1.addCell(label12);
		}
		
		
		book.write();
		book.close();
	}
	
	public static void main(String[] args) throws Exception {
		
		File file = new File("F:\\20180926_F18FTSNWKF1790_I-WAbyS005_Mures_tang\\F18FTSNWKF1790_I-WAbyS005_20180926\\Submit\\BGI_result\\Function_Analyse\\DiffExpressedProtein\\GO_Enrichment\\A-VS-B_GO_enrichment\\3.xls");
		getinfo(file);
	}
}
