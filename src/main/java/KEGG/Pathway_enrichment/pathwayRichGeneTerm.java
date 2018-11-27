package KEGG.Pathway_enrichment;

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

public class pathwayRichGeneTerm {

		public static void getinfo(File file) throws Exception{
			
			InputStream inputStream = new FileInputStream(file.getAbsolutePath());
			Workbook workbook = Workbook.getWorkbook(inputStream);
			
			Sheet sheet = workbook.getSheet(0);
			int rows = sheet.getRows();
			int columns = sheet.getColumns();
			//表格名称
			WritableWorkbook book = Workbook.createWorkbook(new File("F://pathwayRichGeneTerm.xls"));
			 //页面名称
			//TODO 这里可以暴露出一个参数
			WritableSheet sheet1 = book.createSheet("pathwayRichGeneTerm", 0);
			
			Label label1 = new Label(0,0,"gene_id");
			Label label2 = new Label(1,0,"kegg_term_id");
			sheet1.addCell(label1);
			sheet1.addCell(label2);
			
			List<String> gene_id = new ArrayList<String>();
			List<String> kegg_term = new ArrayList<String>();
			
			String title =null;
			
			//这玩意可以让我看到有多少个混蛋KEGG注释
			int flag = 0;
			//KEGG的第一行啥都没有 真是不好意思
			for (int i = 0; i < rows; i++) {
				
				for (int j = 0; j < columns; j++) {
					
					
					Cell cell = sheet.getCell(j,i);
					String kegg_term_id = cell.getContents();
					
					if (kegg_term_id.contains("|")) {
						title = kegg_term_id;
					}
					if (kegg_term_id!=title&&!kegg_term_id.equals("")&&kegg_term_id.length()!=0) {
						flag++;
						kegg_term.add(kegg_term_id);
						//我需要一个累加器
//						Label label4 = new Label(0,j+1,title);
//						sheet1.addCell(label4);
					}
				}
				//title值这里已经有了
				for (int x = 0; x < flag; x++) {
					gene_id.add(title);
				}
				flag=0;
			}
			
			for (int i = 0; i < kegg_term.size(); i++) {
				String kegg_term_id = kegg_term.get(i);
				Label label3 = new Label(1,i+1,kegg_term_id);
				sheet1.addCell(label3);
				
			}
			for(int i = 0;i < gene_id.size();i++)
			{
				String Protein_ID = gene_id.get(i);
				Label label4 = new Label(0,i+1,Protein_ID);
				sheet1.addCell(label4);
			}
			
			
				book.write();
				book.close();
				
		}	
		
		
		
		public static void main(String[] args) {
			File file = new File("F:\\20180926_F18FTSNWKF1790_I-WAbyS005_Mures_tang\\F18FTSNWKF1790_I-WAbyS005_20180926\\Submit\\BGI_result\\Function_Analyse\\DiffExpressedProtein\\Pathway_Enrichment\\A-VS-B_Pathway_enrichment\\1.xls");
			try {
				getinfo(file);
			} catch (Exception e) {
				System.out.println("他娘的怎么又错了");
			}
			
		}

}
