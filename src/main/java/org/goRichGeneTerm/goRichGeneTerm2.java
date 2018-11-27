package org.goRichGeneTerm;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.IdentityHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class goRichGeneTerm2 {
	
	public static void getinfo(File file) throws Exception{
		
		InputStream inputStream = new FileInputStream(file.getAbsolutePath());
		Workbook workbook = Workbook.getWorkbook(inputStream);
		
		Sheet sheet = workbook.getSheet(0);
		int rows = sheet.getRows();
		int columns = sheet.getColumns();
		//表格名称
		WritableWorkbook book = Workbook.createWorkbook(new File("F://GORichGeneTerm.xls"));
		 //页面名称
		WritableSheet sheet1 = book.createSheet("GO", 0);
		
		Label label1 = new Label(0,0,"蛋白ID");
		Label label2 = new Label(1,0,"GOid");
		sheet1.addCell(label1);
		sheet1.addCell(label2);
		
		List<String> Protein_Id = new ArrayList<String>();
		List<String> GO = new ArrayList<String>();
		
		String title =null;
		
		//这玩意可以让我看到有多少个混蛋GO注释
		int flag = 0;
		//第一行的数据不读  因为没什么用
		for (int i = 1; i < rows; i++) {
			
			for (int j = 0; j < columns; j++) {
				
				
				Cell cell = sheet.getCell(j,i);
				String go = cell.getContents();
				
				if (go.contains("|")) {
					title = go;
				}
				if (go!=title&&!go.equals("")&&go.length()!=0) {
					flag++;
					GO.add(go);
					//我需要一个累加器
//					Label label4 = new Label(0,j+1,title);
//					sheet1.addCell(label4);
				}
			}
			//title值这里已经有了
			for (int x = 0; x < flag; x++) {
				Protein_Id.add(title);
			}
			flag=0;
		}
		
		for (int i = 0; i < GO.size(); i++) {
			String GO_id = GO.get(i);
			Label label3 = new Label(1,i+1,GO_id);
			sheet1.addCell(label3);
			
		}
		for(int i = 0;i < Protein_Id.size();i++)
		{
			String Protein_ID = Protein_Id.get(i);
			Label label4 = new Label(0,i+1,Protein_ID);
			sheet1.addCell(label4);
		}
		
		
			book.write();
			book.close();
			
	}	
	
	
	
	public static void main(String[] args) {
		File file = new File("F:\\20180926_F18FTSNWKF1790_I-WAbyS005_Mures_tang\\F18FTSNWKF1790_I-WAbyS005_20180926\\Submit\\BGI_result\\Function_Analyse\\DiffExpressedProtein\\GO_Enrichment\\A-VS-B_GO_enrichment\\2.xls");
		try {
			getinfo(file);
		} catch (Exception e) {
			System.out.println("他娘的怎么又错了");
		}
		
	}
}
