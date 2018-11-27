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


//该类是用IdentityHashMap来保存重复的键值对的，但是虽然我new String("对象")了,但是现在我只看到KEY  看不到Value了  也是奇怪的很
public class goRichGeneTerm {
	
	public static void getinfo(File file) throws Exception{
		
		InputStream inputStream = new FileInputStream(file.getAbsolutePath());
		Workbook workbook = Workbook.getWorkbook(inputStream);
		
		Sheet sheet = workbook.getSheet(0);
		int rows = sheet.getRows();
		int columns = sheet.getColumns();
		
		Map<String, String> has = new IdentityHashMap<String, String>();
		String title = null;
		
		//第一行的数据不读  因为没什么用
		for (int i = 1; i < rows; i++) {
			
			for (int j = 0; j < columns; j++) {
				
				Cell cell = sheet.getCell(j,i);
				String go = cell.getContents();
				if (go.contains("|")) {
					title = go;
				}
				if (go!=title) {
					has.put(new String(title), go);
				}
				
			}
		}
			for (Entry<String, String>  map : has.entrySet()) {
				
				System.out.println(map.getKey()+" "+map.getValue());
			}
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
