package excelparser;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class IpAssignmentParser3 extends AbstractIpAssignmentParser {
	
	public  IpAssignmentParser3(File f) throws FileNotFoundException, IOException{
		FileInputStream file = new FileInputStream(f);
		
		try{
			workbook = new HSSFWorkbook(file);
		}catch(org.apache.poi.poifs.filesystem.OfficeXmlFileException e){
			file.close();
			file = new FileInputStream(f);
			workbook = new XSSFWorkbook(file);
			file.close();
		}
		
	}
	
	public void test() throws Exception{
		int num = workbook.getNumberOfSheets();
		for (int i = 0 ; i < num ; i++){
			Sheet sheet = workbook.getSheetAt(i);
			if(sheet.getSheetName().toLowerCase().contains("assignment") || sheet.getSheetName().toLowerCase().contains("public segment")){
				extract(sheet,9);
			}
			
		}
		
	}
	
	public void extract(Sheet sheet, int startIdx ){
//		System.out.println(sheet.getSheetName());
		int rowidx = 0;
		java.util.Iterator<Row> rowIterator = sheet.iterator();
		while(rowIterator.hasNext()) {
			Row row = rowIterator.next();
			rowidx++;
			
			if(rowidx <startIdx){
				continue;
			}
			if(row.getLastCellNum() < 2){
				continue;
			}
			for(int colidx = 0 ; colidx < row.getLastCellNum() ; colidx+=4){
				String ip = getCellString(row, colidx);
				if(ip == null){
					continue;
				}
				
				if(!ip.matches("(\\d+)\\.(\\d+)\\.(\\d+)\\.(\\d+)")){
					continue;
				}
				String hostname = getCellString(row, colidx+2).trim();
				if("network-id".equals(hostname) || "Network Address".equals(hostname)){
					continue;
				}
				String desc = getCellString(row, colidx+3).trim();
				if((hostname == null || "".equals(hostname) ) && (desc == null || "".equals(desc))){
					continue;
				}
				assignments.add(new String[]{ip, hostname, desc});
			}
			
		}

	}
	
	
	public static void main(String... args) throws FileNotFoundException, IOException, Exception{

		IpAssignmentParser3 x = new IpAssignmentParser3(new File("C:\\Users\\joey_hsiao\\Documents\\ip_assignment\\UDC2_NoneSDI_IPAssignList.xls"));
		x.test();
		x.close();
		for(String[] r : x.assignments){
			System.out.println(Arrays.toString(r));
		}
		System.out.println("***"+x.assignments.size());

		IpAssignmentParser3 x2 = new IpAssignmentParser3(new File("C:\\Users\\joey_hsiao\\Documents\\ip_assignment\\YY-IP-AllocateTable.xls"));
		x2.test();
		x2.close();
		for(String[] r : x.assignments){
			System.out.println(Arrays.toString(r));
		}
		System.out.println("***"+x2.assignments.size());

	}
}
