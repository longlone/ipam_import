package excelparser;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Arrays;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class IpAssignmentParser extends AbstractIpAssignmentParser {
	public  IpAssignmentParser(File f) throws FileNotFoundException, IOException{
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
	boolean has2hostnamefield(Sheet sheet, int headerrownum){
		java.util.Iterator<Row> rowIterator = sheet.iterator();
		int i = 0;
		while(rowIterator.hasNext()) {
			Row row = rowIterator.next();
			i++;
			if(i != headerrownum){
				continue;
			}
			int num = 0;
			for(int j = 0 ; j < row.getLastCellNum() ; j++){
				String tmp = getCellString(row, j);
				if(tmp == null){
					continue;
				}
				if(tmp.toLowerCase().contains("hostname")){
					num++;
				}
			}
			return num == 2;
		}
		return false;
	}
	public void test() throws Exception{
		int num = workbook.getNumberOfSheets();
		for (int i = 0 ; i < num ; i++){
			Sheet sheet = workbook.getSheetAt(i);
			if(sheet.getSheetName().matches("10\\.(\\d+)\\.(\\d+)-(\\d+)(.+)vlans(.*)")){
				extract(sheet,has2hostnamefield(sheet,7));
				getsubnets(sheet,has2hostnamefield(sheet,7));
			}else if(sheet.getSheetName().matches("150\\.70\\.(\\d+)-(\\d+)_IP-alloc")){
				extract(sheet,has2hostnamefield(sheet,7));
			}else if(sheet.getSheetName().matches("10\\.(\\d+)\\.(\\d+)-(\\d+)(.+)shared\\+more(.*)")){
				extract(sheet,has2hostnamefield(sheet,7));
				getsubnets(sheet,has2hostnamefield(sheet,7));
			}
		}
		
	}
	public void getsubnets(Sheet sheet , boolean hasdcshostanmefield){
		System.out.println(sheet.getSheetName());
		
		java.util.Iterator<Row> rowIterator = sheet.iterator();
		while(rowIterator.hasNext()) {
			
			Row row = rowIterator.next();
			if(row.getPhysicalNumberOfCells() < 4){
				continue;
			}
			String tmp = getCellString(row, 4);
			if("vlan".equals(tmp)){
				String vlanid = getCellString(row,5);
			
				String ip = getCellInt(row,0) + "." +
							getCellInt(row,1) + "." +
							getCellInt(row,2) + ".0/24";
				String gateway = getCellInt(row,0) + "." +
						getCellInt(row,1) + "." +
						getCellInt(row,2) + ".1";	
				String desc = getCellString(row, hasdcshostanmefield?7:6).trim();
				String project = getCellString(row, hasdcshostanmefield?8:7).trim();
				String owner = getCellString(row, hasdcshostanmefield?11:10).trim();
					
				subnets.add(new String[]{ip, vlanid, gateway, desc, project, owner});
				
			}
		}
	}
	
	public void extract(Sheet sheet , boolean hasdcshostanmefield){
		System.out.println(sheet.getSheetName());
		
		java.util.Iterator<Row> rowIterator = sheet.iterator();
		while(rowIterator.hasNext()) {
			
			Row row = rowIterator.next();
			if(row.getPhysicalNumberOfCells() < 4){
				continue;
			}
			String tmp = getCellString(row, 4);
			if(tmp.matches("(\\d+)(\\.\\d+)?")){
				int i = Double.valueOf(tmp).intValue();
				
			/*
			Cell c = row.getCell(4);
			if(c.getCellType() ==Cell.CELL_TYPE_NUMERIC){
				int i = (int)c.getNumericCellValue();
				*/
				if( i >=1 && i <=255){
					String ip = getCellInt(row,0) + "." +
							getCellInt(row,1) + "." +
							getCellInt(row,2) + "." + i;
					String hostname1 = getCellString(row, 5).trim();
					String hostname2 = getCellString(row, hasdcshostanmefield?6:5).trim();
					String desc = getCellString(row, hasdcshostanmefield?7:6).trim();
					String project = getCellString(row, hasdcshostanmefield?8:7).trim();
					String owner = getCellString(row, hasdcshostanmefield?11:10).trim();
//					Date date = row.getCell(12).getDateCellValue();
					String hostname = null;
					if(hostname1.equals(hostname2)){
						hostname = hostname1;
					}else if(hostname1.length() == 0){
						hostname = hostname2;
					}else if(hostname2.length() == 0){
						hostname = hostname1;
					}else if(hostname1.matches("\\d+")){
						hostname = hostname2;
					}else{
						hostname = hostname1 + "/" + hostname2;
					}
					
					if("".equals(hostname) && "".equals(desc) && "".equals(project) ){
						continue;
					}
					
//					if(ip.startsWith("10.") && i >=0 && i <=31){
//						continue;
//					}
					if("".equals(hostname) && "".equals(desc)){
//						System.out.println(Arrays.toString(new String[]{ip, hostname, desc, project, owner}));
						continue;
					}
					
					assignments.add(new String[]{ip, hostname, desc, project, owner});
				}
				
			}
		}

	}
	
	
	public static void main(String... args) throws FileNotFoundException, IOException, Exception{

		IpAssignmentParser x = new IpAssignmentParser(new File("C:\\Users\\joey_hsiao\\Documents\\ip_assignment\\IAD1_NTTA_IP-list.xlsx"));
		x.test();
		x.close();
//		for(String[] r : x.result){
//			System.out.println(Arrays.toString(r));
//		}
		System.out.println(x.assignments.size());
		System.out.println(x.subnets.size());
		for(String[] r : x.subnets){
			System.out.println(Arrays.toString(r));
		}

/*
		IpAssignmentParser x2 = new IpAssignmentParser(new File("C:\\Users\\joey_hsiao\\Documents\\ip_assignment\\FRA1_EQX_IP-list.xlsx"));
		x2.test();
		x2.close();
//		for(String[] r : x.result){
//			System.out.println(Arrays.toString(r));
//		}
		System.out.println(x2.result.size());

		IpAssignmentParser x3 = new IpAssignmentParser(new File("C:\\Users\\joey_hsiao\\Documents\\ip_assignment\\SJC1_EQX_IP-list.xlsx"));
		x3.test();
		x3.close();
//		for(String[] r : x.result){
//			System.out.println(Arrays.toString(r));
//		}
		System.out.println(x3.result.size());

		IpAssignmentParser x4 = new IpAssignmentParser(new File("C:\\Users\\joey_hsiao\\Documents\\ip_assignment\\MUC1_EQX_IP-list.xlsx"));
		x4.test();
		x4.close();
//		for(String[] r : x.result){
//			System.out.println(Arrays.toString(r));
//		}
		System.out.println(x4.result.size());
	
		IpAssignmentParser x5 = new IpAssignmentParser(new File("C:\\Users\\joey_hsiao\\Documents\\ip_assignment\\SJDC_NTTA_IP-list (Nexus).xlsx"));
		x5.test();
		x5.close();
//		for(String[] r : x.result){
//			System.out.println(Arrays.toString(r));
//		}
		System.out.println(x5.assignments.size());
	*/	
	}
}
