package excelparser;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class IpAssignmentParser2 extends AbstractIpAssignmentParser {
	
	public  IpAssignmentParser2(File f) throws FileNotFoundException, IOException{
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
			extract(sheet);
			getsubnets(sheet);
		}
		
	}
	
	public void extract(Sheet sheet ){
//		System.out.println(sheet.getSheetName());
		
		java.util.Iterator<Row> rowIterator = sheet.iterator();
		while(rowIterator.hasNext()) {
			Row row = rowIterator.next();
			if(row.getPhysicalNumberOfCells() < 2){
				continue;
			}
			
			String ipseg = getCellString(row, 0);
			if(ipseg.matches("(\\d+)\\.(\\d+)\\.(\\d+)")){
				Integer ipseg4 = getCellInt(row, 1);
				if( ipseg4 == null){
					continue;
				}
				if( ipseg4 >=1 && ipseg4 <=255){
					String ip = ipseg + "." + ipseg4;
					String hostname = getCellString(row, 2).trim();
					String desc = getCellString(row, 3).trim();
					String project = getCellString(row, 4).trim();
					String owner = getCellString(row, 5).trim();
//					Date date = row.getCell(12).getDateCellValue();
					
					if("".equals(hostname) && "".equals(desc) && "".equals(project) ){
						continue;
					}
					
					if("".equals(hostname) && "".equals(desc)){
						continue;
					}
					
					assignments.add(new String[]{ip, hostname, desc, project, owner});
				}
			}
		}
	}

//		Network:  10.48.60.128/26 (Service 12 DB server segment)						
//		Gateway IP:  10.48.60.190						
	public void getsubnets(Sheet sheet){
		System.out.println(sheet.getSheetName());
		String nw = null;
		String gateway = null;
		String vlanid = null;
		String desc = null;
		
		java.util.Iterator<Row> rowIterator = sheet.iterator();
		while(rowIterator.hasNext()) {
			
			Row row = rowIterator.next();
			if(row.getPhysicalNumberOfCells() < 2){
				continue;
			}
					
			String line = getCellString(row, 0);
			Pattern p1 = Pattern.compile("Network\\s*:\\s*+([\\d\\./]+)",Pattern.CASE_INSENSITIVE);
			Pattern p2 = Pattern.compile("\\((.+)\\)");
			Pattern p3 = Pattern.compile("vlan\\s*(\\d+)",Pattern.CASE_INSENSITIVE);
			Pattern p4 = Pattern.compile("Gateway\\s*IP\\s*:\\s*([\\d\\./]+)",Pattern.CASE_INSENSITIVE);
			Matcher matcher;
			matcher = p1.matcher(line);
			if(matcher.find()){
				if(nw != null){
					subnets.add(new String[]{nw, vlanid, gateway, desc});
					nw = null;
					gateway = null;
					vlanid = null;
					desc = null;
				}
				nw = matcher.group(1);
			}
			matcher = p2.matcher(line);
			if(matcher.find()){
				 desc = matcher.group(1);
			}
			matcher = p3.matcher(line);
			if(matcher.find()){
				vlanid = matcher.group(1);
			}
			matcher = p4.matcher(line);
			if(matcher.find()){
				gateway = matcher.group(1);
			}
		}
	}
	
	public static void main(String... args) throws FileNotFoundException, IOException, Exception{

		IpAssignmentParser2 x = new IpAssignmentParser2(new File("C:\\Users\\joey_hsiao\\Documents\\ip_assignment\\SJDC-SDI-BA-CBU-Network-IP-Address.xlsx"));
		x.test();
		x.close();
		for(String[] r : x.subnets){
			System.out.println(Arrays.toString(r));
		}
		System.out.println("***"+x.assignments.size());
/*
		IpAssignmentParser2 x2 = new IpAssignmentParser2(new File("C:\\Users\\joey_hsiao\\Documents\\ip_assignment\\SJDC-SDI-CT2-CBU-Network-IP-Address.xlsx"));
		x2.test();
		x2.close();
//		for(String[] r : x.result){
//			System.out.println(Arrays.toString(r));
//		}
		System.out.println("***"+x2.assignments.size());

		IpAssignmentParser2 x3 = new IpAssignmentParser2(new File("C:\\Users\\joey_hsiao\\Documents\\ip_assignment\\SJDC-SDI-CT3-CBU-Network-IP-AddressV2.xlsx"));
		x3.test();
		x3.close();
//		for(String[] r : x.result){
//			System.out.println(Arrays.toString(r));
//		}
		System.out.println("***"+x3.assignments.size());

		IpAssignmentParser2 x4 = new IpAssignmentParser2(new File("C:\\Users\\joey_hsiao\\Documents\\ip_assignment\\SJDC-SDI-CT-CBU-Network-IP-Address.xls"));
		x4.test();
		x4.close();
//		for(String[] r : x.result){
//			System.out.println(Arrays.toString(r));
//		}
		System.out.println("***"+x4.assignments.size());
		
		IpAssignmentParser2 x5 = new IpAssignmentParser2(new File("C:\\Users\\joey_hsiao\\Documents\\ip_assignment\\SJDC-SDI-NBU-CBU-Network-IP-Address.xls"));
		x5.test();
		x5.close();
//		for(String[] r : x.result){
//			System.out.println(Arrays.toString(r));
//		}
		System.out.println("***"+x5.assignments.size());
		
		IpAssignmentParser2 x6 = new IpAssignmentParser2(new File("C:\\Users\\joey_hsiao\\Documents\\ip_assignment\\UDC2-SDI-SMB-CBU-Network-IP-Address.xls"));
		x6.test();
		x6.close();
//		for(String[] r : x.result){
//			System.out.println(Arrays.toString(r));
//		}
		System.out.println("***"+x6.assignments.size());
		
	*/	
	}
}
