package excelparser;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

public abstract class AbstractIpAssignmentParser {

	protected Workbook workbook;
	protected List<String[]> assignments = new ArrayList<String[]>();
	protected List<String[]> allocations = new ArrayList<String[]>();
	protected List<String[]> subnets = new ArrayList<String[]>();

	public AbstractIpAssignmentParser() {
		super();
	}

	protected String getCellString(Row row, int idx) {
		Cell c = row.getCell(idx);
		if (c == null){
			return "";
		}
		if(c.getCellType() == Cell.CELL_TYPE_BLANK){
			return "";
		}
		if (c.getCellType() == Cell.CELL_TYPE_NUMERIC){
			return String.valueOf(c.getNumericCellValue());
		}else if (c.getCellType() == Cell.CELL_TYPE_BOOLEAN){
			return String.valueOf(c.getBooleanCellValue());
		}else if(c.getCellType() ==Cell.CELL_TYPE_STRING){
			return c.getStringCellValue();
		}
		
		//c.getCellType() == Cell.CELL_TYPE_FORMULA
		try{
			return c.getStringCellValue();
		}catch(IllegalStateException e){
			//
		}
		return String.valueOf((int)c.getNumericCellValue());
	}

	protected int getCellInt(Row row, int idx) {
		Cell c = row.getCell(idx);
		if (c == null){
			return 0;
		}
		if(c.getCellType() == Cell.CELL_TYPE_BLANK){
			return 0;
		}
		if (c.getCellType() == Cell.CELL_TYPE_NUMERIC){
			return (int)c.getNumericCellValue();
		}else if (c.getCellType() == Cell.CELL_TYPE_BOOLEAN){
			return c.getBooleanCellValue()?1:0;
		}else if(c.getCellType() ==Cell.CELL_TYPE_STRING){
			String x = c.getStringCellValue();
			if(x == null){
				return 0;
			}
			if (x.endsWith(".")){
				x = x.replaceAll("\\.$", "");
			}
			if( x.trim().equals("")){
				return 0;
			}
			return Integer.valueOf(x);
		}
		//c.getCellType() == Cell.CELL_TYPE_FORMULA
		try{
			return (int)c.getNumericCellValue();
		}catch(IllegalStateException e){
			//
		}
		return Integer.valueOf(c.getStringCellValue());
	}

	public abstract void test() throws Exception;
	
	public void close() throws IOException {
		workbook.close();
	}

	public static void main(String... args) throws FileNotFoundException, IOException, Exception{
		PrintWriter assignmentOuts = new PrintWriter(new BufferedWriter(new FileWriter("C:\\Users\\joey_hsiao\\Documents\\ip_assignment\\assignments.csv")));
		PrintWriter subnetOuts = new PrintWriter(new BufferedWriter(new FileWriter("C:\\Users\\joey_hsiao\\Documents\\ip_assignment\\subnets.csv")));
		AbstractIpAssignmentParser x = new IpAssignmentParser(new File("C:\\Users\\joey_hsiao\\Documents\\ip_assignment\\IAD1_NTTA_IP-list.xlsx"));
		
		x.test();
		x.close();
		for(String[] r : x.assignments){
			assignmentOuts.println(Arrays.toString(r));
		}
		
		for(String[] r : x.subnets){
			subnetOuts.println(Arrays.toString(r));
		}
		assignmentOuts.println(x.assignments.size());

		x = new IpAssignmentParser(new File("C:\\Users\\joey_hsiao\\Documents\\ip_assignment\\FRA1_EQX_IP-list.xlsx"));
		x.test();
		x.close();
		for(String[] r : x.assignments){
			assignmentOuts.println(Arrays.toString(r));
		}
		
		for(String[] r : x.subnets){
			subnetOuts.println(Arrays.toString(r));
		}
		
		System.out.println(x.assignments.size());

		x = new IpAssignmentParser(new File("C:\\Users\\joey_hsiao\\Documents\\ip_assignment\\SJC1_EQX_IP-list.xlsx"));
		x.test();
		x.close();
		for(String[] r : x.assignments){
			assignmentOuts.println(Arrays.toString(r));
		}
		
		for(String[] r : x.subnets){
			subnetOuts.println(Arrays.toString(r));
		}
		System.out.println(x.assignments.size());

		x = new IpAssignmentParser(new File("C:\\Users\\joey_hsiao\\Documents\\ip_assignment\\MUC1_EQX_IP-list.xlsx"));
		x.test();
		x.close();
		for(String[] r : x.assignments){
			assignmentOuts.println(Arrays.toString(r));
		}
		
		for(String[] r : x.subnets){
			subnetOuts.println(Arrays.toString(r));
		}
		System.out.println(x.assignments.size());
		
		x = new IpAssignmentParser(new File("C:\\Users\\joey_hsiao\\Documents\\ip_assignment\\SJDC_NTTA_IP-list (Nexus).xlsx"));
		x.test();
		x.close();
		for(String[] r : x.assignments){
			assignmentOuts.println(Arrays.toString(r));
		}
		
		for(String[] r : x.subnets){
			subnetOuts.println(Arrays.toString(r));
		}
		System.out.println(x.assignments.size());
		
		x = new IpAssignmentParser2(new File("C:\\Users\\joey_hsiao\\Documents\\ip_assignment\\SJDC-SDI-BA-CBU-Network-IP-Address.xlsx"));
		x.test();
		x.close();
		for(String[] r : x.assignments){
			assignmentOuts.println(Arrays.toString(r));
		}
		for(String[] r : x.subnets){
			subnetOuts.println(Arrays.toString(r));
		}
		System.out.println("***"+x.assignments.size());

		x = new IpAssignmentParser2(new File("C:\\Users\\joey_hsiao\\Documents\\ip_assignment\\SJDC-SDI-CT2-CBU-Network-IP-Address.xlsx"));
		x.test();
		x.close();
		for(String[] r : x.assignments){
			assignmentOuts.println(Arrays.toString(r));
		}
		for(String[] r : x.subnets){
			subnetOuts.println(Arrays.toString(r));
		}
		System.out.println("***"+x.assignments.size());

		x = new IpAssignmentParser2(new File("C:\\Users\\joey_hsiao\\Documents\\ip_assignment\\SJDC-SDI-CT3-CBU-Network-IP-AddressV2.xlsx"));
		x.test();
		x.close();
		for(String[] r : x.assignments){
			assignmentOuts.println(Arrays.toString(r));
		}
		for(String[] r : x.subnets){
			subnetOuts.println(Arrays.toString(r));
		}
		System.out.println("***"+x.assignments.size());

		x = new IpAssignmentParser2(new File("C:\\Users\\joey_hsiao\\Documents\\ip_assignment\\SJDC-SDI-CT-CBU-Network-IP-Address.xls"));
		x.test();
		x.close();
		for(String[] r : x.assignments){
			assignmentOuts.println(Arrays.toString(r));
		}
		System.out.println("***"+x.assignments.size());
		
		x = new IpAssignmentParser2(new File("C:\\Users\\joey_hsiao\\Documents\\ip_assignment\\SJDC-SDI-NBU-CBU-Network-IP-Address.xls"));
		x.test();
		x.close();
		for(String[] r : x.assignments){
			assignmentOuts.println(Arrays.toString(r));
		}
		for(String[] r : x.subnets){
			subnetOuts.println(Arrays.toString(r));
		}
		System.out.println("***"+x.assignments.size());
		/*
		x = new IpAssignmentParser2(new File("C:\\Users\\joey_hsiao\\Documents\\ip_assignment\\UDC2-SDI-SMB-CBU-Network-IP-Address.xls"));
		x.test();
		x.close();
		for(String[] r : x.assignments){
			assignmentOuts.println(Arrays.toString(r));
		}
		for(String[] r : x.subnets){
			subnetOuts.println(Arrays.toString(r));
		}
		System.out.println("***"+x.assignments.size());
		
		x = new IpAssignmentParser3(new File("C:\\Users\\joey_hsiao\\Documents\\ip_assignment\\UDC2_NoneSDI_IPAssignList.xls"));
		x.test();
		x.close();
		for(String[] r : x.assignments){
			assignmentOuts.println(Arrays.toString(r));
		}
		for(String[] r : x.subnets){
			subnetOuts.println(Arrays.toString(r));
		}
		System.out.println("***"+x.assignments.size());
		*/
		x = new IpAssignmentParser3(new File("C:\\Users\\joey_hsiao\\Documents\\ip_assignment\\YY-IP-AllocateTable.xls"));
		x.test();
		x.close();
		for(String[] r : x.assignments){
			assignmentOuts.println(Arrays.toString(r));
		}
		for(String[] r : x.subnets){
			subnetOuts.println(Arrays.toString(r));
		}
		System.out.println("***"+x.assignments.size());
		
		assignmentOuts.close();
		subnetOuts.close();
	}

}