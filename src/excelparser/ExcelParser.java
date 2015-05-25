package excelparser;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelParser {
	Workbook workbook;
	int startRowIdx;
	int endRowIdx;
	String[] header = new String[]{"Vendor","Device Model","Serial Number","Server Standard","DC name","Cage","Rackrow","Rack","Position","System Name","Drac IP","Drac MAC","Phy NIC","Phy NIC2","Trend PO Number","Asset Number","Vendor PO Number","Tracking Number","Purchase Date","Arrival Date"};
	int[] columnIdx = new int[header.length];
	
	public  ExcelParser(File f) throws FileNotFoundException, IOException{
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
	public void valid() throws Exception {
		Sheet sheet = workbook.getSheetAt(0);
		startRowIdx = sheet.getFirstRowNum();
		endRowIdx = sheet.getLastRowNum();
	
		// check version
		java.util.Iterator<Row> rowIterator = sheet.iterator();
		Pattern p = Pattern.compile("Version:([\\d\\.]+)");
		while(rowIterator.hasNext()) {
			Row row = rowIterator.next();
			Cell c= row.getCell(0);
			if (c == null){
				continue;
			}
			Matcher m = p.matcher(c.getStringCellValue());
			if(m.find()){
				String x = m.group(1);
				if("0.2".equals(x)){
					startRowIdx = row.getRowNum();
					return;
				}else{
					throw new Exception("Not version 0.1 document!");
				}
			}
		}
		throw new Exception("No version found!");
	}	
	

	public void valid2() throws Exception {
		Sheet sheet = workbook.getSheetAt(0);
	
		// check head column
		java.util.Iterator<Row> rowIterator = sheet.iterator();
		while(rowIterator.hasNext()) {
			Row row = rowIterator.next();
			if( row.getRowNum() < startRowIdx){
				continue;
			}
			
			java.util.Iterator<Cell> cellIterator = row.cellIterator();
			int i = 0; 
			while(cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
//				if(cell.getCellType() == Cell.)
//				System.out.println(cell);
				if(Cell.CELL_TYPE_STRING == cell.getCellType() && header[i].equals(cell.getStringCellValue())){
					columnIdx[i] = cell.getColumnIndex();
					i++;
				}
			}
			if(i == header.length){
				startRowIdx = row.getRowNum()+1;
				return;
			}
		}
		throw new Exception("No version found!");
	}
	
	public List<Map<String, String>> parse(){
		List<Map<String, String>> result = new ArrayList<Map<String, String>>();
		Sheet sheet = workbook.getSheetAt(0);
		
		// check head column
		java.util.Iterator<Row> rowIterator = sheet.iterator();
		while(rowIterator.hasNext()) {
			Row row = rowIterator.next();
			if( row.getRowNum() < startRowIdx){
				continue;
			}
			if(row.getPhysicalNumberOfCells() == 0){
				continue;
			}
			
			Map<String, String> info = new HashMap<String, String>();
			
			for(int i = 0 ; i < columnIdx.length ; i++){
				String word = getCellString(row, columnIdx[i]);
				info.put(header[i], word);
			}
			result.add(info);
		}
		return result;
	}
	
	private String getCellString(Row row, int idx){
		Cell c = row.getCell(idx);
		if (c == null){
			return "";
		}
		if (c.getCellType() == Cell.CELL_TYPE_NUMERIC){
			return String.valueOf(c.getNumericCellValue());
		}else if (c.getCellType() == Cell.CELL_TYPE_BOOLEAN){
			return String.valueOf(c.getBooleanCellValue());
		}
		String tmp = c.getStringCellValue();
		if(tmp == null){
			return "";
		}
		return tmp;
	}
	
	public void close() throws IOException{
		workbook.close();
	}
	
	public static void main(String... args) throws FileNotFoundException, IOException, Exception{

		ExcelParser x = new ExcelParser(new File("C:\\Users\\joey_hsiao\\Documents\\Omnibus device import spreadsheet.xlsx"));
		x.valid();
		x.valid2();
		List<Map<String, String>> result = x.parse();
		System.err.println(result);
		System.err.println(result.size());
			
	}
}
