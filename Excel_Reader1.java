package excel_reader;
//need to import poi jar files first
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_Reader1 {
	public static void main(String[] args) throws Exception {
		ArrayList data=new ArrayList();
		//pointing to the excel file using the file path
		FileInputStream file= new FileInputStream("E://shams//shams.xlsx");
		XSSFWorkbook workbook=new XSSFWorkbook(file);
		XSSFSheet sheet=workbook.getSheetAt(0);
		
		Iterator rowitr=sheet.iterator();
		while (rowitr.hasNext()) {
		Row row=(Row) rowitr.next();
		
		Iterator cellitr=row.iterator();
		while (cellitr.hasNext()) {
			Cell cell=(Cell) cellitr.next();
			//can be used using  given if condition
			/*if (cell.getCellType()==CellType.STRING) {
				data.add(cell.getStringCellValue());
				
			}
			else if (cell.getCellType()==CellType.NUMERIC) {
				data.add(cell.getNumericCellValue());
			}*/
			switch (cell.getCellType()) {
			case STRING:
				data.add(cell.getStringCellValue());
				break;
			case NUMERIC:
				data.add(cell.getNumericCellValue());
                break;
			case BOOLEAN:
				data.add(cell.getBooleanCellValue());
				break;
			default:
				break;
			}
		}
		}
		System.out.println(data);
		/*for (int i = 0; i < data.size(); i++) {
			if (data.get(i).equals("nisar")) {
				System.out.println(data.get(i));
				System.out.println(data.get(i+1));*/
			}
			
			
		}
	
//}

