package quarterlypdi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import util.Functions;

public class QuarterlyPDI {
	private static final String filePath = "Z:/MIA 2017/Product and Safety Committee/PDI lists ex NZTA and VIRMs for light vehicles/PDI Lists 2017/";
	private static final String fileFolder = "4 December/";
	private static final String pdiFileName = "PDI report run on 2017 12 31 for MIA.xlsx";
	private static final String fileName = " PDI Report 31 December 2017.xlsx";
	
	private static HashMap<String, String> findDis = new HashMap<String, String>();
	private static HashMap<String, Integer> disMap = new HashMap<String, Integer>();
	
	
	private static String[] motorcycles = new String[]{"PEUGEOT", "BMW", "HONDA", "SUZUKI"};

	public static void main(String[] args) {
		try {
			FileInputStream pdiFile = new FileInputStream(new File(filePath+pdiFileName));
			generateDistributorMap();
			
			XSSFWorkbook wb = new XSSFWorkbook(pdiFile);
			XSSFSheet pdiSheet = wb.getSheetAt(0);
			
			int i = 1;
			int rowCount = pdiSheet.getPhysicalNumberOfRows();
			
			String newMake = pdiSheet.getRow(i).getCell(4).getStringCellValue();
			String oldMake = newMake;
			
			String disName;
			String disModel = pdiSheet.getRow(i).getCell(5).getStringCellValue();
			int j;
			ArrayList<XSSFWorkbook> disBooks = new ArrayList<XSSFWorkbook>();
			XSSFWorkbook xb = new XSSFWorkbook();
			XSSFSheet xs = xb.createSheet();
			CreationHelper ch = xb.getCreationHelper();
			CellStyle cs = xb.createCellStyle();
			
			while(i < rowCount) {
				disName = findDistributor(newMake, disModel);
				if (disMap.containsKey(disName)) {
					j = disMap.get(disName);
					for (XSSFWorkbook b: disBooks) {
						if (b.getSheetAt(0).getSheetName().equals(disName)) {
							xs = b.getSheetAt(0);
							cs = xs.getRow(1).getCell(0).getCellStyle();
						}
					}
				} else {
					disMap.put(disName, 1);
					j = 1;
					
					xb = new XSSFWorkbook();
					ch = xb.getCreationHelper();
					cs = xb.createCellStyle();
					cs.setDataFormat(ch.createDataFormat().getFormat("dd/mm/yyyy"));
					xs = xb.createSheet(disName);
					
					Row r = xs.createRow(0);
					for (int k = 0; k < 7; k++) {
						r.createCell(k).setCellValue(pdiSheet.getRow(0).getCell(k).getStringCellValue());
					}
					
					xs.setColumnWidth(0, 13*256);
					xs.setColumnWidth(1, 8*256);
					xs.setColumnWidth(2, 9*256);
					xs.setColumnWidth(3, 21*256);
					xs.setColumnWidth(4, 15*256);
					xs.setColumnWidth(5, 20*256);
					xs.setColumnWidth(6, 11*256);
					
					disBooks.add(xb);
				}
				while(i < rowCount) {
					newMake = pdiSheet.getRow(i).getCell(4).getStringCellValue();
					if (pdiSheet.getRow(i).getCell(5).getCellTypeEnum() == CellType.STRING)
						disModel = pdiSheet.getRow(i).getCell(5).getStringCellValue();
					else {
						double disModelNumeric = pdiSheet.getRow(i).getCell(5).getNumericCellValue();
						disModel = Double.toString(disModelNumeric);
					}
					if (!oldMake.equals(newMake)) {
						oldMake = newMake;
						break;
					}
					
					Row r = xs.createRow(j);
					Cell c = r.createCell(0);
					c.setCellValue(pdiSheet.getRow(i).getCell(0).getNumericCellValue());
					c.setCellStyle(cs);
					
					for (int k = 1; k < 7; k++) {
						if (pdiSheet.getRow(i).getCell(k)!=null) {
							if (pdiSheet.getRow(i).getCell(k).getCellTypeEnum()==CellType.STRING)
								r.createCell(k).setCellValue(pdiSheet.getRow(i).getCell(k).getStringCellValue());
							else
								r.createCell(k).setCellValue(pdiSheet.getRow(i).getCell(k).getNumericCellValue());
						}
					}
					j++;
					i++;
				}
				disMap.put(disName, j);
			}
			
			// Create the physical workbooks on disk
			for (XSSFWorkbook b: disBooks) {
				String name = b.getSheetAt(0).getSheetName();
				FileOutputStream out = new FileOutputStream(new File(filePath+fileFolder+name+fileName));
				b.write(out);
				out.close();
				b.close();
				System.out.println(name +" has written successfully");
			}
			
			wb.close();
			pdiFile.close();
			
			System.out.println("\n*** FINISHED ***");
			System.out.println("\nCheck Auto Distributors, BMW, Honda, and Suzuki for bikes/scooters");
		} catch (Exception e) { e.printStackTrace(); }

	}
	
	private static void generateDistributorMap() {
		findDis = Functions.getDistributorMap();
	}
	
	private static String findDistributor(String make, String model) {
		for (String bike : motorcycles) {
			if (make.equals(bike)) {
				make = Functions.determineIfBike(make, model);
				return make;
			}
		}
		if (findDis.containsKey(make)) {
			return findDis.get(make);
		}
		else return "Other";
	}
}
