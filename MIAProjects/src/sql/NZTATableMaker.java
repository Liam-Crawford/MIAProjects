package sql;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.Calendar;

import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import util.Functions;

public class NZTATableMaker {
	// Working folder and database constants
	private static final String filePath = "Z:/Liam working folder/NZTA Open Data/Test Tables/";
	private static final String fileNameAffix = " - Test.xlsx";
	private static final String connectionUrlMIA = "jdbc:sqlserver://192.168.1.202;databaseName=OpenData;user=MIA;password=Miaapp";
	private static final String connectionUrlSQLite = "jdbc:sqlite:E:/Projects/sqlite/tools/opendataraw.db";

	// Fields for storing which month/year we are using for our sql queries.
	private String month;
	private String ytd;
	private int monthNumber;
	private int year;

	private Connection con;
	private Functions func;

	// Local fields for helper methods
	private ResultSet rs;
	private XSSFWorkbook wb;
	private XSSFSheet s;
	private boolean outOfOrder;
	private int rowIndex;
	private Row row;

	public static void main(String[] args) {
		new NZTATableMaker();
	}
	
	public NZTATableMaker() {
		// Get the current date and extract the number (0-11) of the month.
		// The fact that Java uses 0-11 to represent months works for us as we want
		// to reference the month previous to this one (and in our database they will be 1-12).
		// The only thing to do is check if it's currently January (0) and set to December (12)
		// otherwise the sql queries won't return anything (no date in the db is 0).
		Calendar cal = Calendar.getInstance();
		monthNumber = cal.get(Calendar.MONTH);
		if (monthNumber==0) monthNumber = 12;

		// Hacks for if we run this from a different month than expected.
		monthNumber = 10;

		// We run this report for the previous month, so determine what that is.
		// Start by subtracting 1 to get last month (we do it this way so that if we are
		// in Jan it will to move to Dec of last year).
		cal.add(Calendar.MONTH, -1);
		// Get the month represented as a nice String to use in our sql queries.
		month = (new SimpleDateFormat("MMMM").format(cal.getTime())).toUpperCase();
		month = month + " " + new SimpleDateFormat("YYYY").format(cal.getTime());
		// Get the current year as a number.
		year = cal.get(Calendar.YEAR);
		
		// Create the YTD string which is a list of the months (as number) preceding the current one
		// separated by commas.
		ytd = "('";
		for (int i = 1; i <= monthNumber; i++) {
			if (i < monthNumber) ytd += Integer.toString(i)+"', '";
			else ytd += Integer.toString(i)+"')";
		}

		func = new Functions();

		// Connect to database and run queries
		try {
			con = connectToDB(connectionUrlSQLite);
			makeCountryTLA(Constants.PASS_CLASS_CLEAN, Constants.allImportStatus, Constants.T_001);
			/*makeImportStatus(Constants.passengerClasses, Constants.T_001A);
			makeCountryTLA(Constants.passengerClasses, Constants.newImportStatus, Constants.T_001N);
			makeCountryTLA(Constants.commercialClasses, Constants.allImportStatus, Constants.T_002);
			makeImportStatus(Constants.commercialClasses, Constants.T_002A);
			makeCountryTLA(Constants.commercialClasses, Constants.newImportStatus, Constants.T_002N);
			makeCountryCC(Constants.passengerClasses, Constants.allImportStatus, Constants.T_006);
			makeCountryCC(Constants.passengerClasses, Constants.newImportStatus, Constants.T_006N);
			makeCountryCC(Constants.passengerClasses, Constants.usedImportStatus, Constants.T_006X);
			makeCountryGVM(Constants.commercialClasses, Constants.allImportStatus, Constants.T_008);
			makeCountryGVM(Constants.commercialClasses, Constants.newImportStatus, Constants.T_008N);
			makeCountryGVM(Constants.commercialClasses, Constants.usedImportStatus, Constants.T_008X);
			makeModelCountryTLA(Constants.passengerClasses, Constants.newImportStatus, Constants.T_064N);
			makeModelCountryTLA(Constants.passengerClasses, Constants.usedImportStatus, Constants.T_064X);
			makeModelCountryTLA(Constants.commercialClasses, Constants.newImportStatus, Constants.T_065N);
			makeModelCountryTLA(Constants.commercialClasses, Constants.usedImportStatus, Constants.T_065X);
			makeModelFuelAge(Constants.commercialClasses, Constants.usedImportStatus, Constants.T_U8MM_AGE);
			makeImportYTD(Constants.motorcycleClasses, Constants.newImportStatus, Constants.T_MOTORCYCLES_NEW);
			makeImportYTD(Constants.motorcycleClasses, Constants.allImportStatus, Constants.T_Y_MPC_A);
			makeYTD(Constants.passengerClasses, Constants.newImportStatus, Constants.T_Y_001AN);
			makeYTD(Constants.passengerClasses, Constants.newImportStatus, Constants.T_Y_001AN_2AN);
			makeYTD(Constants.passengerClasses, Constants.newImportStatus, Constants.T_Y_001AX);
			makeYTD(Constants.passengerClasses, Constants.newImportStatus, Constants.T_Y_002AN);
			makeYTD(Constants.passengerClasses, Constants.newImportStatus, Constants.T_Y_002AX);
			makeManyYTD(Constants.passengerClasses, Constants.newImportStatus, Constants.T_Y_080N);
			makeManyYTD(Constants.commercialClasses, Constants.newImportStatus, Constants.T_Y_081N);
			makeModelYTD(Constants.passengerClasses, Constants.newImportStatus, Constants.T_Y_084N);
			makeModelYTD(Constants.commercialClasses, Constants.newImportStatus, Constants.T_Y_085N);
			makeModelSubCCYTD(Constants.motorcycleClasses, Constants.newImportStatus, Constants.T_Y_MPC50);
			makeModelSubCCYTD(Constants.motorcycleClasses, Constants.newImportStatus, Constants.T_Y_MPC51);
			makeModelTypeRental(Constants.allClasses, Constants.newImportStatus, Constants.T_YTD_RENTALS_NEW);
			makeModelYTD(Constants.passengerClasses, Constants.usedImportStatus, Constants.T_YTD_USED_CARS);
			makeModelYTD(Constants.commercialClasses, Constants.usedImportStatus, Constants.T_YTD_USED_COM);*/
			System.out.println("\nTask Complete.");
			con.close();
		} catch (SQLException e){
			e.printStackTrace();
		}
	}

	// ********** REGION CREATE TABLE METHODS **********
	
	private void makeCountryTLA(String segment, String importStatus, String fileName){
		int colOffset = 2;
		int rowOffset = 2;
		
		String sql = Constants.sqlMakeCountryTLA(segment, importStatus, year, monthNumber);
		
		try {
			initialiseWorkbook(
					sql,
					fileName,
					new String[]{Constants.getFileNameHeader(fileName, month)},
					new String[]{"MAKE", "COUNTRY OF ORIGIN"},
					Constants.regionCodeNames,
					colOffset,
					new String[]{"NZ TOTAL"}
					);

			// Algorithm to put the value of each vehicle into the correct TLA cell.
			rowIndex = rowOffset;
			outOfOrder = false;
			String[] oldVars = getEmptyStringArray(colOffset); // oldMake, oldCountry
			while (rs.next()) {
				// Pull variables from sql query
				String make = rs.getString(Constants.OD_MAKE);
				String country = func.getCountryCode(rs.getString(Constants.OD_ORIGINAL_COUNTRY));
				int cellLocation = func.getRegionCode(rs.getString(Constants.OD_TLA), colOffset);
				int total = Integer.parseInt(rs.getString(Constants.OD_TOTAL));

				oldVars = algoMakeCountry(new String[]{make, oldVars[0], country, oldVars[1]}, s, cellLocation, total);
			}
			setTotals(s, rowIndex, colOffset, rowOffset, Constants.regionCodeNames.length);
			writeFile(fileName, wb);
		} catch (Exception e) { e.printStackTrace(); }
	}

	private void makeImportStatus(String segment, String fileName){
		int colOffset = 1;
		int rowOffset = 2;
		
		// Create sql query
		String sql = Constants.sqlMakeImportStatus(segment, year, monthNumber);

		try {
			initialiseWorkbook(sql,
					fileName,
					new String[]{Constants.getFileNameHeader(fileName, month)},
					new String[]{"MAKE", "NEW", "USED", "TOTAL"},
					new String[]{},
					colOffset,
					new String[]{}
					);

			rowIndex = rowOffset;
			String oldMake = "";
			while (rs.next()) {
				// Pull values for each column from query
				String make = rs.getString(Constants.OD_MAKE);
				String importStatus = rs.getString(Constants.OD_IMPORT_STATUS);
				int total = Integer.parseInt(rs.getString(Constants.OD_TOTAL));
				
				if (make.equals(oldMake)) {
					row.createCell(2).setCellValue(total);
				} else {
					oldMake = make;
					row = s.createRow(rowIndex++);
					row.createCell(0).setCellValue(make);
					
					// Determine whether to use NEW or USED column
					Cell c;
					c = (importStatus.equals("NEW")) ? row.createCell(1) : row.createCell(2);
					c.setCellValue(total);
				}
			}
			setTotals(s, rowIndex, colOffset, rowOffset, 2);
			writeFile(fileName, wb);
		} catch (Exception e) { e.printStackTrace(); }
	}
	
	private void makeCountryCC(String segment, String importStatus, String fileName) {
		int colOffset = 2;
		int rowOffset = 2;
		
		String sql = Constants.sqlMakeCountryCC(segment, importStatus, year, monthNumber);

		try {
			initialiseWorkbook(sql,
					fileName,
					new String[]{Constants.getFileNameHeader(fileName, month)},
					new String[]{"MAKE", "COUNTRY OF ORIGIN"},
					Constants.ccBrackets,
					colOffset,
					new String[]{"NZ TOTAL"}
			);

			// Algorithm to put the retail number of each vehicle into the correct cc_rating cell.
			rowIndex = rowOffset;
			outOfOrder = false;
			String[] oldVars = getEmptyStringArray(colOffset);
			while (rs.next()) {
				// Pull variables from sql query
				String make = rs.getString(Constants.OD_MAKE);
				String country = func.getCountryCode(rs.getString(Constants.OD_ORIGINAL_COUNTRY));
				int cellLocation = func.getCCBracketCell(Integer.parseInt(rs.getString(Constants.OD_CC_RATING)), colOffset);
				int total = Integer.parseInt(rs.getString(Constants.OD_TOTAL));

				oldVars = algoMakeCountry(new String[]{make, oldVars[0], country, oldVars[1]}, s, cellLocation, total);
			}
			setTotals(s, rowIndex, colOffset, rowOffset, Constants.ccBrackets.length);
			writeFile(fileName, wb);
		} catch (Exception e) { e.printStackTrace(); }
	}
	
	private void makeCountryGVM(String segment, String importStatus, String fileName) {
		int colOffset = 2;
		int rowOffset = 2;
		
		String sql = Constants.sqlMakeCountryGVM(segment, importStatus, year, monthNumber);

		try {
			initialiseWorkbook(sql,
					fileName,
					new String[]{Constants.getFileNameHeader(fileName, month)},
					new String[]{"MAKE", "COUNTRY OF ORIGIN"},
					Constants.gvmBrackets,
					colOffset,
					new String[]{"TOTAL"}
			);

			// Algorithm to put the sales number of each vehicle into the correct cc_rating cell.
			rowIndex = rowOffset;
			outOfOrder = false;
			String[] oldVars = getEmptyStringArray(colOffset);
			while (rs.next()) {
				// Pull variables from sql query
				String make = rs.getString(Constants.OD_MAKE);
				String country = func.getCountryCode(rs.getString(Constants.OD_ORIGINAL_COUNTRY));
				String vClass = rs.getString(Constants.OD_CLASS);
				int cellLocation;
				int gvm = Integer.parseInt(rs.getString(Constants.OD_GROSS_VEHICLE_MASS));
				// If the vehicle is a bus, the figures go in different cells than the rest
				if (vClass.equals("XMD1")||vClass.equals("XMD2")||vClass.equals("XMD3")||vClass.equals("XMD4")||vClass.equals("XME")) {
					cellLocation = func.getGVMBracketBusCell(gvm, colOffset);
				} else cellLocation = func.getGVMBracketCell(gvm, colOffset);
				int total = Integer.parseInt(rs.getString(Constants.OD_TOTAL));
				
				oldVars = algoMakeCountry(new String[]{make, oldVars[0], country, oldVars[1]}, s, cellLocation, total);
			}
			setTotals(s, rowIndex, colOffset, rowOffset, Constants.gvmBrackets.length);
			writeFile(fileName, wb);
		} catch (Exception e) { e.printStackTrace(); }
	}

	// NEEDS WORK
	private void makeModelCountryTLA(String segment, String importStatus, String fileName){
		int colOffset = 3;
		int rowOffset = 2;
		
		String sql = Constants.sqlMakeModelCountryTLA(segment, importStatus, year, monthNumber);
		
		try {
			initialiseWorkbook(sql,
					fileName,
					new String[]{Constants.getFileNameHeader(fileName, month)},
					new String[]{"MAKE", "MODEL", "COUNTRY OF ORIGIN"},
					Constants.regionCodeNames,
					colOffset,
					new String[]{"TOTAL"}
			);

			// Algorithm to put the value of each vehicle into the correct TLA cell.
			rowIndex = rowOffset;
			String oldMake = "";
			String oldCountry = "";
			String oldModel = "";
			while (rs.next()) {
				// Pull variables from sql query
				String make = rs.getString(Constants.OD_MAKE);
				String model = rs.getString(Constants.OD_MODEL);
				String country = func.getCountryCode(rs.getString(Constants.OD_ORIGINAL_COUNTRY));
				int cellLocation = func.getRegionCode(rs.getString(Constants.OD_TLA), colOffset);
				int total = Integer.parseInt(rs.getString(Constants.OD_TOTAL));
				
				if (make.equals(oldMake)) {
					if (model.equals(oldModel)) {
						if (country.equals(oldCountry)) {
							// if we are still on the same make, model, and country, use the same row
							calcTotal(row, cellLocation, total);
						} else {
							// if we are still on the same make and model but new country, make a new row
							oldCountry = country;
							createNewRow(new String[]{make, model, country}, cellLocation, total);
						}
					} else {
						// if we are onto a new model, create a new row and remember the make, model, and country
						oldModel = model;
						oldCountry = country;
						createNewRow(new String[]{make, model, country}, cellLocation, total);
					}
					
				} else {
					// if we are onto a new make, create a new row and remember the make, model, and country
					oldMake = make;
					oldModel = model;
					oldCountry = country;
					createNewRow(new String[]{make, model, country}, cellLocation, total);
				}
			}
			setTotals(s, rowIndex, colOffset, rowOffset, Constants.regionCodeNames.length);
			writeFile(fileName, wb);
		} catch (Exception e) { e.printStackTrace(); }
	}

	// NEEDS WORK
	private void makeModelFuelAge(String segment, String importStatus, String fileName) {
		int colOffset = 3;
		int rowOffset = 2;
		
		String sql = Constants.sqlMakeModelFuelAge(segment, importStatus, year, monthNumber);
		
		try {
			initialiseWorkbook(sql,
					fileName,
					new String[]{Constants.getFileNameHeader(fileName, month)},
					new String[]{"VEHICLE MAKE", "VEHICLE MODEL", "FUEL TYPE"},
					Constants.ages,
					colOffset,
					new String[]{"TOTAL", "MEAN AGE"}
			);

			// Algorithm to put the value of each vehicle into the correct age cell.
			rowIndex = rowOffset;
			String oldMake = "";
			String oldModel = "";
			String oldFuel = "1";
			while (rs.next()) {
				// Pull variables from sql query
				String make = rs.getString(Constants.OD_MAKE);
				String model = rs.getString(Constants.OD_MODEL);
				String fuel = rs.getString(Constants.OD_MOTIVE_POWER);
				int cellLocation = func.getAgeCell(year-Integer.parseInt(rs.getString(Constants.OD_VEHICLE_YEAR)))+colOffset;
				int total = Integer.parseInt(rs.getString(Constants.OD_TOTAL));
				
				if (make.equals(oldMake)) {
					if (model.equals(oldModel)) {
						if (fuel.equals(oldFuel)) {
							// if we are still on the same make, model, and fuel type, use the same row
							calcTotal(row, cellLocation, total);
						} else {
							// if we are still on the same make and model but new fuel type, make a new row
							oldFuel = fuel;
							createNewRow(new String[]{make, model, fuel}, cellLocation, total);
						}
					} else {
						// if we are onto a new model, create a new row and remember the make, model, and fuel type
						oldModel = model;
						oldFuel = fuel;
						createNewRow(new String[]{make, model, fuel}, cellLocation, total);
					}
					
				} else {
					// if we are onto a new make, create a new row and remember the make, model, and fuel type
					oldMake = make;
					oldModel = model;
					oldFuel = fuel;
					createNewRow(new String[]{make, model, fuel}, cellLocation, total);
				}
			}
			setTotals(s, rowIndex, colOffset, rowOffset, Constants.ages.length);
			writeFile(fileName, wb);
		} catch (Exception e) { e.printStackTrace(); }
	}
	
	private void makeYTD(String segment, String importStatus, String fileName) {
		int colOffset = 1;
		int rowOffset = 2;
		
		String sql = Constants.sqlMakeYTD(segment, importStatus, year, ytd);
		
		try {
			initialiseWorkbook(sql,
					fileName,
					new String[]{Integer.toString(year), Constants.getFileNameHeader(fileName, month)},
					new String[]{"MAKE"},
					Constants.months,
					colOffset,
					new String[]{"YTD"}
			);

			// Algorithm to put the value of each vehicle into the correct cell.
			rowIndex = rowOffset;
			String oldMake = "";
			while (rs.next()) {
				// Pull variables from sql query
				String make = rs.getString(Constants.OD_MAKE);
				int cellLocation = Integer.parseInt(rs.getString(Constants.OD_FIRST_NZ_REGISTRATION_MONTH));
				int total = Integer.parseInt(rs.getString(Constants.OD_TOTAL));
				
				if (make.equals(oldMake)) {
					row.createCell(cellLocation).setCellValue(total);
				} else {
					oldMake = make;
					row = s.createRow(rowIndex++);
					row.createCell(0).setCellValue(make);
					row.createCell(cellLocation).setCellValue(total);
				}
			}
			setTotals(s, rowIndex, colOffset, rowOffset, Constants.months.length);
			writeFile(fileName, wb);
		} catch (Exception e) { e.printStackTrace(); }
	}

	private void makeManyYTD(String segment, String importStatus, String fileName) {
		int colOffset = 10;
		int rowOffset = 1;
		
		String sql = Constants.sqlMakeManyYTD(segment, importStatus, year, ytd);
		
		try {
			initialiseWorkbook(sql,
					fileName,
					getHeaders(new String[]{
							"MAKE",
							"MODEL",
							"SUBMODEL",
							"COUNTRY",
							"ASSEMBLY",
							"CC",
							"BODY",
							"FUEL",
							"AXELCODE",
							"KW"
					}, Constants.months, "YTD"),
					new String[]{},
					new String[]{},
					colOffset,
					new String[]{}
			);

			// Algorithm to put the value of each vehicle into the correct cell.
			rowIndex = rowOffset;
			String[] columns = new String[colOffset];
			String[] oldColumns = getEmptyStringArray(colOffset);
			while (rs.next()) {
				// Pull variables from sql query
				columns[0] = rs.getString(Constants.OD_MAKE);
				columns[1] = rs.getString(Constants.OD_MODEL);
				columns[2] = rs.getString(Constants.OD_SUBMODEL);
				columns[3] = rs.getString(Constants.OD_ORIGINAL_COUNTRY);
				columns[4] = rs.getString(Constants.OD_NZ_ASSEMBLED);
				columns[5] = rs.getString(Constants.OD_CC_RATING);
				columns[6] = rs.getString(Constants.OD_BODY_TYPE);
				columns[7] = rs.getString(Constants.OD_MOTIVE_POWER);
				columns[8] = rs.getString(Constants.OD_NUMBER_OF_AXLES);
				columns[9] = rs.getString(Constants.OD_POWER_RATING);
				int cellLocation = Integer.parseInt(rs.getString(Constants.OD_FIRST_NZ_REGISTRATION_MONTH))+colOffset-1;
				int total = Integer.parseInt(rs.getString(Constants.OD_TOTAL));
				
				algoNewLine(columns, oldColumns, s, cellLocation, total);
			}
			setTotals(s, rowIndex, colOffset, rowOffset, Constants.months.length);
			writeFile(fileName, wb);
		} catch (Exception e) { e.printStackTrace(); }
	}
	
	private void makeModelYTD(String segment, String importStatus, String fileName) {
		int colOffset = 2;
		int rowOffset = 1;
		
		String sql = Constants.sqlMakeModelYTD(segment, importStatus, year, ytd);
		
		try {
			initialiseWorkbook(sql,
					fileName,
					getHeaders(new String[]{
							"MAKE",
							"MODEL"
					}, Constants.months, "YTD"),
					new String[]{},
					new String[]{},
					colOffset,
					new String[]{}
			);

			// Algorithm to put the value of each vehicle into the correct cell.
			rowIndex = rowOffset;
			String[] columns = new String[colOffset];
			String[] oldColumns = getEmptyStringArray(colOffset);
			while (rs.next()) {
				// Pull variables from sql query
				columns[0] = rs.getString(Constants.OD_MAKE);
				columns[1] = rs.getString(Constants.OD_MODEL);
				int cellLocation = Integer.parseInt(rs.getString(Constants.OD_FIRST_NZ_REGISTRATION_MONTH))+colOffset-1;
				int total = Integer.parseInt(rs.getString(Constants.OD_TOTAL));
				
				algoNewLine(columns, oldColumns, s, cellLocation, total);
			}
			setTotals(s, rowIndex, colOffset, rowOffset, Constants.months.length);
			writeFile(fileName, wb);
		} catch (Exception e) { e.printStackTrace(); }
	}
	
	private void makeImportYTD(String segment, String importStatus, String fileName) {
		int colOffset = 2;
		int rowOffset = 2;
		
		String sql = Constants.sqlMakeImportYTD(segment, importStatus, year, ytd);
		
		try {
			initialiseWorkbook(sql,
					fileName,
					new String[]{Integer.toString(year), Constants.getFileNameHeader(fileName, month)},
					new String[]{"MAKE", "NEW-USED"},
					Constants.months,
					colOffset,
					new String[]{"YTD"}
			);

			// Algorithm to put the value of each vehicle into the correct cell.
			rowIndex = rowOffset;
			String[] columns = new String[colOffset];
			String[] oldColumns = getEmptyStringArray(2);
			while (rs.next()) {
				// Pull variables from sql query
				columns[0] = rs.getString(Constants.OD_MAKE);
				columns[1] = rs.getString(Constants.OD_IMPORT_STATUS);
				int cellLocation = Integer.parseInt(rs.getString(Constants.OD_FIRST_NZ_REGISTRATION_MONTH))+colOffset-1;
				int total = Integer.parseInt(rs.getString(Constants.OD_TOTAL));
				
				algoNewLine(columns, oldColumns, s, cellLocation, total);
			}
			setTotals(s, rowIndex, colOffset, rowOffset, Constants.months.length);
			writeFile(fileName, wb);
		} catch (Exception e) { e.printStackTrace(); }
	}
	
	private void makeModelSubCCYTD(String segment, String importStatus, String fileName) {
		int colOffset = 3;
		int rowOffset = 2;
		
		try {
			String[] headers = Constants.getFileNameHeaderMotorcycle(fileName);

			String cc = headers[1];
			String sql = Constants.sqlMakeModelSubCCYTD(segment, importStatus, year, ytd, cc);

			initialiseWorkbook(sql,
					fileName,
					new String[]{Integer.toString(year), headers[0]},
					new String[]{"MAKE", "MODEL", "SUBMOD"},
					Constants.months,
					colOffset,
					new String[]{"YTD"}
			);

			// Algorithm to put the value of each vehicle into the correct cell.
			rowIndex = rowOffset;
			String[] columns = new String[colOffset];
			String[] oldColumns = getEmptyStringArray(colOffset);
			while (rs.next()) {
				// Pull variables from sql query
				columns[0] = rs.getString(Constants.OD_MAKE);
				columns[1] = rs.getString(Constants.OD_MODEL);
				columns[2] = rs.getString(Constants.OD_SUBMODEL);
				int cellLocation = Integer.parseInt(rs.getString(Constants.OD_FIRST_NZ_REGISTRATION_MONTH))+colOffset-1;
				int total = Integer.parseInt(rs.getString(Constants.OD_TOTAL));
				
				algoNewLine(columns, oldColumns, s, cellLocation, total);
			}
			setTotals(s, rowIndex, colOffset, rowOffset, Constants.months.length);
			writeFile(fileName, wb);
		} catch (Exception e) { e.printStackTrace(); }
	}
	
	private void makeModelTypeRental(String segment, String importStatus, String fileName) {
		int colOffset = 3;
		int rowOffset = 1;
		
		String sql = Constants.sqlMakeModelTypeRental(segment, importStatus, year, ytd);
		
		try {
			initialiseWorkbook(sql,
					fileName,
					getHeaders(new String[]{
							"MAKE",
							"MODEL",
							"VEHICLE TYPE"
					}, Constants.months, "YTD"),
					new String[]{},
					new String[]{},
					colOffset,
					new String[]{}
			);

			// Algorithm to put the value of each vehicle into the correct cell.
			rowIndex = rowOffset;
			String[] columns = new String[colOffset];
			String[] oldColumns = getEmptyStringArray(colOffset);
			while (rs.next()) {
				// Pull variables from sql query
				columns[0] = rs.getString(Constants.OD_MAKE);
				columns[1] = rs.getString(Constants.OD_MODEL);
				columns[2] = rs.getString(Constants.OD_VEHICLE_TYPE);
				int cellLocation = Integer.parseInt(rs.getString(Constants.OD_FIRST_NZ_REGISTRATION_MONTH))+colOffset-1;
				int total = Integer.parseInt(rs.getString(Constants.OD_TOTAL));
				
				algoNewLine(columns, oldColumns, s, cellLocation, total);
			}
			setTotals(s, rowIndex, colOffset, rowOffset, Constants.months.length);
			writeFile(fileName, wb);
		} catch (Exception e) { e.printStackTrace(); }
	}

	// ********** END REGION CREATE TABLE METHODS **********

	// ********** REGION HELPER METHODS **********

	/**
	 * A helper method to setup some initial variables common to all the table making methods.
	 * Runs the sql query, creates the excel book and sheet, sets up various headings.
	 *
	 * @param sql The SQL query in String format
	 * @param fileName The name of the file used to create one of the headings and the name of the file on disk
	 * @param headings1 First set of headings on line 1
	 * @param headings2 Second set of headings on line 2
	 * @param headings3 Third set of headings also on line 2 straight after headings2
	 * @param colOffset How many columns are String labels like Make/Model etc.
	 * @param finalHeadings Last column heading, usually some kind of total.
	 * @throws SQLException This method will always be called from within a try/catch
	 */
	private void initialiseWorkbook(String sql, String fileName, String[] headings1, String[] headings2,
									String[] headings3, int colOffset, String[] finalHeadings) throws SQLException{
		// Run the sql query
		rs = readDB(con, sql);

		// Create the workbook and sheet
		wb = new XSSFWorkbook();
		s = wb.createSheet(fileName);

		// Setup the initial headers
		row = s.createRow(0);
		for (int i = 0; i < headings1.length; i++) row.createCell(i).setCellValue(headings1[i]);
		row = s.createRow(1);
		for (int i = 0; i < headings2.length; i++) row.createCell(i).setCellValue(headings2[i]);

		// Setup the rest of the headings and the final cell heading (usually some kind of total)
		int col = setupHeadings(headings3, colOffset);
		for (int i = col; i < col+finalHeadings.length; i++) row.createCell(i).setCellValue(finalHeadings[i-col]);
	}

	/**
	 * Create the database connection
	 */
	private Connection connectToDB(String url) throws SQLException {
		return DriverManager.getConnection(url);
	}

	/**
	 * Processes an sql statement with a given connection and returns the ResultSet
	 */
	private ResultSet readDB(Connection con, String sql) throws SQLException{
		Statement stmt = con.createStatement();
		return stmt.executeQuery(sql);
	}

	/**
	 * Takes a Row and sets it's cell headings for the associated String array, starting at the
	 * column index. Each file will have the first few columns already set to certain things like Make, Model etc.
	 * Then there will be the section which the string array fills out, could be months of the year, TLAs etc.
	 * The column index is returned to allow the calling method to add columns after the ones this method modifies.
	 *
	 * @param headings The array to pull headings from.
	 * @param columnIndex The column to start at.
	 * @return the column index that the array ended at + 1.
	 */
	private int setupHeadings(String[] headings, int columnIndex) {
		for (String s : headings) row.createCell(columnIndex++).setCellValue(s);
		return columnIndex;
	}

	/**
	 * Some tables only have 1 row of headings but initialiseWorkbook is setup to handle
	 * 2 rows. We can hack around this by putting what would normally be on row 2 into row 1
	 * and just passing empty arrays for row 2, which will then be overwritten
	 *
	 * @param initial The hardcoded column headings
	 * @param extras The column headings from one of Constants arrays
	 * @param last The final column heading usually a total or ytd
	 * @return the String[] which will be passed to initialiseWorkbook.
	 */
	private String[] getHeaders(String[] initial, String[] extras, String last) {
		// Setup the array based on the total size of incoming parameters
		String[] headers = new String[initial.length+extras.length+1];
		int i, j;
		// 2 loops to fill out our array with the 2 arrays passed in
		for (i = 0; i < initial.length; i++) headers[i] = initial[i];
		for (j = i; j < extras.length+i; j++) headers[j] = extras[j-i];
		// final index of the array is the string Last.
		headers[j] = last;

		return headers;
	}

	/**
	 * Algorithm to fill out a sheet with sales data in make/country order.
	 *
	 * @param vars The make and country vars
	 * @param s XSSFSheet to modify.
	 * @param cellLocation The cell index which is getting modified
	 * @param total Part of the value that makes up the sales.
	 * @return an array containing the updated oldMake and oldCountry.
	 */
	private String[] algoMakeCountry(String[] vars, XSSFSheet s, int cellLocation, int total) {
		String make = vars[0];
		String oldMake = vars[1];
		String country = vars[2];
		String oldCountry = vars[3];

		if (make.equals(oldMake)) {
			if (country.equals(oldCountry)) {
				// if we are still on the same make and country, use the same row
				// update the sales amount
				calcTotal(row, cellLocation, total);
			} else {
				// if we are still on the same make but new country, make a new row with same make

				// If the countries are out of order, fix it.
				if (outOfOrder) {
					outOfOrder = reorderCountry(s, rowIndex);
					// After reordering, the oldCountry will need to point to the new row above the current one.
					oldCountry = s.getRow(rowIndex-1).getCell(1).getStringCellValue();
				}
				// Check if the countries are out of order.
				if (oldCountry.compareToIgnoreCase(country)>0) outOfOrder = true;

				oldCountry = country;
				createNewRow(new String[]{make, country}, cellLocation, total);
			}
		} else {
			// if we are onto a new make, make a new row and remember the make and country

			// if the countries are out of order, fix it.
			//System.out.println(rowIndex);
			if (outOfOrder) outOfOrder = reorderCountry(s, rowIndex);

			oldMake = make;
			oldCountry = country;
			createNewRow(new String[]{make, country}, cellLocation, total);
		}

		return new String[]{oldMake, oldCountry};
	}

	/**
	 * Algorithm that takes an array of values and compares it against an old array of values.
	 * If any of the values have changed we make a new row in the sheet.
	 *
	 * @param columns Array of new values
	 * @param oldColumns Array of old values
	 * @param s XSSFSheet to modify
	 * @param cellLocation Cell index to modify
	 * @param total Value to put in Cell index.
	 */
	private void algoNewLine(String[] columns, String[] oldColumns, XSSFSheet s, int cellLocation, int total) {
		boolean newLine = false;
		for (int i = 0; i < columns.length; i++) {
			// Determine if any values have changed
			if (!columns[i].equals(oldColumns[i])) newLine = true;
		}

		if (!newLine) {
			// If no values changed, stay on this line and fill in the next month
			row.createCell(cellLocation).setCellValue(total);
		} else {
			// If values changed, we make a new line
			row = s.createRow(rowIndex++);
			for (int i = 0; i < columns.length; i++) {
				row.createCell(i).setCellValue(columns[i]);
				oldColumns[i] = columns[i];
			}
			row.createCell(cellLocation).setCellValue(total);
		}
	}

	private void createNewRow(String[] vars, int cellLocation, int total) {
		row = s.createRow(rowIndex++);
		for (int i = 0; i < vars.length; i++) row.createCell(i).setCellValue(vars[i]);
		row.createCell(cellLocation).setCellValue(total);
	}

	// Adds to the figure already in the cell.
	private void calcTotal(Row r, int cellLocation, int total) {
		if (r.getCell(cellLocation)!=null) {
			int trueTotal = (int)r.getCell(cellLocation).getNumericCellValue() + total;
			r.getCell(cellLocation).setCellValue(trueTotal);
		} else r.createCell(cellLocation).setCellValue(total);
	}

	/**
	 * A helper method to reorder alphabetically within Make, by Country. NZTA tables already have the 3 letter
	 * country codes setup, whereas with opendata we receive them as full country names. This means countries like
	 * United Kingdom will come out of our SQL query ordered with 'U' but NZTA shortens this to 'GBR'.
	 *
	 * @param s The sheet to  modify
	 * @param row The index of the current row
	 */
	private boolean reorderCountry(XSSFSheet s, int row) {
		// Figure out how many rows up the current row needs to go to be in alphabetical order.
		int i = row-2;
		String above = s.getRow(i).getCell(1).getStringCellValue();
		String current = s.getRow(row-1).getCell(1).getStringCellValue();

		while (above.compareToIgnoreCase(current)>0) {
			i--;
			//System.out.println(i);
			if (s.getRow(i)==null) break;
			above = s.getRow(i).getCell(1).getStringCellValue();
		}
		// Move rows down, creating a gap for the current row to move to.
		s.shiftRows(++i, row, 1);

		// Copy current row to new destination, with default copy policy.
		s.copyRows(row, row, i, new CellCopyPolicy());

		return false;
	}

	/**
	 * Helper method that takes an XSSFSheet and applies totals to the rows and columns.
	 * @param s XSSFSheet
	 * @param lastRow The last row in the sheet
	 * @param colOffset The first few columns that don't have numbers
	 * @param rowOffset The first few rows that don't have numbers.
	 * @param length The amount of columns to sum.
	 */
	private void setTotals(XSSFSheet s, int lastRow, int colOffset, int rowOffset, int length) {
		// Setup totals for each row
		for (int j = rowOffset; j < lastRow; j++) {
			String row = Integer.toString(j+1);
			String first = CellReference.convertNumToColString(colOffset);
			String last = CellReference.convertNumToColString(length+colOffset-1);
			s.getRow(j).createCell(length+colOffset).setCellFormula(String.format("SUM(%s%s:%s%s)",first,row,last,row));
		}

		// Setup totals row at the bottom
		Row r = s.createRow(lastRow);
		r.createCell(0).setCellValue("TOTALS");
		for (int j = colOffset; j < length+colOffset+1; j++) {
			String column = CellReference.convertNumToColString(j);
			r.createCell(j).setCellFormula(String.format("SUM(%s%d:%s%d)",column,rowOffset+1,column,lastRow));
		}
	}

	/**
	 * Returns an empty string array of size 'length'. Useful for generating arrays to store temp
	 * variables for comparing to later.
	 * @param length the length of the array
	 * @return an empty array of size length
	 */
	private String[] getEmptyStringArray(int length) {
		String[] array = new String[length];
		for (int i = 0; i < length; i++) array[i] = "";
		return array;
	}

	// Writes the excel file to disk using our filepath set at the top of the class.
	private void writeFile(String fileName, XSSFWorkbook wb) throws IOException{
		FileOutputStream out = new FileOutputStream(new File(filePath + fileName + fileNameAffix));
		wb.write(out);
		out.close();
		wb.close();

		System.out.println("File "+fileName+" written successfully");
	}

	// ********** END REGION HELPER METHODS **********

	/**
	 * Prints the contents of the sql query to the console (useful for debugging)
	 *
	 * @param sql
	 */
	@SuppressWarnings("unused")
	private void printSQLToConsole(String sql) {
		try {
			ResultSet rs = readDB(con, sql);
			// Get the metadata so we can iterate through all columns
			ResultSetMetaData rsmd = rs.getMetaData();
			// As long as there are more rows
			while (rs.next()){
				// Iterate through each column and print the value, note the first starts at 1 not 0.
				for (int i = 1; i <= rsmd.getColumnCount(); i++){
					System.out.print(rs.getString(i)+" ");
				}
				// Drop down a line for the next row
				System.out.print("\n");
			}
		} catch (SQLException e) { e.printStackTrace(); }
	}

}
