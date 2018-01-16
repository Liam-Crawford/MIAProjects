package util;

import java.util.HashMap;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;

public class Functions {
	private HashMap<String, Integer> regionCodeMap = new HashMap<String, Integer>();

	public Functions() {
		// North Island
		regionCodeMap.put("WHA", 0);  // Reg - Northland (Whangarei)
		regionCodeMap.put("AUC", 1);  // Reg - Auckland
		regionCodeMap.put("HAM", 2);  // Reg - Waikato + Taupo
		regionCodeMap.put("THA", 3);  // TLA - Thames-Coromandel + Hauraki
		regionCodeMap.put("TAU", 4);  // Reg - Bay of Plenty (Tauranga)
		regionCodeMap.put("ROT", 5);  // TLA - Rotorua 
		regionCodeMap.put("GIS", 6);  // Reg - Gisborne
		regionCodeMap.put("NAP", 7);  // Reg - Hawke's Bay (Napier)
		regionCodeMap.put("NEW", 8);  // Reg - Taranaki (New Plymouth)
		regionCodeMap.put("WAN", 9); // TLA - Whanganui + Ruapehu
		regionCodeMap.put("PAL", 10); // Reg?- Palmerston North City + Horowhenua + Tararua + Rangitikei + Manawatu
		regionCodeMap.put("MAS", 11); // TLA - Masterton + Carterton
		regionCodeMap.put("WEL", 12); // Reg - Wellington
		
		// South Island
		regionCodeMap.put("NEL", 13); // Reg - Nelson + Tasman
		regionCodeMap.put("BLE", 14); // Reg - Marlborough (Blenheim) 
		regionCodeMap.put("GRE", 15); // TLA - Grey + Buller
		regionCodeMap.put("WES", 16); // Reg - West Coast (Westland)
		regionCodeMap.put("CHR", 17); // Reg - Canterbury (Christchurch)
		regionCodeMap.put("TIM", 18); // TLA - Timaru
		regionCodeMap.put("OAM", 19); // TLA - Waitaki (Oamaru)
		regionCodeMap.put("DUN", 20); // Reg - Otago (Dunedin)
		regionCodeMap.put("INV", 21); // Reg - Southland (Invercargill)
		regionCodeMap.put("OTH", 22); // Anything without a TLA
	}
	
	public void format0012A(HSSFWorkbook wb) {
		HSSFSheet s = wb.getSheetAt(0);
		for (int j = 0; j < 4; j++) {
			s.setColumnWidth(j, 16*256);
		}
		
		setBold(wb);
		
		s.addMergedRegion(new CellRangeAddress(0,0,0,3));
	}
	
	public void format0012(HSSFWorkbook wb) {
		setBold(wb);
		wb.getSheetAt(0).addMergedRegion(new CellRangeAddress(0,0,0,70));
	}
	
	private void setBold(HSSFWorkbook wb) {
		HSSFSheet s = wb.getSheetAt(0);
		HSSFFont f = wb.createFont();
		f.setBold(true);
		HSSFCellStyle style = wb.createCellStyle();
		style.setFont(f);
		s.getRow(0).getCell(0).setCellStyle(style);
	}
	
	/**
	 * Takes the country and returns a 3 character code
	 * @param country
	 * @return
	 */
	public String getCountryCode(String country) {
		switch(country) {
		case "ARGENTINA":return "ARG";
		case "AUSTRALIA":return "AUS";
		case "AUSTRIA":return "AUT";
		case "BELGIUM":return "BEL";
		case "BRAZIL":return "BRA";
		case "CANADA":return "CAN";
		case "CHINA":return "CHN";
		case "CZECH REPUBLIC":return "CZE";
		case "DENMARK":return "DEN";
		case "FRANCE":return "FRA";
		case "GERMANY":return "GER";
		case "GREECE":return "GRE";
		case "HONG KONG":return "HKG";
		case "HUNGARY":return "HUN";
		case "IMPORTED BUILT-UP":return "IMP";
		case "INDIA":return "IND";
		case "INDONESIA":return "INA";
		case "ITALY":return "ITA";
		case "JAPAN":return "JPN";
		case "MALAYSIA":return "MAL";
		case "MEXICO":return "MEX";
		case "NETHERLANDS":return "NLD";
		case "NEW ZEALAND":return "NZL";
		case "NORWAY":return "NOR";
		case "NOT KNOWN":return "XXX";
		case "OTHER":return "OTH";
		case "PHILIPPINES":return "PHI";
		case "POLAND":return "POL";
		case "SINGAPORE":return "SGP";
		case "SLOVAKIA":return "SVK";
		case "SOUTH AFRICA":return "SAF";
		case "SOUTH KOREA":return "KOR";
		case "SPAIN":return "ESP";
		case "SWEDEN":return "SWE";
		case "SWITZERLAND":return "SWZ";
		case "TAIWAN":return "TWN";
		case "THAILAND":return "THA";
		case "TURKEY":return "TUR";
		case "UNITED KINGDOM":return "GBR";
		case "UNITED STATES":return "USA";
		case "UNKNOWN":return "XXX";
		case "USSR - RUSSIA":return "RUS";
		case "YUGOSLAVIA":return "YUG";
		}
		return "???";
	}
	
	/**
	 * Takes the TLA and returns the index of the cell the figure will go into
	 * @param tla
	 * @return
	 */
	public int getRegionCode(String tla, int offset){
		switch(tla) {
		case "ASHBURTON DISTRICT":return regionCodeMap.get("CHR")+offset;
		case "AUCKLAND":return regionCodeMap.get("AUC")+offset;
		case "BULLER DISTRICT":return regionCodeMap.get("GRE")+offset;
		case "CARTERTON DISTRICT":return regionCodeMap.get("MAS")+offset;
		case "CENTRAL HAWKE'S BAY DISTRICT":return regionCodeMap.get("NAP")+offset;
		case "CENTRAL OTAGO DISTRICT":return regionCodeMap.get("DUN")+offset;
		case "CHATHAM ISLANDS TERRITORY":return regionCodeMap.size()+3+offset;				// ***
		case "CHRISTCHURCH CITY":return regionCodeMap.get("CHR")+offset;
		case "CLUTHA DISTRICT":return regionCodeMap.get("DUN")+offset;
		case "DUNEDIN CITY":return regionCodeMap.get("DUN")+offset;
		case "FAR NORTH DISTRICT":return regionCodeMap.get("WHA")+offset;
		case "GISBORNE DISTRICT":return regionCodeMap.get("GIS")+offset;
		case "GORE DISTRICT":return regionCodeMap.get("INV")+offset;
		case "GREY DISTRICT":return regionCodeMap.get("GRE")+offset;
		case "HAMILTON CITY":return regionCodeMap.get("HAM")+offset;
		case "HASTINGS DISTRICT":return regionCodeMap.get("NAP")+offset;
		case "HAURAKI DISTRICT":return regionCodeMap.get("THA")+offset;
		case "HOROWHENUA DISTRICT":return regionCodeMap.get("PAL")+offset;
		case "HURUNUI DISTRICT":return regionCodeMap.get("CHR")+offset;
		case "INVERCARGILL CITY":return regionCodeMap.get("INV")+offset;
		case "KAIKOURA DISTRICT":return regionCodeMap.get("CHR")+offset;
		case "KAIPARA DISTRICT":return regionCodeMap.get("WHA")+offset;
		case "KAPITI COAST DISTRICT":return regionCodeMap.get("WEL")+offset;
		case "KAWERAU DISTRICT":return regionCodeMap.get("TAU")+offset;
		case "LOWER HUTT CITY":return regionCodeMap.get("WEL")+offset;
		case "MACKENZIE DISTRICT":return regionCodeMap.get("CHR")+offset;
		case "MANAWATU DISTRICT":return regionCodeMap.get("PAL")+offset;
		case "MARLBOROUGH DISTRICT":return regionCodeMap.get("BLE")+offset;
		case "MASTERTON DISTRICT":return regionCodeMap.get("MAS")+offset;
		case "MATAMATA-PIAKO DISTRICT":return regionCodeMap.get("HAM")+offset;
		case "NAPIER CITY":return regionCodeMap.get("NAP")+offset;
		case "NELSON CITY":return regionCodeMap.get("NEL")+offset;
		case "NEW PLYMOUTH DISTRICT":return regionCodeMap.get("NEW")+offset;
		case "OPOTIKI DISTRICT":return regionCodeMap.get("TAU")+offset;
		case "OTOROHANGA DISTRICT":return regionCodeMap.get("HAM")+offset;
		case "PALMERSTON NORTH CITY":return regionCodeMap.get("PAL")+offset;
		case "PORIRUA CITY":return regionCodeMap.get("WEL")+offset;
		case "QUEENSTOWN-LAKES DISTRICT":return regionCodeMap.get("DUN")+offset;
		case "RANGITIKEI DISTRICT":return regionCodeMap.get("PAL")+offset;
		case "ROTORUA DISTRICT":return regionCodeMap.get("ROT")+offset;
		case "RUAPEHU DISTRICT":return regionCodeMap.get("WAN")+offset;
		case "SELWYN DISTRICT":return regionCodeMap.get("CHR")+offset;
		case "SOUTH TARANAKI DISTRICT":return regionCodeMap.get("NEW")+offset;
		case "SOUTH WAIKATO DISTRICT":return regionCodeMap.get("HAM")+offset;
		case "SOUTH WAIRARAPA DISTRICT":return regionCodeMap.get("WEL")+offset;
		case "SOUTHLAND DISTRICT":return regionCodeMap.get("INV")+offset;
		case "STRATFORD DISTRICT":return regionCodeMap.get("NEW")+offset;
		case "TARARUA DISTRICT":return regionCodeMap.get("PAL")+offset;
		case "TASMAN DISTRICT":return regionCodeMap.get("NEL")+offset;						// ***
		case "TAUPO DISTRICT":return regionCodeMap.get("HAM")+offset;
		case "TAURANGA CITY":return regionCodeMap.get("TAU")+offset;
		case "THAMES-COROMANDEL DISTRICT":return regionCodeMap.get("THA")+offset;
		case "TIMARU DISTRICT":return regionCodeMap.get("TIM")+offset;
		case "UPPER HUTT CITY":return regionCodeMap.get("WEL")+offset;
		case "WAIKATO DISTRICT":return regionCodeMap.get("HAM")+offset;
		case "WAIMAKARIRI DISTRICT":return regionCodeMap.get("CHR")+offset;
		case "WAIMATE DISTRICT":return regionCodeMap.get("CHR")+offset;
		case "WAIPA DISTRICT":return regionCodeMap.get("HAM")+offset;
		case "WAIROA DISTRICT":return regionCodeMap.get("NAP")+offset;
		case "WAITAKI DISTRICT":return regionCodeMap.get("OAM")+offset;
		case "WAITOMO DISTRICT":return regionCodeMap.get("HAM")+offset;
		case "WELLINGTON CITY":return regionCodeMap.get("WEL")+offset;
		case "WESTERN BAY OF PLENTY DISTRICT":return regionCodeMap.get("TAU")+offset;
		case "WESTLAND DISTRICT":return regionCodeMap.get("WES")+offset;
		case "WHAKATANE DISTRICT":return regionCodeMap.get("TAU")+offset;
		case "WHANGANUI DISTRICT":return regionCodeMap.get("WAN")+offset;
		case "WHANGAREI DISTRICT":return regionCodeMap.get("WHA")+offset;
		}
		return regionCodeMap.get("OTH")+offset;
	}
	
	/**
	 * Takes the cc_rating and returns the cell index based on cc_rating brackets
	 * @param ccRating
	 * @return
	 */
	public int getCCBracketCell(int ccRating, int offset) {
		if (ccRating >=1 && ccRating <= 850) return 0+offset;
		else if (ccRating >=851 && ccRating <= 1000) return 1+offset;
		else if (ccRating >=1001 && ccRating <= 1100) return 2+offset;
		else if (ccRating >=1101 && ccRating <= 1200) return 3+offset;
		else if (ccRating >=1201 && ccRating <= 1300) return 4+offset;
		else if (ccRating >=1301 && ccRating <= 1400) return 5+offset;
		else if (ccRating >=1401 && ccRating <= 1500) return 6+offset;
		else if (ccRating >=1501 && ccRating <= 1600) return 7+offset;
		else if (ccRating >=1601 && ccRating <= 1800) return 8+offset;
		else if (ccRating >=1801 && ccRating <= 2000) return 9+offset;
		else if (ccRating >=2001 && ccRating <= 2500) return 10+offset;
		else if (ccRating >=2501 && ccRating <= 3000) return 11+offset;
		else if (ccRating >=3001 && ccRating <= 3500) return 12+offset;
		else if (ccRating >=3501 && ccRating <= 4000) return 13+offset;
		else if (ccRating >=4001 && ccRating <= 4500) return 14+offset;
		else if (ccRating >=4501 && ccRating <= 5000) return 15+offset;
		else if (ccRating >=5001) return 16+offset;
		
		return 18+offset;
	}
	
	public int getGVMBracketCell(int gvm, int offset) {
		if (gvm >=1 && gvm <= 1500) return 0+offset;
		else if (gvm >=1501 && gvm <= 2000) return 1+offset;
		else if (gvm >=2001 && gvm <= 2500) return 2+offset;
		else if (gvm >=2501 && gvm <= 3500) return 3+offset;
		else if (gvm >=3501 && gvm <= 4500) return 4+offset;
		else if (gvm >=4501 && gvm <= 6500) return 5+offset;
		else if (gvm >=6501 && gvm <= 7500) return 6+offset;
		else if (gvm >=7501 && gvm <= 9000) return 7+offset;
		else if (gvm >=9001 && gvm <= 10500) return 8+offset;
		else if (gvm >=10501 && gvm <= 12000) return 9+offset;
		else if (gvm >=12001 && gvm <= 14500) return 10+offset;
		else if (gvm >=14501 && gvm <= 15000) return 11+offset;
		else if (gvm >=15001 && gvm <= 16000) return 12+offset;
		else if (gvm >=16001 && gvm <= 18000) return 13+offset;
		else if (gvm >=18001 && gvm <= 20500) return 14+offset;
		else if (gvm >=20501 && gvm <= 23000) return 15+offset;
		else if (gvm >23000) return 16+offset;
		
		return 20+offset;
	}
	
	public int getGVMBracketBusCell(int gvm, int offset) {
		if (gvm >=1 && gvm <= 3500) return 17+offset;
		else if (gvm >3500) return 18+offset;
		
		return 20+offset;
	}
	
	public int getAgeCell(int age) {
		if (age>10) return 11;
		else return age;
	}
	
	public static HashMap<String, String> getDistributorMap() {
		HashMap<String, String> dis = new HashMap<String, String>();
		
		dis.put("APRILIA", "Triumph"); dis.put("GAS GAS", "Triumph"); dis.put("KEEWAY", "Triumph");
		dis.put("MOTO GUZZI", "Triumph"); dis.put("PGO", "Triumph"); dis.put("PIAGGIO", "Triumph");
		dis.put("SYM", "Triumph"); dis.put("TRIUMPH", "Triumph"); dis.put("VESPA", "Triumph");
	    
		dis.put("ALFA ROMEO", "Ateco"); dis.put("CHERY", "Ateco"); dis.put("CHRYSLER", "Ateco");
		dis.put("DODGE", "Ateco"); dis.put("FIAT", "Ateco"); dis.put("JEEP", "Ateco");
		dis.put("RAM", "Ateco"); dis.put("MASERATI", "Ateco"); 
		
		dis.put("BMW", "BMW"); dis.put("MINI", "BMW");
	    
		dis.put("AUDI", "EMD"); dis.put("PORSCHE", "EMD"); dis.put("SKODA", "EMD");
		dis.put("VOLKSWAGEN", "EMD");
	    
		dis.put("LDV", "Great Lake"); dis.put("SSANGYONG", "Great Lake");
		
		dis.put("GREAT WALL", "Great Wall"); dis.put("HAVAL", "Great Wall");
	    
		dis.put("HUSQVARNA", "KTM"); dis.put("KTM", "KTM");
	    
		dis.put("JAGUAR", "Motorcorp"); dis.put("LAND ROVER", "Motorcorp");
	    
		dis.put("INDIAN", "Polaris"); dis.put("POLARIS", "Polaris"); dis.put("VICTORY", "Polaris");
	    
		dis.put("PEUGEOT", "Auto Distributors"); dis.put("CITROEN", "Auto Distributors");
	    
		dis.put("TOYOTA", "Toyota"); dis.put("LEXUS", "Toyota");
	    
		dis.put("HONDA", "Honda"); 
		dis.put("CAN-AM", "BRP"); 
		dis.put("DUCATI", "Ducati");
		dis.put("FORD", "Ford"); 
		dis.put("HARLEY DAVIDSON", "Harley Davidson"); 
		dis.put("HOLDEN", "Holden");
		dis.put("HYUNDAI", "Hyundai"); 
		dis.put("ISUZU", "Isuzu Utes"); 
		dis.put("KIA", "Kia");
		dis.put("KAWASAKI", "Kawasaki"); 
		dis.put("MAZDA", "Mazda"); 
		dis.put("MERCEDES-BENZ", "Mercedes-Benz");
		dis.put("MITSUBISHI", "Mitsubishi"); 
		dis.put("NISSAN", "Nissan"); 
		dis.put("SUBARU", "Subaru");
		dis.put("SUZUKI", "Suzuki"); 
		dis.put("TESLA", "Tesla"); 
		dis.put("YAMAHA", "Yamaha");
		
		return dis;
	}
	
	/**
	 *
	 * @param make
	 * @param model
	 * @return the motorcycle make that matches the model (if no model matches, then return the original make)
	 */
	public static String determineIfBike(String make, String model) {
		switch (make) {
		case "BMW":
			String[] BMW = new String[]{"F700", "F800", "R NINET", "S1000", "R1200"};
			for (String m: BMW) {
				if (m.equals(model)) return "Europe Imports";
			}
			break;
		case "HONDA":
			String[] Honda = new String[]{"NCH", "XR", "NVS", "CBR1000", "CB", "NC", "CMX"};
			for (String m: Honda) {
				if (m.equals(model)) return "Blue Wing Honda";
			}
			return "Honda";
		case "SUZUKI":
			String[] Suzuki = new String[]{"UZ50", "UZ", "DL650A", "DL1000", "DR200", "RV200", "GSX-R1000RA"};
			for (String m: Suzuki) {
				if (m.equals(model)) return "Suzuki Motorcycles";
			}
			return "Suzuki";
		case "PEUGEOT":
			String[] peugeot = new String[]{"KISBEE", "DJANGO", "SPEEDFIGHT"};
			for (String m: peugeot) {
				if (m.equals(model)) return "Other";
			}
			return "Auto Distributors";
		}
		
		return make;
	}
	
}
