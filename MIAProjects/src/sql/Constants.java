package sql;

public class Constants {
    // Database name.
    public static final String TABLE_NAME = "t";

    // Opendata constants
    public static final String passengerClasses = "('XMA', 'XMB', 'XMC', 'XLE')";
    public static final String PASS_CLASS_CLEAN = "('MA', 'MB', 'MC', 'LE')";
    public static final String commercialClasses = "('XMD1', 'XMD2', 'XNA', 'XMD', 'XMD3', 'XMD4', 'XME', 'XNB', 'XNC')";
    public static final String allClasses = "('XMA', 'XMB', 'XMC', 'XLE', 'XMD1', 'XMD2', 'XNA', 'XMD', 'XMD3', 'XMD4', 'XME', 'XNB', 'XNC')";
    public static final String motorcycleClasses = "('XLA', 'XLB', 'XLC', 'XLD')";
    public static final String newImportStatus = "('NEW')";
    public static final String usedImportStatus = "('USED')";
    public static final String allImportStatus = "('NEW', 'USED')";
    public static final String electric = "('DIESEL ELECTRIC HYBRID', 'ELECTRIC', 'ELECTRIC[PETROL EXTENDED]', 'PETROL ELECTRIC HYBRID', 'PLUGIN PETROL HYBRID')";

    // Opendata Columns
    public static final String OD_ALTERNATIVE_MOTIVE_POWER = "ALTERNATIVE_MOTIVE_POWER";
    public static final String OD_BASIC_COLOUR = "BASIC_COLOUR";
    public static final String OD_BODY_TYPE = "BODY_TYPE";
    public static final String OD_CC_RATING = "CC_RATING";
    public static final String OD_CHASSIS7 = "CHASSIS7";
    public static final String OD_CLASS = "CLASS";
    public static final String OD_ENGINE_NUMBER = "ENGINE_NUMBER";
    public static final String OD_FIRST_NZ_REGISTRATION_YEAR = "FIRST_NZ_REGISTRATION_YEAR";
    public static final String OD_FIRST_NZ_REGISTRATION_MONTH = "FIRST_NZ_REGISTRATION_MONTH";
    public static final String OD_GROSS_VEHICLE_MASS = "GROSS_VEHICLE_MASS";
    public static final String OD_HEIGHT = "HEIGHT";
    public static final String OD_IMPORT_STATUS = "IMPORT_STATUS";
    public static final String OD_INDUSTRY_CLASS = "INDUSTRY_CLASS";
    public static final String OD_INDUSTRY_MODEL_CODE = "INDUSTRY_MODEL_CODE";
    public static final String OD_MAKE = "MAKE";
    public static final String OD_MODEL = "MODEL";
    public static final String OD_MOTIVE_POWER= "MOTIVE_POWER";
    public static final String OD_MVMA_MODEL_CODE= "MVMA_MODEL_CODE";
    public static final String OD_NUMBER_OF_AXLES= "NUMBER_OF_AXLES";
    public static final String OD_NUMBER_OF_SEATS= "NUMBER_OF_SEATS";
    public static final String OD_NZ_ASSEMBLED= "NZ_ASSEMBLED";
    public static final String OD_ORIGINAL_COUNTRY= "ORIGINAL_COUNTRY";
    public static final String OD_POWER_RATING= "POWER_RATING";
    public static final String OD_PREVIOUS_COUNTRY= "PREVIOUS_COUNTRY";
    public static final String OD_ROAD_TRANSPORT_CODE= "ROAD_TRANSPORT_CODE";
    public static final String OD_SUBMODEL= "SUBMODEL";
    public static final String OD_TLA= "TLA";
    public static final String OD_TRANSMISSION_TYPE= "TRANSMISSION_TYPE";
    public static final String OD_VDAM_WEIGHT= "VDAM_WEIGHT";
    public static final String OD_VEHICLE_TYPE= "VEHICLE_TYPE";
    public static final String OD_VEHICLE_USAGE= "VEHICLE_USAGE";
    public static final String OD_VEHICLE_YEAR= "VEHICLE_YEAR";
    public static final String OD_VIN11= "VIN11";
    public static final String OD_WIDTH = "WIDTH";

    public static final String OD_TOTAL = "TOTAL";

    // NZTA Table names
    public static final String T_001 = "001";
    public static final String T_001A = "001A";
    public static final String T_001N = "001N";
    public static final String T_002 = "002";
    public static final String T_002A = "002A";
    public static final String T_002N = "002N";
    public static final String T_006 = "006";
    public static final String T_006N = "006N";
    public static final String T_006X = "006X";
    public static final String T_008 = "008";
    public static final String T_008N = "008N";
    public static final String T_008X = "008X";
    public static final String T_051 = "051";
    public static final String T_054 = "054";
    public static final String T_064N = "064N";
    public static final String T_064X = "064X";
    public static final String T_065N = "065N";
    public static final String T_065X = "065X";
    public static final String T_MIA_DEREG_MONTHLY = "MIA_DEREG_MONTHLY";
    public static final String T_MOTORCYCLES_NEW = "Motorcycles New";
    public static final String T_N7_USG = "N7-USG";
    public static final String T_U7MM_AGE = "U7MM_AGE";
    public static final String T_U8MM_AGE = "U8MM_AGE";
    public static final String T_VTYP10_13 = "VTyp10-13";
    public static final String T_X_085N = "X-085N";
    public static final String T_Y_MPC_A = "Y_MPC_A";
    public static final String T_Y_001AN = "Y-001AN";
    public static final String T_Y_001AN_2AN = "Y-001AN_2AN";
    public static final String T_Y_001AX = "Y-001AX";
    public static final String T_Y_002AN = "Y-002AN";
    public static final String T_Y_002AX = "Y-002AX";
    public static final String T_Y_065N = "Y-065N";
    public static final String T_Y_080N = "Y-080N";
    public static final String T_Y_081N = "Y-081N";
    public static final String T_Y_084N = "Y-084N";
    public static final String T_Y_085N = "Y-085N";
    public static final String T_Y_MPC50 = "Y-MPC50";
    public static final String T_Y_MPC51 = "Y-MPC51";
    public static final String T_YRY_COMMS_M1 = "YRY-COMMS_M1";
    public static final String T_YTD_RENTALS_NEW = "YTD_RENTALS_NEW";
    public static final String T_YTD_USED_CARS = "YTD_USED_CARS";
    public static final String T_YTD_USED_COM = "YTD_USED_COM";

    // Arrays
    public static final String[] regionCodeNames = new String[]{
            "WHA", "AUC",
            "HAM", "THA",
            "TAU", "ROT",
            "GIS", "NAP",
            "NEW", "WAN",
            "PAL", "MAS",
            "WEL", "NEL",
            "BLE", "GRE",
            "WES", "CHR",
            "TIM", "OAM",
            "DUN", "INV",
            "OTH"};
    public static final String[] ccBrackets = new String[]{
            "1 - 850",
            "851 - 1000",
            "1001 - 1100",
            "1101 - 1200",
            "1201 - 1300",
            "1301 - 1400",
            "1401 - 1500",
            "1501 - 1600",
            "1601 - 1800",
            "1801 - 2000",
            "2001 - 2500",
            "2501 - 3000",
            "3001 - 3500",
            "3501 - 4000",
            "4001 - 4500",
            "4501 - 5000",
            "5001 AND OVER"};
    public static final String[] gvmBrackets = new String[]{
            "1 - 1500",
            "1501 - 2000",
            "2001 - 2500",
            "2501 - 3500",
            "3501 - 4500",
            "4501 - 6500",
            "6501 - 7500",
            "7501 - 9000",
            "9001 - 10500",
            "10501 - 12000",
            "12001 - 14500",
            "14501 - 15000",
            "15001 - 16000",
            "16001 - 18000",
            "18001 - 20500",
            "20501 - 23000",
            "23001 AND OVER",
            "BUS UP TO 3500",
            "BUS OVER 3500"};
    public static final String[] months = new String[]{
            "JAN", "FEB",
            "MAR", "APR",
            "MAY", "JUN",
            "JUL", "AUG",
            "SEP", "OCT",
            "NOV", "DEC"};
    public static final String[] ages = new String[]{
            "0 YEAR",
            "1 YEAR",
            "2 YEARS",
            "3 YEARS",
            "4 YEARS",
            "5 YEARS",
            "6 YEARS",
            "7 YEARS",
            "8 YEARS",
            "9 YEARS",
            "10 YEARS",
            "OVER 10 YEARS"};

    // SQL statements
    public static final String SQL_TOTAL = "\""+OD_TOTAL+"\"";

    public static String sqlMakeCountryTLA(String segment, String importStatus, int year, int monthNumber) {
        return "select MAKE, ORIGINAL_COUNTRY, TLA, count(MAKE) "+SQL_TOTAL+" "+
                "from "+ TABLE_NAME +" "+
                "where class in "+segment+" " +
                "and IMPORT_STATUS in "+importStatus+" "+
                "and FIRST_NZ_REGISTRATION_YEAR = '"+year+"' " +
                "and FIRST_NZ_REGISTRATION_MONTH = '"+monthNumber+"' "+
                "group by make, ORIGINAL_COUNTRY, TLA " +
                "order by make, ORIGINAL_COUNTRY, TLA;";
    }

    public static String sqlMakeImportStatus(String segment, int year, int monthNumber) {
        return "Select MAKE, IMPORT_STATUS, COUNT(MAKE) "+SQL_TOTAL+" " +
                "from "+ TABLE_NAME +" " +
                "where class in "+segment+" and " +
                "IMPORT_STATUS in "+allImportStatus+" " +
                "and FIRST_NZ_REGISTRATION_YEAR = '"+year+"' " +
                "and FIRST_NZ_REGISTRATION_MONTH = '"+monthNumber+"' " +
                "group by make, IMPORT_STATUS " +
                "order by make, IMPORT_STATUS;";
    }

    public static String sqlMakeCountryCC(String segment, String importStatus, int year, int monthNumber) {
        return "select MAKE, ORIGINAL_COUNTRY, CC_RATING, count(MAKE) "+SQL_TOTAL+" "+
                "from "+ TABLE_NAME +" " +
                "where class in "+segment+" " +
                "and IMPORT_STATUS in "+importStatus+" "+
                "and FIRST_NZ_REGISTRATION_YEAR = '"+year+"' " +
                "and FIRST_NZ_REGISTRATION_MONTH = '"+monthNumber+"' "+
                "group by MAKE, ORIGINAL_COUNTRY, CC_RATING " +
                "order by MAKE, ORIGINAL_COUNTRY, CC_RATING;";
    }

    public static String sqlMakeCountryGVM(String segment, String importStatus, int year, int monthNumber) {
        return "select MAKE, ORIGINAL_COUNTRY, CLASS, GROSS_VEHICLE_MASS, count(MAKE) "+SQL_TOTAL+" "+
                "from "+ TABLE_NAME +" " +
                "where class in "+segment+" " +
                "and IMPORT_STATUS in "+importStatus+" "+
                "and FIRST_NZ_REGISTRATION_YEAR = '"+year+"' " +
                "and FIRST_NZ_REGISTRATION_MONTH = '"+monthNumber+"' "+
                "group by MAKE, ORIGINAL_COUNTRY, CLASS, GROSS_VEHICLE_MASS " +
                "order by MAKE, ORIGINAL_COUNTRY, CLASS, GROSS_VEHICLE_MASS;";
    }

    public static String sqlMakeModelCountryTLA(String segment, String importStatus, int year, int monthNumber) {
        return "select MAKE, MODEL, ORIGINAL_COUNTRY, TLA, count(MAKE) "+SQL_TOTAL+" "+
                "from "+ TABLE_NAME +" " +
                "where class in "+segment+" " +
                "and IMPORT_STATUS in "+importStatus+" "+
                "and FIRST_NZ_REGISTRATION_YEAR = '"+year+"' " +
                "and FIRST_NZ_REGISTRATION_MONTH = '"+monthNumber+"' "+
                "group by MAKE, MODEL, ORIGINAL_COUNTRY, TLA " +
                "order by MAKE, MODEL, ORIGINAL_COUNTRY, TLA;";
    }

    public static String sqlMakeModelFuelAge(String segment, String importStatus, int year, int monthNumber) {
        return "select MAKE, MODEL, MOTIVE_POWER, VEHICLE_YEAR, count(MAKE) "+SQL_TOTAL+" "+
                "from "+ TABLE_NAME +" " +
                "where class in "+segment+" " +
                "and IMPORT_STATUS in "+importStatus+" "+
                "and FIRST_NZ_REGISTRATION_YEAR = '"+year+"' " +
                "and FIRST_NZ_REGISTRATION_MONTH = '"+monthNumber+"' "+
                "group by MAKE, MODEL, MOTIVE_POWER, VEHICLE_YEAR " +
                "order by MAKE, MODEL, MOTIVE_POWER, VEHICLE_YEAR;";
    }

    public static String sqlMakeYTD(String segment, String importStatus, int year, String ytd) {
        return "select MAKE, FIRST_NZ_REGISTRATION_MONTH, count(MAKE) "+SQL_TOTAL+" " +
                "from "+ TABLE_NAME +" " +
                "where class in "+segment+" " +
                "and IMPORT_STATUS in "+importStatus+" " +
                "and FIRST_NZ_REGISTRATION_YEAR = '"+year+"' " +
                "and FIRST_NZ_REGISTRATION_MONTH in "+ytd+" " +
                "group by MAKE, FIRST_NZ_REGISTRATION_MONTH " +
                "order by MAKE, FIRST_NZ_REGISTRATION_MONTH;";
    }

    public static String sqlMakeManyYTD(String segment, String importStatus, int year, String ytd) {
        return "select MAKE, MODEL, SUBMODEL, ORIGINAL_COUNTRY, NZ_ASSEMBLED, CC_RATING, BODY_TYPE, MOTIVE_POWER, NUMBER_OF_AXLES, "+
                "POWER_RATING, FIRST_NZ_REGISTRATION_MONTH, count(MAKE) "+SQL_TOTAL+" " +
                "from "+ TABLE_NAME +" " +
                "where class in "+segment+" " +
                "and IMPORT_STATUS in "+importStatus+" " +
                "and FIRST_NZ_REGISTRATION_YEAR = '"+year+"' " +
                "and FIRST_NZ_REGISTRATION_MONTH in "+ytd+" " +
                "group by MAKE, MODEL, SUBMODEL, ORIGINAL_COUNTRY, NZ_ASSEMBLED, CC_RATING, " +
                "BODY_TYPE, MOTIVE_POWER, NUMBER_OF_AXLES, POWER_RATING, FIRST_NZ_REGISTRATION_MONTH "+
                "order by MAKE, MODEL, SUBMODEL, ORIGINAL_COUNTRY, NZ_ASSEMBLED, CC_RATING, " +
                "BODY_TYPE, MOTIVE_POWER, NUMBER_OF_AXLES, POWER_RATING, FIRST_NZ_REGISTRATION_MONTH;";
    }

    public static String sqlMakeModelYTD(String segment, String importStatus, int year, String ytd) {
        return "select MAKE, MODEL, FIRST_NZ_REGISTRATION_MONTH, count(MAKE) "+SQL_TOTAL+" " +
                "from "+ TABLE_NAME +" " +
                "where class in "+segment+" " +
                "and IMPORT_STATUS in "+importStatus+" " +
                "and FIRST_NZ_REGISTRATION_YEAR = '"+year+"' " +
                "and FIRST_NZ_REGISTRATION_MONTH in "+ytd+" " +
                "group by MAKE, MODEL, FIRST_NZ_REGISTRATION_MONTH " +
                "order by MAKE, MODEL, FIRST_NZ_REGISTRATION_MONTH;";
    }

    public static String sqlMakeImportYTD(String segment, String importStatus, int year, String ytd) {
        return "select MAKE, IMPORT_STATUS, FIRST_NZ_REGISTRATION_MONTH, count(MAKE) "+SQL_TOTAL+" " +
                "from "+ TABLE_NAME +" " +
                "where class in "+segment+" " +
                "and IMPORT_STATUS in "+importStatus+" " +
                "and FIRST_NZ_REGISTRATION_YEAR = '"+year+"' " +
                "and FIRST_NZ_REGISTRATION_MONTH in "+ytd+" " +
                "group by MAKE, IMPORT_STATUS, FIRST_NZ_REGISTRATION_MONTH " +
                "order by MAKE, IMPORT_STATUS, FIRST_NZ_REGISTRATION_MONTH;";
    }

    public static String sqlMakeModelSubCCYTD(String segment, String importStatus, int year, String ytd, String cc) {
        return "select MAKE, MODEL, SUBMODEL, FIRST_NZ_REGISTRATION_MONTH, count(MAKE) "+SQL_TOTAL+" " +
                "from "+ TABLE_NAME +" " +
                "where class in "+segment+" " +
                "and IMPORT_STATUS in "+importStatus+" " +
                "and CC_RATING "+cc+" " +
                "and FIRST_NZ_REGISTRATION_YEAR = '"+year+"' " +
                "and FIRST_NZ_REGISTRATION_MONTH in "+ytd+" " +
                "group by MAKE, MODEL, SUBMODEL, FIRST_NZ_REGISTRATION_MONTH " +
                "order by MAKE, MODEL, SUBMODEL, FIRST_NZ_REGISTRATION_MONTH;";
    }

    public static String sqlMakeModelTypeRental(String segment, String importStatus, int year, String ytd) {
        return "select MAKE, MODEL, VEHICLE_TYPE, FIRST_NZ_REGISTRATION_MONTH, count(MAKE) "+SQL_TOTAL+" " +
                "from "+ TABLE_NAME +" " +
                "where class in "+segment+" " +
                "and IMPORT_STATUS in "+importStatus+" " +
                "and VEHICLE_USAGE = 'RENTAL' " +
                "and FIRST_NZ_REGISTRATION_YEAR = '"+year+"' " +
                "and FIRST_NZ_REGISTRATION_MONTH in "+ytd+" " +
                "group by MAKE, MODEL, VEHICLE_TYPE, FIRST_NZ_REGISTRATION_MONTH " +
                "order by MAKE, MODEL, VEHICLE_TYPE, FIRST_NZ_REGISTRATION_MONTH;";
    }

    // Filenames
    public static String getFileNameHeader(String fileName, String month) {
        switch (fileName) {
            case T_001:         return(month + " - NEW AND EX-OVERSEAS CARS REGISTERED BY: MAKE, COUNTRY, POSTAL DISTRICT");
            case T_001A:        return(month + " - CARS NEWLY REGISTERED: NEW vs EX-OVERSEAS");
            case T_001N:        return(month + " - NEW CARS REGISTERED BY: MAKE, COUNTRY, POSTAL DISTRICT");
            case T_002:         return(month + " - NEW AND EX-OVERSEAS COMMERCIALS REGISTERED BY: MAKE, COUNTRY, POSTAL DISTRICT");
            case T_002A:        return(month + " - COMMERCIALS NEWLY REGISTERED: NEW vs EX-OVERSEAS");
            case T_002N:        return(month + " - NEW AND EX-OVERSEAS COMMERCIALS REGISTERED BY: MAKE, COUNTRY, POSTAL DISTRICT");
            case T_006:         return(month + " - NEW AND EX-OVERSEAS CARS REGISTERED BY: MAKE, COUNTRY, CC RATING");
            case T_006N:        return(month + " - NEW CARS REGISTERED BY: MAKE, COUNTRY, CC RATING");
            case T_006X:        return(month + " - EX-OVERSEAS CARS REGISTERED BY: MAKE, COUNTRY, CC RATING");
            case T_008:         return(month + " - NEW AND EX-OVERSEAS COMMERCIALS REGISTERED BY: MAKE, COUNTRY, GROSS VEHICLE MASS (KG)");
            case T_008N:        return(month + " - NEW COMMERCIALS REGISTERED BY: MAKE, COUNTRY, GROSS VEHICLE MASS (KG)");
            case T_008X:        return(month + " - EX-OVERSEAS COMMERCIALS REGISTERED BY: MAKE, COUNTRY, GROSS VEHICLE MASS (KG)");
            case T_064N:        return(month + " - NEW CARS REGISTERED BY: MAKE, MODEL, COUNTRY, POSTAL DISTRICT");
            case T_064X:        return(month + " - EX-OVERSEAS CARS REGISTERED BY: MAKE, MODEL, COUNTRY, POSTAL DISTRICT");
            case T_065N:        return(month + " - NEW COMMERCIALS REGISTERED BY: MAKE, MODEL, COUNTRY, POSTAL DISTRICT");
            case T_065X:        return(month + " - EX-OVERSEAS COMMERCIALS REGISTERED BY: MAKE, MODEL, COUNTRY, POSTAL DISTRICT");
            case T_Y_001AN:     return("NEW CARS BY MAKE, YEAR TO DATE");
            case T_Y_001AN_2AN: return("NEW CARS AND COMMERCIAL BY MAKE, YEAR TO DATE");
            case T_Y_001AX:     return("EX-OVERSEAS CARS BY MAKE, YEAR TO DATE");
            case T_Y_002AN:     return("NEW COMMERCIALS BY MAKE, YEAR TO DATE");
            case T_Y_002AX:     return("EX-OVERSEAS COMMERCIALS BY MAKE, YEAR TO DATE");
            case T_Y_MPC_A:     return("MOPEDS/MOTORCYCLES NEWLY REGISTERED: NEW vs EX-OVERSEAS - YEAR-TO-DATE");
            case T_MOTORCYCLES_NEW:     return("MOPEDS/MOTORCYCLES NEWLY REGISTERED: NEW - YEAR-TO-DATE");
        }

        return "";
    }

    public static String[] getFileNameHeaderMotorcycle(String fileName) {
        switch (fileName) {
            case T_Y_MPC50:     return new String[]{"NEW MOPEDS/MOTORCYCLES UP TO 50CC BY MAKE, YEAR-TO-DATE", "< 51"};
            case T_Y_MPC51:     return new String[]{"NEW MOPEDS/MOTORCYCLES OVER 50CC BY MAKE, YEAR-TO-DATE", "> 50"};
        }

        return new String[]{""};
    }
}