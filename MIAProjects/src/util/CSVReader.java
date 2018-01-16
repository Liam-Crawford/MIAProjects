package util;

import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;

public class CSVReader {
	
	public static void main(String[] args) {
		printCSV("Z:/Liam working folder/NZTA Open Data/Test Tables/001A.csv");
	}
	
	private static void printCSV(String csvFile){
		String line = "";
		String csvSplitBy = ",";
		ArrayList<String[]> vehicles = new ArrayList<String[]>();
		
		try (BufferedReader br = new BufferedReader(new FileReader(csvFile))){
			
			while ((line = br.readLine()) != null){
				String[] v = line.split(csvSplitBy);
				vehicles.add(v);
			}
			
		}catch (IOException e){
			e.printStackTrace();
		}
		
		for (String[] v : vehicles){
			System.out.println(v[0] + " " + v[1] + " " + v[2]);
		}
	}
}
