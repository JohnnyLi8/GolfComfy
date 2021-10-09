
import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Profile {
	String profileExcelName;
	String playerName;
	int avgScore;
	final int SCOREFIRSTLINE=9;
	
	
	public Profile(String pn) {
		playerName = pn;
		profileExcelName = "/Users/apple/Desktop/GolfComfy/sim" + playerName + "_profile.xls";
		if (new File(profileExcelName).exists()==false) {
			System.out.printf("%s's profile does not exist\n");
			return;
		}
	}
	
	int AvgScore() {
		try
		{
			FileInputStream fileIn= new FileInputStream(new File(profileExcelName));
			Workbook wb = new HSSFWorkbook (fileIn);
			Sheet ws = wb.getSheetAt(0);
			int l = SCOREFIRSTLINE;
			int score=0;
			Row r = ws.getRow(l);
			if(r==null)return -1;
			int n=0;
			while(r!=null) {
				Cell c = r.getCell(2);
				String s = c.getStringCellValue();
				int sc = Integer.parseInt(s);
				score=score+sc;
				l++;
				n++;
				r = ws.getRow(l);
			}
			avgScore= (int) (score/n);
			System.out.printf("The average score of player %s is: %d\n",playerName,avgScore);
			wb.close();
			return avgScore;
		}catch(Exception e) {
			System.out.printf("Fail to get the average score of player %s\n",playerName);
			return -1;
		}
	}
	
	int totalBIRDIEs() {
		try
		{
			FileInputStream fileIn= new FileInputStream(new File(profileExcelName));
			Workbook wb = new HSSFWorkbook (fileIn);
			Sheet ws = wb.getSheetAt(0);
			int l = SCOREFIRSTLINE;
			int birdies=0;
			Row r = ws.getRow(l);
			if(r==null)return -1;
			int n=0;
			while(r!=null) {
				Cell c = r.getCell(3);
				String s = c.getStringCellValue();
				int bd = Integer.parseInt(s);
				birdies=birdies+bd;
				l++;
				n++;
				r = ws.getRow(l);
			}
			System.out.printf("The total BIRDIEs of player %s is: %d\n",playerName,birdies);
			wb.close();
			return birdies;
		}catch(Exception e) {
			System.out.printf("Fail to get the total BIRDIEs of player %s\n",playerName);
			return -1;
		}
	}
	
	int totalPARs() {
		try
		{
			FileInputStream fileIn= new FileInputStream(new File(profileExcelName));
			Workbook wb = new HSSFWorkbook (fileIn);
			Sheet ws = wb.getSheetAt(0);
			int l = SCOREFIRSTLINE;
			int pars=0;
			Row r = ws.getRow(l);
			if(r==null)return -1;
			int n=0;
			while(r!=null) {
				Cell c = r.getCell(4);
				String s = c.getStringCellValue();
				int p = Integer.parseInt(s);
				pars=pars+p;
				l++;
				n++;
				r = ws.getRow(l);
			}
			System.out.printf("The total pars of player %s is: %d\n",playerName,pars);
			wb.close();
			return pars;
		}catch(Exception e) {
			System.out.printf("Fail to get the total PARs of player %s\n",playerName);
			return -1;
		}
	}
	
	int totalBOGEYs() {
		try
		{
			FileInputStream fileIn= new FileInputStream(new File(profileExcelName));
			Workbook wb = new HSSFWorkbook (fileIn);
			Sheet ws = wb.getSheetAt(0);
			int l = SCOREFIRSTLINE;
			int bg=0;
			Row r = ws.getRow(l);
			if(r==null)return -1;
			int n=0;
			while(r!=null) {
				Cell c = r.getCell(5);
				String s = c.getStringCellValue();
				int p = Integer.parseInt(s);
				bg=bg+p;
				l++;
				n++;
				r = ws.getRow(l);
			}
			System.out.printf("The total BOGEYs of player %s is: %d\n",playerName,bg);
			wb.close();
			return bg;
		}catch(Exception e) {
			System.out.printf("Fail to get the total BOGEYs of player %s\n",playerName);
			return -1;
		}
	}
	
	int totalDOUBLEBOGEYs() {
		try
		{
			FileInputStream fileIn= new FileInputStream(new File(profileExcelName));
			Workbook wb = new HSSFWorkbook (fileIn);
			Sheet ws = wb.getSheetAt(0);
			int l = SCOREFIRSTLINE;
			int dbg=0;
			Row r = ws.getRow(l);
			if(r==null)return -1;
			int n=0;
			while(r!=null) {
				Cell c = r.getCell(6);
				String s = c.getStringCellValue();
				int p = Integer.parseInt(s);
				dbg=dbg+p;
				l++;
				n++;
				r = ws.getRow(l);
			}
			System.out.printf("The total DOUBLE BOGEYs of player %s is: %d\n",playerName,dbg);
			wb.close();
			return dbg;
		}catch(Exception e) {
			System.out.printf("Fail to get the total DOUBLE BOGEYs of player %s\n",playerName);
			return -1;
		}
	}
	
	int totalSTANDARDONs() {
		try
		{
			FileInputStream fileIn= new FileInputStream(new File(profileExcelName));
			Workbook wb = new HSSFWorkbook (fileIn);
			Sheet ws = wb.getSheetAt(0);
			int l = SCOREFIRSTLINE;
			int so=0;
			Row r = ws.getRow(l);
			if(r==null)return -1;
			int n=0;
			while(r!=null) {
				Cell c = r.getCell(7);
				String s = c.getStringCellValue();
				int p = Integer.parseInt(s);
				so=so+p;
				l++;
				n++;
				r = ws.getRow(l);
			}
			System.out.printf("The total STANDARD ONs of player %s is: %d\n",playerName,so);
			wb.close();
			return so;
		}catch(Exception e) {
			System.out.printf("Fail to get the total STANDARD ONs of player %s\n",playerName);
			return -1;
		}
	}
	
}
