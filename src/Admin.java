
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.GregorianCalendar;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Admin {
	
	String courseName;
	String bookingDate;
	String scorecardExcelName;
	final int ROW_FIRST_GROUP_STATUS = 10;
	final int GROUP_TOTAL = 64;
	
	
	public Admin(String course) {
		courseName = course;
		scorecardExcelName = "/Users/apple/Desktop/GolfComfy/sim/" + courseName + "_Scorecard.xls";
		System.out.println("\nAccessing data as an Admin");
	}
	
	void displayTraffic() {
		try
		{
			System.out.println("Retrieving hole information........");
			FileInputStream fileIn= new FileInputStream(new File(scorecardExcelName));
			Workbook wb = new HSSFWorkbook (fileIn);
			Sheet ws = wb.getSheetAt(0);
			for(int group=1; group<GROUP_TOTAL+1;group++) {
				Row sr = ws.getRow(6*group+4);
				Cell sc = sr.getCell(2);
				if(sc!=null && sc.getStringCellValue().equals("playing")) {
					int location=0;
					Row hr = ws.getRow(6*group);
					for(int c=23;c>4;c--) {
						Cell ce = hr.getCell(c);
						if(ce!=null) {
							location = c-3;
							//System.out.println(location);
						}
					}
					int l=6*group;
					Row groupRow = ws.getRow(l);
					Cell cell_player = groupRow.getCell(2);
					int player_number=0;
					while(cell_player != null) {
						player_number++;
						l++;
						groupRow = ws.getRow(l);
						cell_player = groupRow.getCell(2);
					}
					System.out.printf("There are %d players at hole %d\n",player_number, location);
				}
			}
			fileIn.close();
			wb.close();
		}catch(Exception e){
			System.out.println("Fail to display the location map of playing groups");
			return;
		}
	}

	void checkTraffic() {
		int factor=0;
		try
		{
			System.out.println("\nChecking current traffic........");
			FileInputStream fileIn= new FileInputStream(new File(scorecardExcelName));
			Workbook wb = new HSSFWorkbook (fileIn);
			Sheet ws = wb.getSheetAt(0);
			for(int group=1; group<GROUP_TOTAL+1;group++) {
				Row sr = ws.getRow(6*group+4);
				Cell sc = sr.getCell(2);
				if(sc!=null && sc.getStringCellValue().equals("playing")) {
					int prevlocation=0;
					Row hr = ws.getRow(6*group);
					for(int c=23;c>4;c--) {
						Cell ce = hr.getCell(c);
						if(ce!=null) {
							prevlocation = c-4;
							//System.out.println(location);
						}
					}
					int l=6*group;
					Row groupRow = ws.getRow(l);
					Cell cell_player = groupRow.getCell(2);
					int player_number=0;
					while(cell_player != null) {
						player_number++;
						l++;
						groupRow = ws.getRow(l);
						cell_player = groupRow.getCell(2);
					}
					//System.out.printf(player_number);
					////////////////////////////////////////////////////
					Row tr = ws.getRow(6*group+4);
					Cell tc = tr.getCell(4+prevlocation);
					String prevfinishedtime = tc.getStringCellValue();
					String pft[] = prevfinishedtime.split(":");
					int pftmm = Integer.parseInt(pft[1]);
					SimpleDateFormat SDF = new SimpleDateFormat("HH:mm");
					Calendar c = new GregorianCalendar();
					String currenttime = SDF.format(c.getTime());
					String ct[] = currenttime.split(":");
					int ctmm = Integer.parseInt(ct[1]);
					//System.out.println(pftmm);
					//System.out.println(ctmm);
					int chpt;
					if(ctmm<pftmm) {
						chpt = 60-pftmm+ctmm;
					}else {
						chpt = ctmm-pftmm; 
					}
					//System.out.println(chpt);
					Row spt_row = ws.getRow(2);
					Cell spt_cell = spt_row.getCell(4+prevlocation);
					int suggestedplaytime = (int) spt_cell.getNumericCellValue();
					//System.out.println(suggestedplaytime);
					if(chpt>suggestedplaytime) {
						int diff = chpt-suggestedplaytime;
						System.out.printf("The group with %d player(s) at hole %d is playing too slow, exceeding the suggested time by %d minutes\n",player_number, prevlocation+1,diff);
						factor=1;
					}
				}
			}
			if(factor==0) {
				System.out.println("All groups are proceeding fine");
			}
			fileIn.close();
			wb.close();
		}catch(Exception e) {
			System.out.println("Fail to detect traffic");
		}
	}
	
	void detectGroupsLocation() {	
		try
		{
			int[] activeHoles = new int[18];
			int totalPlayingGroups=0;
			System.out.println();
			FileInputStream fileIn= new FileInputStream(new File(scorecardExcelName));
			Workbook wb = new HSSFWorkbook (fileIn);
			Sheet ws = wb.getSheetAt(0);
			for(int group=1; group<GROUP_TOTAL+1;group++) {
				Row sr = ws.getRow(6*group+4);
				Cell sc = sr.getCell(2);
				if(sc!=null && sc.getStringCellValue().equals("playing")) {
					int prevlocation=0;
					Row hr = ws.getRow(6*group);
					for(int c=23;c>4;c--) {
						Cell ce = hr.getCell(c);
						if(ce!=null) {
							totalPlayingGroups++;
							prevlocation = c-4;
							System.out.printf("There's a group at hole %d\n",prevlocation+1);
						}
					}
				}
			}
			System.out.printf("There are %d groups on the course", totalPlayingGroups);
			//////////////////////////////////////////////////////////////////////////////////////////////
			
			//////////////////////////////////////////////////////////////////////////////////////////////
			wb.close();
			fileIn.close();
		}catch(Exception e) {
			System.out.println("Fail to retrieve all playing groups info");
		}
	}
	
	void markPlayerLate(String pn) {
		String playerProfileExcelName = "/Users/apple/Desktop/GolfComfy/sim/" + pn + "_profile.xls";
		try
		{
			FileInputStream fileIn= new FileInputStream(new File(playerProfileExcelName));
			Workbook wb = new HSSFWorkbook (fileIn);
			Sheet ws = wb.getSheetAt(0);
			Row lr = ws.getRow(2);
			Cell lc = lr.getCell(3);
			String scv = lc.getStringCellValue();
			//System.out.println(scv);
			int lates = Integer.parseInt(scv);
			lates = lates + 1;
			scv = Integer.toString(lates);
			//System.out.println(scv);
			lc.setCellType(CellType.STRING);
			lc.setCellValue(scv);
			FileOutputStream fileOut = new FileOutputStream(new File(playerProfileExcelName));
		    wb.write(fileOut);
		    fileIn.close();
		    wb.close();
		    fileOut.close();
			System.out.printf("\nPlayer %s is marked late\n",pn);
		}catch(Exception e) {
			System.out.printf("Fail to mark player %s late\n",pn);
		}
	}
}
