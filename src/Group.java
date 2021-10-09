
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.Scanner;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Group{
	String bookingDate;
	String courseName;
	String scorecardExcelName;
	String tee_time;
	int group_number;
	final int ROW_FIRST_GROUP = 6;
	final int GROUP_TOTAL = 64;
	final int CELL_FIRST_HOLE = 5;
	final int RECORDSTARTLINE = 9;
	int group_startline;
	int player_number;
	int num_hole;
	String player1;
	String player2;
	String player3;
	String player4;
	int location;
	boolean initilized;
	int holes_thru;
	
	public Group(String course, String bd, String teeTime, int player_num, int h) {
		bookingDate = bd;
		courseName = course;
		tee_time=teeTime;
		if(h!=18 && h!=9) {
			System.out.println("Not allowed to play this many holes");
			return;
		}
		num_hole=h;
		scorecardExcelName = "/Users/apple/Desktop/GolfComfy/sim/" + courseName + "_Scorecard.xls";
		if(player_num >4 || player_num<1) {
			System.out.println("wrong number of people in the group");
			return;
		}
		loadCourseInfo c = new loadCourseInfo(courseName); 
		if (c.pars != null && c.time != null) {
			bookingExcel coursebooking = new bookingExcel(c.name);
			scorecardExcel scorecard = new scorecardExcel(c.name);
			player_number=player_num;
			//////////////////////////////////////////////////////////
			if( coursebooking.book(bookingDate, tee_time, player_number, num_hole)==1 ) {
				if(init()==1) {
					if (player_number==1) {
						createProfile p1 = new createProfile(player1);
					}else if(player_number==2) {
						createProfile p1 = new createProfile(player1);
						createProfile p2 = new createProfile(player2);
					}else if(player_number==3) {
						createProfile p1 = new createProfile(player1);
						createProfile p2 = new createProfile(player2);
						createProfile p3 = new createProfile(player3);
					}else if(player_number==4) {
						createProfile p1 = new createProfile(player1);
						createProfile p2 = new createProfile(player2);
						createProfile p3 = new createProfile(player3);
						createProfile p4 = new createProfile(player4);
					}
					holes_thru=0;
				}else {
					System.out.println("Initialization failure");
					return;
				}
			}else if( coursebooking.book(bookingDate, tee_time, player_number, num_hole)==2 ) {
				if(retrieveGroupInfo()==1) {
					//System.out.println("Group info retrieved");
					return;
				}
				System.out.println("Fail to retreive group info");
				return;
			}else {
				System.out.println("Invalid booking informatoin provided");
			}
		}else {
			System.out.println("wrong course information provided");
			return;
		}

	}
	
	
	int retrieveGroupInfo() {
		try
		{
			System.out.println("Retrieving group information........ ");
			FileInputStream fileIn= new FileInputStream(new File(scorecardExcelName));
			Workbook wb = new HSSFWorkbook (fileIn);
			Sheet ws = wb.getSheet(bookingDate);
			int l = gotoGroup(tee_time);
			Row groupRow = ws.getRow(l);
			Cell cell_teetime = groupRow.getCell(0);
			tee_time = cell_teetime.getStringCellValue();
			Cell cell_groupnum = groupRow.getCell(1);
			String str_groupnum = cell_groupnum.getStringCellValue().substring(1);
			group_number = Integer.parseInt(str_groupnum);
			Cell cell_player = groupRow.getCell(2);
			player_number=0;
			while(cell_player != null) {
				player_number++;
				l++;
				groupRow = ws.getRow(l);
				cell_player = groupRow.getCell(2);
			}
			System.out.printf("The group is starting at %s \n", tee_time);
			System.out.printf("There are %d players in group %d \n", player_number, group_number);
			getLocation();
			System.out.println();
			return 1;
		}catch(Exception e) {
			System.out.println("Fail to retrieve");
			return -1;
		}
	}
	
	int getPlayerName(int player_number) {
		Scanner cnSc = new Scanner(System.in);
		System.out.println("\nPlease enter the names of the players in your group for your scorecard (separated by commas):");
		String names = cnSc.nextLine();
		System.out.println();
		if(player_number==1) {
			player1=names;
		}
		if(player_number>1) {
			String split[] = names.split(",");
			if( split.length != player_number ) {
				System.out.println("Number of name entries do not match player numbers");
				cnSc.close();
				return -1;
			}
			for (int p=0;p<player_number;p++) {
				if(p==0) {
					player1=split[p];
					if(player1==null) {
						System.out.println("Can not leave player1 name blank");
						cnSc.close();
						return -1;
					}
				}else if(p==1) {
					player2=split[p];
					if(player2==null) {
						System.out.println("Can not leave player2 name blank");
						cnSc.close();
						return -1;
					}
				}else if(p==2) {
					player3=split[p];
					if(player3==null) {
						System.out.println("Can not leave player3 name blank");
						cnSc.close();
						return -1;
					}
				}else if(p==3) {
					player4=split[p];
					if(player1==null) {
						System.out.println("Can not leave player4 name blank");
						cnSc.close();
						return -1;
					}
				}
			}
		}
		cnSc.close();
		return 1;
	}
	
	int init() {
		if(gotoGroup(tee_time)==-1) {
			System.out.println("Could not find teet time");
			return -1;
		}
		int editingRow = gotoGroup(tee_time);		
		//if (new File(scorecardExcelName).exists()==false) {
			try
			{
				FileInputStream fileIn= new FileInputStream(new File(scorecardExcelName));
				Workbook wb = new HSSFWorkbook (fileIn);
				Sheet ws = wb.getSheet(bookingDate);
				Row row1 = ws.getRow(editingRow);
				Cell c = row1.getCell(2);
				if(c!=null) {
					System.out.println("Group already created.");
					wb.close();
					return -1;
				}
				if(getPlayerName(player_number)!=1) {
					wb.close();
					return -1;
				}
				group_startline=gotoGroup(tee_time);
				if(player_number==1) {
					Cell c1 = row1.createCell(2);
					c1.setCellType(CellType.STRING);
					c1.setCellValue(player1);
				}else if(player_number==2) {
					Cell c1 = row1.createCell(2);
					c1.setCellType(CellType.STRING);
					c1.setCellValue(player1);
					///
					Row row2 = ws.getRow(editingRow+1);
					Cell c2 = row2.createCell(2);
					c2.setCellType(CellType.STRING);
					c2.setCellValue(player2);
				}else if(player_number==3) {
					Cell c1 = row1.createCell(2);
					c1.setCellType(CellType.STRING);
					c1.setCellValue(player1);
					///
					Row row2 = ws.getRow(editingRow+1);
					Cell c2 = row2.createCell(2);
					c2.setCellType(CellType.STRING);
					c2.setCellValue(player2);
					///
					Row row3 = ws.getRow(editingRow+2);
					Cell c3 = row3.createCell(2);
					c3.setCellType(CellType.STRING);
					c3.setCellValue(player3);
				}else if(player_number==4) {
					///
					Cell c1 = row1.createCell(2);
					c1.setCellType(CellType.STRING);
					c1.setCellValue(player1);
					///
					Row row2 = ws.getRow(editingRow+1);
					Cell c2 = row2.createCell(2);
					c2.setCellType(CellType.STRING);
					c2.setCellValue(player2);
					///
					Row row3 = ws.getRow(editingRow+2);
					Cell c3 = row3.createCell(2);
					c3.setCellType(CellType.STRING);
					c3.setCellValue(player3);
					///
					Row row4 = ws.getRow(editingRow+3);
					Cell c4 = row4.createCell(2);
					c4.setCellType(CellType.STRING);
					c4.setCellValue(player4);
					///
				}else {
					System.out.println("The number of players of your group does not meet course requirements");
					fileIn.close();
					FileOutputStream fileOut = new FileOutputStream(new File(scorecardExcelName));
				    wb.write(fileOut);
				    fileOut.close();
				    wb.close();
					return -1;
				}
				fileIn.close();
				FileOutputStream fileOut = new FileOutputStream(new File(scorecardExcelName));
			    wb.write(fileOut);
			    fileOut.close();
			    //System.out.println("Scorecard prepared for your group.");
			    wb.close();
			    return 1;
			}catch(Exception e){
				e.printStackTrace();
				System.out.println("Fail to initialize the group");
				return -1;
			}
		//}
	}
	
	int gotoGroup(String teeTime) {
		try 
		{
			String cell_value = null;
			int gr=1;
			int ln = ROW_FIRST_GROUP ;
			FileInputStream fileIn= new FileInputStream(new File(scorecardExcelName));
			Workbook wb = new HSSFWorkbook (fileIn);
			Sheet ws = wb.getSheet(bookingDate);
			Row row = ws.getRow(ROW_FIRST_GROUP ); //the first group starts at 5
			Cell cell = row.getCell(0);
			if( cell.getCellType() == CellType.STRING ) {
		    	cell_value = cell.getStringCellValue();
		    }
			while ( !(cell_value) .equals(teeTime)) {
				gr++;
				if(gr>GROUP_TOTAL) {
					//System.out.println("Group number does not exist.");
					gr=0;
					wb.close();
					return -1;
				}
				ln=ln+6;
				row = ws.getRow(ln);
				if (row == null) {
					System.out.print("row not initialized for searching group number");
					wb.close();
					return -1;
				}
			    cell = row.getCell(0);
			    if (cell == null) {
			        System.out.println("cell not initialized for searching group number");
			        wb.close();
			        return -1;
			    }
			    if( cell.getCellType() == CellType.STRING ) {
			    	cell_value = cell.getStringCellValue();
			    }
			}
			//System.out.printf("Your tee time at %s is the %d th group of the day.\n",teeTime,gr);
			//System.out.printf("Your group is at line %d \n",ln);
			fileIn.close();
			wb.close();
			return ln;
		}catch(Exception e){
			e.printStackTrace();
			System.out.println("Fail to read cell");
			return -1;
		}
	}

	int getPlayerNameLine(String playerName) {
		group_startline=gotoGroup(tee_time);
		int player_line = group_startline;
		String cell_player = null;
		try
		{
			FileInputStream fileIn= new FileInputStream(new File(scorecardExcelName));
			Workbook wb = new HSSFWorkbook (fileIn);
			Sheet ws = wb.getSheet(bookingDate);//which sheet to modify: counting from 0
			for(int p=0;p<4;p++) {
				Row r = ws.getRow(player_line+p);
				Cell c = r.getCell(2);
				if( c==null ) {
					System.out.println("No player's name on the scorecard");
					wb.close();
					return -1;
				}
				if( c.getCellType() == CellType.STRING ) {
			    	cell_player = c.getStringCellValue();
			    	if(cell_player.equals(playerName)) {
			    		player_line = player_line+p;
			    		//System.out.print(group_startline+p);
			    		wb.close();
			    		return player_line;
			    	}
			    }
			}
			System.out.println("Can not find player");
			fileIn.close();
			wb.close();
			return -1;
		}catch(Exception e){
			e.printStackTrace();
			System.out.println("Fail to read player's line");
			return -1;
		}
	}
	
	void uploadPersonalScore (String playerName,int hole, int score) {
		try 
		{
			FileInputStream fileIn = new FileInputStream(new File(scorecardExcelName));
			Workbook wb = new HSSFWorkbook (fileIn);
			Sheet ws = wb.getSheet(bookingDate);
			group_startline=gotoGroup(tee_time);
			if(group_startline==0) {
				wb.close();
				return;
			}
			Row r = ws.getRow(getPlayerNameLine(playerName));
			Cell c = r.getCell(CELL_FIRST_HOLE -1 + hole);
			if(c==null) {
				c=r.createCell(CELL_FIRST_HOLE -1 + hole);
			}else {
				System.out.printf("par of player %s at hole %d is already uploaded\n",playerName, hole);
				wb.close();
				return;
			}
			c.setCellType(CellType.STRING);
			c.setCellValue(Integer.toString(score));
			///
			Row tr = ws.getRow(getPlayerNameLine(playerName)+4);
			Cell tc = tr.createCell(CELL_FIRST_HOLE -1 + hole);
			tc.setCellType(CellType.STRING);
			tc.setCellValue(getCurrentTime());
			///
		    fileIn.close();
		    FileOutputStream fileOut = new FileOutputStream(new File(scorecardExcelName));
		    wb.write(fileOut);
		    wb.close();
		    fileOut.close();
		    System.out.printf("Par of player %s at hole %d uploaded successfully.\n",playerName, hole);
		    holes_thru++;
		    if(holes_thru==18) {
		    	System.out.println("18 holes done");
		    	saveScorestoProfile(playerName);
		    }
		    getLocation();
		    return;
		}catch(Exception e){
			e.printStackTrace();
			System.out.println("Fail to upload par to scorecard excel");
			return;
		} 
	}

	int getLocation() {
		location=-1;
		try 
		{
			FileInputStream fileIn = new FileInputStream(new File(scorecardExcelName));
			Workbook wb = new HSSFWorkbook (fileIn);
			Sheet ws = wb.getSheet(bookingDate);
			group_startline=gotoGroup(tee_time);
			if(group_startline==0) {
				wb.close();
				return -1;
			}
			Row r = ws.getRow(group_startline);
			for(int c=23;c>4;c--) {
				Cell ce = r.getCell(c);
				if(ce!=null) {
					location= c-4;
					System.out.printf("The group is moving to hole %d now\n", location+1);
					wb.close();
					return location;
				}
			}
		    fileIn.close();
		    FileOutputStream fileOut = new FileOutputStream(new File(scorecardExcelName));
		    wb.write(fileOut);
		    fileOut.close();
		    wb.close();
		    return location;
		}catch(Exception e){
			e.printStackTrace();
			System.out.println("Fail to get location");
			return -1;
		} 
	}
	/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	///////////////////////////////////////////////to be fixed later/////////////////////////////////////////////////////
	int gotoNextAvailableLine(String pn) {
		int l= RECORDSTARTLINE;
		String profileExcelName = "/Users/apple/Desktop/GolfComfy/sim" + pn + "_profile.xls";
		if (new File(profileExcelName).exists()==true) {
			try {
				FileInputStream fileIn= new FileInputStream(new File(profileExcelName));
				Workbook wb = new HSSFWorkbook (fileIn);
				Sheet ws = wb.getSheet(bookingDate);
				while(ws.getRow(l)!=null) {
					l++;
				}
				//System.out.println(l);
				fileIn.close();
			    FileOutputStream fileOut = new FileOutputStream(new File(profileExcelName));
			    wb.write(fileOut);
			    fileOut.close();
			    wb.close();
			    return l;
			}catch(Exception e) {
				e.printStackTrace();
				System.out.printf("Fail to open %s's profile.\n",pn);
				return -1;
			}
		}else {
			System.out.printf("%s's file does not exist",pn);
			return -1;
		}
	}
	
	int getFinalScore(String pn) {
		int score=0;
		try {
			FileInputStream fileIn = new FileInputStream(new File(scorecardExcelName));
			Workbook wb = new HSSFWorkbook(fileIn);
			Sheet ws = wb.getSheet(bookingDate);
			int scoreLine = getPlayerNameLine(pn);
			Row scoreR = ws.getRow(scoreLine);
			for(int hole=0; hole<18; hole++) {
				Cell ch1 = scoreR.getCell(5+hole);
				if(ch1==null) {
					System.out.println("Missing score(s)");
					return -1;
				}
				String sh1 = ch1.getRichStringCellValue().getString();
				//System.out.println(sh1);
				score = score + Integer.parseInt(sh1);
				//System.out.println(score);
			}
			System.out.printf("%s's final score is %d",pn,score);
			fileIn.close();
		    wb.close();
		    return score;
		}catch(Exception e){
			System.out.println("Fail to get the final score");
			return -1;
		}
		
	}
	
	void saveScorestoProfile(String pn) {
		try {
			String profileExcelName = "/Users/apple/Desktop/GolfComfy/sim" + pn + "_profile.xls";
			FileInputStream fileIn = new FileInputStream(new File(profileExcelName));
			Workbook wb = new HSSFWorkbook(fileIn);
			Sheet ws = wb.getSheet(bookingDate);
			int insertLine = gotoNextAvailableLine(pn);
			Row erow = ws.createRow(insertLine);
			Cell c0 = erow.createCell(0);
			c0.setCellType(CellType.STRING);
			c0.setCellValue(bookingDate);
			Cell c1 = erow.createCell(1);
			c1.setCellType(CellType.STRING);
			c1.setCellValue(courseName);
			/////write scores at cell 2/////
			Cell c2 = erow.createCell(2);
			c2.setCellType(CellType.STRING);
			c2.setCellValue(Integer.toString(getFinalScore(pn)));
			//////////////////////////////
			fileIn.close();
			FileOutputStream fileOut = new FileOutputStream(new File(profileExcelName));
		    wb.write(fileOut);
		    fileOut.close();
		    wb.close();
		}catch(Exception e) {
			e.printStackTrace();
			System.out.printf("Fail to save scores to %s'sprofile\n",pn);
			return;
		}
	}

	String getCurrentTime() {
		SimpleDateFormat SDF = new SimpleDateFormat("HH:mm");
		Calendar c = new GregorianCalendar();
		String holeFinishedTime = SDF.format(c.getTime());
		return holeFinishedTime;
	}

	void setOut() {
		try
		{
			FileInputStream fileIn = new FileInputStream(new File(scorecardExcelName));
			Workbook wb = new HSSFWorkbook(fileIn);
			Sheet ws = wb.getSheetAt(0);
			int l = gotoGroup(tee_time);
			Row r = ws.getRow(l+4);
			Cell c = r.createCell(2);
			c.setCellType(CellType.STRING);
			c.setCellValue("playing");
			fileIn.close();
		    FileOutputStream fileOut = new FileOutputStream(new File(scorecardExcelName));
		    wb.write(fileOut);
		    fileOut.close();
		    wb.close();
		    System.out.printf("A group with %d players is setting out at %s\n",player_number,getCurrentTime());
		    return;
		}catch(Exception e) {
			System.out.println("Fail to start the group");
			return;
		}
	}
	
}
