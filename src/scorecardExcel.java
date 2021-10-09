
import java.io.File;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.Scanner;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

//constructor class by taking in players names 
public class scorecardExcel {
	final int week=7;
	String day[] = new String [week];
	int current_row;
	String todayDate;
	final int GROUP_TOTAL = 64; 
	int pars[] = new int [18];
	String courseName;
	int time[] = new int[18];
	String scorecareExcelName;
	
	public scorecardExcel(String course) {
		courseName = course;
		scorecareExcelName = "/Users/apple/Desktop/GolfComfy/sim/" + courseName + "_Scorecard.xls";
		init();
	}
	
	//create and init
	void init () {
		if (new File(scorecareExcelName).exists()==false) {
			HSSFWorkbook wb = new HSSFWorkbook();
			//
			for (int i = 0; i<week; i++) {
				SimpleDateFormat sdf = new SimpleDateFormat("MMM dd");
				Calendar calendar = new GregorianCalendar();
				calendar.add(Calendar.DATE, i);
				day[i] = sdf.format(calendar.getTime());
			}
			//
			for (int d = 0; d<week; d++) {
				SimpleDateFormat sdf = new SimpleDateFormat("MMM dd");
				Calendar calendar = new GregorianCalendar();
				todayDate = sdf.format(calendar.getTime());
				Sheet sheet = wb.createSheet(day[d]);
				Row table_row = sheet.createRow(0);
				//		
				Row row0 = sheet.createRow(0);
				Cell tt_cell = row0.createCell(0);
			    tt_cell.setCellType(CellType.STRING);
				tt_cell.setCellValue("Tee Time");
				//sheet.setColumnWidth(0,3300);
				//
				Cell groupNum_cell = row0.createCell(1);
			    groupNum_cell.setCellType(CellType.STRING);
				groupNum_cell.setCellValue("Group Number");
				sheet.setColumnWidth(1,3300);
				//
			    Cell playerName_cell = row0.createCell(2);
			    playerName_cell.setCellType(CellType.STRING);
			    playerName_cell.setCellValue("Players Name");
			    sheet.setColumnWidth(2,3500);
			    //
			    Row row1 = sheet.createRow(1);
			    Cell p_cell = row1.createCell(4);
			    p_cell.setCellType(CellType.STRING);
			    p_cell.setCellValue("PAR");
			    //
			    for (int i=0;i<18;i++) {
			    	Cell hole_cell = row0.createCell(5+i);
		    		hole_cell.setCellType(CellType.STRING);
		    		hole_cell.setCellValue("Hole "+Integer.toString(i+1));
			    }
			    //
			    loadCourseInfo courseInfo = new loadCourseInfo(courseName);
				pars = courseInfo.getPars(courseName);
				time = courseInfo.getTime(courseName);
				//init par of each hole
				if( pars!=null) {
				    for (int i=0;i<18;i++) {
				    	Cell par_cell = row1.createCell(5+i);
			    		par_cell.setCellType(CellType.STRING);
			    		par_cell.setCellValue(pars[i]);
				    }    
				}else {
					return;
				}
			    //init suggested play time of each hole
			    Row row2 = sheet.createRow(2);
			    Cell h_time = row2.createCell(4);
			    h_time.setCellType(CellType.STRING);
			    h_time.setCellValue("TIME");
			    for (int i=0;i<18;i++) {
			    	Cell time_cell = row2.createCell(5+i);
		    		time_cell.setCellType(CellType.STRING);
		    		time_cell.setCellValue(time[i]);
			    }
			    //init group number
			    for (int gn=1; gn<GROUP_TOTAL+1;gn++) {
			    	for(int i=0;i<4;i++) {
			    		Row r = sheet.createRow(6*gn+i);
			    		Cell c = r.createCell(1);
			    		c.setCellType(CellType.STRING);
			    		c.setCellValue(Integer.toString(gn));
			    	}
			    }
			    //
			    SimpleDateFormat SDF = new SimpleDateFormat("HH:mm");
				Calendar c = new GregorianCalendar();
				c.set(c.HOUR_OF_DAY,6);
				c.set(c.MINUTE,30);
				String teeTime = SDF.format(c.getTime());
				for (int gn=1; gn<GROUP_TOTAL+1;gn++) {
					teeTime = SDF.format(c.getTime());
					c.add(c.MINUTE, 10);
			    	for(int i=0;i<4;i++) {
			    	Row gr_row = sheet.createRow(6*gn+i);
				    Cell gr_cell = gr_row.createCell(1);
				    gr_cell.setCellType(CellType.STRING);
				    gr_cell.setCellValue("#"+Integer.toString(gn));
					Row tt_row = sheet.getRow(6*gn+i);
				    Cell t_cell = tt_row.createCell(0);
				    t_cell.setCellType(CellType.STRING);
				    t_cell.setCellValue(teeTime);
			    	}
			    	Row tm_row = sheet.createRow(6*gn+4);
			    	Cell tm_cell = tm_row.createCell(0);
			    	tm_cell.setCellType(CellType.STRING);
			    	tm_cell.setCellValue("Status:");
				}
			}
		    
		    //
			try {
				FileOutputStream output = new FileOutputStream(scorecareExcelName);
				wb.write(output);
				output.close();
				wb.close();
				//System.out.println("Scorecard Excel file created and initialized");
				System.out.println();
			}catch(Exception e){
				e.printStackTrace();
				System.out.println("Fail to create Excel");
			}
		}//else {
			//System.out.println("Scorecard Excel file existed/initialized.");
			//System.out.println();
		//}
	}
}