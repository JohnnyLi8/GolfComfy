
import java.util.Scanner;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;


import java.time.LocalDate;
import java.time.Month;
import java.text.DateFormatSymbols;
import java.text.Format;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;

public class bookingExcel {
	
	String bookingDate;
	String firstDate;
	final int week = 7;
	String day[] = new String [week];
	final int group_total = 64; 
	public String courseName;
	String bookingExcelName;
	
	public bookingExcel(String course) {
		courseName = course;
		bookingExcelName = "/Users/apple/Desktop/GolfComfy/sim/" + courseName + "_Booking.xls";
		getBookingDatesExcelSheets();
	}

	
	String getBookingDate() {
		Scanner bkSc = new Scanner(System.in);
		System.out.println("Which date do you want to book?");
		bookingDate = bkSc.nextLine();
		bkSc.close();
		return bookingDate;
	}
	
	
/////////////////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////////////////	

	//check if the booking excel file exists, if not, generate one
	void getBookingDatesExcelSheets () {
		if (new File(bookingExcelName).exists()==false) {
			HSSFWorkbook wb = new HSSFWorkbook();
			
			//store the generated days(in string) in array for later use
			for (int i = 0; i<week; i++) {
				SimpleDateFormat sdf = new SimpleDateFormat("MMM dd");
				Calendar calendar = new GregorianCalendar();
				calendar.add(Calendar.DATE, i);
				day[i] = sdf.format(calendar.getTime());
			}
			
			//set the first tee time to be at 6:30
			SimpleDateFormat sdf = new SimpleDateFormat("HH:mm");
			Calendar calendar = new GregorianCalendar();
			calendar.set(calendar.HOUR_OF_DAY,6);
			calendar.set(calendar.MINUTE,30);
			String teeTime = sdf.format(calendar.getTime());
			
			//generate 7 sheets for each day
			for (int i = 0; i<week; i++) {
				calendar.set(calendar.HOUR_OF_DAY,6);
				calendar.set(calendar.MINUTE,30);
				teeTime = sdf.format(calendar.getTime());
				Sheet sheet = wb.createSheet(day[i]);
				sheet.setColumnWidth(1,3300);
				Row table_row = sheet.createRow(0);
				//
			    Cell groupNum_cell = table_row.createCell(1);
			    groupNum_cell.setCellType(CellType.STRING);
			    groupNum_cell.setCellValue("Group Number");
			    //
			    Cell teeTime_cell = table_row.createCell(0);
			    teeTime_cell.setCellType(CellType.STRING);
			    teeTime_cell.setCellValue("Tee Time");
			    //
			    Cell bookingStatus_cell = table_row.createCell(2);
			    bookingStatus_cell.setCellType(CellType.STRING);
			    bookingStatus_cell.setCellValue("Booking Status");
			    sheet.setColumnWidth(2,3300);
			    //
			    Cell num_ppl_cell = table_row.createCell(3);
			    num_ppl_cell.setCellType(CellType.STRING);
			    num_ppl_cell.setCellValue("Number of Players");
			    sheet.setColumnWidth(3,3800);
			    //
			    Cell num_holes_cell = table_row.createCell(4);
			    num_holes_cell.setCellType(CellType.STRING);
			    num_holes_cell.setCellValue("Number of Holes");
			    sheet.setColumnWidth(4,3700);
			    //
			    //initialize the tee time table for each sheet
			    for (int t = 0; t < group_total; t++) {
			    	Row gr_row = sheet.createRow(t+1);
				    Cell gr_cell = gr_row.createCell(1);
				    gr_cell.setCellType(CellType.STRING);
				    gr_cell.setCellValue(Integer.toString(t+1));
			    	teeTime = sdf.format(calendar.getTime());
					calendar.add(calendar.MINUTE, 10);
					//System.out.println(teeTime);
					Row tt_row = sheet.getRow(t+1);
				    Cell tt_cell = tt_row.createCell(0);
				    tt_cell.setCellType(CellType.STRING);
				    tt_cell.setCellValue(teeTime);
			    }
			}
			try {
				FileOutputStream output = new FileOutputStream(bookingExcelName);
				wb.write(output);
				wb.close();
				output.close();
				//System.out.println("\nBooking Excel file created and initialized");
			}catch(Exception e){
				e.printStackTrace();
				System.out.println("Fail to create Excel");
			}
		}//else {
			//System.out.println("\nBooking Excel file existed/initialized.");
		//}
	}
	
	//create a new sheet to when it's a new day
	void updateExcelSheet() {
		try (FileInputStream fileIn = new FileInputStream(bookingExcelName))
		{
			Workbook wb1 = WorkbookFactory.create (fileIn);
			String firstDate = wb1.getSheetName(0); //get the date on the first Excel sheet
			//get the date in "MM dd" format
			SimpleDateFormat sdf = new SimpleDateFormat("MMM dd");
			Calendar calendar = new GregorianCalendar();
			String todayDate = sdf.format(calendar.getTime());
			calendar.add(calendar.DATE, 7);
			String newDate = sdf.format(calendar.getTime());
			//
			if( ! (firstDate.equals(todayDate)) ) {
				//set the first tee time at "6:30"
			    SimpleDateFormat SDF = new SimpleDateFormat("HH:mm");
				Calendar cal = new GregorianCalendar();
				cal.set(cal.HOUR_OF_DAY,6);
				cal.set(cal.MINUTE,30);
				String teeTime = SDF.format(cal.getTime());
				//
				Sheet sheet = wb1.createSheet(newDate);   //create a new sheet and initialize it
				sheet.setColumnWidth(1,3300);
				Row table_row = sheet.createRow(0);
			    Cell groupNum_cell = table_row.createCell(1);
			    groupNum_cell.setCellType(CellType.STRING);
			    groupNum_cell.setCellValue("Group Number");
			    Cell teeTime_cell = table_row.createCell(0);
			    teeTime_cell.setCellType(CellType.STRING);
			    teeTime_cell.setCellValue("Tee Time");
			    Cell bookingStatus_cell = table_row.createCell(2);
			    bookingStatus_cell.setCellType(CellType.STRING);
			    bookingStatus_cell.setCellValue("Booking Status");
			    sheet.setColumnWidth(2,3300);
			    //initialize the tee time table for each sheet
			    for (int t = 0; t < group_total; t++) {
			    	Row gr_row = sheet.createRow(t+1);      //use "createRow" because the sheet was created just now
				    Cell gr_cell = gr_row.createCell(1);    //same reason
				    gr_cell.setCellType(CellType.STRING);
				    gr_cell.setCellValue(Integer.toString(t+1));
			    	teeTime = SDF.format(cal.getTime());
					cal.add(cal.MINUTE, 10);
					//System.out.println(teeTime);
					Row tt_row = sheet.getRow(t+1);
				    Cell tt_cell = tt_row.createCell(0);
				    tt_cell.setCellType(CellType.STRING);
				    tt_cell.setCellValue(teeTime);
			    }
			    System.out.println("A new sheet is automatically created");
			    fileIn.close();
			    FileOutputStream fileOut = new FileOutputStream(new File(bookingExcelName));
			    wb1.write(fileOut);
			    fileOut.close();
			}else {
				System.out.println("No sheets needed to be update");
				fileIn.close();
			}
		}catch(Exception e){
			e.printStackTrace();
		} 
	}
	
/////////////////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////////////////
	void write(String bookingDate, int num_row, int num_cell, String inputString) {
		try 
		{
			FileInputStream fileIn = new FileInputStream(new File(bookingExcelName));
			Workbook wb = new HSSFWorkbook (fileIn);
			Sheet ws = wb.getSheet(bookingDate);//which sheet to modify: counting from 0
		    
			Row row = ws.getRow(num_row);
			if (row == null) {
				row = ws.createRow(num_row);
			}
			
		    Cell cell = row.getCell(num_cell);
		    if (cell == null) {
		        cell = row.createCell(num_cell);
		        cell.setCellType(CellType.STRING);
		        cell.setCellValue(inputString);
		    }else { //to overwrite the original value
		    	String ori_val = cell.getStringCellValue();
		    	cell.setCellType(CellType.STRING);
		    	cell.setCellValue(inputString);
		    	System.out.printf("The cell value '%s' is changed into '%s'. \n",ori_val,inputString);
		    }
		    //revise the codes so that it updates the value of cell if the cell exists
		    fileIn.close();
		    FileOutputStream fileOut = new FileOutputStream(new File(bookingExcelName));
		    wb.write(fileOut);
		    fileOut.close();
		    wb.close();
		}catch(Exception e){
			e.printStackTrace();
			System.out.println("Fail to add string");
		} 
	}
	
	//read the cell value, assuming it's type string
	String read(String bookingDate,int num_row, int num_cell) {
		String cell_value = null;
		try 
		{
			FileInputStream fileIn= new FileInputStream(new File(bookingExcelName));
			Workbook wb = new HSSFWorkbook (fileIn);
			Sheet ws = wb.getSheet(bookingDate);//which sheet to modify: counting from 0
		    
			Row row = ws.getRow(num_row);
			if (row == null) {
				System.out.print("row null");
			}
			
		    Cell cell = row.getCell(num_cell);
		    if (cell == null) {
		        System.out.println("cell null");
		    }
		    if( cell.getCellType() == CellType.STRING ) {
		    	cell_value = cell.getStringCellValue();
		    	System.out.printf("The cell value is: '%s'.\n",cell_value);
		    }else {
		    	System.out.println("The cell type is not String, can not read");
		    }
		    fileIn.close();
		    wb.close();
		}catch(Exception e){
			e.printStackTrace();
			System.out.println("Fail to read cell");
		}
		return cell_value;
	}
	
	//search for a tee time and return its group number
	int getGroupNumber(String bookingDate, String teeTime) {
		try 
		{
			String cell_value = null;
			int gr=1;
			FileInputStream fileIn= new FileInputStream(new File(bookingExcelName));
			Workbook wb = new HSSFWorkbook (fileIn);
			Sheet ws = wb.getSheet(bookingDate);//which sheet to modify: counting from 0
			if(ws==null) {
				return -1;
			}
			Row row = ws.getRow(1);
			Cell cell = row.getCell(0);
			if( cell.getCellType() == CellType.STRING ) {
		    	cell_value = cell.getStringCellValue();
		    }
			while ( !(cell_value) .equals(teeTime) ) {
				if(gr>63) {
					//System.out.println("Tee time does not meet the course requirements.");
					return -1;
				}
				gr++;
				row = ws.getRow(gr);
				if (row == null) {
					System.out.print("row not initialized for searching tee time");
				}
			    cell = row.getCell(0);
			    if (cell == null) {
			        System.out.println("cell not initialized for searching tee time");
			    }
			    if( cell.getCellType() == CellType.STRING ) {
			    	cell_value = cell.getStringCellValue();
			    }
			}
			//System.out.printf("Your tee time at %s is the %d th group of the day.\n",teeTime,gr);
			fileIn.close();
			wb.close();
			return gr;
		}catch(Exception e){
			e.printStackTrace();
			System.out.println("Fail to read cell");
			return -1;
		}
	}
	
	//check if the tee time is available, will return either true (available) or false (booked)
	int isTeeTimeAvailable(String bookingDate,String teeTime) {
		String cell_value = null;
		try 
		{
			int gr = getGroupNumber(bookingDate, teeTime);
			if(gr!=-1) {
				FileInputStream fileIn= new FileInputStream(new File(bookingExcelName));
				Workbook wb = new HSSFWorkbook (fileIn);
				Sheet ws = wb.getSheet(bookingDate);//which sheet to modify: counting from 0
				int bdexist=0;
				for(int day=0; day<week;day++) {
					if((bookingDate.equals(wb.getSheetName(day)))) bdexist=1;
				}
				if(bdexist==0)return-1;
				Row r = ws.getRow(gr);
				Cell c = r.getCell(2);
				if(c==null) {
					//System.out.println("isTeeTimeAvailable(): true.");
					//System.out.printf("Tee time available at %s and your group number is %d.\n",teeTime,gr);
					fileIn.close();
					wb.close();
					return 1;
				}
				if ( (c.getStringCellValue()).equals("booked")) {
					//System.out.println("isTeeTimeAvailable(): false.");
					fileIn.close();
					wb.close();
					return 2;
				}else {
					//System.out.println("isTeeTimeAvailable(): true.");
					//System.out.printf("Tee time available at %s and your group number is %d.\n",teeTime,gr);
					fileIn.close();
					wb.close();
					return 1;
				}
				
			}
		}catch(Exception e){
			e.printStackTrace();
			System.out.println("Fail to read cell");
			return -1;
		}
		return -1;
	}
	
	//book a group, by writing booked to the cell
	int book(String bookingDate, String teeTime, int n_players, int n_holes) {
		try 
		{
			int gr = getGroupNumber(bookingDate,teeTime);
			if(gr==-1)return -1;
			if(isTeeTimeAvailable(bookingDate,teeTime)==1) {
				//write "booked" to the booking status column
				//groupScorecard group = new groupScorecard(courseName,teeTime,n_players);
				write(bookingDate,gr,2, "booked");
				write(bookingDate,gr,3,Integer.toString(n_players));
				write(bookingDate,gr,4,Integer.toString(n_holes));
				System.out.println("Tee time now booked for you: "+bookingDate+", "+teeTime+".");
				System.out.println("Scorecard also prepared for you.");
				return 1;
			}else {
				//System.out.println("Tee time not available, can not be booked.");
				return 2;
			}
		}catch(Exception e){
			e.printStackTrace();
			System.out.println("Fail to add string");
		}
		return 1; 
			
	}
	

	
	

}

