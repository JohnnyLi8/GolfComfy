
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Simulation {
	
	void write(String fN, int num_row, int num_cell, String inputString) {
		try 
		{
			String excelFileName = "/Users/apple/Desktop/GolfComfy/sim" + fN + ".xls";
			FileInputStream fileIn = new FileInputStream(new File(excelFileName));
			Workbook wb = new HSSFWorkbook (fileIn);
			Sheet ws = wb.getSheetAt(0);//which sheet to modify: counting from 0
		    
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
		    FileOutputStream fileOut = new FileOutputStream(new File(excelFileName));
		    wb.write(fileOut);
		    fileOut.close();
		}catch(Exception e){
			e.printStackTrace();
			System.out.println("Fail to add string");
		} 
	}
}
