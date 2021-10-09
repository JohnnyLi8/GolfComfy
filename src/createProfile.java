
import java.io.File;
import java.io.FileOutputStream;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class createProfile {
	String playername;
	String profileExcelName;
	
	public createProfile(String pn) {
		playername=pn;
		profileExcelName = "/Users/apple/Desktop/GolfComfy/sim/" + playername + "_profile.xls";
		init();
	}
	
	void init() {
		if (new File(profileExcelName).exists()==false) {
			HSSFWorkbook wb = new HSSFWorkbook();
			Sheet ws = wb.createSheet(playername + " profile");
			Row r0 = ws.createRow(0);
			Cell c0 = r0.createCell(0);
			c0.setCellType(CellType.STRING);
			c0.setCellValue("NAME:");
			Cell c1 = r0.createCell(1);
			c1.setCellType(CellType.STRING);
			c1.setCellValue(playername);
			Row r2 = ws.createRow(2);
			Cell c20 = r2.createCell(0);
			c20.setCellType(CellType.STRING);
			c20.setCellValue("GAMES:");
			Cell c21 = r2.createCell(1);
			c21.setCellType(CellType.STRING);
			c21.setCellValue("0");
			Cell c22 = r2.createCell(2);
			c22.setCellType(CellType.STRING);
			c22.setCellValue("Lates:");
			Cell c23 = r2.createCell(3);
			c23.setCellType(CellType.STRING);
			c23.setCellValue("0");
			Row r3 = ws.createRow(3);
			Cell c30 = r3.createCell(0);
			c30.setCellType(CellType.STRING);
			c30.setCellValue("AVG SCORE:");
			ws.setColumnWidth(0,5000);
			Row r4 = ws.createRow(4);
			Cell c40 = r4.createCell(0);
			c40.setCellType(CellType.STRING);
			c40.setCellValue("AVG STANDARD ON:");
			Row r5 = ws.createRow(5);
			Cell c50 = r5.createCell(0);
			c50.setCellType(CellType.STRING);
			c50.setCellValue("AVG PUTS:");
			Row r7 = ws.createRow(7);
			Cell c70 = r7.createCell(0);
			c70.setCellType(CellType.STRING);
			c70.setCellValue("HISTORY RECORD:");
			Row r8 = ws.createRow(8);
			Cell c80 = r8.createCell(0);
			c80.setCellType(CellType.STRING);
			c80.setCellValue("TIME");
			Cell c81 = r8.createCell(1);
			c81.setCellType(CellType.STRING);
			c81.setCellValue("COURSE");
			Cell c82 = r8.createCell(2);
			ws.setColumnWidth(1, 3900);
			c82.setCellType(CellType.STRING);
			c82.setCellValue("SCORES");
			Cell c83 = r8.createCell(3);
			c83.setCellType(CellType.STRING);
			c83.setCellValue("BIRDIEs");
			Cell c84 = r8.createCell(4);
			c84.setCellType(CellType.STRING);
			c84.setCellValue("PARs");
			Cell c85 = r8.createCell(5);
			c85.setCellType(CellType.STRING);
			c85.setCellValue("BOGEYs");
			Cell c86 = r8.createCell(6);
			c86.setCellType(CellType.STRING);
			c86.setCellValue("DOUBLE BOGEYs");
			Cell c87 = r8.createCell(7);
			c87.setCellType(CellType.STRING);
			c87.setCellValue("STANDARD ONs");
			Cell c88 = r8.createCell(8);
			c88.setCellType(CellType.STRING);
			c88.setCellValue("PUTS");
			Cell c89 = r8.createCell(9);
			c89.setCellType(CellType.STRING);
			c89.setCellValue("OBs");
			ws.setColumnWidth(6,4200);
			ws.setColumnWidth(7, 4000);
			
			try {
				FileOutputStream output = new FileOutputStream(profileExcelName);
				wb.write(output);
				output.close();
				wb.close();
				System.out.printf("%s's profile created and initialized\n", playername);
			}catch(Exception e){
				e.printStackTrace();
				System.out.printf("Fail to create %s's profile\n", playername);
			}
		}else {
			System.out.printf("%s's profile existed/initialized\n", playername);
		}
	}
}
