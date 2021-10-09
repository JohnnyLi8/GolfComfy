

import java.util.Scanner;
import java.io.*;
import java.util.*;

public class loadCourseInfo
{	
	String name;
	int pars[] = new int [18];
	int time[] = new int [18];
	int level;
	
	public loadCourseInfo(String cn) {
		name=cn;
		pars = getPars(name);
		time = getTime(name);
		//level = getDLevel(name);
	}
	
	private Scanner fileScanner;
	
	

			
	public int[] getPars(String courseName) {
		try {
			fileScanner = new Scanner(new File(courseName+"_par.txt"));
			int hole=0;
			while(fileScanner.hasNextInt() && hole<18) {
				pars[hole++] = fileScanner.nextInt();
			}
			fileScanner.close();
			return pars;
		}catch(Exception e) {
			System.out.println("can not find the file containing Pars info");
			return null;
		}
	}
	
	public int[] getTime(String courseName) {
		try {
			fileScanner = new Scanner(new File(courseName+"_time.txt"));
			int hole=0;
			while(fileScanner.hasNextInt() && hole<18) {
				time[hole++] = fileScanner.nextInt();
			}
			fileScanner.close();
			return time;
		}catch(Exception e) {
			System.out.println("can not find the file containing Time info");
			return null;
		}
	}
	
	public int getLevel() {
		try {
			fileScanner = new Scanner(new File("CsDiffRate.txt"));
			while(fileScanner.hasNextLine()) {
				final String line = fileScanner.nextLine();
				if (line.contains(name)) {
					level = Integer.parseInt(fileScanner.nextLine());
					break;
				}
			}
			fileScanner.close();
			return level;
		}catch(Exception e) {
			System.out.println("can not find the file containing difficulty info");
			return -1;
		}
	}
}
