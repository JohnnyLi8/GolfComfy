
import java.io.IOException;

public class main {
	
	public static void main(String[] args) {
		/////////////////////must call at first//////////////////////////////////
		System.out.println("Starting Golf App\n");	
		/////////////////////////////////////////////////////////////////////////
		
		Group group1 = new Group("Flamborough Hills","Jan 02","06:30",1,18);
		Group group2 = new Group("Flamborough Hills","Jan 02","06:40",1,18);
		
		/////////////////////////////////////////////////////////////////////////
		
		group1.setOut();
		group2.setOut();
		group1.uploadPersonalScore("p1", 1, 4);
		group2.uploadPersonalScore("p2", 9, 4);
		
		/////////////////////////////////////////////////////////////////////////
		
		Admin admin = new Admin("Flamborough Hills");
		admin.displayTraffic(); //admin.detectGroupsLocation();
		admin.checkTraffic();
		admin.markPlayerLate("p1");
		
		/////////////////////////////////////////////////////////////////////////
	}
}