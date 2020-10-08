package Test;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.util.List;

public class MainEx {
	public static void main(String[] args) {
		ExcelReader excelReader = new ExcelReader();
		
		List<ExData> xlsxList = excelReader.xlsxToList("C:\\Users\\TA공용\\Desktop\\Test.xlsx");
		printList(xlsxList);
		
	}
	
	public static void printList(List<ExData> list) {
		ExData exData;
		//output
		try {
			File file = new File("C:\\Users\\TA공용\\Desktop\\TestOutput.txt");
			BufferedWriter bufferedWriter = new BufferedWriter(new FileWriter(file));
			
			if(file.isFile() && file.canWrite()) {
				for(int i=0; i< list.size(); i++) {
					exData = list.get(i);
					String txt = exData.toString();

					bufferedWriter.write(txt);
					bufferedWriter.newLine();
					
					System.out.println(txt);
				}
				bufferedWriter.close();
			}
			
			
		}catch(Exception e) {
			e.getStackTrace();
		}
	}
}
