package Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {
	//XLSX 파일을 분석하여 List<exData>객체로 반환
	public List<ExData> xlsxToList(String filePath){
		//반환할 객체 생성
		List<ExData> list = new ArrayList<ExData>();
		
		FileInputStream fis = null;
		XSSFWorkbook workbook = null;
		
		try {
			fis = new FileInputStream(filePath); //파일경로
			workbook = new XSSFWorkbook(fis); //엑셀파일 전체내용 담고 있는 객체
			
			//탐색에 사용할 객체
			XSSFSheet curSheet;
			XSSFRow curRow;
			XSSFCell curCell;
			ExData exData;
			
			//sheet 탐색 for문
			for(int sheetIndex=0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
				curSheet = workbook.getSheetAt(sheetIndex); //현재 sheet반환
				//row탐색
				for(int rowIndex=0; rowIndex < curSheet.getPhysicalNumberOfRows(); rowIndex++) {
					if(rowIndex != 0) { //row 0 은 헤더 정보이기 때문에 무시
						curRow = curSheet.getRow(rowIndex); //현재 row
						exData = new ExData();
						String value;
						
						//row의 첫번째 cell값이 비어있지 않은 경우만 탐색
						if(curRow!=null) {
							for(int cellIndex=0; cellIndex < curRow.getPhysicalNumberOfCells(); cellIndex++) {
								curCell = curRow.getCell(cellIndex); //현재 cell
								
								if(true) {
									value = "";
									//cell 스타일이 다르더라도 String으로 반환 받음
									switch(curCell.getCellType()) {
									case FORMULA:
										value = curCell.getCellFormula();
										break;
									case NUMERIC:
										value = curCell.getNumericCellValue()+"";
										break;
									case STRING:
										value = curCell.getStringCellValue()+"";
										break;
									case BLANK:
										value = curCell.getBooleanCellValue()+"";
										break;
									case ERROR:
										value = curCell.getErrorCellValue()+"";
										break;
									default:
										value = new String();
										break;
									}
									
									//현재 column index에 따라 exData에 입력
									switch(cellIndex) {
									case 0: //순번
										exData.setNum(value);
										break;
										
									case 1: //이름
										exData.setName(value);
										break;
										
									case 2: //전화번호
										exData.setPhoneNum(value);
										break;
										
									default:
										break;
									}
								}
							}
							//cell 탐색 후 exData 추가 
							list.add(exData);
						}
					}
				}
			}
		}catch(FileNotFoundException e) {
			e.printStackTrace();
		}catch(IOException e) {
			e.printStackTrace();
		}finally {
			try {
				//사용한 자원은 finally에서 해제
				if(workbook != null) workbook.close();
				if(fis != null) fis.close();
			}catch(IOException e) {
				e.printStackTrace();
			}
		}
		return list;
	}
}
