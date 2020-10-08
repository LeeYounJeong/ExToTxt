package Test;

public class ExData { //엑셀 파일의 칼럼 항목과 동일하게 구성
	private String num; //순번
	private String name; //이름
	private String phoneNum; //전화번호
	
	
	public String getNum() {
		return num;
	}


	public void setNum(String num) {
		this.num = num;
	}


	public String getName() {
		return name;
	}


	public void setName(String name) {
		this.name = name;
	}


	public String getPhoneNum() {
		return phoneNum;
	}


	public void setPhoneNum(String phoneNum) {
		this.phoneNum = phoneNum;
	}

	@Override
	public String toString() { 
		double no = Double.parseDouble(num);
		int iNum;
		iNum = (int)no;
		StringBuffer sb = new StringBuffer();
		
		sb.append(iNum);
		sb.append("," + name);
		sb.append("," + phoneNum);
		return sb.toString();
	}
}
