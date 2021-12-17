package excel.export.data;

import java.util.Map;

public class Student {
	/**
	 * 姓名
	 */
	private String name;
	
	/**
	 * 学生所在班级
	 */
	private ClassRoom classRoom;
	
	/**
	 * 学生的其他信息
	 */
	private Map<String,Object> moreInfo;

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public ClassRoom getClassRoom() {
		return classRoom;
	}

	public void setClassRoom(ClassRoom classRoom) {
		this.classRoom = classRoom;
	}

	public Map<String, Object> getMoreInfo() {
		return moreInfo;
	}

	public void setMoreInfo(Map<String, Object> moreInfo) {
		this.moreInfo = moreInfo;
	}
	
}
