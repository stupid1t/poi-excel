package excel.export;

import java.util.HashMap;
import java.util.Map;

public class  ClassRoom{
	
	/**
	 * 班级名称
	 */
	private String name;
	
	/**
	 * 班级名称
	 */
	private Map<String,Object> school = new HashMap<>();

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}
	
	public ClassRoom(String name) {
		this.name = name;
		getSchool().put("name", "世紀學校");
	}

	public Map<String,Object> getSchool() {
		return school;
	}

	public void setSchool(Map<String,Object> school) {
		this.school = school;
	}
	
}
