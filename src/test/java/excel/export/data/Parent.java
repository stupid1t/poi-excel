package excel.export.data;

public class Parent {
	
	/**
	 * 姓名
	 */
	private String name;

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public int getAge(){
		return 3;
	}
	
	public Parent(String name) {
		this.name = name;
	}
}
