package excel.export;

import java.util.Date;

public class ProjectEvaluate {

	/**
	 * 主键
	 */
	private Long id;

	/**
	 * 项目ID
	 */
	private Long projectId;


	/**
	 * 创建人
	 */
	private Long createUserId;

	/**
	 * 创建时间
	 */
	private Date createTime;


	/**
	 * 项目名称
	 */
	private String projectName;


	/**
	 * 所属区域
	 */
	private String areaName;

	/**
	 * 省份
	 */
	private String province;

	/**
	 * 市
	 */
	private String city;

	/**
	 * 项目状态
	 */
	private String statusName;

	/**
	 * 总分
	 */
	private int scount;

	/**
	 * 历史平均分
	 */
	private double avg;

	/**
	 * 项目所属人
	 */
	private String people;

	/**
	 * 项目领导人
	 */
	private String leader;

	/**
	 * 图片
	 */
	private byte[] img;

	public Long getId() {
		return id;
	}

	public void setId(Long id) {
		this.id = id;
	}

	public Long getProjectId() {
		return projectId;
	}

	public void setProjectId(Long projectId) {
		this.projectId = projectId;
	}

	public Long getCreateUserId() {
		return createUserId;
	}

	public void setCreateUserId(Long createUserId) {
		this.createUserId = createUserId;
	}

	public Date getCreateTime() {
		return createTime;
	}

	public void setCreateTime(Date createTime) {
		this.createTime = createTime;
	}

	public String getProjectName() {
		return projectName;
	}

	public void setProjectName(String projectName) {
		this.projectName = projectName;
	}


	public String getAreaName() {
		return areaName;
	}

	public void setAreaName(String areaName) {
		this.areaName = areaName;
	}

	public String getProvince() {
		return province;
	}

	public void setProvince(String province) {
		this.province = province;
	}

	public String getCity() {
		return city;
	}

	public void setCity(String city) {
		this.city = city;
	}

	public String getStatusName() {
		return statusName;
	}

	public void setStatusName(String statusName) {
		this.statusName = statusName;
	}

	public byte[] getImg() {
		return img;
	}

	public void setImg(byte[] img) {
		this.img = img;
	}

	public int getScount() {
		return scount;
	}

	public void setScount(int scount) {
		this.scount = scount;
	}

	public double getAvg() {
		return avg;
	}

	public void setAvg(double avg) {
		this.avg = avg;
	}

	public String getPeople() {
		return people;
	}

	public void setPeople(String people) {
		this.people = people;
	}

	public String getLeader() {
		return leader;
	}

	public void setLeader(String leader) {
		this.leader = leader;
	}

}
