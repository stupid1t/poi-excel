package excel.export;

import java.io.Serializable;
import java.util.Date;

/**
 * 描述 整体评价
 * 
 * @author 谢宇
 * @version 2017-09-30
 */
public class ProjectEvaluate implements Serializable {

	private static final long serialVersionUID = 1L;

	/**
	 * 主键
	 */
	private Long id;

	/**
	 * 项目ID
	 */
	private Long projectId;

	/**
	 * 区位条件
	 */
	private String areaInfo;

	/**
	 * 资源禀赋
	 */
	private String resourceInfo;

	/**
	 * 经营现状
	 */
	private String manageInfo;

	/**
	 * 考察印象
	 */
	private String reviewInfo;

	/**
	 * 管理团队
	 */
	private String teamInfo;

	/**
	 * 成长潜力
	 */
	private String potentialInfo;

	/**
	 * 创建人
	 */
	private Long createUserId;

	/**
	 * 创建时间
	 */
	private Date createTime;

	/**
	 * 区位分数
	 */
	private Double areaScore;

	/**
	 * 资源分数
	 */
	private Double resourceScore;

	/**
	 * 经营分数
	 */
	private Double manageScore;

	/**
	 * 考察分数
	 */
	private Double reviewScore;

	/**
	 * 管理分数
	 */
	private Double teamScore;

	/**
	 * 成长分数
	 */
	private Double potentialScore;

	/**
	 * 项目名称
	 */
	private String projectName;

	/**
	 * 导出参数
	 */
	private String ids;

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
	private String scount;

	/**
	 * 关键词
	 */
	private String keyWords;
	/**
	 * 标记是否为当前数据:1当前数据,0历史数据
	 */
	private Integer IsCurrent;

	/**
	 * 流程ID
	 */
	private Long flowId;

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

	public String getAreaInfo() {
		return areaInfo;
	}

	public void setAreaInfo(String areaInfo) {
		this.areaInfo = areaInfo;
	}

	public String getResourceInfo() {
		return resourceInfo;
	}

	public void setResourceInfo(String resourceInfo) {
		this.resourceInfo = resourceInfo;
	}

	public String getManageInfo() {
		return manageInfo;
	}

	public void setManageInfo(String manageInfo) {
		this.manageInfo = manageInfo;
	}

	public String getReviewInfo() {
		return reviewInfo;
	}

	public void setReviewInfo(String reviewInfo) {
		this.reviewInfo = reviewInfo;
	}

	public String getTeamInfo() {
		return teamInfo;
	}

	public void setTeamInfo(String teamInfo) {
		this.teamInfo = teamInfo;
	}

	public String getPotentialInfo() {
		return potentialInfo;
	}

	public void setPotentialInfo(String potentialInfo) {
		this.potentialInfo = potentialInfo;
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

	public Double getAreaScore() {
		return areaScore;
	}

	public void setAreaScore(Double areaScore) {
		this.areaScore = areaScore;
	}

	public Double getResourceScore() {
		return resourceScore;
	}

	public void setResourceScore(Double resourceScore) {
		this.resourceScore = resourceScore;
	}

	public Double getManageScore() {
		return manageScore;
	}

	public void setManageScore(Double manageScore) {
		this.manageScore = manageScore;
	}

	public Double getReviewScore() {
		return reviewScore;
	}

	public void setReviewScore(Double reviewScore) {
		this.reviewScore = reviewScore;
	}

	public Double getTeamScore() {
		return teamScore;
	}

	public void setTeamScore(Double teamScore) {
		this.teamScore = teamScore;
	}

	public Double getPotentialScore() {
		return potentialScore;
	}

	public void setPotentialScore(Double potentialScore) {
		this.potentialScore = potentialScore;
	}

	public String getProjectName() {
		return projectName;
	}

	public void setProjectName(String projectName) {
		this.projectName = projectName;
	}

	public String getIds() {
		return ids;
	}

	public void setIds(String ids) {
		this.ids = ids;
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

	public String getScount() {
		return scount;
	}

	public void setScount(String scount) {
		this.scount = scount;
	}

	public String getKeyWords() {
		return keyWords;
	}

	public void setKeyWords(String keyWords) {
		this.keyWords = keyWords;
	}

	public Integer getIsCurrent() {
		return IsCurrent;
	}

	public void setIsCurrent(Integer isCurrent) {
		IsCurrent = isCurrent;
	}

	public Long getFlowId() {
		return flowId;
	}

	public void setFlowId(Long flowId) {
		this.flowId = flowId;
	}

	@Override
	public String toString() {
		return "ProjectEvaluate [id=" + id + ", projectId=" + projectId + ", areaInfo=" + areaInfo + ", resourceInfo=" + resourceInfo + ", manageInfo=" + manageInfo + ", reviewInfo=" + reviewInfo + ", teamInfo=" + teamInfo
				+ ", potentialInfo=" + potentialInfo + ", createUserId=" + createUserId + ", createTime=" + createTime + ", areaScore=" + areaScore + ", resourceScore=" + resourceScore + ", manageScore=" + manageScore + ", reviewScore="
				+ reviewScore + ", teamScore=" + teamScore + ", potentialScore=" + potentialScore + ", projectName=" + projectName + ", ids=" + ids + ", areaName=" + areaName + ", province=" + province + ", city=" + city + ", statusName="
				+ statusName + ", scount=" + scount + ", keyWords=" + keyWords + ", IsCurrent=" + IsCurrent + ", flowId=" + flowId + "]";
	}

	public byte[] getImg() {
		return img;
	}

	public void setImg(byte[] img) {
		this.img = img;
	}


}
