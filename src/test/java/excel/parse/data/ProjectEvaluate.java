package excel.parse.data;

import java.io.Serializable;
import java.util.Date;

public class ProjectEvaluate implements Serializable {


    /**
     * 项目名称
     */
    private String projectName;


    /**
     * 图片
     */
    private byte[] img;

    /**
     * 所属区域
     */
    private String areaName;

    /**
     * 省份
     */
    private String province;

    /**
     * 市Key
     */
    private Integer cityKey;

    /**
     * 市
     */
    private String city;

    /**
     * 项目所属人
     */
    private String people;


    /**
     * 项目领导人
     */
    private String leader;

    /**
     * 总分
     */
    private double score;

    /**
     * 历史平均分
     */
    private String avg;

    /**
     * 创建时间
     */
    private Date createTime;

    public String getProjectName() {
        return projectName;
    }

    public void setProjectName(String projectName) {
        this.projectName = projectName;
    }

    public byte[] getImg() {
        return img;
    }

    public void setImg(byte[] img) {
        this.img = img;
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

    public double getScore() {
        return score;
    }

    public void setScore(double score) {
        this.score = score;
    }

    public String getAvg() {
        return avg;
    }

    public void setAvg(String avg) {
        this.avg = avg;
    }

    public Date getCreateTime() {
        return createTime;
    }

    public void setCreateTime(Date createTime) {
        this.createTime = createTime;
    }

    @Override
    public String toString() {
        return "ProjectEvaluate{" +
                "projectName='" + projectName + '\'' +
                ", img=" + (img != null ? img.length : null) +
                ", areaName='" + areaName + '\'' +
                ", province='" + province + '\'' +
                ", city='" + city + '\'' +
                ", cityKey='" + cityKey + '\'' +
                ", people='" + people + '\'' +
                ", leader='" + leader + '\'' +
                ", score=" + score +
                ", avg=" + avg +
                ", createTime=" + createTime +
                '}';
    }
}
