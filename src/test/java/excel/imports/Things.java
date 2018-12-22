/*
 * @(#)Things.java        3.0.0   2018年01月03日 上午8:21:29
 *
 * Copyright (c) 2007-2017 A cable international early education，Xi'an，China.
 * All rights reserved.
 *
 * This software is the confidential and proprietary information of Asuo
 * ("Confidential Information").  You shall not
 * disclose such Confidential Information and shall use it only in
 * accordance with the terms of the license agreement you entered into
 * with Asuo.
 */
package excel.imports;

import java.util.Arrays;

/**
 * @ClassName: Things
 * @Description: 物料管理实体
 * @author  杨志华
 * @date: 2018年01月03日 上午08:26:29
 * @version: 3.0.0
 */
public class Things {
    /**
     *
     */
    private static final long serialVersionUID = 1697650729747277733L;
    /**
     * @Fields name : 名称
     */
    private String            name;
    /**
     * @Fields status : 状态 0/1 删除/正常
     */
    private String            status;
    /**
     * @Fields buyPrice : 采购价钱
     */
	private Double buyPrice;
    /**
     * @Fields marketPrice : 市场价钱
     */
	private Double marketPrice;
    /**
     * @Fields thingsDesc : 描述
     */
    private String            thingsDesc;
    /**
     * @Fields createTime : 创建时间
     */
    private long              createTime;
    /**
     * @Fields uuid : uuid
     */
    private String            uuid;

    /**
     * @Fields applyThingsHeadId : 申请单id
     */
    private String            applyThingsHeadId;


    /**
     * @Fields specifications : 规格
     */
    private String            specifications;


    /**
     * @Fields cover : 覆盖（首页提供以展示）
     */
    private String cover;

    /**
     * @Fields conversion : 单位
     */
    private String conversion;
    /**
     * @Fields sumnumber : 物品数量 本字段无值，仅在出入库选择物品 列表时 放置查询出物品总数
     */
    private int sumnumber;

	/**
	 * 市面名称
	 */
	private String extraName;

	/**
	 * 品牌
	 */
	private String brand;

	/**
	 * 型号
	 */
	private String modelType;

	/**
	 * 包装
	 */
	private String pack;

	/**
	 * 到货周期
	 */
	private String week;

	/**
	 * 起订量
	 */
	private Integer startNum;

	/**
	 * 材质
	 */
	private String quality;

	/**
	 * 分类，导入使用
	 */
	private String thingsTypeName;

	/**
	 * 图片流信息，导入使用
	 */
	private byte[] pictureData;

    public Things() {
    	this.status="1";
    }


	public String getExtraName() {
		return extraName;
	}

	public void setExtraName(String extraName) {
		this.extraName = extraName;
	}

	public String getBrand() {
		return brand;
	}

	public void setBrand(String brand) {
		this.brand = brand;
	}

	public String getModelType() {
		return modelType;
	}

	public void setModelType(String modelType) {
		this.modelType = modelType;
	}

	public String getPack() {
		return pack;
	}

	public void setPack(String pack) {
		this.pack = pack;
	}

	public String getWeek() {
		return week;
	}

	public void setWeek(String week) {
		this.week = week;
	}

	public Integer getStartNum() {
		return startNum;
	}

	public void setStartNum(Integer startNum) {
		this.startNum = startNum;
	}

	public Things(String name, String status, double buyPrice, double marketPrice,
            String thingsDesc, String uuid,String specifications,String conversion,String cover) {
        this.name = name;
        this.status = status;
        this.buyPrice = buyPrice;
        this.marketPrice = marketPrice;
        this.thingsDesc = thingsDesc;
        this.uuid = uuid;
        this.specifications = specifications;
        this.conversion = conversion;
        this.cover = cover;
    }


    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getStatus() {
        return status;
    }

    public void setStatus(String status) {
        this.status = status;
    }

	public double getBuyPrice() {
        return buyPrice;
    }

	public void setBuyPrice(double buyPrice) {
        this.buyPrice = buyPrice;
    }

	public double getMarketPrice() {
        return marketPrice;
    }

	public void setMarketPrice(double marketPrice) {
        this.marketPrice = marketPrice;
    }

    public String getThingsDesc() {
        return thingsDesc;
    }

    public void setThingsDesc(String thingsDesc) {
        this.thingsDesc = thingsDesc;
    }

    public long getCreateTime() {
        return createTime;
    }

    public void setCreateTime(long createTime) {
        this.createTime = createTime;
    }

    public String getUuid() {
        return uuid;
    }

    public void setUuid(String uuid) {
        this.uuid = uuid;
    }

    public String getApplyThingsHeadId() {
        return applyThingsHeadId;
    }

    public void setApplyThingsHeadId(String applyThingsHeadId) {
        this.applyThingsHeadId = applyThingsHeadId;
    }

    public String getSpecifications() {
        return specifications;
    }

    public void setSpecifications(String specifications) {
        this.specifications = specifications;
    }

    public String getCover() {
        return cover;
    }

    public void setCover(String cover) {
        this.cover = cover;
    }


    public String getConversion() {
        return conversion;
    }


    public void setConversion(String conversion) {
        this.conversion = conversion;
    }


	public int getSumnumber() {
		return sumnumber;
	}


	public void setSumnumber(int sumnumber) {
		this.sumnumber = sumnumber;
	}


	public String getQuality() {
		return quality;
	}

	public void setQuality(String quality) {
		this.quality = quality;
	}


	public String getThingsTypeName() {
		return thingsTypeName;
	}


	public void setThingsTypeName(String thingsTypeName) {
		this.thingsTypeName = thingsTypeName;
	}


	public byte[] getPictureData() {
		return pictureData;
	}


	public void setPictureData(byte[] pictureData) {
		this.pictureData = pictureData;
	}


	@Override
	public String toString() {
		return "Things [name=" + name + ", status=" + status + ", buyPrice=" + buyPrice + ", marketPrice=" + marketPrice + ", thingsDesc=" + thingsDesc + ", createTime=" + createTime + ", uuid="
				+ uuid
				+ ", applyThingsHeadId=" + applyThingsHeadId + ", specifications=" + specifications + ", cover=" + cover + ", conversion=" + conversion + ", sumnumber=" + sumnumber + ", extraName=" + extraName + ", brand=" + brand
				+ ", modelType=" + modelType + ", pack=" + pack + ", week=" + week + ", startNum=" + startNum + ", quality=" + quality + ", thingsTypeName=" + thingsTypeName + ", pictureData=" + Arrays.toString(pictureData) + "]";
	}
}
