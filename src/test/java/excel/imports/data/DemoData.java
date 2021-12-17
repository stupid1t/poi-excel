package excel.imports.data;

import java.math.BigDecimal;
import java.util.Date;
import java.util.Map;

public class DemoData {

	private BigDecimal bigDecimalHandler;

	private Boolean booleanHandler;

	private Character charHandler;

	private Date dateHandler;

	private Double doubleHandler;

	private Float floatHandler;

	private Integer integerHandler;

	private Long longHandler;

	private Map<String,Object> objectHandler;

	private byte[] imgHandler;

	private String stringHandler;

	private Short shortHandler;

	public Short getShortHandler() {
		return shortHandler;
	}

	public void setShortHandler(Short shortHandler) {
		this.shortHandler = shortHandler;
	}

	public BigDecimal getBigDecimalHandler() {
		return bigDecimalHandler;
	}

	public void setBigDecimalHandler(BigDecimal bigDecimalHandler) {
		this.bigDecimalHandler = bigDecimalHandler;
	}

	public Boolean getBooleanHandler() {
		return booleanHandler;
	}

	public void setBooleanHandler(Boolean booleanHandler) {
		this.booleanHandler = booleanHandler;
	}

	public Character getCharHandler() {
		return charHandler;
	}

	public void setCharHandler(Character charHandler) {
		this.charHandler = charHandler;
	}

	public Date getDateHandler() {
		return dateHandler;
	}

	public void setDateHandler(Date dateHandler) {
		this.dateHandler = dateHandler;
	}

	public Double getDoubleHandler() {
		return doubleHandler;
	}

	public void setDoubleHandler(Double doubleHandler) {
		this.doubleHandler = doubleHandler;
	}

	public Float getFloatHandler() {
		return floatHandler;
	}

	public void setFloatHandler(Float floatHandler) {
		this.floatHandler = floatHandler;
	}

	public Integer getIntegerHandler() {
		return integerHandler;
	}

	public void setIntegerHandler(Integer integerHandler) {
		this.integerHandler = integerHandler;
	}

	public Long getLongHandler() {
		return longHandler;
	}

	public void setLongHandler(Long longHandler) {
		this.longHandler = longHandler;
	}

	public Map<String, Object> getObjectHandler() {
		return objectHandler;
	}

	public void setObjectHandler(Map<String, Object> objectHandler) {
		this.objectHandler = objectHandler;
	}

	public byte[] getImgHandler() {
		return imgHandler;
	}

	public void setImgHandler(byte[] imgHandler) {
		this.imgHandler = imgHandler;
	}

	public String getStringHandler() {
		return stringHandler;
	}

	public void setStringHandler(String stringHandler) {
		this.stringHandler = stringHandler;
	}

	@Override
	public String toString() {
		return "DemoData{" +
				"bigDecimalHandler=" + bigDecimalHandler +
				", booleanHandler=" + booleanHandler +
				", charHandler=" + charHandler +
				", dateHandler=" + dateHandler +
				", doubleHandler=" + doubleHandler +
				", floatHandler=" + floatHandler +
				", integerHandler=" + integerHandler +
				", longHandler=" + longHandler +
				", objectHandler=" + objectHandler +
				", stringHandler='" + stringHandler + '\'' +
				", shortHandler=" + shortHandler +
				'}';
	}
}
