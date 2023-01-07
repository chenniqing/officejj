package cn.javaex.officejj.common.entity;

/**
 * 图片属性
 * 
 * @author 陈霓清
 */
public class Picture {
	private Integer width;          // 图片宽度
	private Integer height;         // 图片高度
	private String url;             // 图片路径
	private String description;     // 图片描述

	public Picture() {
		
	}

	public Picture(String url) {
		this.url = url;
	}

	public Picture(Integer width, Integer height, String url) {
		this.width = width;
		this.height = height;
		this.url = url;
	}

	public Picture(Integer width, Integer height, String url, String description) {
		this.width = width;
		this.height = height;
		this.url = url;
		this.description = description;
	}
	
	/**
	 * 得到图片宽度
	 * @return
	 */
	public Integer getWidth() {
		return width;
	}
	/**
	 * 设置图片宽度
	 * @param width
	 */
	public void setWidth(Integer width) {
		this.width = width;
	}

	/**
	 * 得到图片高度
	 * @return
	 */
	public Integer getHeight() {
		return height;
	}
	/**
	 * 设置图片高度
	 * @param height
	 */
	public void setHeight(Integer height) {
		this.height = height;
	}

	/**
	 * 得到图片路径
	 * @return
	 */
	public String getUrl() {
		return url;
	}
	/**
	 * 设置图片路径
	 * @param url
	 */
	public void setUrl(String url) {
		this.url = url;
	}

	/**
	 * 得到图片描述
	 * @return
	 */
	public String getDescription() {
		return description;
	}
	/**
	 * 设置图片描述（可不填）
	 *     填写后，在图片下方显示
	 * @param description
	 */
	public void setDescription(String description) {
		this.description = description;
	}

}
