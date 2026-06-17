package cn.javaex.officejj.common.entity;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

/**
 * 图片属性
 * 
 * @author 陈霓清
 */
public class Picture {
	/** Excel里未指定宽度时的默认图片展示宽度，单位：像素 */
	private static final Integer DEFAULT_WIDTH = 120;
	/** Excel里未指定高度时的默认图片展示高度，单位：像素 */
	private static final Integer DEFAULT_HEIGHT = 80;
	/** Excel里图片默认内边距，单位：像素 */
	private static final Integer DEFAULT_PADDING = 5;
	/** 一个单元格里插入多张图片时，默认每行最多显示几张 */
	private static final Integer DEFAULT_MAX_COLUMNS = 3;

	private Integer width;          // 图片宽度
	private Integer height;         // 图片高度
	private String url;             // 图片路径
	private String description;     // 图片描述
	private byte[] data;            // 图片字节内容
	private String fileSuffix;      // 图片类型（如"png","jpg"等，方便ImageIO处理）
	private Boolean keepRatio = true;    // 是否保持原图比例，默认保持，避免图片被拉伸变形
	private Boolean resizeCell = true;   // 是否自动撑开Excel单元格，默认撑开，避免图片挤在很矮的行里
	private Integer padding = DEFAULT_PADDING;    // 图片和单元格边框的内边距
	private Integer maxColumns = DEFAULT_MAX_COLUMNS;    // 多图写入同一单元格时，每行最多显示几张
	
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
	 * 创建按默认大小等比展示的图片。
	 * @param url 图片路径
	 * @return
	 */
	public static Picture of(String url) {
		return new Picture(url).fit(DEFAULT_WIDTH, DEFAULT_HEIGHT);
	}

	/**
	 * 创建按指定最大宽高等比展示的图片。
	 * @param url 图片路径
	 * @param width 最大展示宽度，单位：像素
	 * @param height 最大展示高度，单位：像素
	 * @return
	 */
	public static Picture fit(String url, Integer width, Integer height) {
		return new Picture(url).fit(width, height);
	}

	/**
	 * 创建按指定宽高强制展示的图片，不保持原图比例。
	 * @param url 图片路径
	 * @param width 展示宽度，单位：像素
	 * @param height 展示高度，单位：像素
	 * @return
	 */
	public static Picture fixed(String url, Integer width, Integer height) {
		return new Picture(url).fixed(width, height);
	}

	/**
	 * 把图片路径集合转换成 Picture 集合，便于 DTO 字段直接声明 List<Picture>。
	 * @param urls 图片路径集合
	 * @return
	 */
	public static List<Picture> list(Collection<String> urls) {
		List<Picture> list = new ArrayList<Picture>();
		if (urls==null || urls.isEmpty()) {
			return list;
		}

		for (String url : urls) {
			if (url!=null && url.length()>0) {
				list.add(new Picture(url));
			}
		}
		return list;
	}

	/**
	 * 把图片路径集合转换成 Picture 集合，并设置每行最多显示几张。
	 * @param urls 图片路径集合
	 * @param maxColumns 每行最多图片数
	 * @return
	 */
	public static List<Picture> list(Collection<String> urls, Integer maxColumns) {
		List<Picture> list = list(urls);
		for (Picture picture : list) {
			picture.columns(maxColumns);
		}
		return list;
	}

	/**
	 * 把图片路径集合转换成按指定最大宽高等比展示的 Picture 集合。
	 * @param urls 图片路径集合
	 * @param width 单张图片最大展示宽度，单位：像素
	 * @param height 单张图片最大展示高度，单位：像素
	 * @return
	 */
	public static List<Picture> fitList(Collection<String> urls, Integer width, Integer height) {
		List<Picture> list = list(urls);
		for (Picture picture : list) {
			picture.fit(width, height);
		}
		return list;
	}

	/**
	 * 把图片路径集合转换成按指定最大宽高等比展示的 Picture 集合，并设置每行最多显示几张。
	 * @param urls 图片路径集合
	 * @param width 单张图片最大展示宽度，单位：像素
	 * @param height 单张图片最大展示高度，单位：像素
	 * @param maxColumns 每行最多图片数
	 * @return
	 */
	public static List<Picture> fitList(Collection<String> urls, Integer width, Integer height, Integer maxColumns) {
		List<Picture> list = fitList(urls, width, height);
		for (Picture picture : list) {
			picture.columns(maxColumns);
		}
		return list;
	}

	/**
	 * 设置图片最大展示宽高，并保持原图比例。
	 * @param width 最大展示宽度，单位：像素
	 * @param height 最大展示高度，单位：像素
	 * @return
	 */
	public Picture fit(Integer width, Integer height) {
		this.width = width;
		this.height = height;
		this.keepRatio = true;
		this.resizeCell = true;
		return this;
	}

	/**
	 * 设置图片固定展示宽高，不保持原图比例。
	 * @param width 展示宽度，单位：像素
	 * @param height 展示高度，单位：像素
	 * @return
	 */
	public Picture fixed(Integer width, Integer height) {
		this.width = width;
		this.height = height;
		this.keepRatio = false;
		this.resizeCell = true;
		return this;
	}

	/**
	 * 设置图片内边距。
	 * @param padding 内边距，单位：像素
	 * @return
	 */
	public Picture padding(Integer padding) {
		this.padding = padding;
		return this;
	}

	/**
	 * 设置一个单元格里多张图片每行最多显示几张。
	 * @param maxColumns 每行最多图片数
	 * @return
	 */
	public Picture columns(Integer maxColumns) {
		this.maxColumns = maxColumns;
		return this;
	}

	/**
	 * 设置是否保持原图比例。
	 * @param keepRatio 是否保持比例
	 * @return
	 */
	public Picture keepRatio(Boolean keepRatio) {
		this.keepRatio = keepRatio;
		return this;
	}

	/**
	 * 设置是否自动撑开Excel单元格。
	 * @param resizeCell 是否自动撑开单元格
	 * @return
	 */
	public Picture resizeCell(Boolean resizeCell) {
		this.resizeCell = resizeCell;
		return this;
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

	public byte[] getData() {
		return data;
	}

	public void setData(byte[] data) {
		this.data = data;
	}

	public String getFileSuffix() {
		return fileSuffix;
	}

	public void setFileSuffix(String fileSuffix) {
		this.fileSuffix = fileSuffix;
	}

	public Boolean getKeepRatio() {
		return keepRatio;
	}

	public void setKeepRatio(Boolean keepRatio) {
		this.keepRatio = keepRatio;
	}

	public Boolean getResizeCell() {
		return resizeCell;
	}

	public void setResizeCell(Boolean resizeCell) {
		this.resizeCell = resizeCell;
	}

	public Integer getPadding() {
		return padding;
	}

	public void setPadding(Integer padding) {
		this.padding = padding;
	}

	public Integer getMaxColumns() {
		return maxColumns;
	}

	public void setMaxColumns(Integer maxColumns) {
		this.maxColumns = maxColumns;
	}

	/**
	 * 得到有效展示宽度，未设置时使用默认值。
	 * @return
	 */
	public Integer getEffectiveWidth() {
		return width==null || width<=0 ? DEFAULT_WIDTH : width;
	}

	/**
	 * 得到有效展示高度，未设置时使用默认值。
	 * @return
	 */
	public Integer getEffectiveHeight() {
		return height==null || height<=0 ? DEFAULT_HEIGHT : height;
	}

	/**
	 * 得到有效内边距，未设置或小于0时使用默认值。
	 * @return
	 */
	public Integer getEffectivePadding() {
		return padding==null || padding<0 ? DEFAULT_PADDING : padding;
	}

	/**
	 * 得到有效最大列数，未设置或小于1时使用默认值。
	 * @return
	 */
	public Integer getEffectiveMaxColumns() {
		return maxColumns==null || maxColumns<1 ? DEFAULT_MAX_COLUMNS : maxColumns;
	}

}
