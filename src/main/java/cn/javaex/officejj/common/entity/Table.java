package cn.javaex.officejj.common.entity;

import java.util.List;

/**
 * 表格设置
 * 
 * @author 陈霓清
 */
public class Table {
	private List<String[]> dataList;      // 数据内容
	private List<int[]> mergeColList;     // 合并列（水平合并）
	private List<int[]> mergeRowList;     // 合并行（垂直合并）
	
	/**
	 * 得到表格数据集合
	 * @return
	 */
	public List<String[]> getDataList() {
		return dataList;
	}
	/**
	 * 设置表格数据集合
	 * @param dataList
	 */
	public void setDataList(List<String[]> dataList) {
		this.dataList = dataList;
	}
	
	/**
	 * 得到表格合并列集合
	 * @return
	 */
	public List<int[]> getMergeColList() {
		return mergeColList;
	}
	/**
	 * 设置表格合并列集合
	 * @param mergeColList
	 */
	public void setMergeColList(List<int[]> mergeColList) {
		this.mergeColList = mergeColList;
	}
	
	/**
	 * 得到表格合并行集合
	 * @return
	 */
	public List<int[]> getMergeRowList() {
		return mergeRowList;
	}
	/**
	 * 设置表格合并行集合
	 * @param mergeRowList
	 */
	public void setMergeRowList(List<int[]> mergeRowList) {
		this.mergeRowList = mergeRowList;
	}
	
}
