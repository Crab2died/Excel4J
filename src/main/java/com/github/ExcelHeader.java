package com.github;

/**
 *
 * 功能说明: 用来存储Excel标题的对象，通过该对象可以获取标题和方法的对应关系
 * 
 * <br/>
 * 
 * 修改历史:<br/>
 *
 */
public class ExcelHeader implements Comparable<ExcelHeader> {
	/**
	 * excel的标题名称
	 */
	private String title;
	/**
	 * 每一个标题的顺序
	 */
	private int order;
	/**
	 * 注解域
	 */
	private String filed;

	public String getTitle() {
		return title;
	}

	public void setTitle(String title) {
		this.title = title;
	}

	public int getOrder() {
		return order;
	}

	public void setOrder(int order) {
		this.order = order;
	}

	public String getFiled() {
		return filed;
	}

	public void setFiled(String filed) {
		this.filed = filed;
	}

	public int compareTo(ExcelHeader o) {
		return order > o.order ? 1 : (order < o.order ? -1 : 0);
	}

	public ExcelHeader(String title, int order, String filed) {
		super();
		this.title = title;
		this.order = order;
		this.filed = filed;
	}

	@Override
	public String toString() {
		return "ExcelHeader [title=" + title + ", order=" + order + ", filed=" + filed + "]";
	}

}
