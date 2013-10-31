package com.generate.template.excel;

import org.apache.poi.ss.usermodel.IndexedColors;

public class SheetStyle {

	private Integer leftMargin;

	private Integer topMargin;

	private short titleFontColor;

	private short titleBackgroundColor;

	private short dataFontColor;

	private short dataBackgroundColor;

	/**
	 * Naming all details for a sheets besides data
	 * 
	 * @param tBgColor
	 *            background color of title like "RED", "BLACK"
	 * @param tfgColor
	 *            foreground color of title
	 * @param dbgColor
	 *            background color of data
	 * @param dfgColor
	 *            foreground color of data
	 * @param lMargin
	 *            number of cell skipped on the left
	 * @param tMargin
	 *            number of cell skipped on the top
	 * @throws SecurityException
	 * @throws NoSuchFieldException
	 * @throws IllegalAccessException
	 * @throws IllegalArgumentException
	 */
	public SheetStyle(String tbgColor, String tftColor, 
			String dbgColor, String dftColor, Integer lMargin, Integer tMargin)
			throws IllegalArgumentException, IllegalAccessException,
			NoSuchFieldException, SecurityException {
		titleBackgroundColor = colorTranslate(tbgColor);
		titleFontColor = colorTranslate(tftColor);
		dataBackgroundColor = colorTranslate(dbgColor);
		dataFontColor = colorTranslate(dftColor);
		leftMargin = lMargin;
		topMargin = tMargin;
	}

	public SheetStyle() throws IllegalArgumentException,
			IllegalAccessException, NoSuchFieldException, SecurityException {
		this("BLACK", "WHITE", "WHITE", "BLACK", 1, 1);
	}

	public short colorTranslate(String colorName)
			throws IllegalArgumentException, IllegalAccessException,
			NoSuchFieldException, SecurityException {
		Class<?> fields = IndexedColors.class;
		return ((IndexedColors) fields.getField(colorName).get(this))
				.getIndex();
	}

	/**
	 * @param args
	 * @throws IllegalAccessException
	 * @throws IllegalArgumentException
	 * @throws SecurityException
	 * @throws NoSuchFieldException
	 */
	public static void main(String[] args) throws IllegalArgumentException,
			IllegalAccessException, NoSuchFieldException, SecurityException {
		// TODO Auto-generated method stub
		SheetStyle ss = new SheetStyle();
		System.out.println(ss.colorTranslate("RED"));
	}

	/**
	 * @return the leftMargin
	 */
	public Integer getLeftMargin() {
		return leftMargin;
	}

	/**
	 * @param leftMargin the leftMargin to set
	 */
	public void setLeftMargin(Integer leftMargin) {
		this.leftMargin = leftMargin;
	}

	/**
	 * @return the topMargin
	 */
	public Integer getTopMargin() {
		return topMargin;
	}

	/**
	 * @param topMargin the topMargin to set
	 */
	public void setTopMargin(Integer topMargin) {
		this.topMargin = topMargin;
	}

	/**
	 * @return the titleFontColor
	 */
	public short getTitleFontColor() {
		return titleFontColor;
	}

	/**
	 * @param titleFontColor the titleFontColor to set
	 */
	public void setTitleFontColor(short titleFontColor) {
		this.titleFontColor = titleFontColor;
	}

	/**
	 * @return the titleBackgroundColor
	 */
	public short getTitleBackgroundColor() {
		return titleBackgroundColor;
	}

	/**
	 * @param titleBackgroundColor the titleBackgroundColor to set
	 */
	public void setTitleBackgroundColor(short titleBackgroundColor) {
		this.titleBackgroundColor = titleBackgroundColor;
	}

	/**
	 * @return the dataFontColor
	 */
	public short getDataFontColor() {
		return dataFontColor;
	}

	/**
	 * @param dataFontColor the dataFontColor to set
	 */
	public void setDataFontColor(short dataFontColor) {
		this.dataFontColor = dataFontColor;
	}

	/**
	 * @return the dataBackgroundColor
	 */
	public short getDataBackgroundColor() {
		return dataBackgroundColor;
	}

	/**
	 * @param dataBackgroundColor the dataBackgroundColor to set
	 */
	public void setDataBackgroundColor(short dataBackgroundColor) {
		this.dataBackgroundColor = dataBackgroundColor;
	}

}
