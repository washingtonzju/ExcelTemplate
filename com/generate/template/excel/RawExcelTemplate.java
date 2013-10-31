package com.generate.template.excel;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * TO-DO List: The root template of excel file. TO-Do Margin: L,R,T,B unit cell
 * Color and Fonts Background Color etc. Change the margins for each Sheet.
 * 
 * @author Zhongkai Mi
 * @since 2013-10-30
 */
public class RawExcelTemplate {

	private Workbook wb = null;

	private String fileName;

	private List<Sheet> tabs = null;

	/**
	 * Naming the file name, create a excel file.
	 * 
	 * @param fName
	 * @param sheetNum
	 */
	public RawExcelTemplate(String fName) {
		fileName = fName;
		tabs = new ArrayList<Sheet>();
		wb = new XSSFWorkbook();
	}

	public void fillCell(Object obj, Cell cell) {
		if (obj instanceof Integer) {
			cell.setCellValue((Integer) obj);
		} else if (obj instanceof Double) {
			cell.setCellValue((Double) obj);
		} else if (obj instanceof Long) {
			cell.setCellValue((Long) obj);
		} else if (obj instanceof Boolean) {
			cell.setCellValue((Boolean) obj);
		} else {
			cell.setCellValue(obj.toString());
		}
	}

	/**
	 * Fill the excel with a named new Sheet and responding data
	 * 
	 * @param sName
	 * @param data
	 */
	public void addSheet(String sName, List<List<Object>> data, SheetStyle ss) {
		
		CellStyle style = wb.createCellStyle();
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style.setFillBackgroundColor(ss.getTitleBackgroundColor());
		style.setFillForegroundColor(ss.getTitleBackgroundColor());
		
		Font font = wb.createFont();
		font.setColor(ss.getTitleFontColor());
		style.setFont(font);
		
		Sheet sheet = wb.createSheet(sName);
		Row titles = sheet.createRow(ss.getTopMargin());
		int cnt = 0;
		for (Object obj : data.get(0)) {
			Cell cell = titles.createCell(ss.getLeftMargin() + cnt);
			cell.setCellStyle(style);
			fillCell(obj, cell);
			++cnt;
		}

		Iterator<List<Object>> itr = data.iterator();
		if (itr.hasNext())
			itr.next();
		else
			return;
		
		
		CellStyle dStyle = wb.createCellStyle();
		dStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		dStyle.setFillBackgroundColor(ss.getDataBackgroundColor());
		dStyle.setFillForegroundColor(ss.getDataBackgroundColor());
		
		Font dFont = wb.createFont();
		dFont.setColor(ss.getDataFontColor());
		dStyle.setFont(dFont);
		
		int rntCnt = 1;
		while (itr.hasNext()) {
			cnt = 0;
			List<Object> unit = itr.next();
			Iterator<Object> cellIter = unit.iterator();
			Row dataRow = sheet.createRow(ss.getTopMargin() + rntCnt);
			while (cellIter.hasNext()) {
				Object obj = cellIter.next();
				Cell cell = dataRow.createCell(ss.getTopMargin() + cnt);
				cell.setCellStyle(dStyle);
				fillCell(obj, cell);
				++cnt;
			}
			++rntCnt;
		}
	}

	public void addSheet(List<List<Object>> data)
			throws IllegalArgumentException, IllegalAccessException,
			NoSuchFieldException, SecurityException {
		SheetStyle ss = new SheetStyle();
		addSheet(String.format("Sheet %d", tabs.size()), data, ss);
	}

	public void writeToFile() throws IOException {
		FileOutputStream fileOut = new FileOutputStream(fileName);
		wb.write(fileOut);
		fileOut.close();
	}

	/**
	 * @param args
	 * @throws IOException
	 * @throws SecurityException 
	 * @throws NoSuchFieldException 
	 * @throws IllegalAccessException 
	 * @throws IllegalArgumentException 
	 */
	public static void main(String[] args) throws IOException, IllegalArgumentException, IllegalAccessException, NoSuchFieldException, SecurityException {
		// TODO Auto-generated method stub
		RawExcelTemplate ret = new RawExcelTemplate("template.xlsx");
		List<List<Object>> data = new ArrayList<List<Object>>();

		ArrayList<Object> titles = new ArrayList<Object>() {
			/**
			 * 
			 */
			private static final long serialVersionUID = 1L;

			{
				this.add("Name");
				this.add("英文的");
				this.add("Titles");
				this.add("成功人士");
			}
		};

		data.add(titles);

		for (int i = 0; i < 5; ++i) {
			ArrayList<Object> datum = new ArrayList<Object>() {
				/**
				 * 
				 */
				private static final long serialVersionUID = -2684033447098261702L;

				{
					this.add("Washington");
					this.add(1.02);
					this.add(100L);
					this.add(50);
				}
			};
			data.add(datum);
		}
		ret.addSheet(data);
		SheetStyle ss = new SheetStyle("BLUE", "WHITE","LIGHT_YELLOW", "BLACK", 2, 2);
		ret.addSheet("今天是星期三啦", data, ss);
		ret.writeToFile();
	}
}
