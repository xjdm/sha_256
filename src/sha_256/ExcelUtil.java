package sha_256;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

public class ExcelUtil {
	public static Workbook createConditionWorkBook(Workbook wb, int start, String data) {
		Sheet sheet = wb.getSheet("sheet1");
		Row row = sheet.createRow(start);
		Cell cell = row.createCell(0);
		cell.setCellValue(data);
		return wb;
	}
	/**
	 * 创建excel文档，
	 * @param list 数据
	 * @param keys list中map的key数组集合
	 * @param columnNames excel的列名
	 */
	public static Workbook createWorkBook(int start, List<Map<String, Object>> list, String[] keys, String columnNames[]) {
		// 创建excel工作簿
		Workbook wb = new HSSFWorkbook();
		if (list == null || list.size() == 0) {
			return wb;
		}
		// 创建第一个sheet（页），并命名
		wb.createSheet(list.get(0).get("sheetName").toString());
		// 手动设置列宽。第一个参数表示要为第几列设；，第二个参数表示列的宽度，n为列高的像素数。
		return writeWorkBook(start, wb, 0, list.subList(1, list.size()), keys, columnNames);
	}
	/**
	 * 创建excel文档，
	 * @param list 数据
	 * @param keys list中map的key数组集合
	 * @param columnNames excel的列名
	 */
	public static Workbook writeWorkBook(int start, Workbook wb, int sheetIndex, List<Map<String, Object>> list, String[] keys, String[] columnNames) {
		Sheet sheet = wb.getSheetAt(sheetIndex);
		// 手动设置列宽。第一个参数表示要为第几列设；，第二个参数表示列的宽度，n为列高的像素数。
		for (int i = 0; i < keys.length; i++) {
			sheet.setColumnWidth((short) i, (short) (35.7 * 150));
		}
		// 创建第start行
		Row row = sheet.createRow((short) start);
		// 设置第一种单元格的样式（用于列名）
		CellStyle cs = wb.createCellStyle();
		setTitleFontStyle(wb, cs);
		setCellStyle(cs);
		// 设置第二种单元格的样式（用于值）
		CellStyle cs2 = wb.createCellStyle();
		setFontStyle(wb, cs2);
		setCellStyle(cs2);
		// 设置列名 ：表头
		for (int i = 0; i < columnNames.length; i++) {
			Cell cell = row.createCell(i);
			cell.setCellValue(columnNames[i]);
			cell.setCellStyle(cs);
		}
		// 设置每行每列的值
		for (int i = 1; i <= list.size(); i++) {
			// Row 行,Cell 方格 , Row 和 Cell 都是从0开始计数的
			// 创建一行，在页sheet上
			Row row1 = sheet.createRow(start + i);
			// 在row行上创建一个方格
			for (int j = 0; j < keys.length; j++) {
				Cell cell = row1.createCell(j);
				cell.setCellValue(list.get(i - 1).get(keys[j]) == null ? " " : list.get(i - 1).get(keys[j]).toString());
				cell.setCellStyle(cs2);
			}
		}
		return wb;
	}
	private static CellStyle setTitleFontStyle(Workbook wb, CellStyle cs) {
		Font f = wb.createFont();
		f.setFontHeightInPoints((short) 10);
		f.setColor(IndexedColors.BLACK.getIndex());
		f.setBoldweight(Font.BOLDWEIGHT_BOLD);
		cs.setFont(f);
		return cs;
	}
	private static CellStyle setFontStyle(Workbook wb, CellStyle cs2) {
		Font f2 = wb.createFont();
		f2.setFontHeightInPoints((short) 10);
		f2.setColor(IndexedColors.BLACK.getIndex());
		cs2.setFont(f2);
		return cs2;
	}
	private static CellStyle setCellStyle(CellStyle cs2) {
		cs2.setBorderLeft(CellStyle.BORDER_THIN);
		cs2.setBorderRight(CellStyle.BORDER_THIN);
		cs2.setBorderTop(CellStyle.BORDER_THIN);
		cs2.setBorderBottom(CellStyle.BORDER_THIN);
		cs2.setAlignment(CellStyle.ALIGN_CENTER);
		return cs2;
	}
	/**
	 * 创建excel文档，
	 * @param list 数据
	 * @param keys list中map的key数组集合
	 * @param columnNames excel的列名
	 */
	public static Workbook createWorkBook(List<Map<String, Object>> list, String[] keys, String columnNames[]) {
		return createWorkBook(0, list, keys, columnNames);
	}
	public static float getExcelCellAutoHeight(String str, float fontCountInline) {
		float defaultRowHeight = 12.00f;// 每一行的高度指定
		float defaultCount = 0.00f;
		for (int i = 0; i < str.length(); i++) {
			float ff = getregex(str.substring(i, i + 1));
			defaultCount = defaultCount + ff;
		}
		return ((int) (defaultCount / fontCountInline) + 1) * defaultRowHeight;// 计算
	}
	public static float getregex(String charStr) {
		if (charStr.equals(" ")) {
			return 0.5f;
		}
		// 判断是否为字母或字符
		if (Pattern.compile("^[A-Za-z0-9]+$").matcher(charStr).matches()) {
			return 0.5f;
		}
		// 判断是否为全角
		if (Pattern.compile("[\u4e00-\u9fa5]+$").matcher(charStr).matches()) {
			return 1.00f;
		}
		// 全角符号 及中文
		if (Pattern.compile("[^x00-xff]").matcher(charStr).matches()) {
			return 1.00f;
		}
		return 0.5f;
	}
}
