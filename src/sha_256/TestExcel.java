package sha_256;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author chenjie
 * @date 2019/9/2 16:42
 */
public class TestExcel {
	public static void main(String[] args) {
		File file = new File("d://1.xls");
		OutputStream outputStream = null;
		try {
			outputStream = new FileOutputStream(file);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		List<Map<String, Object>> list = new ArrayList<>();
		Map<String, Object> sheetMap = new HashMap<String, Object>();
		sheetMap.put("sheetName", "sheet1");
		list.add(sheetMap);

		Map<String, Object> map = new HashMap<>();
		map.put("name","cj");
		map.put("age","26");
		list.add(map);

		Map<String, Object> map2 = new HashMap<>();
		map2.put("name","td");
		map2.put("age","28");
		list.add(map2);
		// 列名
		String columnNames[] = { "姓名","年龄" };
		// map中的key
		String keys[] = { "name","age"};
		try {
			Workbook wb = ExcelUtil.createWorkBook(0, list, keys, columnNames);
			wb.write(outputStream);
		} catch (IOException e) {

		}
	}

}
