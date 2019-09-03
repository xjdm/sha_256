package sha_256;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

/**
 * SHA256 摘要算法工具类
 * 
 * @author Administrator
 *
 */
public class SHA256Util {
	public static String getSHA256StrJava(String str) {
		MessageDigest messageDigest;
		String encodeStr = "";
		try {
			
			messageDigest = MessageDigest.getInstance("SHA-256");
			messageDigest.update(str.getBytes("UTF-8"));
			encodeStr = byte2Hex(messageDigest.digest());
		} catch (NoSuchAlgorithmException e) {
			e.printStackTrace();
		} catch (UnsupportedEncodingException e) {
			e.printStackTrace();
		}
		return encodeStr;
	}

	private static String byte2Hex(byte[] bytes) {
		StringBuffer stringBuffer = new StringBuffer();
		String temp = null;
		for (int i = 0; i < bytes.length; i++) {
			temp = Integer.toHexString(bytes[i] & 0xFF);
			// 1得到一位的进行补0操作
			if (temp.length() == 1) {
				stringBuffer.append("0");
			}
			stringBuffer.append(temp);
		}
		return stringBuffer.toString();
	}
	
	public static void readWorkBook() throws Exception { 
        // poi读取excel 
        //创建要读入的文件的输入流 
        POIFSFileSystem fs=new POIFSFileSystem(new FileInputStream("d://核验联系电话的金鹏卡号.xls")); 
        
        //根据上述创建的输入流 创建工作簿对象 
        HSSFWorkbook wb = new HSSFWorkbook(fs); 
        
        //得到第一页 sheet 
        //页Sheet是从0开始索引的 
        HSSFSheet sheet = wb.getSheetAt(0); 
        File file = new File("d://2017_11_20_核验联系电话的金鹏卡号_加密.xls"); 
        try { 
          OutputStream outputStream = new FileOutputStream(file); 
          // 创建excel文件，注意这里的hssf是excel2007及以前版本可用，2007版以后的不可用，要用xssf 
          HSSFWorkbook workbook = new HSSFWorkbook(); 
          // 创建excel工作表 
          HSSFSheet sheet2 = workbook.createSheet("user"); 
          // 为工作表增加一行 
          HSSFRow row = sheet2.createRow(0); 
          for (int i=0;i<sheet.getLastRowNum();i++) {
        	  HSSFRow row2 = sheet.getRow(i);
        	  Cell cell = row2.getCell(0);
          	  cell.setCellType(Cell.CELL_TYPE_STRING);
              String stringCellValue = cell.getStringCellValue();
          	
          	  System.out.print(cell.getStringCellValue());
        	  
        	  HSSFRow newRow = sheet2.createRow(i);
        	  newRow.createCell(0).setCellValue(stringCellValue);
        	  newRow.createCell(1).setCellValue(getSHA256StrJava(stringCellValue));
          }
          // 调用输出流把excel文件写入到磁盘 
          workbook.write(outputStream); 
          // 关闭输出流 
          outputStream.close(); 
        } catch (Exception e) { 
          e.printStackTrace(); 
        } 
        //fs.close();
    } 
	
	public static void main(String[] args) throws Exception {
		readWorkBook();
	}
	
	
}
