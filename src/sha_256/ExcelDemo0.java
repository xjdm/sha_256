package sha_256; 
   
import java.io.File; 
import java.io.FileOutputStream; 
import java.io.OutputStream; 
import org.apache.poi.hssf.usermodel.HSSFRow; 
import org.apache.poi.hssf.usermodel.HSSFSheet; 
import org.apache.poi.hssf.usermodel.HSSFWorkbook; 
   
public class ExcelDemo0 { 
  /** 
   * java生成excel文件并写入磁盘 
   * 
   * @author：tuzongxun 
   * @Title: main 
   * @param@param args 
   * @return void 
   * @date Apr 28,2016 7:32:52 PM 
   * @throws 
   */
  public static void main(String[] args) { 
    //C:\Users\tuzongxun123\Desktop桌面，windows和linux的斜杠不一样，而且java对于“/”需要转义处理，File.separator可以实现跨平台 
    File file = new File("d://1.xls"); 
    try { 
      OutputStream outputStream = new FileOutputStream(file); 
      // 创建excel文件，注意这里的hssf是excel2007及以前版本可用，2007版以后的不可用，要用xssf 
      HSSFWorkbook workbook = new HSSFWorkbook(); 
      // 创建excel工作表 
      HSSFSheet sheet = workbook.createSheet("user"); 
      // 为工作表增加一行 
      HSSFRow row = sheet.createRow(0); 
      // 在指定的行上增加两个单元格 
      row.createCell(0).setCellValue("name"); 
      row.createCell(1).setCellValue("password"); 
      // 调用输出流把excel文件写入到磁盘 
      workbook.write(outputStream); 
      // 关闭输出流 
      outputStream.close(); 
    } catch (Exception e) { 
      e.printStackTrace(); 
    } 
  } 
}