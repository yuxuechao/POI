package poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.DecimalFormat;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class readExcel {
	//读取excel文件
	public static void readHX(String filepath) {
		// 文件后缀名 即类型
		String fileType = filepath.substring(filepath.lastIndexOf(".") + 1, filepath.length());
		Workbook workbook = null;// 读取excel文件
		Sheet sheet = null;
		Row row = null;
		Cell cell = null;
		try {
			File file = new File(filepath);// 读取excel文件
			if (file.exists() && file != null) {// 判断文件是否存在
				if (fileType.equals("xls")) {// 判断文件类型是xls 还是xlsx
					workbook = new HSSFWorkbook(new FileInputStream(file));
				} else if (fileType.equals("xlsx")) {
					workbook = new XSSFWorkbook(new FileInputStream(file));
				}
				for (int i = 0; i < workbook.getNumberOfSheets(); i++) {// 获取sheet数量 i
					sheet = workbook.getSheetAt(i);
					System.out.println("sheet: " + sheet.getSheetName());
					for (int j = 0; j < sheet.getPhysicalNumberOfRows(); j++) {// 获取row
						row = sheet.getRow(j);
						for (int k = 0; k < row.getPhysicalNumberOfCells(); k++) {// 获取cell
							cell = row.getCell(k);
							System.out.println(cell.getStringCellValue());
						}
					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	public static void main(String[] args) {
		readHX("d:\\2.xls");
		readHX("d:\\1.xlsx");
	}
}
