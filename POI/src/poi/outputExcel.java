package poi;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class outputExcel {
	//输出 excel2003
	public static void HssfWork() {
		HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
		HSSFSheet hssfSheet = hssfWorkbook.createSheet("TEST");
		// row HSSFROW行 cell HSSFCELL列
		int sum = 3;
		for (int i = 0; i < sum; i++) {
			HSSFRow hssfRow = hssfSheet.createRow(i);
			System.out.println("i:" + i);
			for (int j = 0; j < sum; j++) {
				hssfRow.createCell(j).setCellValue(j);
				System.out.println("j:" + j);
			}
		}
		try {
			FileOutputStream fileOutputStream = new FileOutputStream("d:\\2.xls");
			hssfWorkbook.write(fileOutputStream);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	//输出excel2007及以上
	public static void XssfWork() {
		XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
		XSSFSheet xssfSheet = xssfWorkbook.createSheet("TEST");
		int sum = 3;
		for (int i = 0; i < sum; i++) {
			XSSFRow xssfRow = xssfSheet.createRow(i);
			System.out.println("i:" + i);
			for (int j = 0; j < sum; j++) {
				xssfRow.createCell(j).setCellValue(j);
				System.out.println("j:" + j);
			}
		}
		try {
			FileOutputStream fileOutputStream = new FileOutputStream("d:\\1.xlsx");
			xssfWorkbook.write(fileOutputStream);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	public static void main(String[] args) {
		HssfWork();
		
		XssfWork();
	}
}
