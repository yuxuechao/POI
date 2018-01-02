package poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * filepath 文件路径 filename 文件名 savepath保存路径
 */
public class excelCopy {
	public static void readAndWriteExcel(String filepath, String filename, String savepath) {
		// 文件后缀名 即类型
		String fileType = filepath.substring(filepath.lastIndexOf(".") + 1, filepath.length());
		Workbook workbook = null;// 读取excel文件
		Sheet sheet = null;
		Row row = null;
		Cell cell = null;
		Workbook wbCreat = null;// 创建一个新的excel容器 接收excel类型
		Sheet sheetCreat = null;
		Row rowCreat = null;
		CellStyle cellstyle = null;
		try {
			File file = new File(filepath);// 读取excel文件
			if (file.exists() && file != null) {// 判断文件是否存在
				if (fileType.equals("xls")) {// 判断文件类型是xls 还是xlsx
					workbook = new HSSFWorkbook(new FileInputStream(file));
					wbCreat = new HSSFWorkbook();
				} else if (fileType.equals("xlsx")) {
					workbook = new XSSFWorkbook(new FileInputStream(file));
					wbCreat = new XSSFWorkbook();
				}
				cellstyle = wbCreat.createCellStyle();
				DecimalFormat df = new DecimalFormat("0");
				for (int i = 0; i < workbook.getNumberOfSheets(); i++) {// 获取sheet数量 i
					sheet = workbook.getSheetAt(i);
					sheetCreat = wbCreat.createSheet(sheet.getSheetName());// 新建sheet 获取相同的name
					for (int j = 0; j < sheet.getPhysicalNumberOfRows(); j++) {// 获取row
						rowCreat = sheetCreat.createRow(j);// 创建新建excel Sheet的行
						row = sheet.getRow(j);
						for (int k = 0; k < row.getPhysicalNumberOfCells(); k++) {// 获取cell
							sheetCreat.autoSizeColumn(k);// 获取自适应宽度
							cell = row.getCell(k);
							rowCreat.createCell(k);// 创建新建excel Sheet的列
							rowCreat.getCell(k).setCellStyle(cellstyle);
							String cellVal = "";
							if (cell.getCellType() == cell.CELL_TYPE_NUMERIC) {
								cellVal = df.format(cell.getNumericCellValue()).trim();
							} else if (cell.getCellType() == cell.CELL_TYPE_STRING) {
								cellVal = cell.getStringCellValue().trim();
							} else if (cell.getCellType() == cell.CELL_TYPE_BOOLEAN) {
								cellVal = String.valueOf(cell.getBooleanCellValue()).trim();
							}
							// 填充cell的值
							rowCreat.getCell(k).setCellValue(cellVal);

						}
					}
				}
			}
			try {
				// 输出复制的excel
				FileOutputStream fileOutputStream = new FileOutputStream(savepath + "\\copy" + filename);
				wbCreat.write(fileOutputStream);
				fileOutputStream.close();
			} catch (Exception e) {
				e.printStackTrace();
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				if (workbook != null) {
					workbook.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
		}

	}
}
