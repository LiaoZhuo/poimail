import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Solution {
	// 使用POI创建excel工作簿
	public static void createWorkBook() throws IOException {
		// 创建excel工作簿
		Workbook wb = new HSSFWorkbook();
		// 创建第一个sheet（页），命名为 new sheet
		Sheet sheet = wb.createSheet("new sheet");
		// Row 行
		// Cell 方格
		// Row 和 Cell 都是从0开始计数的

		// 创建一行，在页sheet上
		Row row = sheet.createRow((short) 0);
		// 在row行上创建一个方格
		Cell cell = row.createCell(0);
		// 设置方格的显示
		cell.setCellValue(1);

		// Or do it on one line.
		row.createCell(1).setCellValue(1.2);
		row.createCell(2).setCellValue("This is a string 速度反馈链接");
		row.createCell(3).setCellValue(true);

		// 创建一个文件 命名为workbook.xls
		FileOutputStream fileOut = new FileOutputStream("workbook.xls");
		// 把上面创建的工作簿输出到文件中
		wb.write(fileOut);
		// 关闭输出流
		fileOut.close();
	}

	// 使用POI读入excel工作簿文件
	public static void readWorkBook() throws Exception {
		// poi读取excel
		// 创建要读入的文件的输入流
		InputStream inp = new FileInputStream("D:\\tmp\\8月份考勤明细表2.xls");

		// 根据上述创建的输入流 创建工作簿对象
		Workbook wb = WorkbookFactory.create(inp);
		// 得到第一页 sheet
		// 页Sheet是从0开始索引的
		Sheet sheet = wb.getSheetAt(0);
		// 在第一行中找到上班时间和下班时间的索引
		int startWorkTimeIndex = -1;
		int stopWorkTimeIndex = -1;
		int dateIndex = -1;
		for (Cell cell : sheet.getRow(0)) {
			if (cell.getStringCellValue().equals("上班时间")) {
				startWorkTimeIndex = cell.getColumnIndex();
			} else if (cell.getStringCellValue().equals("下班时间")) {
				stopWorkTimeIndex = cell.getColumnIndex();
			} else if (cell.getStringCellValue().equals("日期")) {
				dateIndex = cell.getColumnIndex();
			}
		}
		if (startWorkTimeIndex == -1 || stopWorkTimeIndex == -1) {
			System.out.println("没有找到上班时间或者下班时间");
			return;
		}
		SimpleDateFormat sdf = new SimpleDateFormat("HH:mm:ss");
		SimpleDateFormat sdf2 = new SimpleDateFormat("yyyy-MM-dd");
		// 利用foreach循环 遍历sheet中的所有行
		Calendar c = Calendar.getInstance();
		for (int i = 1; i < sheet.getLastRowNum(); ++i) {
			// 遍历row中的所有方格
			long startTime = 0;
			long endTime = 0;
			long date = 0;
			String name = null;
			boolean isSundayOrSaturday = false;
			for (Cell cell : sheet.getRow(i)) {
				if (cell.getColumnIndex() == dateIndex) {
					date = sdf2.parse(cell.getStringCellValue()).getTime();
					c.setTimeInMillis(date);
					int week = c.get(Calendar.DAY_OF_WEEK);
					if (week == Calendar.SUNDAY || week == Calendar.SATURDAY) {
						isSundayOrSaturday = true;
						break;
					}
				}
				if (cell.getColumnIndex() == startWorkTimeIndex) {
					if (!"".equals(cell.getStringCellValue().trim())) {
						startTime = sdf.parse(cell.getStringCellValue()).getTime();
					}
				} else if (cell.getColumnIndex() == stopWorkTimeIndex) {
					if (!"".equals(cell.getStringCellValue().trim())) {
						endTime = sdf.parse(cell.getStringCellValue()).getTime();
					}
				} else if (cell.getColumnIndex() == 1) {
					name = cell.getStringCellValue();
				}
			}
			if (isSundayOrSaturday)
				continue;
			System.out.print(name + " " + sdf2.format(new Date(date)) + " " + getDayOfWeek(c));
			if (startTime == 0) {
				System.out.print("->上班未打卡");
			} else {
				System.out.print("->上班正常");
			}
			if (endTime == 0) {
				System.out.print("->下班未打卡");
			} else {
				System.out.print("->下班正常");
			}
//			else {
				long duration = endTime - startTime;
				// 每一个行输出之后换行
				if (startTime != 0 && endTime != 0) {
					if ( duration >= (8 * 3600 * 1000)) {
						//System.out.print("<正常>");
					} else {
						System.out.print("<迟到>");
					}
				}
//			}
			System.out.println();
		}
		// 关闭输入流
		inp.close();
	}

	public static String getDayOfWeek(Calendar c) {
		int dayOfWeek = c.get(Calendar.DAY_OF_WEEK);
		if (dayOfWeek == Calendar.MONDAY) {
			return "星期一";
		}
		if (dayOfWeek == Calendar.TUESDAY) {
			return "星期二";
		}
		if (dayOfWeek == Calendar.WEDNESDAY) {
			return "星期三";
		}
		if (dayOfWeek == Calendar.THURSDAY) {
			return "星期四";
		}
		if (dayOfWeek == Calendar.FRIDAY) {
			return "星期五";
		}
		if (dayOfWeek == Calendar.SATURDAY) {
			return "星期六";
		}
		return "星期日";
	}

	public static void main(String[] args) {
		try {
			Solution.readWorkBook();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
